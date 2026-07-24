import json

import pytest

from training import db as training_db
from training import migrations
from training import time_policy
from training.errors import ErrorCode, TrainingError
from tests.test_training_attempt_lifecycle import _make_open_exam_with_assignment, _setup
from services import training_attempt_service as attempts
from services import training_exam_service as exams
from services import training_report_service as reports


def test_finalize_creates_immutable_revision(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    exam_id, assignment_id = _make_open_exam_with_assignment(db_path)
    started = attempts.start_attempt(
        db_path, unit_code="son_tay", actor="learner1", assignment_id=assignment_id
    )
    view = attempts.get_attempt_learner_view(db_path, started["attempt_id"])
    for item in view["items"]:
        attempts.save_response(
            db_path, attempt_id=started["attempt_id"], attempt_item_id=item["item_id"],
            selected_option_ids=["B"], client_revision=1,
        )
    attempts.submit_attempt(db_path, unit_code="son_tay", actor="learner1", attempt_id=started["attempt_id"])
    exams.close_exam(db_path, unit_code="son_tay", actor="mgr", exam_id=exam_id)

    report = reports.finalize_exam(db_path, unit_code="son_tay", actor="mgr", exam_id=exam_id)

    assert report["revision"] == 1
    assert report["payload"]["summary"]["assigned"] == 1
    assert len(report["payload"]["individual"]) == 1


def test_finalize_retry_returns_existing_revision(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    exam_id, _ = _make_open_exam_with_assignment(db_path)
    exams.close_exam(db_path, unit_code="son_tay", actor="mgr", exam_id=exam_id)

    first = reports.finalize_exam(db_path, unit_code="son_tay", actor="mgr", exam_id=exam_id)
    second = reports.finalize_exam(db_path, unit_code="son_tay", actor="mgr", exam_id=exam_id)

    assert first["revision"] == second["revision"] == 1


def test_finalize_rejects_open_exam_before_its_end_without_closing(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    exam_id, assignment_id = _make_open_exam_with_assignment(db_path)
    started = attempts.start_attempt(
        db_path, unit_code="son_tay", actor="learner1", assignment_id=assignment_id,
    )

    with pytest.raises(TrainingError) as exc_info:
        reports.finalize_exam(db_path, unit_code="son_tay", actor="mgr", exam_id=exam_id)

    assert exc_info.value.code == ErrorCode.EXAM_NOT_OPEN
    assert exams.get_exam(db_path, exam_id)["status"] == "open"
    assert attempts.get_attempt(db_path, started["attempt_id"])["status"] == "active"
    assert reports.get_report_snapshot(db_path, exam_id) is None


@pytest.mark.parametrize("status", ["draft", "ready", "cancelled"])
def test_finalize_rejects_non_finalizable_exam_states(monkeypatch, tmp_path, status):
    db_path = _setup(monkeypatch, tmp_path)
    exam_id, _ = _make_open_exam_with_assignment(db_path)
    conn = training_db.write_connection(db_path)
    try:
        conn.execute("UPDATE exam_events SET status=? WHERE id=?", (status, exam_id))
        conn.commit()
    finally:
        conn.close()

    with pytest.raises(TrainingError) as exc_info:
        reports.finalize_exam(db_path, unit_code="son_tay", actor="mgr", exam_id=exam_id)

    assert exc_info.value.code == ErrorCode.CONFLICT
    assert exc_info.value.status == 409
    assert reports.get_report_snapshot(db_path, exam_id) is None


def test_finalize_expired_open_exam_times_out_active_attempt_before_closing(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    exam_id, assignment_id = _make_open_exam_with_assignment(db_path)
    started = attempts.start_attempt(
        db_path, unit_code="son_tay", actor="learner1", assignment_id=assignment_id,
    )
    now = 2_000_000
    conn = training_db.write_connection(db_path)
    try:
        conn.execute("UPDATE exam_events SET end_at_ms=? WHERE id=?", (now, exam_id))
        conn.execute("UPDATE exam_attempts SET deadline_at_ms=? WHERE id=?", (now + 10_000, started["attempt_id"]))
        conn.commit()
    finally:
        conn.close()
    monkeypatch.setattr(time_policy, "utc_now_ms", lambda: now + 1)

    reports.finalize_exam(db_path, unit_code="son_tay", actor="mgr", exam_id=exam_id)

    attempt = attempts.get_attempt(db_path, started["attempt_id"])
    assert (attempt["status"], attempt["ended_reason"]) == ("timed_out", "timeout")


def test_finalize_expires_unstarted_assignments(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    exam_id, assignment_id = _make_open_exam_with_assignment(db_path)
    exams.close_exam(db_path, unit_code="son_tay", actor="mgr", exam_id=exam_id)

    reports.finalize_exam(db_path, unit_code="son_tay", actor="mgr", exam_id=exam_id)

    conn = training_db.read_connection(db_path)
    status = conn.execute(
        "SELECT status FROM exam_assignments WHERE id=?", (assignment_id,)
    ).fetchone()["status"]
    conn.close()
    assert status == "expired"


def test_finalize_administratively_submits_active_attempt_and_is_idempotent(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    exam_id, assignment_id = _make_open_exam_with_assignment(db_path)
    started = attempts.start_attempt(db_path, unit_code="son_tay", actor="learner1", assignment_id=assignment_id)
    item_id = attempts.get_attempt_learner_view(db_path, started["attempt_id"])["items"][0]["item_id"]
    attempts.save_response(db_path, attempt_id=started["attempt_id"], attempt_item_id=item_id,
                           selected_option_ids=["B"], client_revision=1)
    exams.close_exam(db_path, unit_code="son_tay", actor="mgr", exam_id=exam_id)

    first = reports.finalize_exam(db_path, unit_code="son_tay", actor="mgr", exam_id=exam_id)
    second = reports.finalize_exam(db_path, unit_code="son_tay", actor="mgr", exam_id=exam_id)

    attempt = attempts.get_attempt(db_path, started["attempt_id"])
    assert attempt["status"] == "administratively_submitted"
    assert attempt["ended_reason"] == "exam_closed"
    assert first == second
    conn = training_db.read_connection(db_path)
    try:
        assert conn.execute("SELECT COUNT(*) AS c FROM exam_results WHERE attempt_id=?", (started["attempt_id"],)).fetchone()["c"] == 1
        assert conn.execute("SELECT COUNT(*) AS c FROM exam_report_snapshots WHERE exam_event_id=?", (exam_id,)).fetchone()["c"] == 1
    finally:
        conn.close()


def test_finalize_administratively_submits_expired_active_closed_attempt(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    exam_id, assignment_id = _make_open_exam_with_assignment(db_path)
    started = attempts.start_attempt(db_path, unit_code="son_tay", actor="learner1", assignment_id=assignment_id)
    conn = training_db.write_connection(db_path)
    try:
        conn.execute("UPDATE exam_events SET status='closed' WHERE id=?", (exam_id,))
        conn.execute("UPDATE exam_attempts SET deadline_at_ms=0 WHERE id=?", (started["attempt_id"],))
        conn.commit()
    finally:
        conn.close()

    reports.finalize_exam(db_path, unit_code="son_tay", actor="mgr", exam_id=exam_id)

    attempt = attempts.get_attempt(db_path, started["attempt_id"])
    assert attempt["status"] == "administratively_submitted"
    assert attempt["ended_reason"] == "exam_closed"


def test_finalize_recovers_active_closed_exam_with_exam_closed_reason(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    exam_id, assignment_id = _make_open_exam_with_assignment(db_path)
    started = attempts.start_attempt(db_path, unit_code="son_tay", actor="learner1", assignment_id=assignment_id)
    conn = training_db.write_connection(db_path)
    try:
        conn.execute("UPDATE exam_events SET status='closed' WHERE id=?", (exam_id,))
        conn.commit()
    finally:
        conn.close()

    reports.finalize_exam(db_path, unit_code="son_tay", actor="mgr", exam_id=exam_id)

    attempt = attempts.get_attempt(db_path, started["attempt_id"])
    assert attempt["status"] == "administratively_submitted"
    assert attempt["ended_reason"] == "exam_closed"


@pytest.mark.parametrize("operation", ["close", "finalize"])
def test_unknown_exam_transitions_raise_not_found(monkeypatch, tmp_path, operation):
    db_path = _setup(monkeypatch, tmp_path)
    service = reports.finalize_exam if operation == "finalize" else __import__(
        "services.training_exam_service", fromlist=["close_exam"]
    ).close_exam

    with pytest.raises(TrainingError) as exc_info:
        service(db_path, unit_code="son_tay", actor="mgr", exam_id="missing")

    assert exc_info.value.code == ErrorCode.NOT_FOUND
    assert exc_info.value.status == 404
