import json

import pytest

from training import db as training_db
from training import migrations
from training.errors import ErrorCode, TrainingError
from tests.test_training_attempt_lifecycle import _make_open_exam_with_assignment, _setup
from services import training_attempt_service as attempts
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

    report = reports.finalize_exam(db_path, unit_code="son_tay", actor="mgr", exam_id=exam_id)

    assert report["revision"] == 1
    assert report["payload"]["summary"]["assigned"] == 1
    assert len(report["payload"]["individual"]) == 1


def test_finalize_retry_returns_existing_revision(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    exam_id, _ = _make_open_exam_with_assignment(db_path)

    first = reports.finalize_exam(db_path, unit_code="son_tay", actor="mgr", exam_id=exam_id)
    second = reports.finalize_exam(db_path, unit_code="son_tay", actor="mgr", exam_id=exam_id)

    assert first["revision"] == second["revision"] == 1


def test_finalize_expires_unstarted_assignments(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    exam_id, assignment_id = _make_open_exam_with_assignment(db_path)

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


def test_finalize_preserves_timeout_status_for_expired_active_attempt(monkeypatch, tmp_path):
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
    assert attempt["status"] == "timed_out"
    assert attempt["ended_reason"] == "timeout"


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
