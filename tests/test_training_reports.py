import json

from training import db as training_db
from training import migrations
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
