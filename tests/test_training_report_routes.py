"""Route-level tests for report/finalize/export RBAC and guardrails."""

import json
from pathlib import Path

import pytest
from flask import Flask

from blueprints.auth_routes import auth_bp
from blueprints.training_routes import training_bp
from services.training_catalog_service import grant_role, seed_defaults
from training import db as training_db, migrations, time_policy
from training.errors import ErrorCode, TrainingError
from services import training_attempt_service as attempts
from services import training_exam_service as exams
from services import training_report_service as reports
from tests.test_training_attempt_lifecycle import _make_open_exam_with_assignment, VALID_BATCH
from services import training_question_service as qs


def _setup(monkeypatch, tmp_path):
    db_path = str(tmp_path / "training.db")
    monkeypatch.setattr(training_db, "TRAINING_DB_PATH", db_path)
    monkeypatch.setattr(training_db, "UNIT_CODE", "son_tay")
    migrations.run_migrations(db_path, "son_tay")
    seed_defaults(db_path, "son_tay")
    qs.reset_duplicate_cache()
    return db_path


def _route_client(monkeypatch, tmp_path, *, username, roles=(), dashboard_role="user"):
    from blueprints import training_routes

    db_path = str(tmp_path / "training.db")
    migrations.run_migrations(db_path, "son_tay")
    seed_defaults(db_path, "son_tay")
    for role in roles:
        grant_role(db_path, "son_tay", "seed", username, role)

    monkeypatch.setattr(training_routes.config, "TRAINING_DB_PATH", db_path)
    monkeypatch.setattr(training_routes.config, "TRAINING_EXPORT_DIR", str(tmp_path / "exports"))
    monkeypatch.setattr(training_routes.config, "UNIT_CODE", "son_tay")
    monkeypatch.setattr(
        training_routes,
        "get_user_by_username",
        lambda candidate: {
            "username": candidate,
            "fullname": candidate,
            "role": dashboard_role,
        } if candidate == username else None,
    )

    root = Path(__file__).resolve().parents[1]
    app = Flask(__name__, template_folder=str(root / "templates"), static_folder=str(root / "static"))
    app.config["SECRET_KEY"] = "test"
    app.config["TESTING"] = True
    app.jinja_env.globals["is_endpoint_enabled"] = lambda endpoint: endpoint == "training.page_index"
    app.url_build_error_handlers.append(lambda error, endpoint, values: "#")
    app.register_blueprint(auth_bp)
    app.register_blueprint(training_bp)
    with app.test_client() as client:
        with client.session_transaction() as sess:
            sess["username"] = username
            sess["_csrf_token"] = "test-csrf"
        yield client, db_path


# --- RBAC: learner cannot access report/finalize/export ---

def test_learner_cannot_get_report(monkeypatch, tmp_path):
    for client, db_path in _route_client(monkeypatch, tmp_path, username="learner1", roles=("learner",)):
        resp = client.get("/api/training/exams/fake-exam-id/report")
    assert resp.status_code in (401, 403)


def test_learner_cannot_finalize(monkeypatch, tmp_path):
    for client, db_path in _route_client(monkeypatch, tmp_path, username="learner1", roles=("learner",)):
        resp = client.post(
            "/api/training/exams/fake-exam-id/finalize",
            json={},
            headers={"X-CSRF-Token": "test-csrf"},
        )
    assert resp.status_code in (401, 403)


def test_learner_cannot_export_excel(monkeypatch, tmp_path):
    for client, db_path in _route_client(monkeypatch, tmp_path, username="learner1", roles=("learner",)):
        resp = client.get("/download/training/exams/fake-exam-id/report.xlsx")
    assert resp.status_code in (401, 403)


def test_unauthenticated_cannot_get_report(monkeypatch, tmp_path):
    from blueprints import training_routes
    db_path = str(tmp_path / "training.db")
    migrations.run_migrations(db_path, "son_tay")
    monkeypatch.setattr(training_routes.config, "TRAINING_DB_PATH", db_path)
    monkeypatch.setattr(training_routes.config, "UNIT_CODE", "son_tay")

    root = Path(__file__).resolve().parents[1]
    app = Flask(__name__, template_folder=str(root / "templates"), static_folder=str(root / "static"))
    app.config["SECRET_KEY"] = "test"
    app.config["TESTING"] = True
    app.register_blueprint(auth_bp)
    app.register_blueprint(training_bp)
    with app.test_client() as client:
        resp = client.get("/api/training/exams/fake-exam-id/report")
    assert resp.status_code in (401, 403)


# --- Finalize guardrails ---

def test_finalize_unknown_exam_returns_not_found(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    with pytest.raises(TrainingError) as exc_info:
        reports.finalize_exam(db_path, unit_code="son_tay", actor="mgr", exam_id="nonexistent")
    assert exc_info.value.code == ErrorCode.NOT_FOUND
    assert exc_info.value.status == 404


@pytest.mark.parametrize("status", ["draft", "ready", "cancelled"])
def test_finalize_rejects_invalid_states(monkeypatch, tmp_path, status):
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


def test_finalize_retry_returns_same_revision(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    exam_id, _ = _make_open_exam_with_assignment(db_path)
    exams.close_exam(db_path, unit_code="son_tay", actor="mgr", exam_id=exam_id)

    first = reports.finalize_exam(db_path, unit_code="son_tay", actor="mgr", exam_id=exam_id)
    second = reports.finalize_exam(db_path, unit_code="son_tay", actor="mgr", exam_id=exam_id)

    assert first["revision"] == second["revision"] == 1
    assert first["payload"] == second["payload"]
    conn = training_db.read_connection(db_path)
    try:
        count = conn.execute(
            "SELECT COUNT(*) AS c FROM exam_report_snapshots WHERE exam_event_id=?", (exam_id,)
        ).fetchone()["c"]
        assert count == 1
    finally:
        conn.close()


# --- Report snapshot immutability ---

def test_report_snapshot_is_immutable(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    exam_id, assignment_id = _make_open_exam_with_assignment(db_path)
    started = attempts.start_attempt(
        db_path, unit_code="son_tay", actor="learner1", assignment_id=assignment_id
    )
    attempts.submit_attempt(
        db_path, unit_code="son_tay", actor="learner1", attempt_id=started["attempt_id"]
    )
    exams.close_exam(db_path, unit_code="son_tay", actor="mgr", exam_id=exam_id)
    reports.finalize_exam(db_path, unit_code="son_tay", actor="mgr", exam_id=exam_id)

    r1 = reports.get_report_snapshot(db_path, exam_id)
    r2 = reports.get_report_snapshot(db_path, exam_id)
    assert r1 is not None
    assert r2 is not None
    assert r1 == r2
    assert r1["revision"] == 1


def test_report_does_not_leak_correct_answers(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    exam_id, assignment_id = _make_open_exam_with_assignment(db_path)
    started = attempts.start_attempt(
        db_path, unit_code="son_tay", actor="learner1", assignment_id=assignment_id
    )
    attempts.submit_attempt(
        db_path, unit_code="son_tay", actor="learner1", attempt_id=started["attempt_id"]
    )
    exams.close_exam(db_path, unit_code="son_tay", actor="mgr", exam_id=exam_id)
    report = reports.finalize_exam(db_path, unit_code="son_tay", actor="mgr", exam_id=exam_id)

    report_str = json.dumps(report["payload"])
    assert "correct_option_ids" not in report_str
    assert "explanation" not in report_str
    assert "evidence" not in report_str
    assert "distractor_rationales" not in report_str


# --- Export guardrails ---

def test_export_returns_404_when_not_finalized(monkeypatch, tmp_path):
    for client, db_path in _route_client(monkeypatch, tmp_path, username="admin", roles=("admin",), dashboard_role="admin"):
        resp = client.get("/download/training/exams/nonexistent/report.xlsx")
    assert resp.status_code == 404


def test_manager_can_access_report_after_finalize(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    exam_id, _ = _make_open_exam_with_assignment(db_path)
    exams.close_exam(db_path, unit_code="son_tay", actor="mgr", exam_id=exam_id)
    reports.finalize_exam(db_path, unit_code="son_tay", actor="mgr", exam_id=exam_id)

    for client, _ in _route_client(monkeypatch, tmp_path, username="admin", roles=("admin",), dashboard_role="admin"):
        resp = client.get("/api/training/exams/" + exam_id + "/report")
    assert resp.status_code == 200
    data = resp.get_json()
    assert data["revision"] == 1
    assert "payload" in data
