import flask
import pytest

from training import db as training_db
from training import migrations, time_policy
from training.permissions import has_module_role
from services import training_exam_service as exams
from services import training_question_service as qs
from services.training_catalog_service import seed_defaults, grant_role
from blueprints import training_bp


def _setup_db(tmp_path):
    db_path = str(tmp_path / "training.db")
    migrations.run_migrations(db_path, "son_tay")
    seed_defaults(db_path, "son_tay")
    grant_role(db_path, "son_tay", "admin", "admin", "exam_manager")
    grant_role(db_path, "son_tay", "admin", "admin", "learner")
    grant_role(db_path, "son_tay", "admin", "learner1", "learner")
    grant_role(db_path, "son_tay", "admin", "learner2", "learner")
    return db_path


def _publish_questions(db_path):
    from tests.test_training_exam_states import VALID_BATCH
    result = qs.import_question_batch(
        db_path, unit_code="son_tay", actor="admin",
        batch=VALID_BATCH, status="draft",
    )
    for vid in result["version_ids"]:
        qs.add_review_action(db_path, unit_code="son_tay", actor="admin",
                             version_id=vid, action="approve")
        qs.publish_question_version(db_path, unit_code="son_tay", actor="admin", version_id=vid)
    return result["version_ids"]


def _make_exam_with_assignment(db_path, *, username="learner1"):
    version_ids = _publish_questions(db_path)
    template = exams.create_template(
        db_path, unit_code="son_tay", actor="admin", code="TPL-MY",
        title="My Assignments Test", target_audience_code="nvkt",
        question_version_ids=version_ids, duration_seconds=600,
        pass_score_percent=80.0,
    )
    now = time_policy.utc_now_ms()
    exam = exams.create_exam(
        db_path, unit_code="son_tay", actor="admin", code="EXAM-MY",
        title="Kỳ thi thử", template_id=template["id"],
        target_audience_code="nvkt", start_at_ms=now,
        end_at_ms=now + 3_600_000, duration_seconds=600,
        pass_score_percent=80.0,
    )
    exams.create_assignments(
        db_path, unit_code="son_tay", actor="admin", exam_id=exam["id"],
        users=[{"username": username}], audience_code="nvkt",
    )
    exams.ready_exam(db_path, unit_code="son_tay", actor="admin", exam_id=exam["id"])
    exams.open_exam(db_path, unit_code="son_tay", actor="admin", exam_id=exam["id"])
    return exam


def _client(db_path, monkeypatch, username="learner1"):
    app = flask.Flask(__name__)
    app.config["SECRET_KEY"] = "test"
    app.register_blueprint(training_bp)
    monkeypatch.setattr("config.TRAINING_DB_PATH", db_path)
    monkeypatch.setattr("config.UNIT_CODE", "son_tay")
    client = app.test_client()
    with client.session_transaction() as sess:
        sess["username"] = username
        sess["_csrf_token"] = "csrf"
    return client


def test_my_assignments_returns_only_own(tmp_path, monkeypatch):
    db_path = _setup_db(tmp_path)
    _make_exam_with_assignment(db_path, username="learner1")

    learner2 = _client(db_path, monkeypatch, username="learner2")
    resp = learner2.get("/api/training/my-assignments", headers={"X-CSRF-Token": "csrf"})
    assert resp.status_code == 200
    assert resp.get_json()["items"] == []

    mine = _client(db_path, monkeypatch, username="learner1")
    resp = mine.get("/api/training/my-assignments", headers={"X-CSRF-Token": "csrf"})
    assert resp.status_code == 200
    body = resp.get_json()
    assert len(body["items"]) == 1
    assert body["items"][0]["exam_code"] == "EXAM-MY"


def test_my_assignments_dto_hides_secrets(tmp_path, monkeypatch):
    db_path = _setup_db(tmp_path)
    _make_exam_with_assignment(db_path)
    client = _client(db_path, monkeypatch)
    resp = client.get("/api/training/my-assignments", headers={"X-CSRF-Token": "csrf"})
    body = resp.get_json()
    item = body["items"][0]
    hidden_keys = {
        "correct_option_ids", "explanation", "evidence", "distractor_rationales",
        "scoring_policy", "password", "team_code", "organization_code",
    }
    assert hidden_keys.isdisjoint(item.keys())


def test_my_assignments_empty_for_no_assignments(tmp_path, monkeypatch):
    db_path = _setup_db(tmp_path)
    client = _client(db_path, monkeypatch)
    resp = client.get("/api/training/my-assignments", headers={"X-CSRF-Token": "csrf"})
    assert resp.status_code == 200
    assert resp.get_json()["items"] == []
