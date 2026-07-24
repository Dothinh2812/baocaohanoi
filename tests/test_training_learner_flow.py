import flask
import pytest

from training import db as training_db
from training import migrations, time_policy
from services import training_exam_service as exams
from services import training_attempt_service as attempts
from services import training_question_service as qs
from services.training_catalog_service import seed_defaults, grant_role
from blueprints import training_bp
from tests.test_training_exam_states import _publish_questions


def _setup_db(tmp_path):
    db_path = str(tmp_path / "training.db")
    migrations.run_migrations(db_path, "son_tay")
    seed_defaults(db_path, "son_tay")
    grant_role(db_path, "son_tay", "seed", "admin", "exam_manager")
    grant_role(db_path, "son_tay", "seed", "admin", "learner")
    grant_role(db_path, "son_tay", "seed", "learner1", "learner")
    grant_role(db_path, "son_tay", "seed", "learner2", "learner")
    return db_path


def _setup_learner_exam(db_path, *, learner="learner1"):
    version_ids = _publish_questions(db_path)
    template = exams.create_template(
        db_path, unit_code="son_tay", actor="admin", code="TPL-LEARN",
        title="Learner Flow Test", target_audience_code="nvkt",
        question_version_ids=version_ids, duration_seconds=600, pass_score_percent=80.0,
    )
    now = time_policy.utc_now_ms()
    exam = exams.create_exam(
        db_path, unit_code="son_tay", actor="admin", code="EXAM-LEARN",
        title="Kỳ thi learner", template_id=template["id"],
        target_audience_code="nvkt", start_at_ms=now,
        end_at_ms=now + 3_600_000, duration_seconds=600, pass_score_percent=80.0,
    )
    exams.create_assignments(
        db_path, unit_code="son_tay", actor="admin", exam_id=exam["id"],
        users=[{"username": learner}], audience_code="nvkt",
    )
    exams.ready_exam(db_path, unit_code="son_tay", actor="admin", exam_id=exam["id"])
    exams.open_exam(db_path, unit_code="son_tay", actor="admin", exam_id=exam["id"])
    assignment = exams.get_assignment_for_user(db_path, exam["id"], learner)[0]
    return exam, assignment


def _client(db_path, monkeypatch, username="learner1"):
    from blueprints import training_routes

    monkeypatch.setattr(training_routes.config, "TRAINING_DB_PATH", db_path)
    monkeypatch.setattr(training_routes.config, "UNIT_CODE", "son_tay")
    app = flask.Flask(__name__)
    app.config["SECRET_KEY"] = "test"
    app.register_blueprint(training_bp)
    client = app.test_client()
    with client.session_transaction() as sess:
        sess["username"] = username
        sess["_csrf_token"] = "csrf"
    return client


def _create_shared_exam(db_path):
    version_ids = _publish_questions(db_path)
    template = exams.create_template(
        db_path, unit_code="son_tay", actor="admin", code="TPL-OWN",
        title="Ownership Test", target_audience_code="nvkt",
        question_version_ids=version_ids, duration_seconds=600, pass_score_percent=80.0,
    )
    now = time_policy.utc_now_ms()
    exam = exams.create_exam(
        db_path, unit_code="son_tay", actor="admin", code="EXAM-OWN",
        title="Ownership", template_id=template["id"],
        target_audience_code="nvkt", start_at_ms=now,
        end_at_ms=now + 3_600_000, duration_seconds=600, pass_score_percent=80.0,
    )
    exams.create_assignments(
        db_path, unit_code="son_tay", actor="admin", exam_id=exam["id"],
        users=[{"username": "learner1"}, {"username": "learner2"}], audience_code="nvkt",
    )
    exams.ready_exam(db_path, unit_code="son_tay", actor="admin", exam_id=exam["id"])
    exams.open_exam(db_path, unit_code="son_tay", actor="admin", exam_id=exam["id"])
    return exam


class TestLearnerOwnership:
    def test_learner_a_cannot_start_learner_b_attempt(self, tmp_path, monkeypatch):
        db_path = _setup_db(tmp_path)
        exam = _create_shared_exam(db_path)
        a1 = exams.get_assignment_for_user(db_path, exam["id"], "learner1")[0]

        client_b = _client(db_path, monkeypatch, username="learner2")
        resp = client_b.post(
            f"/api/training/assignments/{a1['id']}/attempts",
            headers={"X-CSRF-Token": "csrf"},
        )
        assert resp.status_code in (403, 404)

    def test_learner_attempt_dto_hides_answers(self, tmp_path, monkeypatch):
        db_path = _setup_db(tmp_path)
        exam, assignment = _setup_learner_exam(db_path, learner="learner1")
        client = _client(db_path, monkeypatch, username="learner1")

        start_resp = client.post(
            f"/api/training/assignments/{assignment['id']}/attempts",
            headers={"X-CSRF-Token": "csrf"},
        )
        assert start_resp.status_code == 200
        attempt_id = start_resp.get_json()["attempt_id"]

        get_resp = client.get(f"/api/training/attempts/{attempt_id}")
        assert get_resp.status_code == 200
        body = get_resp.get_json()
        for item in body["items"]:
            assert "correct_option_ids" not in item
            assert "explanation" not in item
            assert "distractor_rationales" not in item
            assert "evidence" not in item
            assert "scoring_policy" not in item
            assert "max_score" not in item
        assert "no-store" in get_resp.headers.get("Cache-Control", "")

    def test_autosave_revision_cas(self, tmp_path, monkeypatch):
        db_path = _setup_db(tmp_path)
        exam, assignment = _setup_learner_exam(db_path, learner="learner1")
        client = _client(db_path, monkeypatch, username="learner1")

        start_resp = client.post(
            f"/api/training/assignments/{assignment['id']}/attempts",
            headers={"X-CSRF-Token": "csrf"},
        )
        attempt_id = start_resp.get_json()["attempt_id"]

        get_resp = client.get(f"/api/training/attempts/{attempt_id}")
        item_id = get_resp.get_json()["items"][0]["item_id"]

        resp5 = client.put(
            f"/api/training/attempts/{attempt_id}/responses/{item_id}",
            json={"selected_option_ids": ["A"], "client_revision": 5},
            headers={"X-CSRF-Token": "csrf"},
        )
        assert resp5.status_code == 200
        assert resp5.get_json()["accepted"] is True

        resp3 = client.put(
            f"/api/training/attempts/{attempt_id}/responses/{item_id}",
            json={"selected_option_ids": ["B"], "client_revision": 3},
            headers={"X-CSRF-Token": "csrf"},
        )
        assert resp3.status_code == 200
        body = resp3.get_json()
        assert body["accepted"] is False
        assert body["stored_revision"] == 5

    def test_submit_retry_idempotent(self, tmp_path, monkeypatch):
        db_path = _setup_db(tmp_path)
        exam, assignment = _setup_learner_exam(db_path, learner="learner1")
        client = _client(db_path, monkeypatch, username="learner1")

        start_resp = client.post(
            f"/api/training/assignments/{assignment['id']}/attempts",
            headers={"X-CSRF-Token": "csrf"},
        )
        attempt_id = start_resp.get_json()["attempt_id"]

        get_resp = client.get(f"/api/training/attempts/{attempt_id}")
        for item in get_resp.get_json()["items"]:
            client.put(
                f"/api/training/attempts/{attempt_id}/responses/{item['item_id']}",
                json={"selected_option_ids": ["B"], "client_revision": 1},
                headers={"X-CSRF-Token": "csrf"},
            )

        s1 = client.post(
            f"/api/training/attempts/{attempt_id}/submit",
            headers={"X-CSRF-Token": "csrf"},
        )
        assert s1.status_code == 200
        score1 = s1.get_json()["result"]["score"]

        s2 = client.post(
            f"/api/training/attempts/{attempt_id}/submit",
            headers={"X-CSRF-Token": "csrf"},
        )
        assert s2.status_code == 200
        score2 = s2.get_json()["result"]["score"]

        assert score1 == score2
