from flask import Flask

from app_helpers import configure_app
from blueprints.training_routes import training_bp
from training import db as training_db
from training import migrations
from services import training_attempt_service as attempts
from services import training_exam_service as exams
from services.training_catalog_service import grant_role, seed_defaults
from tests.test_training_attempt_lifecycle import _make_open_exam_with_assignment


def _client(monkeypatch, tmp_path):
    from blueprints import training_routes

    db_path = str(tmp_path / "training.db")
    migrations.run_migrations(db_path, "son_tay")
    seed_defaults(db_path, "son_tay")
    grant_role(db_path, "son_tay", "seed", "admin", "exam_manager")
    monkeypatch.setattr(training_routes.config, "TRAINING_DB_PATH", db_path)
    monkeypatch.setattr(training_routes.config, "UNIT_CODE", "son_tay")

    app = Flask(__name__)
    app.config["SECRET_KEY"] = "test"
    app.register_blueprint(training_bp)
    with app.test_client() as client:
        with client.session_transaction() as sess:
            sess["username"] = "admin"
            sess["_csrf_token"] = "csrf"
        yield client


def _learner_client_with_attempt(monkeypatch, tmp_path, *, second_attempt=False):
    from blueprints import training_routes

    db_path = str(tmp_path / "training.db")
    migrations.run_migrations(db_path, "son_tay")
    seed_defaults(db_path, "son_tay")
    monkeypatch.setattr(training_routes.config, "TRAINING_DB_PATH", db_path)
    monkeypatch.setattr(training_routes.config, "UNIT_CODE", "son_tay")
    exam_id, assignment_id = _make_open_exam_with_assignment(db_path)
    attempt = attempts.start_attempt(
        db_path, unit_code="son_tay", actor="learner1", assignment_id=assignment_id,
    )
    item_id = attempts.get_attempt_learner_view(db_path, attempt["attempt_id"])["items"][0]["item_id"]
    other_item_id = None
    if second_attempt:
        second_assignment_id = exams.create_assignments(
            db_path, unit_code="son_tay", actor="mgr", exam_id=exam_id,
            users=[{"username": "learner2", "display_name": "Learner 2"}], audience_code="nvkt",
        )[0]
        second_attempt_result = attempts.start_attempt(
            db_path, unit_code="son_tay", actor="learner2", assignment_id=second_assignment_id,
        )
        other_item_id = attempts.get_attempt_learner_view(
            db_path, second_attempt_result["attempt_id"],
        )["items"][0]["item_id"]

    app = Flask(__name__)
    app.config["SECRET_KEY"] = "test"
    app.register_blueprint(training_bp)
    with app.test_client() as client:
        with client.session_transaction() as sess:
            sess["username"] = "learner1"
            sess["_csrf_token"] = "csrf"
        yield client, attempt["attempt_id"], item_id, other_item_id


def test_operator_can_create_paste_text_knowledge(monkeypatch, tmp_path):
    for client in _client(monkeypatch, tmp_path):
        response = client.post(
            "/api/training/knowledge",
            headers={"X-CSRF-Token": "csrf"},
            json={
                "document_code": "C1.1",
                "title": "Chất lượng sửa chữa",
                "content_text": "Quy định test.",
                "classification": {"domain_code": "quality"},
                "audience_codes": ["nvkt"],
            },
        )

    assert response.status_code == 201
    assert response.get_json()["version_id"]


def test_operator_can_import_question_batch_as_draft(monkeypatch, tmp_path):
    for client in _client(monkeypatch, tmp_path):
        response = client.post(
            "/api/training/questions/import",
            headers={"X-CSRF-Token": "csrf"},
            json={
                "schema_version": "1.0",
                "batch": {
                    "title": "Draft", "language": "vi",
                    "source_document_version_ids": ["docver-001"],
                    "target_audience_codes": ["nvkt"], "requested_count": 1,
                },
                "questions": [{
                    "local_ref": "Q1", "type": "single_choice", "stem": "Câu test?",
                    "options": [{"id": "A", "text": "Sai"}, {"id": "B", "text": "Đúng"}],
                    "correct_option_ids": ["B"], "classification": {"domain_code": "quality"},
                    "difficulty": "easy",
                    "evidence": [{"document_version_id": "docver-001", "block_id": "DOC-B001", "extraction_revision": 1, "supports": "correct_answer"}],
                }],
            },
        )

    assert response.status_code == 201
    assert len(response.get_json()["version_ids"]) == 1


def test_autosave_route_rejects_item_from_another_attempt(monkeypatch, tmp_path):
    for client, attempt_id, _, foreign_item_id in _learner_client_with_attempt(
        monkeypatch, tmp_path, second_attempt=True,
    ):
        response = client.put(
            f"/api/training/attempts/{attempt_id}/responses/{foreign_item_id}",
            headers={"X-CSRF-Token": "csrf"},
            json={"selected_option_ids": ["B"], "client_revision": 1},
        )

    assert response.status_code == 404
    assert response.get_json()["error"]["code"] == "ATTEMPT_ITEM_NOT_FOUND"


def test_autosave_route_rejects_invalid_option(monkeypatch, tmp_path):
    for client, attempt_id, item_id, _ in _learner_client_with_attempt(monkeypatch, tmp_path):
        response = client.put(
            f"/api/training/attempts/{attempt_id}/responses/{item_id}",
            headers={"X-CSRF-Token": "csrf"},
            json={"selected_option_ids": ["missing"], "client_revision": 1},
        )

    assert response.status_code == 400
    assert response.get_json()["error"]["code"] == "INVALID_OPTION_SELECTION"


def test_autosave_route_rejects_multiple_single_choice_options(monkeypatch, tmp_path):
    for client, attempt_id, item_id, _ in _learner_client_with_attempt(monkeypatch, tmp_path):
        response = client.put(
            f"/api/training/attempts/{attempt_id}/responses/{item_id}",
            headers={"X-CSRF-Token": "csrf"},
            json={"selected_option_ids": ["A", "B"], "client_revision": 1},
        )

    assert response.status_code == 400
    assert response.get_json()["error"]["code"] == "SINGLE_CHOICE_REQUIRES_ONE_OPTION"
