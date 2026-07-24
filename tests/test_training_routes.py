from flask import Flask

from app_helpers import configure_app
from blueprints.training_routes import training_bp
from training import db as training_db
from training import migrations
from services.training_catalog_service import grant_role, seed_defaults


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
