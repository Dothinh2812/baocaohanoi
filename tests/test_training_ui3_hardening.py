import threading

from tests.test_training_exam_states import VALID_BATCH, _publish_questions, _setup
from tests.test_training_routes import QUESTION_BATCH, _client
from training import db as training_db
from training.errors import TrainingError
from services import training_question_service as qs


def test_question_bank_detail_has_no_store_cache_header(monkeypatch, tmp_path):
    for client in _client(monkeypatch, tmp_path):
        imported = client.post(
            "/api/training/questions/import", headers={"X-CSRF-Token": "csrf"},
            json=QUESTION_BATCH,
        )
        version_id = imported.get_json()["version_ids"][0]
        response = client.get(f"/api/training/questions/{version_id}")

    assert response.status_code == 200
    assert response.headers["Cache-Control"] == "no-store"
