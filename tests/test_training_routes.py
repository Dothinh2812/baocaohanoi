from flask import Flask

from app_helpers import configure_app
from blueprints.training_routes import training_bp
from training import db as training_db
from training import migrations
from services import training_attempt_service as attempts
from services import training_exam_service as exams
from services.training_catalog_service import grant_role, seed_defaults
from tests.test_training_attempt_lifecycle import _make_open_exam_with_assignment
from tests.test_training_exam_states import _publish_questions
from training import time_policy


def _client(monkeypatch, tmp_path):
    from blueprints import training_routes

    db_path = str(tmp_path / "training.db")
    migrations.run_migrations(db_path, "son_tay")
    seed_defaults(db_path, "son_tay")
    grant_role(db_path, "son_tay", "seed", "admin", "exam_manager")
    grant_role(db_path, "son_tay", "seed", "admin", "editor")
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
    grant_role(db_path, "son_tay", "seed", "learner1", "learner")
    grant_role(db_path, "son_tay", "seed", "learner2", "learner")
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


def _role_client(monkeypatch, tmp_path, *, username, roles=(), dashboard_role=None):
    from blueprints import training_routes

    db_path = str(tmp_path / "training.db")
    migrations.run_migrations(db_path, "son_tay")
    seed_defaults(db_path, "son_tay")
    for role in roles:
        grant_role(db_path, "son_tay", "seed", username, role)
    monkeypatch.setattr(training_routes.config, "TRAINING_DB_PATH", db_path)
    monkeypatch.setattr(training_routes.config, "UNIT_CODE", "son_tay")
    monkeypatch.setattr(
        training_routes,
        "get_user_by_username",
        lambda candidate: {"role": dashboard_role} if candidate == username and dashboard_role else None,
    )

    app = Flask(__name__)
    app.config["SECRET_KEY"] = "test"
    app.register_blueprint(training_bp)
    with app.test_client() as client:
        with client.session_transaction() as sess:
            sess["username"] = username
            sess["_csrf_token"] = "csrf"
        yield client, db_path


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


def test_knowledge_create_requires_editor_role(monkeypatch, tmp_path):
    for client, _ in _role_client(monkeypatch, tmp_path, username="unassigned"):
        response = client.post(
            "/api/training/knowledge",
            headers={"X-CSRF-Token": "csrf"},
            json={"document_code": "C1", "title": "Tài liệu", "content_text": "Nội dung"},
        )

    assert response.status_code == 403
    assert response.get_json()["error"]["code"] == "PERMISSION_SCOPE_DENIED"


def test_editor_can_create_knowledge(monkeypatch, tmp_path):
    for client, _ in _role_client(monkeypatch, tmp_path, username="editor", roles=("editor",)):
        response = client.post(
            "/api/training/knowledge",
            headers={"X-CSRF-Token": "csrf"},
            json={"document_code": "C1", "title": "Tài liệu", "content_text": "Nội dung"},
        )

    assert response.status_code == 201


def test_editor_cannot_publish_question(monkeypatch, tmp_path):
    for client, _ in _role_client(monkeypatch, tmp_path, username="editor", roles=("editor",)):
        response = client.post(
            "/api/training/questions/question-1/publish",
            headers={"X-CSRF-Token": "csrf"},
        )

    assert response.status_code == 403
    assert response.get_json()["error"]["code"] == "PERMISSION_SCOPE_DENIED"


def test_manager_cannot_load_owned_attempt_without_learner_role(monkeypatch, tmp_path):
    for client, db_path in _role_client(monkeypatch, tmp_path, username="manager", roles=("exam_manager",)):
        exam_id, _ = _make_open_exam_with_assignment(db_path)
        assignment_id = exams.create_assignments(
            db_path, unit_code="son_tay", actor="mgr", exam_id=exam_id,
            users=[{"username": "manager", "display_name": "Manager"}], audience_code="nvkt",
        )[0]
        attempt = attempts.start_attempt(
            db_path, unit_code="son_tay", actor="manager", assignment_id=assignment_id,
        )
        response = client.get(f"/api/training/attempts/{attempt['attempt_id']}")

    assert response.status_code == 403
    assert response.get_json()["error"]["code"] == "PERMISSION_SCOPE_DENIED"


def test_attempt_routes_require_learner_role_before_ownership(monkeypatch, tmp_path):
    for client, db_path in _role_client(monkeypatch, tmp_path, username="manager", roles=("exam_manager",)):
        exam_id, _ = _make_open_exam_with_assignment(db_path)
        assignment_id = exams.create_assignments(
            db_path, unit_code="son_tay", actor="mgr", exam_id=exam_id,
            users=[{"username": "manager", "display_name": "Manager"}], audience_code="nvkt",
        )[0]
        attempt = attempts.start_attempt(
            db_path, unit_code="son_tay", actor="manager", assignment_id=assignment_id,
        )
        item_id = attempts.get_attempt_learner_view(db_path, attempt["attempt_id"])["items"][0]["item_id"]
        responses = [
            client.post(f"/api/training/assignments/{assignment_id}/attempts", headers={"X-CSRF-Token": "csrf"}),
            client.get(f"/api/training/attempts/{attempt['attempt_id']}"),
            client.put(
                f"/api/training/attempts/{attempt['attempt_id']}/responses/{item_id}",
                headers={"X-CSRF-Token": "csrf"},
                json={"selected_option_ids": ["B"], "client_revision": 0},
            ),
            client.post(f"/api/training/attempts/{attempt['attempt_id']}/submit", headers={"X-CSRF-Token": "csrf"}),
        ]

    assert [response.status_code for response in responses] == [403, 403, 403, 403]
    assert all(response.get_json()["error"]["code"] == "PERMISSION_SCOPE_DENIED" for response in responses)


def test_learner_cannot_load_foreign_attempt(monkeypatch, tmp_path):
    for client, attempt_id, _, other_attempt_item_id in _learner_client_with_attempt(
        monkeypatch, tmp_path, second_attempt=True,
    ):
        # The second fixture creates learner2's attempt; recover its ID from its item.
        conn = training_db.read_connection(str(tmp_path / "training.db"))
        try:
            foreign_attempt_id = conn.execute(
                "SELECT attempt_id FROM exam_attempt_items WHERE id=?", (other_attempt_item_id,),
            ).fetchone()["attempt_id"]
        finally:
            conn.close()
        response = client.get(f"/api/training/attempts/{foreign_attempt_id}")

    assert response.status_code == 403
    assert response.get_json()["error"]["code"] == "PERMISSION_SCOPE_DENIED"


def test_dashboard_admin_bypasses_module_role_requirement(monkeypatch, tmp_path):
    for client, _ in _role_client(monkeypatch, tmp_path, username="dashboard-admin", dashboard_role="admin"):
        response = client.post(
            "/api/training/knowledge",
            headers={"X-CSRF-Token": "csrf"},
            json={"document_code": "C1", "title": "Tài liệu", "content_text": "Nội dung"},
        )

    assert response.status_code == 201


def test_create_exam_route_returns_stable_error_for_missing_template(monkeypatch, tmp_path):
    for client in _client(monkeypatch, tmp_path):
        now = time_policy.utc_now_ms()
        response = client.post(
            "/api/training/exams",
            headers={"X-CSRF-Token": "csrf"},
            json={
                "code": "EXAM-MISSING-TEMPLATE", "title": "Kỳ thi", "template_id": "missing",
                "target_audience_code": "nvkt", "start_at_ms": now,
                "end_at_ms": now + 3_600_000, "duration_seconds": 600,
            },
        )

    assert response.status_code == 404
    assert response.get_json()["error"]["code"] == "NOT_FOUND"


def test_create_exam_route_returns_audience_mismatch(monkeypatch, tmp_path):
    for client in _client(monkeypatch, tmp_path):
        db_path = str(tmp_path / "training.db")
        version_ids = _publish_questions(db_path)
        template = exams.create_template(
            db_path, unit_code="son_tay", actor="admin", code="TPL-ROUTE-AUDIENCE", title="Template",
            target_audience_code="nvkt", question_version_ids=version_ids,
            duration_seconds=600, pass_score_percent=80.0,
        )
        now = time_policy.utc_now_ms()
        response = client.post(
            "/api/training/exams",
            headers={"X-CSRF-Token": "csrf"},
            json={
                "code": "EXAM-ROUTE-AUDIENCE", "title": "Kỳ thi", "template_id": template["id"],
                "target_audience_code": "b2a", "start_at_ms": now,
                "end_at_ms": now + 3_600_000, "duration_seconds": 600,
            },
        )

    assert response.status_code == 400
    assert response.get_json()["error"]["code"] == "AUDIENCE_MISMATCH"


def test_create_assignments_route_returns_audience_mismatch(monkeypatch, tmp_path):
    for client in _client(monkeypatch, tmp_path):
        db_path = str(tmp_path / "training.db")
        version_ids = _publish_questions(db_path)
        template = exams.create_template(
            db_path, unit_code="son_tay", actor="admin", code="TPL-ASSIGN-ROUTE", title="Template",
            target_audience_code="nvkt", question_version_ids=version_ids,
            duration_seconds=600, pass_score_percent=80.0,
        )
        now = time_policy.utc_now_ms()
        exam = exams.create_exam(
            db_path, unit_code="son_tay", actor="admin", code="EXAM-ASSIGN-ROUTE", title="Kỳ thi",
            template_id=template["id"], target_audience_code="nvkt", start_at_ms=now,
            end_at_ms=now + 3_600_000, duration_seconds=600, pass_score_percent=80.0,
        )
        response = client.post(
            f"/api/training/exams/{exam['id']}/assignments",
            headers={"X-CSRF-Token": "csrf"},
            json={"users": [{"username": "u1"}], "audience_code": "b2a"},
        )

    assert response.status_code == 400
    assert response.get_json()["error"]["code"] == "AUDIENCE_MISMATCH"
