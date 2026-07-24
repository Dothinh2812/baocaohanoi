from pathlib import Path

from flask import Flask

from blueprints.auth_routes import auth_bp
from blueprints.training_routes import training_bp
from services.training_catalog_service import grant_role, seed_defaults
from training import migrations


def _workspace_client(monkeypatch, tmp_path, *, username, roles=(), dashboard_role=None):
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
        lambda candidate: {
            "username": candidate,
            "fullname": candidate,
            "role": dashboard_role or "user",
        } if candidate == username else None,
    )

    root = Path(__file__).resolve().parents[1]
    app = Flask(__name__, template_folder=str(root / "templates"), static_folder=str(root / "static"))
    app.config["SECRET_KEY"] = "test"
    app.jinja_env.globals["is_endpoint_enabled"] = lambda endpoint: endpoint == "training.page_index"
    app.url_build_error_handlers.append(lambda error, endpoint, values: "#")
    app.register_blueprint(auth_bp)
    app.register_blueprint(training_bp)
    with app.test_client() as client:
        with client.session_transaction() as sess:
            sess["username"] = username
        yield client


def test_workspace_shows_only_learner_navigation(monkeypatch, tmp_path):
    for client in _workspace_client(monkeypatch, tmp_path, username="learner", roles=("learner",)):
        response = client.get("/dao-tao-sat-hach")

    assert response.status_code == 200
    for label in ("Tổng quan", "Bài thi của tôi", "Kết quả của tôi"):
        assert label.encode("utf-8") in response.data
    for label in ("Kho tri thức", "Ngân hàng câu hỏi", "Mẫu đề", "Kỳ thi", "Báo cáo"):
        assert label.encode("utf-8") not in response.data


def test_workspace_shows_editor_and_exam_manager_navigation(monkeypatch, tmp_path):
    for client in _workspace_client(
        monkeypatch,
        tmp_path,
        username="operator",
        roles=("editor", "exam_manager"),
    ):
        response = client.get("/dao-tao-sat-hach")

    assert response.status_code == 200
    for label in ("Tổng quan", "Kho tri thức", "Ngân hàng câu hỏi", "Mẫu đề", "Kỳ thi", "Báo cáo"):
        assert label.encode("utf-8") in response.data
    assert b"B\xc3\xa0i thi c\xe1\xbb\xa7a t\xc3\xb4i" not in response.data
    assert b"K\xe1\xba\xbft qu\xe1\xba\xa3 c\xe1\bb\xa7a t\xc3\xb4i" not in response.data


def test_workspace_shows_exam_manager_assignment_navigation(monkeypatch, tmp_path):
    for client in _workspace_client(
        monkeypatch,
        tmp_path,
        username="manager",
        roles=("exam_manager",),
    ):
        response = client.get("/dao-tao-sat-hach")

    assert response.status_code == 200
    assert "Giao bài".encode("utf-8") in response.data


def test_workspace_dashboard_admin_receives_all_module_navigation(monkeypatch, tmp_path):
    for client in _workspace_client(
        monkeypatch,
        tmp_path,
        username="dashboard-admin",
        dashboard_role="admin",
    ):
        response = client.get("/dao-tao-sat-hach")

    assert response.status_code == 200
    for label in (
        "Tổng quan", "Kho tri thức", "Ngân hàng câu hỏi", "Mẫu đề", "Kỳ thi",
        "Bài thi của tôi", "Kết quả của tôi", "Báo cáo",
    ):
        assert label.encode("utf-8") in response.data
    assert b"S\xe1\xba\xafp tri\xe1\xbb\x83n khai" in response.data


def test_workspace_requires_authenticated_dashboard_session():
    import dashboard

    dashboard.app.config["TESTING"] = True
    with dashboard.app.test_client() as client:
        response = client.get("/dao-tao-sat-hach")

    assert response.status_code == 302
    assert "/login" in response.headers["Location"]


def test_workspace_references_shared_client_helper(monkeypatch, tmp_path):
    for client in _workspace_client(monkeypatch, tmp_path, username="learner", roles=("learner",)):
        response = client.get("/dao-tao-sat-hach")

    assert response.status_code == 200
    assert b"js/training-ui.js" in response.data


def test_workspace_question_bank_panel_is_only_rendered_for_operators(monkeypatch, tmp_path):
    for editor in _workspace_client(monkeypatch, tmp_path, username="editor", roles=("editor",)):
        editor_page = editor.get("/dao-tao-sat-hach").get_data(as_text=True)
    for learner in _workspace_client(monkeypatch, tmp_path, username="learner", roles=("learner",)):
        learner_page = learner.get("/dao-tao-sat-hach").get_data(as_text=True)

    assert 'id="training-question-bank"' in editor_page
    assert 'id="training-question-bank"' not in learner_page
    assert "js/training-question-bank.js" in editor_page


def test_question_bank_client_renders_full_detail_and_confirms_all_review_actions():
    root = Path(__file__).resolve().parents[1]
    script = (root / "static" / "js" / "training-question-bank.js").read_text(encoding="utf-8")

    for field in (
        "question.stimulus", "question.language", "question.max_score", "question.scoring_policy",
        "question.created_by", "question.created_at_ms", "question.publication.approved_by",
        "question.publication.approved_at_ms",
    ):
        assert field in script
    for prompt in ("Duyệt câu hỏi này?", "Từ chối câu hỏi này?", "Phát hành câu hỏi này?"):
        assert prompt in script


def test_workspace_renders_template_and_exam_panels_for_exam_manager(monkeypatch, tmp_path):
    for client in _workspace_client(
        monkeypatch, tmp_path, username="manager", roles=("exam_manager",),
    ):
        page = client.get("/dao-tao-sat-hach").get_data(as_text=True)

    assert 'id="training-templates"' in page
    assert 'id="training-exams"' in page
    assert "js/training-templates.js" in page
    assert "js/training-exams.js" in page


def test_workspace_hides_template_and_exam_panels_for_learners(monkeypatch, tmp_path):
    for client in _workspace_client(monkeypatch, tmp_path, username="learner", roles=("learner",)):
        page = client.get("/dao-tao-sat-hach").get_data(as_text=True)

    assert 'id="training-templates"' not in page
    assert 'id="training-exams"' not in page
    assert "js/training-templates.js" not in page
    assert "js/training-exams.js" not in page


def test_workspace_nav_points_template_and_exam_panels_to_real_ids(monkeypatch, tmp_path):
    for client in _workspace_client(
        monkeypatch, tmp_path, username="manager", roles=("exam_manager",),
    ):
        page = client.get("/dao-tao-sat-hach").get_data(as_text=True)

    assert 'href="#training-templates"' in page
    assert 'data-panel="training-templates"' in page
    assert 'href="#training-exams"' in page
    assert 'data-panel="training-exams"' in page
    # Mẫu đề nav must no longer fall back to the shared placeholder.
    templates_nav = page.split('Mẫu đề')[0].rsplit('<a', 1)[-1]
    assert "#training-pending" not in templates_nav
