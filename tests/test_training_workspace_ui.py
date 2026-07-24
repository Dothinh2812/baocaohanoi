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
