import importlib
import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))


def _reload_module(name):
    if name in sys.modules:
        return importlib.reload(sys.modules[name])
    return importlib.import_module(name)


def test_config_uses_instance_scoped_runtime_paths(monkeypatch, tmp_path):
    session_dir = tmp_path / "sessions"
    cache_dir = tmp_path / "cache"
    log_dir = tmp_path / "logs"
    pid_file = tmp_path / "dashv4-ba_dinh.pid"
    user_file = tmp_path / "users.xlsx"
    login_log_file = log_dir / "login_history.csv"

    monkeypatch.setenv("DASHV4_UNIT_CODE", "ba_dinh")
    monkeypatch.setenv("DASHV4_SESSION_FILE_DIR", str(session_dir))
    monkeypatch.setenv("DASHV4_CACHE_DIR", str(cache_dir))
    monkeypatch.setenv("DASHV4_LOG_DIR", str(log_dir))
    monkeypatch.setenv("DASHV4_PID_FILE", str(pid_file))
    monkeypatch.setenv("DASHV4_USER_FILE", str(user_file))
    monkeypatch.setenv("DASHV4_LOGIN_LOG_FILE", str(login_log_file))

    config = _reload_module("config")

    assert config.UNIT_CODE == "ba_dinh"
    assert config.SESSION_FILE_DIR == str(session_dir)
    assert config.CACHE_DIR == str(cache_dir)
    assert config.LOG_DIR == str(log_dir)
    assert config.PID_FILE == str(pid_file)
    assert config.USER_FILE == str(user_file)
    assert config.LOGIN_LOG_FILE == str(login_log_file)
    assert config.DashboardConfig.SESSION_FILE_DIR == str(session_dir)
    assert config.DashboardConfig.LOG_DIR == str(log_dir)


def test_config_defaults_are_scoped_by_unit_code(monkeypatch):
    monkeypatch.setenv("DASHV4_UNIT_CODE", "cau_giay")
    for name in [
        "DASHV4_SESSION_FILE_DIR",
        "DASHV4_CACHE_DIR",
        "DASHV4_LOG_DIR",
        "DASHV4_PID_FILE",
        "DASHV4_USER_FILE",
        "DASHV4_LOGIN_LOG_FILE",
    ]:
        monkeypatch.delenv(name, raising=False)

    config = _reload_module("config")

    assert config.SESSION_FILE_DIR.endswith(os.path.join("runtime_app", "cau_giay", "flask_session"))
    assert config.CACHE_DIR.endswith(os.path.join("runtime_app", "cau_giay", "cache"))
    assert config.LOG_DIR.endswith(os.path.join("logs", "cau_giay"))
    assert config.PID_FILE == "/tmp/dashv4-cau_giay.pid"
    assert config.USER_FILE.endswith(os.path.join("runtime_app", "cau_giay", "users.xlsx"))
    assert config.LOGIN_LOG_FILE.endswith(os.path.join("logs", "cau_giay", "login_history.csv"))


def test_auth_uses_instance_user_and_login_log_files(monkeypatch, tmp_path):
    user_file = tmp_path / "users.xlsx"
    login_log_file = tmp_path / "unit_logs" / "login_history.csv"
    monkeypatch.setenv("DASHV4_USER_FILE", str(user_file))
    monkeypatch.setenv("DASHV4_LOGIN_LOG_FILE", str(login_log_file))

    auth = _reload_module("auth")

    assert Path(auth.EXCEL_FILE) == user_file
    assert Path(auth.LOGIN_LOG_FILE) == login_log_file
    assert login_log_file.exists()
    assert login_log_file.read_text(encoding="utf-8").splitlines()[0] == (
        "timestamp,username,ip_address,user_agent,action,status"
    )


def test_gunicorn_uses_instance_log_and_pid_paths(monkeypatch, tmp_path):
    log_dir = tmp_path / "unit_logs"
    pid_file = tmp_path / "dashv4-ba_dinh.pid"
    monkeypatch.setenv("DASHV4_UNIT_CODE", "ba_dinh")
    monkeypatch.setenv("DASHV4_LOG_DIR", str(log_dir))
    monkeypatch.setenv("DASHV4_PID_FILE", str(pid_file))

    _reload_module("config")
    gunicorn_config = _reload_module("gunicorn_config")

    assert gunicorn_config.proc_name == "dashv4_ba_dinh"
    assert gunicorn_config.pidfile == str(pid_file)
    assert gunicorn_config.accesslog == str(log_dir / "gunicorn_access.log")
    assert gunicorn_config.errorlog == str(log_dir / "gunicorn_error.log")
