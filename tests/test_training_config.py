import importlib
import os

import config


def _reload_config(monkeypatch, **env):
    for key, value in env.items():
        monkeypatch.setenv(key, value)
    return importlib.reload(config)


def test_training_db_default_under_instance_runtime_dir(monkeypatch):
    monkeypatch.setenv("DASHV4_UNIT_CODE", "son_tay")
    cfg = _reload_config(monkeypatch)
    try:
        assert cfg.TRAINING_DB_PATH.endswith(os.path.join("son_tay", "training.db"))
    finally:
        importlib.reload(config)


def test_training_db_env_override(monkeypatch, tmp_path):
    custom = str(tmp_path / "custom_training.db")
    cfg = _reload_config(monkeypatch, DASHV4_TRAINING_DB_PATH=custom)
    try:
        assert cfg.TRAINING_DB_PATH == custom
    finally:
        importlib.reload(config)


def test_training_files_and_export_dirs_default(monkeypatch):
    monkeypatch.setenv("DASHV4_UNIT_CODE", "unit_x")
    cfg = _reload_config(monkeypatch)
    try:
        assert cfg.TRAINING_FILES_DIR.endswith(os.path.join("unit_x", "training_files"))
        assert cfg.TRAINING_EXPORT_DIR.endswith(os.path.join("unit_x", "training_export"))
    finally:
        importlib.reload(config)


def test_ai_disabled_by_default(monkeypatch):
    cfg = _reload_config(monkeypatch)
    try:
        assert cfg.TRAINING_AI_ENABLED is False
    finally:
        importlib.reload(config)


def test_ai_config_defaults(monkeypatch):
    cfg = _reload_config(monkeypatch)
    try:
        assert cfg.TRAINING_AI_PROVIDER == "fake"
        assert cfg.TRAINING_AI_MODEL == ""
        assert cfg.TRAINING_MAX_TEXT_BYTES > 0
        assert cfg.TRAINING_GENERATION_TIMEOUT_SECONDS > 0
        assert cfg.TRAINING_LEASE_SECONDS > 0
        assert cfg.TRAINING_MAX_RETRIES >= 1
    finally:
        importlib.reload(config)


def test_ai_enabled_via_env(monkeypatch):
    cfg = _reload_config(monkeypatch, DASHV4_TRAINING_AI_ENABLED="true")
    try:
        assert cfg.TRAINING_AI_ENABLED is True
    finally:
        importlib.reload(config)


def test_dashboard_config_exposes_training_attrs(monkeypatch):
    monkeypatch.setenv("DASHV4_UNIT_CODE", "son_tay")
    cfg = _reload_config(monkeypatch)
    try:
        assert cfg.DashboardConfig.TRAINING_DB_PATH == cfg.TRAINING_DB_PATH
        assert cfg.DashboardConfig.TRAINING_AI_ENABLED is False
        assert cfg.DashboardConfig.TRAINING_AI_PROVIDER == "fake"
    finally:
        importlib.reload(config)
