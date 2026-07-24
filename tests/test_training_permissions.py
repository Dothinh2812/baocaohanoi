import pytest

from training import db as training_db
from training import migrations
from training.permissions import (
    has_module_role,
    require_module_role,
    user_audience_codes,
    PERMISSION_SCOPE_DENIED,
)
from training.errors import TrainingError, ErrorCode


def _setup(monkeypatch, tmp_path, unit_code="son_tay"):
    db_path = str(tmp_path / "training.db")
    monkeypatch.setattr(training_db, "TRAINING_DB_PATH", db_path)
    monkeypatch.setattr(training_db, "UNIT_CODE", unit_code)
    migrations.run_migrations(db_path, unit_code)
    return db_path


def _grant_role(db_path, username, role):
    conn = training_db.write_connection(db_path)
    try:
        conn.execute(
            "INSERT INTO training_user_roles (id, username, role_code, granted_at_ms) "
            "VALUES (?, ?, ?, 0)",
            (f"r-{username}-{role}", username, role),
        )
        conn.commit()
    finally:
        conn.close()


def _grant_audience(db_path, username, audience):
    conn = training_db.write_connection(db_path)
    try:
        conn.execute(
            "INSERT INTO training_user_audiences (id, username, audience_code) VALUES (?, ?, ?)",
            (f"a-{username}-{audience}", username, audience),
        )
        conn.commit()
    finally:
        conn.close()


def test_catalog_tables_created(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    conn = training_db.read_connection(db_path)
    tables = {row["name"] for row in conn.execute(
        "SELECT name FROM sqlite_master WHERE type='table'"
    ).fetchall()}
    conn.close()
    assert "training_domains" in tables
    assert "training_categories" in tables
    assert "training_audiences" in tables
    assert "training_user_roles" in tables
    assert "training_user_audiences" in tables


def test_has_module_role_true_when_granted(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    _grant_role(db_path, "alice", "exam_manager")
    assert has_module_role(db_path, "alice", "exam_manager") is True


def test_has_module_role_false_when_not_granted(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    assert has_module_role(db_path, "bob", "exam_manager") is False


def test_user_can_have_multiple_roles(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    _grant_role(db_path, "carol", "learner")
    _grant_role(db_path, "carol", "editor")
    assert has_module_role(db_path, "carol", "learner") is True
    assert has_module_role(db_path, "carol", "editor") is True


def test_admin_role_satisfies_any_requirement(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    _grant_role(db_path, "dave", "admin")
    assert has_module_role(db_path, "dave", "exam_manager") is True
    assert has_module_role(db_path, "dave", "learner") is True


def test_user_audience_codes(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    _grant_audience(db_path, "eve", "nvkt")
    _grant_audience(db_path, "eve", "b2a")
    codes = user_audience_codes(db_path, "eve")
    assert set(codes) == {"nvkt", "b2a"}


def test_require_module_role_raises_when_denied(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    with pytest.raises(TrainingError) as exc_info:
        require_module_role(db_path, "frank", "exam_manager")
    assert exc_info.value.code == ErrorCode.PERMISSION_SCOPE_DENIED
    assert exc_info.value.status == 403


def test_require_module_role_passes_when_granted(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    _grant_role(db_path, "grace", "exam_manager")
    require_module_role(db_path, "grace", "exam_manager")


def test_duplicate_role_insert_is_rejected(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    _grant_role(db_path, "henry", "learner")
    with pytest.raises(Exception):
        _grant_role(db_path, "henry", "learner")
