import os
import sqlite3

import pytest

from training import db as training_db
from training import migrations


def _config(monkeypatch, tmp_path, unit_code="son_tay"):
    db_path = str(tmp_path / "training.db")
    monkeypatch.setattr(training_db, "TRAINING_DB_PATH", db_path)
    monkeypatch.setattr(training_db, "UNIT_CODE", unit_code)
    monkeypatch.setattr(training_db, "TRAINING_FILES_DIR", str(tmp_path / "files"))
    return db_path


def test_new_db_migrates_cleanly(monkeypatch, tmp_path):
    db_path = _config(monkeypatch, tmp_path)

    migrations.run_migrations(db_path, unit_code="son_tay")

    conn = training_db.read_connection(db_path)
    tables = {row["name"] for row in conn.execute(
        "SELECT name FROM sqlite_master WHERE type='table'"
    ).fetchall()}
    conn.close()
    assert "training_instance_metadata" in tables
    assert "training_schema_migrations" in tables
    assert "training_audit_log" in tables


def test_migrate_is_idempotent(monkeypatch, tmp_path):
    db_path = _config(monkeypatch, tmp_path)

    migrations.run_migrations(db_path, unit_code="son_tay")
    migrations.run_migrations(db_path, unit_code="son_tay")

    conn = training_db.read_connection(db_path)
    count = conn.execute(
        "SELECT COUNT(*) AS c FROM training_schema_migrations"
    ).fetchone()["c"]
    conn.close()
    assert count >= 1


def test_unit_code_mismatch_rejected(monkeypatch, tmp_path):
    db_path = _config(monkeypatch, tmp_path, unit_code="son_tay")
    migrations.run_migrations(db_path, unit_code="son_tay")

    with pytest.raises(migrations.UnitCodeMismatchError):
        migrations.run_migrations(db_path, unit_code="ba_dinh")


def test_instance_metadata_has_correct_unit_code(monkeypatch, tmp_path):
    db_path = _config(monkeypatch, tmp_path, unit_code="son_tay")
    migrations.run_migrations(db_path, unit_code="son_tay")

    conn = training_db.read_connection(db_path)
    row = conn.execute(
        "SELECT unit_code, schema_version FROM training_instance_metadata"
    ).fetchone()
    conn.close()
    assert row["unit_code"] == "son_tay"
    assert row["schema_version"] >= 8


def test_write_connection_pragmas(monkeypatch, tmp_path):
    db_path = _config(monkeypatch, tmp_path)
    migrations.run_migrations(db_path, unit_code="son_tay")

    conn = training_db.write_connection(db_path)
    journal = conn.execute("PRAGMA journal_mode").fetchone()[0]
    busy = conn.execute("PRAGMA busy_timeout").fetchone()[0]
    fk = conn.execute("PRAGMA foreign_keys").fetchone()[0]
    conn.close()
    assert journal.lower() == "wal"
    assert busy >= 5000
    assert fk == 1


def test_read_connection_has_row_factory(monkeypatch, tmp_path):
    db_path = _config(monkeypatch, tmp_path)
    migrations.run_migrations(db_path, unit_code="son_tay")

    conn = training_db.read_connection(db_path)
    assert conn.row_factory is sqlite3.Row
    conn.close()


def test_each_connection_is_independent(monkeypatch, tmp_path):
    db_path = _config(monkeypatch, tmp_path)
    migrations.run_migrations(db_path, unit_code="son_tay")

    c1 = training_db.write_connection(db_path)
    c2 = training_db.write_connection(db_path)
    assert c1 is not c2
    c1.close()
    c2.close()


def test_audit_log_index_exists(monkeypatch, tmp_path):
    db_path = _config(monkeypatch, tmp_path)
    migrations.run_migrations(db_path, unit_code="son_tay")

    conn = training_db.read_connection(db_path)
    indexes = {row["name"] for row in conn.execute(
        "SELECT name FROM sqlite_master WHERE type='index'"
    ).fetchall()}
    conn.close()
    assert "idx_audit_entity" in indexes


def test_migration_008_preserves_template_items_and_locks_question_versions(monkeypatch, tmp_path):
    db_path = _config(monkeypatch, tmp_path)
    migrations.run_migrations(db_path, unit_code="son_tay", migrations=migrations._MIGRATIONS[:7])
    conn = training_db.write_connection(db_path)
    try:
        conn.execute(
            "INSERT INTO question_items (id, created_at_ms) VALUES ('question', 0)"
        )
        conn.execute(
            """INSERT INTO question_versions
            (id, question_item_id, version_number, type, stem, correct_option_ids_json,
             difficulty, created_by, created_at_ms)
            VALUES ('version', 'question', 1, 'single_choice', 'Question', '[]', 'easy', 'alice', 0)"""
        )
        conn.execute(
            "INSERT INTO question_items (id, created_at_ms) VALUES ('question-2', 0)"
        )
        conn.execute(
            """INSERT INTO question_versions
            (id, question_item_id, version_number, type, stem, correct_option_ids_json,
             difficulty, created_by, created_at_ms)
            VALUES ('version-2', 'question-2', 1, 'single_choice', 'Question 2', '[]', 'easy', 'alice', 0)"""
        )
        conn.execute(
            """INSERT INTO exam_templates
            (id, code, title, target_audience_code, total_questions, duration_seconds,
             created_by, created_at_ms)
            VALUES ('template', 'TPL', 'Template', 'nvkt', 1, 60, 'alice', 0)"""
        )
        conn.execute(
            """INSERT INTO exam_template_items
            (id, template_id, sequence_number, question_version_id, points)
            VALUES ('item', 'template', 1, 'version', 1.0)"""
        )
        conn.execute(
            """INSERT INTO exam_template_items
            (id, template_id, sequence_number, question_version_id, points)
            VALUES ('duplicate-question', 'template', 3, 'version', 1.0)"""
        )
        conn.commit()
        conn.execute("PRAGMA foreign_keys=OFF")
        conn.execute(
            """INSERT INTO exam_template_items
            (id, template_id, sequence_number, question_version_id, points)
            VALUES ('dangling-version', 'template', 2, 'missing-version', 1.0)"""
        )
        conn.execute(
            """INSERT INTO exam_template_items
            (id, template_id, sequence_number, question_version_id, points)
            VALUES ('dangling-template', 'missing-template', 1, 'version-2', 1.0)"""
        )
        conn.commit()
        conn.execute("PRAGMA foreign_keys=ON")
    finally:
        conn.close()

    migrations.run_migrations(db_path, unit_code="son_tay")

    conn = training_db.write_connection(db_path)
    try:
        rows = conn.execute(
            "SELECT id, question_version_id FROM exam_template_items"
        ).fetchall()
        assert len(rows) == 1
        assert [(row["id"], row["question_version_id"]) for row in rows] == [("item", "version")]
        foreign_keys = conn.execute("PRAGMA foreign_key_list('exam_template_items')").fetchall()
        assert any(
            fk["table"] == "question_versions"
            and fk["from"] == "question_version_id"
            and fk["to"] == "id"
            for fk in foreign_keys
        )
        with pytest.raises(sqlite3.IntegrityError):
            conn.execute(
                """INSERT INTO exam_template_items
                (id, template_id, sequence_number, question_version_id, points)
                VALUES ('duplicate', 'template', 2, 'version', 1.0)"""
            )
        with pytest.raises(sqlite3.IntegrityError):
            conn.execute(
                """INSERT INTO exam_template_items
                (id, template_id, sequence_number, question_version_id, points)
                VALUES ('missing', 'template', 2, 'missing-version', 1.0)"""
            )
        with pytest.raises(sqlite3.IntegrityError):
            conn.execute(
                """INSERT INTO exam_template_items
                (id, template_id, sequence_number, question_version_id, points)
                VALUES ('missing-template', 'missing-template', 2, 'version-2', 1.0)"""
            )
        with pytest.raises(sqlite3.IntegrityError):
            conn.execute(
                """INSERT INTO exam_template_items
                (id, template_id, sequence_number, question_version_id, points)
                VALUES ('duplicate-sequence', 'template', 1, 'version-2', 1.0)"""
            )
        indexes = {row["name"] for row in conn.execute(
            "SELECT name FROM sqlite_master WHERE type='index'"
        ).fetchall()}
        assert "idx_titems_template_qv" in indexes
    finally:
        conn.close()
