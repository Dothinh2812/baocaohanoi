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
    assert migrations.CURRENT_SCHEMA_VERSION == 11
    assert migrations.get_schema_version(db_path) == migrations.CURRENT_SCHEMA_VERSION


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


def test_migration_009_backfills_job_unit_from_instance_metadata_with_audit(monkeypatch, tmp_path):
    db_path = _config(monkeypatch, tmp_path, unit_code="ba_vi")
    migrations.run_migrations(db_path, unit_code="ba_vi", migrations=migrations._MIGRATIONS[:8])
    conn = training_db.write_connection(db_path)
    try:
        conn.execute(
            """INSERT INTO ai_generation_jobs
            (id, status, request_payload_json, source_document_version_ids_json,
             target_audience_codes_json, requested_count, created_by, created_at_ms)
            VALUES ('job_legacy', 'pending', '{}', '[]', '[]', 1, 'alice', 0)"""
        )
        conn.commit()
    finally:
        conn.close()

    migrations.run_migrations(db_path, unit_code="ba_vi")

    conn = training_db.read_connection(db_path)
    try:
        job = conn.execute(
            "SELECT unit_code FROM ai_generation_jobs WHERE id='job_legacy'"
        ).fetchone()
        audit = conn.execute(
            "SELECT actor, unit_code FROM training_audit_log "
            "WHERE action='backfill_job_unit_code' AND entity_id='job_legacy'"
        ).fetchone()
    finally:
        conn.close()
    assert job["unit_code"] == "ba_vi"
    assert audit["actor"] == "migration"
    assert audit["unit_code"] == "ba_vi"


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


def test_failed_fresh_migration_leaves_no_schema(monkeypatch, tmp_path):
    db_path = _config(monkeypatch, tmp_path)

    def fail_migration(conn):
        raise RuntimeError("migration failed")

    with pytest.raises(RuntimeError, match="migration failed"):
        migrations.run_migrations(
            db_path,
            unit_code="son_tay",
            migrations=[(1, migrations.migration_001), (2, fail_migration)],
        )

    conn = training_db.write_connection(db_path)
    try:
        tables = {row["name"] for row in conn.execute(
            "SELECT name FROM sqlite_master WHERE type='table'"
        ).fetchall()}
    finally:
        conn.close()

    assert "training_schema_migrations" not in tables
    assert "training_audit_log" not in tables


def test_migration_008_rolls_back_when_recording_fails(monkeypatch, tmp_path):
    db_path = _config(monkeypatch, tmp_path)
    migrations.run_migrations(db_path, unit_code="son_tay", migrations=migrations._MIGRATIONS[:7])

    def fail_to_record(*args):
        raise RuntimeError("record migration failed")

    monkeypatch.setattr(migrations, "_record_migration", fail_to_record)
    with pytest.raises(RuntimeError, match="record migration failed"):
        migrations.run_migrations(db_path, unit_code="son_tay")

    conn = training_db.write_connection(db_path)
    try:
        versions = {row["version"] for row in conn.execute(
            "SELECT version FROM training_schema_migrations"
        ).fetchall()}
        foreign_keys = conn.execute("PRAGMA foreign_key_list('exam_template_items')").fetchall()
        schema_version = conn.execute(
            "SELECT schema_version FROM training_instance_metadata"
        ).fetchone()["schema_version"]
    finally:
        conn.close()

    assert 8 not in versions
    assert schema_version == 7
    assert not any(fk["table"] == "question_versions" for fk in foreign_keys)


def test_migration_008_recovers_orphaned_replacement_table(monkeypatch, tmp_path):
    db_path = _config(monkeypatch, tmp_path)
    migrations.run_migrations(db_path, unit_code="son_tay", migrations=migrations._MIGRATIONS[:7])
    conn = training_db.write_connection(db_path)
    try:
        conn.execute("CREATE TABLE exam_template_items_new (id TEXT NOT NULL PRIMARY KEY)")
        conn.commit()
    finally:
        conn.close()

    migrations.run_migrations(db_path, unit_code="son_tay")

    conn = training_db.write_connection(db_path)
    try:
        orphan = conn.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name='exam_template_items_new'"
        ).fetchone()
    finally:
        conn.close()

    assert orphan is None


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


@pytest.mark.parametrize("starting_version", [7, 8])
def test_migration_010_recounts_templates_after_removing_legacy_items(monkeypatch, tmp_path, starting_version):
    db_path = _config(monkeypatch, tmp_path)
    migrations.run_migrations(
        db_path, unit_code="son_tay", migrations=migrations._MIGRATIONS[:starting_version]
    )
    conn = training_db.write_connection(db_path)
    try:
        conn.execute("INSERT INTO question_items (id, created_at_ms) VALUES ('question', 0)")
        conn.execute(
            """INSERT INTO question_versions
            (id, question_item_id, version_number, type, stem, correct_option_ids_json,
             difficulty, created_by, created_at_ms)
            VALUES ('version', 'question', 1, 'single_choice', 'Question', '[]', 'easy', 'alice', 0)"""
        )
        conn.execute(
            """INSERT INTO exam_templates
            (id, code, title, target_audience_code, total_questions, duration_seconds,
             created_by, created_at_ms)
            VALUES ('template', 'TPL', 'Template', 'nvkt', 99, 60, 'alice', 0)"""
        )
        conn.commit()
        conn.execute("PRAGMA foreign_keys=OFF")
        conn.execute("DROP TABLE exam_template_items")
        conn.execute(
            """CREATE TABLE exam_template_items (
                id TEXT NOT NULL PRIMARY KEY,
                template_id TEXT NOT NULL,
                sequence_number INTEGER NOT NULL,
                question_version_id TEXT NOT NULL,
                section_label TEXT,
                points REAL NOT NULL DEFAULT 1.0
            )"""
        )
        conn.executemany(
            """INSERT INTO exam_template_items
            (id, template_id, sequence_number, question_version_id, points)
            VALUES (?, ?, ?, ?, 1.0)""",
            [
                ('valid', 'template', 1, 'version'),
                ('duplicate', 'template', 2, 'version'),
                ('orphan-question', 'template', 3, 'missing-version'),
                ('orphan-template', 'missing-template', 1, 'version'),
            ],
        )
        conn.commit()
        conn.execute("PRAGMA foreign_keys=ON")
    finally:
        conn.close()

    migrations.run_migrations(db_path, unit_code="son_tay")
    migrations.run_migrations(db_path, unit_code="son_tay")

    conn = training_db.read_connection(db_path)
    try:
        items = conn.execute(
            "SELECT id FROM exam_template_items WHERE template_id='template' ORDER BY sequence_number"
        ).fetchall()
        total = conn.execute(
            "SELECT total_questions FROM exam_templates WHERE id='template'"
        ).fetchone()["total_questions"]
    finally:
        conn.close()
    assert [row["id"] for row in items] == ["valid"]
    assert total == 1


@pytest.mark.parametrize("starting_version", [7, 8])
def test_migration_010_rolls_back_item_rebuild_and_template_counts(monkeypatch, tmp_path, starting_version):
    db_path = _config(monkeypatch, tmp_path)
    migrations.run_migrations(
        db_path, unit_code="son_tay", migrations=migrations._MIGRATIONS[:starting_version]
    )
    conn = training_db.write_connection(db_path)
    try:
        conn.execute("INSERT INTO question_items (id, created_at_ms) VALUES ('question', 0)")
        conn.execute(
            """INSERT INTO question_versions
            (id, question_item_id, version_number, type, stem, correct_option_ids_json,
             difficulty, created_by, created_at_ms)
            VALUES ('version', 'question', 1, 'single_choice', 'Question', '[]', 'easy', 'alice', 0)"""
        )
        conn.execute(
            """INSERT INTO exam_templates
            (id, code, title, target_audience_code, total_questions, duration_seconds,
             created_by, created_at_ms)
            VALUES ('template', 'TPL', 'Template', 'nvkt', 99, 60, 'alice', 0)"""
        )
        conn.commit()
        conn.execute("PRAGMA foreign_keys=OFF")
        conn.execute("DROP TABLE exam_template_items")
        conn.execute(
            """CREATE TABLE exam_template_items (
                id TEXT NOT NULL PRIMARY KEY,
                template_id TEXT NOT NULL,
                sequence_number INTEGER NOT NULL,
                question_version_id TEXT NOT NULL,
                section_label TEXT,
                points REAL NOT NULL DEFAULT 1.0
            )"""
        )
        conn.executemany(
            """INSERT INTO exam_template_items
            (id, template_id, sequence_number, question_version_id, points)
            VALUES (?, ?, ?, ?, 1.0)""",
            [
                ('valid', 'template', 1, 'version'),
                ('duplicate', 'template', 2, 'version'),
                ('orphan-question', 'template', 3, 'missing-version'),
                ('orphan-template', 'missing-template', 1, 'version'),
            ],
        )
        conn.commit()
        conn.execute("PRAGMA foreign_keys=ON")
    finally:
        conn.close()

    monkeypatch.setattr(migrations, "_record_migration", lambda *_: (_ for _ in ()).throw(RuntimeError("record failed")))
    with pytest.raises(RuntimeError, match="record failed"):
        migrations.run_migrations(
            db_path, unit_code="son_tay", migrations=[(10, migrations.migration_010)]
        )

    conn = training_db.read_connection(db_path)
    try:
        version = conn.execute("SELECT MAX(version) AS version FROM training_schema_migrations").fetchone()["version"]
        total = conn.execute("SELECT total_questions FROM exam_templates WHERE id='template'").fetchone()["total_questions"]
        items = conn.execute("SELECT id FROM exam_template_items ORDER BY id").fetchall()
        replacement = conn.execute(
            "SELECT 1 FROM sqlite_master WHERE type='table' AND name='exam_template_items_new'"
        ).fetchone()
    finally:
        conn.close()
    assert version == starting_version
    assert total == 99
    assert [row["id"] for row in items] == ["duplicate", "orphan-question", "orphan-template", "valid"]
    assert replacement is None
