"""Schema migration runner cho training.db.

Có version, chạy trong transaction, dùng file lock per-instance để nhiều
Gunicorn worker không migrate đồng thời. Kiểm tra unit_code trước khi thay schema.
"""

import fcntl
import os
import time

from training import time_policy
from training.db import write_connection

CURRENT_SCHEMA_VERSION = 1


class UnitCodeMismatchError(Exception):
    pass


def _lock_path(db_path):
    return os.path.abspath(db_path) + ".migrate.lock"


def _acquire_lock(db_path):
    lock_file = open(_lock_path(db_path), "w")
    fcntl.flock(lock_file.fileno(), fcntl.LOCK_EX)
    return lock_file


def _table_exists(conn, name):
    row = conn.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name=?", (name,)
    ).fetchone()
    return row is not None


def _applied_versions(conn):
    if not _table_exists(conn, "training_schema_migrations"):
        return set()
    return {
        row["version"]
        for row in conn.execute(
            "SELECT version FROM training_schema_migrations"
        ).fetchall()
    }


def _check_or_init_instance_metadata(conn, unit_code):
    """Kiểm tra unit_code trên DB đã có metadata; trả về True nếu đã có sẵn."""
    if not _table_exists(conn, "training_instance_metadata"):
        return False
    row = conn.execute(
        "SELECT unit_code FROM training_instance_metadata"
    ).fetchone()
    if row and row["unit_code"] != unit_code:
        raise UnitCodeMismatchError(
            f"training.db unit_code={row['unit_code']!r} != expected {unit_code!r}"
        )
    return row is not None


def _init_instance_metadata(conn, unit_code, schema_version):
    conn.execute(
        "INSERT INTO training_instance_metadata (unit_code, schema_version, created_at_ms) "
        "VALUES (?, ?, ?)",
        (unit_code, schema_version, time_policy.utc_now_ms()),
    )


def _record_migration(conn, version):
    conn.execute(
        "INSERT INTO training_schema_migrations (version, applied_at_ms) VALUES (?, ?)",
        (version, time_policy.utc_now_ms()),
    )


# --- migration definitions ---

def migration_001(conn):
    """Foundation: instance metadata, schema migrations, audit log."""
    conn.executescript(
        """
        CREATE TABLE IF NOT EXISTS training_instance_metadata (
            unit_code TEXT NOT NULL PRIMARY KEY,
            schema_version INTEGER NOT NULL,
            created_at_ms INTEGER NOT NULL
        );

        CREATE TABLE IF NOT EXISTS training_schema_migrations (
            version INTEGER NOT NULL PRIMARY KEY,
            applied_at_ms INTEGER NOT NULL
        );

        CREATE TABLE IF NOT EXISTS training_audit_log (
            id TEXT NOT NULL PRIMARY KEY,
            actor TEXT NOT NULL,
            unit_code TEXT NOT NULL,
            action TEXT NOT NULL,
            entity_type TEXT NOT NULL,
            entity_id TEXT NOT NULL,
            before_json TEXT,
            after_json TEXT,
            request_id TEXT,
            ip TEXT,
            user_agent TEXT,
            created_at_ms INTEGER NOT NULL
        );
        CREATE INDEX IF NOT EXISTS idx_audit_entity
            ON training_audit_log (entity_type, entity_id, created_at_ms);
        """
    )


_MIGRATIONS = [
    (1, migration_001),
]


def run_migrations(db_path, unit_code, *, migrations=None):
    """Chạy migration pending với file lock + unit_code check."""
    lock_file = _acquire_lock(db_path)
    try:
        conn = write_connection(db_path)
        try:
            metadata_exists = _check_or_init_instance_metadata(conn, unit_code)
            applied = _applied_versions(conn)
            pending = [
                (version, fn)
                for version, fn in (migrations or _MIGRATIONS)
                if version not in applied
            ]
            if not pending:
                return applied
            for version, fn in pending:
                fn(conn)
                _record_migration(conn, version)
            max_version = max(v for v, _ in pending)
            if metadata_exists:
                conn.execute(
                    "UPDATE training_instance_metadata SET schema_version=?",
                    (max_version,),
                )
            else:
                _init_instance_metadata(conn, unit_code, max_version)
            conn.commit()
            return applied | {v for v, _ in pending}
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()
    finally:
        fcntl.flock(lock_file.fileno(), fcntl.LOCK_UN)
        lock_file.close()


def get_schema_version(db_path):
    conn = write_connection(db_path)
    try:
        if not _table_exists(conn, "training_instance_metadata"):
            return 0
        row = conn.execute(
            "SELECT schema_version FROM training_instance_metadata"
        ).fetchone()
        return row["schema_version"] if row else 0
    finally:
        conn.close()
