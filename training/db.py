"""Quản lý kết nối SQLite cho training.db per-instance.

Mỗi request/worker dùng connection riêng. Write connection bật WAL,
busy_timeout và foreign_keys. Read connection dùng row_factory.
"""

import os
import sqlite3

import config

TRAINING_DB_PATH = config.TRAINING_DB_PATH
UNIT_CODE = config.UNIT_CODE
TRAINING_FILES_DIR = config.TRAINING_FILES_DIR


def read_connection(db_path=None):
    """Kết nối read-only với row_factory."""
    path = db_path or TRAINING_DB_PATH
    conn = sqlite3.connect(f"file:{os.path.abspath(path)}?mode=ro", uri=True, timeout=10)
    conn.execute("PRAGMA busy_timeout=5000")
    conn.row_factory = sqlite3.Row
    return conn


def write_connection(db_path=None):
    """Kết nối ghi: WAL, busy_timeout, foreign_keys ON, row_factory."""
    path = db_path or TRAINING_DB_PATH
    conn = sqlite3.connect(path, timeout=10)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA busy_timeout=5000")
    conn.execute("PRAGMA foreign_keys=ON")
    conn.row_factory = sqlite3.Row
    return conn


def ensure_files_dir():
    os.makedirs(TRAINING_FILES_DIR, exist_ok=True)
