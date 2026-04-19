#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Khoi tao SQLite report_history.db trong api_transition."""

from __future__ import annotations

import argparse
import sqlite3
from pathlib import Path


MODULE_DIR = Path(__file__).resolve().parent
API_TRANSITION_DIR = MODULE_DIR.parent
DEFAULT_DB_PATH = API_TRANSITION_DIR / "report_history.db"
DEFAULT_SCHEMA_PATH = MODULE_DIR / "report_history_schema.sql"
DEFAULT_VIEWS_PATH = MODULE_DIR / "report_history_views.sql"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Khoi tao SQLite report_history.db")
    parser.add_argument("--db-path", default=str(DEFAULT_DB_PATH), help="Duong dan file SQLite can tao.")
    parser.add_argument("--schema-path", default=str(DEFAULT_SCHEMA_PATH), help="Duong dan file schema SQL.")
    parser.add_argument("--views-path", default=str(DEFAULT_VIEWS_PATH), help="Duong dan file SQL tao views.")
    parser.add_argument("--reset", action="store_true", help="Xoa file DB cu truoc khi khoi tao lai.")
    return parser.parse_args()


def init_database(db_path: Path, schema_path: Path, views_path: Path, reset: bool = False) -> None:
    if reset and db_path.exists():
        db_path.unlink()

    db_path.parent.mkdir(parents=True, exist_ok=True)
    schema_sql = schema_path.read_text(encoding="utf-8")
    views_sql = views_path.read_text(encoding="utf-8") if views_path.exists() else ""

    conn = sqlite3.connect(db_path)
    try:
        conn.execute("PRAGMA foreign_keys = ON")
        conn.execute("PRAGMA journal_mode = WAL")
        conn.execute("PRAGMA synchronous = NORMAL")
        conn.executescript(schema_sql)
        if views_sql:
            conn.executescript(views_sql)
        conn.commit()

        table_count = conn.execute(
            "SELECT COUNT(*) FROM sqlite_master WHERE type = 'table' AND name NOT LIKE 'sqlite_%'"
        ).fetchone()[0]
        index_count = conn.execute(
            "SELECT COUNT(*) FROM sqlite_master WHERE type = 'index' AND name NOT LIKE 'sqlite_%'"
        ).fetchone()[0]
        view_count = conn.execute(
            "SELECT COUNT(*) FROM sqlite_master WHERE type = 'view' AND name NOT LIKE 'sqlite_%'"
        ).fetchone()[0]

        print(f"Da khoi tao DB: {db_path}")
        print(f"So bang: {table_count}")
        print(f"So index: {index_count}")
        print(f"So view: {view_count}")
    finally:
        conn.close()


def main() -> None:
    args = parse_args()
    db_path = Path(args.db_path).expanduser().resolve()
    schema_path = Path(args.schema_path).expanduser().resolve()
    views_path = Path(args.views_path).expanduser().resolve()

    if not schema_path.exists():
        raise FileNotFoundError(f"Khong tim thay schema: {schema_path}")

    init_database(db_path, schema_path, views_path, reset=args.reset)


if __name__ == "__main__":
    main()
