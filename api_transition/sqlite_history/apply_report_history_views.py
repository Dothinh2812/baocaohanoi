#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Apply view layer cho report_history.db."""

from __future__ import annotations

import argparse
import sqlite3
from pathlib import Path


MODULE_DIR = Path(__file__).resolve().parent
API_TRANSITION_DIR = MODULE_DIR.parent
DEFAULT_DB_PATH = API_TRANSITION_DIR / "report_history.db"
DEFAULT_VIEWS_PATH = MODULE_DIR / "report_history_views.sql"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Apply view layer cho report_history.db")
    parser.add_argument("--db-path", default=str(DEFAULT_DB_PATH), help="Duong dan file SQLite.")
    parser.add_argument("--views-path", default=str(DEFAULT_VIEWS_PATH), help="Duong dan file SQL chua view.")
    return parser.parse_args()


def apply_views(db_path: Path, views_path: Path) -> int:
    if not db_path.exists():
        raise FileNotFoundError(f"Khong tim thay DB: {db_path}")
    if not views_path.exists():
        raise FileNotFoundError(f"Khong tim thay file views: {views_path}")

    sql = views_path.read_text(encoding="utf-8")
    conn = sqlite3.connect(db_path)
    try:
        conn.execute("PRAGMA foreign_keys = ON")
        existing_views = conn.execute(
            "SELECT name FROM sqlite_master WHERE type = 'view' AND name NOT LIKE 'sqlite_%'"
        ).fetchall()
        for (view_name,) in existing_views:
            quoted_name = '"' + str(view_name).replace('"', '""') + '"'
            conn.execute(f"DROP VIEW IF EXISTS {quoted_name}")
        conn.executescript(sql)
        conn.commit()
        view_count = conn.execute(
            "SELECT COUNT(*) FROM sqlite_master WHERE type = 'view' AND name NOT LIKE 'sqlite_%'"
        ).fetchone()[0]
        print(f"Da apply views vao DB: {db_path}")
        print(f"So view: {view_count}")
        return int(view_count)
    finally:
        conn.close()


def main() -> int:
    args = parse_args()
    db_path = Path(args.db_path).expanduser().resolve()
    views_path = Path(args.views_path).expanduser().resolve()
    apply_views(db_path, views_path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
