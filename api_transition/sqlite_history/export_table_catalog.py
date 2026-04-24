#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Export danh sach table, file nguon, sheet nguon va cac cot ra CSV."""

from __future__ import annotations

import argparse
import csv
import sqlite3
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, List, Sequence


MODULE_DIR = Path(__file__).resolve().parent
API_TRANSITION_DIR = MODULE_DIR.parent
DEFAULT_OUTPUT_PATH = MODULE_DIR / "table_catalog_export.csv"
DEFAULT_ROOT_DB_PATH = API_TRANSITION_DIR / "report_history.db"
DEFAULT_RUNTIME_DIR = API_TRANSITION_DIR / "runtime"
SYSTEM_TABLE_PREFIXES = ("sqlite_",)


@dataclass(frozen=True)
class DbRecord:
    unit_code: str
    db_path: Path


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Export table catalog tu cac DB SQLite sang CSV.")
    parser.add_argument(
        "--output",
        default=str(DEFAULT_OUTPUT_PATH),
        help=f"Duong dan file CSV dau ra. Mac dinh: {DEFAULT_OUTPUT_PATH}",
    )
    parser.add_argument(
        "--db-path",
        action="append",
        default=[],
        help="Chi export DB nay. Co the truyen nhieu lan.",
    )
    parser.add_argument(
        "--include-root-db",
        action="store_true",
        help="Bao gom report_history.db o root neu co du lieu.",
    )
    return parser.parse_args(argv)


def _runtime_db_records() -> Iterable[DbRecord]:
    if not DEFAULT_RUNTIME_DIR.exists():
        return
    for db_path in sorted(DEFAULT_RUNTIME_DIR.glob("*/sqlite_history/report_history.db")):
        unit_code = db_path.parent.parent.name
        yield DbRecord(unit_code=unit_code, db_path=db_path.resolve())


def _explicit_db_records(paths: Sequence[str]) -> List[DbRecord]:
    records: List[DbRecord] = []
    for raw_path in paths:
        db_path = Path(raw_path).expanduser().resolve()
        unit_code = db_path.parent.parent.name if db_path.parent.name == "sqlite_history" else db_path.stem
        records.append(DbRecord(unit_code=unit_code, db_path=db_path))
    return records


def collect_db_records(args: argparse.Namespace) -> List[DbRecord]:
    if args.db_path:
        return _explicit_db_records(args.db_path)

    records = list(_runtime_db_records())
    if args.include_root_db and DEFAULT_ROOT_DB_PATH.exists():
        records.append(DbRecord(unit_code="root", db_path=DEFAULT_ROOT_DB_PATH.resolve()))
    return records


def _table_columns(conn: sqlite3.Connection, table_name: str) -> List[str]:
    rows = conn.execute(f'PRAGMA table_info("{table_name.replace(chr(34), chr(34) * 2)}")').fetchall()
    return [str(row[1]) for row in rows]


def _has_imported_data(conn: sqlite3.Connection) -> bool:
    row = conn.execute(
        "SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = 'sheet_bao_cao_tong_hop'"
    ).fetchone()
    if row is None:
        return False
    row = conn.execute("SELECT 1 FROM sheet_bao_cao_tong_hop LIMIT 1").fetchone()
    return row is not None


def _table_rows_for_db(record: DbRecord) -> List[dict]:
    if not record.db_path.exists():
        return []

    conn = sqlite3.connect(record.db_path)
    try:
        if not _has_imported_data(conn):
            return []

        metadata_rows = conn.execute(
            """
            SELECT
                s.ten_bang_du_lieu,
                GROUP_CONCAT(DISTINCT b.ten_tep_nguon) AS source_files,
                GROUP_CONCAT(DISTINCT s.ten_sheet) AS source_sheets
            FROM sheet_bao_cao_tong_hop s
            JOIN bao_cao_tong_hop_ngay b
              ON b.id = s.bao_cao_tong_hop_ngay_id
            GROUP BY s.ten_bang_du_lieu
            ORDER BY s.ten_bang_du_lieu
            """
        ).fetchall()

        rows: List[dict] = []
        for table_name, source_files, source_sheets in metadata_rows:
            if not table_name or str(table_name).startswith(SYSTEM_TABLE_PREFIXES):
                continue
            column_names = _table_columns(conn, str(table_name))
            row = {
                "unit_code": record.unit_code,
                "db_path": str(record.db_path),
                "table_name": str(table_name),
                "source_file": source_files or "",
                "source_sheet": source_sheets or "",
            }
            for index, column_name in enumerate(column_names, start=1):
                row[f"db_column_{index}"] = column_name
            rows.append(row)
        return rows
    finally:
        conn.close()


def export_catalog(records: Sequence[DbRecord], output_path: Path) -> int:
    all_rows: List[dict] = []
    for record in records:
        all_rows.extend(_table_rows_for_db(record))

    base_fields = ["unit_code", "db_path", "table_name", "source_file", "source_sheet"]
    dynamic_fields = sorted(
        {
            key
            for row in all_rows
            for key in row.keys()
            if key.startswith("db_column_")
        },
        key=lambda value: int(value.split("_")[-1]),
    )
    fieldnames = base_fields + dynamic_fields

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(all_rows)
    return len(all_rows)


def main(argv: Sequence[str] | None = None) -> int:
    args = parse_args(argv)
    records = collect_db_records(args)
    output_path = Path(args.output).expanduser().resolve()
    row_count = export_catalog(records, output_path)
    print(f"Da ghi {row_count} dong vao {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
