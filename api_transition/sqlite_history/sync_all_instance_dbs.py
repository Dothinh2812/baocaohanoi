#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Dong bo SQLite DB cho tat ca instance multi-unit.

Script nay la admin utility, KHONG nam trong full pipeline hang ngay.
Dung no khi can:

- kiem tra trang thai DB cua cac instance
- apply lai views cho tat ca DB
- tao DB cho cac instance chua co
- reset va init lai DB cho tat ca instance
"""

from __future__ import annotations

import argparse
import sqlite3
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, List, Optional, Sequence

from api_transition.runtime_config import RuntimeContext, load_runtime_context
from api_transition.sqlite_history.apply_report_history_views import (
    DEFAULT_VIEWS_PATH,
    apply_views,
)
from api_transition.sqlite_history.init_report_history_db import (
    DEFAULT_SCHEMA_PATH,
    init_database,
)


API_TRANSITION_DIR = Path(__file__).resolve().parent.parent
DEFAULT_CONFIGS_DIR = API_TRANSITION_DIR / "configs" / "units"


@dataclass(frozen=True)
class InstanceRecord:
    config_path: Path
    runtime_context: RuntimeContext


@dataclass
class SyncResult:
    unit_code: str
    config_path: str
    db_path: str
    action: str
    status: str
    detail: str = ""


def _config_candidates(configs_dir: Path) -> Iterable[Path]:
    for path in sorted(configs_dir.glob("*.yaml")):
        if path.name.startswith("_"):
            continue
        yield path.resolve()


def _load_instances(configs_dir: Path, units: Sequence[str]) -> List[InstanceRecord]:
    if not configs_dir.exists():
        raise FileNotFoundError(f"Khong tim thay configs dir: {configs_dir}")

    unit_filter = {item.strip() for item in units if item.strip()}
    records: List[InstanceRecord] = []
    for config_path in _config_candidates(configs_dir):
        context = load_runtime_context(config_path)
        if unit_filter and context.unit.code not in unit_filter:
            continue
        records.append(
            InstanceRecord(
                config_path=config_path,
                runtime_context=context,
            )
        )

    if unit_filter:
        found = {record.runtime_context.unit.code for record in records}
        missing = sorted(unit_filter - found)
        if missing:
            raise ValueError(f"Khong tim thay config cho unit: {', '.join(missing)}")

    return records


def _count_sqlite_objects(db_path: Path) -> tuple[int, int]:
    conn = sqlite3.connect(db_path)
    try:
        table_count = conn.execute(
            "SELECT COUNT(*) FROM sqlite_master WHERE type = 'table' AND name NOT LIKE 'sqlite_%'"
        ).fetchone()[0]
        view_count = conn.execute(
            "SELECT COUNT(*) FROM sqlite_master WHERE type = 'view' AND name NOT LIKE 'sqlite_%'"
        ).fetchone()[0]
        return int(table_count), int(view_count)
    finally:
        conn.close()


def _status_detail(db_path: Path) -> str:
    if not db_path.exists():
        return "missing"
    table_count, view_count = _count_sqlite_objects(db_path)
    size_mb = db_path.stat().st_size / (1024 * 1024)
    return f"exists tables={table_count} views={view_count} size_mb={size_mb:.1f}"


def _run_status(records: Sequence[InstanceRecord]) -> List[SyncResult]:
    results: List[SyncResult] = []
    for record in records:
        db_path = record.runtime_context.paths.sqlite_db_path
        results.append(
            SyncResult(
                unit_code=record.runtime_context.unit.code,
                config_path=str(record.config_path),
                db_path=str(db_path),
                action="status",
                status="ok",
                detail=_status_detail(db_path),
            )
        )
    return results


def _run_apply_views(
    records: Sequence[InstanceRecord],
    *,
    views_path: Path,
) -> List[SyncResult]:
    results: List[SyncResult] = []
    for record in records:
        db_path = record.runtime_context.paths.sqlite_db_path
        if not db_path.exists():
            results.append(
                SyncResult(
                    unit_code=record.runtime_context.unit.code,
                    config_path=str(record.config_path),
                    db_path=str(db_path),
                    action="apply-views",
                    status="skipped",
                    detail="DB missing",
                )
            )
            continue
        try:
            view_count = apply_views(db_path, views_path)
            results.append(
                SyncResult(
                    unit_code=record.runtime_context.unit.code,
                    config_path=str(record.config_path),
                    db_path=str(db_path),
                    action="apply-views",
                    status="ok",
                    detail=f"views={view_count}",
                )
            )
        except Exception as exc:
            results.append(
                SyncResult(
                    unit_code=record.runtime_context.unit.code,
                    config_path=str(record.config_path),
                    db_path=str(db_path),
                    action="apply-views",
                    status="failed",
                    detail=f"{type(exc).__name__}: {exc}",
                )
            )
    return results


def _run_apply_schema(
    records: Sequence[InstanceRecord],
    *,
    schema_path: Path,
    views_path: Path,
) -> List[SyncResult]:
    results: List[SyncResult] = []
    for record in records:
        db_path = record.runtime_context.paths.sqlite_db_path
        try:
            init_database(db_path, schema_path, views_path, reset=False)
            table_count, view_count = _count_sqlite_objects(db_path)
            results.append(
                SyncResult(
                    unit_code=record.runtime_context.unit.code,
                    config_path=str(record.config_path),
                    db_path=str(db_path),
                    action="apply-schema",
                    status="ok",
                    detail=f"tables={table_count} views={view_count}",
                )
            )
        except Exception as exc:
            results.append(
                SyncResult(
                    unit_code=record.runtime_context.unit.code,
                    config_path=str(record.config_path),
                    db_path=str(db_path),
                    action="apply-schema",
                    status="failed",
                    detail=f"{type(exc).__name__}: {exc}",
                )
            )
    return results


def _run_init(
    records: Sequence[InstanceRecord],
    *,
    schema_path: Path,
    views_path: Path,
    reset: bool,
) -> List[SyncResult]:
    action = "reset-and-init" if reset else "init-if-missing"
    results: List[SyncResult] = []
    for record in records:
        db_path = record.runtime_context.paths.sqlite_db_path
        if db_path.exists() and not reset:
            results.append(
                SyncResult(
                    unit_code=record.runtime_context.unit.code,
                    config_path=str(record.config_path),
                    db_path=str(db_path),
                    action=action,
                    status="skipped",
                    detail="DB already exists",
                )
            )
            continue
        try:
            init_database(db_path, schema_path, views_path, reset=reset)
            table_count, view_count = _count_sqlite_objects(db_path)
            results.append(
                SyncResult(
                    unit_code=record.runtime_context.unit.code,
                    config_path=str(record.config_path),
                    db_path=str(db_path),
                    action=action,
                    status="ok",
                    detail=f"tables={table_count} views={view_count}",
                )
            )
        except Exception as exc:
            results.append(
                SyncResult(
                    unit_code=record.runtime_context.unit.code,
                    config_path=str(record.config_path),
                    db_path=str(db_path),
                    action=action,
                    status="failed",
                    detail=f"{type(exc).__name__}: {exc}",
                )
            )
    return results


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Dong bo SQLite DB cho tat ca instance multi-unit. Day la admin utility, khong nam trong full pipeline hang ngay.",
    )
    parser.add_argument(
        "--configs-dir",
        default=str(DEFAULT_CONFIGS_DIR),
        help=f"Thu muc chua config don vi. Mac dinh: {DEFAULT_CONFIGS_DIR}",
    )
    parser.add_argument(
        "--mode",
        required=True,
        choices=("status", "apply-views", "apply-schema", "init-if-missing", "reset-and-init"),
        help="Che do dong bo DB.",
    )
    parser.add_argument(
        "--unit",
        action="append",
        default=[],
        help="Chi ap dung cho unit code nay. Co the lap lai nhieu lan.",
    )
    parser.add_argument(
        "--schema-path",
        default=str(DEFAULT_SCHEMA_PATH),
        help=f"Duong dan schema SQL. Mac dinh: {DEFAULT_SCHEMA_PATH}",
    )
    parser.add_argument(
        "--views-path",
        default=str(DEFAULT_VIEWS_PATH),
        help=f"Duong dan views SQL. Mac dinh: {DEFAULT_VIEWS_PATH}",
    )
    parser.add_argument(
        "--quiet",
        action="store_true",
        help="Giam log chi tiet.",
    )
    return parser


def _print_results(results: Sequence[SyncResult]) -> None:
    for item in results:
        print(
            f"[{item.status.upper():<7}]"
            f" unit={item.unit_code}"
            f" action={item.action}"
            f" db={item.db_path}"
            f" detail={item.detail}"
        )

    ok_count = sum(1 for item in results if item.status == "ok")
    skipped_count = sum(1 for item in results if item.status == "skipped")
    failed_count = sum(1 for item in results if item.status == "failed")
    print(
        f"[summary] total={len(results)} ok={ok_count} skipped={skipped_count} failed={failed_count}"
    )


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = _build_parser()
    args = parser.parse_args(argv)

    configs_dir = Path(args.configs_dir).expanduser().resolve()
    schema_path = Path(args.schema_path).expanduser().resolve()
    views_path = Path(args.views_path).expanduser().resolve()

    if args.mode in {"apply-schema", "init-if-missing", "reset-and-init"} and not schema_path.exists():
        raise FileNotFoundError(f"Khong tim thay schema: {schema_path}")
    if args.mode in {"apply-views", "apply-schema", "init-if-missing", "reset-and-init"} and not views_path.exists():
        raise FileNotFoundError(f"Khong tim thay views SQL: {views_path}")

    records = _load_instances(configs_dir, args.unit)
    if not records:
        print("Khong co instance nao phu hop.")
        return 0

    if not args.quiet:
        print(f"[sync-db] mode={args.mode} instances={len(records)} configs_dir={configs_dir}")

    if args.mode == "status":
        results = _run_status(records)
    elif args.mode == "apply-views":
        results = _run_apply_views(records, views_path=views_path)
    elif args.mode == "apply-schema":
        results = _run_apply_schema(records, schema_path=schema_path, views_path=views_path)
    elif args.mode == "init-if-missing":
        results = _run_init(records, schema_path=schema_path, views_path=views_path, reset=False)
    else:
        results = _run_init(records, schema_path=schema_path, views_path=views_path, reset=True)

    _print_results(results)
    return 1 if any(item.status == "failed" for item in results) else 0


if __name__ == "__main__":
    raise SystemExit(main())
