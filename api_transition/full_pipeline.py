#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Full pipeline cho api_transition: download -> process -> import SQLite."""

from __future__ import annotations

import argparse
import sys
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence

if __package__ in (None, ""):
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from api_transition.batch_download import (
    REPORT_MONTH,
    REPORT_YEAR,
    run_batch_download,
)
from api_transition.processors import run_all_processors
from api_transition.runtime_config import RuntimeContext, load_runtime_context
from api_transition.sqlite_history.apply_report_history_views import (
    DEFAULT_VIEWS_PATH,
    apply_views,
)
from api_transition.sqlite_history.import_processed_to_sqlite import (
    DEFAULT_ARCHIVE_ROOT,
    DEFAULT_DB_PATH,
    DEFAULT_PROCESSED_ROOT,
    archive_processed_file,
    connect_sqlite,
    import_workbook,
    parse_optional_date,
)
from api_transition.sqlite_history.init_report_history_db import (
    DEFAULT_SCHEMA_PATH,
    init_database,
)


@dataclass
class FullPipelineSummary:
    snapshot_date: str
    db_path: str
    download_success: int
    download_failed: int
    processor_success: int
    processor_failed: int
    imported: int
    import_failed: int
    skipped: int
    archived: int


def _effective_report_month(value: Optional[int]) -> int:
    return value if value is not None else REPORT_MONTH


def _effective_report_year(value: Optional[int]) -> int:
    return value if value is not None else REPORT_YEAR


def _ensure_database(
    db_path: Path,
    schema_path: Path,
    views_path: Path,
    *,
    reset_db: bool,
    verbose: bool,
) -> None:
    if verbose:
        if reset_db:
            print(f"[db] Recreating SQLite at {db_path}")
        elif db_path.exists():
            print(f"[db] Ensuring summary schema at {db_path}")
        else:
            print(f"[db] Initializing SQLite at {db_path}")
    init_database(db_path, schema_path, views_path, reset=reset_db)


def _iter_xlsx_paths(value: Any) -> Iterable[Path]:
    if isinstance(value, Path):
        yield value
        return
    if isinstance(value, str):
        if value.lower().endswith(".xlsx"):
            yield Path(value)
        return
    if isinstance(value, dict):
        for item in value.values():
            yield from _iter_xlsx_paths(item)
        return
    if isinstance(value, (list, tuple, set)):
        for item in value:
            yield from _iter_xlsx_paths(item)


def _collect_successful_processed_workbooks(
    processor_results: Dict[str, List[Any]],
    processed_root: Path,
) -> List[Path]:
    processed_root = processed_root.resolve()
    workbook_paths: List[Path] = []
    seen: set[Path] = set()
    for item in processor_results.get("success", []):
        for candidate in _iter_xlsx_paths(item.result):
            resolved = candidate.expanduser().resolve()
            try:
                resolved.relative_to(processed_root)
            except ValueError:
                continue
            if not resolved.exists() or resolved in seen:
                continue
            seen.add(resolved)
            workbook_paths.append(resolved)
    return workbook_paths


def archive_processed_outputs(
    *,
    workbook_paths: Sequence[Path],
    processed_root: Path,
    archive_root: Path,
    snapshot_date: date,
    verbose: bool = True,
) -> Dict[Path, Path]:
    archived_paths: Dict[Path, Path] = {}
    processed_root = processed_root.expanduser().resolve()
    archive_root = archive_root.expanduser().resolve()

    for workbook_path in workbook_paths:
        resolved = workbook_path.expanduser().resolve()
        archive_path = archive_processed_file(resolved, processed_root, archive_root, snapshot_date)
        archived_paths[resolved] = archive_path
        if verbose:
            rel_path = resolved.relative_to(processed_root)
            print(f"[archive] {rel_path} -> {archive_path.relative_to(archive_root.parent)}")
    return archived_paths


def import_processed_to_report_history(
    *,
    db_path: Path = DEFAULT_DB_PATH,
    processed_root: Path = DEFAULT_PROCESSED_ROOT,
    archive_root: Path = DEFAULT_ARCHIVE_ROOT,
    snapshot_date: date,
    period_start: Optional[date] = None,
    period_end: Optional[date] = None,
    report_month: Optional[int] = None,
    report_year: Optional[int] = None,
    path_contains: Optional[Sequence[str]] = None,
    workbook_paths: Optional[Sequence[Path]] = None,
    dry_run: bool = False,
    skip_archive: bool = False,
    skip_if_same_hash: bool = False,
    pre_archived_paths: Optional[Dict[Path, Path]] = None,
    verbose: bool = True,
) -> Dict[str, Any]:
    processed_root = Path(processed_root).expanduser().resolve()
    archive_root = Path(archive_root).expanduser().resolve()
    db_path = Path(db_path).expanduser().resolve()

    if not processed_root.exists():
        raise FileNotFoundError(f"Khong tim thay Processed root: {processed_root}")

    results: List[Dict[str, Any]] = []
    import_candidates: List[Path]
    if workbook_paths is None:
        from api_transition.sqlite_history.import_processed_to_sqlite import iter_workbooks

        import_candidates = list(iter_workbooks(processed_root, path_contains or []))
    else:
        import_candidates = [Path(path).expanduser().resolve() for path in workbook_paths]

    conn = connect_sqlite(db_path)
    try:
        for workbook_path in import_candidates:
            try:
                rel_path = workbook_path.relative_to(processed_root)
            except ValueError:
                rel_path = workbook_path

            if verbose:
                print(f"[import] {rel_path}")
            try:
                results.append(
                    import_workbook(
                        conn,
                        workbook_path,
                        processed_root,
                        archive_root,
                        snapshot_date,
                        period_start,
                        period_end,
                        report_month,
                        report_year,
                        dry_run=dry_run,
                        skip_archive=skip_archive,
                        skip_if_same_hash=skip_if_same_hash,
                        pre_archived_path=(pre_archived_paths or {}).get(workbook_path),
                    )
                )
            except Exception as exc:
                results.append(
                    {
                        "report_code": None,
                        "report_name": workbook_path.stem,
                        "processed_rel_path": str(rel_path).replace("\\", "/"),
                        "snapshot_date": snapshot_date.isoformat(),
                        "file_sha256": None,
                        "raw_rows": 0,
                        "summary_rows": 0,
                        "detail_rows": 0,
                        "business_rows": 0,
                        "status": "failed",
                        "archive_path": None,
                        "error": f"{type(exc).__name__}: {exc}",
                    }
                )
                if verbose:
                    print(f"[import] FAILED {rel_path}: {type(exc).__name__}: {exc}")
    finally:
        conn.close()

    return {
        "db_path": str(db_path),
        "processed_root": str(processed_root),
        "archive_root": str(archive_root),
        "count": len(results),
        "results": results,
    }


def run_full_pipeline(
    *,
    config_path: Optional[str] = None,
    runtime_context: Optional[RuntimeContext] = None,
    report_month: Optional[int] = None,
    report_year: Optional[int] = None,
    month_id: Optional[str] = None,
    month_label: Optional[str] = None,
    vattu_start_date: Optional[str] = None,
    headed: Optional[bool] = None,
    download_only: Optional[Sequence[str]] = None,
    download_skip: Optional[Sequence[str]] = None,
    overwrite_processed: bool = False,
    processor_only: Optional[Sequence[str]] = None,
    processor_skip: Optional[Sequence[str]] = None,
    processor_groups: Optional[Sequence[str]] = None,
    processor_stop_on_error: bool = False,
    db_path: Optional[Path] = None,
    processed_root: Optional[Path] = None,
    archive_root: Optional[Path] = None,
    snapshot_date: Optional[date] = None,
    period_start: Optional[date] = None,
    period_end: Optional[date] = None,
    skip_archive: bool = False,
    skip_if_same_hash: bool = False,
    reset_db: bool = False,
    allow_partial: bool = True,
    dry_run_import: bool = False,
    verbose: bool = True,
) -> Dict[str, Any]:
    """Chay full pipeline tu download den import SQLite.

    Theo mac dinh, pipeline dung truoc khi import neu download hoac processor
    co loi de tranh nap du lieu cu/stale vao DB.
    """

    active_runtime_context = runtime_context
    if active_runtime_context is None and config_path:
        active_runtime_context = load_runtime_context(config_path)

    if active_runtime_context is not None:
        if report_month is None:
            report_month = active_runtime_context.period.report_month
        if report_year is None:
            report_year = active_runtime_context.period.report_year
        if month_id is None:
            month_id = active_runtime_context.period.month_id or None
        if month_label is None:
            month_label = active_runtime_context.period.month_label or None
        if vattu_start_date is None:
            vattu_start_date = active_runtime_context.period.vattu_start_date or None
        if headed is None:
            headed = active_runtime_context.download.headed
        if db_path is None:
            db_path = active_runtime_context.paths.sqlite_db_path
        if processed_root is None:
            processed_root = active_runtime_context.paths.processed_root
        if archive_root is None:
            archive_root = active_runtime_context.paths.archive_root

    snapshot_date = snapshot_date or date.today()
    eff_report_month = _effective_report_month(report_month)
    eff_report_year = _effective_report_year(report_year)
    db_path = Path(db_path or DEFAULT_DB_PATH).expanduser().resolve()
    processed_root = Path(processed_root or DEFAULT_PROCESSED_ROOT).expanduser().resolve()
    archive_root = Path(archive_root or DEFAULT_ARCHIVE_ROOT).expanduser().resolve()

    if verbose:
        print("[pipeline] Starting full pipeline")
        print(f"[pipeline] snapshot_date={snapshot_date.isoformat()} report_month={eff_report_month} report_year={eff_report_year}")
        if active_runtime_context is not None:
            print(
                f"[pipeline] unit={active_runtime_context.unit.code}"
                f" downloads_root={active_runtime_context.paths.downloads_root}"
                f" processed_root={processed_root}"
                f" archive_root={archive_root}"
                f" db_path={db_path}"
            )

    download_results = run_batch_download(
        config_path=config_path if active_runtime_context is None else None,
        runtime_context=active_runtime_context,
        report_month=report_month,
        report_year=report_year,
        month_id=month_id,
        month_label=month_label,
        vattu_start_date=vattu_start_date,
        headed=headed,
        skip_reports=list(download_skip or []),
        only_reports=list(download_only or []),
    )
    if download_results["failed"] and verbose:
        print(f"[pipeline] Download failed for {len(download_results['failed'])} report(s)")

    processor_results = run_all_processors(
        overwrite_processed=overwrite_processed,
        only=processor_only,
        skip=processor_skip,
        groups=processor_groups,
        runtime_context=active_runtime_context,
        stop_on_error=processor_stop_on_error,
        verbose=verbose,
    )
    if processor_results["failed"] and verbose:
        print(f"[pipeline] Processor failed for {len(processor_results['failed'])} task(s)")

    successful_workbooks = _collect_successful_processed_workbooks(processor_results, processed_root)
    if not successful_workbooks:
        raise RuntimeError("Khong co workbook processed thanh cong nao de archive/import.")

    if verbose:
        print(f"[pipeline] Successful processed workbooks: {len(successful_workbooks)}")

    archived_paths: Dict[Path, Path] = {}
    if not skip_archive:
        archived_paths = archive_processed_outputs(
            workbook_paths=successful_workbooks,
            processed_root=processed_root,
            archive_root=archive_root,
            snapshot_date=snapshot_date,
            verbose=verbose,
        )
    elif verbose:
        print("[pipeline] Skip archiving ProcessedDaily by request")

    if (download_results["failed"] or processor_results["failed"]) and not allow_partial:
        raise RuntimeError(
            "Pipeline co loi o download/process va dang chay strict mode; dung truoc khi import."
        )

    _ensure_database(
        db_path,
        DEFAULT_SCHEMA_PATH,
        DEFAULT_VIEWS_PATH,
        reset_db=reset_db,
        verbose=verbose,
    )

    import_results = import_processed_to_report_history(
        db_path=db_path,
        processed_root=processed_root,
        archive_root=archive_root,
        snapshot_date=snapshot_date,
        period_start=period_start,
        period_end=period_end,
        report_month=eff_report_month,
        report_year=eff_report_year,
        workbook_paths=successful_workbooks,
        dry_run=dry_run_import,
        skip_archive=True if archived_paths else skip_archive,
        skip_if_same_hash=skip_if_same_hash,
        pre_archived_paths=archived_paths,
        verbose=verbose,
    )
    apply_views(db_path, DEFAULT_VIEWS_PATH)

    summary = FullPipelineSummary(
        snapshot_date=snapshot_date.isoformat(),
        db_path=str(db_path),
        download_success=len(download_results["success"]),
        download_failed=len(download_results["failed"]),
        processor_success=len(processor_results["success"]),
        processor_failed=len(processor_results["failed"]),
        imported=sum(1 for item in import_results["results"] if item["status"] == "imported"),
        import_failed=sum(1 for item in import_results["results"] if item["status"] == "failed"),
        skipped=sum(1 for item in import_results["results"] if item["status"] in {"skipped", "dry_run"}),
        archived=len(archived_paths),
    )

    if verbose:
        print(
            "[pipeline] Done:"
            f" download_ok={summary.download_success}"
            f" download_failed={summary.download_failed}"
            f" processor_ok={summary.processor_success}"
            f" processor_failed={summary.processor_failed}"
            f" archived={summary.archived}"
            f" imported={summary.imported}"
            f" import_failed={summary.import_failed}"
            f" skipped={summary.skipped}"
        )

    return {
        "summary": summary,
        "download": download_results,
        "processors": processor_results,
        "import": import_results,
    }


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Run full api_transition pipeline.")
    parser.add_argument("--config", default=None, help="Duong dan file config runtime theo don vi.")
    parser.add_argument("--month", type=int, default=None, help=f"Thang bao cao (mac dinh: {REPORT_MONTH})")
    parser.add_argument("--year", type=int, default=None, help=f"Nam bao cao (mac dinh: {REPORT_YEAR})")
    parser.add_argument("--month-id", default=None, help="Month ID dung cho batch download.")
    parser.add_argument("--month-label", default=None, help="Nhan ky bao cao.")
    parser.add_argument("--vattu-start-date", default=None, help="Ngay bat dau rieng cho vat tu thu hoi.")
    parser.add_argument("--headed", action="store_true", default=None, help="Mo trinh duyet co giao dien.")
    parser.add_argument("--download-only", nargs="+", default=[], metavar="NAME", help="Chi tai cac report nay.")
    parser.add_argument("--download-skip", nargs="+", default=[], metavar="NAME", help="Bo qua cac report nay khi download.")
    parser.add_argument("--processor-only", action="append", default=[], help="Chi chay processor co ten nay. Co the lap lai.")
    parser.add_argument("--processor-skip", action="append", default=[], help="Bo qua processor co ten nay. Co the lap lai.")
    parser.add_argument("--processor-group", action="append", default=[], help="Chi chay processor group nay. Co the lap lai.")
    parser.add_argument("--overwrite-processed", action="store_true", help="Ghi de workbook processed neu processor ho tro.")
    parser.add_argument("--processor-stop-on-error", action="store_true", help="Dung processor stage ngay khi gap loi.")
    parser.add_argument("--db-path", default=None, help=f"Duong dan SQLite DB. Mac dinh: {DEFAULT_DB_PATH}")
    parser.add_argument("--processed-root", default=None, help=f"Thu muc Processed. Mac dinh: {DEFAULT_PROCESSED_ROOT}")
    parser.add_argument("--archive-root", default=None, help=f"Thu muc ProcessedDaily. Mac dinh: {DEFAULT_ARCHIVE_ROOT}")
    parser.add_argument("--snapshot-date", default=None, help="Ngay du lieu YYYY-MM-DD. Mac dinh: hom nay.")
    parser.add_argument("--period-start", default=None, help="Ky du lieu bat dau YYYY-MM-DD.")
    parser.add_argument("--period-end", default=None, help="Ky du lieu ket thuc YYYY-MM-DD.")
    parser.add_argument("--skip-archive", action="store_true", help="Khong tao ban sao vao ProcessedDaily khi import.")
    parser.add_argument("--skip-if-same-hash", action="store_true", help="Bo qua import neu cung ngay va cung hash file.")
    parser.add_argument("--reset-db", action="store_true", help="Reset report_history.db truoc khi import.")
    parser.add_argument("--allow-partial", action="store_true", help="Tuong thich nguoc. Pipeline hien mac dinh da import phan thanh cong.")
    parser.add_argument("--strict", action="store_true", help="Dung truoc khi import neu download/process co loi.")
    parser.add_argument("--dry-run-import", action="store_true", help="Chi parse import, khong ghi vao SQLite.")
    parser.add_argument("--quiet", action="store_true", help="Giam log pipeline.")
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = _build_parser()
    args = parser.parse_args(argv)
    runtime_context = load_runtime_context(args.config) if args.config else None

    result = run_full_pipeline(
        config_path=args.config,
        runtime_context=runtime_context,
        report_month=args.month,
        report_year=args.year,
        month_id=args.month_id,
        month_label=args.month_label,
        vattu_start_date=args.vattu_start_date,
        headed=args.headed,
        download_only=args.download_only,
        download_skip=args.download_skip,
        overwrite_processed=args.overwrite_processed,
        processor_only=args.processor_only,
        processor_skip=args.processor_skip,
        processor_groups=args.processor_group,
        processor_stop_on_error=args.processor_stop_on_error,
        db_path=Path(args.db_path) if args.db_path else None,
        processed_root=Path(args.processed_root) if args.processed_root else None,
        archive_root=Path(args.archive_root) if args.archive_root else None,
        snapshot_date=parse_optional_date(args.snapshot_date),
        period_start=parse_optional_date(args.period_start),
        period_end=parse_optional_date(args.period_end),
        skip_archive=args.skip_archive,
        skip_if_same_hash=args.skip_if_same_hash,
        reset_db=args.reset_db,
        allow_partial=False if args.strict else True,
        dry_run_import=args.dry_run_import,
        verbose=not args.quiet,
    )
    summary: FullPipelineSummary = result["summary"]
    return 1 if summary.download_failed or summary.processor_failed or summary.import_failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
