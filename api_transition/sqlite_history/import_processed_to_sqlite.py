#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Import workbook processed vao SQLite theo mo hinh moi sheet tong hop mot bang rieng."""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import numbers
import re
import shutil
import sqlite3
import tempfile
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import pandas as pd

if __package__ in (None, ""):
    import sys

    sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent))


MODULE_DIR = Path(__file__).resolve().parent
API_TRANSITION_DIR = MODULE_DIR.parent
DEFAULT_DB_PATH = API_TRANSITION_DIR / "report_history.db"
DEFAULT_PROCESSED_ROOT = API_TRANSITION_DIR / "Processed"
DEFAULT_ARCHIVE_ROOT = API_TRANSITION_DIR / "ProcessedDaily"
LIST_OF_TABLE_CSV_PATH = MODULE_DIR / "list of table.csv"
DATE_IN_NAME_RE = re.compile(r"(\d{8})")
SKIPPED_PROCESSED_REL_PATHS = {
    "vat_tu_thu_hoi/quyet_toan_vat_tu_processed.xlsx",
}

SOURCE_SHEET_NAMES = {"sheet", "sheet1"}
RAW_DETAIL_SHEET_NAMES = {
    "data",
    "data_combined",
    "data_tam_dung",
    "data_khoi_phuc",
    "chi tiết vật tư",
    "chi_tiet_chua_khoi_phuc",
}
NOTE_SHEET_NAMES = {"thong_bao"}
DETAIL_NAME_HINTS = ("khong_dat", "chi tiết vật tư")
METRIC_EXCLUDE_COLUMNS = {
    "stt",
    "id",
    "baohong_id",
    "hdtb_id",
    "ma_tb",
    "ma_kh",
    "ma_gd",
    "ma_men",
    "ma_vt",
    "ma_spdv",
    "so_dt",
    "dien_thoai_lh",
    "serial",
    "serial_number",
    "ma_nv",
    "mã nv",
    "mã nhân viên",
}
METRIC_NAME_HINTS = (
    "sm",
    "tỷ lệ",
    "ty le",
    "điểm",
    "diem",
    "tổng",
    "tong",
    "số",
    "so ",
    "sl ",
    "kq ",
    "hoàn thành",
    "hoan thanh",
    "chỉ tiêu",
    "chi tieu",
)

META_SNAPSHOT_ID = "__snapshot_id"
META_SHEET_ID = "__sheet_id"
META_ROW_NUM = "__row_num"
META_ROW_HASH = "__row_hash"
META_IMPORTED_AT = "__imported_at"
SYSTEM_COLUMNS = (
    (META_SNAPSHOT_ID, "INTEGER NOT NULL"),
    (META_SHEET_ID, "INTEGER NOT NULL"),
    (META_ROW_NUM, "INTEGER NOT NULL"),
    (META_ROW_HASH, "TEXT NOT NULL"),
    (META_IMPORTED_AT, "TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP"),
)


@dataclass(frozen=True)
class ReportMeta:
    report_code: str
    report_name: str
    report_group: str
    processed_rel_path: str
    folder_name: str
    file_name: str


@dataclass
class SheetData:
    sheet_name: str
    sheet_order: int
    sheet_kind: str
    row_count: int
    column_names: List[str]
    measure_columns: List[str]
    rows: List[Dict[str, Any]]


@dataclass(frozen=True)
class SheetImportTarget:
    sheet_name: str
    sheet_code: str
    table_name: str
    sheet: SheetData


def load_import_allowlist(csv_path: Path = LIST_OF_TABLE_CSV_PATH) -> Dict[Tuple[str, str], set[str]]:
    if not csv_path.exists():
        raise FileNotFoundError(f"Khong tim thay file whitelist sheet: {csv_path}")

    allowlist: Dict[Tuple[str, str], set[str]] = {}
    with csv_path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        required_columns = {"folder_name", "file_name", "sheet_name"}
        if reader.fieldnames is None or set(reader.fieldnames) < required_columns:
            raise ValueError(
                f"File whitelist {csv_path} phai co cac cot: folder_name,file_name,sheet_name"
            )

        for row in reader:
            folder_name = row["folder_name"].strip().lower()
            file_name = row["file_name"].strip().lower()
            sheet_name = normalize_key(row["sheet_name"])
            if not folder_name or not file_name or not sheet_name:
                continue
            allowlist.setdefault((folder_name, file_name), set()).add(sheet_name)
    return allowlist


def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Import workbook processed vao SQLite summary per report sheet")
    parser.add_argument("--db-path", default=str(DEFAULT_DB_PATH), help="Duong dan file SQLite.")
    parser.add_argument("--processed-root", default=str(DEFAULT_PROCESSED_ROOT), help="Thu muc Processed.")
    parser.add_argument(
        "--archive-root",
        default=str(DEFAULT_ARCHIVE_ROOT),
        help="Thu muc ProcessedDaily luu file theo ngay.",
    )
    parser.add_argument("--snapshot-date", help="Ngay du lieu dang YYYY-MM-DD.")
    parser.add_argument("--period-start", help="Tu ngay dang YYYY-MM-DD.")
    parser.add_argument("--period-end", help="Den ngay dang YYYY-MM-DD.")
    parser.add_argument("--report-month", type=int, help="Thang bao cao.")
    parser.add_argument("--report-year", type=int, help="Nam bao cao.")
    parser.add_argument(
        "--path-contains",
        action="append",
        default=[],
        help="Chi import cac file co duong dan chua chuoi nay. Co the dung nhieu lan.",
    )
    parser.add_argument("--dry-run", action="store_true", help="Chi parse va in thong ke, khong ghi SQLite.")
    parser.add_argument("--skip-archive", action="store_true", help="Khong copy file sang ProcessedDaily.")
    parser.add_argument(
        "--skip-if-same-hash",
        action="store_true",
        help="Neu cung ngay va cung hash file thi bo qua import.",
    )
    parser.add_argument("--json", action="store_true", help="In ket qua dang JSON.")
    return parser.parse_args(argv)


def strip_accents(text: str) -> str:
    replacements = {
        "à": "a",
        "á": "a",
        "ạ": "a",
        "ả": "a",
        "ã": "a",
        "ă": "a",
        "ằ": "a",
        "ắ": "a",
        "ặ": "a",
        "ẳ": "a",
        "ẵ": "a",
        "â": "a",
        "ầ": "a",
        "ấ": "a",
        "ậ": "a",
        "ẩ": "a",
        "ẫ": "a",
        "đ": "d",
        "è": "e",
        "é": "e",
        "ẹ": "e",
        "ẻ": "e",
        "ẽ": "e",
        "ê": "e",
        "ề": "e",
        "ế": "e",
        "ệ": "e",
        "ể": "e",
        "ễ": "e",
        "ì": "i",
        "í": "i",
        "ị": "i",
        "ỉ": "i",
        "ĩ": "i",
        "ò": "o",
        "ó": "o",
        "ọ": "o",
        "ỏ": "o",
        "õ": "o",
        "ô": "o",
        "ồ": "o",
        "ố": "o",
        "ộ": "o",
        "ổ": "o",
        "ỗ": "o",
        "ơ": "o",
        "ờ": "o",
        "ớ": "o",
        "ợ": "o",
        "ở": "o",
        "ỡ": "o",
        "ù": "u",
        "ú": "u",
        "ụ": "u",
        "ủ": "u",
        "ũ": "u",
        "ư": "u",
        "ừ": "u",
        "ứ": "u",
        "ự": "u",
        "ử": "u",
        "ữ": "u",
        "ỳ": "y",
        "ý": "y",
        "ỵ": "y",
        "ỷ": "y",
        "ỹ": "y",
    }
    return "".join(replacements.get(ch, ch) for ch in text.lower())


def sanitize_slug(value: str) -> str:
    text = strip_accents(value).strip().lower()
    text = re.sub(r"_processed$", "", text)
    text = re.sub(r"\.xlsx$", "", text)
    text = re.sub(r"[^a-z0-9]+", "_", text)
    text = re.sub(r"_+", "_", text).strip("_")
    return text


def normalize_key(value: str) -> str:
    return re.sub(r"\s+", " ", strip_accents(str(value).strip())).strip()


def parse_optional_date(value: Optional[str]) -> Optional[date]:
    if not value:
        return None
    return date.fromisoformat(value)


def parse_snapshot_date(explicit_value: Optional[str], workbook_path: Path) -> date:
    if explicit_value:
        return date.fromisoformat(explicit_value)
    match = DATE_IN_NAME_RE.search(workbook_path.name)
    if not match:
        raise ValueError(
            f"Khong xac dinh duoc ngay_du_lieu tu ten file {workbook_path.name}. "
            "Hay truyen --snapshot-date YYYY-MM-DD."
        )
    return datetime.strptime(match.group(1), "%d%m%Y").date()


def sha256_bytes(content: bytes) -> str:
    return hashlib.sha256(content).hexdigest()


def short_hash(value: str, length: int = 8) -> str:
    return sha256_bytes(value.encode("utf-8"))[:length]


def jsonable_value(value: Any) -> Any:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime().isoformat()
    if isinstance(value, datetime):
        return value.isoformat()
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, bool):
        return bool(value)
    if isinstance(value, numbers.Integral):
        return int(value)
    if isinstance(value, numbers.Real):
        if float(value).is_integer():
            return int(value)
        return float(value)
    text = str(value).strip()
    return text or None


def parse_numeric(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, numbers.Number) and not pd.isna(value):
        return float(value)
    text = str(value).strip()
    if not text:
        return None
    text = text.replace("%", "").replace(",", "")
    if re.fullmatch(r"-?\d+(?:\.\d+)?", text):
        return float(text)
    return None


def clean_column_names(columns: Sequence[Any]) -> List[str]:
    seen: Dict[str, int] = {}
    cleaned: List[str] = []
    for idx, column in enumerate(columns, start=1):
        value = str(column).strip() if column is not None else ""
        if not value or value.lower().startswith("unnamed:"):
            value = f"column_{idx}"
        count = seen.get(value, 0) + 1
        seen[value] = count
        if count > 1:
            value = f"{value}_{count}"
        cleaned.append(value)
    return cleaned


def drop_empty_records(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    trimmed = df.dropna(axis=0, how="all").dropna(axis=1, how="all")
    return trimmed.reset_index(drop=True)


def looks_like_metric_name(column: str) -> bool:
    normalized = normalize_key(column)
    return any(token in normalized for token in METRIC_NAME_HINTS)


def detect_measure_columns(df: pd.DataFrame) -> List[str]:
    measure_columns: List[str] = []
    for column in df.columns:
        normalized_slug = sanitize_slug(column)
        if normalized_slug in METRIC_EXCLUDE_COLUMNS:
            continue
        series = df[column].dropna()
        if series.empty:
            continue
        parsed = series.map(parse_numeric)
        ratio = parsed.notna().sum() / len(series)
        if ratio >= 0.5 or (parsed.notna().sum() > 0 and looks_like_metric_name(column)):
            measure_columns.append(column)
    return measure_columns


def classify_sheet(sheet_name: str, df: pd.DataFrame, measure_columns: Sequence[str]) -> str:
    lower_name = normalize_key(sheet_name)
    row_count = len(df.index)
    if lower_name in SOURCE_SHEET_NAMES:
        return "source"
    if lower_name in NOTE_SHEET_NAMES:
        return "note"
    if lower_name in {normalize_key(value) for value in RAW_DETAIL_SHEET_NAMES}:
        return "detail"
    if any(token in lower_name for token in DETAIL_NAME_HINTS):
        if measure_columns:
            return "summary"
        return "other"
    if measure_columns:
        return "summary"
    if row_count <= 200:
        return "other"
    return "other"


def read_workbook_sheets(workbook_path: Path) -> List[SheetData]:
    excel = pd.ExcelFile(workbook_path, engine="openpyxl")
    results: List[SheetData] = []
    for sheet_order, sheet_name in enumerate(excel.sheet_names, start=1):
        df = excel.parse(sheet_name=sheet_name, dtype=object)
        df = drop_empty_records(df.copy())
        df.columns = clean_column_names(df.columns)
        measure_columns = detect_measure_columns(df)
        rows = [{key: jsonable_value(value) for key, value in row.items()} for row in df.to_dict(orient="records")]
        results.append(
            SheetData(
                sheet_name=sheet_name,
                sheet_order=sheet_order,
                sheet_kind=classify_sheet(sheet_name, df, measure_columns),
                row_count=len(df.index),
                column_names=list(df.columns),
                measure_columns=measure_columns,
                rows=rows,
            )
        )
    return results


def build_report_meta(processed_root: Path, workbook_path: Path) -> ReportMeta:
    rel_path = workbook_path.relative_to(processed_root)
    rel_text = str(rel_path).replace("\\", "/")
    group = rel_path.parts[0]
    file_name = workbook_path.name
    stem = workbook_path.stem
    stem = re.sub(r"_processed$", "", stem)
    clean_stem = DATE_IN_NAME_RE.sub("", stem).replace("__", "_").strip("_ ")
    if group == "xac_minh_tam_dung":
        report_code = "xac_minh_tam_dung"
    else:
        report_code = sanitize_slug(f"{group}_{clean_stem}")
    report_name = clean_stem.replace("_", " ").strip() or stem
    if group == "xac_minh_tam_dung":
        report_name = "xac minh tam dung"
    return ReportMeta(
        report_code=report_code,
        report_name=report_name,
        report_group=group,
        processed_rel_path=rel_text,
        folder_name=group,
        file_name=file_name,
    )


def is_skipped_processed_workbook(processed_root: Path, workbook_path: Path) -> bool:
    rel_path = str(workbook_path.relative_to(processed_root)).replace("\\", "/").lower()
    return rel_path in SKIPPED_PROCESSED_REL_PATHS


def get_allowed_sheet_names(
    report_meta: ReportMeta,
    allowlist: Dict[Tuple[str, str], set[str]],
) -> set[str]:
    return allowlist.get((report_meta.folder_name.lower(), report_meta.file_name.lower()), set())


def current_timestamp() -> str:
    return datetime.now().isoformat(sep=" ", timespec="seconds")


def connect_sqlite(db_path: Path) -> sqlite3.Connection:
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    conn.execute("PRAGMA journal_mode = WAL")
    conn.execute("PRAGMA synchronous = NORMAL")
    return conn


def insert_many(conn: sqlite3.Connection, table: str, rows: Sequence[Dict[str, Any]]) -> int:
    if not rows:
        return 0
    columns: List[str] = []
    for row in rows:
        for key in row:
            if key not in columns:
                columns.append(key)
    placeholders = ", ".join("?" for _ in columns)
    sql = f"INSERT INTO {table} ({', '.join(columns)}) VALUES ({placeholders})"
    values = [tuple(row.get(column) for column in columns) for row in rows]
    conn.executemany(sql, values)
    return len(rows)


def quote_ident(identifier: str) -> str:
    return '"' + identifier.replace('"', '""') + '"'


def table_exists(conn: sqlite3.Connection, table_name: str) -> bool:
    row = conn.execute(
        "SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = ?",
        (table_name,),
    ).fetchone()
    return row is not None


def dynamic_index_name(table_name: str, suffix: str) -> str:
    base = sanitize_slug(f"{table_name}_{suffix}")[:48]
    return f"{base}_{short_hash(f'{table_name}:{suffix}')}"


def build_sheet_code(sheet_name: str) -> str:
    slug = sanitize_slug(sheet_name) or "sheet"
    return f"{slug[:40]}_{short_hash(normalize_key(sheet_name))}"


def shortened_file_name(file_name: str) -> str:
    stem = Path(file_name).stem
    if stem.endswith("_processed"):
        return f"{stem[:-len('_processed')]}_"
    return f"{stem}_"


def sheet_data_table_name(folder_name: str, file_name: str, sheet_name: str) -> str:
    return sanitize_slug(f"{folder_name}_{shortened_file_name(file_name)}_{sheet_name}") or "sheet_data"


def build_sheet_targets(report_meta: ReportMeta, sheets: Sequence[SheetData]) -> List[SheetImportTarget]:
    return [
        SheetImportTarget(
            sheet_name=sheet.sheet_name,
            sheet_code=build_sheet_code(sheet.sheet_name),
            table_name=sheet_data_table_name(report_meta.folder_name, report_meta.file_name, sheet.sheet_name),
            sheet=sheet,
        )
        for sheet in sheets
    ]


def infer_measure_column_type(values: Sequence[Any]) -> str:
    numeric_values = [parse_numeric(value) for value in values if jsonable_value(value) is not None]
    if not numeric_values:
        return "NUMERIC"
    if all(value is not None and float(value).is_integer() for value in numeric_values):
        return "INTEGER"
    return "NUMERIC"


def infer_sheet_column_types(sheet: SheetData) -> Dict[str, str]:
    measure_set = set(sheet.measure_columns)
    column_types: Dict[str, str] = {}
    for column_name in sheet.column_names:
        values = [row.get(column_name) for row in sheet.rows]
        if column_name in measure_set:
            column_types[column_name] = infer_measure_column_type(values)
        else:
            column_types[column_name] = "TEXT"
    return column_types


def ensure_sheet_data_table(
    conn: sqlite3.Connection,
    table_name: str,
    column_types: Dict[str, str],
) -> None:
    quoted_table = quote_ident(table_name)
    idx_snapshot = quote_ident(dynamic_index_name(table_name, "snapshot"))
    idx_sheet = quote_ident(dynamic_index_name(table_name, "sheet"))

    if not table_exists(conn, table_name):
        fixed_columns = [f"{quote_ident(name)} {sql_type}" for name, sql_type in SYSTEM_COLUMNS]
        dynamic_columns = [
            f"{quote_ident(column_name)} {column_types.get(column_name, 'TEXT')}"
            for column_name in column_types
        ]
        conn.executescript(
            f"""
            CREATE TABLE IF NOT EXISTS {quoted_table} (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                {', '.join(fixed_columns + dynamic_columns)},
                FOREIGN KEY ({quote_ident(META_SNAPSHOT_ID)}) REFERENCES bao_cao_tong_hop_ngay(id) ON DELETE CASCADE,
                FOREIGN KEY ({quote_ident(META_SHEET_ID)}) REFERENCES sheet_bao_cao_tong_hop(id) ON DELETE CASCADE,
                UNIQUE ({quote_ident(META_SNAPSHOT_ID)}, {quote_ident(META_ROW_NUM)})
            );
            CREATE INDEX IF NOT EXISTS {idx_snapshot}
                ON {quoted_table}({quote_ident(META_SNAPSHOT_ID)});
            CREATE INDEX IF NOT EXISTS {idx_sheet}
                ON {quoted_table}({quote_ident(META_SHEET_ID)}, {quote_ident(META_ROW_NUM)});
            """
        )
        return

    existing_columns = {
        str(row["name"])
        for row in conn.execute(f"PRAGMA table_info({quoted_table})").fetchall()
    }
    for column_name, column_type in column_types.items():
        if column_name in existing_columns:
            continue
        conn.execute(
            f"ALTER TABLE {quoted_table} ADD COLUMN {quote_ident(column_name)} {column_type}"
        )


def insert_rows_with_quoted_columns(
    conn: sqlite3.Connection,
    table_name: str,
    rows: Sequence[Dict[str, Any]],
) -> int:
    if not rows:
        return 0
    columns: List[str] = []
    for row in rows:
        for key in row:
            if key not in columns:
                columns.append(key)
    placeholders = ", ".join("?" for _ in columns)
    quoted_columns = ", ".join(quote_ident(column) for column in columns)
    sql = f"INSERT INTO {quote_ident(table_name)} ({quoted_columns}) VALUES ({placeholders})"
    conn.executemany(sql, [tuple(row.get(column) for column in columns) for row in rows])
    return len(rows)


def archive_processed_file(workbook_path: Path, processed_root: Path, archive_root: Path, snapshot_date: date) -> Path:
    rel_path = workbook_path.relative_to(processed_root)
    archive_path = archive_root / snapshot_date.isoformat() / rel_path
    archive_path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(delete=False, dir=archive_path.parent, suffix=archive_path.suffix) as tmp:
        tmp_path = Path(tmp.name)
    shutil.copy2(workbook_path, tmp_path)
    tmp_path.replace(archive_path)
    return archive_path


def upsert_report_catalog(conn: sqlite3.Connection, report_meta: ReportMeta) -> None:
    now = current_timestamp()
    conn.execute(
        """
        INSERT INTO danh_muc_bao_cao_tong_hop (
            ma_bao_cao,
            ten_bao_cao,
            nhom_bao_cao,
            duong_dan_processed_mac_dinh,
            mo_ta,
            thoi_gian_tao,
            thoi_gian_cap_nhat
        )
        VALUES (?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(ma_bao_cao) DO UPDATE SET
            ten_bao_cao = excluded.ten_bao_cao,
            nhom_bao_cao = excluded.nhom_bao_cao,
            duong_dan_processed_mac_dinh = excluded.duong_dan_processed_mac_dinh,
            mo_ta = excluded.mo_ta,
            dang_su_dung = 1,
            thoi_gian_cap_nhat = excluded.thoi_gian_cap_nhat
        """,
        (
            report_meta.report_code,
            report_meta.report_name,
            report_meta.report_group,
            report_meta.processed_rel_path,
            "Auto-import du lieu tong hop theo tung sheet processed",
            now,
            now,
        ),
    )


def upsert_sheet_table_catalog(
    conn: sqlite3.Connection,
    report_meta: ReportMeta,
    target: SheetImportTarget,
) -> None:
    now = current_timestamp()
    conn.execute(
        """
        INSERT INTO danh_muc_bang_du_lieu_bao_cao (
            ma_bao_cao,
            ten_sheet_goc,
            ma_sheet,
            ten_bang_du_lieu,
            che_do_luu_tru,
            tong_so_cot,
            danh_sach_cot_json,
            mo_ta,
            thoi_gian_tao,
            thoi_gian_cap_nhat
        )
        VALUES (?, ?, ?, ?, 'processed_summary_sheet_only', ?, ?, ?, ?, ?)
        ON CONFLICT(ma_bao_cao, ten_sheet_goc) DO UPDATE SET
            ma_sheet = excluded.ma_sheet,
            ten_bang_du_lieu = excluded.ten_bang_du_lieu,
            che_do_luu_tru = excluded.che_do_luu_tru,
            tong_so_cot = excluded.tong_so_cot,
            danh_sach_cot_json = excluded.danh_sach_cot_json,
            mo_ta = excluded.mo_ta,
            thoi_gian_cap_nhat = excluded.thoi_gian_cap_nhat
        """,
        (
            report_meta.report_code,
            target.sheet_name,
            target.sheet_code,
            target.table_name,
            len(target.sheet.column_names),
            json.dumps(target.sheet.column_names, ensure_ascii=False),
            f"Bang du lieu cho sheet {target.sheet_name} cua bao cao {report_meta.report_code}",
            now,
            now,
        ),
    )


def start_import_log(
    conn: sqlite3.Connection,
    report_meta: ReportMeta,
    snapshot_date: date,
    source_hash: str,
    table_names: Sequence[str],
) -> int:
    cursor = conn.execute(
        """
        INSERT INTO nhat_ky_nap_tong_hop (
            ma_bao_cao,
            ngay_du_lieu,
            so_bang_du_lieu,
            danh_sach_bang_du_lieu,
            che_do_nap,
            bat_dau_luc,
            trang_thai,
            ma_hash_tep
        )
        VALUES (?, ?, ?, ?, 'summary_per_report_sheet_overwrite_same_day', ?, 'dang_chay', ?)
        """,
        (
            report_meta.report_code,
            snapshot_date.isoformat(),
            len(table_names),
            ", ".join(table_names) if table_names else None,
            current_timestamp(),
            source_hash,
        ),
    )
    conn.commit()
    return int(cursor.lastrowid)


def finish_import_log(
    conn: sqlite3.Connection,
    log_id: int,
    *,
    status: str,
    message: Optional[str],
    table_names: Sequence[str],
    sheet_count: int = 0,
    row_count: int = 0,
    metric_count: int = 0,
) -> None:
    conn.execute(
        """
        UPDATE nhat_ky_nap_tong_hop
        SET ket_thuc_luc = ?,
            trang_thai = ?,
            thong_diep = ?,
            so_bang_du_lieu = ?,
            danh_sach_bang_du_lieu = ?,
            so_sheet_tong_hop = ?,
            so_dong_tong_hop = ?,
            so_chi_tieu_tong_hop = ?
        WHERE id = ?
        """,
        (
            current_timestamp(),
            status,
            message,
            len(table_names),
            ", ".join(table_names) if table_names else None,
            sheet_count,
            row_count,
            metric_count,
            log_id,
        ),
    )
    conn.commit()


def get_existing_report_day(conn: sqlite3.Connection, report_code: str, snapshot_date: date) -> Optional[sqlite3.Row]:
    return conn.execute(
        """
        SELECT id, ma_hash_tep
        FROM bao_cao_tong_hop_ngay
        WHERE ma_bao_cao = ? AND ngay_du_lieu = ?
        """,
        (report_code, snapshot_date.isoformat()),
    ).fetchone()


def prepare_report_day(
    conn: sqlite3.Connection,
    report_meta: ReportMeta,
    snapshot_date: date,
    period_start: Optional[date],
    period_end: Optional[date],
    report_month: Optional[int],
    report_year: Optional[int],
    source_hash: str,
    file_path: str,
    file_name: str,
    existing_row: Optional[sqlite3.Row],
) -> int:
    if existing_row:
        conn.execute("DELETE FROM bao_cao_tong_hop_ngay WHERE id = ?", (int(existing_row["id"]),))
    cursor = conn.execute(
        """
        INSERT INTO bao_cao_tong_hop_ngay (
            ma_bao_cao,
            ngay_du_lieu,
            tu_ngay,
            den_ngay,
            thang_bao_cao,
            nam_bao_cao,
            ten_tep_nguon,
            duong_dan_tep_nguon,
            ma_hash_tep,
            trang_thai_nap,
            thoi_gian_tao,
            thoi_gian_cap_nhat
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'cho_xu_ly', ?, ?)
        """,
        (
            report_meta.report_code,
            snapshot_date.isoformat(),
            period_start.isoformat() if period_start else None,
            period_end.isoformat() if period_end else None,
            report_month,
            report_year,
            file_name,
            file_path,
            source_hash,
            current_timestamp(),
            current_timestamp(),
        ),
    )
    return int(cursor.lastrowid)


def insert_file_records(
    conn: sqlite3.Connection,
    report_day_id: int,
    processed_path: Path,
    archive_path: Optional[Path],
    source_hash: str,
    size_bytes: int,
) -> None:
    processed_stat = processed_path.stat()
    rows = [
        {
            "bao_cao_tong_hop_ngay_id": report_day_id,
            "loai_tep": "processed",
            "ten_tep": processed_path.name,
            "duong_dan_tep": str(processed_path.relative_to(API_TRANSITION_DIR)),
            "la_tep_nap_chinh": 1,
            "kich_thuoc_byte": size_bytes,
            "ma_hash_tep": source_hash,
            "thoi_gian_sua_file": datetime.fromtimestamp(processed_stat.st_mtime).isoformat(sep=" ", timespec="seconds"),
            "thoi_gian_tao": current_timestamp(),
        }
    ]
    if archive_path is not None:
        archive_stat = archive_path.stat()
        rows.append(
            {
                "bao_cao_tong_hop_ngay_id": report_day_id,
                "loai_tep": "daily",
                "ten_tep": archive_path.name,
                "duong_dan_tep": str(archive_path.relative_to(API_TRANSITION_DIR)),
                "la_tep_nap_chinh": 0,
                "kich_thuoc_byte": size_bytes,
                "ma_hash_tep": source_hash,
                "thoi_gian_sua_file": datetime.fromtimestamp(archive_stat.st_mtime).isoformat(sep=" ", timespec="seconds"),
                "thoi_gian_tao": current_timestamp(),
            }
        )
    insert_many(conn, "tep_nguon_bao_cao_tong_hop", rows)


def select_importable_processed_sheets(
    report_meta: ReportMeta,
    sheets: Sequence[SheetData],
    allowlist: Dict[Tuple[str, str], set[str]],
) -> List[SheetData]:
    allowed_sheet_names = get_allowed_sheet_names(report_meta, allowlist)
    return [
        sheet
        for sheet in sheets
        if normalize_key(sheet.sheet_name) in allowed_sheet_names
    ]


def get_dimension_columns(sheet: SheetData) -> List[str]:
    return [column for column in sheet.column_names if column not in sheet.measure_columns]


def row_has_values(record: Dict[str, Any]) -> bool:
    return any(jsonable_value(value) is not None for value in record.values())


def measure_value_count(record: Dict[str, Any], measure_columns: Sequence[str]) -> int:
    count = 0
    for column_name in measure_columns:
        if jsonable_value(record.get(column_name)) is not None:
            count += 1
    return count


def analyse_importable_sheets(importable_sheets: Sequence[SheetData]) -> Tuple[int, int, int, int]:
    sheet_count = len(importable_sheets)
    table_count = len(importable_sheets)
    row_count = 0
    metric_count = 0
    for sheet in importable_sheets:
        for record in sheet.rows:
            if not row_has_values(record):
                continue
            row_count += 1
            metric_count += measure_value_count(record, sheet.measure_columns)
    return sheet_count, table_count, row_count, metric_count


def insert_sheet_rows(
    conn: sqlite3.Connection,
    report_meta: ReportMeta,
    report_day_id: int,
    targets: Sequence[SheetImportTarget],
) -> Tuple[int, int, int, int, List[str]]:
    total_sheet_count = 0
    total_table_count = 0
    total_row_count = 0
    total_metric_count = 0
    imported_tables: List[str] = []

    for target in targets:
        sheet = target.sheet
        dimension_columns = get_dimension_columns(sheet)
        column_types = infer_sheet_column_types(sheet)
        ensure_sheet_data_table(conn, target.table_name, column_types)
        upsert_sheet_table_catalog(conn, report_meta, target)

        sheet_cursor = conn.execute(
            """
            INSERT INTO sheet_bao_cao_tong_hop (
                bao_cao_tong_hop_ngay_id,
                ten_sheet,
                ma_sheet,
                ten_bang_du_lieu,
                thu_tu_sheet,
                tong_so_cot,
                so_dong_tong_hop,
                so_chi_tieu_tong_hop,
                danh_sach_cot_json,
                cot_chieu_json,
                cot_chi_tieu_json,
                thoi_gian_tao
            )
            VALUES (?, ?, ?, ?, ?, ?, 0, 0, ?, ?, ?, ?)
            """,
            (
                report_day_id,
                target.sheet_name,
                target.sheet_code,
                target.table_name,
                sheet.sheet_order,
                len(sheet.column_names),
                json.dumps(sheet.column_names, ensure_ascii=False),
                json.dumps(dimension_columns, ensure_ascii=False),
                json.dumps(sheet.measure_columns, ensure_ascii=False),
                current_timestamp(),
            ),
        )
        sheet_id = int(sheet_cursor.lastrowid)

        rows_to_insert: List[Dict[str, Any]] = []
        imported_row_count = 0
        imported_metric_count = 0

        for row_number, record in enumerate(sheet.rows, start=1):
            if not row_has_values(record):
                continue
            full_row_json = json.dumps(record, ensure_ascii=False, sort_keys=True)
            row_payload: Dict[str, Any] = {
                META_SNAPSHOT_ID: report_day_id,
                META_SHEET_ID: sheet_id,
                META_ROW_NUM: row_number,
                META_ROW_HASH: sha256_bytes(f"{target.sheet_name}|{row_number}|{full_row_json}".encode("utf-8")),
                META_IMPORTED_AT: current_timestamp(),
            }
            for column_name in sheet.column_names:
                row_payload[column_name] = jsonable_value(record.get(column_name))
            rows_to_insert.append(row_payload)
            imported_row_count += 1
            imported_metric_count += measure_value_count(record, sheet.measure_columns)

        insert_rows_with_quoted_columns(conn, target.table_name, rows_to_insert)
        conn.execute(
            """
            UPDATE sheet_bao_cao_tong_hop
            SET so_dong_tong_hop = ?,
                so_chi_tieu_tong_hop = ?
            WHERE id = ?
            """,
            (imported_row_count, imported_metric_count, sheet_id),
        )

        imported_tables.append(target.table_name)
        total_sheet_count += 1
        total_table_count += 1
        total_row_count += imported_row_count
        total_metric_count += imported_metric_count

    return total_sheet_count, total_table_count, total_row_count, total_metric_count, imported_tables


def import_workbook(
    conn: sqlite3.Connection,
    workbook_path: Path,
    processed_root: Path,
    archive_root: Path,
    snapshot_date: date,
    period_start: Optional[date],
    period_end: Optional[date],
    report_month: Optional[int],
    report_year: Optional[int],
    *,
    dry_run: bool,
    skip_archive: bool,
    skip_if_same_hash: bool,
    pre_archived_path: Optional[Path] = None,
) -> Dict[str, Any]:
    if is_skipped_processed_workbook(processed_root, workbook_path):
        rel_path = str(workbook_path.relative_to(processed_root)).replace("\\", "/")
        return {
            "report_code": None,
            "report_name": workbook_path.stem,
            "processed_rel_path": rel_path,
            "snapshot_date": snapshot_date.isoformat(),
            "file_sha256": None,
            "data_tables": [],
            "data_table_count": 0,
            "raw_rows": 0,
            "summary_rows": 0,
            "detail_rows": 0,
            "business_rows": 0,
            "summary_sheet_count": 0,
            "metric_count": 0,
            "status": "skipped",
            "archive_path": None,
            "error": "Removed from pipeline.",
        }

    report_meta = build_report_meta(processed_root, workbook_path)
    allowlist = load_import_allowlist()
    allowed_sheet_names = get_allowed_sheet_names(report_meta, allowlist)
    if not allowed_sheet_names:
        return {
            "report_code": report_meta.report_code,
            "report_name": report_meta.report_name,
            "processed_rel_path": report_meta.processed_rel_path,
            "snapshot_date": snapshot_date.isoformat(),
            "file_sha256": None,
            "data_tables": [],
            "data_table_count": 0,
            "raw_rows": 0,
            "summary_rows": 0,
            "detail_rows": 0,
            "business_rows": 0,
            "summary_sheet_count": 0,
            "metric_count": 0,
            "status": "skipped",
            "archive_path": None,
            "error": "Workbook not listed in list of table.csv.",
        }

    workbook_bytes = workbook_path.read_bytes()
    source_hash = sha256_bytes(workbook_bytes)
    archive_path = pre_archived_path
    if not dry_run and not skip_archive and archive_path is None:
        archive_path = archive_processed_file(workbook_path, processed_root, archive_root, snapshot_date)

    sheets = read_workbook_sheets(workbook_path)
    importable_sheets = select_importable_processed_sheets(report_meta, sheets, allowlist)
    targets = build_sheet_targets(report_meta, importable_sheets)
    table_names = [target.table_name for target in targets]
    sheet_count, table_count, summary_row_count, metric_count = analyse_importable_sheets(importable_sheets)

    result = {
        "report_code": report_meta.report_code,
        "report_name": report_meta.report_name,
        "processed_rel_path": report_meta.processed_rel_path,
        "snapshot_date": snapshot_date.isoformat(),
        "file_sha256": source_hash,
        "data_tables": table_names,
        "data_table_count": table_count,
        "raw_rows": 0,
        "summary_rows": summary_row_count,
        "detail_rows": 0,
        "business_rows": metric_count,
        "summary_sheet_count": sheet_count,
        "metric_count": metric_count,
        "status": "dry_run" if dry_run else "pending",
        "archive_path": str(archive_path.relative_to(API_TRANSITION_DIR)) if archive_path is not None else None,
    }

    if dry_run:
        return result

    existing_row = get_existing_report_day(conn, report_meta.report_code, snapshot_date)
    log_id = start_import_log(conn, report_meta, snapshot_date, source_hash, table_names)

    if skip_if_same_hash and existing_row and existing_row["ma_hash_tep"] == source_hash:
        finish_import_log(
            conn,
            log_id,
            status="bo_qua",
            message="Bo qua vi cung ngay va cung hash file.",
            table_names=table_names,
            sheet_count=sheet_count,
            row_count=summary_row_count,
            metric_count=metric_count,
        )
        result["status"] = "skipped"
        return result

    try:
        conn.execute("BEGIN")
        upsert_report_catalog(conn, report_meta)
        file_path = str((archive_path or workbook_path).relative_to(API_TRANSITION_DIR))
        report_day_id = prepare_report_day(
            conn,
            report_meta,
            snapshot_date,
            period_start,
            period_end,
            report_month,
            report_year,
            source_hash,
            file_path,
            (archive_path or workbook_path).name,
            existing_row,
        )
        insert_file_records(
            conn,
            report_day_id,
            workbook_path,
            archive_path,
            source_hash,
            len(workbook_bytes),
        )

        if not targets:
            conn.execute(
                """
                UPDATE bao_cao_tong_hop_ngay
                SET so_sheet_tong_hop = 0,
                    so_bang_du_lieu = 0,
                    so_dong_tong_hop = 0,
                    so_chi_tieu_tong_hop = 0,
                    trang_thai_nap = 'khong_co_sheet_tong_hop',
                    ghi_chu = ?,
                    thoi_gian_cap_nhat = ?
                WHERE id = ?
                """,
                ("Khong tim thay sheet sau xu ly hop le de import.", current_timestamp(), report_day_id),
            )
            conn.commit()
            finish_import_log(
                conn,
                log_id,
                status="khong_co_sheet_tong_hop",
                message="Khong tim thay sheet sau xu ly de import.",
                table_names=table_names,
                sheet_count=0,
                row_count=0,
                metric_count=0,
            )
            result["status"] = "no_summary_sheets"
            return result

        imported_sheet_count, imported_table_count, imported_row_count, imported_metric_count, imported_tables = (
            insert_sheet_rows(conn, report_meta, report_day_id, targets)
        )
        conn.execute(
            """
            UPDATE bao_cao_tong_hop_ngay
            SET so_sheet_tong_hop = ?,
                so_bang_du_lieu = ?,
                so_dong_tong_hop = ?,
                so_chi_tieu_tong_hop = ?,
                trang_thai_nap = 'thanh_cong',
                ghi_chu = ?,
                thoi_gian_cap_nhat = ?
            WHERE id = ?
            """,
            (
                imported_sheet_count,
                imported_table_count,
                imported_row_count,
                imported_metric_count,
                f"Da import {imported_sheet_count} sheet vao {imported_table_count} bang du lieu.",
                current_timestamp(),
                report_day_id,
            ),
        )
        conn.commit()
        finish_import_log(
            conn,
            log_id,
            status="thanh_cong",
            message=f"Thanh cong. tables={imported_table_count} metric_count={imported_metric_count}",
            table_names=imported_tables,
            sheet_count=imported_sheet_count,
            row_count=imported_row_count,
            metric_count=imported_metric_count,
        )
        result["status"] = "imported"
        result["summary_rows"] = imported_row_count
        result["business_rows"] = imported_metric_count
        result["metric_count"] = imported_metric_count
        result["summary_sheet_count"] = imported_sheet_count
        result["data_tables"] = imported_tables
        result["data_table_count"] = imported_table_count
        return result
    except Exception as exc:
        conn.rollback()
        finish_import_log(
            conn,
            log_id,
            status="that_bai",
            message=str(exc),
            table_names=table_names,
            sheet_count=sheet_count,
            row_count=summary_row_count,
            metric_count=metric_count,
        )
        raise


def iter_workbooks(processed_root: Path, path_filters: Sequence[str]) -> Iterable[Path]:
    allowlist = load_import_allowlist()
    lowered_filters = [value.lower() for value in path_filters if value]
    for workbook_path in sorted(processed_root.rglob("*.xlsx")):
        rel_path = str(workbook_path.relative_to(processed_root)).replace("\\", "/").lower()
        if rel_path in SKIPPED_PROCESSED_REL_PATHS:
            continue
        report_meta = build_report_meta(processed_root, workbook_path)
        if not get_allowed_sheet_names(report_meta, allowlist):
            continue
        if lowered_filters and not all(value in rel_path for value in lowered_filters):
            continue
        yield workbook_path


def ensure_database(db_path: Path) -> None:
    from api_transition.sqlite_history.apply_report_history_views import apply_views
    from api_transition.sqlite_history.init_report_history_db import DEFAULT_SCHEMA_PATH, DEFAULT_VIEWS_PATH, init_database

    init_database(db_path, DEFAULT_SCHEMA_PATH, DEFAULT_VIEWS_PATH, reset=False)
    apply_views(db_path, DEFAULT_VIEWS_PATH)


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = parse_args(argv)
    db_path = Path(args.db_path).expanduser().resolve()
    processed_root = Path(args.processed_root).expanduser().resolve()
    archive_root = Path(args.archive_root).expanduser().resolve()
    period_start = parse_optional_date(args.period_start)
    period_end = parse_optional_date(args.period_end)
    results: List[Dict[str, Any]] = []

    if not processed_root.exists():
        raise FileNotFoundError(f"Khong tim thay Processed root: {processed_root}")

    ensure_database(db_path)
    conn = connect_sqlite(db_path)
    try:
        for workbook_path in iter_workbooks(processed_root, args.path_contains):
            snapshot_date = parse_snapshot_date(args.snapshot_date, workbook_path)
            result = import_workbook(
                conn,
                workbook_path,
                processed_root,
                archive_root,
                snapshot_date,
                period_start,
                period_end,
                args.report_month or snapshot_date.month,
                args.report_year or snapshot_date.year,
                dry_run=args.dry_run,
                skip_archive=args.skip_archive,
                skip_if_same_hash=args.skip_if_same_hash,
            )
            results.append(result)
    finally:
        conn.close()

    output = {
        "db_path": str(db_path),
        "processed_root": str(processed_root),
        "archive_root": str(archive_root),
        "count": len(results),
        "results": results,
    }
    if args.json:
        print(json.dumps(output, ensure_ascii=False, indent=2))
    else:
        print(f"Da xu ly {len(results)} workbook")
        for item in results:
            print(
                f"- {item['processed_rel_path']}: status={item['status']} "
                f"tables={item['data_table_count']} sheets={item['summary_sheet_count']} "
                f"rows={item['summary_rows']} metrics={item['metric_count']}"
            )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
