#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Import workbook processed vao SQLite voi co che ghi de cung ngay."""

from __future__ import annotations

import argparse
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


MODULE_DIR = Path(__file__).resolve().parent
API_TRANSITION_DIR = MODULE_DIR.parent
DEFAULT_DB_PATH = API_TRANSITION_DIR / "report_history.db"
DEFAULT_PROCESSED_ROOT = API_TRANSITION_DIR / "Processed"
DEFAULT_ARCHIVE_ROOT = API_TRANSITION_DIR / "ProcessedDaily"
DATE_IN_NAME_RE = re.compile(r"(\d{8})")

SOURCE_SHEET_NAMES = {"sheet", "sheet1"}
DETAIL_SHEET_NAMES = {
    "data",
    "data_combined",
    "data_tam_dung",
    "data_khoi_phuc",
    "chi tiết vật tư",
    "chi_tiet_chua_khoi_phuc",
    "chi_tiet",
    "chi tiet",
}
NOTE_SHEET_NAMES = {"thong_bao"}
SUMMARY_NAME_HINTS = (
    "th_",
    "tong_hop",
    "tổng hợp",
    "kq_",
    "thuc_tang",
    "thang",
    "fiber_",
    "mytv_",
    "c11 kpi nvkt",
    "c12 kpi nvkt",
    "c13 kpi nvkt",
    "du_lieu_sach",
)
DETAIL_NAME_HINTS = ("khong_dat", "chi tiết vật tư", "chi_tiet")
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
}
TTVT_KEYS = (
    "ttvt",
    "ten_ttvt",
    "trung tâm viễn thông",
    "trung tam vien thong",
)
TEAM_KEYS = (
    "doivt",
    "doi vt",
    "đội vt",
    "đội viễn thông",
    "doi vien thong",
    "ten_doi",
    "diemchia",
    "nhom_dia_ban",
    "nhóm địa bàn",
)
UNIT_KEYS = (
    "đơn vị",
    "don vi",
    "đơn vị/nhân viên kt",
)
EMPLOYEE_CODE_KEYS = (
    "mã nv",
    "ma nv",
    "mã nhân viên",
    "ma nhan vien",
)
EMPLOYEE_NAME_KEYS = (
    "nvkt",
    "ten nv",
    "tên nv",
    "nhanvien_kt",
    "nhân viên kt",
    "nhanvien_thu",
    "nhanvien_nhapkho",
    "nvkt_diaban_giao",
    "mã nhân viên",
    "mã nv",
    "ma nv",
    "ma nhan vien",
)
ENTITY_KEY_KEYS = (
    "ma_tb",
    "baohong_id",
    "hdtb_id",
    "ma_men",
    "ma_gd",
    "ma_vt",
    "serial number",
    "serial",
    "ma_tkyt",
)
DATETIME_KEYS = (
    "ngay_thuchien",
    "ngay_hoan_thanh",
    "ngay_nghiem_thu",
    "ngay_hc",
    "ngay_yc",
    "ngay_td",
    "ngay_huy",
    "ngaylap_hd",
    "ngay_ins",
    "thời gian cập nhật",
    "thoi gian cap nhat",
)


@dataclass(frozen=True)
class ReportMeta:
    report_code: str
    report_name: str
    report_group: str
    processed_rel_path: str


@dataclass(frozen=True)
class SourceField:
    target: str
    sources: Tuple[str, ...]
    kind: str = "text"


@dataclass
class SheetData:
    sheet_name: str
    sheet_order: int
    sheet_kind: str
    row_count: int
    column_names: List[str]
    numeric_columns: List[str]
    rows: List[Dict[str, Any]]


def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Import workbook processed vao SQLite report_history.db")
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


def sanitize_slug(value: str) -> str:
    text = value.strip().lower()
    text = re.sub(r"_processed$", "", text)
    text = re.sub(r"\.xlsx$", "", text)
    text = re.sub(r"[^a-z0-9]+", "_", text)
    text = re.sub(r"_+", "_", text).strip("_")
    return text


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


def normalize_key(value: str) -> str:
    return re.sub(r"\s+", " ", strip_accents(str(value).strip())).strip()


def sha256_bytes(content: bytes) -> str:
    return hashlib.sha256(content).hexdigest()


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


def build_report_meta(processed_root: Path, workbook_path: Path) -> ReportMeta:
    rel_path = workbook_path.relative_to(processed_root)
    rel_text = str(rel_path).replace("\\", "/")
    group = rel_path.parts[0]
    stem = workbook_path.stem
    stem = re.sub(r"_processed$", "", stem)
    clean_stem = DATE_IN_NAME_RE.sub("", stem).replace("__", "_").strip("_ ")
    report_code = sanitize_slug(f"{group}_{clean_stem}")
    report_name = clean_stem.replace("_", " ").strip() or stem
    return ReportMeta(
        report_code=report_code,
        report_name=report_name,
        report_group=group,
        processed_rel_path=rel_text,
    )


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


def parse_numeric(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, numbers.Number) and not isinstance(value, bool) and not pd.isna(value):
        return float(value)
    text = str(value).strip()
    if not text:
        return None
    text = text.replace("%", "").replace(",", "")
    if re.fullmatch(r"-?\d+(?:\.\d+)?", text):
        return float(text)
    return None


def parse_datetime_value(value: Any) -> Optional[datetime]:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime.combine(value, datetime.min.time())
    text = str(value).strip()
    if not text:
        return None
    parse_candidates = [False, True] if re.search(r"\b(am|pm)\b", text, flags=re.IGNORECASE) else [True, False]
    parsed = pd.NaT
    for dayfirst in parse_candidates:
        parsed = pd.to_datetime(text, errors="coerce", dayfirst=dayfirst)
        if not pd.isna(parsed):
            break
    if pd.isna(parsed):
        return None
    return parsed.to_pydatetime()


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
        return value
    if isinstance(value, numbers.Integral):
        return int(value)
    if isinstance(value, numbers.Real):
        if float(value).is_integer():
            return int(value)
        return float(value)
    text = str(value).strip()
    return text or None


def detect_numeric_columns(df: pd.DataFrame) -> List[str]:
    numeric_columns: List[str] = []
    for column in df.columns:
        if sanitize_slug(column) in METRIC_EXCLUDE_COLUMNS:
            continue
        series = df[column].dropna()
        if series.empty:
            continue
        parsed = series.map(parse_numeric)
        if len(series) > 0 and parsed.notna().sum() / len(series) >= 0.8:
            numeric_columns.append(column)
    return numeric_columns


def classify_sheet(sheet_name: str, df: pd.DataFrame, numeric_columns: Sequence[str]) -> str:
    lower_name = normalize_key(sheet_name)
    row_count = len(df.index)
    col_count = len(df.columns)
    if lower_name in SOURCE_SHEET_NAMES:
        return "source"
    if lower_name in NOTE_SHEET_NAMES:
        return "note"
    if lower_name in {normalize_key(value) for value in DETAIL_SHEET_NAMES}:
        return "detail"
    if any(token in lower_name for token in DETAIL_NAME_HINTS):
        return "detail"
    if "chi_tiet" in lower_name or "chi tiet" in lower_name:
        if row_count <= 500 and col_count <= 12 and numeric_columns:
            return "summary"
        return "detail"
    if any(token in lower_name for token in SUMMARY_NAME_HINTS):
        return "summary"
    if row_count <= 200 and numeric_columns:
        return "summary"
    return "other"


def read_workbook_sheets(workbook_path: Path) -> List[SheetData]:
    excel = pd.ExcelFile(workbook_path, engine="openpyxl")
    results: List[SheetData] = []
    for sheet_order, sheet_name in enumerate(excel.sheet_names, start=1):
        df = excel.parse(sheet_name=sheet_name, dtype=object)
        df = drop_empty_records(df.copy())
        df.columns = clean_column_names(df.columns)
        numeric_columns = detect_numeric_columns(df)
        rows = [{key: jsonable_value(value) for key, value in row.items()} for row in df.to_dict(orient="records")]
        results.append(
            SheetData(
                sheet_name=sheet_name,
                sheet_order=sheet_order,
                sheet_kind=classify_sheet(sheet_name, df, numeric_columns),
                row_count=len(df.index),
                column_names=list(df.columns),
                numeric_columns=numeric_columns,
                rows=rows,
            )
        )
    return results


def first_value(record: Dict[str, Any], candidates: Sequence[str]) -> Tuple[Optional[str], Optional[str]]:
    normalized = {normalize_key(key): key for key in record}
    for candidate in candidates:
        key = normalized.get(normalize_key(candidate))
        if not key:
            continue
        value = jsonable_value(record.get(key))
        if value is None:
            continue
        text = str(value).strip()
        if text:
            return text, key
    return None, None


def infer_entity_key(record: Dict[str, Any]) -> Optional[str]:
    value, _ = first_value(record, ENTITY_KEY_KEYS)
    return value


def extract_context(record: Dict[str, Any]) -> Dict[str, Optional[str]]:
    ttvt, _ = first_value(record, TTVT_KEYS)
    team_name, _ = first_value(record, TEAM_KEYS)
    unit_name, _ = first_value(record, UNIT_KEYS)
    employee_code, _ = first_value(record, EMPLOYEE_CODE_KEYS)
    employee_name, _ = first_value(record, EMPLOYEE_NAME_KEYS)
    return {
        "ttvt": ttvt,
        "doi_vien_thong": team_name,
        "don_vi": unit_name,
        "ma_nv": employee_code,
        "ten_nv": employee_name,
    }


def infer_occurred_at(record: Dict[str, Any]) -> Optional[datetime]:
    normalized = {normalize_key(key): key for key in record}
    for candidate in DATETIME_KEYS:
        key = normalized.get(candidate)
        if not key:
            continue
        parsed = parse_datetime_value(record.get(key))
        if parsed:
            return parsed
    return None


def text_value(value: Any) -> Optional[str]:
    normalized = jsonable_value(value)
    if normalized is None:
        return None
    if isinstance(normalized, (int, float)):
        return str(normalized)
    return str(normalized).strip() or None


def int_value(value: Any) -> Optional[int]:
    parsed = parse_numeric(value)
    if parsed is None:
        return None
    return int(parsed)


def float_value(value: Any) -> Optional[float]:
    return parse_numeric(value)


def date_value(value: Any) -> Optional[str]:
    parsed = parse_datetime_value(value)
    if not parsed:
        return None
    return parsed.date().isoformat()


def datetime_value(value: Any) -> Optional[str]:
    parsed = parse_datetime_value(value)
    if not parsed:
        return None
    return parsed.isoformat(sep=" ", timespec="seconds")


def convert_value(value: Any, kind: str) -> Any:
    if kind == "int":
        return int_value(value)
    if kind == "float":
        return float_value(value)
    if kind == "date":
        return date_value(value)
    if kind == "datetime":
        return datetime_value(value)
    return text_value(value)


def build_extra_json(record: Dict[str, Any], used_keys: Sequence[str], extra: Optional[Dict[str, Any]] = None) -> Optional[str]:
    used = {normalize_key(key) for key in used_keys}
    payload: Dict[str, Any] = {}
    for key, value in record.items():
        if normalize_key(key) in used:
            continue
        normalized = jsonable_value(value)
        if normalized is None:
            continue
        payload[key] = normalized
    if extra:
        for key, value in extra.items():
            if value is not None:
                payload[key] = value
    if not payload:
        return None
    return json.dumps(payload, ensure_ascii=False, sort_keys=True)


def map_record(record: Dict[str, Any], fields: Sequence[SourceField]) -> Tuple[Dict[str, Any], List[str]]:
    normalized = {normalize_key(key): key for key in record}
    mapped: Dict[str, Any] = {}
    used_keys: List[str] = []
    for field in fields:
        chosen_key: Optional[str] = None
        for source in field.sources:
            chosen_key = normalized.get(normalize_key(source))
            if chosen_key:
                break
        mapped[field.target] = convert_value(record.get(chosen_key) if chosen_key else None, field.kind)
        if chosen_key:
            used_keys.append(chosen_key)
    return mapped, used_keys


def find_sheet(sheets: Sequence[SheetData], *names: str) -> Optional[SheetData]:
    normalized_targets = {normalize_key(name) for name in names}
    for sheet in sheets:
        if normalize_key(sheet.sheet_name) in normalized_targets:
            return sheet
    return None


def find_sheets_by_prefix(sheets: Sequence[SheetData], prefixes: Sequence[str]) -> List[SheetData]:
    normalized_prefixes = [normalize_key(prefix) for prefix in prefixes]
    return [sheet for sheet in sheets if any(normalize_key(sheet.sheet_name).startswith(prefix) for prefix in normalized_prefixes)]


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


def insert_or_ignore_many(conn: sqlite3.Connection, table: str, rows: Sequence[Dict[str, Any]]) -> int:
    if not rows:
        return 0
    columns: List[str] = []
    for row in rows:
        for key in row:
            if key not in columns:
                columns.append(key)
    placeholders = ", ".join("?" for _ in columns)
    sql = f"INSERT OR IGNORE INTO {table} ({', '.join(columns)}) VALUES ({placeholders})"
    values = [tuple(row.get(column) for column in columns) for row in rows]
    conn.executemany(sql, values)
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
        INSERT INTO danh_muc_bao_cao (
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
            f"Auto-import tu {report_meta.processed_rel_path}",
            now,
            now,
        ),
    )


def start_import_log(
    conn: sqlite3.Connection,
    report_meta: ReportMeta,
    snapshot_date: date,
    source_hash: str,
) -> int:
    cursor = conn.execute(
        """
        INSERT INTO nhat_ky_nap_bao_cao (
            ma_bao_cao,
            ngay_du_lieu,
            che_do_ghi_de,
            bat_dau_luc,
            trang_thai,
            ma_hash_tep
        )
        VALUES (?, ?, 'overwrite_cung_ngay', ?, 'dang_chay', ?)
        """,
        (report_meta.report_code, snapshot_date.isoformat(), current_timestamp(), source_hash),
    )
    conn.commit()
    return int(cursor.lastrowid)


def finish_import_log(
    conn: sqlite3.Connection,
    log_id: int,
    *,
    status: str,
    message: Optional[str],
    raw_count: int = 0,
    summary_count: int = 0,
    detail_count: int = 0,
) -> None:
    conn.execute(
        """
        UPDATE nhat_ky_nap_bao_cao
        SET ket_thuc_luc = ?,
            trang_thai = ?,
            thong_diep = ?,
            so_dong_goc = ?,
            so_dong_tong_hop = ?,
            so_dong_chi_tiet = ?
        WHERE id = ?
        """,
        (current_timestamp(), status, message, raw_count, summary_count, detail_count, log_id),
    )
    conn.commit()


def get_existing_report_day(conn: sqlite3.Connection, report_code: str, snapshot_date: date) -> Optional[sqlite3.Row]:
    return conn.execute(
        "SELECT id, ma_hash_tep FROM bao_cao_ngay WHERE ma_bao_cao = ? AND ngay_du_lieu = ?",
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
    existing_row: Optional[sqlite3.Row],
) -> int:
    if existing_row:
        conn.execute("DELETE FROM bao_cao_ngay WHERE id = ?", (int(existing_row["id"]),))
    cursor = conn.execute(
        """
        INSERT INTO bao_cao_ngay (
            ma_bao_cao,
            ngay_du_lieu,
            tu_ngay,
            den_ngay,
            thang_bao_cao,
            nam_bao_cao,
            duong_dan_tep,
            ma_hash_tep,
            trang_thai_nap,
            thoi_gian_tao,
            thoi_gian_cap_nhat
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'cho_xu_ly', ?, ?)
        """,
        (
            report_meta.report_code,
            snapshot_date.isoformat(),
            period_start.isoformat() if period_start else None,
            period_end.isoformat() if period_end else None,
            report_month,
            report_year,
            file_path,
            source_hash,
            current_timestamp(),
            current_timestamp(),
        ),
    )
    return int(cursor.lastrowid)


def insert_sheet_and_raw_rows(
    conn: sqlite3.Connection,
    report_day_id: int,
    report_meta: ReportMeta,
    sheets: Sequence[SheetData],
) -> Tuple[int, int, int]:
    raw_rows: List[Dict[str, Any]] = []
    sheet_rows: List[Dict[str, Any]] = []
    don_vi_rows: Dict[str, Dict[str, Any]] = {}
    nhan_vien_rows: Dict[str, Dict[str, Any]] = {}
    raw_count = 0
    summary_count = 0
    detail_count = 0

    for sheet in sheets:
        cursor = conn.execute(
            """
            INSERT INTO sheet_bao_cao (
                bao_cao_ngay_id,
                ten_sheet,
                loai_sheet,
                thu_tu_sheet,
                so_dong,
                danh_sach_cot_json,
                thoi_gian_tao
            )
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (
                report_day_id,
                sheet.sheet_name,
                sheet.sheet_kind,
                sheet.sheet_order,
                sheet.row_count,
                json.dumps(sheet.column_names, ensure_ascii=False),
                current_timestamp(),
            ),
        )
        sheet_id = int(cursor.lastrowid)
        if sheet.sheet_kind == "summary":
            summary_count += sheet.row_count
        elif sheet.sheet_kind in {"detail", "other"}:
            detail_count += sheet.row_count

        for row_number, record in enumerate(sheet.rows, start=1):
            raw_count += 1
            context = extract_context(record)
            entity_key = infer_entity_key(record)
            payload_json = json.dumps(record, ensure_ascii=False, sort_keys=True)
            payload_hash = hashlib.sha256(payload_json.encode("utf-8")).hexdigest()
            raw_rows.append(
                {
                    "bao_cao_ngay_id": report_day_id,
                    "sheet_bao_cao_id": sheet_id,
                    "ten_sheet": sheet.sheet_name,
                    "so_dong": row_number,
                    "khoa_ban_ghi": entity_key,
                    "ttvt": context["ttvt"],
                    "doi_vien_thong": context["doi_vien_thong"],
                    "don_vi": context["don_vi"],
                    "ma_nv": context["ma_nv"],
                    "ten_nv": context["ten_nv"],
                    "du_lieu_json": payload_json,
                    "ma_hash_dong": payload_hash,
                    "thoi_gian_tao": current_timestamp(),
                }
            )
            if context["ttvt"] or context["doi_vien_thong"] or context["don_vi"]:
                khoa_chuan_hoa = normalize_key(
                    "|".join(
                        [
                            context["ttvt"] or "",
                            context["doi_vien_thong"] or "",
                            context["don_vi"] or "",
                        ]
                    )
                )
                if khoa_chuan_hoa and khoa_chuan_hoa not in don_vi_rows:
                    don_vi_rows[khoa_chuan_hoa] = {
                        "ttvt": context["ttvt"],
                        "doi_vien_thong": context["doi_vien_thong"],
                        "don_vi": context["don_vi"],
                        "ma_don_vi": None,
                        "khoa_chuan_hoa": khoa_chuan_hoa,
                        "du_lieu_bo_sung_json": json.dumps({"nguon_du_lieu": report_meta.report_code}, ensure_ascii=False),
                        "thoi_gian_tao": current_timestamp(),
                    }
            if context["ten_nv"] or context["ma_nv"]:
                ten_nv = context["ten_nv"] or ""
                ma_nv = context["ma_nv"] or ""
                doi_vien_thong = context["doi_vien_thong"] or ""
                nhan_vien_key = normalize_key(f"{ma_nv}|{ten_nv}|{doi_vien_thong}")
                if ten_nv and nhan_vien_key not in nhan_vien_rows:
                    nhan_vien_rows[nhan_vien_key] = {
                        "ma_nv": ma_nv,
                        "ten_nv": ten_nv,
                        "ten_nv_chuan_hoa": normalize_key(ten_nv),
                        "ttvt": context["ttvt"],
                        "doi_vien_thong": doi_vien_thong,
                        "don_vi": context["don_vi"],
                        "nguon_du_lieu": report_meta.report_code,
                        "du_lieu_bo_sung_json": None,
                        "thoi_gian_tao": current_timestamp(),
                    }

    insert_many(conn, "dong_bao_cao_goc", raw_rows)
    insert_or_ignore_many(conn, "danh_muc_don_vi", list(don_vi_rows.values()))
    insert_or_ignore_many(conn, "danh_muc_nhan_vien", list(nhan_vien_rows.values()))
    return raw_count, summary_count, detail_count


C11_TONG_HOP_FIELDS = [
    SourceField("don_vi", ("Đơn vị",)),
    SourceField("sm1", ("SM1",), "int"),
    SourceField("sm2", ("SM2",), "int"),
    SourceField(
        "ty_le_sua_chua_chat_luong_chu_dong",
        ("Tỷ lệ sửa chữa phiếu chất lượng chủ động dịch vụ FiberVNN, MyTV đạt yêu cầu",),
        "float",
    ),
    SourceField("sm3", ("SM3",), "int"),
    SourceField("sm4", ("SM4",), "int"),
    SourceField(
        "ty_le_bao_hong_brcd_dung_quy_dinh",
        ("Tỷ lệ phiếu sửa chữa báo hỏng dịch vụ BRCD đúng quy định không tính hẹn",),
        "float",
    ),
    SourceField("sm5", ("SM5",), "int"),
    SourceField("sm6", ("SM6",), "int"),
    SourceField("ty_le_sua_chua_trong_ngay_tai_ccco", ("Tỷ lệ phiếu sửa chữa trong ngày tại CCCO",), "float"),
    SourceField("chi_tieu_bsc", ("Chỉ tiêu BSC",), "float"),
]

C11_CHI_TIET_FIELDS = [
    SourceField("doi_vien_thong", ("TEN_DOI", "DOIVT")),
    SourceField("nvkt", ("NVKT",)),
    SourceField("tong_phieu", ("Tổng phiếu",), "int"),
    SourceField("so_phieu_dat", ("Số phiếu đạt",), "int"),
    SourceField(
        "ty_le_dat",
        ("Tỷ lệ phiếu sửa chữa báo hỏng dịch vụ BRCD đúng quy định không tính hẹn",),
        "float",
    ),
]

C12_TONG_HOP_FIELDS = [
    SourceField("don_vi", ("Đơn vị",)),
    SourceField("sm1", ("SM1",), "int"),
    SourceField("sm2", ("SM2",), "int"),
    SourceField("ty_le_hong_lap_lai", ("Tỷ lệ thuê bao báo hỏng dịch vụ BRCĐ lặp lại",), "float"),
    SourceField("sm3", ("SM3",), "int"),
    SourceField("sm4", ("SM4",), "int"),
    SourceField("ty_le_su_co_brcd", ("Tỷ lệ sự cố dịch vụ BRCĐ",), "float"),
    SourceField("chi_tieu_bsc", ("Chỉ tiêu BSC",), "float"),
]

C12_HLL_FIELDS = [
    SourceField("doi_vien_thong", ("TEN_DOI",)),
    SourceField("nvkt", ("NVKT",)),
    SourceField("so_phieu_hll", ("Số phiếu HLL",), "int"),
    SourceField("so_phieu_bao_hong", ("Số phiếu báo hỏng",), "int"),
    SourceField("ty_le_hll", ("Tỉ lệ HLL tháng (2.5%)",), "float"),
]

C13_TONG_HOP_FIELDS = [
    SourceField("don_vi", ("Đơn vị",)),
    SourceField("sm1", ("SM1",), "int"),
    SourceField("sm2", ("SM2",), "int"),
    SourceField("ty_le_sua_chua_dung_han", ("Tỷ lệ sửa chữa dịch vụ kênh TSL hoàn thành đúng thời gian quy định",), "float"),
    SourceField("sm3", ("SM3",), "int"),
    SourceField("sm4", ("SM4",), "int"),
    SourceField("ty_le_hong_lap_lai_kenh_tsl", ("Tỷ lệ thuê bao báo hỏng dịch vụ kênh TSL lặp lại",), "float"),
    SourceField("sm5", ("SM5",), "int"),
    SourceField("sm6", ("SM6",), "int"),
    SourceField("ty_le_su_co_kenh_tsl", ("Tỷ lệ sự cố dịch vụ kênh TSL",), "float"),
    SourceField("chi_tieu_bsc", ("Chỉ tiêu BSC",), "float"),
]

C14_TONG_HOP_FIELDS = [
    SourceField("don_vi", ("Đơn vị",)),
    SourceField("tong_phieu", ("Tổng phiếu",), "int"),
    SourceField("so_luong_da_khao_sat", ("SL đã KS",), "int"),
    SourceField("so_luong_khao_sat_thanh_cong", ("SL KS thành công",), "int"),
    SourceField("so_luong_khach_hang_hai_long", ("SL KH hài lòng",), "int"),
    SourceField("khong_hai_long_ky_thuat_phuc_vu", ("Không HL KT phục vụ",), "int"),
    SourceField("ty_le_hai_long_ky_thuat_phuc_vu", ("Tỷ lệ HL KT phục vụ",), "float"),
    SourceField("khong_hai_long_ky_thuat_dich_vu", ("Không HL KT dịch vụ",), "int"),
    SourceField("ty_le_hai_long_ky_thuat_dich_vu", ("Tỷ lệ HL KT dịch vụ",), "float"),
    SourceField("tong_phieu_hai_long_ky_thuat", ("Tổng phiếu hài lòng KT",), "int"),
    SourceField("ty_le_khach_hang_hai_long", ("Tỷ lệ KH hài lòng",), "float"),
    SourceField("diem_bsc", ("Điểm BSC",), "float"),
]

C14_NVKT_FIELDS = [
    SourceField("doi_vien_thong", ("DOIVT",)),
    SourceField("nvkt", ("NVKT",)),
    SourceField("tong_phieu_khao_sat_thanh_cong", ("Tổng phiếu KS thành công",), "int"),
    SourceField("tong_phieu_khong_hai_long", ("Tổng phiếu KHL",), "int"),
    SourceField("ty_le_hai_long_nvkt", ("Tỉ lệ HL NVKT (%)",), "float"),
]

GHTT_FIELDS = [
    SourceField("don_vi", ("Đơn vị",)),
    SourceField("ttvt", ("TTVT",)),
    SourceField("hoan_thanh_t", ("Hoàn thành T",), "int"),
    SourceField("giao_nvkt_t", ("Giao NVKT T",), "int"),
    SourceField("ty_le_t", ("Tỷ lệ T",), "float"),
    SourceField("hoan_thanh_t_cong_1", ("Hoàn thành T+1",), "int"),
    SourceField("giao_nvkt_t_cong_1", ("Giao NVKT T+1",), "int"),
    SourceField("ty_le_t_cong_1", ("Tỷ lệ T+1",), "float"),
    SourceField("so_luong_ghtt_lon_hon_6_thang", ("SL GHTT >=6T",), "int"),
    SourceField("hoan_thanh_lon_hon_6_thang_t_cong_1", ("Hoàn thành >=6T T+1",), "int"),
    SourceField("ty_le_lon_hon_6_thang_t_cong_1", ("Tỷ lệ >=6T T+1",), "float"),
    SourceField("ty_le_tong", ("Tỷ lệ Tổng",), "float"),
]

GHTT_NVKT_FIELDS = [
    SourceField("nvkt", ("NVKT",)),
    *GHTT_FIELDS,
]

KPI_C11_FIELDS = [
    SourceField("don_vi", ("đơn vị", "Đơn vị")),
    SourceField("nvkt", ("NVKT",)),
    SourceField("sm1", ("SM1",), "int"),
    SourceField("sm2", ("SM2",), "int"),
    SourceField(
        "ty_le_sua_chua_chat_luong_chu_dong",
        ("Tỷ lệ sửa chữa phiếu chất lượng chủ động dịch vụ FiberVNN, MyTV đạt yêu cầu",),
        "float",
    ),
    SourceField("sm3", ("SM3",), "int"),
    SourceField("sm4", ("SM4",), "int"),
    SourceField(
        "ty_le_bao_hong_brcd_dung_quy_dinh",
        ("Tỷ lệ phiếu sửa chữa báo hỏng dịch vụ BRCĐ đúng quy định không tính hẹn",),
        "float",
    ),
    SourceField("chi_tieu_bsc", ("Chỉ tiêu BSC",), "float"),
]

KPI_C12_FIELDS = [
    SourceField("don_vi", ("đơn vị", "Đơn vị")),
    SourceField("nvkt", ("NVKT",)),
    SourceField("sm1", ("SM1",), "int"),
    SourceField("sm2", ("SM2",), "int"),
    SourceField("ty_le_hong_lap_lai", ("Tỷ lệ thuê bao báo hỏng dịch vụ BRCĐ lặp lại",), "float"),
    SourceField("sm3", ("SM3",), "int"),
    SourceField("sm4", ("SM4",), "int"),
    SourceField("ty_le_su_co_brcd", ("Tỷ lệ sự cố dịch vụ BRCĐ",), "float"),
    SourceField("chi_tieu_bsc", ("Chỉ tiêu BSC",), "float"),
]

KPI_C13_FIELDS = [
    SourceField("don_vi", ("đơn vị", "Đơn vị")),
    SourceField("nvkt", ("NVKT",)),
    SourceField("sm1", ("SM1",), "int"),
    SourceField("sm2", ("SM2",), "int"),
    SourceField("ty_le_sua_chua_dung_han", ("Tỷ lệ sửa chữa dịch vụ kênh TSL hoàn thành đúng thời gian quy định",), "float"),
    SourceField("sm3", ("SM3",), "int"),
    SourceField("sm4", ("SM4",), "int"),
    SourceField("ty_le_hong_lap_lai_kenh_tsl", ("Tỷ lệ thuê bao báo hỏng dịch vụ kênh TSL lặp lại",), "float"),
    SourceField("sm5", ("SM5",), "int"),
    SourceField("sm6", ("SM6",), "int"),
    SourceField("ty_le_su_co_kenh_tsl", ("Tỷ lệ sự cố dịch vụ kênh TSL",), "float"),
    SourceField("chi_tieu_bsc", ("Chỉ tiêu BSC",), "float"),
]

KQ_TIEP_THI_NV_FIELDS = [
    SourceField("don_vi", ("Đơn vị",)),
    SourceField("ma_nv", ("Mã NV",)),
    SourceField("ten_nv", ("Tên NV",)),
    SourceField("dich_vu_brcd", ("Dịch vụ BRCĐ",), "int"),
    SourceField("dich_vu_mytv", ("Dịch vụ MyTV",), "int"),
    SourceField("tong", ("Tổng",), "int"),
]

KQ_TIEP_THI_DV_FIELDS = [
    SourceField("don_vi", ("Đơn vị",)),
    SourceField("dich_vu_brcd", ("Dịch vụ BRCĐ",), "int"),
    SourceField("dich_vu_mytv", ("Dịch vụ MyTV",), "int"),
    SourceField("tong", ("Tổng",), "int"),
]

HOAN_CONG_FIBER_FIELDS = [
    SourceField("ma_thue_bao", ("MA_TB",)),
    SourceField("so_may", ("SO_MAY",)),
    SourceField("hdtb_id", ("HDTB_ID",)),
    SourceField("ma_khach_hang", ("MA_KH",)),
    SourceField("ma_giao_dich", ("MA_GD",)),
    SourceField("loai_hinh_thue_bao", ("LOAIHINH_TB",)),
    SourceField("ten_dich_vu", ("TEN_DVVT_HNI",)),
    SourceField("ten_goi", ("TEN_GOI",)),
    SourceField("ten_thue_bao", ("TEN_TB",)),
    SourceField("dia_chi_thue_bao", ("DIACHI_TB",)),
    SourceField("ngay_yeu_cau", ("NGAYLAP_HD",), "datetime"),
    SourceField("ngay_nghiem_thu", ("NGAY_HC",), "datetime"),
    SourceField("ttvt", ("TTVT",)),
    SourceField("doi_vien_thong", ("DOIVT",)),
    SourceField("nvkt", ("NVKT",)),
    SourceField("trang_thai_hop_dong", ("TRANGTHAI_HD",)),
]

NGUNG_KHOI_PHUC_FIBER_FIELDS = [
    SourceField("ma_thue_bao", ("MA_TB",)),
    SourceField("ten_thue_bao", ("TEN_TB",)),
    SourceField("ngay_lap_hop_dong", ("NGAYLAP_HD",), "datetime"),
    SourceField("ngay_thuc_hien", ("NGAY_THUCHIEN",), "datetime"),
    SourceField("ten_kieu_lenh", ("TEN_KIEULD",)),
    SourceField("ly_do_huy", ("LYDOHUY",)),
    SourceField("trang_thai_thue_bao", ("TRANGTHAI_TB",)),
    SourceField("ten_loai_hop_dong", ("TEN_LOAIHD",)),
    SourceField("trang_thai_hop_dong", ("TRANGTHAI_HD",)),
    SourceField("ma_giao_dich", ("MA_GD",)),
    SourceField("ttvt", ("TTVT",)),
    SourceField("doi_vien_thong", ("DOIVT",)),
    SourceField("nvkt", ("NVKT",)),
]

HOAN_CONG_MYTV_FIELDS = [
    SourceField("ma_thue_bao", ("MA_TB",)),
    SourceField("hdtb_id", ("HDTB_ID",)),
    SourceField("ngay_ins", ("NGAY_INS",), "datetime"),
    SourceField("ngay_yeu_cau", ("NGAY_YC",), "datetime"),
    SourceField("nhom_dia_ban", ("NHOM_DIABAN",)),
    SourceField("doi_vien_thong", ("DOIVT",)),
    SourceField("ten_ttvt", ("TEN_TTVT",)),
    SourceField("nhan_vien_ky_thuat", ("NHANVIEN_KT",)),
    SourceField("ma_giao_dich", ("MA_GD",)),
    SourceField("trang_thai_hop_dong", ("TRANGTHAI_HD",)),
]

NGUNG_PSC_MYTV_FIELDS = [
    SourceField("ma_thue_bao", ("MA_TB",)),
    SourceField("ngay_tam_dung", ("NGAY_TD",), "datetime"),
    SourceField("ngay_huy", ("NGAY_HUY",), "datetime"),
    SourceField("ma_khu_vuc", ("MA_KV",)),
    SourceField("ten_khu_vuc", ("TEN_KV",)),
    SourceField("ten_doi", ("TEN_DOI",)),
    SourceField("ten_ttvt", ("TEN_TTVT",)),
    SourceField("trang_thai_thue_bao", ("TRANGTHAI_TB",)),
    SourceField("loai_hinh_thue_bao", ("LOAIHINH_TB",)),
]

THUC_TANG_FIELDS = [
    SourceField("ttvt", ("TTVT",)),
    SourceField("doi_vien_thong", ("Đội VT", "DOIVT")),
    SourceField("nvkt", ("NVKT",)),
    SourceField("hoan_cong", ("Hoàn công",), "int"),
    SourceField("ngung_phat_sinh_cuoc", ("Ngưng phát sinh cước",), "int"),
    SourceField("thuc_tang", ("Thực tăng",), "int"),
    SourceField("ty_le_ngung_psc", ("Tỷ lệ ngưng/psc",), "float"),
]

XAC_MINH_CHI_TIET_FIELDS = [
    SourceField("ma_thue_bao", ("MA_TB",)),
    SourceField("ten_loai_hinh_thue_bao", ("TEN_LOAIHINHTB",)),
    SourceField("ten_thue_bao", ("TEN_TB",)),
    SourceField("ten_kieu_lenh", ("TEN_KIEULD", "TEN_KIEULD.1")),
    SourceField("ngay_lap_hop_dong", ("NGAYLAP_HD",), "datetime"),
    SourceField("ngay_hoan_thanh", ("NGAY_HOAN_THANH",), "datetime"),
    SourceField("loai_phieu", ("LOAI_PHIEU",)),
    SourceField("ten_khu_vuc", ("TEN_KV",)),
    SourceField("doi_vien_thong", ("DOIVT",)),
    SourceField("ttvt", ("TTVT",)),
    SourceField("nvkt", ("NVKT",)),
]

XAC_MINH_NVKT_FIELDS = [
    SourceField("ttvt", ("TTVT",)),
    SourceField("doi_vien_thong", ("DOIVT",)),
    SourceField("nvkt", ("NVKT",)),
    SourceField("so_phieu_xac_minh", ("SỐ PHIẾU XM",), "int"),
]

XAC_MINH_LOAI_PHIEU_FIELDS = [
    SourceField("loai_phieu", ("LOAI_PHIEU",)),
    SourceField("ten_kieu_lenh", ("TEN_KIEULD.1",)),
    SourceField("so_phieu_xac_minh", ("SỐ PHIẾU XM",), "int"),
]

CAU_HINH_CHI_TIET_FIELDS = [
    SourceField("serial_number", ("Serial Number",)),
    SourceField("ma_thue_bao", ("Mã thuê bao",)),
    SourceField("loai_hop_dong", ("Loại hợp đồng",)),
    SourceField("loai_cau_hinh", ("Loại cấu hình",)),
    SourceField("trang_thai", ("Trang thái",)),
    SourceField("trang_thai_chuan_hoa", ("Trang thái chuẩn hóa",)),
    SourceField("ma_loi", ("Mã lỗi",)),
    SourceField("thoi_gian_cap_nhat", ("Thời gian cập nhật",), "datetime"),
    SourceField("ttvt", ("Trung tâm Viễn thông",)),
    SourceField("doi_vien_thong", ("Đội Viễn thông",)),
    SourceField("ma_nhan_vien", ("Mã nhân viên",)),
    SourceField("nvkt", ("NVKT",)),
]

CAU_HINH_TONG_HOP_FIELDS = [
    SourceField("ttvt", ("TTVT",)),
    SourceField("don_vi", ("Đơn vị",)),
    SourceField("loai_dong", ("Loại dòng",)),
    SourceField("tong_hop_dong", ("Tổng hợp đồng",), "int"),
    SourceField("khong_thuc_hien_cau_hinh_tu_dong", ("Không thực hiện cấu hình tự động",), "int"),
    SourceField("da_day_cau_hinh_tu_dong", ("Đã đẩy cấu hình tự động",), "int"),
    SourceField("khong_day_do_loi_he_thong", ("Không đẩy do lỗi hệ thống",), "int"),
    SourceField("khong_day_do_tbi_da_co_cau_hinh", ("Không đẩy do TBI đã có cấu hình",), "int"),
    SourceField("cau_hinh_thanh_cong", ("Cấu hình thành công",), "int"),
    SourceField("ty_le_day_tu_dong", ("Tỷ lệ đẩy tự động (%)",), "float"),
    SourceField("ty_le_tbi_da_co_cau_hinh", ("Tỷ lệ TBI đã có cấu hình (%)",), "float"),
    SourceField("ty_le_cau_hinh_thanh_cong", ("Tỷ lệ cấu hình thành công (%)",), "float"),
]

CAU_HINH_LOI_FIELDS = [
    SourceField("ma_loi", ("Mã lỗi",)),
    SourceField("so_luong", ("Số lượng",), "int"),
]

VAT_TU_THU_HOI_FIELDS = [
    SourceField("nvkt_dia_ban_giao", ("NVKT_DIABAN_GIAO",)),
    SourceField("trang_thai_thu_hoi", ("TRANGTHAI_THUHOI",)),
    SourceField("loai_vat_tu", ("LOAI_VT",)),
    SourceField("loai_phieu", ("LOAI_PHIEU",)),
    SourceField("ma_men", ("MA_MEN",)),
    SourceField("ma_thue_bao", ("MA_TB",)),
    SourceField("ten_thue_bao", ("TEN_TB",)),
    SourceField("dia_chi_khach_hang", ("DIACHI_KH",)),
    SourceField("nhan_vien_khoa", ("NHANVIEN_KHOA",)),
    SourceField("nhan_vien_thu", ("NHANVIEN_THU",)),
    SourceField("nhan_vien_nhap_kho", ("NHANVIEN_NHAPKHO",)),
    SourceField("ngay_khoa", ("NGAY_KHOA",), "date"),
    SourceField("ngay_hoan_cong", ("NGAY_HOAN_CONG",), "date"),
    SourceField("ngay_hoan_ung", ("NGAY_HOAN_UNG",), "date"),
]

CHI_TIET_VAT_TU_FIELDS = [
    SourceField("diem_chia", ("DIEMCHIA",)),
    SourceField("nvkt_dia_ban_giao", ("NVKT_DIABAN_GIAO",)),
    SourceField("ma_thue_bao", ("MA_TB",)),
    SourceField("ten_thue_bao", ("TEN_TB",)),
    SourceField("ten_thiet_bi", ("TEN_TBI",)),
    SourceField("ngay_giao", ("NGAY_GIAO",), "date"),
    SourceField("ten_loai_hop_dong", ("TEN_LOAIHD",)),
    SourceField("ten_kieu_lenh", ("TEN_KIEULD",)),
    SourceField("so_dien_thoai", ("SO_DT",)),
    SourceField("ngay_su_dung_thiet_bi", ("NGAY_SD_TB",), "date"),
]

QUYET_TOAN_VAT_TU_FIELDS = [
    SourceField("ma_tkyt", ("MA_TKYT",)),
    SourceField("ma_spdv", ("MA_SPDV",)),
    SourceField("loai", ("LOAI",)),
    SourceField("ten_vat_tu", ("TEN_VT",)),
    SourceField("don_vi_tinh", ("DVI_TINH",)),
    SourceField("so_luong", ("SOLUONG",), "float"),
    SourceField("don_gia", ("DONGIA",), "float"),
    SourceField("thanh_tien", ("THANHTIEN",), "float"),
    SourceField("ma_vat_tu", ("MA_VT",)),
]


def insert_from_sheet(
    conn: sqlite3.Connection,
    table: str,
    report_day_id: int,
    sheet: Optional[SheetData],
    fields: Sequence[SourceField],
    *,
    required_keys: Sequence[str],
    base_values: Optional[Dict[str, Any]] = None,
    extra_base: Optional[Dict[str, Any]] = None,
    include_extra_json: bool = True,
    empty_string_fields: Sequence[str] = (),
) -> int:
    if sheet is None:
        return 0
    rows: List[Dict[str, Any]] = []
    for record in sheet.rows:
        mapped, used_keys = map_record(record, fields)
        if base_values:
            mapped.update(base_values)
        for field_name in empty_string_fields:
            if mapped.get(field_name) is None:
                mapped[field_name] = ""
        if any(mapped.get(key) in (None, "") for key in required_keys):
            continue
        mapped["bao_cao_ngay_id"] = report_day_id
        if include_extra_json:
            mapped["du_lieu_bo_sung_json"] = build_extra_json(
                record,
                used_keys,
                extra={"sheet_name": sheet.sheet_name, **(extra_base or {})},
            )
        rows.append(mapped)
    return insert_many(conn, table, rows)


def import_c11(conn: sqlite3.Connection, report_day_id: int, sheets: Sequence[SheetData]) -> int:
    count = 0
    count += insert_from_sheet(
        conn,
        "c11_tong_hop",
        report_day_id,
        find_sheet(sheets, "TH_C1.1"),
        C11_TONG_HOP_FIELDS,
        required_keys=("don_vi",),
    )
    for sheet in find_sheets_by_prefix(sheets, ["chi_tiet", "chi_tieu_ko_hen"]):
        if not sheet.rows or "Tổng phiếu" not in sheet.column_names:
            continue
        moc_gio = "tong"
        match = re.search(r"(\d{2})h", normalize_key(sheet.sheet_name))
        if match:
            moc_gio = f"{match.group(1)}h"
        count += insert_from_sheet(
            conn,
            "c11_chi_tiet_nvkt",
            report_day_id,
            sheet,
            C11_CHI_TIET_FIELDS,
            required_keys=("doi_vien_thong", "nvkt"),
            base_values={"moc_gio": moc_gio},
        )
    return count


def import_c12(conn: sqlite3.Connection, report_day_id: int, sheets: Sequence[SheetData]) -> int:
    count = 0
    count += insert_from_sheet(
        conn,
        "c12_tong_hop",
        report_day_id,
        find_sheet(sheets, "TH_C1.2"),
        C12_TONG_HOP_FIELDS,
        required_keys=("don_vi",),
    )
    count += insert_from_sheet(
        conn,
        "c12_hong_lap_lai_nvkt",
        report_day_id,
        find_sheet(sheets, "TH_SM1C12_HLL_Thang"),
        C12_HLL_FIELDS,
        required_keys=("doi_vien_thong", "nvkt"),
    )
    return count


def import_c13(conn: sqlite3.Connection, report_day_id: int, sheets: Sequence[SheetData]) -> int:
    return insert_from_sheet(
        conn,
        "c13_tong_hop",
        report_day_id,
        find_sheet(sheets, "TH_C1.3"),
        C13_TONG_HOP_FIELDS,
        required_keys=("don_vi",),
    )


def import_c14(conn: sqlite3.Connection, report_day_id: int, sheets: Sequence[SheetData]) -> int:
    count = 0
    count += insert_from_sheet(
        conn,
        "c14_tong_hop",
        report_day_id,
        find_sheet(sheets, "TH_C1.4"),
        C14_TONG_HOP_FIELDS,
        required_keys=("don_vi",),
    )
    count += insert_from_sheet(
        conn,
        "c14_hai_long_nvkt",
        report_day_id,
        find_sheet(sheets, "TH_HL_NVKT"),
        C14_NVKT_FIELDS,
        required_keys=("doi_vien_thong", "nvkt"),
    )
    return count


def import_ghtt_don_vi(conn: sqlite3.Connection, report_day_id: int, sheets: Sequence[SheetData], sheet_name: str) -> int:
    return insert_from_sheet(
        conn,
        "ghtt_don_vi",
        report_day_id,
        find_sheet(sheets, sheet_name),
        GHTT_FIELDS,
        required_keys=("don_vi",),
        empty_string_fields=("ttvt",),
    )


def import_ghtt_nvkt(conn: sqlite3.Connection, report_day_id: int, sheets: Sequence[SheetData]) -> int:
    return insert_from_sheet(
        conn,
        "ghtt_nvkt",
        report_day_id,
        find_sheet(sheets, "kq_nvktdb"),
        GHTT_NVKT_FIELDS,
        required_keys=("nvkt",),
        empty_string_fields=("don_vi", "ttvt"),
    )


def import_kpi_nvkt(conn: sqlite3.Connection, report_day_id: int, sheets: Sequence[SheetData], table: str, sheet_name: str, fields: Sequence[SourceField]) -> int:
    return insert_from_sheet(
        conn,
        table,
        report_day_id,
        find_sheet(sheets, sheet_name),
        fields,
        required_keys=("don_vi", "nvkt"),
    )


def import_kq_tiep_thi(conn: sqlite3.Connection, report_day_id: int, sheets: Sequence[SheetData]) -> int:
    count = 0
    count += insert_from_sheet(
        conn,
        "ket_qua_tiep_thi_nv",
        report_day_id,
        find_sheet(sheets, "kq_tiep_thi"),
        KQ_TIEP_THI_NV_FIELDS,
        required_keys=("don_vi", "ten_nv"),
        empty_string_fields=("ma_nv",),
    )
    count += insert_from_sheet(
        conn,
        "ket_qua_tiep_thi_don_vi",
        report_day_id,
        find_sheet(sheets, "kq_th"),
        KQ_TIEP_THI_DV_FIELDS,
        required_keys=("don_vi",),
    )
    return count


def import_hoan_cong_fiber(conn: sqlite3.Connection, report_day_id: int, sheets: Sequence[SheetData]) -> int:
    return insert_from_sheet(
        conn,
        "hoan_cong_fiber",
        report_day_id,
        find_sheet(sheets, "Data", "Sheet1"),
        HOAN_CONG_FIBER_FIELDS,
        required_keys=("ma_thue_bao",),
    )


def import_ngung_khoi_phuc_fiber(
    conn: sqlite3.Connection,
    report_day_id: int,
    sheets: Sequence[SheetData],
    table: str,
) -> int:
    return insert_from_sheet(
        conn,
        table,
        report_day_id,
        find_sheet(sheets, "Data", "Sheet1"),
        NGUNG_KHOI_PHUC_FIBER_FIELDS,
        required_keys=("ma_thue_bao",),
    )


def import_hoan_cong_mytv(conn: sqlite3.Connection, report_day_id: int, sheets: Sequence[SheetData]) -> int:
    return insert_from_sheet(
        conn,
        "hoan_cong_mytv",
        report_day_id,
        find_sheet(sheets, "Data"),
        HOAN_CONG_MYTV_FIELDS,
        required_keys=("ma_thue_bao",),
    )


def import_ngung_psc_mytv(conn: sqlite3.Connection, report_day_id: int, sheets: Sequence[SheetData]) -> int:
    return insert_from_sheet(
        conn,
        "ngung_psc_mytv",
        report_day_id,
        find_sheet(sheets, "Data"),
        NGUNG_PSC_MYTV_FIELDS,
        required_keys=("ma_thue_bao",),
    )


def import_thuc_tang(conn: sqlite3.Connection, report_day_id: int, sheets: Sequence[SheetData], table: str) -> int:
    count = 0
    for sheet_name in ("thuc_tang_theo_to", "thuc_tang_theo_NVKT"):
        sheet = find_sheet(sheets, sheet_name)
        if sheet is None:
            continue
        cap_tong_hop = "nvkt" if "nvkt" in normalize_key(sheet.sheet_name) else "to"
        rows: List[Dict[str, Any]] = []
        for record in sheet.rows:
            mapped, used_keys = map_record(record, THUC_TANG_FIELDS)
            if not mapped.get("doi_vien_thong"):
                continue
            mapped["bao_cao_ngay_id"] = report_day_id
            mapped["cap_tong_hop"] = cap_tong_hop
            mapped["ttvt"] = mapped.get("ttvt") or ""
            mapped["doi_vien_thong"] = mapped.get("doi_vien_thong") or ""
            mapped["nvkt"] = mapped.get("nvkt") or ""
            mapped["du_lieu_bo_sung_json"] = build_extra_json(record, used_keys, extra={"sheet_name": sheet.sheet_name})
            rows.append(mapped)
        count += insert_many(conn, table, rows)
    return count


def import_xac_minh(conn: sqlite3.Connection, report_day_id: int, sheets: Sequence[SheetData]) -> int:
    count = 0
    count += insert_from_sheet(
        conn,
        "xac_minh_chi_tiet",
        report_day_id,
        find_sheet(sheets, "Data"),
        XAC_MINH_CHI_TIET_FIELDS,
        required_keys=("ma_thue_bao",),
    )
    count += insert_from_sheet(
        conn,
        "xac_minh_tong_hop_nvkt",
        report_day_id,
        find_sheet(sheets, "tong_hop_theo_nvkt"),
        XAC_MINH_NVKT_FIELDS,
        required_keys=("doi_vien_thong",),
        empty_string_fields=("ttvt",),
    )
    count += insert_from_sheet(
        conn,
        "xac_minh_tong_hop_loai_phieu",
        report_day_id,
        find_sheet(sheets, "tong_hop_theo_loai_phieu"),
        XAC_MINH_LOAI_PHIEU_FIELDS,
        required_keys=("loai_phieu",),
        empty_string_fields=("ten_kieu_lenh",),
    )
    return count


def import_cau_hinh(conn: sqlite3.Connection, report_day_id: int, sheets: Sequence[SheetData]) -> int:
    count = 0
    count += insert_from_sheet(
        conn,
        "cau_hinh_tu_dong_chi_tiet",
        report_day_id,
        find_sheet(sheets, "chi_tiet"),
        CAU_HINH_CHI_TIET_FIELDS,
        required_keys=("serial_number",),
    )
    count += insert_from_sheet(
        conn,
        "cau_hinh_tu_dong_tong_hop",
        report_day_id,
        find_sheet(sheets, "du_lieu_sach"),
        CAU_HINH_TONG_HOP_FIELDS,
        required_keys=("don_vi",),
        empty_string_fields=("ttvt",),
    )
    count += insert_from_sheet(
        conn,
            "tong_hop_loi_cau_hinh_tu_dong",
            report_day_id,
            find_sheet(sheets, "tong_hop_loi"),
            CAU_HINH_LOI_FIELDS,
            required_keys=("ma_loi",),
            include_extra_json=False,
        )
    return count


def import_vat_tu(conn: sqlite3.Connection, report_day_id: int, sheets: Sequence[SheetData]) -> int:
    count = 0
    count += insert_from_sheet(
        conn,
        "vat_tu_thu_hoi",
        report_day_id,
        find_sheet(sheets, "Chi tiết"),
        VAT_TU_THU_HOI_FIELDS,
        required_keys=("ma_men",),
    )
    count += insert_from_sheet(
        conn,
        "chi_tiet_vat_tu_thu_hoi",
        report_day_id,
        find_sheet(sheets, "Chi tiết vật tư"),
        CHI_TIET_VAT_TU_FIELDS,
        required_keys=("ma_thue_bao",),
    )
    return count


def import_quyet_toan_vat_tu(conn: sqlite3.Connection, report_day_id: int, sheets: Sequence[SheetData]) -> int:
    return insert_from_sheet(
        conn,
        "quyet_toan_vat_tu",
        report_day_id,
        find_sheet(sheets, "Data", "Sheet1"),
        QUYET_TOAN_VAT_TU_FIELDS,
        required_keys=("ma_tkyt",),
    )


def populate_business_tables(
    conn: sqlite3.Connection,
    report_day_id: int,
    report_meta: ReportMeta,
    sheets: Sequence[SheetData],
) -> int:
    rel_path = report_meta.processed_rel_path.lower()
    if rel_path == "chi_tieu_c/c1.1 report_processed.xlsx":
        return import_c11(conn, report_day_id, sheets)
    if rel_path == "chi_tieu_c/c1.2 report_processed.xlsx":
        return import_c12(conn, report_day_id, sheets)
    if rel_path == "chi_tieu_c/c1.3 report_processed.xlsx":
        return import_c13(conn, report_day_id, sheets)
    if rel_path == "chi_tieu_c/c1.4 report_processed.xlsx":
        return import_c14(conn, report_day_id, sheets)
    if rel_path == "chi_tieu_c/c1.1_chitiet_report_processed.xlsx":
        return import_c11(conn, report_day_id, sheets)
    if rel_path == "chi_tieu_c/c1.2_chitiet_sm1_report_processed.xlsx":
        return import_c12(conn, report_day_id, sheets)
    if rel_path == "ghtt/ghtt_hni report_processed.xlsx":
        return import_ghtt_don_vi(conn, report_day_id, sheets, "kq_hni")
    if rel_path == "ghtt/ghtt_sontay report_processed.xlsx":
        return import_ghtt_don_vi(conn, report_day_id, sheets, "kq_sontay")
    if rel_path == "ghtt/ghtt_nvktdb report_processed.xlsx":
        return import_ghtt_nvkt(conn, report_day_id, sheets)
    if rel_path == "kpi_nvkt/c11-nvktdb report_processed.xlsx":
        return import_kpi_nvkt(conn, report_day_id, sheets, "kpi_nvkt_c11", "c11 kpi nvkt", KPI_C11_FIELDS)
    if rel_path == "kpi_nvkt/c12-nvktdb report_processed.xlsx":
        return import_kpi_nvkt(conn, report_day_id, sheets, "kpi_nvkt_c12", "c12 kpi nvkt", KPI_C12_FIELDS)
    if rel_path == "kpi_nvkt/c13-nvktdb report_processed.xlsx":
        return import_kpi_nvkt(conn, report_day_id, sheets, "kpi_nvkt_c13", "c13 kpi nvkt", KPI_C13_FIELDS)
    if rel_path == "kq_tiep_thi/kq_tiep_thi report_processed.xlsx":
        return import_kq_tiep_thi(conn, report_day_id, sheets)
    if "phieu_hoan_cong_dich_vu_chi_tiet_processed.xlsx" in rel_path:
        return import_hoan_cong_fiber(conn, report_day_id, sheets)
    if "tam_dung_khoi_phuc_dich_vu_chi_tiet_processed.xlsx" in rel_path:
        return import_ngung_khoi_phuc_fiber(conn, report_day_id, sheets, "ngung_psc_fiber")
    if "tam_dung_khoi_phuc_dich_vu_chi_tiet_khoi_phuc_processed.xlsx" in rel_path:
        return import_ngung_khoi_phuc_fiber(conn, report_day_id, sheets, "khoi_phuc_fiber")
    if "mytv_hoan_cong_" in rel_path:
        return import_hoan_cong_mytv(conn, report_day_id, sheets)
    if "mytv_ngung_psc_" in rel_path:
        return import_ngung_psc_mytv(conn, report_day_id, sheets)
    if rel_path == "thuc_tang_ngung_psc/fiber_thuc_tang_processed.xlsx":
        return import_thuc_tang(conn, report_day_id, sheets, "thuc_tang_fiber")
    if rel_path == "mytv_dich_vu/mytv_thuc_tang_processed.xlsx":
        return import_thuc_tang(conn, report_day_id, sheets, "thuc_tang_mytv")
    if rel_path == "ty_le_xac_minh/ty_le_xac_minh_dung_thoi_gian_quy_dinh_chi_tiet_processed.xlsx":
        return import_xac_minh(conn, report_day_id, sheets)
    if rel_path.startswith("cau_hinh_tu_dong/"):
        return import_cau_hinh(conn, report_day_id, sheets)
    if rel_path == "vat_tu_thu_hoi/bc_thu_hoi_vat_tu_processed.xlsx":
        return import_vat_tu(conn, report_day_id, sheets)
    if rel_path == "vat_tu_thu_hoi/quyet_toan_vat_tu_processed.xlsx":
        return import_quyet_toan_vat_tu(conn, report_day_id, sheets)
    return 0


def insert_file_records(
    conn: sqlite3.Connection,
    report_day_id: int,
    processed_path: Path,
    processed_root: Path,
    archive_path: Optional[Path],
    source_hash: str,
    size_bytes: int,
) -> None:
    rows = [
        {
            "bao_cao_ngay_id": report_day_id,
            "loai_tep": "processed",
            "duong_dan_tep": str(processed_path.relative_to(API_TRANSITION_DIR)),
            "kich_thuoc_byte": size_bytes,
            "ma_hash_tep": source_hash,
            "thoi_gian_tao": current_timestamp(),
        }
    ]
    if archive_path is not None:
        rows.append(
            {
                "bao_cao_ngay_id": report_day_id,
                "loai_tep": "daily",
                "duong_dan_tep": str(archive_path.relative_to(API_TRANSITION_DIR)),
                "kich_thuoc_byte": size_bytes,
                "ma_hash_tep": source_hash,
                "thoi_gian_tao": current_timestamp(),
            }
        )
    insert_many(conn, "tep_luu_tru_bao_cao", rows)


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
    report_meta = build_report_meta(processed_root, workbook_path)
    workbook_bytes = workbook_path.read_bytes()
    source_hash = sha256_bytes(workbook_bytes)
    archive_path = pre_archived_path
    if not dry_run and not skip_archive and archive_path is None:
        archive_path = archive_processed_file(workbook_path, processed_root, archive_root, snapshot_date)
    sheets = read_workbook_sheets(workbook_path)
    raw_count = sum(len(sheet.rows) for sheet in sheets)
    summary_count = sum(sheet.row_count for sheet in sheets if sheet.sheet_kind == "summary")
    detail_count = sum(sheet.row_count for sheet in sheets if sheet.sheet_kind in {"detail", "other"})

    result = {
        "report_code": report_meta.report_code,
        "report_name": report_meta.report_name,
        "processed_rel_path": report_meta.processed_rel_path,
        "snapshot_date": snapshot_date.isoformat(),
        "file_sha256": source_hash,
        "raw_rows": raw_count,
        "summary_rows": summary_count,
        "detail_rows": detail_count,
        "business_rows": 0,
        "status": "dry_run" if dry_run else "pending",
        "archive_path": str(archive_path.relative_to(API_TRANSITION_DIR)) if archive_path is not None else None,
    }

    if dry_run:
        return result

    existing_row = get_existing_report_day(conn, report_meta.report_code, snapshot_date)
    log_id = start_import_log(conn, report_meta, snapshot_date, source_hash)

    if skip_if_same_hash and existing_row and existing_row["ma_hash_tep"] == source_hash:
        finish_import_log(
            conn,
            log_id,
            status="bo_qua",
            message="Bo qua vi cung ngay va cung hash file.",
            raw_count=raw_count,
            summary_count=summary_count,
            detail_count=detail_count,
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
            existing_row,
        )
        insert_file_records(
            conn,
            report_day_id,
            workbook_path,
            processed_root,
            archive_path,
            source_hash,
            len(workbook_bytes),
        )
        raw_row_count, summary_row_count, detail_row_count = insert_sheet_and_raw_rows(conn, report_day_id, report_meta, sheets)
        business_rows = populate_business_tables(conn, report_day_id, report_meta, sheets)
        conn.execute(
            """
            UPDATE bao_cao_ngay
            SET so_dong_goc = ?,
                so_dong_tong_hop = ?,
                so_dong_chi_tiet = ?,
                trang_thai_nap = 'thanh_cong',
                ghi_chu = ?,
                thoi_gian_cap_nhat = ?
            WHERE id = ?
            """,
            (
                raw_row_count,
                summary_row_count,
                detail_row_count,
                f"Da import {business_rows} dong nghiep vu.",
                current_timestamp(),
                report_day_id,
            ),
        )
        conn.commit()
        finish_import_log(
            conn,
            log_id,
            status="thanh_cong",
            message=f"Thanh cong. business_rows={business_rows}",
            raw_count=raw_row_count,
            summary_count=summary_row_count,
            detail_count=detail_row_count,
        )
        result["status"] = "imported"
        result["business_rows"] = business_rows
        return result
    except Exception as exc:
        conn.rollback()
        finish_import_log(
            conn,
            log_id,
            status="that_bai",
            message=str(exc),
            raw_count=raw_count,
            summary_count=summary_count,
            detail_count=detail_count,
        )
        raise


def iter_workbooks(processed_root: Path, path_filters: Sequence[str]) -> Iterable[Path]:
    lowered_filters = [value.lower() for value in path_filters if value]
    for workbook_path in sorted(processed_root.rglob("*.xlsx")):
        rel_path = str(workbook_path.relative_to(processed_root)).replace("\\", "/").lower()
        if lowered_filters and not all(value in rel_path for value in lowered_filters):
            continue
        yield workbook_path


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
                f"raw={item['raw_rows']} summary={item['summary_rows']} "
                f"detail={item['detail_rows']} business={item['business_rows']}"
            )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
