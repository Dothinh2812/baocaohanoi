#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Dong bo du lieu workbook processed len Supabase."""

from __future__ import annotations

import argparse
import hashlib
import json
import mimetypes
import numbers
import os
import re
import sys
import uuid
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import pandas as pd
import requests
from dotenv import load_dotenv


PROJECT_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_PROCESSED_ROOT = Path(__file__).resolve().parent / "Processed"
DATE_IN_NAME_RE = re.compile(r"(\d{8})")

GROUP_CATEGORY_MAP = {
    "chi_tieu_c": "quality",
    "ghtt": "quality",
    "kpi_nvkt": "quality",
    "kq_tiep_thi": "sales",
    "phieu_hoan_cong_dich_vu": "service",
    "tam_dung_khoi_phuc_dich_vu": "service",
    "thuc_tang_ngung_psc": "service",
    "mytv_dich_vu": "service",
    "ty_le_xac_minh": "verification",
    "cau_hinh_tu_dong": "automation",
    "vat_tu_thu_hoi": "equipment",
}

SOURCE_SHEET_NAMES = {"sheet", "sheet1"}
DETAIL_SHEET_NAMES = {
    "data",
    "data_combined",
    "data_tam_dung",
    "data_khoi_phuc",
    "chi tiết vật tư",
    "chi_tiet_chua_khoi_phuc",
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
    "ngay_yc",
    "ngay_td",
    "ngay_huy",
    "ngaylap_hd",
    "thời gian cập nhật",
    "thoi gian cap nhat",
    "ngay_ins",
)


@dataclass(frozen=True)
class ReportMeta:
    report_code: str
    report_name: str
    report_group: str
    report_category: str
    service_type: Optional[str]
    action_type: Optional[str]
    processed_rel_path: str


@dataclass
class SheetPlan:
    sheet_name: str
    sheet_kind: str
    row_count: int
    column_names: List[str]
    numeric_columns: List[str]
    text_columns: List[str]
    raw_rows: List[Dict[str, Any]]
    metric_rows: List[Dict[str, Any]]
    detail_rows: List[Dict[str, Any]]
    fallback_detail_rows: List[Dict[str, Any]]


class SupabaseClient:
    def __init__(self, base_url: str, service_role_key: str, timeout: int = 60):
        self.base_url = base_url.rstrip("/")
        self.timeout = timeout
        self.session = requests.Session()
        self.session.headers.update(
            {
                "apikey": service_role_key,
                "Authorization": f"Bearer {service_role_key}",
            }
        )

    def _check(self, response: requests.Response, context: str) -> None:
        if response.ok:
            return
        detail = response.text.strip()
        raise RuntimeError(f"{context} that bai: HTTP {response.status_code} - {detail}")

    def delete_rows(self, table: str, filters: Dict[str, str]) -> None:
        response = self.session.delete(
            f"{self.base_url}/rest/v1/{table}",
            params=filters,
            timeout=self.timeout,
        )
        if response.status_code not in (200, 204):
            self._check(response, f"DELETE {table}")

    def insert_rows(
        self,
        table: str,
        rows: Sequence[Dict[str, Any]],
        *,
        upsert: bool = False,
        on_conflict: Optional[str] = None,
        returning: str = "minimal",
    ) -> None:
        if not rows:
            return
        headers = {"Prefer": f"return={returning}"}
        if upsert:
            headers["Prefer"] = f"resolution=merge-duplicates,return={returning}"
        params: Dict[str, str] = {}
        if on_conflict:
            params["on_conflict"] = on_conflict
        response = self.session.post(
            f"{self.base_url}/rest/v1/{table}",
            params=params,
            headers=headers,
            json=list(rows),
            timeout=self.timeout,
        )
        self._check(response, f"INSERT {table}")

    def patch_rows(self, table: str, filters: Dict[str, str], payload: Dict[str, Any]) -> None:
        response = self.session.patch(
            f"{self.base_url}/rest/v1/{table}",
            params=filters,
            headers={"Prefer": "return=minimal"},
            json=payload,
            timeout=self.timeout,
        )
        self._check(response, f"PATCH {table}")

    def upload_bytes(
        self,
        bucket: str,
        object_path: str,
        content: bytes,
        content_type: str,
    ) -> None:
        response = self.session.post(
            f"{self.base_url}/storage/v1/object/{bucket}/{object_path}",
            params={"upsert": "true"},
            headers={"Content-Type": content_type, "x-upsert": "true"},
            data=content,
            timeout=max(self.timeout, 120),
        )
        self._check(response, f"UPLOAD storage {bucket}/{object_path}")


def load_environment() -> None:
    load_dotenv(PROJECT_ROOT / ".env")
    load_dotenv(Path.cwd() / ".env")


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


def build_report_meta(processed_root: Path, workbook_path: Path) -> ReportMeta:
    rel_path = workbook_path.relative_to(processed_root)
    rel_text = str(rel_path).replace("\\", "/")
    group = rel_path.parts[0]
    stem = workbook_path.stem
    stem = re.sub(r"_processed$", "", stem)
    clean_stem = DATE_IN_NAME_RE.sub("", stem).replace("__", "_").strip("_ ")
    report_code = sanitize_slug(f"{group}_{clean_stem}")
    report_name = clean_stem.replace("_", " ").strip() or stem
    category = GROUP_CATEGORY_MAP.get(group, "general")
    lower_meta = f"{group}/{stem}".lower()

    service_type = None
    if "mytv" in lower_meta:
        service_type = "MyTV"
    elif any(token in lower_meta for token in ("fiber", "phieu_hoan_cong", "tam_dung", "ngung_psc")):
        service_type = "Fiber"

    action_type = None
    if "hoan_cong" in lower_meta:
        action_type = "hoan_cong"
    elif "thuc_tang" in lower_meta:
        action_type = "thuc_tang"
    elif "khoi_phuc" in lower_meta:
        action_type = "khoi_phuc"
    elif any(token in lower_meta for token in ("ngung_psc", "tam_dung")):
        action_type = "ngung_psc"

    return ReportMeta(
        report_code=report_code,
        report_name=report_name,
        report_group=group,
        report_category=category,
        service_type=service_type,
        action_type=action_type,
        processed_rel_path=rel_text,
    )


def parse_snapshot_date(explicit_value: Optional[str], workbook_path: Path) -> date:
    if explicit_value:
        return date.fromisoformat(explicit_value)

    match = DATE_IN_NAME_RE.search(workbook_path.name)
    if not match:
        raise ValueError(
            f"Khong xac dinh duoc snapshot_date tu ten file {workbook_path.name}. "
            "Hay truyen --snapshot-date YYYY-MM-DD."
        )

    digits = match.group(1)
    return datetime.strptime(digits, "%d%m%Y").date()


def sha256_bytes(content: bytes) -> str:
    return hashlib.sha256(content).hexdigest()


def file_content_type(path: Path) -> str:
    mime, _ = mimetypes.guess_type(path.name)
    return mime or "application/octet-stream"


def make_run_key(report_code: str, snapshot_date: date, processed_rel_path: str) -> str:
    return f"{report_code}:{snapshot_date.isoformat()}:{processed_rel_path}"


def chunked(rows: Sequence[Dict[str, Any]], size: int) -> Iterable[Sequence[Dict[str, Any]]]:
    for idx in range(0, len(rows), size):
        yield rows[idx : idx + size]


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
    parse_candidates = []
    if re.search(r"\b(am|pm)\b", text, flags=re.IGNORECASE):
        parse_candidates = [False, True]
    else:
        parse_candidates = [True, False]
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
        if value.is_integer():
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
        ratio = parsed.notna().sum() / len(series)
        if ratio >= 0.8:
            numeric_columns.append(column)
    return numeric_columns


def first_value(record: Dict[str, Any], candidates: Sequence[str]) -> Optional[str]:
    normalized = {normalize_key(key): key for key in record}
    for candidate in candidates:
        key = normalized.get(normalize_key(candidate))
        if not key:
            continue
        value = record.get(key)
        if value is None:
            continue
        text = str(value).strip()
        if text:
            return text
    return None


def extract_context(
    record: Dict[str, Any],
    report_meta: ReportMeta,
) -> Dict[str, Optional[str]]:
    ttvt = first_value(record, TTVT_KEYS)
    team_name = first_value(record, TEAM_KEYS)
    unit_name = first_value(record, UNIT_KEYS)
    employee_code = first_value(record, EMPLOYEE_CODE_KEYS)
    employee_name = first_value(record, EMPLOYEE_NAME_KEYS)

    service_type = report_meta.service_type
    if not service_type:
        raw_service = first_value(record, ("ten_dvvt_hni", "ten_loaihinh_tb", "loaihinh_tb", "loai_dich_vu"))
        if raw_service:
            service_type = infer_service_type(raw_service)

    action_type = report_meta.action_type
    if not action_type:
        raw_action = first_value(record, ("ten_kieuld", "ten_loaihd", "loai_phieu"))
        if raw_action:
            action_type = infer_action_type(raw_action)

    return {
        "ttvt": ttvt,
        "team_name": team_name,
        "unit_name": unit_name,
        "employee_code": employee_code,
        "employee_name": employee_name,
        "service_type": service_type,
        "action_type": action_type,
    }


def infer_service_type(value: str) -> Optional[str]:
    text = strip_accents(value)
    if "mytv" in text:
        return "MyTV"
    if any(token in text for token in ("fiber", "cap quang", "internet")):
        return "Fiber"
    return None


def infer_action_type(value: str) -> Optional[str]:
    text = strip_accents(value)
    if "hoan cong" in text:
        return "hoan_cong"
    if any(token in text for token in ("khoi phuc", "mo khoa")):
        return "khoi_phuc"
    if any(token in text for token in ("ngung", "tam dung", "khoa may")):
        return "ngung_psc"
    if "thuc tang" in text:
        return "thuc_tang"
    if "xac minh" in text:
        return "xac_minh"
    return None


def infer_entity_key(record: Dict[str, Any]) -> Optional[str]:
    normalized = {normalize_key(key): key for key in record}
    for candidate in ENTITY_KEY_KEYS:
        key = normalized.get(candidate)
        if not key:
            continue
        value = record.get(key)
        if value is None:
            continue
        text = str(value).strip()
        if text:
            return text
    return None


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


def classify_sheet(sheet_name: str, df: pd.DataFrame, numeric_columns: Sequence[str]) -> str:
    lower_name = sheet_name.strip().lower()
    row_count = len(df.index)
    col_count = len(df.columns)

    if lower_name in SOURCE_SHEET_NAMES:
        return "source"
    if lower_name in NOTE_SHEET_NAMES:
        return "note"
    if lower_name in DETAIL_SHEET_NAMES:
        return "detail"
    if any(token in lower_name for token in DETAIL_NAME_HINTS):
        return "detail"
    if "chi_tiet" in lower_name or "chi tiết" in lower_name:
        if row_count <= 500 and col_count <= 12 and numeric_columns:
            return "summary"
        return "detail"
    if any(token in lower_name for token in SUMMARY_NAME_HINTS):
        return "summary"
    if row_count <= 200 and numeric_columns:
        return "summary"
    return "other"


def build_dimensions(record: Dict[str, Any], numeric_columns: Sequence[str]) -> Dict[str, Any]:
    numeric_set = set(numeric_columns)
    dimensions: Dict[str, Any] = {}
    for key, value in record.items():
        if key in numeric_set:
            continue
        normalized = jsonable_value(value)
        if normalized is None:
            continue
        dimensions[key] = normalized
    return dimensions


def build_metrics(record: Dict[str, Any], numeric_columns: Sequence[str]) -> Dict[str, Any]:
    metrics: Dict[str, Any] = {}
    for column in numeric_columns:
        parsed = parse_numeric(record.get(column))
        if parsed is None:
            continue
        metrics[column] = int(parsed) if float(parsed).is_integer() else parsed
    return metrics


def infer_metric_level(context: Dict[str, Optional[str]], subject_type: Optional[str]) -> str:
    if context.get("employee_name") or context.get("employee_code"):
        return "nvkt"
    if context.get("team_name"):
        return "team"
    if context.get("unit_name"):
        return "unit"
    if context.get("ttvt"):
        return "ttvt"
    if subject_type:
        return "subject"
    return "total"


def prepare_sheet(
    report_meta: ReportMeta,
    run_id: uuid.UUID,
    snapshot_date: date,
    period_start: Optional[date],
    period_end: Optional[date],
    report_month: Optional[int],
    report_year: Optional[int],
    sheet_name: str,
    sheet_order: int,
    df: pd.DataFrame,
) -> SheetPlan:
    df = drop_empty_records(df.copy())
    df.columns = clean_column_names(df.columns)
    numeric_columns = detect_numeric_columns(df)
    text_columns = [column for column in df.columns if column not in numeric_columns]
    sheet_kind = classify_sheet(sheet_name, df, numeric_columns)

    raw_rows: List[Dict[str, Any]] = []
    metric_rows: List[Dict[str, Any]] = []
    detail_rows: List[Dict[str, Any]] = []
    fallback_detail_rows: List[Dict[str, Any]] = []
    sheet_id = uuid.uuid5(run_id, sheet_name)
    standard_dimension_keys = {
        normalize_key(value)
        for value in TTVT_KEYS + TEAM_KEYS + UNIT_KEYS + EMPLOYEE_CODE_KEYS + EMPLOYEE_NAME_KEYS
    }
    source_can_be_fallback_detail = sheet_kind == "source" and len(df.index) >= 50

    for row_number, row in enumerate(df.to_dict(orient="records"), start=1):
        record = {key: jsonable_value(value) for key, value in row.items()}
        context = extract_context(record, report_meta)
        entity_key = infer_entity_key(record)
        payload_json = json.dumps(record, ensure_ascii=False, sort_keys=True)
        payload_hash = hashlib.sha256(payload_json.encode("utf-8")).hexdigest()

        raw_rows.append(
            {
                "run_id": str(run_id),
                "sheet_id": str(sheet_id),
                "sheet_name": sheet_name,
                "row_number": row_number,
                "entity_key": entity_key,
                "ttvt": context["ttvt"],
                "team_name": context["team_name"],
                "unit_name": context["unit_name"],
                "employee_code": context["employee_code"],
                "employee_name": context["employee_name"],
                "service_type": context["service_type"],
                "action_type": context["action_type"],
                "payload": record,
                "payload_hash": payload_hash,
            }
        )

        if sheet_kind == "summary":
            metrics = build_metrics(record, numeric_columns)
            if metrics:
                dimensions = build_dimensions(record, numeric_columns)
                subject_type = None
                subject_name = None
                for key, value in dimensions.items():
                    normalized_key = normalize_key(key)
                    if normalized_key in standard_dimension_keys:
                        continue
                    subject_type = key
                    subject_name = str(value)
                    break
                metric_level = infer_metric_level(context, subject_type)
                metric_hash_source = json.dumps(
                    {
                        "sheet_name": sheet_name,
                        "metric_level": metric_level,
                        "subject_type": subject_type,
                        "subject_name": subject_name,
                        "ttvt": context["ttvt"],
                        "team_name": context["team_name"],
                        "unit_name": context["unit_name"],
                        "employee_code": context["employee_code"],
                        "employee_name": context["employee_name"],
                        "dimensions": dimensions,
                        "metrics": metrics,
                    },
                    ensure_ascii=False,
                    sort_keys=True,
                )
                metric_rows.append(
                    {
                        "run_id": str(run_id),
                        "report_code": report_meta.report_code,
                        "snapshot_date": snapshot_date.isoformat(),
                        "period_start": period_start.isoformat() if period_start else None,
                        "period_end": period_end.isoformat() if period_end else None,
                        "report_month": report_month,
                        "report_year": report_year,
                        "sheet_name": sheet_name,
                        "metric_level": metric_level,
                        "subject_type": subject_type,
                        "subject_name": subject_name,
                        "ttvt": context["ttvt"],
                        "team_name": context["team_name"],
                        "unit_name": context["unit_name"],
                        "employee_code": context["employee_code"],
                        "employee_name": context["employee_name"],
                        "service_type": context["service_type"],
                        "action_type": context["action_type"],
                        "dimensions": dimensions,
                        "metrics": metrics,
                        "metric_hash": hashlib.sha256(metric_hash_source.encode("utf-8")).hexdigest(),
                    }
                )

        if sheet_kind in {"detail", "other"}:
            occurred_at = infer_occurred_at(record)
            detail_rows.append(
                {
                    "run_id": str(run_id),
                    "report_code": report_meta.report_code,
                    "snapshot_date": snapshot_date.isoformat(),
                    "period_start": period_start.isoformat() if period_start else None,
                    "period_end": period_end.isoformat() if period_end else None,
                    "report_month": report_month,
                    "report_year": report_year,
                    "sheet_name": sheet_name,
                    "detail_level": "record",
                    "entity_key": entity_key,
                    "ttvt": context["ttvt"],
                    "team_name": context["team_name"],
                    "unit_name": context["unit_name"],
                    "employee_code": context["employee_code"],
                    "employee_name": context["employee_name"],
                    "service_type": context["service_type"],
                    "action_type": context["action_type"],
                    "occurred_at": occurred_at.isoformat() if occurred_at else None,
                    "payload": record,
                    "payload_hash": payload_hash,
                }
            )
        elif source_can_be_fallback_detail:
            occurred_at = infer_occurred_at(record)
            fallback_detail_rows.append(
                {
                    "run_id": str(run_id),
                    "report_code": report_meta.report_code,
                    "snapshot_date": snapshot_date.isoformat(),
                    "period_start": period_start.isoformat() if period_start else None,
                    "period_end": period_end.isoformat() if period_end else None,
                    "report_month": report_month,
                    "report_year": report_year,
                    "sheet_name": sheet_name,
                    "detail_level": "record",
                    "entity_key": entity_key,
                    "ttvt": context["ttvt"],
                    "team_name": context["team_name"],
                    "unit_name": context["unit_name"],
                    "employee_code": context["employee_code"],
                    "employee_name": context["employee_name"],
                    "service_type": context["service_type"],
                    "action_type": context["action_type"],
                    "occurred_at": occurred_at.isoformat() if occurred_at else None,
                    "payload": record,
                    "payload_hash": payload_hash,
                }
            )

    return SheetPlan(
        sheet_name=sheet_name,
        sheet_kind=sheet_kind,
        row_count=len(df.index),
        column_names=list(df.columns),
        numeric_columns=list(numeric_columns),
        text_columns=text_columns,
        raw_rows=raw_rows,
        metric_rows=metric_rows,
        detail_rows=detail_rows,
        fallback_detail_rows=fallback_detail_rows,
    )


def read_workbook_sheets(workbook_path: Path) -> List[Tuple[str, pd.DataFrame]]:
    excel = pd.ExcelFile(workbook_path, engine="openpyxl")
    sheets: List[Tuple[str, pd.DataFrame]] = []
    for sheet_name in excel.sheet_names:
        df = excel.parse(sheet_name=sheet_name, dtype=object)
        sheets.append((sheet_name, df))
    return sheets


def upload_local_file(
    client: SupabaseClient,
    bucket: str,
    object_path: str,
    path: Path,
) -> Dict[str, Any]:
    content = path.read_bytes()
    client.upload_bytes(bucket, object_path, content, file_content_type(path))
    return {
        "bucket": bucket,
        "object_path": object_path,
        "sha256": sha256_bytes(content),
        "size_bytes": len(content),
        "content_type": file_content_type(path),
    }


def build_storage_object_path(snapshot_date: date, report_code: str, suffix: str) -> str:
    prefix = snapshot_date.strftime("%Y/%m/%d")
    return f"{prefix}/{report_code}/{suffix}"


def ingest_workbook(
    client: Optional[SupabaseClient],
    workbook_path: Path,
    processed_root: Path,
    snapshot_date: date,
    period_kind: str,
    period_start: Optional[date],
    period_end: Optional[date],
    report_month: Optional[int],
    report_year: Optional[int],
    *,
    dry_run: bool,
    skip_storage: bool,
    replace_existing: bool,
    chunk_size: int,
) -> Dict[str, Any]:
    report_meta = build_report_meta(processed_root, workbook_path)
    run_key = make_run_key(report_meta.report_code, snapshot_date, report_meta.processed_rel_path)
    run_id = uuid.uuid5(uuid.NAMESPACE_URL, run_key)
    workbook_bytes = workbook_path.read_bytes()
    source_hash = sha256_bytes(workbook_bytes)

    run_row = {
        "run_id": str(run_id),
        "run_key": run_key,
        "report_code": report_meta.report_code,
        "snapshot_date": snapshot_date.isoformat(),
        "period_kind": period_kind,
        "period_start": period_start.isoformat() if period_start else None,
        "period_end": period_end.isoformat() if period_end else None,
        "report_month": report_month,
        "report_year": report_year,
        "local_file_name": workbook_path.name,
        "processed_rel_path": report_meta.processed_rel_path,
        "source_hash": source_hash,
        "status": "running",
        "params": {
            "snapshot_date": snapshot_date.isoformat(),
            "period_kind": period_kind,
            "period_start": period_start.isoformat() if period_start else None,
            "period_end": period_end.isoformat() if period_end else None,
            "report_month": report_month,
            "report_year": report_year,
        },
    }
    catalog_row = {
        "report_code": report_meta.report_code,
        "report_name": report_meta.report_name,
        "report_group": report_meta.report_group,
        "report_category": report_meta.report_category,
        "processed_rel_path": report_meta.processed_rel_path,
        "service_type": report_meta.service_type,
        "action_type": report_meta.action_type,
        "description": f"Auto-discovered from {report_meta.processed_rel_path}",
    }

    sheets = read_workbook_sheets(workbook_path)
    sheet_plans: List[SheetPlan] = []
    raw_count = 0
    metric_count = 0
    detail_count = 0

    for sheet_order, (sheet_name, df) in enumerate(sheets, start=1):
        plan = prepare_sheet(
            report_meta=report_meta,
            run_id=run_id,
            snapshot_date=snapshot_date,
            period_start=period_start,
            period_end=period_end,
            report_month=report_month,
            report_year=report_year,
            sheet_name=sheet_name,
            sheet_order=sheet_order,
            df=df,
        )
        sheet_plans.append(plan)
        raw_count += len(plan.raw_rows)
        metric_count += len(plan.metric_rows)
        detail_count += len(plan.detail_rows)

    if detail_count == 0:
        for plan in sheet_plans:
            if plan.fallback_detail_rows:
                plan.detail_rows.extend(plan.fallback_detail_rows)
                detail_count += len(plan.fallback_detail_rows)

    manifest = {
        "run_id": str(run_id),
        "run_key": run_key,
        "report_code": report_meta.report_code,
        "report_name": report_meta.report_name,
        "snapshot_date": snapshot_date.isoformat(),
        "processed_rel_path": report_meta.processed_rel_path,
        "file_sha256": source_hash,
        "counts": {
            "sheets": len(sheet_plans),
            "raw_rows": raw_count,
            "metric_rows": metric_count,
            "detail_rows": detail_count,
        },
        "sheets": [
            {
                "sheet_name": plan.sheet_name,
                "sheet_kind": plan.sheet_kind,
                "row_count": plan.row_count,
                "columns": plan.column_names,
                "numeric_columns": plan.numeric_columns,
            }
            for plan in sheet_plans
        ],
    }

    if not dry_run and client is not None:
        if replace_existing:
            client.delete_rows("meta_ingest_run", {"run_id": f"eq.{run_id}"})

        client.insert_rows(
            "meta_report_catalog",
            [catalog_row],
            upsert=True,
            on_conflict="report_code",
        )
        client.insert_rows(
            "meta_ingest_run",
            [run_row],
            upsert=True,
            on_conflict="run_key",
        )

        ingest_files: List[Dict[str, Any]] = []

        if not skip_storage:
            processed_object = build_storage_object_path(
                snapshot_date,
                report_meta.report_code,
                f"{workbook_path.stem}__{source_hash[:12]}.xlsx",
            )
            processed_meta = upload_local_file(
                client,
                "processed-reports",
                processed_object,
                workbook_path,
            )
            ingest_files.append(
                {
                    "run_id": str(run_id),
                    "file_kind": "processed",
                    "storage_bucket": processed_meta["bucket"],
                    "storage_object_path": processed_meta["object_path"],
                    "local_rel_path": report_meta.processed_rel_path,
                    "content_type": processed_meta["content_type"],
                    "size_bytes": processed_meta["size_bytes"],
                    "sha256": processed_meta["sha256"],
                    "sheet_count": len(sheet_plans),
                }
            )

            manifest_content = json.dumps(manifest, ensure_ascii=False, indent=2).encode("utf-8")
            manifest_object = build_storage_object_path(
                snapshot_date,
                report_meta.report_code,
                f"{workbook_path.stem}__{source_hash[:12]}.json",
            )
            client.upload_bytes(
                "report-manifests",
                manifest_object,
                manifest_content,
                "application/json",
            )
            ingest_files.append(
                {
                    "run_id": str(run_id),
                    "file_kind": "manifest",
                    "storage_bucket": "report-manifests",
                    "storage_object_path": manifest_object,
                    "local_rel_path": f"{report_meta.processed_rel_path}.manifest.json",
                    "content_type": "application/json",
                    "size_bytes": len(manifest_content),
                    "sha256": sha256_bytes(manifest_content),
                    "sheet_count": len(sheet_plans),
                }
            )

        for plan in sheet_plans:
            sheet_id = str(uuid.uuid5(run_id, plan.sheet_name))
            client.insert_rows(
                "raw_workbook_sheet",
                [
                    {
                        "sheet_id": sheet_id,
                        "run_id": str(run_id),
                        "sheet_name": plan.sheet_name,
                        "sheet_order": next(
                            idx for idx, (name, _) in enumerate(sheets, start=1) if name == plan.sheet_name
                        ),
                        "sheet_kind": plan.sheet_kind,
                        "column_names": plan.column_names,
                        "numeric_columns": plan.numeric_columns,
                        "text_columns": plan.text_columns,
                        "row_count": plan.row_count,
                    }
                ],
                upsert=True,
                on_conflict="run_id,sheet_name",
            )

            for batch in chunked(plan.raw_rows, chunk_size):
                client.insert_rows("raw_workbook_row", batch)
            for batch in chunked(plan.metric_rows, chunk_size):
                client.insert_rows("mart_metric_snapshot", batch)
            for batch in chunked(plan.detail_rows, chunk_size):
                client.insert_rows("mart_detail_record", batch)

        if ingest_files:
            client.insert_rows("meta_ingest_file", ingest_files, upsert=True, on_conflict="run_id,file_kind,local_rel_path")

        client.patch_rows(
            "meta_ingest_run",
            {"run_id": f"eq.{run_id}"},
            {
                "status": "success",
                "raw_row_count": raw_count,
                "metric_row_count": metric_count,
                "detail_row_count": detail_count,
                "summary": manifest["counts"],
                "finished_at": datetime.utcnow().isoformat(),
            },
        )

    return {
        "run_id": str(run_id),
        "report_code": report_meta.report_code,
        "report_name": report_meta.report_name,
        "processed_rel_path": report_meta.processed_rel_path,
        "snapshot_date": snapshot_date.isoformat(),
        "raw_rows": raw_count,
        "metric_rows": metric_count,
        "detail_rows": detail_count,
        "sheets": [
            {
                "sheet_name": plan.sheet_name,
                "sheet_kind": plan.sheet_kind,
                "row_count": plan.row_count,
            }
            for plan in sheet_plans
        ],
    }


def parse_optional_date(value: Optional[str]) -> Optional[date]:
    if not value:
        return None
    return date.fromisoformat(value)


def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Dong bo workbook processed len Supabase.")
    parser.add_argument("--processed-root", default=str(DEFAULT_PROCESSED_ROOT), help="Thu muc root cua workbook processed.")
    parser.add_argument("--snapshot-date", help="Ngay snapshot YYYY-MM-DD. Khuyen nghi truyen ro cho batch ngay.")
    parser.add_argument("--period-kind", default="snapshot", choices=["snapshot", "daily", "monthly", "custom"])
    parser.add_argument("--period-start", help="Ngay bat dau ky bao cao YYYY-MM-DD.")
    parser.add_argument("--period-end", help="Ngay ket thuc ky bao cao YYYY-MM-DD.")
    parser.add_argument("--report-month", type=int, help="Thang bao cao 1-12.")
    parser.add_argument("--report-year", type=int, help="Nam bao cao.")
    parser.add_argument("--path-contains", action="append", default=[], help="Chi ingest cac file co duong dan chua chuoi nay.")
    parser.add_argument("--limit", type=int, help="Chi ingest N file dau tien.")
    parser.add_argument("--chunk-size", type=int, default=200, help="So dong gui moi batch REST.")
    parser.add_argument("--skip-storage", action="store_true", help="Khong upload workbook/manifest len Supabase Storage.")
    parser.add_argument("--dry-run", action="store_true", help="Khong goi Supabase, chi doc workbook va in manifest.")
    parser.add_argument("--no-replace-existing", action="store_true", help="Khong xoa run cu cung run_id truoc khi import.")
    parser.add_argument("--json", action="store_true", help="In ket qua theo JSON.")
    return parser.parse_args(argv)


def main(argv: Optional[Sequence[str]] = None) -> int:
    load_environment()
    args = parse_args(argv)

    processed_root = Path(args.processed_root).expanduser()
    if not processed_root.is_absolute():
        processed_root = (Path.cwd() / processed_root).resolve()
    else:
        processed_root = processed_root.resolve()

    if not processed_root.exists():
        raise FileNotFoundError(f"Khong tim thay processed root: {processed_root}")

    workbook_paths = sorted(processed_root.rglob("*.xlsx"))
    if args.path_contains:
        workbook_paths = [
            path
            for path in workbook_paths
            if all(token.lower() in str(path.relative_to(processed_root)).lower() for token in args.path_contains)
        ]
    if args.limit:
        workbook_paths = workbook_paths[: args.limit]
    if not workbook_paths:
        raise ValueError("Khong co workbook nao khop bo loc.")

    client: Optional[SupabaseClient] = None
    if not args.dry_run:
        supabase_url = os.getenv("SUPABASE_URL")
        service_role_key = os.getenv("SUPABASE_SERVICE_ROLE_KEY")
        if not supabase_url or not service_role_key:
            raise EnvironmentError("Can SUPABASE_URL va SUPABASE_SERVICE_ROLE_KEY trong .env de ingest that.")
        client = SupabaseClient(supabase_url, service_role_key)

    results: List[Dict[str, Any]] = []
    exit_code = 0

    for workbook_path in workbook_paths:
        snapshot_date = parse_snapshot_date(args.snapshot_date, workbook_path)
        report_month = args.report_month or snapshot_date.month
        report_year = args.report_year or snapshot_date.year
        period_start = parse_optional_date(args.period_start)
        period_end = parse_optional_date(args.period_end)

        try:
            result = ingest_workbook(
                client=client,
                workbook_path=workbook_path,
                processed_root=processed_root,
                snapshot_date=snapshot_date,
                period_kind=args.period_kind,
                period_start=period_start,
                period_end=period_end,
                report_month=report_month,
                report_year=report_year,
                dry_run=args.dry_run,
                skip_storage=args.skip_storage,
                replace_existing=not args.no_replace_existing,
                chunk_size=args.chunk_size,
            )
            results.append(result)
        except Exception as exc:
            exit_code = 1
            error_payload = {
                "file": str(workbook_path.relative_to(processed_root)),
                "error": str(exc),
            }
            results.append(error_payload)
            if not args.dry_run and client is not None:
                try:
                    report_meta = build_report_meta(processed_root, workbook_path)
                    failed_snapshot_date = parse_snapshot_date(args.snapshot_date, workbook_path)
                    failed_run_key = make_run_key(
                        report_meta.report_code,
                        failed_snapshot_date,
                        report_meta.processed_rel_path,
                    )
                    failed_run_id = uuid.uuid5(uuid.NAMESPACE_URL, failed_run_key)
                    client.patch_rows(
                        "meta_ingest_run",
                        {"run_id": f"eq.{failed_run_id}"},
                        {
                            "status": "failed",
                            "error_message": str(exc),
                            "finished_at": datetime.utcnow().isoformat(),
                        },
                    )
                except Exception:
                    pass

    if args.json:
        print(json.dumps(results, ensure_ascii=False, indent=2))
    else:
        for item in results:
            if "error" in item:
                print(f"[ERROR] {item['file']}: {item['error']}")
            else:
                print(
                    f"[OK] {item['processed_rel_path']} -> {item['report_code']} | "
                    f"raw={item['raw_rows']} metric={item['metric_rows']} detail={item['detail_rows']}"
                )

    return exit_code


if __name__ == "__main__":
    sys.exit(main())
