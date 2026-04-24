# -*- coding: utf-8 -*-
"""Processors cho nhom vat tu trong api_transition."""

from pathlib import Path
import re

import pandas as pd

from api_transition.processors.common import append_or_replace_sheet, ensure_processed_workbook


DEFAULT_VATTU_THU_HOI_INPUT = (
    Path(__file__).resolve().parent.parent
    / "downloads"
    / "vat_tu_thu_hoi"
    / "bc_thu_hoi_vat_tu.xlsx"
)
def _resolve_path(input_path):
    path = Path(input_path).expanduser()
    if not path.is_absolute():
        path = (Path.cwd() / path).resolve()
    else:
        path = path.resolve()

    if not path.exists():
        raise FileNotFoundError(f"Khong tim thay file input: {path}")
    return path


def _normalize_person_name(value):
    if pd.isna(value):
        return None

    text = str(value).strip()
    if not text:
        return None

    if "-" in text:
        text = text.split("-", 1)[1].strip()

    text = re.sub(r"\([^)]*\)", "", text)
    text = re.sub(r"\s+", " ", text)
    return text.lower().title()


def _prepare_vattu_thu_hoi_df(input_path):
    raw_path = _resolve_path(input_path)
    df = pd.read_excel(raw_path).copy()
    df.columns = [str(col).strip() for col in df.columns]

    if "NVKT_DIABAN_GIAO" in df.columns:
        df["NVKT_DIABAN_GIAO"] = df["NVKT_DIABAN_GIAO"].apply(_normalize_person_name)

    filter_cols = {
        "TRANGTHAI_THUHOI": "CHUA_THU_HOI",
        "LOAI_VT": "ONT",
        "LOAI_PHIEU": "PTTB",
        "ONT_TBDC_CHAT_LUONG_CAO": 1,
        "LOAI_GIAO_GIAMTRU_LAN_HAI": 0,
    }
    for col, value in filter_cols.items():
        if col in df.columns:
            df = df[df[col] == value]

    preferred_cols = [
        "NVKT_DIABAN_GIAO",
        "TRANGTHAI_THUHOI",
        "LOAI_VT",
        "LOAI_PHIEU",
        "ONT_TBDC_CHAT_LUONG_CAO",
        "LOAI_GIAO_GIAMTRU_LAN_HAI",
    ]
    ordered = [col for col in preferred_cols if col in df.columns] + [col for col in df.columns if col not in preferred_cols]
    df = df[ordered].reset_index(drop=True)
    return raw_path, df


def _build_vattu_thuhoi_summary(df):
    if "DIEMCHIA" not in df.columns or "NVKT_DIABAN_GIAO" not in df.columns:
        return pd.DataFrame(columns=["DIEMCHIA", "NVKT_DIABAN_GIAO", "Số lượng"])

    temp = df.copy()
    temp["DIEMCHIA"] = temp["DIEMCHIA"].fillna("(Chưa xác định)").astype(str).str.strip()
    temp["NVKT_DIABAN_GIAO"] = temp["NVKT_DIABAN_GIAO"].fillna("(Chưa xác định)").astype(str).str.strip()
    return (
        temp.groupby(["DIEMCHIA", "NVKT_DIABAN_GIAO"], as_index=False)
        .size()
        .rename(columns={"size": "Số lượng"})
        .sort_values(["DIEMCHIA", "NVKT_DIABAN_GIAO"], ascending=[True, True])
        .reset_index(drop=True)
    )


def _build_vattu_chi_tiet(df):
    detail_cols = [
        "DIEMCHIA",
        "NVKT_DIABAN_GIAO",
        "MA_TB",
        "TEN_TB",
        "TEN_TBI",
        "NGAY_GIAO",
        "TEN_LOAIHD",
        "TEN_KIEULD",
        "SO_DT",
        "NGAY_SD_TB",
    ]
    existing = [col for col in detail_cols if col in df.columns]
    if not existing:
        return pd.DataFrame()
    return df[existing].copy().reset_index(drop=True)


def process_vat_tu_thu_hoi_api_output(
    input_path=DEFAULT_VATTU_THU_HOI_INPUT,
    overwrite_processed=False,
):
    """Xu ly bao cao vat tu thu hoi."""
    raw_path, df = _prepare_vattu_thu_hoi_df(input_path)
    df_detail = _build_vattu_chi_tiet(df)
    df_summary = _build_vattu_thuhoi_summary(df)

    processed_path = ensure_processed_workbook(raw_path, overwrite=overwrite_processed)
    append_or_replace_sheet(processed_path, "Chi tiết", df)
    if not df_detail.empty:
        append_or_replace_sheet(processed_path, "Chi tiết vật tư", df_detail)
    if not df_summary.empty:
        append_or_replace_sheet(processed_path, "Tổng hợp", df_summary)
    return processed_path
