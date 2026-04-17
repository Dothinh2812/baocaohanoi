# -*- coding: utf-8 -*-
"""Processors cho nhom ty le xac minh dung thoi gian quy dinh."""

from pathlib import Path
import re

import pandas as pd

from api_transition.processors.common import append_or_replace_sheet, ensure_processed_workbook


DEFAULT_XM_TTVTKV_INPUT = (
    Path(__file__).resolve().parent.parent
    / "downloads"
    / "ty_le_xac_minh"
    / "ty_le_xac_minh_dung_thoi_gian_quy_dinh_ttvtkv.xlsx"
)
DEFAULT_XM_CHI_TIET_INPUT = (
    Path(__file__).resolve().parent.parent
    / "downloads"
    / "ty_le_xac_minh"
    / "ty_le_xac_minh_dung_thoi_gian_quy_dinh_chi_tiet.xlsx"
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


def _extract_nvkt_from_ten_kv(value):
    if pd.isna(value):
        return None

    text = str(value).strip()
    if not text:
        return None

    parts = [part.strip() for part in text.split("-") if part.strip()]
    if len(parts) >= 2:
        return _normalize_person_name(parts[-1])
    return _normalize_person_name(text)


def _append_total_row(df, total_col, fixed_values):
    total_row = {col: "" for col in df.columns}
    total_row[total_col] = int(pd.to_numeric(df[total_col], errors="coerce").fillna(0).sum())
    for key, value in fixed_values.items():
        total_row[key] = value
    return pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)


def process_ty_le_xac_minh_dung_thoi_gian_quy_dinh_ttvtkv_api_output(
    input_path=DEFAULT_XM_TTVTKV_INPUT,
    overwrite_processed=False,
    sheet_name="tong_hop_ttvtkv",
):
    """Xu ly bao cao tong hop ty le xac minh dung thoi gian quy dinh cap to."""
    raw_path = _resolve_path(input_path)
    df = pd.read_excel(raw_path).copy()
    df.columns = [str(col).strip() for col in df.columns]

    numeric_cols = [col for col in df.columns if col != "Đơn vị"]
    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    if "Tỷ lệ phiếu đã phê duyệt XM đúng thời gian QĐ" in df.columns:
        df = df.sort_values(
            "Tỷ lệ phiếu đã phê duyệt XM đúng thời gian QĐ",
            ascending=True,
            na_position="last",
        ).reset_index(drop=True)

    if "Tổng số phiếu giao XM" in df.columns:
        total_row = {col: "" for col in df.columns}
        total_row["Đơn vị"] = "TỔNG CỘNG"
        for col in numeric_cols:
            if "Tỷ lệ" not in col:
                total_row[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).sum()
        df = pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)

    processed_path = ensure_processed_workbook(raw_path, overwrite=overwrite_processed)
    append_or_replace_sheet(processed_path, sheet_name, df)
    return processed_path


def _prepare_xm_chi_tiet_df(input_path):
    raw_path = _resolve_path(input_path)
    df = pd.read_excel(raw_path).copy()
    df.columns = [str(col).strip() for col in df.columns]

    df["NVKT"] = df["TEN_KV"].apply(_extract_nvkt_from_ten_kv)
    df["DOIVT"] = df["DOIVT"].fillna("(Chưa xác định)").astype(str).str.strip()
    df["TTVT"] = df["TTVT"].fillna("(Chưa xác định)").astype(str).str.strip()
    df["LOAI_PHIEU"] = df["LOAI_PHIEU"].fillna("(Chưa xác định)").astype(str).str.strip()
    df["TEN_KIEULD.1"] = df["TEN_KIEULD.1"].fillna("(Chưa xác định)").astype(str).str.strip()
    df["NVKT"] = df["NVKT"].fillna("(Chưa xác định)")
    return raw_path, df


def _build_xm_nvkt_summary(df):
    grouped = (
        df.groupby(["TTVT", "DOIVT", "NVKT"], dropna=False)
        .size()
        .reset_index(name="SỐ PHIẾU XM")
        .sort_values(["TTVT", "DOIVT", "SỐ PHIẾU XM", "NVKT"], ascending=[True, True, False, True])
        .reset_index(drop=True)
    )
    return _append_total_row(
        grouped,
        "SỐ PHIẾU XM",
        {"TTVT": "TỔNG CỘNG", "DOIVT": "", "NVKT": ""},
    )


def _build_xm_team_summary(df):
    grouped = (
        df.groupby(["TTVT", "DOIVT"], dropna=False)
        .size()
        .reset_index(name="SỐ PHIẾU XM")
        .sort_values(["TTVT", "SỐ PHIẾU XM", "DOIVT"], ascending=[True, False, True])
        .reset_index(drop=True)
    )
    return _append_total_row(
        grouped,
        "SỐ PHIẾU XM",
        {"TTVT": "TỔNG CỘNG", "DOIVT": ""},
    )


def _build_xm_loai_phieu_summary(df):
    grouped = (
        df.groupby(["LOAI_PHIEU", "TEN_KIEULD.1"], dropna=False)
        .size()
        .reset_index(name="SỐ PHIẾU XM")
        .sort_values(["SỐ PHIẾU XM", "LOAI_PHIEU", "TEN_KIEULD.1"], ascending=[False, True, True])
        .reset_index(drop=True)
    )
    return _append_total_row(
        grouped,
        "SỐ PHIẾU XM",
        {"LOAI_PHIEU": "TỔNG CỘNG", "TEN_KIEULD.1": ""},
    )


def process_ty_le_xac_minh_dung_thoi_gian_quy_dinh_chi_tiet_api_output(
    input_path=DEFAULT_XM_CHI_TIET_INPUT,
    overwrite_processed=False,
):
    """Xu ly bao cao chi tiet ty le xac minh dung thoi gian quy dinh."""
    raw_path, df = _prepare_xm_chi_tiet_df(input_path)
    df_nvkt = _build_xm_nvkt_summary(df)
    df_team = _build_xm_team_summary(df)
    df_loai = _build_xm_loai_phieu_summary(df)

    processed_path = ensure_processed_workbook(raw_path, overwrite=overwrite_processed)
    append_or_replace_sheet(processed_path, "Data", df)
    append_or_replace_sheet(processed_path, "tong_hop_theo_nvkt", df_nvkt)
    append_or_replace_sheet(processed_path, "tong_hop_theo_to", df_team)
    append_or_replace_sheet(processed_path, "tong_hop_theo_loai_phieu", df_loai)
    return processed_path
