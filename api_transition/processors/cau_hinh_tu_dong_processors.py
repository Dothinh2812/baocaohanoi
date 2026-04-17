# -*- coding: utf-8 -*-
"""Processors cho nhom cau hinh tu dong trong api_transition."""

from pathlib import Path
import re

import pandas as pd

from api_transition.processors.common import append_or_replace_sheet, ensure_processed_workbook


DEFAULT_CHTD_PTM_INPUT = (
    Path(__file__).resolve().parent.parent
    / "downloads"
    / "cau_hinh_tu_dong"
    / "cau_hinh_tu_dong_ptm.xlsx"
)
DEFAULT_CHTD_THAY_THE_INPUT = (
    Path(__file__).resolve().parent.parent
    / "downloads"
    / "cau_hinh_tu_dong"
    / "cau_hinh_tu_dong_thay_the.xlsx"
)
DEFAULT_CHTD_DETAIL_INPUT = (
    Path(__file__).resolve().parent.parent
    / "downloads"
    / "cau_hinh_tu_dong"
    / "cau_hinh_tu_dong_chi_tiet.xlsx"
)


SUMMARY_RENAME_MAP = {
    "Số hợp đồng (2)": "Tổng hợp đồng",
    "Số hợp đồng không thực hiện cấu hình tự động (3)= (2)-(4)": "Không thực hiện cấu hình tự động",
    "Số hợp đồng đã đẩy cấu hình tự động (4)": "Đã đẩy cấu hình tự động",
    "Không đẩy cấu hình tự động do lỗi hệ thống (5)": "Không đẩy do lỗi hệ thống",
    "Không đẩy cấu hình tự động do TBI có cấu hình trước (6)": "Không đẩy do TBI đã có cấu hình",
    "Số phiếu cấu hình thành công (7)": "Cấu hình thành công",
    "Số phiếu đẩy cầu hình tự động/Số phiếu PTM (8) = (4)/(2)": "Tỷ lệ đẩy tự động (%)",
    "Tỉ lệ Số phiếu không thành công do cấu hình TBI đã có cấu hình trước/Số phiếu đã đẩy cấu hình tự động (9)=(6)/(4)": "Tỷ lệ TBI đã có cấu hình (%)",
    "Số phiếu cấu hình thành công/Số phiếu đã đẩy cấu hình tự động (10)=(7)/(4)": "Tỷ lệ cấu hình thành công (%)",
}

SUMMARY_NUMERIC_COLUMNS = [
    "Tổng hợp đồng",
    "Không thực hiện cấu hình tự động",
    "Đã đẩy cấu hình tự động",
    "Không đẩy do lỗi hệ thống",
    "Không đẩy do TBI đã có cấu hình",
    "Cấu hình thành công",
    "Tỷ lệ đẩy tự động (%)",
    "Tỷ lệ TBI đã có cấu hình (%)",
    "Tỷ lệ cấu hình thành công (%)",
]


def _resolve_path(input_path):
    path = Path(input_path).expanduser()
    if not path.is_absolute():
        path = (Path.cwd() / path).resolve()
    else:
        path = path.resolve()

    if not path.exists():
        raise FileNotFoundError(f"Khong tim thay file input: {path}")
    return path


def _strip_html_tags(value):
    if pd.isna(value):
        return None
    text = str(value).strip()
    if not text:
        return None
    text = re.sub(r"<[^>]+>", "", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


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


def _extract_employee_code(value):
    if pd.isna(value):
        return None
    text = str(value).strip()
    if not text:
        return None
    if "-" in text:
        return text.split("-", 1)[0].strip()
    return None


def _prepare_cau_hinh_summary_df(input_path):
    raw_path = _resolve_path(input_path)
    df = pd.read_excel(raw_path).copy()
    df.columns = [str(col).strip() for col in df.columns]
    df = df.rename(columns=SUMMARY_RENAME_MAP)

    raw_units = df["Đơn vị"].copy()
    cleaned_units = raw_units.apply(_strip_html_tags)
    is_ttvt_row = raw_units.astype(str).str.contains(r"<\s*b\s*>", case=False, regex=True, na=False)
    is_ttvt_row = is_ttvt_row | cleaned_units.fillna("").str.startswith("TTVT ")

    df["Đơn vị"] = cleaned_units
    df["Loại dòng"] = is_ttvt_row.map({True: "TTVT", False: "Tổ"})
    df["TTVT"] = cleaned_units.where(is_ttvt_row).ffill()

    for col in SUMMARY_NUMERIC_COLUMNS:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    ordered_cols = ["TTVT", "Đơn vị", "Loại dòng"] + SUMMARY_NUMERIC_COLUMNS
    df = df[ordered_cols].copy()

    df_ttvt = df[df["Loại dòng"] == "TTVT"].reset_index(drop=True)
    df_to = df[df["Loại dòng"] == "Tổ"].reset_index(drop=True)
    return raw_path, df.reset_index(drop=True), df_ttvt, df_to


def _prepare_cau_hinh_detail_df(input_path):
    raw_path = _resolve_path(input_path)
    df = pd.read_excel(raw_path).copy()
    df.columns = [str(col).strip() for col in df.columns]

    df["Mã nhân viên"] = df["Nhân viên phụ trách"].apply(_extract_employee_code)
    df["NVKT"] = df["Nhân viên phụ trách"].apply(_normalize_person_name)
    df["Trang thái chuẩn hóa"] = df["Trang thái"].fillna("Chưa có trạng thái").astype(str).str.strip()
    df["Loại hợp đồng"] = df["Loại hợp đồng"].fillna("").astype(str).str.strip()
    df["Loại cấu hình"] = df["Loại cấu hình"].fillna("").astype(str).str.strip()
    df["Mã lỗi"] = df["Mã lỗi"].fillna("").astype(str).str.strip()

    ordered_cols = [
        "STT",
        "Serial Number",
        "Mã thuê bao",
        "Loại hợp đồng",
        "Loại cấu hình",
        "Trang thái",
        "Trang thái chuẩn hóa",
        "Mã lỗi",
        "Thời gian cập nhật",
        "Trung tâm Viễn thông",
        "Đội Viễn thông",
        "Mã nhân viên",
        "NVKT",
    ]
    df = df[ordered_cols].copy()
    return raw_path, df


def _build_cau_hinh_detail_summary(df, group_cols):
    df_summary_base = df.dropna(subset=group_cols).copy()
    grouped = (
        df_summary_base.groupby(group_cols, dropna=False)
        .agg(
            **{
                "Tổng hợp đồng": ("Serial Number", "size"),
                "Lắp mới": ("Loại hợp đồng", lambda s: (s == "Lắp mới").sum()),
                "Thay thế": ("Loại hợp đồng", lambda s: (s == "Thay thế").sum()),
                "Cấu hình WAN": ("Loại cấu hình", lambda s: (s == "Cấu hình WAN").sum()),
                "Cấu hình WiFi": ("Loại cấu hình", lambda s: (s == "Cấu hình WiFi").sum()),
                "Thành công": ("Trang thái chuẩn hóa", lambda s: (s == "Thành công").sum()),
                "Thất bại": ("Trang thái chuẩn hóa", lambda s: (s == "Thất bại").sum()),
                "Chưa có trạng thái": (
                    "Trang thái chuẩn hóa",
                    lambda s: (s == "Chưa có trạng thái").sum(),
                ),
            }
        )
        .reset_index()
    )

    grouped["Tỷ lệ thành công (%)"] = (
        grouped["Thành công"].div(grouped["Tổng hợp đồng"]).replace([float("inf")], 0).fillna(0) * 100
    ).round(2)
    grouped["Tỷ lệ thất bại (%)"] = (
        grouped["Thất bại"].div(grouped["Tổng hợp đồng"]).replace([float("inf")], 0).fillna(0) * 100
    ).round(2)

    return grouped


def _build_error_summary(df):
    error_df = df[df["Mã lỗi"].astype(str).str.strip() != ""].copy()
    if error_df.empty:
        return pd.DataFrame(columns=["Mã lỗi", "Số lượng"])

    return (
        error_df.groupby("Mã lỗi")
        .size()
        .reset_index(name="Số lượng")
        .sort_values(["Số lượng", "Mã lỗi"], ascending=[False, True])
        .reset_index(drop=True)
    )


def process_cau_hinh_tu_dong_ptm_api_output(
    input_path=DEFAULT_CHTD_PTM_INPUT,
    overwrite_processed=False,
):
    """Xu ly bao cao cau hinh tu dong PTM va ghi vao file processed."""
    raw_path, df_clean, df_ttvt, df_to = _prepare_cau_hinh_summary_df(input_path)
    processed_path = ensure_processed_workbook(raw_path, overwrite=overwrite_processed)
    append_or_replace_sheet(processed_path, "du_lieu_sach", df_clean)
    append_or_replace_sheet(processed_path, "tong_hop_ttvt", df_ttvt)
    append_or_replace_sheet(processed_path, "tong_hop_to", df_to)
    return processed_path


def process_cau_hinh_tu_dong_thay_the_api_output(
    input_path=DEFAULT_CHTD_THAY_THE_INPUT,
    overwrite_processed=False,
):
    """Xu ly bao cao cau hinh tu dong thay the va ghi vao file processed."""
    raw_path, df_clean, df_ttvt, df_to = _prepare_cau_hinh_summary_df(input_path)
    processed_path = ensure_processed_workbook(raw_path, overwrite=overwrite_processed)
    append_or_replace_sheet(processed_path, "du_lieu_sach", df_clean)
    append_or_replace_sheet(processed_path, "tong_hop_ttvt", df_ttvt)
    append_or_replace_sheet(processed_path, "tong_hop_to", df_to)
    return processed_path


def process_cau_hinh_tu_dong_chi_tiet_api_output(
    input_path=DEFAULT_CHTD_DETAIL_INPUT,
    overwrite_processed=False,
):
    """Xu ly bao cao cau hinh tu dong chi tiet, tao tong hop theo to va theo NVKT."""
    raw_path, df_detail = _prepare_cau_hinh_detail_df(input_path)
    df_team = _build_cau_hinh_detail_summary(
        df_detail,
        ["Trung tâm Viễn thông", "Đội Viễn thông"],
    )
    df_nvkt = _build_cau_hinh_detail_summary(
        df_detail,
        ["Trung tâm Viễn thông", "Đội Viễn thông", "NVKT"],
    )
    df_errors = _build_error_summary(df_detail)

    processed_path = ensure_processed_workbook(raw_path, overwrite=overwrite_processed)
    append_or_replace_sheet(processed_path, "chi_tiet", df_detail)
    append_or_replace_sheet(processed_path, "th_theo_to", df_team)
    append_or_replace_sheet(processed_path, "th_theo_nvkt", df_nvkt)
    append_or_replace_sheet(processed_path, "tong_hop_loi", df_errors)
    return processed_path
