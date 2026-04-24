# -*- coding: utf-8 -*-
"""Processors cho nhom chi tieu C trong api_transition."""

from pathlib import Path
import re

import pandas as pd

from api_transition.processors.common import (
    DOWNLOADS_DIR,
    append_or_replace_sheet,
    ensure_processed_workbook,
)


DEFAULT_C11_INPUT = DOWNLOADS_DIR / "chi_tieu_c" / "c1.1 report.xlsx"
DEFAULT_C11_SHEET = "TH_C1.1"
DEFAULT_C12_INPUT = DOWNLOADS_DIR / "chi_tieu_c" / "c1.2 report.xlsx"
DEFAULT_C12_SHEET = "TH_C1.2"
DEFAULT_C13_INPUT = DOWNLOADS_DIR / "chi_tieu_c" / "c1.3 report.xlsx"
DEFAULT_C13_SHEET = "TH_C1.3"
DEFAULT_C14_INPUT = DOWNLOADS_DIR / "chi_tieu_c" / "c1.4 report.xlsx"
DEFAULT_C14_SHEET = "TH_C1.4"
DEFAULT_C15_INPUT = DOWNLOADS_DIR / "chi_tieu_c" / "c1.5 report.xlsx"
DEFAULT_C15_SHEET = "TH_C1.5"
DEFAULT_C15_DETAIL_INPUT = DOWNLOADS_DIR / "chi_tieu_c" / "c1.5_chitiet_report.xlsx"
DEFAULT_C15_DETAIL_SHEETS = (
    "KQ_C15_chitiet",
    "TH_TTVTST",
    "Chi_tiet_TG",
    "TH_KIEULD",
    "TH_DVVT",
    "TH_DVVT_DOI",
    "TH_DVVT_TTVT",
)
DEFAULT_C14_DETAIL_INPUT = DOWNLOADS_DIR / "chi_tieu_c" / "c1.4_chitiet_report.xlsx"
DEFAULT_C14_DETAIL_SHEET = "TH_HL_NVKT"
DEFAULT_C11_DETAIL_INPUT = DOWNLOADS_DIR / "chi_tieu_c" / "c1.1_chitiet_report.xlsx"
DEFAULT_C12_DETAIL_SM1_INPUT = DOWNLOADS_DIR / "chi_tieu_c" / "c1.2_chitiet_sm1_report.xlsx"
DEFAULT_C12_DETAIL_SM2_INPUT = DOWNLOADS_DIR / "chi_tieu_c" / "c1.2_chitiet_sm2_report.xlsx"
C11_DEFAULT_EXCLUDE_UNIT_PATTERNS = [
    "Tổ Kỹ thuật Địa bàn Bất Bạt",
    "Tổ Kỹ thuật Địa bàn Tùng Thiện",
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


def _apply_first_column_exclusions(df, exclude_patterns):
    """Loai tru dong dua tren cot dau tien neu co pattern cau hinh."""
    if not exclude_patterns:
        return df

    escaped_patterns = [
        re.escape(str(item)) for item in exclude_patterns if str(item).strip()
    ]
    if not escaped_patterns:
        return df

    pattern = "|".join(escaped_patterns)
    return df[
        ~df.iloc[:, 0].astype(str).str.contains(pattern, na=False)
    ].reset_index(drop=True)


def _flatten_excel_columns(columns):
    """Chuan hoa header Excel 2 dong thanh ten cot on dinh."""
    flattened = []
    for column in columns:
        if isinstance(column, tuple):
            parts = [str(part).strip() for part in column if str(part).strip() and not str(part).startswith("Unnamed:")]
            if not parts:
                flattened.append("")
                continue
            if len(parts) == 1 or parts[0] == parts[-1]:
                flattened.append(parts[-1])
            else:
                flattened.append(" - ".join(parts))
        else:
            flattened.append(str(column).strip())
    return flattened


def _calculate_c15_bsc(rate_percent):
    """Tinh diem BSC cho C1.5 tu ty le dang phan tram."""
    if pd.isna(rate_percent):
        return pd.NA

    rate_decimal = float(rate_percent) / 100 if float(rate_percent) > 1 else float(rate_percent)
    if rate_decimal >= 0.995:
        return 5.0
    if rate_decimal > 0.895:
        return round(1 + 4 * (rate_decimal - 0.895) / 0.10, 2)
    return 1.0


def _extract_nvkt_name(ten_kv):
    """Trich xuat ten NVKT tu TEN_KV theo cac mau thuong gap trong bao cao."""
    if pd.isna(ten_kv):
        return None

    ten_kv = str(ten_kv).strip()
    if not ten_kv:
        return None

    if "-" in ten_kv:
        nvkt_name = ten_kv.split("-")[-1].strip()
    else:
        nvkt_name = ten_kv

    if "(" in nvkt_name:
        nvkt_name = nvkt_name.split("(")[0].strip()

    return _normalize_person_name(nvkt_name)


def _normalize_person_name(name):
    """Chuan hoa ten nguoi de gom cac bien the khac nhau ve hoa/thuong/khoang trang."""
    if pd.isna(name):
        return None

    name = str(name).strip()
    if not name:
        return None

    name = re.sub(r"\s+", " ", name)
    return name.lower().title()


def _build_c11_detail_summary(df, has_team_column):
    """Tong hop C1.1 chi tiet theo NVKT hoac TEN_DOI + NVKT."""
    working_df = df.copy()
    working_df["DAT_TT_KO_HEN_NUM"] = pd.to_numeric(
        working_df["DAT_TT_KO_HEN"], errors="coerce"
    ).fillna(0)

    group_columns = ["TEN_DOI", "NVKT"] if has_team_column else ["NVKT"]
    grouped = (
        working_df.groupby(group_columns, dropna=False, as_index=False)
        .agg(
            tong_phieu=("NVKT", "size"),
            so_phieu_dat=("DAT_TT_KO_HEN_NUM", lambda s: int((s == 1).sum())),
        )
    )
    grouped = grouped.rename(
        columns={
            "tong_phieu": "Tổng phiếu",
            "so_phieu_dat": "Số phiếu đạt",
        }
    )
    grouped[
        "Tỷ lệ phiếu sửa chữa báo hỏng dịch vụ BRCD đúng quy định không tính hẹn"
    ] = grouped.apply(
        lambda row: round(row["Số phiếu đạt"] / row["Tổng phiếu"] * 100, 2)
        if row["Tổng phiếu"] > 0
        else 0,
        axis=1,
    )

    sort_columns = ["TEN_DOI", "NVKT"] if has_team_column else ["NVKT"]
    return grouped.sort_values(sort_columns, kind="stable").reset_index(drop=True)


def _filter_c11_detail_by_gio(df, max_hour_inclusive):
    """Loc ban ghi NGAYGIAO trong khoang 22h-den-max_hour cua ngay tiep theo."""
    if "NGAYGIAO" not in df.columns:
        raise ValueError("Bao cao C1.1 chi tiet thieu cot NGAYGIAO")

    filtered_df = df.copy()
    filtered_df["NGAYGIAO"] = pd.to_datetime(
        filtered_df["NGAYGIAO"],
        format="%d/%m/%Y %H:%M:%S",
        errors="coerce",
    )
    filtered_df = filtered_df[filtered_df["NGAYGIAO"].notna()].copy()
    filtered_df["gio_giao"] = filtered_df["NGAYGIAO"].dt.hour
    filtered_df = filtered_df[
        (filtered_df["gio_giao"] >= 22) | (filtered_df["gio_giao"] <= max_hour_inclusive)
    ].copy()
    return filtered_df


def _build_unique_ma_tb_summary(df, value_column_name, has_team_column):
    """Tong hop so MA_TB duy nhat theo NVKT hoac TEN_DOI + NVKT."""
    group_columns = ["TEN_DOI", "NVKT"] if has_team_column else ["NVKT"]
    grouped = (
        df.dropna(subset=["NVKT"])
        .groupby(group_columns, dropna=False, as_index=False)
        .agg(**{value_column_name: ("MA_TB", "nunique")})
    )
    sort_columns = ["TEN_DOI", "NVKT"] if has_team_column else ["NVKT"]
    return grouped.sort_values(sort_columns, kind="stable").reset_index(drop=True)


def _empty_c15_detail_outputs():
    return {
        "KQ_C15_chitiet": pd.DataFrame(
            columns=[
                "DOIVT",
                "NVKT",
                "Phiếu đạt",
                "Phiếu không đạt",
                "Tổng Hoàn công",
                "Tỉ lệ đạt (%)",
            ]
        ),
        "TH_TTVTST": pd.DataFrame(
            columns=[
                "DOIVT",
                "Phiếu đạt",
                "Phiếu không đạt",
                "Tổng Hoàn công",
                "Tỉ lệ đạt (%)",
            ]
        ),
        "Chi_tiet_TG": pd.DataFrame(
            columns=[
                "DOIVT",
                "MA_TB",
                "NGAY_LHD",
                "NGAY_HC",
                "TG_THI_CONG",
                "KET_QUA",
                "NVKT",
            ]
        ),
        "TH_KIEULD": pd.DataFrame(
            columns=[
                "DOIVT",
                "NVKT",
                "TEN_KIEULD",
                "Phiếu đạt",
                "Phiếu không đạt",
                "Tổng Hoàn công",
                "Tỉ lệ đạt (%)",
            ]
        ),
        "TH_DVVT": pd.DataFrame(
            columns=[
                "DOIVT",
                "NVKT",
                "TEN_DVVT",
                "Phiếu đạt",
                "Phiếu không đạt",
                "Tổng Hoàn công",
                "Tỉ lệ đạt (%)",
            ]
        ),
        "TH_DVVT_DOI": pd.DataFrame(
            columns=[
                "DOIVT",
                "TEN_DVVT",
                "Phiếu đạt",
                "Phiếu không đạt",
                "Tổng Hoàn công",
                "Tỉ lệ đạt (%)",
            ]
        ),
        "TH_DVVT_TTVT": pd.DataFrame(
            columns=[
                "TEN_DVVT",
                "Phiếu đạt",
                "Phiếu không đạt",
                "Tổng Hoàn công",
                "Tỉ lệ đạt (%)",
            ]
        ),
    }


def _summarize_c15_detail(df, group_columns, rename_map):
    grouped = (
        df.groupby(group_columns, dropna=False, as_index=False)
        .agg(
            **{
                "Phiếu đạt": ("KET_QUA", lambda s: int((s == "Đạt").sum())),
                "Phiếu không đạt": ("KET_QUA", lambda s: int((s == "Không đạt").sum())),
                "Tổng Hoàn công": ("KET_QUA", "size"),
            }
        )
    )
    grouped["Tỉ lệ đạt (%)"] = grouped.apply(
        lambda row: round(row["Phiếu đạt"] / row["Tổng Hoàn công"] * 100, 2)
        if row["Tổng Hoàn công"] > 0
        else 0,
        axis=1,
    )
    grouped = grouped.rename(columns=rename_map)
    return grouped


def process_c11_report_api_output(
    input_path=DEFAULT_C11_INPUT,
    overwrite_processed=False,
    sheet_name=DEFAULT_C11_SHEET,
    exclude_unit_patterns=None,
):
    """Xu ly bao cao C1.1 va ghi ket qua vao file _processed."""
    raw_path = _resolve_path(input_path)
    exclude_patterns = (
        C11_DEFAULT_EXCLUDE_UNIT_PATTERNS
        if exclude_unit_patterns is None
        else exclude_unit_patterns
    )

    df = pd.read_excel(raw_path)
    if df.shape[1] < 11:
        raise ValueError(
            f"Bao cao C1.1 khong du cot de xu ly. So cot hien tai: {df.shape[1]}"
        )

    df = df.iloc[1:].reset_index(drop=True)
    df = _apply_first_column_exclusions(df, exclude_patterns)

    df = df.iloc[:, :11].copy()

    df.columns = [
        "Đơn vị",
        "SM1",
        "SM2",
        "Tỷ lệ sửa chữa phiếu chất lượng chủ động dịch vụ FiberVNN, MyTV đạt yêu cầu",
        "SM3",
        "SM4",
        "Tỷ lệ phiếu sửa chữa báo hỏng dịch vụ BRCD đúng quy định không tính hẹn",
        "SM5",
        "SM6",
        "Tỷ lệ phiếu sửa chữa trong ngày tại CCCO",
        "Chỉ tiêu BSC",
    ]

    processed_path = ensure_processed_workbook(raw_path, overwrite=overwrite_processed)
    append_or_replace_sheet(processed_path, sheet_name, df)
    return processed_path


def process_c12_report_api_output(
    input_path=DEFAULT_C12_INPUT,
    overwrite_processed=False,
    sheet_name=DEFAULT_C12_SHEET,
    exclude_unit_patterns=None,
):
    """Xu ly bao cao C1.2 va ghi ket qua vao file _processed."""
    raw_path = _resolve_path(input_path)
    exclude_patterns = (
        C11_DEFAULT_EXCLUDE_UNIT_PATTERNS
        if exclude_unit_patterns is None
        else exclude_unit_patterns
    )

    df = pd.read_excel(raw_path)
    if df.shape[1] < 8:
        raise ValueError(
            f"Bao cao C1.2 khong du cot de xu ly. So cot hien tai: {df.shape[1]}"
        )

    df = df.iloc[1:].reset_index(drop=True)
    df = _apply_first_column_exclusions(df, exclude_patterns)

    df = df.iloc[:, :8].copy()

    df.columns = [
        "Đơn vị",
        "SM1",
        "SM2",
        "Tỷ lệ thuê bao báo hỏng dịch vụ BRCĐ lặp lại",
        "SM3",
        "SM4",
        "Tỷ lệ sự cố dịch vụ BRCĐ",
        "Chỉ tiêu BSC",
    ]

    processed_path = ensure_processed_workbook(raw_path, overwrite=overwrite_processed)
    append_or_replace_sheet(processed_path, sheet_name, df)
    return processed_path


def process_c13_report_api_output(
    input_path=DEFAULT_C13_INPUT,
    overwrite_processed=False,
    sheet_name=DEFAULT_C13_SHEET,
    exclude_unit_patterns=None,
):
    """Xu ly bao cao C1.3 va ghi ket qua vao file _processed."""
    raw_path = _resolve_path(input_path)
    exclude_patterns = (
        C11_DEFAULT_EXCLUDE_UNIT_PATTERNS
        if exclude_unit_patterns is None
        else exclude_unit_patterns
    )

    df = pd.read_excel(raw_path)
    if df.shape[1] < 11:
        raise ValueError(
            f"Bao cao C1.3 khong du cot de xu ly. So cot hien tai: {df.shape[1]}"
        )

    df = df.iloc[1:].reset_index(drop=True)
    df = _apply_first_column_exclusions(df, exclude_patterns)

    df = df.iloc[:, :11].copy()

    df.columns = [
        "Đơn vị",
        "SM1",
        "SM2",
        "Tỷ lệ sửa chữa dịch vụ kênh TSL hoàn thành đúng thời gian quy định",
        "SM3",
        "SM4",
        "Tỷ lệ thuê bao báo hỏng dịch vụ kênh TSL lặp lại",
        "SM5",
        "SM6",
        "Tỷ lệ sự cố dịch vụ kênh TSL",
        "Chỉ tiêu BSC",
    ]

    processed_path = ensure_processed_workbook(raw_path, overwrite=overwrite_processed)
    append_or_replace_sheet(processed_path, sheet_name, df)
    return processed_path


def process_c14_report_api_output(
    input_path=DEFAULT_C14_INPUT,
    overwrite_processed=False,
    sheet_name=DEFAULT_C14_SHEET,
):
    """Xu ly bao cao C1.4 va ghi ket qua vao file _processed."""
    raw_path = _resolve_path(input_path)

    df = pd.read_excel(raw_path)
    if df.shape[1] < 12:
        raise ValueError(
            f"Bao cao C1.4 khong du cot de xu ly. So cot hien tai: {df.shape[1]}"
        )

    df = df.iloc[1:].reset_index(drop=True)

    columns_to_keep = [i for i in range(df.shape[1]) if i not in [5, 6, 7, 8]]
    df_filtered = df.iloc[:, columns_to_keep].copy()

    for i in range(1, df_filtered.shape[1]):
        df_filtered.iloc[:, i] = pd.to_numeric(df_filtered.iloc[:, i], errors="coerce").fillna(0)

    tong_phieu = df_filtered.iloc[:, 1].sum()
    sl_da_ks = df_filtered.iloc[:, 2].sum()
    sl_ks_thanh_cong = df_filtered.iloc[:, 3].sum()
    sl_kh_hai_long = df_filtered.iloc[:, 4].sum()
    khong_hl_kt_phuc_vu = df_filtered.iloc[:, 5].sum()
    khong_hl_kt_dich_vu = df_filtered.iloc[:, 7].sum()
    tong_phieu_hai_long_kt = df_filtered.iloc[:, 9].sum()

    ty_le_hl_kt_phuc_vu = round((sl_kh_hai_long / sl_ks_thanh_cong * 100), 2) if sl_ks_thanh_cong > 0 else 0
    ty_le_hl_kt_dich_vu = round((khong_hl_kt_dich_vu / sl_ks_thanh_cong * 100), 2) if sl_ks_thanh_cong > 0 else 0
    ty_le_kh_hai_long = round((tong_phieu_hai_long_kt / sl_ks_thanh_cong * 100), 2) if sl_ks_thanh_cong > 0 else 0

    if ty_le_kh_hai_long >= 99.5:
        diem_bsc = 5
    elif ty_le_kh_hai_long <= 95:
        diem_bsc = 1
    else:
        diem_bsc = round(1 + 4 * (ty_le_kh_hai_long - 95) / 4.5, 2)

    tong_row = pd.Series({
        df_filtered.columns[0]: "Tổng",
        df_filtered.columns[1]: tong_phieu,
        df_filtered.columns[2]: sl_da_ks,
        df_filtered.columns[3]: sl_ks_thanh_cong,
        df_filtered.columns[4]: sl_kh_hai_long,
        df_filtered.columns[5]: khong_hl_kt_phuc_vu,
        df_filtered.columns[6]: ty_le_hl_kt_phuc_vu,
        df_filtered.columns[7]: khong_hl_kt_dich_vu,
        df_filtered.columns[8]: ty_le_hl_kt_dich_vu,
        df_filtered.columns[9]: tong_phieu_hai_long_kt,
        df_filtered.columns[10]: ty_le_kh_hai_long,
        df_filtered.columns[11]: diem_bsc,
    })
    df_filtered = pd.concat([df_filtered, tong_row.to_frame().T], ignore_index=True)

    processed_path = ensure_processed_workbook(raw_path, overwrite=overwrite_processed)
    append_or_replace_sheet(processed_path, sheet_name, df_filtered)
    return processed_path


def process_c15_report_api_output(
    input_path=DEFAULT_C15_INPUT,
    overwrite_processed=False,
    sheet_name=DEFAULT_C15_SHEET,
):
    """Xu ly bao cao C1.5 va ghi ket qua tong hop vao file _processed."""
    raw_path = _resolve_path(input_path)

    df = pd.read_excel(raw_path, header=[0, 1])
    if df.shape[1] < 17:
        raise ValueError(
            f"Bao cao C1.5 khong du cot de xu ly. So cot hien tai: {df.shape[1]}"
        )

    df = df.copy()
    df.columns = _flatten_excel_columns(df.columns)
    df = df.iloc[:, :17].copy()
    df.columns = [
        "Đơn vị",
        "Tổng - SM1",
        "Tổng - SM2",
        "Tổng - KQ thực hiện chỉ tiêu",
        "Tổng - Điểm BSC",
        "CCCO - SM1",
        "CCCO - SM2",
        "CCCO - Tỷ lệ",
        "CCCO - Điểm BSC",
        "Không CCCO - SM1",
        "Không CCCO - SM2",
        "Không CCCO - Tỷ lệ",
        "Không CCCO - Điểm BSC",
        "CCCO xã hội hóa - SM1",
        "CCCO xã hội hóa - SM2",
        "CCCO xã hội hóa - Tỷ lệ",
        "CCCO xã hội hóa - Điểm BSC",
    ]

    df["Đơn vị"] = df["Đơn vị"].astype(str).str.strip()
    df = df[df["Đơn vị"].ne("")].reset_index(drop=True)

    numeric_columns = [column for column in df.columns if column != "Đơn vị"]
    for column in numeric_columns:
        df[column] = pd.to_numeric(df[column], errors="coerce")

    total_sm1 = df["Tổng - SM1"].sum(min_count=1)
    total_sm2 = df["Tổng - SM2"].sum(min_count=1)
    total_rate = round(total_sm1 / total_sm2 * 100, 2) if total_sm2 and pd.notna(total_sm2) else pd.NA

    ccco_sm1 = df["CCCO - SM1"].sum(min_count=1)
    ccco_sm2 = df["CCCO - SM2"].sum(min_count=1)
    ccco_rate = round(ccco_sm1 / ccco_sm2 * 100, 2) if ccco_sm2 and pd.notna(ccco_sm2) else pd.NA

    khong_ccco_sm1 = df["Không CCCO - SM1"].sum(min_count=1)
    khong_ccco_sm2 = df["Không CCCO - SM2"].sum(min_count=1)
    khong_ccco_rate = (
        round(khong_ccco_sm1 / khong_ccco_sm2 * 100, 2)
        if khong_ccco_sm2 and pd.notna(khong_ccco_sm2)
        else pd.NA
    )

    xhh_sm1 = df["CCCO xã hội hóa - SM1"].sum(min_count=1)
    xhh_sm2 = df["CCCO xã hội hóa - SM2"].sum(min_count=1)
    xhh_rate = round(xhh_sm1 / xhh_sm2 * 100, 2) if xhh_sm2 and pd.notna(xhh_sm2) else pd.NA

    total_row = pd.DataFrame(
        [
            {
                "Đơn vị": "Tổng",
                "Tổng - SM1": total_sm1,
                "Tổng - SM2": total_sm2,
                "Tổng - KQ thực hiện chỉ tiêu": total_rate,
                "Tổng - Điểm BSC": _calculate_c15_bsc(total_rate),
                "CCCO - SM1": ccco_sm1,
                "CCCO - SM2": ccco_sm2,
                "CCCO - Tỷ lệ": ccco_rate,
                "CCCO - Điểm BSC": _calculate_c15_bsc(ccco_rate),
                "Không CCCO - SM1": khong_ccco_sm1,
                "Không CCCO - SM2": khong_ccco_sm2,
                "Không CCCO - Tỷ lệ": khong_ccco_rate,
                "Không CCCO - Điểm BSC": _calculate_c15_bsc(khong_ccco_rate),
                "CCCO xã hội hóa - SM1": xhh_sm1,
                "CCCO xã hội hóa - SM2": xhh_sm2,
                "CCCO xã hội hóa - Tỷ lệ": xhh_rate,
                "CCCO xã hội hóa - Điểm BSC": _calculate_c15_bsc(xhh_rate),
            }
        ]
    )
    df = pd.concat([df, total_row], ignore_index=True)

    processed_path = ensure_processed_workbook(raw_path, overwrite=overwrite_processed)
    append_or_replace_sheet(processed_path, sheet_name, df)
    return processed_path


def process_c15_chitiet_report_api_output(
    input_path=DEFAULT_C15_DETAIL_INPUT,
    overwrite_processed=False,
):
    """Xu ly bao cao chi tiet C1.5 va tao cac sheet tong hop on dinh cho multi-instance."""
    raw_path = _resolve_path(input_path)
    processed_path = ensure_processed_workbook(raw_path, overwrite=overwrite_processed)

    df = pd.read_excel(raw_path)
    if df.empty or df.shape[1] == 0:
        for sheet_name, empty_df in _empty_c15_detail_outputs().items():
            append_or_replace_sheet(processed_path, sheet_name, empty_df)
        return processed_path

    required_columns = ["NGAY_LHD", "NGAY_HC"]
    missing_columns = [column for column in required_columns if column not in df.columns]
    if missing_columns:
        raise ValueError(
            "Bao cao C1.5 chi tiet thieu cac cot can thiet: "
            + ", ".join(missing_columns)
        )

    nvkt_source_column = None
    for candidate in ("TEN_NVKT", "TEN_KV"):
        if candidate in df.columns:
            nvkt_source_column = candidate
            break
    if nvkt_source_column is None:
        raise ValueError("Bao cao C1.5 chi tiet thieu cot TEN_NVKT/TEN_KV de xac dinh NVKT")

    has_team_column = "DOIVT" in df.columns
    has_service_column = "TEN_DVVT" in df.columns
    has_kieuld_column = "TEN_KIEULD" in df.columns
    unit_name = ""
    if "TTVT" in df.columns:
        non_empty_units = (
            df["TTVT"].dropna().astype(str).str.strip().replace("", pd.NA).dropna().unique()
        )
        if len(non_empty_units):
            unit_name = str(non_empty_units[0]).strip()

    working_df = df.copy()
    working_df["NGAY_LHD"] = pd.to_datetime(
        working_df["NGAY_LHD"], errors="coerce", dayfirst=True
    )
    working_df["NGAY_HC"] = pd.to_datetime(
        working_df["NGAY_HC"], errors="coerce", dayfirst=True
    )
    working_df["TG_THI_CONG"] = (
        working_df["NGAY_HC"] - working_df["NGAY_LHD"]
    ).dt.total_seconds() / 3600
    working_df["KET_QUA"] = working_df["TG_THI_CONG"].apply(
        lambda value: "Không đạt" if pd.notna(value) and value > 24 else "Đạt"
    )
    working_df["NVKT"] = working_df[nvkt_source_column].apply(_extract_nvkt_name)

    filtered_df = working_df[
        working_df["NVKT"].notna() & working_df["TG_THI_CONG"].notna()
    ].copy()

    if has_team_column:
        filtered_df["DOIVT"] = (
            filtered_df["DOIVT"].astype(str).str.strip().replace("nan", "")
        )

    summary_outputs = _empty_c15_detail_outputs()
    if filtered_df.empty:
        for sheet_name, empty_df in summary_outputs.items():
            append_or_replace_sheet(processed_path, sheet_name, empty_df)
        return processed_path

    group_columns = ["DOIVT", "NVKT"] if has_team_column else ["NVKT"]
    detail_summary = _summarize_c15_detail(filtered_df, group_columns, {})
    sort_columns = ["DOIVT", "NVKT"] if has_team_column else ["NVKT"]
    detail_summary = detail_summary.sort_values(sort_columns, kind="stable").reset_index(drop=True)
    if not has_team_column:
        detail_summary.insert(0, "DOIVT", "")
    summary_outputs["KQ_C15_chitiet"] = detail_summary[
        [
            "DOIVT",
            "NVKT",
            "Phiếu đạt",
            "Phiếu không đạt",
            "Tổng Hoàn công",
            "Tỉ lệ đạt (%)",
        ]
    ]

    team_summary = _summarize_c15_detail(
        filtered_df,
        ["DOIVT"] if has_team_column else ["NVKT"],
        {},
    )
    if has_team_column:
        team_summary = team_summary.sort_values(["DOIVT"], kind="stable").reset_index(drop=True)
        total_row_name = unit_name or "Toàn đơn vị"
    else:
        team_summary = pd.DataFrame(
            [
                {
                    "DOIVT": unit_name or "Tất cả",
                    "Phiếu đạt": int((filtered_df["KET_QUA"] == "Đạt").sum()),
                    "Phiếu không đạt": int((filtered_df["KET_QUA"] == "Không đạt").sum()),
                    "Tổng Hoàn công": int(len(filtered_df)),
                    "Tỉ lệ đạt (%)": round(
                        (filtered_df["KET_QUA"] == "Đạt").sum() / len(filtered_df) * 100,
                        2,
                    ),
                }
            ]
        )
        total_row_name = unit_name or "Tất cả"

    total_row = pd.DataFrame(
        [
            {
                "DOIVT": total_row_name,
                "Phiếu đạt": int((filtered_df["KET_QUA"] == "Đạt").sum()),
                "Phiếu không đạt": int((filtered_df["KET_QUA"] == "Không đạt").sum()),
                "Tổng Hoàn công": int(len(filtered_df)),
                "Tỉ lệ đạt (%)": round(
                    (filtered_df["KET_QUA"] == "Đạt").sum() / len(filtered_df) * 100,
                    2,
                ),
            }
        ]
    )
    if has_team_column and (
        team_summary.empty or str(team_summary.iloc[-1]["DOIVT"]).strip() != total_row_name
    ):
        team_summary = pd.concat([team_summary, total_row], ignore_index=True)
    summary_outputs["TH_TTVTST"] = team_summary[
        ["DOIVT", "Phiếu đạt", "Phiếu không đạt", "Tổng Hoàn công", "Tỉ lệ đạt (%)"]
    ]

    detail_columns = ["DOIVT", "MA_TB", "NGAY_LHD", "NGAY_HC", "TG_THI_CONG", "KET_QUA", "NVKT"]
    available_detail_columns = [column for column in detail_columns if column in filtered_df.columns]
    detail_sheet = filtered_df[available_detail_columns].copy()
    for column in detail_columns:
        if column not in detail_sheet.columns:
            detail_sheet[column] = pd.NA
    if "TG_THI_CONG" in detail_sheet.columns:
        detail_sheet["TG_THI_CONG"] = detail_sheet["TG_THI_CONG"].round(2)
    summary_outputs["Chi_tiet_TG"] = detail_sheet[
        ["DOIVT", "MA_TB", "NGAY_LHD", "NGAY_HC", "TG_THI_CONG", "KET_QUA", "NVKT"]
    ]

    if has_kieuld_column:
        kieuld_df = filtered_df[filtered_df["TEN_KIEULD"].notna()].copy()
        if not kieuld_df.empty:
            kieuld_summary = _summarize_c15_detail(
                kieuld_df,
                (["DOIVT"] if has_team_column else [])
                + ["NVKT", "TEN_KIEULD"],
                {},
            )
            sort_columns = (["DOIVT"] if has_team_column else []) + ["NVKT", "TEN_KIEULD"]
            kieuld_summary = kieuld_summary.sort_values(sort_columns, kind="stable").reset_index(drop=True)
            if not has_team_column:
                kieuld_summary.insert(0, "DOIVT", "")
            summary_outputs["TH_KIEULD"] = kieuld_summary[
                [
                    "DOIVT",
                    "NVKT",
                    "TEN_KIEULD",
                    "Phiếu đạt",
                    "Phiếu không đạt",
                    "Tổng Hoàn công",
                    "Tỉ lệ đạt (%)",
                ]
            ]

    if has_service_column:
        dvvt_df = filtered_df[filtered_df["TEN_DVVT"].notna()].copy()
        if not dvvt_df.empty:
            dvvt_summary = _summarize_c15_detail(
                dvvt_df,
                (["DOIVT"] if has_team_column else []) + ["NVKT", "TEN_DVVT"],
                {},
            )
            sort_columns = (["DOIVT"] if has_team_column else []) + ["NVKT", "TEN_DVVT"]
            dvvt_summary = dvvt_summary.sort_values(sort_columns, kind="stable").reset_index(drop=True)
            if not has_team_column:
                dvvt_summary.insert(0, "DOIVT", "")
            summary_outputs["TH_DVVT"] = dvvt_summary[
                [
                    "DOIVT",
                    "NVKT",
                    "TEN_DVVT",
                    "Phiếu đạt",
                    "Phiếu không đạt",
                    "Tổng Hoàn công",
                    "Tỉ lệ đạt (%)",
                ]
            ]

            if has_team_column:
                dvvt_team_summary = _summarize_c15_detail(
                    dvvt_df,
                    ["DOIVT", "TEN_DVVT"],
                    {},
                )
                dvvt_team_summary = dvvt_team_summary.sort_values(
                    ["DOIVT", "TEN_DVVT"], kind="stable"
                ).reset_index(drop=True)
                summary_outputs["TH_DVVT_DOI"] = dvvt_team_summary[
                    [
                        "DOIVT",
                        "TEN_DVVT",
                        "Phiếu đạt",
                        "Phiếu không đạt",
                        "Tổng Hoàn công",
                        "Tỉ lệ đạt (%)",
                    ]
                ]

            dvvt_unit_summary = _summarize_c15_detail(dvvt_df, ["TEN_DVVT"], {})
            dvvt_unit_summary = dvvt_unit_summary.sort_values(["TEN_DVVT"], kind="stable").reset_index(drop=True)
            total_dvvt_row = pd.DataFrame(
                [
                    {
                        "TEN_DVVT": f"{unit_name} (Tổng)" if unit_name else "Tổng",
                        "Phiếu đạt": int((dvvt_df["KET_QUA"] == "Đạt").sum()),
                        "Phiếu không đạt": int((dvvt_df["KET_QUA"] == "Không đạt").sum()),
                        "Tổng Hoàn công": int(len(dvvt_df)),
                        "Tỉ lệ đạt (%)": round(
                            (dvvt_df["KET_QUA"] == "Đạt").sum() / len(dvvt_df) * 100,
                            2,
                        ),
                    }
                ]
            )
            dvvt_unit_summary = pd.concat([dvvt_unit_summary, total_dvvt_row], ignore_index=True)
            summary_outputs["TH_DVVT_TTVT"] = dvvt_unit_summary[
                [
                    "TEN_DVVT",
                    "Phiếu đạt",
                    "Phiếu không đạt",
                    "Tổng Hoàn công",
                    "Tỉ lệ đạt (%)",
                ]
            ]

    for sheet_name in DEFAULT_C15_DETAIL_SHEETS:
        append_or_replace_sheet(processed_path, sheet_name, summary_outputs[sheet_name])

    return processed_path


def process_c14_chitiet_report_api_output(
    input_path=DEFAULT_C14_DETAIL_INPUT,
    overwrite_processed=False,
    sheet_name=DEFAULT_C14_DETAIL_SHEET,
    exclude_team_patterns=None,
    exclude_person_patterns=None,
):
    """Xu ly bao cao chi tiet C1.4 va tong hop hai long theo to/NVKT."""
    raw_path = _resolve_path(input_path)
    df = pd.read_excel(raw_path)

    required_columns = ["TEN_KV", "DO_HL", "KHL_KT"]
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        raise ValueError(
            "Bao cao C1.4 chi tiet thieu cac cot can thiet: "
            + ", ".join(missing_columns)
        )

    has_doivt = "DOIVT" in df.columns
    df = df.copy()
    df["NVKT"] = df["TEN_KV"].apply(_extract_nvkt_name)

    if has_doivt and exclude_team_patterns:
        pattern = "|".join(
            re.escape(str(item)) for item in exclude_team_patterns if str(item).strip()
        )
        if pattern:
            df = df[
                ~df["DOIVT"].astype(str).str.contains(pattern, na=False)
            ].reset_index(drop=True)

    if exclude_person_patterns:
        pattern = "|".join(
            re.escape(str(item)) for item in exclude_person_patterns if str(item).strip()
        )
        if pattern:
            df = df[
                ~df["NVKT"].astype(str).str.contains(pattern, na=False)
            ].reset_index(drop=True)

    success_mask = df["DO_HL"].isin(["HL", "KHL"])
    khl_mask = df["KHL_KT"].notna()
    hl_mask = success_mask & df["KHL_KT"].isna()

    if has_doivt:
        group_columns = ["DOIVT", "NVKT"]
    else:
        group_columns = ["NVKT"]

    grouped = (
        df.assign(
            _ks_thanh_cong=success_mask.astype(int),
            _phieu_khl=khl_mask.astype(int),
            _phieu_hl=hl_mask.astype(int),
        )
        .dropna(subset=["NVKT"])
        .groupby(group_columns, dropna=False, as_index=False)
        .agg(
            {
                "_ks_thanh_cong": "sum",
                "_phieu_khl": "sum",
                "_phieu_hl": "sum",
            }
        )
    )

    grouped = grouped.rename(
        columns={
            "_ks_thanh_cong": "Tổng phiếu KS thành công",
            "_phieu_khl": "Tổng phiếu KHL",
        }
    )
    grouped["Tỉ lệ HL NVKT (%)"] = grouped.apply(
        lambda row: round(
            row["_phieu_hl"] / row["Tổng phiếu KS thành công"] * 100, 2
        )
        if row["Tổng phiếu KS thành công"] > 0
        else 100,
        axis=1,
    )
    grouped = grouped.drop(columns=["_phieu_hl"])

    sort_columns = ["DOIVT", "NVKT"] if has_doivt else ["NVKT"]
    grouped = grouped.sort_values(sort_columns, kind="stable").reset_index(drop=True)

    processed_path = ensure_processed_workbook(raw_path, overwrite=overwrite_processed)
    append_or_replace_sheet(processed_path, sheet_name, grouped)
    return processed_path


def process_c11_chitiet_report_api_output(
    input_path=DEFAULT_C11_DETAIL_INPUT,
    overwrite_processed=False,
    exclude_team_patterns=None,
    exclude_person_patterns=None,
):
    """Xu ly bao cao chi tiet C1.1 va tao cac sheet tong hop theo KPI khung gio."""
    raw_path = _resolve_path(input_path)
    df = pd.read_excel(raw_path)

    required_columns = ["TEN_KV", "DAT_TT_KO_HEN"]
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        raise ValueError(
            "Bao cao C1.1 chi tiet thieu cac cot can thiet: "
            + ", ".join(missing_columns)
        )

    has_team_column = "TEN_DOI" in df.columns

    df = df.copy()
    df["NVKT"] = df["TEN_KV"].apply(_extract_nvkt_name)
    df = df[df["NVKT"].notna()].copy()

    if has_team_column and exclude_team_patterns:
        pattern = "|".join(
            re.escape(str(item)) for item in exclude_team_patterns if str(item).strip()
        )
        if pattern:
            df = df[
                ~df["TEN_DOI"].astype(str).str.contains(pattern, na=False)
            ].reset_index(drop=True)

    if exclude_person_patterns:
        pattern = "|".join(
            re.escape(str(item)) for item in exclude_person_patterns if str(item).strip()
        )
        if pattern:
            df = df[
                ~df["NVKT"].astype(str).str.contains(pattern, na=False)
            ].reset_index(drop=True)

    processed_path = ensure_processed_workbook(raw_path, overwrite=overwrite_processed)

    detail_summary = _build_c11_detail_summary(df, has_team_column)
    append_or_replace_sheet(processed_path, "chi_tiet", detail_summary)

    hour_sheet_map = {
        15: "chi_tieu_ko_hen_15h",
        16: "chi_tieu_ko_hen_16h",
        17: "chi_tieu_ko_hen_17h",
        18: "chi_tieu_ko_hen_18h",
    }

    for max_hour, sheet_name in hour_sheet_map.items():
        filtered_df = _filter_c11_detail_by_gio(df, max_hour)
        summary_df = _build_c11_detail_summary(filtered_df, has_team_column)
        append_or_replace_sheet(processed_path, sheet_name, summary_df)

        failed_df = filtered_df[
            pd.to_numeric(filtered_df["DAT_TT_KO_HEN"], errors="coerce").fillna(0) != 1
        ].copy()
        if "gio_giao" in failed_df.columns:
            failed_df = failed_df.drop(columns=["gio_giao"])
        append_or_replace_sheet(
            processed_path,
            f"chi_tiet_khong_dat_{max_hour}h",
            failed_df,
        )

    return processed_path


def process_c12_chitiet_sm1_report_api_output(
    input_path=DEFAULT_C12_DETAIL_SM1_INPUT,
    overwrite_processed=False,
    sheet_name="TH_phieu_hong_lai_7_ngay",
    exclude_team_patterns=None,
    exclude_person_patterns=None,
):
    """Xu ly file C1.2 chi tiet SM1 va tao sheet tong hop phieu hong lai 7 ngay."""
    raw_path = _resolve_path(input_path)
    df = pd.read_excel(raw_path)

    required_columns = ["TEN_KV", "MA_TB"]
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        raise ValueError(
            "Bao cao C1.2 chi tiet SM1 thieu cac cot can thiet: "
            + ", ".join(missing_columns)
        )

    has_team_column = "TEN_DOI" in df.columns
    df = df.copy()
    df["NVKT"] = df["TEN_KV"].apply(_extract_nvkt_name)

    if has_team_column and exclude_team_patterns:
        pattern = "|".join(
            re.escape(str(item)) for item in exclude_team_patterns if str(item).strip()
        )
        if pattern:
            df = df[
                ~df["TEN_DOI"].astype(str).str.contains(pattern, na=False)
            ].reset_index(drop=True)

    if exclude_person_patterns:
        pattern = "|".join(
            re.escape(str(item)) for item in exclude_person_patterns if str(item).strip()
        )
        if pattern:
            df = df[
                ~df["NVKT"].astype(str).str.contains(pattern, na=False)
            ].reset_index(drop=True)

    summary_df = _build_unique_ma_tb_summary(df, "Số phiếu HLL", has_team_column)
    processed_path = ensure_processed_workbook(raw_path, overwrite=overwrite_processed)
    append_or_replace_sheet(processed_path, sheet_name, summary_df)
    return processed_path, summary_df


def process_c12_chitiet_sm2_report_api_output(
    input_path=DEFAULT_C12_DETAIL_SM2_INPUT,
    overwrite_processed=False,
    sheet_name="Tong_phieu_bao_hong_thang",
    exclude_team_patterns=None,
    exclude_person_patterns=None,
):
    """Xu ly file C1.2 chi tiet SM2 va tao sheet tong hop tong phieu bao hong thang."""
    raw_path = _resolve_path(input_path)
    df = pd.read_excel(raw_path)

    required_columns = ["TEN_KV", "MA_TB"]
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        raise ValueError(
            "Bao cao C1.2 chi tiet SM2 thieu cac cot can thiet: "
            + ", ".join(missing_columns)
        )

    has_team_column = "TEN_DOI" in df.columns
    df = df.copy()
    df["NVKT"] = df["TEN_KV"].apply(_extract_nvkt_name)

    if has_team_column and exclude_team_patterns:
        pattern = "|".join(
            re.escape(str(item)) for item in exclude_team_patterns if str(item).strip()
        )
        if pattern:
            df = df[
                ~df["TEN_DOI"].astype(str).str.contains(pattern, na=False)
            ].reset_index(drop=True)

    if exclude_person_patterns:
        pattern = "|".join(
            re.escape(str(item)) for item in exclude_person_patterns if str(item).strip()
        )
        if pattern:
            df = df[
                ~df["NVKT"].astype(str).str.contains(pattern, na=False)
            ].reset_index(drop=True)

    summary_df = _build_unique_ma_tb_summary(df, "Số phiếu báo hỏng", has_team_column)
    processed_path = ensure_processed_workbook(raw_path, overwrite=overwrite_processed)
    append_or_replace_sheet(processed_path, sheet_name, summary_df)
    return processed_path, summary_df


def process_c12_chitiet_reports_api_output(
    sm1_input_path=DEFAULT_C12_DETAIL_SM1_INPUT,
    sm2_input_path=DEFAULT_C12_DETAIL_SM2_INPUT,
    overwrite_processed=False,
    exclude_team_patterns=None,
    exclude_person_patterns=None,
):
    """Xu ly tron bo C1.2 chi tiet SM1 + SM2 va tao sheet tong hop cuoi tren file SM1."""
    sm1_processed_path, sm1_summary_df = process_c12_chitiet_sm1_report_api_output(
        input_path=sm1_input_path,
        overwrite_processed=overwrite_processed,
        exclude_team_patterns=exclude_team_patterns,
        exclude_person_patterns=exclude_person_patterns,
    )
    sm2_processed_path, sm2_summary_df = process_c12_chitiet_sm2_report_api_output(
        input_path=sm2_input_path,
        overwrite_processed=overwrite_processed,
        exclude_team_patterns=exclude_team_patterns,
        exclude_person_patterns=exclude_person_patterns,
    )

    sm1_has_team = "TEN_DOI" in sm1_summary_df.columns
    sm2_has_team = "TEN_DOI" in sm2_summary_df.columns

    if sm1_has_team and sm2_has_team:
        merged_df = pd.merge(
            sm1_summary_df[["TEN_DOI", "NVKT", "Số phiếu HLL"]],
            sm2_summary_df[["TEN_DOI", "NVKT", "Số phiếu báo hỏng"]],
            on=["TEN_DOI", "NVKT"],
            how="outer",
        )
    elif sm1_has_team:
        merged_df = pd.merge(
            sm1_summary_df[["TEN_DOI", "NVKT", "Số phiếu HLL"]],
            sm2_summary_df[["NVKT", "Số phiếu báo hỏng"]],
            on="NVKT",
            how="outer",
        )
    elif sm2_has_team:
        merged_df = pd.merge(
            sm1_summary_df[["NVKT", "Số phiếu HLL"]],
            sm2_summary_df[["TEN_DOI", "NVKT", "Số phiếu báo hỏng"]],
            on="NVKT",
            how="outer",
        )
    else:
        merged_df = pd.merge(
            sm1_summary_df[["NVKT", "Số phiếu HLL"]],
            sm2_summary_df[["NVKT", "Số phiếu báo hỏng"]],
            on="NVKT",
            how="outer",
        )

    merged_df["Số phiếu HLL"] = pd.to_numeric(
        merged_df["Số phiếu HLL"], errors="coerce"
    ).fillna(0)
    merged_df["Số phiếu báo hỏng"] = pd.to_numeric(
        merged_df["Số phiếu báo hỏng"], errors="coerce"
    ).fillna(0)
    merged_df["Tỉ lệ HLL tháng (2.5%)"] = merged_df.apply(
        lambda row: round(row["Số phiếu HLL"] / row["Số phiếu báo hỏng"] * 100, 2)
        if row["Số phiếu báo hỏng"] > 0
        else 0,
        axis=1,
    )

    if "TEN_DOI" in merged_df.columns:
        result_df = merged_df[
            ["TEN_DOI", "NVKT", "Số phiếu HLL", "Số phiếu báo hỏng", "Tỉ lệ HLL tháng (2.5%)"]
        ].copy()
        result_df = result_df.sort_values(["TEN_DOI", "NVKT"], kind="stable").reset_index(drop=True)
    else:
        result_df = merged_df[
            ["NVKT", "Số phiếu HLL", "Số phiếu báo hỏng", "Tỉ lệ HLL tháng (2.5%)"]
        ].copy()
        result_df = result_df.sort_values(["NVKT"], kind="stable").reset_index(drop=True)

    append_or_replace_sheet(sm1_processed_path, "TH_SM1C12_HLL_Thang", result_df)
    return {
        "sm1_processed_path": sm1_processed_path,
        "sm2_processed_path": sm2_processed_path,
        "result_df": result_df,
    }
