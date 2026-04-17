# -*- coding: utf-8 -*-
"""Processors cho nhom luong dich vu trong api_transition."""

from pathlib import Path
import re

import pandas as pd
from openpyxl import Workbook, load_workbook

from api_transition.processors.common import (
    PROCESSED_DIR,
    append_or_replace_sheet,
    copy_raw_to_processed,
    ensure_processed_workbook,
)


DEFAULT_PHIEU_HOAN_CONG_INPUT = (
    Path(__file__).resolve().parent.parent
    / "downloads"
    / "phieu_hoan_cong_dich_vu"
    / "phieu_hoan_cong_dich_vu_chi_tiet.xlsx"
)
DEFAULT_TAM_DUNG_CHI_TIET_INPUT = (
    Path(__file__).resolve().parent.parent
    / "downloads"
    / "tam_dung_khoi_phuc_dich_vu"
    / "tam_dung_khoi_phuc_dich_vu_chi_tiet.xlsx"
)
DEFAULT_TAM_DUNG_TONG_HOP_INPUT = (
    Path(__file__).resolve().parent.parent
    / "downloads"
    / "tam_dung_khoi_phuc_dich_vu"
    / "tam_dung_khoi_phuc_dich_vu_tong_hop.xlsx"
)
DEFAULT_MYTV_NGUNG_PSC_INPUT = (
    Path(__file__).resolve().parent.parent
    / "downloads"
    / "mytv_dich_vu"
    / "mytv_ngung_psc.xlsx"
)
DEFAULT_MYTV_HOAN_CONG_INPUT = (
    Path(__file__).resolve().parent.parent
    / "downloads"
    / "mytv_dich_vu"
    / "mytv_hoan_cong.xlsx"
)
DEFAULT_MYTV_THUC_TANG_OUTPUT = (
    PROCESSED_DIR
    / "mytv_dich_vu"
    / "mytv_thuc_tang_processed.xlsx"
)
DEFAULT_SON_TAY_MYTV_NGUNG_T_MINUS_1_INPUT = (
    Path(__file__).resolve().parent.parent
    / "downloads"
    / "mytv_dich_vu"
    / "ngung_psc_mytv_thang_t-1_sontay.xlsx"
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


def _extract_nvkt_from_nhanvien_congviec(value):
    if pd.isna(value):
        return None

    text = str(value).strip()
    if not text:
        return None

    first_line = text.splitlines()[0].strip()
    if ":" in first_line:
        first_line = first_line.split(":", 1)[0].strip()
    return _normalize_person_name(first_line)


def _coalesce(*values):
    for value in values:
        if value is not None and str(value).strip():
            return value
    return None


def _append_total_row(df, total_col, fixed_values):
    total_row = {col: "" for col in df.columns}
    total_row[total_col] = int(pd.to_numeric(df[total_col], errors="coerce").fillna(0).sum())
    for key, value in fixed_values.items():
        total_row[key] = value
    return pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)


def _processed_target_for_group(raw_path, group_name):
    raw_path = _resolve_path(raw_path)
    return PROCESSED_DIR / group_name / f"{raw_path.stem}_processed{raw_path.suffix}"


def _ensure_processed_workbook_for_group(raw_path, group_name, overwrite=False):
    target_path = _processed_target_for_group(raw_path, group_name)
    return copy_raw_to_processed(raw_path, processed_path=target_path, overwrite=overwrite)


def _ensure_generated_workbook(output_path, overwrite=False):
    output_path = Path(output_path).expanduser()
    if not output_path.is_absolute():
        output_path = (Path.cwd() / output_path).resolve()
    else:
        output_path = output_path.resolve()

    output_path.parent.mkdir(parents=True, exist_ok=True)
    if overwrite or not output_path.exists():
        workbook = Workbook()
        workbook.save(output_path)
    return output_path


def _remove_empty_default_sheet(workbook_path):
    workbook = load_workbook(workbook_path)
    if len(workbook.sheetnames) > 1:
        for sheet_name in ["Sheet", "Sheet1"]:
            if sheet_name in workbook.sheetnames:
                sheet = workbook[sheet_name]
                if sheet.max_row == 1 and sheet.max_column == 1 and sheet["A1"].value is None:
                    workbook.remove(sheet)
                    workbook.save(workbook_path)
                    break
    return workbook_path


def _prepare_phieu_hoan_cong_df(input_path):
    raw_path = _resolve_path(input_path)
    df = pd.read_excel(raw_path).copy()
    df.columns = [str(col).strip() for col in df.columns]

    required_cols = ["DOIVT", "TTVT"]
    missing = [col for col in required_cols if col not in df.columns]
    if missing:
        raise ValueError(f"Bao cao phieu hoan cong thieu cot bat buoc: {missing}")

    nvkt_from_db = df["TEN_NVKT_DB"].apply(_normalize_person_name) if "TEN_NVKT_DB" in df.columns else pd.Series([None] * len(df))
    nvkt_from_ten_kv = df["TEN_KV"].apply(_extract_nvkt_from_ten_kv) if "TEN_KV" in df.columns else pd.Series([None] * len(df))
    nvkt_from_congviec = (
        df["NHANVIEN_CONGVIEC"].apply(_extract_nvkt_from_nhanvien_congviec)
        if "NHANVIEN_CONGVIEC" in df.columns
        else pd.Series([None] * len(df))
    )

    df["NVKT"] = [
        _coalesce(a, b, c)
        for a, b, c in zip(nvkt_from_db, nvkt_from_ten_kv, nvkt_from_congviec)
    ]

    df["DOIVT"] = df["DOIVT"].fillna("(Chưa xác định)").astype(str).str.strip()
    df["TTVT"] = df["TTVT"].fillna("(Chưa xác định)").astype(str).str.strip()
    df["NVKT"] = df["NVKT"].fillna("(Chưa xác định)")

    return raw_path, df


def _build_nvkt_summary(df):
    grouped = (
        df.groupby(["TTVT", "DOIVT", "NVKT"], dropna=False)
        .size()
        .reset_index(name="SỐ LƯỢNG HOÀN CÔNG THÁNG")
        .sort_values(["TTVT", "DOIVT", "SỐ LƯỢNG HOÀN CÔNG THÁNG", "NVKT"], ascending=[True, True, False, True])
        .reset_index(drop=True)
    )
    return _append_total_row(
        grouped,
        "SỐ LƯỢNG HOÀN CÔNG THÁNG",
        {"TTVT": "TỔNG CỘNG", "DOIVT": "", "NVKT": ""},
    )


def _build_team_summary(df):
    grouped = (
        df.groupby(["TTVT", "DOIVT"], dropna=False)
        .size()
        .reset_index(name="SỐ LƯỢNG HOÀN CÔNG THÁNG")
        .sort_values(["TTVT", "SỐ LƯỢNG HOÀN CÔNG THÁNG", "DOIVT"], ascending=[True, False, True])
        .reset_index(drop=True)
    )
    return _append_total_row(
        grouped,
        "SỐ LƯỢNG HOÀN CÔNG THÁNG",
        {"TTVT": "TỔNG CỘNG", "DOIVT": ""},
    )


def _build_ttvt_summary(df):
    grouped = (
        df.groupby(["TTVT"], dropna=False)
        .size()
        .reset_index(name="SỐ LƯỢNG HOÀN CÔNG THÁNG")
        .sort_values(["SỐ LƯỢNG HOÀN CÔNG THÁNG", "TTVT"], ascending=[False, True])
        .reset_index(drop=True)
    )
    return _append_total_row(
        grouped,
        "SỐ LƯỢNG HOÀN CÔNG THÁNG",
        {"TTVT": "TỔNG CỘNG"},
    )


def _build_reason_summary(df, reason_col):
    if reason_col not in df.columns:
        return pd.DataFrame(columns=[reason_col, "Số lượng"])

    reason_df = df.copy()
    reason_df[reason_col] = reason_df[reason_col].fillna("(Không có lý do)").astype(str).str.strip()
    reason_df.loc[reason_df[reason_col] == "", reason_col] = "(Không có lý do)"

    return (
        reason_df.groupby(reason_col)
        .size()
        .reset_index(name="Số lượng")
        .sort_values(["Số lượng", reason_col], ascending=[False, True])
        .reset_index(drop=True)
    )


def _build_group_count_summary(df, group_cols, count_col):
    grouped = (
        df.groupby(group_cols, dropna=False)
        .size()
        .reset_index(name=count_col)
        .sort_values(group_cols[:-1] + [count_col, group_cols[-1]], ascending=[True] * (len(group_cols) - 1) + [False, True])
        .reset_index(drop=True)
    )
    return grouped


def _prepare_mytv_ngung_psc_df(input_path):
    raw_path = _resolve_path(input_path)
    df = pd.read_excel(raw_path).copy()
    df.columns = [str(col).strip() for col in df.columns]

    required_cols = ["TEN_KV", "TEN_DOI"]
    missing = [col for col in required_cols if col not in df.columns]
    if missing:
        raise ValueError(f"Bao cao MyTV ngung PSC thieu cot bat buoc: {missing}")

    df["TEN_KV"] = df["TEN_KV"].apply(_normalize_person_name).fillna("(Chưa xác định)")
    df["TEN_DOI"] = df["TEN_DOI"].fillna("(Chưa xác định)").astype(str).str.strip()
    if "TEN_TTVT" in df.columns:
        df["TEN_TTVT"] = df["TEN_TTVT"].fillna("(Chưa xác định)").astype(str).str.strip()
    return raw_path, df


def _prepare_mytv_hoan_cong_df(input_path):
    raw_path = _resolve_path(input_path)
    df = pd.read_excel(raw_path).copy()
    df.columns = [str(col).strip() for col in df.columns]

    required_cols = ["NHANVIEN_KT", "DOIVT"]
    missing = [col for col in required_cols if col not in df.columns]
    if missing:
        raise ValueError(f"Bao cao MyTV hoan cong thieu cot bat buoc: {missing}")

    df["NHANVIEN_KT"] = df["NHANVIEN_KT"].apply(_normalize_person_name).fillna("(Chưa xác định)")
    df["DOIVT"] = df["DOIVT"].fillna("(Chưa xác định)").astype(str).str.strip()
    if "TEN_TTVT" in df.columns:
        df["TEN_TTVT"] = df["TEN_TTVT"].fillna("(Chưa xác định)").astype(str).str.strip()
    return raw_path, df


def _append_mytv_total_row(df, team_col, person_col, value_col):
    total_row = pd.DataFrame(
        [{team_col: "TỔNG CỘNG", person_col: "", value_col: int(pd.to_numeric(df[value_col], errors="coerce").fillna(0).sum())}]
    )
    return pd.concat([df, total_row], ignore_index=True)


def _append_mytv_team_total_row(df, team_col, value_col):
    total_row = pd.DataFrame(
        [{team_col: "TỔNG CỘNG", value_col: int(pd.to_numeric(df[value_col], errors="coerce").fillna(0).sum())}]
    )
    return pd.concat([df, total_row], ignore_index=True)


def _extract_son_tay_ngung_psc_subset(df_raw):
    data_row_indices = [3, 4, 5, 6]
    if len(df_raw) <= max(data_row_indices):
        raise ValueError(
            f"File Son Tay khong du so dong du lieu. Can it nhat {max(data_row_indices) + 1} dong, nhung chi co {len(df_raw)} dong."
        )

    header_row_count = min(3, len(df_raw))
    header_map = {}
    for col_idx in range(df_raw.shape[1]):
        header_values = [
            str(df_raw.iat[row_idx, col_idx]).strip()
            for row_idx in range(header_row_count)
            if pd.notna(df_raw.iat[row_idx, col_idx]) and str(df_raw.iat[row_idx, col_idx]).strip()
        ]
        for value in header_values:
            header_map.setdefault(value, col_idx)

    required_columns = {
        "Đơn vị/Nhân viên KT": "Đơn vị/Nhân viên KT",
        "Hoàn công(*) (1.5)": "Hoàn công(*) (1.5)",
        "Lũy kế tháng(1.6)": "Lũy kế tháng(1.6)",
        "Lũy kế năm(1.7)": "Lũy kế năm(1.7)",
        "Ngưng PSC tạm tính tháng T(5.1)": "Ngưng PSC tạm tính tháng T(5.1)",
        "TB Ngưng PSC lũy kế năm(4.6) (4.7-4.4)": "TB Ngưng PSC lũy kế năm(4.6) (4.7-4.4)",
    }
    missing_headers = [header for header in required_columns if header not in header_map]
    if missing_headers:
        raise ValueError(
            "Khong tim thay cac cot bat buoc trong bao cao Son Tay: " + ", ".join(missing_headers)
        )

    selected_col_indices = [header_map[header] for header in required_columns]
    df_subset = df_raw.iloc[data_row_indices, selected_col_indices].copy()
    df_subset.columns = list(required_columns.values())
    return df_subset


def process_phieu_hoan_cong_dich_vu_chi_tiet_api_output(
    input_path=DEFAULT_PHIEU_HOAN_CONG_INPUT,
    overwrite_processed=False,
):
    """Xu ly bao cao phieu hoan cong dich vu chi tiet."""
    raw_path, df = _prepare_phieu_hoan_cong_df(input_path)
    df_nvkt = _build_nvkt_summary(df)
    df_team = _build_team_summary(df)
    df_ttvt = _build_ttvt_summary(df)

    processed_path = ensure_processed_workbook(raw_path, overwrite=overwrite_processed)
    append_or_replace_sheet(processed_path, "Data", df)
    append_or_replace_sheet(processed_path, "fiber_hoan_cong_thang", df_nvkt)
    append_or_replace_sheet(processed_path, "fiber_hoan_cong_thang_theo_to", df_team)
    append_or_replace_sheet(processed_path, "fiber_hoan_cong_thang_theo_ttvt", df_ttvt)
    return processed_path


def _prepare_tam_dung_chi_tiet_df(input_path):
    raw_path = _resolve_path(input_path)
    df = pd.read_excel(raw_path).copy()
    df.columns = [str(col).strip() for col in df.columns]

    required_cols = ["DOIVT", "TTVT"]
    missing = [col for col in required_cols if col not in df.columns]
    if missing:
        raise ValueError(f"Bao cao tam dung/khoi phuc thieu cot bat buoc: {missing}")

    nvkt_from_db = df["TEN_NVKT_DB"].apply(_normalize_person_name) if "TEN_NVKT_DB" in df.columns else pd.Series([None] * len(df))
    nvkt_from_ten_kv = df["TEN_KV"].apply(_extract_nvkt_from_ten_kv) if "TEN_KV" in df.columns else pd.Series([None] * len(df))
    nvkt_from_xm = df["NVKT_XM"].apply(_normalize_person_name) if "NVKT_XM" in df.columns else pd.Series([None] * len(df))

    df["NVKT"] = [
        _coalesce(a, b, c)
        for a, b, c in zip(nvkt_from_db, nvkt_from_ten_kv, nvkt_from_xm)
    ]

    df["DOIVT"] = df["DOIVT"].fillna("(Chưa xác định)").astype(str).str.strip()
    df["TTVT"] = df["TTVT"].fillna("(Chưa xác định)").astype(str).str.strip()
    df["NVKT"] = df["NVKT"].fillna("(Chưa xác định)")
    return raw_path, df


def process_tam_dung_khoi_phuc_dich_vu_chi_tiet_api_output(
    input_path=DEFAULT_TAM_DUNG_CHI_TIET_INPUT,
    overwrite_processed=False,
):
    """Xu ly bao cao tam dung/khoi phuc dich vu chi tiet."""
    raw_path, df = _prepare_tam_dung_chi_tiet_df(input_path)
    df_nvkt = _build_nvkt_summary(df).rename(
        columns={"SỐ LƯỢNG HOÀN CÔNG THÁNG": "SỐ LƯỢNG NGƯNG PSC THÁNG"}
    )
    df_team = _build_team_summary(df).rename(
        columns={"SỐ LƯỢNG HOÀN CÔNG THÁNG": "SỐ LƯỢNG NGƯNG PSC THÁNG"}
    )
    df_ttvt = _build_ttvt_summary(df).rename(
        columns={"SỐ LƯỢNG HOÀN CÔNG THÁNG": "SỐ LƯỢNG NGƯNG PSC THÁNG"}
    )
    df_reasons = _build_reason_summary(df, "LYDOHUY")

    processed_path = ensure_processed_workbook(raw_path, overwrite=overwrite_processed)
    append_or_replace_sheet(processed_path, "Data", df)
    append_or_replace_sheet(processed_path, "fiber_ngung_psc_thang", df_nvkt)
    append_or_replace_sheet(processed_path, "fiber_ngung_psc_thang_theo_to", df_team)
    append_or_replace_sheet(processed_path, "fiber_ngung_psc_thang_theo_ttvt", df_ttvt)
    append_or_replace_sheet(processed_path, "tong_hop_ly_do_huy", df_reasons)
    return processed_path


def process_tam_dung_khoi_phuc_dich_vu_tong_hop_api_output(
    input_path=DEFAULT_TAM_DUNG_TONG_HOP_INPUT,
    overwrite_processed=False,
):
    """Xu ly workbook tong hop tam dung/khoi phuc dich vu.

    Hien tai file raw co the chua du lieu. Truong hop do, van tao file processed
    va ghi sheet thong bao de luong sau khong bi gay.
    """
    raw_path = _resolve_path(input_path)
    try:
        df = pd.read_excel(raw_path)
    except ValueError:
        df = pd.DataFrame()

    processed_path = ensure_processed_workbook(raw_path, overwrite=overwrite_processed)
    if df.empty and len(df.columns) == 0:
        note_df = pd.DataFrame(
            [{"Trạng thái": "Không có dữ liệu", "Ghi chú": "Workbook nguồn đang rỗng, chưa có dữ liệu để chuẩn hóa."}]
        )
        append_or_replace_sheet(processed_path, "thong_bao", note_df)
        return processed_path

    df.columns = [str(col).strip() for col in df.columns]
    append_or_replace_sheet(processed_path, "du_lieu_sach", df)
    return processed_path


def process_mytv_ngung_psc_api_output(
    input_path=DEFAULT_MYTV_NGUNG_PSC_INPUT,
    overwrite_processed=False,
):
    """Xu ly bao cao MyTV ngung PSC."""
    raw_path, df = _prepare_mytv_ngung_psc_df(input_path)
    df_nvkt = _build_group_count_summary(df, ["TEN_DOI", "TEN_KV"], "SỐ LƯỢNG NGƯNG PSC THÁNG")
    df_nvkt = _append_mytv_total_row(df_nvkt, "TEN_DOI", "TEN_KV", "SỐ LƯỢNG NGƯNG PSC THÁNG")
    df_team = (
        df.groupby("TEN_DOI", dropna=False)
        .size()
        .reset_index(name="SỐ LƯỢNG NGƯNG PSC THÁNG")
        .sort_values(["SỐ LƯỢNG NGƯNG PSC THÁNG", "TEN_DOI"], ascending=[False, True])
        .reset_index(drop=True)
    )
    df_team = _append_mytv_team_total_row(df_team, "TEN_DOI", "SỐ LƯỢNG NGƯNG PSC THÁNG")

    processed_path = _ensure_processed_workbook_for_group(raw_path, "mytv_dich_vu", overwrite=overwrite_processed)
    append_or_replace_sheet(processed_path, "Data", df)
    append_or_replace_sheet(processed_path, "mytv_ngung_psc_thang", df_nvkt)
    append_or_replace_sheet(processed_path, "mytv_ngung_psc_thang_theo_to", df_team)
    return processed_path


def process_mytv_hoan_cong_api_output(
    input_path=DEFAULT_MYTV_HOAN_CONG_INPUT,
    overwrite_processed=False,
):
    """Xu ly bao cao MyTV hoan cong."""
    raw_path, df = _prepare_mytv_hoan_cong_df(input_path)
    df_nvkt = _build_group_count_summary(df, ["DOIVT", "NHANVIEN_KT"], "SỐ LƯỢNG HOÀN CÔNG THÁNG")
    df_nvkt = _append_mytv_total_row(df_nvkt, "DOIVT", "NHANVIEN_KT", "SỐ LƯỢNG HOÀN CÔNG THÁNG")
    df_team = (
        df.groupby("DOIVT", dropna=False)
        .size()
        .reset_index(name="SỐ LƯỢNG HOÀN CÔNG THÁNG")
        .sort_values(["SỐ LƯỢNG HOÀN CÔNG THÁNG", "DOIVT"], ascending=[False, True])
        .reset_index(drop=True)
    )
    df_team = _append_mytv_team_total_row(df_team, "DOIVT", "SỐ LƯỢNG HOÀN CÔNG THÁNG")

    processed_path = _ensure_processed_workbook_for_group(raw_path, "mytv_dich_vu", overwrite=overwrite_processed)
    append_or_replace_sheet(processed_path, "Data", df)
    append_or_replace_sheet(processed_path, "mytv_hoan_cong_thang", df_nvkt)
    append_or_replace_sheet(processed_path, "mytv_hoan_cong_thang_theo_to", df_team)
    return processed_path


def process_mytv_thuc_tang_api_output(
    ngung_psc_input=DEFAULT_MYTV_NGUNG_PSC_INPUT,
    hoan_cong_input=DEFAULT_MYTV_HOAN_CONG_INPUT,
    output_path=DEFAULT_MYTV_THUC_TANG_OUTPUT,
    overwrite_processed=False,
):
    """Tao bao cao MyTV thuc tang tu 2 nguon raw: hoan cong va ngung PSC."""
    _, df_ngung = _prepare_mytv_ngung_psc_df(ngung_psc_input)
    _, df_hc = _prepare_mytv_hoan_cong_df(hoan_cong_input)

    df_ngung_team = (
        df_ngung.groupby("TEN_DOI", dropna=False)
        .size()
        .reset_index(name="Ngưng phát sinh cước")
        .rename(columns={"TEN_DOI": "Đội VT"})
    )
    df_hc_team = (
        df_hc.groupby("DOIVT", dropna=False)
        .size()
        .reset_index(name="Hoàn công")
        .rename(columns={"DOIVT": "Đội VT"})
    )
    df_thuc_tang_to = pd.merge(df_hc_team, df_ngung_team, on="Đội VT", how="outer").fillna(0)
    df_thuc_tang_to["Hoàn công"] = pd.to_numeric(df_thuc_tang_to["Hoàn công"], errors="coerce").fillna(0).astype(int)
    df_thuc_tang_to["Ngưng phát sinh cước"] = pd.to_numeric(df_thuc_tang_to["Ngưng phát sinh cước"], errors="coerce").fillna(0).astype(int)
    df_thuc_tang_to["Thực tăng"] = df_thuc_tang_to["Hoàn công"] - df_thuc_tang_to["Ngưng phát sinh cước"]
    df_thuc_tang_to["Tỷ lệ ngưng/psc"] = df_thuc_tang_to.apply(
        lambda row: round((row["Ngưng phát sinh cước"] / row["Hoàn công"]) * 100, 2) if row["Hoàn công"] else 0.0,
        axis=1,
    )
    df_thuc_tang_to = df_thuc_tang_to.sort_values(["Thực tăng", "Đội VT"], ascending=[False, True]).reset_index(drop=True)
    total_row_to = pd.DataFrame(
        [{
            "Đội VT": "TỔNG CỘNG",
            "Hoàn công": int(df_thuc_tang_to["Hoàn công"].sum()),
            "Ngưng phát sinh cước": int(df_thuc_tang_to["Ngưng phát sinh cước"].sum()),
            "Thực tăng": int(df_thuc_tang_to["Thực tăng"].sum()),
            "Tỷ lệ ngưng/psc": round(
                (df_thuc_tang_to["Ngưng phát sinh cước"].sum() / df_thuc_tang_to["Hoàn công"].sum()) * 100,
                2,
            ) if df_thuc_tang_to["Hoàn công"].sum() else 0.0,
        }]
    )
    df_thuc_tang_to = pd.concat([df_thuc_tang_to, total_row_to], ignore_index=True)

    df_ngung_nvkt = (
        df_ngung.groupby(["TEN_DOI", "TEN_KV"], dropna=False)
        .size()
        .reset_index(name="Ngưng phát sinh cước")
        .rename(columns={"TEN_DOI": "Đội VT", "TEN_KV": "NVKT"})
    )
    df_hc_nvkt = (
        df_hc.groupby(["DOIVT", "NHANVIEN_KT"], dropna=False)
        .size()
        .reset_index(name="Hoàn công")
        .rename(columns={"DOIVT": "Đội VT", "NHANVIEN_KT": "NVKT"})
    )
    df_thuc_tang_nvkt = pd.merge(df_hc_nvkt, df_ngung_nvkt, on=["Đội VT", "NVKT"], how="outer").fillna(0)
    df_thuc_tang_nvkt["Hoàn công"] = pd.to_numeric(df_thuc_tang_nvkt["Hoàn công"], errors="coerce").fillna(0).astype(int)
    df_thuc_tang_nvkt["Ngưng phát sinh cước"] = pd.to_numeric(df_thuc_tang_nvkt["Ngưng phát sinh cước"], errors="coerce").fillna(0).astype(int)
    df_thuc_tang_nvkt["Thực tăng"] = df_thuc_tang_nvkt["Hoàn công"] - df_thuc_tang_nvkt["Ngưng phát sinh cước"]
    df_thuc_tang_nvkt["Tỷ lệ ngưng/psc"] = df_thuc_tang_nvkt.apply(
        lambda row: round((row["Ngưng phát sinh cước"] / row["Hoàn công"]) * 100, 2) if row["Hoàn công"] else 0.0,
        axis=1,
    )
    df_thuc_tang_nvkt = df_thuc_tang_nvkt.sort_values(["Thực tăng", "Đội VT", "NVKT"], ascending=[False, True, True]).reset_index(drop=True)
    total_row_nvkt = pd.DataFrame(
        [{
            "Đội VT": "TỔNG CỘNG",
            "NVKT": "",
            "Hoàn công": int(df_thuc_tang_nvkt["Hoàn công"].sum()),
            "Ngưng phát sinh cước": int(df_thuc_tang_nvkt["Ngưng phát sinh cước"].sum()),
            "Thực tăng": int(df_thuc_tang_nvkt["Thực tăng"].sum()),
            "Tỷ lệ ngưng/psc": round(
                (df_thuc_tang_nvkt["Ngưng phát sinh cước"].sum() / df_thuc_tang_nvkt["Hoàn công"].sum()) * 100,
                2,
            ) if df_thuc_tang_nvkt["Hoàn công"].sum() else 0.0,
        }]
    )
    df_thuc_tang_nvkt = pd.concat([df_thuc_tang_nvkt, total_row_nvkt], ignore_index=True)

    processed_path = _ensure_generated_workbook(output_path, overwrite=overwrite_processed)
    append_or_replace_sheet(processed_path, "thuc_tang_theo_to", df_thuc_tang_to)
    append_or_replace_sheet(processed_path, "thuc_tang_theo_NVKT", df_thuc_tang_nvkt)
    _remove_empty_default_sheet(processed_path)
    return processed_path


def process_son_tay_mytv_ngung_psc_t_minus_1_api_output(
    input_path=DEFAULT_SON_TAY_MYTV_NGUNG_T_MINUS_1_INPUT,
    overwrite_processed=False,
):
    """Xu ly bao cao MyTV Son Tay ngung PSC thang T-1."""
    raw_path = _resolve_path(input_path)
    df_raw = pd.read_excel(raw_path, header=None)
    df_subset = _extract_son_tay_ngung_psc_subset(df_raw)

    numeric_cols = list(df_subset.columns[1:])
    for col in numeric_cols:
        df_subset[col] = pd.to_numeric(df_subset[col], errors="coerce").fillna(0).astype(int)

    total_row = pd.DataFrame(
        [{
            "Đơn vị/Nhân viên KT": "Tổng",
            "Hoàn công(*) (1.5)": int(df_subset["Hoàn công(*) (1.5)"].sum()),
            "Lũy kế tháng(1.6)": int(df_subset["Lũy kế tháng(1.6)"].sum()),
            "Lũy kế năm(1.7)": int(df_subset["Lũy kế năm(1.7)"].sum()),
            "Ngưng PSC tạm tính tháng T(5.1)": int(df_subset["Ngưng PSC tạm tính tháng T(5.1)"].sum()),
            "TB Ngưng PSC lũy kế năm(4.6) (4.7-4.4)": int(df_subset["TB Ngưng PSC lũy kế năm(4.6) (4.7-4.4)"].sum()),
        }]
    )
    df_final = pd.concat([df_subset, total_row], ignore_index=True)

    processed_path = _ensure_processed_workbook_for_group(raw_path, "mytv_dich_vu", overwrite=overwrite_processed)
    append_or_replace_sheet(processed_path, "TH_ngung_PSC-Thang T-1", df_final)
    return processed_path
