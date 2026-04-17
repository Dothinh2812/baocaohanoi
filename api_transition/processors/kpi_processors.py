# -*- coding: utf-8 -*-
"""Processors cho nhom KPI NVKT trong api_transition."""

from pathlib import Path
import re

import pandas as pd

from api_transition.processors.common import (
    DOWNLOADS_DIR,
    append_or_replace_sheet,
    ensure_processed_workbook,
)


DEFAULT_DSNV_FILE = Path.cwd() / "dsnv.xlsx"
DEFAULT_KPI_C11_INPUT = DOWNLOADS_DIR / "kpi_nvkt" / "c11-nvktdb report.xlsx"
DEFAULT_KPI_C12_INPUT = DOWNLOADS_DIR / "kpi_nvkt" / "c12-nvktdb report.xlsx"
DEFAULT_KPI_C13_INPUT = DOWNLOADS_DIR / "kpi_nvkt" / "c13-nvktdb report.xlsx"


def _resolve_path(input_path):
    path = Path(input_path).expanduser()
    if not path.is_absolute():
        path = (Path.cwd() / path).resolve()
    else:
        path = path.resolve()

    if not path.exists():
        raise FileNotFoundError(f"Khong tim thay file input: {path}")
    return path


def _normalize_person_name(name):
    if pd.isna(name):
        return None

    name = str(name).strip()
    if not name:
        return None

    name = re.sub(r"\s+", " ", name)
    return name.lower().title()


def _extract_nvkt_from_unit_cell(raw_value):
    """Boc ten NVKT tu chuoi dang ma-ma-ten."""
    if pd.isna(raw_value):
        return None

    raw_value = str(raw_value).strip()
    if not raw_value:
        return None

    parts = raw_value.split("-")
    candidate = parts[-1].strip() if parts else raw_value
    if "(" in candidate:
        candidate = candidate.split("(")[0].strip()
    return _normalize_person_name(candidate)


def _build_donvi_lookup(dsnv_file):
    dsnv_path = _resolve_path(dsnv_file)
    df_dsnv = pd.read_excel(dsnv_path)
    df_dsnv.columns = [str(col).strip() for col in df_dsnv.columns]

    if "Họ tên" not in df_dsnv.columns or "đơn vị" not in df_dsnv.columns:
        raise ValueError(
            "File dsnv.xlsx phai co cac cot 'Họ tên' va 'đơn vị'"
        )

    normalized_names = df_dsnv["Họ tên"].apply(_normalize_person_name)
    donvi_values = df_dsnv["đơn vị"].astype(str).str.strip()
    return dict(zip(normalized_names, donvi_values))


def _process_kpi_nvkt_report(
    input_path,
    dsnv_file,
    output_columns,
    output_slice,
    sheet_name,
    overwrite_processed=False,
):
    raw_path = _resolve_path(input_path)
    lookup = _build_donvi_lookup(dsnv_file)

    df_raw = pd.read_excel(raw_path, header=[0, 1])
    col_donvi_raw = df_raw.iloc[:, 0]
    nvkt_series = col_donvi_raw.apply(_extract_nvkt_from_unit_cell)

    mask = nvkt_series.notna() & nvkt_series.astype(str).str.strip().ne("")
    nvkt_series = nvkt_series[mask].reset_index(drop=True)

    df_metrics = df_raw.iloc[mask.values, output_slice].copy().reset_index(drop=True)
    df_metrics.columns = output_columns

    donvi_series = nvkt_series.map(lookup).fillna("")

    df_result = pd.DataFrame()
    df_result.insert(0, "STT", range(1, len(nvkt_series) + 1))
    df_result.insert(1, "đơn vị", donvi_series.values)
    df_result.insert(2, "NVKT", nvkt_series.values)
    for index, col in enumerate(output_columns, start=3):
        df_result.insert(index, col, df_metrics[col].values)

    df_result = df_result.sort_values(
        by=["đơn vị", "NVKT"], na_position="last", kind="stable"
    ).reset_index(drop=True)
    df_result["STT"] = range(1, len(df_result) + 1)

    processed_path = ensure_processed_workbook(raw_path, overwrite=overwrite_processed)
    append_or_replace_sheet(processed_path, sheet_name, df_result)
    return processed_path


def process_kpi_nvkt_c11_api_output(
    input_path=DEFAULT_KPI_C11_INPUT,
    dsnv_file=DEFAULT_DSNV_FILE,
    overwrite_processed=False,
    sheet_name="c11 kpi nvkt",
):
    """Xu ly KPI NVKT C11 va ghi sheet ket qua vao file processed."""
    output_columns = [
        "SM1",
        "SM2",
        "Tỷ lệ sửa chữa phiếu chất lượng chủ động dịch vụ FiberVNN, MyTV đạt yêu cầu",
        "SM3",
        "SM4",
        "Tỷ lệ phiếu sửa chữa báo hỏng dịch vụ BRCĐ đúng quy định không tính hẹn",
        "Chỉ tiêu BSC",
    ]
    return _process_kpi_nvkt_report(
        input_path=input_path,
        dsnv_file=dsnv_file,
        output_columns=output_columns,
        output_slice=slice(1, 8),
        sheet_name=sheet_name,
        overwrite_processed=overwrite_processed,
    )


def process_kpi_nvkt_c12_api_output(
    input_path=DEFAULT_KPI_C12_INPUT,
    dsnv_file=DEFAULT_DSNV_FILE,
    overwrite_processed=False,
    sheet_name="c12 kpi nvkt",
):
    """Xu ly KPI NVKT C12 va ghi sheet ket qua vao file processed."""
    output_columns = [
        "SM1",
        "SM2",
        "Tỷ lệ thuê bao báo hỏng dịch vụ BRCĐ lặp lại",
        "SM3",
        "SM4",
        "Tỷ lệ sự cố dịch vụ BRCĐ",
        "Chỉ tiêu BSC",
    ]
    return _process_kpi_nvkt_report(
        input_path=input_path,
        dsnv_file=dsnv_file,
        output_columns=output_columns,
        output_slice=slice(1, 8),
        sheet_name=sheet_name,
        overwrite_processed=overwrite_processed,
    )


def process_kpi_nvkt_c13_api_output(
    input_path=DEFAULT_KPI_C13_INPUT,
    dsnv_file=DEFAULT_DSNV_FILE,
    overwrite_processed=False,
    sheet_name="c13 kpi nvkt",
):
    """Xu ly KPI NVKT C13 va ghi sheet ket qua vao file processed."""
    output_columns = [
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
    return _process_kpi_nvkt_report(
        input_path=input_path,
        dsnv_file=dsnv_file,
        output_columns=output_columns,
        output_slice=slice(1, 11),
        sheet_name=sheet_name,
        overwrite_processed=overwrite_processed,
    )
