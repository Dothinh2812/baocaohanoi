# -*- coding: utf-8 -*-
"""Processors cho nhom GHTT trong api_transition."""

from pathlib import Path
import re

import pandas as pd

from api_transition.processors.common import append_or_replace_sheet, ensure_processed_workbook


DEFAULT_DSNV_FILE = Path.cwd() / "dsnv.xlsx"
DEFAULT_GHTT_HNI_INPUT = Path(__file__).resolve().parent.parent / "downloads" / "ghtt" / "ghtt_hni report.xlsx"
DEFAULT_GHTT_SONTAY_INPUT = Path(__file__).resolve().parent.parent / "downloads" / "ghtt" / "ghtt_sontay report.xlsx"
DEFAULT_GHTT_NVKT_INPUT = Path(__file__).resolve().parent.parent / "downloads" / "ghtt" / "ghtt_nvktdb report.xlsx"


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


def _build_donvi_lookup(dsnv_file):
    dsnv_path = _resolve_path(dsnv_file)
    df_dsnv = pd.read_excel(dsnv_path)
    df_dsnv.columns = [str(col).strip() for col in df_dsnv.columns]
    if "Họ tên" not in df_dsnv.columns or "đơn vị" not in df_dsnv.columns:
        raise ValueError("File dsnv.xlsx phai co cot 'Họ tên' va 'đơn vị'")
    return dict(
        zip(
            df_dsnv["Họ tên"].apply(_normalize_person_name),
            df_dsnv["đơn vị"].astype(str).str.strip(),
        )
    )


def _format_percent_value(value):
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if not text:
        return ""
    if text.endswith("%"):
        return text
    try:
        return f"{float(str(text).replace(',', '.')):.2f}%"
    except Exception:
        return text


def _prepare_ghtt_summary_df(input_path, *, keep_ttvt=True):
    raw_path = _resolve_path(input_path)
    df_raw = pd.read_excel(raw_path, header=None, skiprows=2, usecols=range(12), dtype=object)
    df_raw = df_raw.dropna(how="all").reset_index(drop=True)
    if df_raw.shape[1] < 12:
        raise ValueError(f"Bao cao GHTT khong du cot de xu ly: {raw_path}")

    if keep_ttvt:
        df = df_raw.iloc[:, [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11]].copy()
        df.columns = [
            "Đơn vị",
            "TTVT",
            "Hoàn thành T",
            "Giao NVKT T",
            "Tỷ lệ T",
            "Hoàn thành T+1",
            "Giao NVKT T+1",
            "Tỷ lệ T+1",
            "SL GHTT >=6T",
            "Hoàn thành >=6T T+1",
            "Tỷ lệ >=6T T+1",
            "Tỷ lệ Tổng",
        ]
    else:
        df = df_raw.iloc[:, [0, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11]].copy()
        df.columns = [
            "Đơn vị",
            "Hoàn thành T",
            "Giao NVKT T",
            "Tỷ lệ T",
            "Hoàn thành T+1",
            "Giao NVKT T+1",
            "Tỷ lệ T+1",
            "SL GHTT >=6T",
            "Hoàn thành >=6T T+1",
            "Tỷ lệ >=6T T+1",
            "Tỷ lệ Tổng",
        ]

    for col in ["Tỷ lệ T", "Tỷ lệ T+1", "Tỷ lệ >=6T T+1", "Tỷ lệ Tổng"]:
        df[col] = df[col].apply(_format_percent_value)

    numeric_cols = ["Hoàn thành T", "Giao NVKT T", "Hoàn thành T+1", "Giao NVKT T+1", "SL GHTT >=6T", "Hoàn thành >=6T T+1"]
    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)

    return raw_path, df


def process_ghtt_hni_api_output(
    input_path=DEFAULT_GHTT_HNI_INPUT,
    overwrite_processed=False,
    sheet_name="kq_hni",
):
    """Xu ly bao cao GHTT HNI va ghi ket qua vao file processed."""
    raw_path, df = _prepare_ghtt_summary_df(input_path, keep_ttvt=False)
    processed_path = ensure_processed_workbook(raw_path, overwrite=overwrite_processed)
    append_or_replace_sheet(processed_path, sheet_name, df)
    return processed_path


def process_ghtt_sontay_api_output(
    input_path=DEFAULT_GHTT_SONTAY_INPUT,
    overwrite_processed=False,
    sheet_name="kq_sontay",
):
    """Xu ly bao cao GHTT Son Tay va ghi ket qua vao file processed."""
    raw_path, df = _prepare_ghtt_summary_df(input_path, keep_ttvt=True)
    processed_path = ensure_processed_workbook(raw_path, overwrite=overwrite_processed)
    append_or_replace_sheet(processed_path, sheet_name, df)
    return processed_path


def process_ghtt_nvktdb_api_output(
    input_path=DEFAULT_GHTT_NVKT_INPUT,
    dsnv_file=DEFAULT_DSNV_FILE,
    overwrite_processed=False,
    sheet_name="kq_nvktdb",
):
    """Xu ly bao cao GHTT NVKTDB va bo sung cot don vi tu dsnv.xlsx."""
    raw_path, df = _prepare_ghtt_summary_df(input_path, keep_ttvt=True)

    df = df.rename(columns={"Đơn vị": "NVKT raw"})
    df["NVKT"] = df["NVKT raw"].apply(
        lambda x: _normalize_person_name(
            str(x).split("-", 1)[1].strip() if isinstance(x, str) and "-" in x else x
        )
    )

    lookup = _build_donvi_lookup(dsnv_file)
    df["Đơn vị"] = df["NVKT"].map(lookup)

    df = df[
        [
            "NVKT",
            "Đơn vị",
            "TTVT",
            "Hoàn thành T",
            "Giao NVKT T",
            "Tỷ lệ T",
            "Hoàn thành T+1",
            "Giao NVKT T+1",
            "Tỷ lệ T+1",
            "SL GHTT >=6T",
            "Hoàn thành >=6T T+1",
            "Tỷ lệ >=6T T+1",
            "Tỷ lệ Tổng",
        ]
    ].copy()

    cols = df.columns.tolist()
    cols.remove("Đơn vị")
    nvkt_idx = cols.index("NVKT")
    cols.insert(nvkt_idx + 1, "Đơn vị")
    df = df[cols].reset_index(drop=True)

    processed_path = ensure_processed_workbook(raw_path, overwrite=overwrite_processed)
    append_or_replace_sheet(processed_path, sheet_name, df)
    return processed_path
