# -*- coding: utf-8 -*-
"""Helper chung cho luong xu ly file trong api_transition."""

from pathlib import Path
import shutil

import pandas as pd


API_TRANSITION_DIR = Path(__file__).resolve().parent.parent
DEFAULT_DOWNLOADS_DIR = API_TRANSITION_DIR / "downloads"
DEFAULT_PROCESSED_DIR = API_TRANSITION_DIR / "Processed"
DOWNLOADS_DIR = DEFAULT_DOWNLOADS_DIR
PROCESSED_DIR = DEFAULT_PROCESSED_DIR


def get_downloads_dir():
    """Tra ve downloads root dang duoc cau hinh cho processor runtime."""
    return DOWNLOADS_DIR


def get_processed_dir():
    """Tra ve Processed root dang duoc cau hinh cho processor runtime."""
    return PROCESSED_DIR


def processed_group_dir(group_name):
    """Tra ve thu muc Processed/<group> theo runtime hien tai."""
    return get_processed_dir() / group_name


def configure_runtime_roots(*, downloads_root=None, processed_root=None):
    """Cap nhat root downloads/Processed cho 1 instance runtime cu the."""
    global DOWNLOADS_DIR, PROCESSED_DIR

    DOWNLOADS_DIR = (
        Path(downloads_root).expanduser().resolve()
        if downloads_root is not None
        else DEFAULT_DOWNLOADS_DIR
    )
    PROCESSED_DIR = (
        Path(processed_root).expanduser().resolve()
        if processed_root is not None
        else DEFAULT_PROCESSED_DIR
    )


def reset_runtime_roots():
    """Tra lai root downloads/Processed mac dinh cua codebase."""
    configure_runtime_roots()


def _normalize_input_path(input_path):
    """Chuyen input_path ve Path tuyet doi va validate ton tai."""
    path = Path(input_path).expanduser()
    if not path.is_absolute():
        path = (Path.cwd() / path).resolve()
    else:
        path = path.resolve()

    if not path.exists():
        raise FileNotFoundError(f"Khong tim thay file input: {path}")
    return path


def _processed_filename_for(path):
    """Them hau to _processed truoc extension."""
    return f"{path.stem}_processed{path.suffix}"


def build_processed_path(input_path):
    """Sinh duong dan file processed tu 1 file raw nam trong api_transition/downloads."""
    source_path = _normalize_input_path(input_path)

    try:
        relative_path = source_path.relative_to(get_downloads_dir().resolve())
    except ValueError as exc:
        raise ValueError(
            f"Input phai nam trong thu muc downloads: {source_path}"
        ) from exc

    target_dir = get_processed_dir() / relative_path.parent
    return target_dir / _processed_filename_for(relative_path)


def copy_raw_to_processed(input_path, processed_path=None, overwrite=False):
    """Copy file raw sang Processed va tra ve path dich."""
    source_path = _normalize_input_path(input_path)
    target_path = Path(processed_path) if processed_path else build_processed_path(source_path)
    target_path = target_path.expanduser()
    if not target_path.is_absolute():
        target_path = (Path.cwd() / target_path).resolve()
    else:
        target_path = target_path.resolve()

    target_path.parent.mkdir(parents=True, exist_ok=True)
    if overwrite or not target_path.exists():
        shutil.copy2(source_path, target_path)
    return target_path


def ensure_processed_workbook(input_path, overwrite=False):
    """Tao ban copy processed neu chua co, hoac ghi de neu overwrite=True."""
    return copy_raw_to_processed(input_path, overwrite=overwrite)


def append_or_replace_sheet(workbook_path, sheet_name, df):
    """Ghi DataFrame vao 1 sheet trong workbook processed."""
    workbook_path = _normalize_input_path(workbook_path)
    if not isinstance(df, pd.DataFrame):
        raise TypeError("df phai la pandas.DataFrame")

    with pd.ExcelWriter(
        workbook_path,
        engine="openpyxl",
        mode="a",
        if_sheet_exists="replace",
    ) as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)

    return workbook_path
