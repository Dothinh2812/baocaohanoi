# -*- coding: utf-8 -*-
"""Processors cho bao cao I1.5 / I1.5 K2 trong api_transition."""

from __future__ import annotations

import sqlite3
from pathlib import Path
import re
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import pandas as pd

from api_transition.processors.common import (
    DOWNLOADS_DIR,
    append_or_replace_sheet,
    ensure_processed_workbook,
)


API_TRANSITION_DIR = Path(__file__).resolve().parent.parent
DEFAULT_I15_INPUT = DOWNLOADS_DIR / "chi_tieu_i" / "i1.5 report.xlsx"
DEFAULT_I15_K2_INPUT = DOWNLOADS_DIR / "chi_tieu_i" / "i1.5_k2 report.xlsx"
DEFAULT_HISTORY_DB_PATH = API_TRANSITION_DIR / "report_history.db"
DEFAULT_DSNV_DB_PATH = API_TRANSITION_DIR.parent / "danhba.db"
I15_HISTORY_COLUMNS = (
    "k_suffix",
    "ngay_bao_cao",
    "account_cts",
    "ten_tb_one",
    "dt_onediachi_one",
    "doi_one",
    "nvkt_db",
    "nvkt_db_normalized",
    "sa",
    "olt_cts",
    "port_cts",
    "thietbi",
    "ketcuoi",
    "trangthai_tb",
    "olt_rx",
    "onu_rx",
)


def _resolve_path(input_path: str | Path) -> Path:
    path = Path(input_path).expanduser()
    if not path.is_absolute():
        path = (Path.cwd() / path).resolve()
    else:
        path = path.resolve()
    if not path.exists():
        raise FileNotFoundError(f"Khong tim thay file input: {path}")
    return path


def _normalize_text(value: Any) -> Optional[str]:
    if pd.isna(value):
        return None
    text = str(value).strip()
    if not text:
        return None
    return text


def normalize_nvkt(value: Any) -> Optional[str]:
    """Chuan hoa NVKT_DB ve phan ten sau dau '-'. """
    text = _normalize_text(value)
    if text is None:
        return None
    if "-" in text:
        text = text.split("-")[-1].strip()
    text = re.sub(r"\([^)]*\)", "", text).strip()
    text = re.sub(r"\s+", " ", text)
    return text or None


def add_tt_column(df: pd.DataFrame) -> pd.DataFrame:
    """Them cot TT vao dau DataFrame."""
    if df.empty:
        return df.copy()
    result = df.copy()
    if "TT" in result.columns:
        result = result.drop(columns=["TT"])
    result.insert(0, "TT", range(1, len(result) + 1))
    return result


def _safe_numeric(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce")


def _read_danhba_tables(dsnv_db_path: Path) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    if not dsnv_db_path.exists():
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    try:
        conn = sqlite3.connect(dsnv_db_path)
        try:
            df_danhba = pd.read_sql_query(
                "SELECT MA_TB, THIETBI, SA, KETCUOI, DOI_VT, NVKT FROM danhba",
                conn,
            )
        except Exception:
            df_danhba = pd.DataFrame()
        try:
            df_thong_ke = pd.read_sql_query(
                "SELECT DOI_VT, NVKT, so_thue_bao_pon_qly FROM thong_ke",
                conn,
            )
        except Exception:
            df_thong_ke = pd.DataFrame()
        try:
            df_thong_ke_dv = pd.read_sql_query(
                "SELECT don_vi, so_thue_bao_pon_qly FROM thong_ke_theo_don_vi",
                conn,
            )
        except Exception:
            df_thong_ke_dv = pd.DataFrame()
        conn.close()
        return df_danhba, df_thong_ke, df_thong_ke_dv
    except Exception:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()


def _derive_report_date(df: pd.DataFrame) -> str:
    for column in ("NGAY_SUYHAO", "NGAY_SH"):
        if column in df.columns:
            parsed = pd.to_datetime(df[column], dayfirst=True, errors="coerce")
            parsed = parsed.dropna()
            if not parsed.empty:
                return parsed.iloc[0].strftime("%Y-%m-%d")
    return pd.Timestamp.today().strftime("%Y-%m-%d")


def _ensure_history_schema(conn: sqlite3.Connection) -> None:
    conn.executescript(
        """
        PRAGMA foreign_keys = ON;

        CREATE TABLE IF NOT EXISTS i15_snapshots (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            k_suffix TEXT NOT NULL,
            ngay_bao_cao DATE NOT NULL,
            account_cts TEXT NOT NULL,
            ten_tb_one TEXT,
            dt_onediachi_one TEXT,
            doi_one TEXT,
            nvkt_db TEXT,
            nvkt_db_normalized TEXT,
            sa TEXT,
            olt_cts TEXT,
            port_cts TEXT,
            thietbi TEXT,
            ketcuoi TEXT,
            trangthai_tb TEXT,
            olt_rx REAL,
            onu_rx REAL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            UNIQUE (k_suffix, ngay_bao_cao, account_cts)
        );

        CREATE TABLE IF NOT EXISTS i15_tracking (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            k_suffix TEXT NOT NULL,
            account_cts TEXT NOT NULL,
            ngay_xuat_hien_dau_tien DATE NOT NULL,
            ngay_thay_cuoi_cung DATE NOT NULL,
            so_ngay_lien_tuc INTEGER DEFAULT 1,
            doi_one TEXT,
            nvkt_db TEXT,
            nvkt_db_normalized TEXT,
            sa TEXT,
            trang_thai TEXT,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            UNIQUE (k_suffix, account_cts)
        );

        CREATE TABLE IF NOT EXISTS i15_daily_changes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            k_suffix TEXT NOT NULL,
            ngay_bao_cao DATE NOT NULL,
            account_cts TEXT NOT NULL,
            loai_bien_dong TEXT NOT NULL,
            doi_one TEXT,
            nvkt_db TEXT,
            nvkt_db_normalized TEXT,
            sa TEXT,
            so_ngay_lien_tuc INTEGER,
            ten_tb_one TEXT,
            dt_onediachi_one TEXT,
            olt_cts TEXT,
            port_cts TEXT,
            thietbi TEXT,
            ketcuoi TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            UNIQUE (k_suffix, ngay_bao_cao, account_cts, loai_bien_dong)
        );

        CREATE TABLE IF NOT EXISTS i15_daily_summary (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            k_suffix TEXT NOT NULL,
            ngay_bao_cao DATE NOT NULL,
            doi_one TEXT,
            nvkt_db_normalized TEXT,
            tong_so_hien_tai INTEGER DEFAULT 0,
            so_tang_moi INTEGER DEFAULT 0,
            so_giam_het INTEGER DEFAULT 0,
            so_van_con INTEGER DEFAULT 0,
            so_tb_quan_ly INTEGER DEFAULT 0,
            ty_le_shc REAL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            UNIQUE (k_suffix, ngay_bao_cao, doi_one, nvkt_db_normalized)
        );

        CREATE INDEX IF NOT EXISTS idx_i15_snapshots_variant_date ON i15_snapshots(k_suffix, ngay_bao_cao);
        CREATE INDEX IF NOT EXISTS idx_i15_snapshots_account ON i15_snapshots(k_suffix, account_cts);
        CREATE INDEX IF NOT EXISTS idx_i15_tracking_account ON i15_tracking(k_suffix, account_cts);
        CREATE INDEX IF NOT EXISTS idx_i15_daily_changes_variant_date ON i15_daily_changes(k_suffix, ngay_bao_cao);
        CREATE INDEX IF NOT EXISTS idx_i15_daily_summary_variant_date ON i15_daily_summary(k_suffix, ngay_bao_cao);
        """
    )


def _history_date(report_date: str, offset_days: int) -> str:
    return (pd.Timestamp(report_date) + pd.Timedelta(days=offset_days)).strftime("%Y-%m-%d")


def _load_previous_snapshot(
    conn: sqlite3.Connection,
    k_suffix: str,
    report_date: str,
) -> pd.DataFrame:
    prev_date = _history_date(report_date, -1)
    try:
        return pd.read_sql_query(
            """
            SELECT *
            FROM i15_snapshots
            WHERE k_suffix = ? AND ngay_bao_cao = ?
            """,
            conn,
            params=(k_suffix, prev_date),
        )
    except Exception:
        return pd.DataFrame()


def _load_tracking(conn: sqlite3.Connection, k_suffix: str) -> pd.DataFrame:
    try:
        return pd.read_sql_query(
            """
            SELECT *
            FROM i15_tracking
            WHERE k_suffix = ?
            """,
            conn,
            params=(k_suffix,),
        )
    except Exception:
        return pd.DataFrame()


def _delete_history_rows(conn: sqlite3.Connection, k_suffix: str, report_date: str) -> None:
    conn.execute("DELETE FROM i15_snapshots WHERE k_suffix = ? AND ngay_bao_cao = ?", (k_suffix, report_date))
    conn.execute("DELETE FROM i15_daily_changes WHERE k_suffix = ? AND ngay_bao_cao = ?", (k_suffix, report_date))
    conn.execute("DELETE FROM i15_daily_summary WHERE k_suffix = ? AND ngay_bao_cao = ?", (k_suffix, report_date))


def _upsert_history(
    conn: sqlite3.Connection,
    k_suffix: str,
    report_date: str,
    df: pd.DataFrame,
    account_col: str,
    prev_snapshot: pd.DataFrame,
    df_thong_ke: pd.DataFrame,
) -> None:
    _delete_history_rows(conn, k_suffix, report_date)

    prev_accounts = set(
        prev_snapshot[account_col].dropna().astype(str).str.strip().tolist()
    ) if not prev_snapshot.empty and account_col in prev_snapshot.columns else set()
    today_accounts = set(df[account_col].dropna().astype(str).str.strip().tolist())

    tang_moi = today_accounts - prev_accounts
    giam_het = prev_accounts - today_accounts
    van_con = today_accounts & prev_accounts

    tracking_df = _load_tracking(conn, k_suffix)
    tracking_map: Dict[str, Dict[str, Any]] = {}
    for _, row in tracking_df.iterrows():
        tracking_map[str(row["account_cts"])]= row.to_dict()

    # Snapshot hiện tại
    for _, row in df.iterrows():
        account = _normalize_text(row.get(account_col))
        if not account:
            continue
        conn.execute(
            """
            INSERT OR REPLACE INTO i15_snapshots (
                k_suffix, ngay_bao_cao, account_cts, ten_tb_one, dt_onediachi_one,
                doi_one, nvkt_db, nvkt_db_normalized, sa, olt_cts, port_cts,
                thietbi, ketcuoi, trangthai_tb, olt_rx, onu_rx
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                k_suffix,
                report_date,
                account,
                _normalize_text(row.get("TEN_TB_ONE")),
                _normalize_text(row.get("DT_ONEDIACHI_ONE") or row.get("DT_ONE")),
                _normalize_text(row.get("DOI_ONE")),
                _normalize_text(row.get("NVKT_DB")),
                _normalize_text(row.get("NVKT_DB_NORMALIZED")),
                _normalize_text(row.get("SA")),
                _normalize_text(row.get("OLT_CTS") or row.get("OLT_RX")),
                _normalize_text(row.get("PORT_CTS")),
                _normalize_text(row.get("THIETBI")),
                _normalize_text(row.get("KETCUOI")),
                _normalize_text(row.get("TRANGTHAI_TB")),
                row.get("OLT_RX"),
                row.get("ONU_RX"),
            ),
        )

    # Tracking cập nhật cho hôm nay
    def upsert_tracking(account: str, row: pd.Series, status: str, so_ngay: int) -> None:
        existing = tracking_map.get(account)
        first_date = existing.get("ngay_xuat_hien_dau_tien") if existing else report_date
        if not first_date:
            first_date = report_date
        conn.execute(
            """
            INSERT OR REPLACE INTO i15_tracking (
                k_suffix, account_cts, ngay_xuat_hien_dau_tien, ngay_thay_cuoi_cung,
                so_ngay_lien_tuc, doi_one, nvkt_db, nvkt_db_normalized, sa, trang_thai
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                k_suffix,
                account,
                first_date,
                report_date,
                so_ngay,
                _normalize_text(row.get("DOI_ONE")),
                _normalize_text(row.get("NVKT_DB")),
                _normalize_text(row.get("NVKT_DB_NORMALIZED")),
                _normalize_text(row.get("SA")),
                status,
            ),
        )

    for account in tang_moi:
        row = df[df[account_col].astype(str).str.strip() == account].iloc[0]
        upsert_tracking(account, row, "DANG_SUY_HAO", 1)

    for account in van_con:
        row = df[df[account_col].astype(str).str.strip() == account].iloc[0]
        existing = tracking_map.get(account)
        so_ngay = int(existing.get("so_ngay_lien_tuc", 1)) + 1 if existing else 2
        upsert_tracking(account, row, "DANG_SUY_HAO", so_ngay)

    for account in giam_het:
        if not prev_snapshot.empty and account_col in prev_snapshot.columns:
            prev_row = prev_snapshot[prev_snapshot[account_col].astype(str).str.strip() == account]
            if prev_row.empty:
                continue
            row = prev_row.iloc[0]
        else:
            continue
        existing = tracking_map.get(account)
        so_ngay = int(existing.get("so_ngay_lien_tuc", 1)) if existing else 1
        upsert_tracking(account, row, "DA_HET_SUY_HAO", so_ngay)

    def save_changes(records: pd.DataFrame, loai: str, source_df: pd.DataFrame, source_col: str) -> None:
        if records.empty:
            return
        for account in records[source_col].dropna().astype(str).str.strip().unique():
            if not account:
                continue
            if source_df.empty:
                continue
            match = source_df[source_df[source_col].astype(str).str.strip() == account]
            if match.empty:
                continue
            row = match.iloc[0]
            existing = tracking_map.get(account, {})
            conn.execute(
                """
                INSERT OR REPLACE INTO i15_daily_changes (
                    k_suffix, ngay_bao_cao, account_cts, loai_bien_dong,
                    doi_one, nvkt_db, nvkt_db_normalized, sa, so_ngay_lien_tuc,
                    ten_tb_one, dt_onediachi_one, olt_cts, port_cts, thietbi, ketcuoi
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    k_suffix,
                    report_date,
                    account,
                    loai,
                    _normalize_text(row.get("DOI_ONE")),
                    _normalize_text(row.get("NVKT_DB")),
                    _normalize_text(row.get("NVKT_DB_NORMALIZED")),
                    _normalize_text(row.get("SA")),
                    int(existing.get("so_ngay_lien_tuc", 1)) if existing else 1,
                    _normalize_text(row.get("TEN_TB_ONE")),
                    _normalize_text(row.get("DT_ONEDIACHI_ONE") or row.get("DT_ONE")),
                    _normalize_text(row.get("OLT_CTS") or row.get("OLT_RX")),
                    _normalize_text(row.get("PORT_CTS")),
                    _normalize_text(row.get("THIETBI")),
                    _normalize_text(row.get("KETCUOI")),
                ),
            )

    save_changes(df[df[account_col].astype(str).str.strip().isin(tang_moi)], "TANG_MOI", df, account_col)
    if not prev_snapshot.empty and account_col in prev_snapshot.columns:
        save_changes(prev_snapshot[prev_snapshot[account_col].astype(str).str.strip().isin(giam_het)], "GIAM_HET", prev_snapshot, account_col)
    save_changes(df[df[account_col].astype(str).str.strip().isin(van_con)], "VAN_CON", df, account_col)

    # Tổng hợp lịch sử theo đơn vị và NVKT
    group_cols = ["DOI_ONE", "NVKT_DB_NORMALIZED"]
    curr_summary = (
        df.assign(
            _account=df[account_col].astype(str).str.strip(),
            _is_tang=df[account_col].astype(str).str.strip().isin(tang_moi).astype(int),
            _is_van=df[account_col].astype(str).str.strip().isin(van_con).astype(int),
        )
        .groupby(group_cols, dropna=False, as_index=False)
        .agg(
            tong_so_hien_tai=("_account", "size"),
            so_tang_moi=("_is_tang", "sum"),
            so_van_con=("_is_van", "sum"),
        )
    )
    curr_summary["so_giam_het"] = 0
    prev_summary = pd.DataFrame()
    if not prev_snapshot.empty and account_col in prev_snapshot.columns:
        prev_summary = (
            prev_snapshot.assign(
                _account=prev_snapshot[account_col].astype(str).str.strip(),
            )
            .groupby(group_cols, dropna=False, as_index=False)
            .agg(so_giam_het=("_account", "size"))
        )
    if not prev_summary.empty:
        curr_summary = curr_summary.merge(prev_summary, on=group_cols, how="left", suffixes=("", "_prev"))
        curr_summary["so_giam_het"] = curr_summary["so_giam_het_prev"].fillna(0).astype(int)
        curr_summary = curr_summary.drop(columns=["so_giam_het_prev"])

    if not curr_summary.empty:
        curr_summary["doi_one"] = curr_summary["DOI_ONE"]
        curr_summary["nvkt_db_normalized"] = curr_summary["NVKT_DB_NORMALIZED"]
        curr_summary["k_suffix"] = k_suffix
        curr_summary["ngay_bao_cao"] = report_date
        curr_summary["so_tb_quan_ly"] = 0
        curr_summary["ty_le_shc"] = 0.0

        if not df_thong_ke.empty:
            thong_ke = df_thong_ke.rename(
                columns={
                    "DOI_VT": "DOI_ONE",
                    "NVKT": "NVKT_DB_NORMALIZED",
                    "so_thue_bao_pon_qly": "so_tb_quan_ly",
                }
            )
            curr_summary = curr_summary.merge(
                thong_ke[["DOI_ONE", "NVKT_DB_NORMALIZED", "so_tb_quan_ly"]],
                on=["DOI_ONE", "NVKT_DB_NORMALIZED"],
                how="left",
                suffixes=("", "_ref"),
            )
            if "so_tb_quan_ly_ref" in curr_summary.columns:
                curr_summary["so_tb_quan_ly"] = curr_summary["so_tb_quan_ly_ref"].fillna(curr_summary["so_tb_quan_ly"])
                curr_summary = curr_summary.drop(columns=["so_tb_quan_ly_ref"])

        curr_summary["so_tb_quan_ly"] = pd.to_numeric(curr_summary["so_tb_quan_ly"], errors="coerce").fillna(0).astype(int)
        curr_summary["ty_le_shc"] = curr_summary.apply(
            lambda row: round(row["tong_so_hien_tai"] / row["so_tb_quan_ly"] * 100, 2) if row["so_tb_quan_ly"] else 0.0,
            axis=1,
        )

    conn.execute(
        "DELETE FROM i15_daily_summary WHERE k_suffix = ? AND ngay_bao_cao = ?",
        (k_suffix, report_date),
    )
    for _, row in curr_summary.iterrows():
        conn.execute(
            """
            INSERT INTO i15_daily_summary (
                k_suffix, ngay_bao_cao, doi_one, nvkt_db_normalized,
                tong_so_hien_tai, so_tang_moi, so_giam_het, so_van_con,
                so_tb_quan_ly, ty_le_shc
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                k_suffix,
                report_date,
                _normalize_text(row.get("DOI_ONE")),
                _normalize_text(row.get("NVKT_DB_NORMALIZED")),
                int(row.get("tong_so_hien_tai", 0) or 0),
                int(row.get("so_tang_moi", 0) or 0),
                int(row.get("so_giam_het", 0) or 0),
                int(row.get("so_van_con", 0) or 0),
                int(row.get("so_tb_quan_ly", 0) or 0),
                float(row.get("ty_le_shc", 0) or 0),
            ),
        )


def _build_today_summary(
    df: pd.DataFrame,
    k_suffix: str,
    report_date: str,
    df_thong_ke: pd.DataFrame,
    df_thong_ke_dv: pd.DataFrame,
    account_col: str,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    df_work = df.copy()
    df_work["NVKT_DB_NORMALIZED"] = df_work["NVKT_DB"].apply(normalize_nvkt)
    if "DOI_ONE" not in df_work.columns and "DOI_CTS" in df_work.columns:
        df_work["DOI_ONE"] = df_work["DOI_CTS"]
    if "TTVT_ONE" not in df_work.columns and "TTVT_CTS" in df_work.columns:
        df_work["TTVT_ONE"] = df_work["TTVT_CTS"]

    result = df_work.groupby(["NVKT_DB_NORMALIZED", "DOI_ONE"], dropna=False).size().reset_index(name=f"Số TB Suy hao cao {k_suffix}")
    result = result.rename(columns={"NVKT_DB_NORMALIZED": "NVKT_DB", "DOI_ONE": "Đơn vị"})
    result = result[["Đơn vị", "NVKT_DB", f"Số TB Suy hao cao {k_suffix}"]]
    result = result.sort_values(by="Đơn vị", kind="stable").reset_index(drop=True)

    if not df_thong_ke.empty:
        thong_ke = df_thong_ke.rename(columns={
            "DOI_VT": "Đơn vị",
            "NVKT": "NVKT_DB",
            "so_thue_bao_pon_qly": "Số TB quản lý",
        })
        result = result.merge(thong_ke, on=["Đơn vị", "NVKT_DB"], how="left")
        result["Tỉ lệ SHC (%)"] = (
            result[f"Số TB Suy hao cao {k_suffix}"] / result["Số TB quản lý"] * 100
        ).round(2)
        result["Tỉ lệ SHC (%)"] = result["Tỉ lệ SHC (%)"].fillna(0)

    by_to = result.groupby("Đơn vị", dropna=False)[f"Số TB Suy hao cao {k_suffix}"].sum().reset_index()
    by_to = by_to.sort_values(by="Đơn vị", kind="stable").reset_index(drop=True)
    if not df_thong_ke_dv.empty:
        by_to = by_to.merge(
            df_thong_ke_dv.rename(columns={"don_vi": "Đơn vị", "so_thue_bao_pon_qly": "Số TB quản lý"}),
            on="Đơn vị",
            how="left",
        )
        by_to["Tỉ lệ SHC (%)"] = (
            by_to[f"Số TB Suy hao cao {k_suffix}"] / by_to["Số TB quản lý"] * 100
        ).round(2)
        by_to["Tỉ lệ SHC (%)"] = by_to["Tỉ lệ SHC (%)"].fillna(0)

    total_shc = by_to[f"Số TB Suy hao cao {k_suffix}"].sum() if not by_to.empty else 0
    total_ql = by_to["Số TB quản lý"].sum() if "Số TB quản lý" in by_to.columns else 0
    total_rate = round(total_shc / total_ql * 100, 2) if total_ql else 0
    total_row = pd.DataFrame(
        [{"Đơn vị": "Tổng", f"Số TB Suy hao cao {k_suffix}": total_shc, **({"Số TB quản lý": total_ql} if total_ql else {}), **({"Tỉ lệ SHC (%)": total_rate} if total_ql else {})}]
    )
    by_to = pd.concat([by_to, total_row], ignore_index=True)

    sa_df = pd.DataFrame()
    if "SA" in df_work.columns:
        sa_df = df_work.groupby("SA", dropna=False).size().reset_index(name="Số lượng")
        sa_df = sa_df.sort_values(by="Số lượng", ascending=False, kind="stable").reset_index(drop=True)
        sa_df = pd.concat([sa_df, pd.DataFrame([{"SA": "Tổng", "Số lượng": int(sa_df["Số lượng"].sum())}])], ignore_index=True)

    return result, by_to, sa_df, df_work


def _build_delta_frames(
    df_work: pd.DataFrame,
    prev_snapshot: pd.DataFrame,
    account_col: str,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    curr_accounts = set(df_work[account_col].dropna().astype(str).str.strip())
    prev_accounts = set()
    if not prev_snapshot.empty and account_col in prev_snapshot.columns:
        prev_accounts = set(prev_snapshot[account_col].dropna().astype(str).str.strip())

    tang_moi = curr_accounts - prev_accounts
    giam_het = prev_accounts - curr_accounts
    van_con = curr_accounts & prev_accounts

    df_tang = df_work[df_work[account_col].astype(str).str.strip().isin(tang_moi)].copy()
    df_van = df_work[df_work[account_col].astype(str).str.strip().isin(van_con)].copy()

    if not prev_snapshot.empty and account_col in prev_snapshot.columns:
        df_giam = prev_snapshot[prev_snapshot[account_col].astype(str).str.strip().isin(giam_het)].copy()
    else:
        df_giam = pd.DataFrame(columns=df_work.columns)

    return df_tang, df_giam, df_van, pd.DataFrame()


def _build_group_change_summary(
    df_work: pd.DataFrame,
    prev_snapshot: pd.DataFrame,
    account_col: str,
) -> pd.DataFrame:
    curr_accounts = df_work[account_col].astype(str).str.strip()
    prev_accounts = prev_snapshot[account_col].astype(str).str.strip() if (not prev_snapshot.empty and account_col in prev_snapshot.columns) else pd.Series(dtype=str)

    tang_moi = set(curr_accounts.dropna()) - set(prev_accounts.dropna())
    giam_het = set(prev_accounts.dropna()) - set(curr_accounts.dropna())
    van_con = set(curr_accounts.dropna()) & set(prev_accounts.dropna())

    def _group(df_source: pd.DataFrame, accounts: set[str], label: str) -> pd.DataFrame:
        if df_source.empty or not accounts:
            return pd.DataFrame(columns=["Đơn vị", "NVKT_DB", label])
        grouped = (
            df_source[df_source[account_col].astype(str).str.strip().isin(accounts)]
            .groupby(["DOI_ONE", "NVKT_DB_NORMALIZED"], dropna=False)
            .size()
            .reset_index(name=label)
        )
        grouped = grouped.rename(columns={"DOI_ONE": "Đơn vị", "NVKT_DB_NORMALIZED": "NVKT_DB"})
        return grouped

    curr_total = (
        df_work.groupby(["DOI_ONE", "NVKT_DB_NORMALIZED"], dropna=False)
        .size()
        .reset_index(name="Tổng số hiện tại")
        .rename(columns={"DOI_ONE": "Đơn vị", "NVKT_DB_NORMALIZED": "NVKT_DB"})
    )
    current_tang = _group(df_work, tang_moi, "Tăng mới")
    current_van = _group(df_work, van_con, "Vẫn còn")
    prev_giam = _group(prev_snapshot, giam_het, "Giảm/Hết") if not prev_snapshot.empty else pd.DataFrame(columns=["Đơn vị", "NVKT_DB", "Giảm/Hết"])

    summary = curr_total.merge(current_tang, on=["Đơn vị", "NVKT_DB"], how="left")
    summary = summary.merge(current_van, on=["Đơn vị", "NVKT_DB"], how="left")
    summary = summary.merge(prev_giam, on=["Đơn vị", "NVKT_DB"], how="left")
    for column in ("Tăng mới", "Vẫn còn", "Giảm/Hết"):
        if column not in summary.columns:
            summary[column] = 0
        summary[column] = pd.to_numeric(summary[column], errors="coerce").fillna(0).astype(int)

    summary["Số TB quản lý"] = 0
    summary["Tỉ lệ SHC (%)"] = 0.0
    summary = summary[["Đơn vị", "NVKT_DB", "Tổng số hiện tại", "Tăng mới", "Giảm/Hết", "Vẫn còn", "Số TB quản lý", "Tỉ lệ SHC (%)"]]
    return summary.sort_values(["Đơn vị", "NVKT_DB"], kind="stable").reset_index(drop=True)


def _build_detail_sheet(df_work: pd.DataFrame) -> pd.DataFrame:
    columns = [
        "MA_TB",
        "ACCOUNT_CTS",
        "TEN_TB_ONE",
        "DIACHI_ONE",
        "DT_ONE",
        "DT_ONEDIACHI_ONE",
        "NGAY_SUYHAO",
        "THIETBI",
        "SA",
        "KETCUOI",
        "NVKT_DB_NORMALIZED",
        "OLT_RX",
        "ONU_RX",
    ]
    available = [column for column in columns if column in df_work.columns]
    detail = df_work[available].copy()
    return detail


def _write_workbook(
    raw_path: Path,
    overwrite_processed: bool,
    df_work: pd.DataFrame,
    result_df: pd.DataFrame,
    by_to_df: pd.DataFrame,
    sa_df: pd.DataFrame,
    change_df: pd.DataFrame,
    tang_df: pd.DataFrame,
    giam_df: pd.DataFrame,
    van_df: pd.DataFrame,
) -> Path:
    processed_path = ensure_processed_workbook(raw_path, overwrite=overwrite_processed)

    append_or_replace_sheet(processed_path, "Sheet1", add_tt_column(df_work))
    append_or_replace_sheet(processed_path, "TH_SHC_I15", add_tt_column(result_df))
    append_or_replace_sheet(processed_path, "TH_SHC_theo_to", add_tt_column(by_to_df))
    append_or_replace_sheet(processed_path, "shc_theo_SA", add_tt_column(sa_df) if not sa_df.empty else sa_df)
    append_or_replace_sheet(processed_path, "Bien_dong_tong_hop", add_tt_column(change_df) if not change_df.empty else change_df)

    if not tang_df.empty:
        append_or_replace_sheet(processed_path, "Tang_moi", add_tt_column(tang_df))
    else:
        append_or_replace_sheet(processed_path, "Tang_moi", tang_df)
    if not giam_df.empty:
        giam_out = giam_df.copy()
        if "so_ngay_lien_tuc" in giam_out.columns:
            giam_out = giam_out.rename(columns={"so_ngay_lien_tuc": "Số ngày suy hao"})
        append_or_replace_sheet(processed_path, "Giam_het", add_tt_column(giam_out))
    else:
        append_or_replace_sheet(processed_path, "Giam_het", giam_df)
    if not van_df.empty:
        van_out = van_df.copy()
        if "so_ngay_lien_tuc" in van_out.columns:
            van_out = van_out.rename(columns={"so_ngay_lien_tuc": "Số ngày liên tục"})
        append_or_replace_sheet(processed_path, "Van_con", add_tt_column(van_out))
    else:
        append_or_replace_sheet(processed_path, "Van_con", van_df)

    nvkt_col = "NVKT_DB_NORMALIZED"
    if nvkt_col in df_work.columns:
        for nvkt in sorted(df_work[nvkt_col].dropna().astype(str).str.strip().unique()):
            nvkt_df = df_work[df_work[nvkt_col].astype(str).str.strip() == nvkt].copy()
            if "SA" in nvkt_df.columns:
                nvkt_df = nvkt_df.sort_values(by="SA", kind="stable").reset_index(drop=True)
            sheet_name = nvkt[:31]
            detail_df = _build_detail_sheet(nvkt_df)
            if nvkt_col in detail_df.columns:
                detail_df = detail_df.drop(columns=[nvkt_col])
            append_or_replace_sheet(processed_path, sheet_name, add_tt_column(detail_df))

    return processed_path


def _process_i15_generic(
    input_path: str | Path,
    *,
    k_suffix: str,
    history_db_path: str | Path | None = None,
    dsnv_db_path: str | Path = DEFAULT_DSNV_DB_PATH,
    overwrite_processed: bool = False,
) -> Path:
    raw_path = _resolve_path(input_path)
    df = pd.read_excel(raw_path)
    if df.empty:
        processed_path = ensure_processed_workbook(raw_path, overwrite=overwrite_processed)
        for sheet in ("TH_SHC_I15", "TH_SHC_theo_to", "shc_theo_SA", "Bien_dong_tong_hop", "Tang_moi", "Giam_het", "Van_con"):
            append_or_replace_sheet(processed_path, sheet, pd.DataFrame())
        return processed_path

    if "NVKT_DB" not in df.columns and "NVKT_DB_NORMALIZED" in df.columns:
        df["NVKT_DB"] = df["NVKT_DB_NORMALIZED"]
    if "NVKT_DB_NORMALIZED" not in df.columns:
        df["NVKT_DB_NORMALIZED"] = df["NVKT_DB"].apply(normalize_nvkt)
    else:
        df["NVKT_DB_NORMALIZED"] = df["NVKT_DB_NORMALIZED"].apply(normalize_nvkt)
    if "DOI_ONE" not in df.columns and "DOI_CTS" in df.columns:
        df["DOI_ONE"] = df["DOI_CTS"]
    if "DT_ONEDIACHI_ONE" not in df.columns and "DT_ONE" in df.columns:
        df["DT_ONEDIACHI_ONE"] = df["DT_ONE"]

    account_col = "ACCOUNT_CTS" if "ACCOUNT_CTS" in df.columns else "MA_TB"
    if account_col not in df.columns:
        raise ValueError("Bao cao I1.5 thieu cot ACCOUNT_CTS/MA_TB")

    report_date = _derive_report_date(df)
    if "NGAY_SUYHAO" in df.columns:
        df["NGAY_SUYHAO"] = pd.to_datetime(df["NGAY_SUYHAO"], dayfirst=True, errors="coerce").dt.strftime("%d/%m/%Y")
    else:
        df["NGAY_SUYHAO"] = pd.Timestamp(report_date).strftime("%d/%m/%Y")

    df_danhba, df_thong_ke, df_thong_ke_dv = _read_danhba_tables(Path(dsnv_db_path))
    if not df_danhba.empty and account_col in df.columns:
        merge_key = "ACCOUNT_CTS" if "ACCOUNT_CTS" in df.columns else "MA_TB"
        if merge_key in df.columns:
            df = df.merge(df_danhba, left_on=merge_key, right_on="MA_TB", how="left", suffixes=("", "_db"))
            for col in ("THIETBI", "SA", "KETCUOI", "DOI_VT", "NVKT"):
                db_col = f"{col}_db"
                if db_col in df.columns:
                    df[col] = df[col].where(df[col].notna(), df[db_col])
                    df = df.drop(columns=[db_col])
            if "DOI_VT" in df.columns and "DOI_ONE" in df.columns:
                df["DOI_ONE"] = df["DOI_ONE"].where(df["DOI_ONE"].notna(), df["DOI_VT"])
                df = df.drop(columns=["DOI_VT"])
            if "NVKT" in df.columns and "NVKT_DB" in df.columns:
                df["NVKT_DB"] = df["NVKT_DB"].where(df["NVKT_DB"].notna(), df["NVKT"])
                df = df.drop(columns=["NVKT"])
            if "NVKT_DB" in df.columns:
                df["NVKT_DB_NORMALIZED"] = df["NVKT_DB"].apply(normalize_nvkt)

    prev_snapshot = pd.DataFrame()
    if history_db_path is not None:
        history_db = Path(history_db_path)
        history_db.parent.mkdir(parents=True, exist_ok=True)
        conn = sqlite3.connect(history_db)
        try:
            _ensure_history_schema(conn)
            prev_snapshot = _load_previous_snapshot(conn, k_suffix, report_date)
        finally:
            conn.close()

    result_df, by_to_df, sa_df, df_work = _build_today_summary(
        df,
        k_suffix,
        report_date,
        df_thong_ke,
        df_thong_ke_dv,
        account_col,
    )
    change_df = _build_group_change_summary(df_work, prev_snapshot, account_col)
    df_tang, df_giam, df_van, _ = _build_delta_frames(df_work, prev_snapshot, account_col)

    processed_path = _write_workbook(
        raw_path,
        overwrite_processed,
        df_work,
        result_df,
        by_to_df,
        sa_df,
        change_df,
        df_tang,
        df_giam,
        df_van,
    )

    if history_db_path is not None:
        conn = sqlite3.connect(Path(history_db_path))
        try:
            _ensure_history_schema(conn)
            _upsert_history(conn, k_suffix, report_date, df_work, account_col, prev_snapshot, df_thong_ke)
            conn.commit()
        finally:
            conn.close()

    return processed_path


def process_i15_report_api_output(
    input_path: str | Path = DEFAULT_I15_INPUT,
    overwrite_processed: bool = False,
    history_db_path: str | Path | None = DEFAULT_HISTORY_DB_PATH,
    dsnv_db_path: str | Path = DEFAULT_DSNV_DB_PATH,
) -> Path:
    return _process_i15_generic(
        input_path,
        k_suffix="K1",
        history_db_path=history_db_path,
        dsnv_db_path=dsnv_db_path,
        overwrite_processed=overwrite_processed,
    )


def process_i15_k2_report_api_output(
    input_path: str | Path = DEFAULT_I15_K2_INPUT,
    overwrite_processed: bool = False,
    history_db_path: str | Path | None = DEFAULT_HISTORY_DB_PATH,
    dsnv_db_path: str | Path = DEFAULT_DSNV_DB_PATH,
) -> Path:
    return _process_i15_generic(
        input_path,
        k_suffix="K2",
        history_db_path=history_db_path,
        dsnv_db_path=dsnv_db_path,
        overwrite_processed=overwrite_processed,
    )
