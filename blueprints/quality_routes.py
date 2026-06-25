import io
import os
import sqlite3
import unicodedata
from collections import OrderedDict
from datetime import datetime
from threading import Lock

import pandas as pd
from flask import Blueprint, current_app, jsonify, render_template, request, send_file, session

from app_helpers import (
    build_file_info,
    build_multi_sheet_payload,
    build_sheet_payload,
    latest_matching_file,
    read_excel_sheet_cached,
    safe_file_response,
    serialize_dataframe,
)
from auth import get_user_by_username, login_required
from config import (
    BAOCAO_HANOI_DOWNLOADS_DIR,
    SHC_CTS_HISTORY_DB_PATH,
    SHC_NVKT_DETAIL_REPORTS,
    SHC_NVKT_TEAMS,
)
from repositories import (
    load_many_tables_by_date,
    load_c11_nvkt_df,
    load_c12_nvkt_df,
    load_c14_nvkt_df,
    load_c14_tong_hop_df,
    load_i15_dashboard_df,
    load_i15_tracking_df,
    load_table_by_date,
    resolve_date_context,
)
from repositories.sqlite_runtime import read_sql_rows
from services.shc_processing_report import get_shc_processing_report


quality_bp = Blueprint('quality', __name__)


SHC_CTS_REPORT_PATH = '/home/vtst/shc/processed/reports/So_sanh_SHC_theo_ngay_T-1.xlsx'
SHC_CTS_SUMMARY_SHEET = 'Theo_don_vi'
SHC_CTS_DETAIL_SHEET = 'Chi_tiet_NVKT'
SHC_CTS_INTRADAY_REPORT_DIR = '/home/vtst/shc/processed/intraday/reports'
SHC_CTS_INTRADAY_REPORT_PATTERN = 'Bao_cao_tien_trinh_*.xlsx'
SHC_CTS_INTRADAY_PROGRESS_SHEET = 'Theo NVKT'
SHC_CTS_NVKT_DETAIL_ROOT = '/home/vtst/shc/processed'
SHC_CTS_NVKT_DETAIL_PREFIX = 'shc_NVKT_danh_sach_chi_tiet_K1'


CHAT_LUONG_DATE_BINDINGS = [
    {
        'key': 'v_chi_tieu_c_c1_1_report_th_c1_1',
        'report_code': 'chi_tieu_c_c1_1_report',
        'table_name': 'chi_tieu_c_c1_1_report_th_c1_1',
        'order_by': 't."Đơn vị"',
    },
    {
        'key': 'v_chi_tieu_c_c1_1_chitiet_report_chi_tiet',
        'report_code': 'chi_tieu_c_c1_1_chitiet_report',
        'table_name': 'chi_tieu_c_c1_1_chitiet_report_chi_tiet',
        'order_by': 't."TEN_DOI", t."NVKT"',
    },
    {
        'key': 'v_chi_tieu_c_c1_1_chitiet_report_chi_tieu_ko_hen_18h',
        'report_code': 'chi_tieu_c_c1_1_chitiet_report',
        'table_name': 'chi_tieu_c_c1_1_chitiet_report_chi_tieu_ko_hen_18h',
        'order_by': 't."TEN_DOI", t."NVKT"',
    },
    {
        'key': 'v_chi_tieu_c_c1_2_report_th_c1_2',
        'report_code': 'chi_tieu_c_c1_2_report',
        'table_name': 'chi_tieu_c_c1_2_report_th_c1_2',
        'order_by': 't."Đơn vị"',
    },
    {
        'key': 'v_chi_tieu_c_c1_2_chitiet_sm1_report_th_sm1c12_hll_thang',
        'report_code': 'chi_tieu_c_c1_2_chitiet_sm1_report',
        'table_name': 'chi_tieu_c_c1_2_chitiet_sm1_report_th_sm1c12_hll_thang',
        'order_by': 't."TEN_DOI", t."NVKT"',
    },
    {
        'key': 'v_chi_tieu_c_c1_3_report_th_c1_3',
        'report_code': 'chi_tieu_c_c1_3_report',
        'table_name': 'chi_tieu_c_c1_3_report_th_c1_3',
        'order_by': 't."Đơn vị"',
    },
    {
        'key': 'v_chi_tieu_c_c1_4_report_th_c1_4',
        'report_code': 'chi_tieu_c_c1_4_report',
        'table_name': 'chi_tieu_c_c1_4_report_th_c1_4',
        'order_by': 't."Đơn vị"',
    },
    {
        'key': 'v_chi_tieu_c_c1_4_chitiet_report_th_hl_nvkt',
        'report_code': 'chi_tieu_c_c1_4_chitiet_report',
        'table_name': 'chi_tieu_c_c1_4_chitiet_report_th_hl_nvkt',
        'order_by': 't."DOIVT", t."NVKT"',
    },
    {
        'key': 'v_chi_tieu_c_c1_5_report_th_c1_5',
        'report_code': 'chi_tieu_c_c1_5_report',
        'table_name': 'chi_tieu_c_c1_5_report_th_c1_5',
        'order_by': 't."Đơn vị"',
    },
]


I15_K1_DATE_BINDINGS = [
    {
        'key': 'bien_dong_tong_hop',
        'report_code': 'chi_tieu_i_i1_5_report',
        'table_name': 'chi_tieu_i_i1_5_report_bien_dong_tong_hop',
        'order_by': 't."Đơn vị", t."NVKT_DB"',
    },
    {
        'key': 'shc_theo_sa',
        'report_code': 'chi_tieu_i_i1_5_report',
        'table_name': 'chi_tieu_i_i1_5_report_shc_theo_sa',
        'order_by': 't."TT"',
    },
    {
        'key': 'th_shc_i15',
        'report_code': 'chi_tieu_i_i1_5_report',
        'table_name': 'chi_tieu_i_i1_5_report_th_shc_i15',
        'order_by': 't."Đơn vị", t."NVKT_DB"',
    },
    {
        'key': 'th_shc_theo_to',
        'report_code': 'chi_tieu_i_i1_5_report',
        'table_name': 'chi_tieu_i_i1_5_report_th_shc_theo_to',
        'order_by': 't."TT"',
    },
]


I15_K2_DATE_BINDINGS = [
    {
        'key': 'bien_dong_tong_hop',
        'report_code': 'chi_tieu_i_i1_5_k2_report',
        'table_name': 'chi_tieu_i_i1_5_k2_report_bien_dong_tong_hop',
        'order_by': 't."Đơn vị", t."NVKT_DB"',
    },
    {
        'key': 'shc_theo_sa',
        'report_code': 'chi_tieu_i_i1_5_k2_report',
        'table_name': 'chi_tieu_i_i1_5_k2_report_shc_theo_sa',
        'order_by': 't."TT"',
    },
    {
        'key': 'th_shc_i15',
        'report_code': 'chi_tieu_i_i1_5_k2_report',
        'table_name': 'chi_tieu_i_i1_5_k2_report_th_shc_i15',
        'order_by': 't."Đơn vị", t."NVKT_DB"',
    },
    {
        'key': 'th_shc_theo_to',
        'report_code': 'chi_tieu_i_i1_5_k2_report',
        'table_name': 'chi_tieu_i_i1_5_k2_report_th_shc_theo_to',
        'order_by': 't."TT"',
    },
]


I15_DETAIL_DB_CONFIG = {
    'k1': {
        'label': 'Suy hao cao K1',
        'report_code': 'chi_tieu_i_i1_5_report',
        'summary_table': 'chi_tieu_i_i1_5_report_th_shc_i15',
        'table_prefix': 'chi_tieu_i_i1_5_report_',
    },
    'k2': {
        'label': 'Suy hao cao K2',
        'report_code': 'chi_tieu_i_i1_5_k2_report',
        'summary_table': 'chi_tieu_i_i1_5_k2_report_th_shc_i15',
        'table_prefix': 'chi_tieu_i_i1_5_k2_report_',
    },
}

I15_DETAIL_EXCLUDED_TABLE_SUFFIXES = {
    'bien_dong_tong_hop',
    'shc_theo_sa',
    'th_shc_i15',
    'th_shc_theo_to',
}


def _current_user():
    return get_user_by_username(session['username'])


def _sorted_group_payload(df, group_column, *, rate_col=None, ascending=False, strip_percent=False):
    grouped = OrderedDict()
    if group_column not in df.columns:
        return grouped

    for group_value in df[group_column].dropna().unique():
        if not str(group_value).strip():
            continue
        filtered_df = df[df[group_column] == group_value].copy()
        if rate_col and rate_col in filtered_df.columns:
            rate_series = filtered_df[rate_col]
            if strip_percent:
                rate_series = rate_series.astype(str).str.replace('%', '', regex=False)
            filtered_df[rate_col] = pd.to_numeric(rate_series, errors='coerce')
            filtered_df = filtered_df.sort_values(by=rate_col, ascending=ascending, na_position='last')
        grouped[str(group_value)] = build_sheet_payload(filtered_df)
    return grouped


def _round_float_columns(df):
    rounded = df.copy()
    for col in rounded.columns:
        if pd.api.types.is_float_dtype(rounded[col]):
            rounded[col] = rounded[col].round(2)
    return rounded


def _build_i15_team_summary_df(detail_df):
    summary_rows = []
    if 'Đơn vị' not in detail_df.columns:
        return pd.DataFrame(summary_rows)

    grouped = detail_df.groupby('Đơn vị', dropna=False, sort=False)
    for don_vi, group_df in grouped:
        if not str(don_vi).strip():
            continue

        so_tb_suy_hao_k1 = _numeric_sum(group_df['Số tăng mới']) if 'Số tăng mới' in group_df.columns else 0
        so_tb_quan_ly = _numeric_sum(group_df['Số TB quản lý']) if 'Số TB quản lý' in group_df.columns else 0
        ti_le_shc = round((so_tb_suy_hao_k1 * 100.0 / so_tb_quan_ly) if so_tb_quan_ly else 0, 2)

        summary_rows.append(
            {
                'stt': len(summary_rows) + 1,
                'tt': len(summary_rows) + 1,
                'don_vi': str(don_vi),
                'so_tb_suy_hao_k1': int(so_tb_suy_hao_k1),
                'so_tb_quan_ly': int(so_tb_quan_ly),
                'ti_le_shc': ti_le_shc,
            }
        )

    if summary_rows:
        total_so_tb_suy_hao_k1 = sum(row['so_tb_suy_hao_k1'] for row in summary_rows)
        total_so_tb_quan_ly = sum(row['so_tb_quan_ly'] for row in summary_rows)
        summary_rows.append(
            {
                'stt': len(summary_rows) + 1,
                'tt': len(summary_rows) + 1,
                'don_vi': 'Tổng',
                'so_tb_suy_hao_k1': int(total_so_tb_suy_hao_k1),
                'so_tb_quan_ly': int(total_so_tb_quan_ly),
                'ti_le_shc': round(
                    (total_so_tb_suy_hao_k1 * 100.0 / total_so_tb_quan_ly) if total_so_tb_quan_ly else 0,
                    2,
                ),
            }
        )

    return _round_float_columns(pd.DataFrame(summary_rows))


def _detail_group_response(excel_path, sheet_name, group_column, *, rate_col=None, ascending=False, strip_percent=False):
    if not os.path.exists(excel_path):
        return None

    df = read_excel_sheet_cached(excel_path, sheet_name)
    return {
        'file_info': build_file_info(excel_path),
        'doi_data': _sorted_group_payload(
            df,
            group_column,
            rate_col=rate_col,
            ascending=ascending,
            strip_percent=strip_percent,
        ),
    }


def _resolve_excel_sheet_name(excel_path, preferred_sheet):
    workbook = pd.ExcelFile(excel_path)
    sheet_names = workbook.sheet_names

    if preferred_sheet in sheet_names:
        return preferred_sheet

    if len(sheet_names) == 1:
        return sheet_names[0]

    normalized_sheet_map = {str(name).strip().lower(): name for name in sheet_names}
    normalized_preferred = str(preferred_sheet).strip().lower()
    if normalized_preferred in normalized_sheet_map:
        return normalized_sheet_map[normalized_preferred]

    raise ValueError(
        f"Worksheet named '{preferred_sheet}' not found. Available sheets: {', '.join(sheet_names)}"
    )


def _numeric_sum(series):
    return pd.to_numeric(series, errors='coerce').fillna(0).sum()


def _last_non_empty(series, fallback=None):
    for value in series.tolist():
        if value is None:
            continue
        text = str(value).strip()
        if text:
            return value
    return fallback


def _normalize_text_key(value):
    normalized = unicodedata.normalize('NFKD', str(value or '').strip())
    ascii_text = ''.join(char for char in normalized if not unicodedata.combining(char))
    return ' '.join(ascii_text.lower().split())


def _slugify_table_suffix(value):
    normalized = unicodedata.normalize('NFKD', str(value or '').strip())
    ascii_text = ''.join(char for char in normalized if not unicodedata.combining(char))
    cleaned = []
    last_was_separator = False
    for char in ascii_text.lower():
        if char.isalnum():
            cleaned.append(char)
            last_was_separator = False
            continue
        if not last_was_separator:
            cleaned.append('_')
            last_was_separator = True
    return ''.join(cleaned).strip('_')


def _get_i15_detail_db_source(report_type):
    return I15_DETAIL_DB_CONFIG.get(report_type)


def _list_i15_personal_tables(report_type):
    db_source = _get_i15_detail_db_source(report_type)
    if not db_source:
        return []

    status_placeholders = ', '.join('?' for _ in ('thanh_cong', 'khong_co_sheet_tong_hop'))
    rows = read_sql_rows(
        f'''
        SELECT DISTINCT
            s.ten_bang_du_lieu AS table_name,
            COALESCE(NULLIF(TRIM(s.ten_sheet), ''), s.ten_bang_du_lieu) AS sheet_name
        FROM sheet_bao_cao_tong_hop s
        JOIN bao_cao_tong_hop_ngay b
          ON b.id = s.bao_cao_tong_hop_ngay_id
        WHERE b.ma_bao_cao = ?
          AND s.ten_bang_du_lieu LIKE ?
          AND b.trang_thai_nap IN ({status_placeholders})
        ORDER BY s.ten_bang_du_lieu
        ''',
        (
            db_source['report_code'],
            f"{db_source['table_prefix']}%",
            'thanh_cong',
            'khong_co_sheet_tong_hop',
        ),
    )

    personal_tables = []
    for row in rows:
        table_name = str(row.get('table_name') or '').strip()
        if not table_name.startswith(db_source['table_prefix']):
            continue
        suffix = table_name[len(db_source['table_prefix']):]
        if not suffix or suffix in I15_DETAIL_EXCLUDED_TABLE_SUFFIXES:
            continue
        sheet_name = str(row.get('sheet_name') or '').strip() or table_name
        personal_tables.append(
            {
                'table_name': table_name,
                'sheet_name': sheet_name,
                'normalized_name': _normalize_text_key(sheet_name),
                'slug_name': _slugify_table_suffix(sheet_name),
            }
        )
    return personal_tables


def _resolve_i15_detail_context(report_type):
    db_source = _get_i15_detail_db_source(report_type)
    if not db_source:
        return None, None

    date_context = resolve_date_context(request.args.get('date'), [db_source['report_code']])
    return db_source, date_context


def _build_i15_personal_file_items(report_type, team):
    db_source, date_context = _resolve_i15_detail_context(report_type)
    team_name = SHC_NVKT_TEAMS.get(team)
    if not db_source:
        return None, None, None, 'Loại báo cáo không hợp lệ'
    if not team_name:
        return None, None, None, 'Tổ không hợp lệ'

    if not date_context['date_has_data']:
        return db_source, team_name, [], None

    detail_df = load_table_by_date(
        db_source['report_code'],
        db_source['summary_table'],
        date_context['selected_date'],
        order_by='t."Đơn vị", t."NVKT_DB"',
    )
    if detail_df.empty or 'Đơn vị' not in detail_df.columns or 'NVKT_DB' not in detail_df.columns:
        return db_source, team_name, [], None

    personal_tables = _list_i15_personal_tables(report_type)
    by_table_name = {item['table_name']: item for item in personal_tables}
    by_normalized_name = {item['normalized_name']: item for item in personal_tables}
    by_slug_name = {item['slug_name']: item for item in personal_tables}

    team_df = detail_df[detail_df['Đơn vị'].astype(str).str.strip() == team_name].copy()
    if team_df.empty:
        return db_source, team_name, [], None

    file_items = []
    seen_tables = set()
    for nvkt_name in team_df['NVKT_DB'].dropna().tolist():
        nvkt_text = str(nvkt_name).strip()
        if not nvkt_text:
            continue

        catalog_item = by_normalized_name.get(_normalize_text_key(nvkt_text))
        if not catalog_item:
            expected_table_name = f"{db_source['table_prefix']}{_slugify_table_suffix(nvkt_text)}"
            catalog_item = by_table_name.get(expected_table_name) or by_slug_name.get(_slugify_table_suffix(nvkt_text))
        if not catalog_item:
            continue
        if catalog_item['table_name'] in seen_tables:
            continue

        seen_tables.add(catalog_item['table_name'])
        file_items.append(
            {
                'name': catalog_item['table_name'],
                'display_name': catalog_item['sheet_name'],
                'size': None,
                'modified': date_context['selected_date'],
            }
        )

    file_items.sort(key=lambda item: _normalize_text_key(item['display_name']))
    return db_source, team_name, file_items, None


def _resolve_i15_personal_selection(report_type, team, table_name):
    db_source, date_context = _resolve_i15_detail_context(report_type)
    team_name = SHC_NVKT_TEAMS.get(team)
    if not db_source:
        return None, None, None, None, None, 'Loại báo cáo không hợp lệ'
    if not team_name:
        return None, None, None, None, None, 'Tổ không hợp lệ'

    normalized_table_name = str(table_name or '').strip()
    if not normalized_table_name:
        return None, None, None, None, None, 'Chưa chọn cá nhân chi tiết'

    personal_tables = {item['table_name']: item for item in _list_i15_personal_tables(report_type)}
    selected_item = personal_tables.get(normalized_table_name)
    if not selected_item:
        return None, None, None, None, None, 'Cá nhân chi tiết không hợp lệ'

    if not date_context['date_has_data']:
        return db_source, date_context, team_name, normalized_table_name, selected_item, None

    available_items = {item['name'] for item in (_build_i15_personal_file_items(report_type, team)[2] or [])}
    if normalized_table_name not in available_items:
        return None, None, None, None, None, 'Cá nhân chi tiết không thuộc tổ đã chọn'

    return db_source, date_context, team_name, normalized_table_name, selected_item, None


def _sheet_payload_to_dataframe(sheet_payload):
    columns = list(sheet_payload.get('columns') or [])
    rows = list(sheet_payload.get('data') or [])
    if not columns and not rows:
        return pd.DataFrame()
    df = pd.DataFrame(rows)
    if columns:
        for column in columns:
            if column not in df.columns:
                df[column] = ''
        df = df[columns]
    return df


def _safe_excel_sheet_name(name, used_names):
    base_name = str(name or 'Sheet').strip() or 'Sheet'
    cleaned = ''.join('_' if char in '[]:*?/\\\\' else char for char in base_name)
    cleaned = cleaned[:31] or 'Sheet'

    candidate = cleaned
    counter = 2
    while candidate in used_names:
        suffix = f'_{counter}'
        candidate = f"{cleaned[:31-len(suffix)]}{suffix}"
        counter += 1

    used_names.add(candidate)
    return candidate


def _build_i15_download_workbook_response(payload, filename_prefix):
    workbook_sheets = OrderedDict()
    workbook_sheets['Tong_hop_theo_to'] = _sheet_payload_to_dataframe(payload.get('tong_hop') or {})
    workbook_sheets['Bien_dong_theo_don_vi'] = _sheet_payload_to_dataframe(payload.get('theo_don_vi') or {})
    workbook_sheets['Theo_doi_SHC_theo_SA'] = _sheet_payload_to_dataframe(payload.get('shc_theo_sa') or {})

    for don_vi_name, sheet_payload in (payload.get('don_vi') or {}).items():
        workbook_sheets[f"Chi_tiet_{don_vi_name}"] = _sheet_payload_to_dataframe(sheet_payload)

    for don_vi_name, sheet_payload in (payload.get('chi_tiet_nvkt') or {}).items():
        workbook_sheets[f"Bien_dong_{don_vi_name}"] = _sheet_payload_to_dataframe(sheet_payload)

    buffer = io.BytesIO()
    used_sheet_names = set()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        wrote_any_sheet = False
        for suggested_name, df in workbook_sheets.items():
            sheet_name = _safe_excel_sheet_name(suggested_name, used_sheet_names)
            export_df = serialize_dataframe(df)
            export_df.to_excel(writer, sheet_name=sheet_name, index=False)
            wrote_any_sheet = True
        if not wrote_any_sheet:
            pd.DataFrame().to_excel(writer, sheet_name='Sheet1', index=False)
    buffer.seek(0)

    selected_date = payload.get('selected_date') or 'latest'
    return send_file(
        buffer,
        as_attachment=True,
        download_name=f"{filename_prefix}_{selected_date}.xlsx",
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )


def _file_created_timestamp(file_path):
    stat = os.stat(file_path)
    created_at = getattr(stat, 'st_birthtime', stat.st_mtime)
    return datetime.fromtimestamp(created_at).strftime('%d/%m/%Y %H:%M:%S')


def _add_progress_timestamp_column(progress_df, timestamp):
    progress_df = progress_df.copy()
    if 'Timestamp' in progress_df.columns:
        progress_df = progress_df.drop(columns=['Timestamp'])
    insert_at = 1 if 'Đơn vị' in progress_df.columns else 0
    progress_df.insert(insert_at, 'Timestamp', timestamp)
    return progress_df


SHC_CTS_KIEMSOAT_NOI_DUNG_MAX = 2000
SHC_CTS_INTRADAY_PROGRESS_COLUMNS = [
    'Đơn vị',
    'NVKT_DB',
    'Tổng số',
    'Đạt baseline',
    'Đã xử lý trong ngày',
    'Tổng đã đạt',
    'Chưa đạt',
    'OFF/Lỗi',
    '% đạt',
]

_shc_cts_schema_lock = Lock()
_shc_cts_schema_ready_path = None


def _shc_cts_write_connection():
    conn = sqlite3.connect(SHC_CTS_HISTORY_DB_PATH, timeout=5)
    conn.execute('PRAGMA journal_mode=WAL')
    conn.execute('PRAGMA busy_timeout=5000')
    conn.row_factory = sqlite3.Row
    return conn


def _shc_cts_read_connection():
    conn = sqlite3.connect(
        f'file:{os.path.abspath(SHC_CTS_HISTORY_DB_PATH)}?mode=ro',
        uri=True,
        timeout=5,
    )
    conn.row_factory = sqlite3.Row
    return conn


def _ensure_shc_cts_schema():
    global _shc_cts_schema_ready_path
    db_path = os.path.abspath(SHC_CTS_HISTORY_DB_PATH)
    with _shc_cts_schema_lock:
        if _shc_cts_schema_ready_path == db_path:
            return
        db_dir = os.path.dirname(db_path)
        if db_dir:
            os.makedirs(db_dir, exist_ok=True)
        with _shc_cts_write_connection() as conn:
            conn.execute(
                '''
                CREATE TABLE IF NOT EXISTS shc_cts_tien_do (
                    ngay_xu_ly        TEXT NOT NULL,
                    don_vi            TEXT,
                    nvkt_db           TEXT NOT NULL,
                    tong_so           INTEGER,
                    dat_baseline      INTEGER,
                    da_xu_ly_ngay     INTEGER,
                    tong_dat          INTEGER,
                    chua_dat          INTEGER,
                    off_loi           INTEGER,
                    ty_le_dat         REAL,
                    captured_at       TEXT,
                    PRIMARY KEY (ngay_xu_ly, nvkt_db)
                )
                '''
            )
            conn.execute(
                'CREATE INDEX IF NOT EXISTS idx_shc_cts_tien_do_ngay ON shc_cts_tien_do(ngay_xu_ly)'
            )
            conn.execute(
                'CREATE INDEX IF NOT EXISTS idx_shc_cts_tien_do_don_vi ON shc_cts_tien_do(don_vi)'
            )
            conn.execute(
                '''
                CREATE TABLE IF NOT EXISTS shc_cts_kiemsoat (
                    ngay_xu_ly         TEXT NOT NULL,
                    nvkt_db            TEXT NOT NULL,
                    don_vi             TEXT,
                    noi_dung_kiem_soat TEXT NOT NULL DEFAULT '',
                    nguoi_nhap         TEXT,
                    thoi_diem_nhap     TEXT,
                    thoi_diem_cap_nhat TEXT,
                    PRIMARY KEY (ngay_xu_ly, nvkt_db)
                )
                '''
            )
        _shc_cts_schema_ready_path = db_path


def _shc_cts_payload_from_excel():
    if not os.path.exists(SHC_CTS_REPORT_PATH):
        raise FileNotFoundError(f'File Excel SHC CTS không tồn tại: {SHC_CTS_REPORT_PATH}')

    intraday_pattern = os.path.join(SHC_CTS_INTRADAY_REPORT_DIR, SHC_CTS_INTRADAY_REPORT_PATTERN)
    intraday_report_path = latest_matching_file(intraday_pattern)
    if not intraday_report_path:
        raise FileNotFoundError(
            f'Không tìm thấy file Excel tiến trình SHC trong ngày theo mẫu: {intraday_pattern}'
        )

    summary_df = read_excel_sheet_cached(SHC_CTS_REPORT_PATH, SHC_CTS_SUMMARY_SHEET)
    detail_df = read_excel_sheet_cached(SHC_CTS_REPORT_PATH, SHC_CTS_DETAIL_SHEET)
    progress_df = read_excel_sheet_cached(intraday_report_path, SHC_CTS_INTRADAY_PROGRESS_SHEET)
    progress_df = _add_progress_timestamp_column(progress_df, _file_created_timestamp(intraday_report_path))

    don_vi_data = OrderedDict()
    if 'Đơn vị' in detail_df.columns:
        for don_vi in detail_df['Đơn vị'].dropna().unique():
            don_vi_text = str(don_vi).strip()
            if not don_vi_text:
                continue
            don_vi_df = detail_df[detail_df['Đơn vị'] == don_vi].copy()
            if 'NVKT' in don_vi_df.columns:
                don_vi_df = don_vi_df.sort_values(by='NVKT', ascending=True, na_position='last')
            don_vi_data[don_vi_text] = build_sheet_payload(don_vi_df)

    progress_by_unit = OrderedDict()
    if 'Đơn vị' in progress_df.columns:
        for don_vi in progress_df['Đơn vị'].dropna().unique():
            don_vi_text = str(don_vi).strip()
            if not don_vi_text:
                continue
            unit_df = progress_df[progress_df['Đơn vị'] == don_vi].copy()
            if 'NVKT_DB' in unit_df.columns:
                unit_df = unit_df.sort_values(by='NVKT_DB', ascending=True, na_position='last')
            progress_by_unit[don_vi_text] = build_sheet_payload(unit_df)

    return {
        'file_info': build_file_info(SHC_CTS_REPORT_PATH, include_name=True),
        'tien_do_file_info': build_file_info(intraday_report_path, include_name=True),
        'tong_hop': build_sheet_payload(summary_df),
        'don_vi': don_vi_data,
        'tien_do_xu_ly': build_sheet_payload(progress_df),
        'tien_do_theo_don_vi': progress_by_unit,
    }


def _build_shc_cts_download_workbook_response(payload):
    buffer = io.BytesIO()
    used_sheet_names = set()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        _sheet_payload_to_dataframe(payload.get('tong_hop') or {}).to_excel(
            writer,
            sheet_name=_safe_excel_sheet_name('Tong_hop_SHC_CTS_theo_to', used_sheet_names),
            index=False,
        )
        for don_vi_name, sheet_payload in (payload.get('don_vi') or {}).items():
            _sheet_payload_to_dataframe(sheet_payload).to_excel(
                writer,
                sheet_name=_safe_excel_sheet_name(f'Chi_tiet_{don_vi_name}', used_sheet_names),
                index=False,
            )
        _sheet_payload_to_dataframe(payload.get('tien_do_xu_ly') or {}).to_excel(
            writer,
            sheet_name=_safe_excel_sheet_name('Tien_do_xu_ly_SHC_trong_ngay', used_sheet_names),
            index=False,
        )
    buffer.seek(0)

    return send_file(
        buffer,
        as_attachment=True,
        download_name='SHC_CTS_theo_ngay_T-1.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )


def _parse_shc_cts_nvkt_detail_dir_date(dir_name):
    prefix = f'{SHC_CTS_NVKT_DETAIL_PREFIX}-'
    if not str(dir_name).startswith(prefix):
        return None
    date_text = str(dir_name)[len(prefix):]
    try:
        return datetime.strptime(date_text, '%d-%m-%Y').date()
    except ValueError:
        return None


def _latest_shc_cts_nvkt_detail_dir():
    if not os.path.isdir(SHC_CTS_NVKT_DETAIL_ROOT):
        return None

    candidates = []
    for name in os.listdir(SHC_CTS_NVKT_DETAIL_ROOT):
        path = os.path.join(SHC_CTS_NVKT_DETAIL_ROOT, name)
        if not os.path.isdir(path):
            continue
        if name == SHC_CTS_NVKT_DETAIL_PREFIX or name.startswith(f'{SHC_CTS_NVKT_DETAIL_PREFIX}-'):
            parsed_date = _parse_shc_cts_nvkt_detail_dir_date(name)
            candidates.append((parsed_date is not None, parsed_date, os.path.getmtime(path), path))

    if not candidates:
        return None

    candidates.sort(key=lambda item: (item[0], item[1] or datetime.min.date(), item[2]), reverse=True)
    return candidates[0][3]


def _list_shc_cts_nvkt_detail_teams():
    source_dir = _latest_shc_cts_nvkt_detail_dir()
    if not source_dir:
        return None, []

    teams = []
    for name in sorted(os.listdir(source_dir), key=lambda value: _normalize_text_key(value)):
        path = os.path.join(source_dir, name)
        if os.path.isdir(path):
            teams.append({'key': name, 'label': name})
    return source_dir, teams


def _resolve_shc_cts_nvkt_detail_team(team):
    source_dir = _latest_shc_cts_nvkt_detail_dir()
    if not source_dir:
        return None, None, 'Không tìm thấy thư mục chi tiết SHC NVKT K1'

    team_name = str(team or '').strip()
    if not team_name or os.path.basename(team_name) != team_name:
        return None, None, 'Tổ không hợp lệ'

    team_dir = os.path.abspath(os.path.join(source_dir, team_name))
    source_abs = os.path.abspath(source_dir)
    try:
        if os.path.commonpath([team_dir, source_abs]) != source_abs:
            return None, None, 'Đường dẫn tổ không hợp lệ'
    except ValueError:
        return None, None, 'Đường dẫn tổ không hợp lệ'

    if not os.path.isdir(team_dir):
        return None, None, 'Tổ không tồn tại'

    return source_dir, team_dir, None


def _list_shc_cts_nvkt_detail_files(team):
    source_dir, team_dir, error = _resolve_shc_cts_nvkt_detail_team(team)
    if error:
        return source_dir, None, [], error

    file_items = []
    for file_name in sorted(os.listdir(team_dir), key=lambda value: _normalize_text_key(os.path.splitext(value)[0])):
        file_path = os.path.join(team_dir, file_name)
        if not os.path.isfile(file_path) or not file_name.lower().endswith(('.xlsx', '.xls')):
            continue
        file_items.append(
            {
                'name': file_name,
                'display_name': os.path.splitext(file_name)[0],
                'size': os.path.getsize(file_path),
                'modified': build_file_info(file_path)['modified'],
            }
        )

    return source_dir, team_dir, file_items, None


def _resolve_shc_cts_nvkt_detail_file(team, filename):
    source_dir, team_dir, error = _resolve_shc_cts_nvkt_detail_team(team)
    if error:
        return source_dir, None, None, error

    normalized_filename = os.path.basename(filename)
    if not normalized_filename or normalized_filename != filename:
        return source_dir, team_dir, None, 'Tên file không hợp lệ'
    if not normalized_filename.lower().endswith(('.xlsx', '.xls')):
        return source_dir, team_dir, None, 'Định dạng file không được hỗ trợ'

    file_path = os.path.abspath(os.path.join(team_dir, normalized_filename))
    try:
        if os.path.commonpath([file_path, team_dir]) != team_dir:
            return source_dir, team_dir, None, 'Đường dẫn file không hợp lệ'
    except ValueError:
        return source_dir, team_dir, None, 'Đường dẫn file không hợp lệ'

    if not os.path.isfile(file_path):
        return source_dir, team_dir, None, 'File không tồn tại'

    return source_dir, team_dir, file_path, None


def _i15_k1_payload_from_tables():
    date_context = resolve_date_context(
        request.args.get('date'),
        [binding['report_code'] for binding in I15_K1_DATE_BINDINGS],
    )
    file_info = build_file_info(current_app.config['REPORT_HISTORY_DB_PATH'], include_name=True)

    if not date_context['date_has_data']:
        empty_df = pd.DataFrame()
        return {
            'file_info': file_info,
            'selected_date': date_context['selected_date'],
            'latest_available_date': date_context['latest_available_date'],
            'available_dates': date_context['available_dates'],
            'date_has_data': date_context['date_has_data'],
            'default_k_suffix': 'K1',
            'tong_hop': build_sheet_payload(empty_df),
            'don_vi': OrderedDict(),
            'theo_don_vi': build_sheet_payload(empty_df),
            'chi_tiet_nvkt': OrderedDict(),
            'shc_theo_sa': build_sheet_payload(empty_df),
        }

    tables = load_many_tables_by_date(I15_K1_DATE_BINDINGS, date_context['selected_date'])
    tong_hop_df = tables['th_shc_theo_to'].copy()
    detail_df = tables['th_shc_i15'].copy()
    bien_dong_df = tables['bien_dong_tong_hop'].copy()
    shc_theo_sa_df = tables['shc_theo_sa'].copy()

    don_vi_data = OrderedDict()
    if 'Đơn vị' in detail_df.columns:
        for don_vi in detail_df['Đơn vị'].dropna().unique():
            if not str(don_vi).strip():
                continue
            don_vi_df = detail_df[detail_df['Đơn vị'] == don_vi].copy()
            don_vi_df = don_vi_df.sort_values(by=['NVKT_DB', 'Số TB Suy hao cao K1'], ascending=[True, False], na_position='last')
            don_vi_data[str(don_vi)] = build_sheet_payload(don_vi_df)

    variation_rows = []
    if 'Đơn vị' in bien_dong_df.columns:
        for don_vi, group_df in bien_dong_df.groupby('Đơn vị', dropna=False, sort=False):
            if not str(don_vi).strip():
                continue
            variation_rows.append(
                {
                    'Đơn vị': str(don_vi),
                    'Tổng số hiện tại': int(_numeric_sum(group_df['Tổng số hiện tại'])) if 'Tổng số hiện tại' in group_df.columns else 0,
                    'Tăng mới': int(_numeric_sum(group_df['Tăng mới'])) if 'Tăng mới' in group_df.columns else 0,
                    'Giảm/Hết': int(_numeric_sum(group_df['Giảm/Hết'])) if 'Giảm/Hết' in group_df.columns else 0,
                    'Vẫn còn': int(_numeric_sum(group_df['Vẫn còn'])) if 'Vẫn còn' in group_df.columns else 0,
                    'Số TB quản lý': int(_numeric_sum(group_df['Số TB quản lý'])) if 'Số TB quản lý' in group_df.columns else 0,
                    'Tỉ lệ SHC (%)': round(
                        (_numeric_sum(group_df['Tăng mới']) * 100.0 / _numeric_sum(group_df['Số TB quản lý']))
                        if 'Tăng mới' in group_df.columns and 'Số TB quản lý' in group_df.columns and _numeric_sum(group_df['Số TB quản lý'])
                        else 0,
                        2,
                    ),
                }
            )

    variation_df = pd.DataFrame(variation_rows)

    tracking_by_unit = OrderedDict()
    if 'Đơn vị' in bien_dong_df.columns:
        for don_vi in bien_dong_df['Đơn vị'].dropna().unique():
            if not str(don_vi).strip():
                continue
            sheet_df = bien_dong_df[bien_dong_df['Đơn vị'] == don_vi].copy()
            sheet_df = sheet_df.sort_values(
                by=['Tăng mới', 'Tổng số hiện tại', 'NVKT_DB'],
                ascending=[False, False, True],
                na_position='last',
            )
            tracking_by_unit[str(don_vi)] = build_sheet_payload(sheet_df)

    return {
        'file_info': file_info,
        'selected_date': date_context['selected_date'],
        'latest_available_date': date_context['latest_available_date'],
        'available_dates': date_context['available_dates'],
        'date_has_data': date_context['date_has_data'],
        'default_k_suffix': 'K1',
        'tong_hop': build_sheet_payload(tong_hop_df),
        'don_vi': don_vi_data,
        'theo_don_vi': build_sheet_payload(variation_df),
        'chi_tiet_nvkt': tracking_by_unit,
        'shc_theo_sa': build_sheet_payload(shc_theo_sa_df),
    }


def _i15_k2_payload_from_tables():
    date_context = resolve_date_context(
        request.args.get('date'),
        [binding['report_code'] for binding in I15_K2_DATE_BINDINGS],
    )
    file_info = build_file_info(current_app.config['REPORT_HISTORY_DB_PATH'], include_name=True)

    if not date_context['date_has_data']:
        empty_df = pd.DataFrame()
        return {
            'file_info': file_info,
            'selected_date': date_context['selected_date'],
            'latest_available_date': date_context['latest_available_date'],
            'available_dates': date_context['available_dates'],
            'date_has_data': date_context['date_has_data'],
            'default_k_suffix': 'K2',
            'tong_hop': build_sheet_payload(empty_df),
            'don_vi': OrderedDict(),
            'theo_don_vi': build_sheet_payload(empty_df),
            'chi_tiet_nvkt': OrderedDict(),
            'shc_theo_sa': build_sheet_payload(empty_df),
        }

    tables = load_many_tables_by_date(I15_K2_DATE_BINDINGS, date_context['selected_date'])
    tong_hop_df = tables['th_shc_theo_to'].copy()
    detail_df = tables['th_shc_i15'].copy()
    bien_dong_df = tables['bien_dong_tong_hop'].copy()
    shc_theo_sa_df = tables['shc_theo_sa'].copy()

    don_vi_data = OrderedDict()
    if 'Đơn vị' in detail_df.columns:
        for don_vi in detail_df['Đơn vị'].dropna().unique():
            if not str(don_vi).strip():
                continue
            don_vi_df = detail_df[detail_df['Đơn vị'] == don_vi].copy()
            don_vi_df = don_vi_df.sort_values(by=['NVKT_DB', 'Số TB Suy hao cao K2'], ascending=[True, False], na_position='last')
            don_vi_data[str(don_vi)] = build_sheet_payload(don_vi_df)

    variation_rows = []
    if 'Đơn vị' in bien_dong_df.columns:
        for don_vi, group_df in bien_dong_df.groupby('Đơn vị', dropna=False, sort=False):
            if not str(don_vi).strip():
                continue
            variation_rows.append(
                {
                    'Đơn vị': str(don_vi),
                    'Tổng số hiện tại': int(_numeric_sum(group_df['Tổng số hiện tại'])) if 'Tổng số hiện tại' in group_df.columns else 0,
                    'Tăng mới': int(_numeric_sum(group_df['Tăng mới'])) if 'Tăng mới' in group_df.columns else 0,
                    'Giảm/Hết': int(_numeric_sum(group_df['Giảm/Hết'])) if 'Giảm/Hết' in group_df.columns else 0,
                    'Vẫn còn': int(_numeric_sum(group_df['Vẫn còn'])) if 'Vẫn còn' in group_df.columns else 0,
                    'Số TB quản lý': int(_numeric_sum(group_df['Số TB quản lý'])) if 'Số TB quản lý' in group_df.columns else 0,
                    'Tỉ lệ SHC (%)': round(
                        (_numeric_sum(group_df['Tăng mới']) * 100.0 / _numeric_sum(group_df['Số TB quản lý']))
                        if 'Tăng mới' in group_df.columns and 'Số TB quản lý' in group_df.columns and _numeric_sum(group_df['Số TB quản lý'])
                        else 0,
                        2,
                    ),
                }
            )

    variation_df = pd.DataFrame(variation_rows)

    tracking_by_unit = OrderedDict()
    if 'Đơn vị' in bien_dong_df.columns:
        for don_vi in bien_dong_df['Đơn vị'].dropna().unique():
            if not str(don_vi).strip():
                continue
            sheet_df = bien_dong_df[bien_dong_df['Đơn vị'] == don_vi].copy()
            sheet_df = sheet_df.sort_values(
                by=['Tăng mới', 'Tổng số hiện tại', 'NVKT_DB'],
                ascending=[False, False, True],
                na_position='last',
            )
            tracking_by_unit[str(don_vi)] = build_sheet_payload(sheet_df)

    return {
        'file_info': file_info,
        'selected_date': date_context['selected_date'],
        'latest_available_date': date_context['latest_available_date'],
        'available_dates': date_context['available_dates'],
        'date_has_data': date_context['date_has_data'],
        'default_k_suffix': 'K2',
        'tong_hop': build_sheet_payload(tong_hop_df),
        'don_vi': don_vi_data,
        'theo_don_vi': build_sheet_payload(variation_df),
        'chi_tiet_nvkt': tracking_by_unit,
        'shc_theo_sa': build_sheet_payload(shc_theo_sa_df),
    }


def _i15_payload(default_k_suffix):
    detail_df = load_i15_dashboard_df()
    tracking_df = load_i15_tracking_df()
    file_info = build_file_info(current_app.config['REPORT_HISTORY_DB_PATH'], include_name=True)

    if detail_df.empty:
        return None

    selected_detail_df = detail_df[detail_df['Kỳ'] == default_k_suffix].copy()
    if selected_detail_df.empty:
        selected_detail_df = detail_df.copy()

    selected_tracking_df = tracking_df[tracking_df['Kỳ'] == default_k_suffix].copy()
    if selected_tracking_df.empty:
        selected_tracking_df = tracking_df.copy()

    summary_df = _build_i15_team_summary_df(selected_detail_df)

    don_vi_data = OrderedDict()
    if 'Đơn vị' in selected_detail_df.columns:
        for don_vi in selected_detail_df['Đơn vị'].dropna().unique():
            if not str(don_vi).strip():
                continue
            sheet_df = selected_detail_df[selected_detail_df['Đơn vị'] == don_vi].copy()
            sheet_df = sheet_df.sort_values(
                by=['NVKT', 'Tổng số hiện tại'],
                ascending=[True, False],
                na_position='last',
            )
            don_vi_data[str(don_vi)] = build_sheet_payload(sheet_df)

    variation_rows = []
    if 'Đơn vị' in selected_detail_df.columns:
        grouped = selected_detail_df.groupby('Đơn vị', dropna=False)
        for don_vi, group_df in grouped:
            if not str(don_vi).strip():
                continue
            tong_so_hien_tai = _numeric_sum(group_df['Tổng số hiện tại'])
            so_tang_moi = _numeric_sum(group_df['Số tăng mới'])
            so_giam_het = _numeric_sum(group_df['Số giảm/hết'])
            so_van_con = _numeric_sum(group_df['Số vẫn còn'])
            so_tb_quan_ly = _numeric_sum(group_df['Số TB quản lý'])
            ty_le_shc = round((so_tang_moi * 100.0 / so_tb_quan_ly) if so_tb_quan_ly else 0, 2)
            variation_rows.append(
                {
                    'Kỳ': default_k_suffix,
                    'Đơn vị': str(don_vi),
                    'Tổng số hiện tại': int(tong_so_hien_tai),
                    'Số tăng mới': int(so_tang_moi),
                    'Số giảm/hết': int(so_giam_het),
                    'Số vẫn còn': int(so_van_con),
                    'Số TB quản lý': int(so_tb_quan_ly),
                    'Tỷ lệ SHC (%)': ty_le_shc,
                }
            )

    variation_df = pd.DataFrame(variation_rows)
    variation_df = _round_float_columns(variation_df)

    tracking_by_unit = OrderedDict()
    if 'Đơn vị' in selected_tracking_df.columns:
        for don_vi in selected_tracking_df['Đơn vị'].dropna().unique():
            if not str(don_vi).strip():
                continue
            sheet_df = selected_tracking_df[selected_tracking_df['Đơn vị'] == don_vi].copy()
            sheet_df = sheet_df.sort_values(
                by=['Số ngày liên tục', 'NVKT', 'Account CTS'],
                ascending=[False, True, True],
                na_position='last',
            )
            tracking_by_unit[str(don_vi)] = build_sheet_payload(sheet_df)

    return {
        'file_info': file_info,
        'default_k_suffix': default_k_suffix,
        'tong_hop': build_sheet_payload(summary_df),
        'don_vi': don_vi_data,
        'theo_don_vi': build_sheet_payload(variation_df),
        'chi_tiet_nvkt': tracking_by_unit,
        'shc_theo_sa': build_sheet_payload(selected_tracking_df),
    }


def _get_shc_nvkt_report_and_team(report_type, team):
    report_info = SHC_NVKT_DETAIL_REPORTS.get(report_type)
    team_name = SHC_NVKT_TEAMS.get(team)
    return report_info, team_name


def _list_shc_nvkt_detail_files(report_type, team):
    report_info, team_name = _get_shc_nvkt_report_and_team(report_type, team)
    if not report_info:
        return None, None, None, 'Loại báo cáo không hợp lệ'
    if not team_name:
        return None, None, None, 'Tổ không hợp lệ'

    report_dir = report_info['path']
    team_dir = os.path.join(report_dir, team_name)
    if not os.path.isdir(team_dir):
        return report_info, team_name, [], None

    file_items = []
    for file_name in sorted(os.listdir(team_dir), key=lambda name: name.lower()):
        file_path = os.path.join(team_dir, file_name)
        if not os.path.isfile(file_path):
            continue
        if not file_name.lower().endswith(('.xlsx', '.xls')):
            continue

        file_items.append({
            'name': file_name,
            'display_name': os.path.splitext(file_name)[0],
            'size': os.path.getsize(file_path),
            'modified': build_file_info(file_path)['modified'],
        })

    return report_info, team_name, file_items, None


def _resolve_shc_nvkt_detail_file(report_type, team, filename):
    report_info, team_name = _get_shc_nvkt_report_and_team(report_type, team)
    if not report_info:
        return None, None, None, None, 'Loại báo cáo không hợp lệ'
    if not team_name:
        return None, None, None, None, 'Tổ không hợp lệ'

    normalized_filename = os.path.basename(filename)
    if not normalized_filename or normalized_filename != filename:
        return None, None, None, None, 'Tên file không hợp lệ'
    if not normalized_filename.lower().endswith(('.xlsx', '.xls')):
        return None, None, None, None, 'Định dạng file không được hỗ trợ'

    team_dir = os.path.abspath(os.path.join(report_info['path'], team_name))
    file_path = os.path.abspath(os.path.join(team_dir, normalized_filename))

    try:
        if os.path.commonpath([file_path, team_dir]) != team_dir:
            return None, None, None, None, 'Đường dẫn file không hợp lệ'
    except ValueError:
        return None, None, None, None, 'Đường dẫn file không hợp lệ'

    if not os.path.exists(file_path):
        return None, None, None, None, 'File không tồn tại'

    return report_info, team_name, normalized_filename, file_path, None


@quality_bp.route('/chatluong')
@login_required
def page_chatluong():
    return render_template('pages/chatluong.html', current_user=_current_user(), active_page='chatluong')


@quality_bp.route('/i15')
@login_required
def page_i15():
    return render_template('pages/i15.html', current_user=_current_user(), active_page='i15')


@quality_bp.route('/i15k2')
@login_required
def page_i15k2():
    return render_template('pages/i15k2.html', current_user=_current_user(), active_page='i15k2')


@quality_bp.route('/shc-cts')
@login_required
def page_shc_cts():
    return render_template('pages/shc_cts.html', current_user=_current_user(), active_page='shc_cts')


@quality_bp.route('/shc-processing')
@login_required
def page_shc_processing():
    return render_template('pages/shc_processing_report.html', current_user=_current_user(), active_page='shc_processing')


@quality_bp.route('/download/excel-i15')
def download_excel_i15():
    try:
        payload = _i15_k1_payload_from_tables()
        return _build_i15_download_workbook_response(payload, 'I1.5_report')
    except Exception as exc:
        return jsonify({'error': f'Lỗi khi kết xuất file I1.5 từ SQLite: {exc}'}), 500


@quality_bp.route('/download/excel-i15k2')
def download_excel_i15k2():
    try:
        payload = _i15_k2_payload_from_tables()
        return _build_i15_download_workbook_response(payload, 'I1.5_k2_report')
    except Exception as exc:
        return jsonify({'error': f'Lỗi khi kết xuất file I1.5 K2 từ SQLite: {exc}'}), 500


@quality_bp.route('/download/excel-shc-cts')
def download_excel_shc_cts():
    try:
        payload = _shc_cts_payload_from_excel()
        return _build_shc_cts_download_workbook_response(payload)
    except FileNotFoundError as exc:
        return jsonify({'error': str(exc)}), 404
    except Exception as exc:
        return jsonify({'error': f'Lỗi khi kết xuất file SHC CTS từ Excel: {exc}'}), 500


@quality_bp.route('/download/excel-chatluong/<file_type>')
def download_excel_chatluong(file_type):
    file_mapping = {
        'c1.1': ('c1.1 report.xlsx', 'C1.1_report.xlsx'),
        'c1.1-chitiet-suachua': ('SM2-C11.xlsx', 'C1.1_ChiTiet_PhieuSuaChua_BRCD.xlsx'),
        'c1.1-chitiet-chatluong': ('SM4-C11.xlsx', 'C1.1_ChiTiet_PhieuChatLuong_FiberMyTV.xlsx'),
        'c1.2': ('c1.2 report.xlsx', 'C1.2_report.xlsx'),
        'c1.2-chitiet-laplai': ('SM1-C12.xlsx', 'C1.2_ChiTiet_HongLapLai_7Ngay.xlsx'),
        'c1.3': ('c1.3 report.xlsx', 'C1.3_report.xlsx'),
        'c1.4': ('c1.4 report.xlsx', 'C1.4_report.xlsx'),
        'c1.4-chitiet': ('c1.4_chitiet_report.xlsx', 'C1.4_ChiTiet_DoHaiLong_NVKT.xlsx'),
        'c1.5': ('c1.5 report.xlsx', 'C1.5_report.xlsx'),
        'c1.5-chitiet': ('c1.5_chitiet_report.xlsx', 'C1.5_ChiTiet_ThietLapDichVu_NVKT.xlsx'),
    }
    if file_type not in file_mapping:
        return jsonify({'error': 'File type không hợp lệ'}), 400

    filename, download_name = file_mapping[file_type]
    excel_path = os.path.join(BAOCAO_HANOI_DOWNLOADS_DIR, filename)
    return safe_file_response(
        excel_path,
        as_attachment=True,
        download_name=download_name,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )


@quality_bp.route('/api/c1-chat-luong-data')
def get_c1_chat_luong_data():
    try:
        date_context = resolve_date_context(
            request.args.get('date'),
            [binding['report_code'] for binding in CHAT_LUONG_DATE_BINDINGS],
        )
        file_info = build_file_info(current_app.config['REPORT_HISTORY_DB_PATH'])
        if date_context['date_has_data']:
            dataframes = load_many_tables_by_date(CHAT_LUONG_DATE_BINDINGS, date_context['selected_date'])
        else:
            dataframes = {binding['key']: pd.DataFrame() for binding in CHAT_LUONG_DATE_BINDINGS}

        return jsonify(
            {
                'file_info': file_info,
                'selected_date': date_context['selected_date'],
                'latest_available_date': date_context['latest_available_date'],
                'available_dates': date_context['available_dates'],
                'date_has_data': date_context['date_has_data'],
                'sheets': OrderedDict(
                    [
                        (
                            'v_chi_tieu_c_c1_1_report_th_c1_1',
                            build_sheet_payload(dataframes['v_chi_tieu_c_c1_1_report_th_c1_1']),
                        ),
                        (
                            'v_chi_tieu_c_c1_1_chitiet_report_chi_tiet',
                            build_sheet_payload(dataframes['v_chi_tieu_c_c1_1_chitiet_report_chi_tiet']),
                        ),
                        (
                            'v_chi_tieu_c_c1_1_chitiet_report_chi_tieu_ko_hen_18h',
                            build_sheet_payload(dataframes['v_chi_tieu_c_c1_1_chitiet_report_chi_tieu_ko_hen_18h']),
                        ),
                        (
                            'v_chi_tieu_c_c1_2_report_th_c1_2',
                            build_sheet_payload(dataframes['v_chi_tieu_c_c1_2_report_th_c1_2']),
                        ),
                        (
                            'v_chi_tieu_c_c1_2_chitiet_sm1_report_th_sm1c12_hll_thang',
                            build_sheet_payload(dataframes['v_chi_tieu_c_c1_2_chitiet_sm1_report_th_sm1c12_hll_thang']),
                        ),
                        (
                            'v_chi_tieu_c_c1_3_report_th_c1_3',
                            build_sheet_payload(dataframes['v_chi_tieu_c_c1_3_report_th_c1_3']),
                        ),
                        (
                            'v_chi_tieu_c_c1_4_report_th_c1_4',
                            build_sheet_payload(dataframes['v_chi_tieu_c_c1_4_report_th_c1_4']),
                        ),
                        (
                            'v_chi_tieu_c_c1_4_chitiet_report_th_hl_nvkt',
                            build_sheet_payload(dataframes['v_chi_tieu_c_c1_4_chitiet_report_th_hl_nvkt']),
                        ),
                        (
                            'v_chi_tieu_c_c1_5_report_th_c1_5',
                            build_sheet_payload(dataframes['v_chi_tieu_c_c1_5_report_th_c1_5']),
                        ),
                    ]
                ),
            }
        )
    except Exception as exc:
        current_app.logger.exception('Khong the tai du lieu chat luong tong hop tu SQLite: %s', exc)
        return jsonify({'error': f'Lỗi khi đọc dữ liệu chất lượng từ SQLite: {exc}'}), 500


@quality_bp.route('/api/c11-chi-tiet-data')
def get_c11_chi_tiet_data():
    try:
        df = load_c11_nvkt_df('tong')
        data = {'file_info': build_file_info(current_app.config['REPORT_HISTORY_DB_PATH']), 'doi_data': _sorted_group_payload(df, 'TEN_DOI', rate_col='Tỷ lệ phiếu sửa chữa báo hỏng dịch vụ BRCD đúng quy định không tính hẹn', ascending=False)}
        return jsonify(data)
    except Exception as exc:
        return jsonify({'error': f'Lỗi khi đọc dữ liệu C1.1 từ SQLite: {exc}'}), 500


@quality_bp.route('/api/c11-chi-tiet-15h-data')
def get_c11_chi_tiet_15h_data():
    try:
        df = load_c11_nvkt_df('15h')
        data = {'file_info': build_file_info(current_app.config['REPORT_HISTORY_DB_PATH']), 'doi_data': _sorted_group_payload(df, 'TEN_DOI', rate_col='Tỷ lệ phiếu sửa chữa báo hỏng dịch vụ BRCD đúng quy định không tính hẹn', ascending=False)}
        return jsonify(data)
    except Exception as exc:
        return jsonify({'error': f'Lỗi khi đọc dữ liệu C1.1 15h từ SQLite: {exc}'}), 500


@quality_bp.route('/api/c11-chi-tiet-16h-data')
def get_c11_chi_tiet_16h_data():
    try:
        df = load_c11_nvkt_df('16h')
        data = {'file_info': build_file_info(current_app.config['REPORT_HISTORY_DB_PATH']), 'doi_data': _sorted_group_payload(df, 'TEN_DOI', rate_col='Tỷ lệ phiếu sửa chữa báo hỏng dịch vụ BRCD đúng quy định không tính hẹn', ascending=False)}
        return jsonify(data)
    except Exception as exc:
        return jsonify({'error': f'Lỗi khi đọc dữ liệu C1.1 16h từ SQLite: {exc}'}), 500


@quality_bp.route('/api/c11-chi-tiet-17h-data')
def get_c11_chi_tiet_17h_data():
    try:
        df = load_c11_nvkt_df('17h')
        data = {'file_info': build_file_info(current_app.config['REPORT_HISTORY_DB_PATH']), 'doi_data': _sorted_group_payload(df, 'TEN_DOI', rate_col='Tỷ lệ phiếu sửa chữa báo hỏng dịch vụ BRCD đúng quy định không tính hẹn', ascending=False)}
        return jsonify(data)
    except Exception as exc:
        return jsonify({'error': f'Lỗi khi đọc dữ liệu C1.1 17h từ SQLite: {exc}'}), 500


@quality_bp.route('/api/c11-chi-tiet-18h-data')
def get_c11_chi_tiet_18h_data():
    try:
        df = load_c11_nvkt_df('18h')
        data = {'file_info': build_file_info(current_app.config['REPORT_HISTORY_DB_PATH']), 'doi_data': _sorted_group_payload(df, 'TEN_DOI', rate_col='Tỷ lệ phiếu sửa chữa báo hỏng dịch vụ BRCD đúng quy định không tính hẹn', ascending=False)}
        return jsonify(data)
    except Exception as exc:
        return jsonify({'error': f'Lỗi khi đọc dữ liệu C1.1 18h từ SQLite: {exc}'}), 500


@quality_bp.route('/api/c12-repeat-failure-data')
def get_c12_repeat_failure_data():
    try:
        df = load_c12_nvkt_df()
        data = {'file_info': build_file_info(current_app.config['REPORT_HISTORY_DB_PATH']), 'doi_data': _sorted_group_payload(df, 'TEN_DOI', rate_col='Tỉ lệ HLL tháng (2.5%)', ascending=False)}
        return jsonify(data)
    except Exception as exc:
        return jsonify({'error': f'Lỗi khi đọc dữ liệu C1.2 từ SQLite: {exc}'}), 500


@quality_bp.route('/api/c11-chi-tiet-quality-data')
def get_c11_chi_tiet_quality_data():
    try:
        df = load_c11_nvkt_df('tong')
        data = {'file_info': build_file_info(current_app.config['REPORT_HISTORY_DB_PATH']), 'doi_data': _sorted_group_payload(df, 'TEN_DOI', rate_col='Tỷ lệ phiếu sửa chữa báo hỏng dịch vụ BRCD đúng quy định không tính hẹn', ascending=False)}
        return jsonify(data)
    except Exception as exc:
        return jsonify({'error': f'Lỗi khi đọc dữ liệu C1.1 chất lượng chủ động từ SQLite: {exc}'}), 500


@quality_bp.route('/api/c14-chi-tiet-data')
def get_c14_chi_tiet_data():
    try:
        df = load_c14_nvkt_df()
        data = {'file_info': build_file_info(current_app.config['REPORT_HISTORY_DB_PATH']), 'doi_data': _sorted_group_payload(df, 'DOIVT', rate_col='Tỉ lệ HL NVKT (%)', ascending=True)}
        return jsonify(data)
    except Exception as exc:
        return jsonify({'error': f'Lỗi khi đọc dữ liệu C1.4 từ SQLite: {exc}'}), 500


@quality_bp.route('/api/c15-chi-tiet-data')
def get_c15_chi_tiet_data():
    excel_path = os.path.join(BAOCAO_HANOI_DOWNLOADS_DIR, 'c1.5_chitiet_report.xlsx')
    if not os.path.exists(excel_path):
        return jsonify({'error': 'File Excel c1.5_chitiet_report.xlsx không tồn tại'}), 404

    try:
        df_ttvtst = read_excel_sheet_cached(excel_path, 'TH_TTVTST')
        df_detail = read_excel_sheet_cached(excel_path, 'KQ_C15_chitiet')
        df_dvvt_ttvt = read_excel_sheet_cached(excel_path, 'TH_DVVT_TTVT')

        doi_data = _sorted_group_payload(
            df_detail,
            'DOIVT',
            rate_col='Tỉ lệ đạt',
            ascending=True,
            strip_percent=True,
        )

        return jsonify({
            'file_info': build_file_info(excel_path),
            'ttvtst': build_sheet_payload(df_ttvtst),
            'dvvt_ttvt': build_sheet_payload(df_dvvt_ttvt),
            'doi_data': doi_data,
        })
    except Exception as exc:
        return jsonify({'error': f'Lỗi khi đọc file Excel c1.5_chitiet_report.xlsx: {exc}'}), 500


@quality_bp.route('/api/i15-data')
def get_i15_data():
    data = _i15_k1_payload_from_tables()
    if data is None:
        return jsonify({'error': 'Không tìm thấy dữ liệu I1.5 trong report_history.db'}), 404
    try:
        return jsonify(data)
    except Exception as exc:
        return jsonify({'error': f'Lỗi khi đọc dữ liệu I1.5 từ SQLite: {exc}'}), 500


@quality_bp.route('/api/shc-variation-data')
def get_shc_variation_data():
    try:
        data = _i15_k1_payload_from_tables()
        return jsonify(data)
    except Exception as exc:
        return jsonify({'error': f'Lỗi khi đọc dữ liệu SHC từ SQLite: {exc}'}), 500


@quality_bp.route('/api/i15k2-data')
def get_i15k2_data():
    data = _i15_k2_payload_from_tables()
    if data is None:
        return jsonify({'error': 'Không tìm thấy dữ liệu I1.5 K2 trong report_history.db'}), 404
    try:
        return jsonify(data)
    except Exception as exc:
        return jsonify({'error': f'Lỗi khi đọc dữ liệu I1.5 K2 từ SQLite: {exc}'}), 500


@quality_bp.route('/api/shc-cts-data')
def get_shc_cts_data():
    try:
        return jsonify(_shc_cts_payload_from_excel())
    except FileNotFoundError as exc:
        return jsonify({'error': str(exc)}), 404
    except Exception as exc:
        return jsonify({'error': f'Lỗi khi đọc dữ liệu SHC CTS từ Excel: {exc}'}), 500


@quality_bp.route('/api/shc-cts-nvkt-detail/options')
@login_required
def get_shc_cts_nvkt_detail_options():
    source_dir, teams = _list_shc_cts_nvkt_detail_teams()
    if not source_dir:
        return jsonify({'error': 'Không tìm thấy thư mục chi tiết SHC NVKT K1'}), 404

    return jsonify(
        {
            'source_dir': source_dir,
            'source_name': os.path.basename(source_dir),
            'teams': teams,
        }
    )


@quality_bp.route('/api/shc-cts-nvkt-detail/files')
@login_required
def get_shc_cts_nvkt_detail_files():
    team = request.args.get('team', '').strip()
    source_dir, team_dir, files, error = _list_shc_cts_nvkt_detail_files(team)
    if error:
        status = 404 if error in {'Không tìm thấy thư mục chi tiết SHC NVKT K1', 'Tổ không tồn tại'} else 400
        return jsonify({'error': error}), status

    return jsonify(
        {
            'source_dir': source_dir,
            'source_name': os.path.basename(source_dir),
            'team': {'key': os.path.basename(team_dir), 'label': os.path.basename(team_dir)},
            'files': files,
        }
    )


@quality_bp.route('/api/shc-cts-nvkt-detail/preview')
@login_required
def preview_shc_cts_nvkt_detail_file():
    team = request.args.get('team', '').strip()
    filename = request.args.get('file_name', '').strip()
    source_dir, team_dir, file_path, error = _resolve_shc_cts_nvkt_detail_file(team, filename)
    if error:
        status = 404 if error in {'Không tìm thấy thư mục chi tiết SHC NVKT K1', 'Tổ không tồn tại', 'File không tồn tại'} else 400
        return jsonify({'error': error}), status

    try:
        payload = build_multi_sheet_payload(
            file_path,
            include_file_name=True,
            include_sheet_errors=True,
        )
        payload['source_dir'] = source_dir
        payload['source_name'] = os.path.basename(source_dir)
        payload['team'] = {'key': os.path.basename(team_dir), 'label': os.path.basename(team_dir)}
        payload['selected_file'] = os.path.basename(file_path)
        return jsonify(payload)
    except Exception as exc:
        return jsonify({'error': f'Lỗi khi đọc file Excel chi tiết SHC CTS: {exc}'}), 500


@quality_bp.route('/download/shc-cts-nvkt-detail/<team>/<path:filename>')
@login_required
def download_shc_cts_nvkt_detail_file(team, filename):
    _, _, file_path, error = _resolve_shc_cts_nvkt_detail_file(team, filename)
    if error:
        status = 404 if error in {'Không tìm thấy thư mục chi tiết SHC NVKT K1', 'Tổ không tồn tại', 'File không tồn tại'} else 400
        return jsonify({'error': error}), status

    normalized_filename = os.path.basename(file_path)
    return safe_file_response(
        file_path,
        as_attachment=True,
        download_name=normalized_filename,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )


@quality_bp.route('/api/shc-variation-k2-data')
def get_shc_variation_k2_data():
    try:
        data = _i15_k2_payload_from_tables()
        return jsonify(data)
    except Exception as exc:
        return jsonify({'error': f'Lỗi khi đọc dữ liệu SHC K2 từ SQLite: {exc}'}), 500


@quality_bp.route('/api/shc-processing-report')
@login_required
def get_shc_processing_report_data():
    report_date = request.args.get('report_date', '').strip()
    team_filter = request.args.get('team_filter', '').strip()

    try:
        return jsonify(get_shc_processing_report(report_date=report_date, team_filter=team_filter))
    except FileNotFoundError as exc:
        return jsonify({'error': str(exc)}), 404
    except Exception as exc:
        return jsonify({'error': f'Lỗi khi tổng hợp báo cáo xử lý SHC: {exc}'}), 500


@quality_bp.route('/api/shc-nvkt-detail/options')
@login_required
def get_shc_nvkt_detail_options():
    requested_report_type = request.args.get('report_type', '').strip()
    report_types = [{'key': key, 'label': value['label']} for key, value in SHC_NVKT_DETAIL_REPORTS.items()]
    teams = [{'key': key, 'label': value} for key, value in SHC_NVKT_TEAMS.items()]
    _, date_context = _resolve_i15_detail_context(requested_report_type) if requested_report_type else (None, None)
    payload = {'report_types': report_types, 'teams': teams}
    if date_context:
        payload.update(
            {
                'selected_date': date_context['selected_date'],
                'latest_available_date': date_context['latest_available_date'],
                'available_dates': date_context['available_dates'],
                'date_has_data': date_context['date_has_data'],
            }
        )
    return jsonify(payload)


@quality_bp.route('/api/shc-nvkt-detail/files')
@login_required
def get_shc_nvkt_detail_files():
    report_type = request.args.get('report_type', '').strip()
    team = request.args.get('team', '').strip()

    db_source, team_name, files, error = _build_i15_personal_file_items(report_type, team)
    if db_source:
        if error:
            return jsonify({'error': error}), 400
        _, date_context = _resolve_i15_detail_context(report_type)
        return jsonify(
            {
                'report_type': {'key': report_type, 'label': db_source['label']},
                'team': {'key': team, 'label': team_name},
                'files': files,
                'selected_date': date_context['selected_date'],
                'latest_available_date': date_context['latest_available_date'],
                'available_dates': date_context['available_dates'],
                'date_has_data': date_context['date_has_data'],
                'source': 'report_history_db',
            }
        )

    report_info, team_name, files, error = _list_shc_nvkt_detail_files(report_type, team)
    if error:
        return jsonify({'error': error}), 400

    return jsonify({
        'report_type': {'key': report_type, 'label': report_info['label']},
        'team': {'key': team, 'label': team_name},
        'files': files,
    })


@quality_bp.route('/api/shc-nvkt-detail/preview')
@login_required
def preview_shc_nvkt_detail_file():
    report_type = request.args.get('report_type', '').strip()
    team = request.args.get('team', '').strip()
    filename = request.args.get('file_name', '').strip()

    db_source, date_context, team_name, normalized_table_name, selected_item, error = _resolve_i15_personal_selection(
        report_type,
        team,
        filename,
    )
    if db_source:
        if error:
            status = 404 if error == 'Cá nhân chi tiết không thuộc tổ đã chọn' else 400
            return jsonify({'error': error}), status

        if not date_context['date_has_data']:
            return jsonify(
                {
                    'error': f"Ngày {date_context['selected_date']} hiện chưa có dữ liệu cho báo cáo {db_source['label']}",
                    'selected_date': date_context['selected_date'],
                    'latest_available_date': date_context['latest_available_date'],
                    'available_dates': date_context['available_dates'],
                    'date_has_data': date_context['date_has_data'],
                }
            ), 404

        detail_df = load_table_by_date(
            db_source['report_code'],
            normalized_table_name,
            date_context['selected_date'],
            order_by='t."TT"',
        )
        payload = {
            'file_info': {
                'name': f"{selected_item['sheet_name']}.xlsx",
                'modified': date_context['selected_date'],
                'size': len(detail_df.index),
            },
            'report_type': {'key': report_type, 'label': db_source['label']},
            'team': {'key': team, 'label': team_name},
            'selected_file': normalized_table_name,
            'selected_date': date_context['selected_date'],
            'latest_available_date': date_context['latest_available_date'],
            'available_dates': date_context['available_dates'],
            'date_has_data': date_context['date_has_data'],
            'source': 'report_history_db',
            'sheets': OrderedDict([(selected_item['sheet_name'], build_sheet_payload(detail_df))]),
        }
        return jsonify(payload)

    report_info, team_name, normalized_filename, file_path, error = _resolve_shc_nvkt_detail_file(
        report_type,
        team,
        filename,
    )
    if error:
        status = 404 if error == 'File không tồn tại' else 400
        return jsonify({'error': error}), status

    try:
        payload = build_multi_sheet_payload(
            file_path,
            include_file_name=True,
            include_sheet_errors=True,
        )
        payload['report_type'] = {'key': report_type, 'label': report_info['label']}
        payload['team'] = {'key': team, 'label': team_name}
        payload['selected_file'] = normalized_filename
        return jsonify(payload)
    except Exception as exc:
        return jsonify({'error': f'Lỗi khi đọc file Excel chi tiết: {exc}'}), 500


@quality_bp.route('/download/shc-nvkt-detail/<report_type>/<team>/<path:filename>')
@login_required
def download_shc_nvkt_detail_file(report_type, team, filename):
    db_source, date_context, team_name, normalized_table_name, selected_item, error = _resolve_i15_personal_selection(
        report_type,
        team,
        filename,
    )
    if db_source:
        if error:
            status = 404 if error == 'Cá nhân chi tiết không thuộc tổ đã chọn' else 400
            return jsonify({'error': error}), status
        if not date_context['date_has_data']:
            return (
                jsonify(
                    {
                        'error': f"Ngày {date_context['selected_date']} hiện chưa có dữ liệu cho báo cáo {db_source['label']}",
                        'selected_date': date_context['selected_date'],
                        'latest_available_date': date_context['latest_available_date'],
                        'available_dates': date_context['available_dates'],
                        'date_has_data': date_context['date_has_data'],
                    }
                ),
                404,
            )

        detail_df = load_table_by_date(
            db_source['report_code'],
            normalized_table_name,
            date_context['selected_date'],
            order_by='t."TT"',
        )
        export_df = serialize_dataframe(detail_df)
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            export_df.to_excel(writer, sheet_name=selected_item['sheet_name'][:31], index=False)
        buffer.seek(0)

        download_name = f"{selected_item['sheet_name']}_{date_context['selected_date']}.xlsx"
        return send_file(
            buffer,
            as_attachment=True,
            download_name=download_name,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        )

    _, _, normalized_filename, file_path, error = _resolve_shc_nvkt_detail_file(report_type, team, filename)
    if error:
        status = 404 if error == 'File không tồn tại' else 400
        return jsonify({'error': error}), status

    return safe_file_response(
        file_path,
        as_attachment=True,
        download_name=normalized_filename,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )
