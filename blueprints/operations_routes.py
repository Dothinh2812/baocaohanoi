import io
import os
import re
import sqlite3
import unicodedata
from datetime import date, datetime, timedelta
from threading import Lock

import pandas as pd
from flask import Blueprint, abort, current_app, jsonify, render_template, request, send_file, session

from app_helpers import (
    build_file_info,
    build_multi_sheet_payload,
    build_sheet_payload,
    read_excel_sheet_cached,
    safe_directory_response,
    safe_file_response,
    serialize_dataframe,
)
from auth import get_user_by_username, login_required
from config import (
    BASE_DATA_PATH,
    BAOCAO_HANOI_KPI_DIR,
    BAOCAO_HANOI_PATH,
    BRCD_KIEMSOAT_DB_PATH,
    KPI_TONGHOP_NVKT_FILE,
    REPORT_HISTORY_DB_PATH,
    UNIT_NAME,
)
from repositories import (
    load_bsc_kpi_cac_to_df,
    load_don_vi_tong_hop_da_nguon_df,
    load_many_tables_by_date,
    load_nvkt_tong_hop_da_nguon_by_name_df,
    load_nvkt_tong_hop_da_nguon_df,
    resolve_date_context,
    load_tong_hop_bsc_kpi_rows,
)
from repositories import load_dashboard_kpi_nvkt_df, load_kpi_nvkt_tong_hop_df


operations_bp = Blueprint('operations', __name__)


BRCD_RUNTIME_DOWNLOADS_DIR = '/home/vtst/1bss/runtime/default/downloads'
BRCD_SUMMARY_FILE = os.path.join(BRCD_RUNTIME_DOWNLOADS_DIR, 'kq_dhsc', 'bc_BRCD.xlsx')
BRCD_DETAIL_MAIN_FILE = os.path.join(BRCD_RUNTIME_DOWNLOADS_DIR, 'chiaTheoDoi', 'chiTietBrcd5Doi.xlsx')
BRCD_DETAIL_OFF_FILE = os.path.join(BRCD_RUNTIME_DOWNLOADS_DIR, 'chiaTheoDoi', 'chiTietBrcd5Doi_OFF.xlsx')
PTTB_RUNTIME_DOWNLOADS_DIR = '/home/vtst/1bss/runtime/default/downloads'
PTTB_SUMMARY_FILE = os.path.join(PTTB_RUNTIME_DOWNLOADS_DIR, 'ton_pttb', 'baoCaoPTTB.xlsx')


# ---------------------------------------------------------------------------
# Kiểm soát tổ trưởng (BRCD)
# Lớp annotation ghi được, per-instance, keyed bằng baohong_id.
# Tồn live vẫn đọc Excel read-only; annotation lưu riêng để không bị 1bss ghi đè.
# ---------------------------------------------------------------------------
BRCD_KIEMSOAT_DISPLAY_COLUMNS = [
    'baohong_id',
    'ma_tb',
    'TEN_TB',
    'DIACHI_LD',
    'LOAIHINH_TB',
    'GHICHU_HONG',
    'NVKT',
    'DOI_VT',
    'ngay_bh',
    'Trạng thái cổng',
    'ttvt_ton',
    'chitieu_tg',
    'thời gian tồn thực',
    'giờ còn lại thực',
    'SA',
]
BRCD_KIEMSOAT_NOI_DUNG_MAX = 2000

_brcd_kiemsoat_schema_lock = Lock()
_brcd_kiemsoat_schema_ready_path = None


def _brcd_kiemsoat_write_connection():
    conn = sqlite3.connect(BRCD_KIEMSOAT_DB_PATH, timeout=5)
    conn.execute('PRAGMA journal_mode=WAL')
    conn.execute('PRAGMA busy_timeout=5000')
    conn.row_factory = sqlite3.Row
    return conn


def _brcd_kiemsoat_read_connection():
    conn = sqlite3.connect(
        f'file:{os.path.abspath(BRCD_KIEMSOAT_DB_PATH)}?mode=ro',
        uri=True,
        timeout=5,
    )
    conn.row_factory = sqlite3.Row
    return conn


def _ensure_brcd_kiemsoat_schema():
    global _brcd_kiemsoat_schema_ready_path
    db_path = os.path.abspath(BRCD_KIEMSOAT_DB_PATH)
    with _brcd_kiemsoat_schema_lock:
        if _brcd_kiemsoat_schema_ready_path == db_path:
            return
        db_dir = os.path.dirname(db_path)
        if db_dir:
            os.makedirs(db_dir, exist_ok=True)
        with _brcd_kiemsoat_write_connection() as conn:
            conn.execute(
                '''
                CREATE TABLE IF NOT EXISTS brcd_kiemsoat (
                    baohong_id INTEGER PRIMARY KEY,
                    ma_tb TEXT,
                    doi_vt TEXT,
                    nvkt TEXT,
                    noi_dung_kiem_soat TEXT NOT NULL DEFAULT '',
                    nguoi_nhap TEXT,
                    thoi_diem_nhap TEXT,
                    thoi_diem_cap_nhat TEXT
                )
                '''
            )
            conn.execute(
                '''
                CREATE TABLE IF NOT EXISTS brcd_phieu (
                    baohong_id         INTEGER PRIMARY KEY,
                    ma_tb              TEXT,
                    ten_tb             TEXT,
                    diachi_ld          TEXT,
                    loaihinh_tb        TEXT,
                    ghichu_hong        TEXT,
                    nvkt               TEXT,
                    doi_vt             TEXT,
                    ngay_bh            TEXT,
                    trang_thai_cong    TEXT,
                    ttvt_ton           TEXT,
                    chitieu_tg         REAL,
                    thoi_gian_ton_thuc REAL,
                    gio_con_lai_thuc   REAL,
                    sa                 TEXT,
                    sheet              TEXT,
                    first_seen         TEXT,
                    last_seen          TEXT
                )
                '''
            )
            conn.execute(
                'CREATE INDEX IF NOT EXISTS idx_brcd_phieu_last_seen ON brcd_phieu(last_seen)'
            )
            conn.execute(
                'CREATE INDEX IF NOT EXISTS idx_brcd_phieu_doi_vt ON brcd_phieu(doi_vt)'
            )
        _brcd_kiemsoat_schema_ready_path = db_path


def get_brcd_kiemsoat_map(baohong_ids):
    """Trả dict {baohong_id(int): row(dict)} cho danh sách id đang tồn."""
    _ensure_brcd_kiemsoat_schema()
    ids = [int(i) for i in baohong_ids if i is not None]
    if not ids:
        return {}
    placeholders = ', '.join('?' for _ in ids)
    with _brcd_kiemsoat_read_connection() as conn:
        rows = conn.execute(
            f'''
            SELECT baohong_id, ma_tb, doi_vt, nvkt, noi_dung_kiem_soat,
                   nguoi_nhap, thoi_diem_nhap, thoi_diem_cap_nhat
            FROM brcd_kiemsoat
            WHERE baohong_id IN ({placeholders})
            ''',
            ids,
        ).fetchall()
    return {row['baohong_id']: dict(row) for row in rows}


# ---------------------------------------------------------------------------
# Snapshot Vũ trụ tổng vào brcd_phieu (lịch sử phiếu tồn)
# ---------------------------------------------------------------------------

_BRCD_PHIEU_UPSERT_SQL = """
    INSERT INTO brcd_phieu (
        baohong_id, ma_tb, ten_tb, diachi_ld, loaihinh_tb, ghichu_hong,
        nvkt, doi_vt, ngay_bh, trang_thai_cong, ttvt_ton,
        chitieu_tg, thoi_gian_ton_thuc, gio_con_lai_thuc, sa,
        sheet, first_seen, last_seen
    )
    VALUES (
        ?, ?, ?, ?, ?, ?,
        ?, ?, ?, ?, ?,
        ?, ?, ?, ?,
        ?, ?, ?
    )
    ON CONFLICT(baohong_id) DO UPDATE SET
        ma_tb = excluded.ma_tb,
        ten_tb = excluded.ten_tb,
        diachi_ld = excluded.diachi_ld,
        loaihinh_tb = excluded.loaihinh_tb,
        ghichu_hong = excluded.ghichu_hong,
        nvkt = excluded.nvkt,
        doi_vt = excluded.doi_vt,
        ngay_bh = excluded.ngay_bh,
        trang_thai_cong = excluded.trang_thai_cong,
        ttvt_ton = excluded.ttvt_ton,
        chitieu_tg = excluded.chitieu_tg,
        thoi_gian_ton_thuc = excluded.thoi_gian_ton_thuc,
        gio_con_lai_thuc = excluded.gio_con_lai_thuc,
        sa = excluded.sa,
        sheet = excluded.sheet,
        last_seen = excluded.last_seen
"""


def _phieu_row_to_params(row, sheet, now_iso):
    """Map 1 dòng Excel (Series) sang tuple params cho UPSERT."""
    def _val(col):
        v = row.get(col)
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return None
        return v
    return (
        int(row['baohong_id']),
        _val('ma_tb'),
        _val('TEN_TB'),
        _val('DIACHI_LD'),
        _val('LOAIHINH_TB'),
        _val('GHICHU_HONG'),
        _val('NVKT'),
        _val('DOI_VT'),
        _val('ngay_bh'),
        _val('Trạng thái cổng'),
        _val('ttvt_ton'),
        _val('chitieu_tg'),
        _val('thời gian tồn thực'),
        _val('giờ còn lại thực'),
        _val('SA'),
        sheet,
        now_iso,  # first_seen (chỉ tác dụng khi INSERT; UPSERT bỏ qua trên UPDATE)
        now_iso,  # last_seen (luôn update)
    )


def _sync_brcd_phieu_to_db():
    """Đọc Vũ trụ tổng hiện tại (Excel), upsert vào brcd_phieu.

    Idempotent. Không raise khi Excel thiếu hoặc DB lock — trả dict với skipped=1.
    Trả: {'synced': N, 'new': M, 'updated': K, 'skipped': 0|1, 'reason': str}
    """
    if not os.path.exists(BRCD_DETAIL_MAIN_FILE):
        return {'synced': 0, 'new': 0, 'updated': 0, 'skipped': 1, 'reason': 'excel_missing'}
    _ensure_brcd_kiemsoat_schema()

    try:
        all_sheets = pd.ExcelFile(BRCD_DETAIL_MAIN_FILE).sheet_names
        team_sheets = [s for s in all_sheets
                       if s.startswith('ToKT_') and not s.endswith('_rut_gon')]
    except Exception:
        return {'synced': 0, 'new': 0, 'updated': 0, 'skipped': 1, 'reason': 'excel_unreadable'}

    incoming = []  # list[tuple[sheet_name, DataFrame]]
    for sheet in team_sheets:
        df = read_excel_sheet_cached(BRCD_DETAIL_MAIN_FILE, sheet)
        cols = [c for c in BRCD_KIEMSOAT_DISPLAY_COLUMNS if c in df.columns]
        df = df[cols].copy()
        df['baohong_id'] = pd.to_numeric(df['baohong_id'], errors='coerce')
        df = df.dropna(subset=['baohong_id']).copy()
        df['baohong_id'] = df['baohong_id'].astype(int)
        if len(df):
            incoming.append((sheet, df))

    if not incoming:
        return {'synced': 0, 'new': 0, 'updated': 0, 'skipped': 0, 'reason': ''}

    all_ids = set()
    for _, df in incoming:
        all_ids.update(df['baohong_id'].astype(int).tolist())

    now_iso = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    try:
        with _brcd_kiemsoat_write_connection() as conn:
            # Tính new vs updated trước bằng 1 query IN
            placeholders = ', '.join('?' for _ in all_ids)
            existing_rows = conn.execute(
                f'SELECT baohong_id FROM brcd_phieu WHERE baohong_id IN ({placeholders})',
                list(all_ids),
            ).fetchall()
            existing_ids = {r['baohong_id'] for r in existing_rows}

            for sheet, df in incoming:
                for _, row in df.iterrows():
                    params = _phieu_row_to_params(row, sheet, now_iso)
                    conn.execute(_BRCD_PHIEU_UPSERT_SQL, params)

        new_count = len(all_ids - existing_ids)
        updated_count = len(all_ids & existing_ids)
        return {
            'synced': len(all_ids),
            'new': new_count,
            'updated': updated_count,
            'skipped': 0,
            'reason': '',
        }
    except sqlite3.OperationalError as e:
        msg = str(e).lower()
        if 'locked' in msg or 'busy' in msg:
            return {'synced': 0, 'new': 0, 'updated': 0, 'skipped': 1, 'reason': 'db_locked'}
        raise


CAU_HINH_TU_DONG_DATE_BINDINGS = [
    {
        'key': 'tong_hop',
        'report_code': 'cau_hinh_tu_dong_cau_hinh_tu_dong_chi_tiet',
        'table_name': 'cau_hinh_tu_dong_cau_hinh_tu_dong_chi_tiet_th_theo_to',
        'order_by': 't."Trung tâm Viễn thông", t."Đội Viễn thông"',
    },
    {
        'key': 'team_summary',
        'report_code': 'cau_hinh_tu_dong_cau_hinh_tu_dong_chi_tiet',
        'table_name': 'cau_hinh_tu_dong_cau_hinh_tu_dong_chi_tiet_th_theo_nvkt',
        'order_by': 't."Trung tâm Viễn thông", t."Đội Viễn thông", t."NVKT"',
    },
]


KPI_DATE_BINDINGS = [
    {
        'key': 'c11',
        'report_code': 'kpi_nvkt_c11_nvktdb_report',
        'table_name': 'kpi_nvkt_c11_nvktdb_report_c11_kpi_nvkt',
        'order_by': 't."đơn vị", t."NVKT"',
    },
    {
        'key': 'c12',
        'report_code': 'kpi_nvkt_c12_nvktdb_report',
        'table_name': 'kpi_nvkt_c12_nvktdb_report_c12_kpi_nvkt',
        'order_by': 't."đơn vị", t."NVKT"',
    },
    {
        'key': 'c13',
        'report_code': 'kpi_nvkt_c13_nvktdb_report',
        'table_name': 'kpi_nvkt_c13_nvktdb_report_c13_kpi_nvkt',
        'order_by': 't."đơn vị", t."NVKT"',
    },
]


TONG_HOP_BSC_KPI_CHAT_LUONG_C_DATE_BINDINGS = [
    {
        'key': 'c1_1',
        'report_code': 'chi_tieu_c_c1_1_report',
        'table_name': 'chi_tieu_c_c1_1_report_th_c1_1',
        'order_by': 't."Đơn vị"',
    },
    {
        'key': 'c1_2',
        'report_code': 'chi_tieu_c_c1_2_report',
        'table_name': 'chi_tieu_c_c1_2_report_th_c1_2',
        'order_by': 't."Đơn vị"',
    },
    {
        'key': 'c1_3',
        'report_code': 'chi_tieu_c_c1_3_report',
        'table_name': 'chi_tieu_c_c1_3_report_th_c1_3',
        'order_by': 't."Đơn vị"',
    },
    {
        'key': 'c1_4',
        'report_code': 'chi_tieu_c_c1_4_report',
        'table_name': 'chi_tieu_c_c1_4_report_th_c1_4',
        'order_by': 't."Đơn vị"',
    },
    {
        'key': 'c1_5',
        'report_code': 'chi_tieu_c_c1_5_report',
        'table_name': 'chi_tieu_c_c1_5_report_th_c1_5',
        'order_by': 't."Đơn vị"',
    },
]


def _current_user():
    username = session.get('username')
    return get_user_by_username(username) if username else None


def _slugify_vi(text):
    value = str(text or '').strip().lower()
    normalized = unicodedata.normalize('NFD', value)
    ascii_text = ''.join(ch for ch in normalized if unicodedata.category(ch) != 'Mn')
    ascii_text = ascii_text.replace('đ', 'd')
    ascii_text = re.sub(r'[^a-z0-9]+', '-', ascii_text)
    return ascii_text.strip('-')


def _team_slug(team_name):
    raw = str(team_name or '').strip()
    if not raw:
        return ''

    prefixes = [
        'Tổ Kỹ thuật Địa bàn ',
        'To Ky thuat Dia ban ',
        'Tổ ',
        'To ',
    ]
    for prefix in prefixes:
        if raw.startswith(prefix):
            raw = raw[len(prefix):]
            break
    return _slugify_vi(raw)


def _build_nvkt_slug(row):
    team_part = _team_slug(row.get('to_doi_hoac_don_vi'))
    name_part = _slugify_vi(row.get('nvkt_hoac_ten_nv'))
    if team_part and name_part:
        return f'{team_part}-{name_part}'
    return name_part or team_part


def _find_nvkt_row_by_slug(nvkt_slug):
    df = load_nvkt_tong_hop_da_nguon_df()
    if df.empty:
        return None

    for _, row in df.iterrows():
        if _build_nvkt_slug(row) == nvkt_slug:
            return row.to_dict()
    return None


def _filter_pending_brcd_sheet(df, _sheet_name):
    if 'lydoton' not in df.columns:
        return build_sheet_payload(df.iloc[0:0])

    filtered_df = df[
        df['lydoton'].isna()
        | (df['lydoton'].astype(str).str.strip() == '')
        | (df['lydoton'].astype(str).str.lower() == 'nan')
    ].copy()
    return build_sheet_payload(filtered_df)


def _pttb_chitiet_to_sheet(df, _sheet_name):
    required_columns = [
        'MA_THUE_BAO', 'LOAIHINH_TB', 'TEN_THUEBAO', 'DIACHI_LAPDAT', 'SO_DT',
        'KIEU_LAPDAT', 'NGAYLAP_HOPDONG', 'NGUOI_GIAO_VIEC', 'TEN_KV', 'NGAYHEN_DEN',
        'NOIDUNG_HEN', 'TG_THICONG_H', 'NHOM_TON', 'LYDOTON', 'GHICHU_TON', 'NGAY_BAO_TON',
    ]
    available_cols = [col for col in required_columns if col in df.columns]
    filtered_df = df[available_cols].copy()

    if 'NGAYLAP_HOPDONG' in filtered_df.columns:
        filtered_df = filtered_df.sort_values(by='NGAYLAP_HOPDONG', ascending=True, na_position='last')

    return build_sheet_payload(filtered_df, datetime_format='%d/%m/%Y %H:%M')


def _collect_png_file_info(directory, path_prefix):
    files = {}
    if not os.path.isdir(directory):
        return files

    for filename in sorted(os.listdir(directory)):
        if not filename.lower().endswith('.png'):
            continue

        file_path = os.path.join(directory, filename)
        if not os.path.isfile(file_path):
            continue

        file_info = build_file_info(file_path)
        file_info['path'] = f'{path_prefix}/{filename}'
        files[filename] = file_info

    return files


def _collect_recursive_png_file_info(directory, path_prefix):
    files = {}
    if not os.path.isdir(directory):
        return files

    for root, _, filenames in os.walk(directory):
        for filename in sorted(filenames):
            if not filename.lower().endswith('.png'):
                continue

            file_path = os.path.join(root, filename)
            relative_path = os.path.relpath(file_path, directory).replace(os.sep, '/')
            file_info = build_file_info(file_path)
            file_info['path'] = f'{path_prefix}/{relative_path}'
            files[relative_path] = file_info

    return files


def _build_cau_hinh_tu_dong_son_tay_payload(date_context):
    if date_context['date_has_data']:
        dataframes = load_many_tables_by_date(CAU_HINH_TU_DONG_DATE_BINDINGS, date_context['selected_date'])
        df = dataframes['tong_hop']
        nvkt_summary = dataframes['team_summary']
    else:
        df = pd.DataFrame()
        nvkt_summary = pd.DataFrame()

    file_info = build_file_info(REPORT_HISTORY_DB_PATH, include_name=True)
    error_counts = pd.DataFrame(columns=['Mã lỗi', 'Số lượng'])
    detail_df = pd.DataFrame()

    total_contracts = int(df['Tổng hợp đồng'].sum()) if 'Tổng hợp đồng' in df.columns else 0
    total_success = int(df['Thành công'].sum()) if 'Thành công' in df.columns else 0
    total_fail = int(df['Thất bại'].sum()) if 'Thất bại' in df.columns else 0
    total_pending = int(df['Chưa có trạng thái'].sum()) if 'Chưa có trạng thái' in df.columns else 0
    total_new = int(df['Lắp mới'].sum()) if 'Lắp mới' in df.columns else 0
    total_replace = int(df['Thay thế'].sum()) if 'Thay thế' in df.columns else 0
    total_wan = int(df['Cấu hình WAN'].sum()) if 'Cấu hình WAN' in df.columns else 0
    total_wifi = int(df['Cấu hình WiFi'].sum()) if 'Cấu hình WiFi' in df.columns else 0

    summary = {
        'tong_so': total_contracts,
        'thanh_cong': total_success,
        'that_bai': total_fail,
        'chua_co_trang_thai': total_pending,
        'lap_moi': total_new,
        'thay_the': total_replace,
        'cau_hinh_wan': total_wan,
        'cau_hinh_wifi': total_wifi,
    }
    summary['ty_le_thanh_cong'] = round(
        (summary['thanh_cong'] / summary['tong_so'] * 100) if summary['tong_so'] else 0,
        2,
    )

    return {
        'file_info': file_info,
        'selected_date': date_context['selected_date'],
        'latest_available_date': date_context['latest_available_date'],
        'available_dates': date_context['available_dates'],
        'date_has_data': date_context['date_has_data'],
        'summary': summary,
        'tong_hop': build_sheet_payload(df),
        'team_summary': build_sheet_payload(nvkt_summary),
        'nvkt_summary': build_sheet_payload(nvkt_summary),
        'error_summary': build_sheet_payload(error_counts),
        'detail': build_sheet_payload(detail_df),
    }


def _normalize_kpi_numeric(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, str):
        text = value.strip().replace('%', '').replace(',', '.')
        if not text:
            return None
        try:
            return float(text)
        except ValueError:
            return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _safe_percent(numerator, denominator):
    numerator_value = _normalize_kpi_numeric(numerator)
    denominator_value = _normalize_kpi_numeric(denominator)
    if numerator_value is None or denominator_value in (None, 0):
        return None
    return round(numerator_value * 100.0 / denominator_value, 2)


def _build_kpi_detail_frames(date_context):
    if not date_context['date_has_data']:
        return pd.DataFrame(), pd.DataFrame()

    source_frames = load_many_tables_by_date(KPI_DATE_BINDINGS, date_context['selected_date'])
    normalized_frames = []

    def append_frame(frame, nhom_chi_tieu, *, chi_so_1, ten_chi_so_1, chi_so_2=None, ten_chi_so_2=None, chi_so_3=None, ten_chi_so_3=None):
        if frame.empty:
            return
        normalized = pd.DataFrame({
            'ngay_du_lieu': date_context['selected_date'],
            'nhom_chi_tieu': nhom_chi_tieu,
            'don_vi': frame.get('đơn vị', pd.Series(dtype='object')).fillna('').astype(str).str.strip(),
            'nvkt': frame.get('NVKT', pd.Series(dtype='object')).fillna('').astype(str).str.strip(),
            'sm1': frame.get('SM1'),
            'sm2': frame.get('SM2'),
            'sm3': frame.get('SM3'),
            'sm4': frame.get('SM4'),
            'sm5': frame.get('SM5') if 'SM5' in frame.columns else None,
            'sm6': frame.get('SM6') if 'SM6' in frame.columns else None,
            'chi_so_1': frame.get(chi_so_1),
            'ten_chi_so_1': ten_chi_so_1,
            'chi_so_2': frame.get(chi_so_2) if chi_so_2 else None,
            'ten_chi_so_2': ten_chi_so_2 or '',
            'chi_so_3': frame.get(chi_so_3) if chi_so_3 else None,
            'ten_chi_so_3': ten_chi_so_3 or '',
            'chi_tieu_bsc': frame.get('Chỉ tiêu BSC'),
        })
        normalized_frames.append(normalized)

    append_frame(
        source_frames['c11'],
        'C11',
        chi_so_1='Tỷ lệ sửa chữa phiếu chất lượng chủ động dịch vụ FiberVNN, MyTV đạt yêu cầu',
        ten_chi_so_1='Tỷ lệ sửa chữa phiếu chất lượng chủ động dịch vụ FiberVNN, MyTV đạt yêu cầu',
        chi_so_2='Tỷ lệ phiếu sửa chữa báo hỏng dịch vụ BRCĐ đúng quy định không tính hẹn',
        ten_chi_so_2='Tỷ lệ phiếu sửa chữa báo hỏng dịch vụ BRCĐ đúng quy định không tính hẹn',
    )
    append_frame(
        source_frames['c12'],
        'C12',
        chi_so_1='Tỷ lệ thuê bao báo hỏng dịch vụ BRCĐ lặp lại',
        ten_chi_so_1='Tỷ lệ thuê bao báo hỏng dịch vụ BRCĐ lặp lại',
        chi_so_2='Tỷ lệ sự cố dịch vụ BRCĐ',
        ten_chi_so_2='Tỷ lệ sự cố dịch vụ BRCĐ',
    )
    append_frame(
        source_frames['c13'],
        'C13',
        chi_so_1='Tỷ lệ sửa chữa dịch vụ kênh TSL hoàn thành đúng thời gian quy định',
        ten_chi_so_1='Tỷ lệ sửa chữa dịch vụ kênh TSL hoàn thành đúng thời gian quy định',
        chi_so_2='Tỷ lệ thuê bao báo hỏng dịch vụ kênh TSL lặp lại',
        ten_chi_so_2='Tỷ lệ thuê bao báo hỏng dịch vụ kênh TSL lặp lại',
        chi_so_3='Tỷ lệ sự cố dịch vụ kênh TSL',
        ten_chi_so_3='Tỷ lệ sự cố dịch vụ kênh TSL',
    )

    normalized_frames = [frame for frame in normalized_frames if not frame.empty]
    if not normalized_frames:
        return pd.DataFrame(), pd.DataFrame()

    detail_df = pd.concat(normalized_frames, ignore_index=True)
    detail_df = detail_df[(detail_df['don_vi'] != '') & (detail_df['nvkt'] != '') & (detail_df['nvkt'] != 'Tổng')].copy()
    if detail_df.empty:
        return pd.DataFrame(), pd.DataFrame()

    for column in ['sm1', 'sm2', 'sm3', 'sm4', 'sm5', 'sm6', 'chi_so_1', 'chi_so_2', 'chi_so_3', 'chi_tieu_bsc']:
        if column in detail_df.columns:
            detail_df[column] = detail_df[column].apply(_normalize_kpi_numeric)

    summary_rows = []
    for (don_vi, nhom_chi_tieu), group in detail_df.groupby(['don_vi', 'nhom_chi_tieu'], dropna=False, sort=True):
        row = {
            'ngay_du_lieu': date_context['selected_date'],
            'nhom_chi_tieu': nhom_chi_tieu,
            'don_vi': don_vi,
            'nvkt': 'Tổng',
            'sm1': group['sm1'].sum(min_count=1),
            'sm2': group['sm2'].sum(min_count=1),
            'sm3': group['sm3'].sum(min_count=1),
            'sm4': group['sm4'].sum(min_count=1),
            'sm5': group['sm5'].sum(min_count=1),
            'sm6': group['sm6'].sum(min_count=1),
            'ten_chi_so_1': group['ten_chi_so_1'].iloc[0],
            'ten_chi_so_2': group['ten_chi_so_2'].iloc[0],
            'ten_chi_so_3': group['ten_chi_so_3'].iloc[0],
            'chi_tieu_bsc': round(group['chi_tieu_bsc'].mean(), 2) if group['chi_tieu_bsc'].notna().any() else None,
        }
        row['chi_so_1'] = _safe_percent(row['sm1'], row['sm2'])
        row['chi_so_2'] = _safe_percent(row['sm3'], row['sm4'])
        row['chi_so_3'] = _safe_percent(row['sm5'], row['sm6']) if nhom_chi_tieu == 'C13' else None
        summary_rows.append(row)

    summary_df = pd.DataFrame(summary_rows)
    detail_df = detail_df.sort_values(by=['don_vi', 'nvkt', 'nhom_chi_tieu'], na_position='last').reset_index(drop=True)
    summary_df = summary_df.sort_values(by=['don_vi', 'nhom_chi_tieu'], na_position='last').reset_index(drop=True)
    return summary_df, detail_df


def _load_tong_hop_bsc_kpi_rows():
    return load_tong_hop_bsc_kpi_rows(UNIT_NAME)


def _load_tong_hop_bsc_kpi_chat_luong_c_frames(selected_date):
    if not selected_date:
        return {binding['key']: pd.DataFrame() for binding in TONG_HOP_BSC_KPI_CHAT_LUONG_C_DATE_BINDINGS}
    return load_many_tables_by_date(TONG_HOP_BSC_KPI_CHAT_LUONG_C_DATE_BINDINGS, selected_date)


def _build_tong_hop_bsc_kpi_chat_luong_c_sheet(date_context, frames):
    column_specs = [
        ('c1_1', 'Chỉ tiêu BSC', 'C1.1 - Chỉ tiêu BSC'),
        ('c1_2', 'Chỉ tiêu BSC', 'C1.2 - Chỉ tiêu BSC'),
        ('c1_3', 'Chỉ tiêu BSC', 'C1.3 - Chỉ tiêu BSC'),
        ('c1_4', 'Điểm BSC', 'C1.4 - Điểm BSC'),
        ('c1_5', 'Tổng - Điểm BSC', 'C1.5 - Tổng - Điểm BSC'),
    ]

    merged_df = None
    for key, source_column, target_column in column_specs:
        frame = frames.get(key, pd.DataFrame())
        if frame.empty or 'Đơn vị' not in frame.columns or source_column not in frame.columns:
            current_df = pd.DataFrame(columns=['Đơn vị', target_column])
        else:
            current_df = frame[['Đơn vị', source_column]].copy()
            current_df = current_df.rename(columns={source_column: target_column})
            current_df = current_df.drop_duplicates(subset=['Đơn vị'], keep='first')

        merged_df = current_df if merged_df is None else merged_df.merge(current_df, on='Đơn vị', how='outer')

    if merged_df is None:
        merged_df = pd.DataFrame(columns=['Đơn vị', *[spec[2] for spec in column_specs]])

    merged_df = merged_df.fillna(pd.NA)
    merged_df['_sort_order'] = merged_df['Đơn vị'].apply(lambda value: 1 if str(value).strip().lower() == 'tổng' else 0)
    merged_df = merged_df.sort_values(by=['_sort_order', 'Đơn vị'], na_position='last').drop(columns=['_sort_order'])
    return build_sheet_payload(merged_df.reset_index(drop=True))


def _build_tong_hop_bsc_kpi_payload(date_context=None):
    file_info = build_file_info(REPORT_HISTORY_DB_PATH, include_name=True)
    if date_context is None:
        date_context = resolve_date_context(
            None,
            [binding['report_code'] for binding in TONG_HOP_BSC_KPI_CHAT_LUONG_C_DATE_BINDINGS],
        )

    if date_context['date_has_data']:
        chat_luong_frames = _load_tong_hop_bsc_kpi_chat_luong_c_frames(date_context['selected_date'])
    else:
        chat_luong_frames = {binding['key']: pd.DataFrame() for binding in TONG_HOP_BSC_KPI_CHAT_LUONG_C_DATE_BINDINGS}
    chat_luong_c_sheet = _build_tong_hop_bsc_kpi_chat_luong_c_sheet(date_context, chat_luong_frames)

    return {
        'file_info': file_info,
        'selected_date': date_context['selected_date'],
        'latest_available_date': date_context['latest_available_date'],
        'available_dates': date_context['available_dates'],
        'date_has_data': date_context['date_has_data'],
        'chat_luong_c': chat_luong_c_sheet,
    }


@operations_bp.route('/brcd')
@login_required
def page_brcd():
    ton_dv_data = []
    ton_nvkt_data = []
    ton_5doi_data = []
    ton_dv_columns = []
    ton_nvkt_columns = []
    ton_5doi_columns = []

    try:
        df = read_excel_sheet_cached(BRCD_SUMMARY_FILE, 'ton_dv_theo_to')
        ton_dv_columns = df.columns.tolist()
        ton_dv_data = df.fillna('').to_dict('records')

        df_nvkt = read_excel_sheet_cached(BRCD_SUMMARY_FILE, 'ton_dv_theo_nvkt')
        ton_nvkt_columns = df_nvkt.columns.tolist()
        ton_nvkt_data = df_nvkt.fillna('').to_dict('records')

        df_5doi = read_excel_sheet_cached(BRCD_SUMMARY_FILE, 'tong_hop_5doi')
        ton_5doi_columns = df_5doi.columns.tolist()
        ton_5doi_data = df_5doi.fillna('').to_dict('records')
    except Exception as exc:
        current_app.logger.warning('Khong the doc du lieu tong hop BRCD: %s', exc)

    return render_template(
        'pages/brcd.html',
        current_user=_current_user(),
        active_page='brcd',
        ton_dv_data=ton_dv_data,
        ton_nvkt_data=ton_nvkt_data,
        ton_dv_columns=ton_dv_columns,
        ton_nvkt_columns=ton_nvkt_columns,
        ton_5doi_data=ton_5doi_data,
        ton_5doi_columns=ton_5doi_columns,
    )


@operations_bp.route('/download/bc-brcd')
@login_required
def download_bc_brcd():
    return safe_file_response(
        BRCD_SUMMARY_FILE,
        as_attachment=True,
        download_name='bc_BRCD.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )


@operations_bp.route('/api/brcd-kiemsoat/luu', methods=['POST'])
@login_required
def api_brcd_kiemsoat_luu():
    payload = request.get_json(silent=True) or {}
    try:
        baohong_id = int(payload.get('baohong_id'))
    except (TypeError, ValueError):
        return jsonify({'ok': False, 'error': 'baohong_id không hợp lệ'}), 400

    noi_dung = str(payload.get('noi_dung') or '').strip()
    if len(noi_dung) > BRCD_KIEMSOAT_NOI_DUNG_MAX:
        return jsonify({'ok': False, 'error': 'Nội dung không được vượt quá 2000 ký tự'}), 400

    ma_tb = str(payload.get('ma_tb') or '')
    doi_vt = str(payload.get('doi_vt') or '')
    nvkt = str(payload.get('nvkt') or '')
    username = session.get('username') or ''
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    _ensure_brcd_kiemsoat_schema()
    try:
        with _brcd_kiemsoat_write_connection() as conn:
            if not noi_dung:
                conn.execute('DELETE FROM brcd_kiemsoat WHERE baohong_id = ?', (baohong_id,))
                return jsonify({'ok': True, 'baohong_id': baohong_id, 'noi_dung': '', 'nguoi_nhap': ''})

            existing = conn.execute(
                'SELECT thoi_diem_nhap FROM brcd_kiemsoat WHERE baohong_id = ?',
                (baohong_id,),
            ).fetchone()
            thoi_diem_nhap = existing['thoi_diem_nhap'] if existing else now
            conn.execute(
                '''
                INSERT INTO brcd_kiemsoat
                    (baohong_id, ma_tb, doi_vt, nvkt, noi_dung_kiem_soat,
                     nguoi_nhap, thoi_diem_nhap, thoi_diem_cap_nhat)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(baohong_id) DO UPDATE SET
                    ma_tb = excluded.ma_tb,
                    doi_vt = excluded.doi_vt,
                    nvkt = excluded.nvkt,
                    noi_dung_kiem_soat = excluded.noi_dung_kiem_soat,
                    nguoi_nhap = excluded.nguoi_nhap,
                    thoi_diem_cap_nhat = excluded.thoi_diem_cap_nhat
                ''',
                (baohong_id, ma_tb, doi_vt, nvkt, noi_dung, username, thoi_diem_nhap, now),
            )
    except sqlite3.Error as exc:
        return jsonify({'ok': False, 'error': str(exc)}), 500

    return jsonify({'ok': True, 'baohong_id': baohong_id, 'noi_dung': noi_dung, 'nguoi_nhap': username})


def _load_brcd_kiemsoat_df():
    """Đọc sheet đầy đủ ToKT_<doi> + join annotation kiểm soát. Trả DataFrame hoặc None."""
    if not os.path.exists(BRCD_DETAIL_MAIN_FILE):
        return None
    _ensure_brcd_kiemsoat_schema()
    # Snapshot Vũ trụ tổng hiện tại vào brcd_phieu (best-effort, không raise).
    try:
        _sync_brcd_phieu_to_db()
    except Exception:
        current_app.logger.warning('brcd_phieu sync thất bại trong _load', exc_info=True)

    all_sheets = pd.ExcelFile(BRCD_DETAIL_MAIN_FILE).sheet_names
    team_sheets = [s for s in all_sheets if s.startswith('ToKT_') and not s.endswith('_rut_gon')]

    frames = []
    for sheet in team_sheets:
        df = read_excel_sheet_cached(BRCD_DETAIL_MAIN_FILE, sheet)
        cols = [c for c in BRCD_KIEMSOAT_DISPLAY_COLUMNS if c in df.columns]
        df = df[cols].copy()
        df['_sheet'] = sheet
        df['baohong_id'] = pd.to_numeric(df['baohong_id'], errors='coerce')
        frames.append(df)

    if not frames:
        return None

    combined = pd.concat(frames, ignore_index=True).dropna(subset=['baohong_id']).copy()
    combined['baohong_id'] = combined['baohong_id'].astype(int)

    kiemsoat_map = get_brcd_kiemsoat_map(combined['baohong_id'].unique().tolist())
    ks_keys = set(kiemsoat_map.keys())
    combined['kiemsoat_noi_dung'] = combined['baohong_id'].map(
        lambda i: kiemsoat_map.get(i, {}).get('noi_dung_kiem_soat', '')
    )
    combined['kiemsoat_nguoi_nhap'] = combined['baohong_id'].map(
        lambda i: kiemsoat_map.get(i, {}).get('nguoi_nhap', '')
    )
    combined['kiemsoat_thoi_diem'] = combined['baohong_id'].map(
        lambda i: kiemsoat_map.get(i, {}).get('thoi_diem_cap_nhat')
        or kiemsoat_map.get(i, {}).get('thoi_diem_nhap', '')
    )
    combined['kiemsoat_da_nhap'] = combined['baohong_id'].isin(ks_keys)

    for col in ('giờ còn lại thực', 'thời gian tồn thực', 'chitieu_tg'):
        if col in combined.columns:
            combined[col] = pd.to_numeric(combined[col], errors='coerce')

    return combined


def _apply_brcd_kiemsoat_filters(df, args):
    """Lọc DataFrame tồn+kiemsoat theo query args. df đã numeric."""
    filtered = df
    doi = args.get('doi')
    if doi and 'DOI_VT' in filtered.columns:
        filtered = filtered[filtered['DOI_VT'].astype(str) == doi]

    loaihinh = args.get('loaihinh')
    if loaihinh and 'LOAIHINH_TB' in filtered.columns:
        filtered = filtered[filtered['LOAIHINH_TB'].astype(str) == loaihinh]

    nhom = args.get('nhom')
    if nhom and 'giờ còn lại thực' in filtered.columns:
        remaining = filtered['giờ còn lại thực']
        if nhom == 'qua_gio':
            filtered = filtered[remaining.fillna(0) <= 0]
        elif nhom == 'trong_gio':
            filtered = filtered[remaining.fillna(0) > 0]

    trangthai = args.get('trangthai')
    if trangthai == 'da':
        filtered = filtered[filtered['kiemsoat_da_nhap']]
    elif trangthai == 'chua':
        filtered = filtered[~filtered['kiemsoat_da_nhap']]

    return filtered


def _compute_lich_su(filtered_df, args):
    """Tính 'phiếu đã rời tồn' từ snapshot brcd_phieu trong khoảng thời gian.

    "Rời tồn" = có trong brcd_phieu (trong khoảng) nhưng KHÔNG có trong
    current universe (filtered_df). Filter doi/loaihinh áp dụng cho snapshot.
    """
    khoang = (args.get('khoang') or 'thang_nay').strip()
    today = date.today()
    if khoang == 'tuan_nay':
        tu_ngay = today - timedelta(days=today.weekday())  # thứ 2
        den_ngay = today
    elif khoang == 'nam_nay':
        tu_ngay = today.replace(month=1, day=1)
        den_ngay = today
    elif khoang == 'tat_ca':
        tu_ngay = date(1970, 1, 1)
        den_ngay = today
    else:  # thang_nay (default)
        tu_ngay = today.replace(day=1)
        den_ngay = today

    # Current universe IDs (sau filter doi/loaihinh/nhom)
    current_ids = set()
    if filtered_df is not None and 'baohong_id' in filtered_df.columns:
        current_ids = set(filtered_df['baohong_id'].astype(int).tolist())

    sql = ("SELECT baohong_id FROM brcd_phieu "
           "WHERE DATE(last_seen) >= ? AND DATE(last_seen) <= ?")
    params = [tu_ngay.isoformat(), den_ngay.isoformat()]

    doi = args.get('doi')
    if doi:
        sql += " AND doi_vt = ?"
        params.append(doi)
    loaihinh = args.get('loaihinh')
    if loaihinh:
        sql += " AND loaihinh_tb = ?"
        params.append(loaihinh)

    _ensure_brcd_kiemsoat_schema()
    with _brcd_kiemsoat_read_connection() as conn:
        snap_rows = conn.execute(sql, params).fetchall()
        snap_ids = {r['baohong_id'] for r in snap_rows}

        ks_ids = set()
        if snap_ids:
            snap_list = list(snap_ids)
            for i in range(0, len(snap_list), 500):
                batch = snap_list[i:i+500]
                placeholders = ', '.join('?' for _ in batch)
                ks_rows = conn.execute(
                    f"SELECT baohong_id FROM brcd_kiemsoat "
                    f"WHERE baohong_id IN ({placeholders}) "
                    f"AND COALESCE(noi_dung_kiem_soat, '') != ''",
                    batch,
                ).fetchall()
                ks_ids.update(r['baohong_id'] for r in ks_rows)

    roi_ids = snap_ids - current_ids
    roi_da_ks = len(roi_ids & ks_ids)
    roi_chua_ks = len(roi_ids - ks_ids)

    return {
        'tu_ngay': tu_ngay.isoformat(),
        'den_ngay': den_ngay.isoformat(),
        'roi_da_ks': roi_da_ks,
        'roi_chua_ks': roi_chua_ks,
    }


@operations_bp.route('/api/brcd-kiemsoat/detail')
@login_required
def api_brcd_kiemsoat_detail():
    df = _load_brcd_kiemsoat_df()
    if df is None:
        return jsonify({'error': 'File Excel chi tiết BRCD không tồn tại'}), 404

    sheets = {}
    for sheet, group in df.groupby('_sheet'):
        group = group.drop(columns=['_sheet'])
        sheets[str(sheet)] = build_sheet_payload(group)

    return jsonify({'sheets': sheets, 'file_info': build_file_info(BRCD_DETAIL_MAIN_FILE)})


@operations_bp.route('/api/brcd-kiemsoat/thongke')
@login_required
def api_brcd_kiemsoat_thongke():
    df = _load_brcd_kiemsoat_df()
    if df is None:
        return jsonify({'error': 'File Excel chi tiết BRCD không tồn tại'}), 404

    filtered = _apply_brcd_kiemsoat_filters(df, request.args)

    total = len(filtered)
    da = int(filtered['kiemsoat_da_nhap'].sum()) if total else 0
    chua = total - da
    ty_le = round(da * 100.0 / total, 1) if total else 0.0

    def _agg(group_col):
        if group_col not in filtered.columns:
            return []
        result = []
        for key, group in filtered.dropna(subset=[group_col]).groupby(group_col):
            g_total = len(group)
            g_da = int(group['kiemsoat_da_nhap'].sum())
            result.append({
                group_col: '' if key is None else str(key),
                'total': g_total,
                'da_kiem_soat': g_da,
                'chua': g_total - g_da,
                'ty_le': round(g_da * 100.0 / g_total, 1) if g_total else 0.0,
            })
        result.sort(key=lambda r: r['total'], reverse=True)
        return result

    chi_tiet = serialize_dataframe(filtered.drop(columns=['_sheet'])).to_dict('records')

    return jsonify({
        'summary': {
            'total': total,
            'da_kiem_soat': da,
            'chua': chua,
            'ty_le': ty_le,
        },
        'by_doi': _agg('DOI_VT'),
        'by_nvkt': _agg('NVKT'),
        'chi_tiet': chi_tiet,
        'lich_su': _compute_lich_su(filtered, request.args),
    })


@operations_bp.route('/download/brcd-kiemsoat-report')
@login_required
def download_brcd_kiemsoat_report():
    df = _load_brcd_kiemsoat_df()
    if df is None:
        return jsonify({'error': 'File Excel chi tiết BRCD không tồn tại'}), 404

    filtered = _apply_brcd_kiemsoat_filters(df, request.args)
    export_df = serialize_dataframe(filtered.drop(columns=['_sheet']))

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        export_df.to_excel(writer, index=False, sheet_name='Kiem_soat_BRCD')
    output.seek(0)

    download_name = f'brcd_kiemsoat_{datetime.now():%Y%m%d_%H%M}.xlsx'
    return send_file(
        output,
        as_attachment=True,
        download_name=download_name,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )


@operations_bp.route('/pttb')
@login_required
def page_pttb():
    return render_template('pages/pttb.html', current_user=_current_user(), active_page='pttb')


@operations_bp.route('/cau-hinh-tu-dong')
@login_required
def page_cau_hinh_tu_dong():
    return render_template(
        'pages/cau_hinh_tu_dong.html',
        current_user=_current_user(),
        active_page='cau_hinh_tu_dong',
    )


@operations_bp.route('/tong-hop-bsc-kpi')
@login_required
def page_tong_hop_bsc_kpi():
    return render_template(
        'pages/tong_hop_bsc_kpi.html',
        current_user=_current_user(),
        active_page='tong_hop_bsc_kpi',
    )


@operations_bp.route('/thuctang')
@login_required
def page_thuctang():
    return render_template('pages/thuctang.html', current_user=_current_user(), active_page='thuctang')


@operations_bp.route('/kpi')
def page_kpi():
    return render_template('pages/kpi.html', current_user=_current_user(), active_page='kpi')


@operations_bp.route('/kpi-nvkt-bchn')
def page_kpi_nvkt_bchn():
    return render_template('pages/kpi_nvkt_bchn.html', current_user=_current_user(), active_page='kpi_nvkt_bchn')


@operations_bp.route('/tong-hop-cap-to')
@login_required
def page_tong_hop_cap_to():
    return render_template('pages/tong_hop_cap_to.html', current_user=_current_user(), active_page='tong_hop_cap_to')


@operations_bp.route('/bsc-kpi-cac-to')
@login_required
def page_bsc_kpi_cac_to():
    return render_template('pages/bsc_kpi_cac_to.html', current_user=_current_user(), active_page='bsc_kpi_cac_to')


@operations_bp.route('/api/kpi-nvkt-bchn-data')
def get_kpi_nvkt_bchn_data():
    try:
        df = load_nvkt_tong_hop_da_nguon_df()
        if df.empty:
            return jsonify({
                'file_modified': build_file_info(current_app.config['REPORT_HISTORY_DB_PATH'])['modified'],
                'sheet': {'columns': [], 'data': []},
            })

        return jsonify({
            'file_modified': build_file_info(current_app.config['REPORT_HISTORY_DB_PATH'])['modified'],
            'sheet': build_sheet_payload(df),
        })
    except Exception as exc:
        current_app.logger.exception('Khong the tai du lieu KPI BCHN tu SQLite: %s', exc)
        return jsonify({'error': f'Lỗi khi đọc dữ liệu KPI BCHN từ SQLite: {exc}'}), 500


@operations_bp.route('/api/tong-hop-cap-to-data')
@login_required
def get_tong_hop_cap_to_data():
    try:
        df = load_don_vi_tong_hop_da_nguon_df()
        return jsonify({
            'file_modified': build_file_info(current_app.config['REPORT_HISTORY_DB_PATH'])['modified'],
            'sheet': build_sheet_payload(df),
        })
    except Exception as exc:
        current_app.logger.exception('Khong the tai du lieu tong hop cap to tu SQLite: %s', exc)
        return jsonify({'error': f'Lỗi khi đọc dữ liệu tổng hợp cấp tổ từ SQLite: {exc}'}), 500


@operations_bp.route('/api/bsc-kpi-cac-to-data')
@login_required
def get_bsc_kpi_cac_to_data():
    try:
        df = load_bsc_kpi_cac_to_df()
        return jsonify({
            'file_modified': build_file_info(current_app.config['REPORT_HISTORY_DB_PATH'])['modified'],
            'sheet': build_sheet_payload(df),
        })
    except Exception as exc:
        current_app.logger.exception('Khong the tai du lieu bsc kpi cac to tu SQLite: %s', exc)
        return jsonify({'error': f'Lỗi khi đọc dữ liệu BSC KPI các tổ từ SQLite: {exc}'}), 500


@operations_bp.route('/api/nvkt-chi-tiet/<path:nvkt_slug>')
@login_required
def get_nvkt_chi_tiet_data(nvkt_slug):
    try:
        row = _find_nvkt_row_by_slug(nvkt_slug)
        if not row:
            return jsonify({'error': f'Không tìm thấy NVKT với slug: {nvkt_slug}'}), 404

        df = load_nvkt_tong_hop_da_nguon_by_name_df(row['nvkt_hoac_ten_nv'])
        if df.empty:
            return jsonify({'error': f'Không tìm thấy NVKT: {row["nvkt_hoac_ten_nv"]}'}), 404

        return jsonify({
            'file_modified': build_file_info(current_app.config['REPORT_HISTORY_DB_PATH'])['modified'],
            'sheet': build_sheet_payload(df),
        })
    except Exception as exc:
        current_app.logger.exception('Khong the tai du lieu chi tiet NVKT slug %s: %s', nvkt_slug, exc)
        return jsonify({'error': f'Lỗi khi đọc dữ liệu NVKT từ SQLite: {exc}'}), 500


@operations_bp.route('/api/tong-hop-bsc-kpi')
def get_tong_hop_bsc_kpi_data():
    try:
        date_context = resolve_date_context(
            request.args.get('date'),
            [binding['report_code'] for binding in TONG_HOP_BSC_KPI_CHAT_LUONG_C_DATE_BINDINGS],
        )
        return jsonify(_build_tong_hop_bsc_kpi_payload(date_context))
    except FileNotFoundError as exc:
        return jsonify({'error': str(exc)}), 404
    except sqlite3.Error as exc:
        current_app.logger.exception('Khong the doc report_history.db: %s', exc)
        return jsonify({'error': f'Lỗi khi đọc SQLite: {exc}'}), 500
    except Exception as exc:
        current_app.logger.exception('Khong the tao payload tong hop BSC KPI: %s', exc)
        return jsonify({'error': f'Lỗi khi tải dashboard tổng hợp BSC KPI: {exc}'}), 500


@operations_bp.route('/chart/<filename>')
def serve_chart(filename):
    return safe_directory_response(os.path.join(BASE_DATA_PATH, 'chart'), filename, no_cache=True)


@operations_bp.route('/image/<filename>')
def serve_image(filename):
    return safe_directory_response(os.path.join(BASE_DATA_PATH, 'image'), filename, no_cache=True)


@operations_bp.route('/pttb_3_to/<filename>')
def serve_pttb_3to_chart(filename):
    return safe_directory_response(os.path.join(BASE_DATA_PATH, 'pttb_3_to'), filename, no_cache=True)


@operations_bp.route('/<path:nvkt_slug>')
@login_required
def page_nvkt_chi_tiet(nvkt_slug):
    row = _find_nvkt_row_by_slug(nvkt_slug)
    if not row:
        abort(404)

    return render_template(
        'pages/nvkt_chi_tiet.html',
        current_user=_current_user(),
        active_page='kpi_nvkt_bchn',
        nvkt_name=row['nvkt_hoac_ten_nv'],
        nvkt_slug=nvkt_slug,
    )


@operations_bp.route('/baocaohanoi/chart/<path:filepath>')
def serve_baocaohanoi_chart(filepath):
    return safe_directory_response(os.path.join(BAOCAO_HANOI_PATH, 'chart'), filepath, no_cache=True)


@operations_bp.route('/api/file-info')
def get_file_info():
    return jsonify({
        'charts': _collect_png_file_info(os.path.join(BASE_DATA_PATH, 'chart'), '/chart'),
        'images': _collect_png_file_info(os.path.join(BASE_DATA_PATH, 'image'), '/image'),
        'baocaohanoi_charts': _collect_recursive_png_file_info(
            os.path.join(BAOCAO_HANOI_PATH, 'chart'),
            '/baocaohanoi/chart',
        ),
    })


@operations_bp.route('/api/excel-data')
def get_excel_data():
    if not os.path.exists(BRCD_DETAIL_OFF_FILE):
        return jsonify({'error': 'File Excel không tồn tại'}), 404

    try:
        return jsonify(build_multi_sheet_payload(BRCD_DETAIL_OFF_FILE))
    except Exception as exc:
        return jsonify({'error': f'Lỗi khi đọc file Excel: {exc}'}), 500


@operations_bp.route('/api/excel-data-main')
def get_excel_data_main():
    if not os.path.exists(BRCD_DETAIL_MAIN_FILE):
        return jsonify({'error': 'File Excel không tồn tại'}), 404

    try:
        return jsonify(build_multi_sheet_payload(BRCD_DETAIL_MAIN_FILE))
    except Exception as exc:
        return jsonify({'error': f'Lỗi khi đọc file Excel: {exc}'}), 500


@operations_bp.route('/download/excel')
def download_excel():
    return safe_file_response(
        BRCD_DETAIL_OFF_FILE,
        as_attachment=True,
        download_name='ChiTietBrcd5Doi_OFF.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )


@operations_bp.route('/api/excel-data-pending')
def get_excel_data_pending():
    if not os.path.exists(BRCD_DETAIL_MAIN_FILE):
        return jsonify({'error': 'File Excel không tồn tại'}), 404

    try:
        return jsonify(build_multi_sheet_payload(BRCD_DETAIL_MAIN_FILE, transform=_filter_pending_brcd_sheet))
    except Exception as exc:
        return jsonify({'error': f'Lỗi khi đọc file Excel: {exc}'}), 500


@operations_bp.route('/download/excel-main')
def download_excel_main():
    return safe_file_response(
        BRCD_DETAIL_MAIN_FILE,
        as_attachment=True,
        download_name='ChiTietBrcd5Doi.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )


@operations_bp.route('/api/pttb-data-summary')
def get_pttb_data_summary():
    if not os.path.exists(PTTB_SUMMARY_FILE):
        return jsonify({'error': 'File Excel PTTB không tồn tại'}), 404

    try:
        return jsonify(
            build_multi_sheet_payload(
                PTTB_SUMMARY_FILE,
                sheet_names=['TK_TongHop_DonGian', 'tong_hop_5doi', 'TK_TongHop_TrangThai'],
                include_sheet_errors=True,
            )
        )
    except Exception as exc:
        return jsonify({'error': f'Lỗi khi đọc file Excel PTTB: {exc}'}), 500


@operations_bp.route('/api/pttb-data-detail')
def get_pttb_data_detail():
    if not os.path.exists(PTTB_SUMMARY_FILE):
        return jsonify({'error': 'File Excel PTTB không tồn tại'}), 404

    try:
        return jsonify(
            build_multi_sheet_payload(
                PTTB_SUMMARY_FILE,
                sheet_names=['tong_hop_dia_ban', 'TK_ToKT_SonTay', 'TK_ToKT_SuoiHai', 'TK_ToKT_QuangOai'],
                include_sheet_errors=True,
            )
        )
    except Exception as exc:
        return jsonify({'error': f'Lỗi khi đọc file Excel PTTB: {exc}'}), 500


@operations_bp.route('/api/pttb-data-pending')
def get_pttb_data_pending():
    if not os.path.exists(PTTB_SUMMARY_FILE):
        return jsonify({'error': 'File Excel PTTB không tồn tại'}), 404

    try:
        return jsonify(
            build_multi_sheet_payload(
                PTTB_SUMMARY_FILE,
                sheet_names=['chua_co_lydoton', 'phieu_qua_gio'],
                include_sheet_errors=True,
            )
        )
    except Exception as exc:
        return jsonify({'error': f'Lỗi khi đọc file Excel PTTB: {exc}'}), 500


@operations_bp.route('/api/pttb-data-chitiet-to')
def get_pttb_data_chitiet_to():
    if not os.path.exists(PTTB_SUMMARY_FILE):
        return jsonify({'error': 'File Excel PTTB không tồn tại'}), 404

    try:
        return jsonify(
            build_multi_sheet_payload(
                PTTB_SUMMARY_FILE,
                sheet_names=['ToKT_SonTay', 'ToKT_SuoiHai', 'ToKT_QuangOai', 'ToKT_PhucTho'],
                include_sheet_errors=True,
                transform=_pttb_chitiet_to_sheet,
            )
        )
    except Exception as exc:
        return jsonify({'error': f'Lỗi khi đọc file Excel PTTB: {exc}'}), 500


@operations_bp.route('/api/cau-hinh-tu-dong/son-tay')
@login_required
def get_cau_hinh_tu_dong_son_tay():
    try:
        date_context = resolve_date_context(
            request.args.get('date'),
            [binding['report_code'] for binding in CAU_HINH_TU_DONG_DATE_BINDINGS],
        )
        return jsonify(_build_cau_hinh_tu_dong_son_tay_payload(date_context))
    except FileNotFoundError as exc:
        return jsonify({'error': str(exc)}), 404
    except Exception as exc:
        return jsonify({'error': f'Lỗi khi đọc dữ liệu cấu hình tự động: {exc}'}), 500


@operations_bp.route('/api/kpi-data')
def get_kpi_data():
    try:
        date_context = resolve_date_context(
            request.args.get('date'),
            [binding['report_code'] for binding in KPI_DATE_BINDINGS],
        )
        summary_df, detail_df = _build_kpi_detail_frames(date_context)
        file_info = build_file_info(current_app.config['REPORT_HISTORY_DB_PATH'])
        return jsonify({
            'file_info': {
                'tomtat_modified': file_info['modified'],
                'chitiet_modified': file_info['modified'],
            },
            'selected_date': date_context['selected_date'],
            'latest_available_date': date_context['latest_available_date'],
            'available_dates': date_context['available_dates'],
            'date_has_data': date_context['date_has_data'],
            'sheets': {
                'tomtat': build_sheet_payload(summary_df),
                'chitiet': build_sheet_payload(detail_df),
            },
        })
    except Exception as exc:
        current_app.logger.exception('Khong the tai du lieu KPI tu SQLite: %s', exc)
        return jsonify({'error': f'Lỗi khi đọc dữ liệu KPI từ SQLite: {exc}'}), 500


@operations_bp.route('/download/excel-pttb')
def download_excel_pttb():
    return safe_file_response(
        PTTB_SUMMARY_FILE,
        as_attachment=True,
        download_name='BaoCaoPTTB.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )
