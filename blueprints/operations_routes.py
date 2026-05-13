import os
import re
import sqlite3
import unicodedata

import pandas as pd
from flask import Blueprint, abort, current_app, jsonify, render_template, request, session

from app_helpers import (
    build_file_info,
    build_multi_sheet_payload,
    build_sheet_payload,
    read_excel_sheet_cached,
    safe_directory_response,
    safe_file_response,
)
from auth import get_user_by_username, login_required
from config import (
    BASE_DATA_PATH,
    BAOCAO_HANOI_KPI_DIR,
    BAOCAO_HANOI_PATH,
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
