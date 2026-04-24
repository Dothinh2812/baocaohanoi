import os
from collections import OrderedDict
from datetime import datetime

import pandas as pd
from flask import Blueprint, current_app, flash, jsonify, redirect, render_template, request, session, url_for

from app_helpers import build_sheet_payload, get_file_metadata, read_excel_sheet_cached, safe_file_response
from auth import get_user_by_username, login_required
from config import (
    INVENTORY_COMMON_GOOD_FILE,
    INVENTORY_PROCESSED_FILE,
    INVENTORY_SOURCE_FILE,
    INVENTORY_TEAM_CONFIGS,
    UNIT_NAME,
)
from repositories import (
    load_many_tables_by_date,
    load_xac_minh_chi_tiet_df,
    resolve_date_context,
)


inventory_bp = Blueprint('inventory', __name__)


TAM_DUNG_KHOI_PHUC_DATE_BINDINGS = [
    {
        'key': 'v_tam_dung_khoi_phuc_dich_vu_chi_tiet_combined_tong_hop_theo_to',
        'report_code': 'tam_dung_khoi_phuc_dich_vu_tam_dung_khoi_phuc_dich_vu_chi_tiet_combined',
        'table_name': 'tam_dung_khoi_phuc_dich_vu_tam_dung_khoi_phuc_dich_vu_chi_tiet_combined_tong_hop_theo_to',
        'order_by': 't."TTVT", t."DOIVT"',
    },
    {
        'key': 'v_tam_dung_khoi_phuc_dich_vu_chi_tiet_combined_tong_hop_theo_nvkt',
        'report_code': 'tam_dung_khoi_phuc_dich_vu_tam_dung_khoi_phuc_dich_vu_chi_tiet_combined',
        'table_name': 'tam_dung_khoi_phuc_dich_vu_tam_dung_khoi_phuc_dich_vu_chi_tiet_combined_tong_hop_theo_nvkt',
        'order_by': 't."TTVT", t."DOIVT", t."NVKT"',
    },
]


def _current_user():
    return get_user_by_username(session['username'])


def _normalize_stt(value):
    if pd.isna(value):
        return ''
    text = str(value).strip()
    if not text:
        return ''
    try:
        number = float(text)
        if number.is_integer():
            return str(int(number))
    except Exception:
        pass
    return text


def _format_money_value(value):
    if pd.isna(value):
        return ''
    if isinstance(value, (int, float)):
        if float(value).is_integer():
            return f"{int(value):,}".replace(',', '.')
        formatted = f"{float(value):,.2f}"
        return formatted.replace(',', 'X').replace('.', ',').replace('X', '.')
    text = str(value).strip()
    return text or ''


def _select_columns(df, prefix_cols, body_cols):
    selected = [col for col in prefix_cols if col in df.columns]
    selected.extend(col for col in body_cols if col in df.columns and col not in selected)
    return df[selected].copy()


def _quantity_columns(df):
    return [col for col in df.columns if '(SL)' in str(col)]


def _money_columns(df):
    return [col for col in df.columns if 'Tổng' in str(col) and 'thành tiền' in str(col)]


def _apply_common_formatting(df):
    formatted = df.copy()
    for col in formatted.columns:
        if pd.api.types.is_datetime64_any_dtype(formatted[col]):
            formatted[col] = formatted[col].dt.strftime('%Y-%m-%d %H:%M:%S')
    return formatted


def _move_columns_to_front(df, columns):
    front = [col for col in columns if col in df.columns]
    remaining = [col for col in df.columns if col not in front]
    return df[front + remaining].copy()


def _team_inventory_payload(config):
    base_file = config['base_file']
    stat_sheet = 'thong-ke-theo-vat-tu'

    df_team = read_excel_sheet_cached(base_file, config['sheet_name'])
    df_team = _apply_common_formatting(
        _select_columns(df_team, ['Tên Kho'], _quantity_columns(df_team))
    )

    df_stats = read_excel_sheet_cached(base_file, stat_sheet)
    df_stats = _apply_common_formatting(
        _select_columns(df_stats, ['Chủng loại vật tư'], _quantity_columns(df_stats))
    )

    return {
        'file_info': get_file_metadata(base_file),
        'sheets': {
            'ton-kho-nvkt': build_sheet_payload(df_team),
            'ton-theo-loai': build_sheet_payload(df_stats),
        },
    }


def _team_good_inventory_payload(config):
    df = read_excel_sheet_cached(config['good_file'], config['sheet_name'])
    formatted = _apply_common_formatting(
        _select_columns(df, ['Tên Kho'], _quantity_columns(df) + _money_columns(df))
    )
    return {
        'file_info': get_file_metadata(config['good_file']),
        'columns': formatted.columns.tolist(),
        'data': formatted.fillna('').to_dict('records'),
    }


@inventory_bp.route('/ton-kho-vat-tu')
@login_required
def page_ton_kho_vat_tu():
    return render_template('pages/ton_kho_vat_tu.html', current_user=_current_user(), active_page='ton_kho_vat_tu')


@inventory_bp.route('/Tong_hop_tien')
@login_required
def page_tong_hop_tien():
    user = _current_user()
    source_file_timestamp = 'Không đọc được thời gian file gốc'

    if os.path.exists(INVENTORY_SOURCE_FILE):
        source_file_timestamp = datetime.fromtimestamp(
            os.path.getmtime(INVENTORY_SOURCE_FILE)
        ).strftime('%d/%m/%Y %H:%M:%S')

    if not os.path.exists(INVENTORY_PROCESSED_FILE):
        flash('Không tìm thấy file dữ liệu: 66 bc ton vat tu_processed.xlsx', 'error')
        return redirect(url_for('inventory.page_ton_kho_vat_tu'))

    try:
        df_to = _apply_common_formatting(read_excel_sheet_cached(INVENTORY_PROCESSED_FILE, 'Tong_hop_tien'))
        df_ca_nhan = _apply_common_formatting(read_excel_sheet_cached(INVENTORY_PROCESSED_FILE, 'tong_hop_tien_den_ca_nhan'))

        for frame in (df_to, df_ca_nhan):
            if 'STT' in frame.columns:
                frame['STT'] = frame['STT'].apply(_normalize_stt)
            for col in frame.columns:
                if 'tổng thành tiền' in ' '.join(str(col).split()).lower():
                    frame[col] = frame[col].apply(_format_money_value)

        return render_template(
            'pages/tong_hop_tien_vat_tu.html',
            current_user=user,
            active_page='tong_hop_tien',
            tong_hop_tien_columns=df_to.columns.tolist(),
            tong_hop_tien_data=df_to.fillna('').to_dict('records'),
            tong_hop_tien_ca_nhan_columns=df_ca_nhan.columns.tolist(),
            tong_hop_tien_ca_nhan_data=df_ca_nhan.fillna('').to_dict('records'),
            source_file_timestamp=source_file_timestamp,
        )
    except Exception as exc:
        flash(f'Lỗi khi đọc dữ liệu tổng hợp tiền: {exc}', 'error')
        return redirect(url_for('inventory.page_ton_kho_vat_tu'))


@inventory_bp.route('/tra-cuu-nhanh-vat-tu')
@login_required
def page_tra_cuu_nhanh_vat_tu():
    return render_template(
        'pages/tra_cuu_nhanh_vat_tu.html',
        current_user=_current_user(),
        active_page='tra_cuu_nhanh_vat_tu',
    )


@inventory_bp.route('/api/tra-cuu-nhanh-vat-tu')
@login_required
def api_tra_cuu_nhanh_vat_tu():
    if not os.path.exists(INVENTORY_PROCESSED_FILE):
        return jsonify({'error': 'File Excel không tồn tại'}), 404

    try:
        df = _apply_common_formatting(read_excel_sheet_cached(INVENTORY_PROCESSED_FILE, 'Sheet1'))
        return jsonify(
            {
                'file_info': get_file_metadata(INVENTORY_PROCESSED_FILE),
                'columns': df.columns.tolist(),
                'data': df.fillna('').to_dict('records'),
            }
        )
    except Exception as exc:
        return jsonify({'error': f'Lỗi khi đọc file Excel: {exc}'}), 500


@inventory_bp.route('/download/ton-kho-vat-tu')
@login_required
def download_ton_kho_vat_tu():
    return safe_file_response(
        INVENTORY_PROCESSED_FILE,
        as_attachment=True,
        download_name='66_bc_ton_vat_tu.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )


@inventory_bp.route('/api/ton-kho-vat-tu')
@login_required
def api_ton_kho_vat_tu():
    if not os.path.exists(INVENTORY_PROCESSED_FILE):
        return jsonify({'error': 'File Excel không tồn tại'}), 404

    try:
        df_summary_raw = read_excel_sheet_cached(INVENTORY_PROCESSED_FILE, 'Tổng hợp')
        df_summary = _apply_common_formatting(
            _select_columns(
                df_summary_raw,
                ['STT', 'Đơn Vị', 'Tên Kho'],
                _quantity_columns(df_summary_raw),
            )
        )
        df_stats_raw = read_excel_sheet_cached(INVENTORY_PROCESSED_FILE, 'thong-ke-theo-vat-tu')
        df_stats = _apply_common_formatting(
            _select_columns(
                df_stats_raw,
                ['STT', 'Chủng loại vật tư'],
                _quantity_columns(df_stats_raw),
            )
        )

        return jsonify(
            {
                'file_info': get_file_metadata(INVENTORY_PROCESSED_FILE),
                'sheets': {
                    'Tổng hợp': build_sheet_payload(df_summary),
                    'thong-ke-theo-vat-tu': build_sheet_payload(df_stats),
                },
            }
        )
    except Exception as exc:
        return jsonify({'error': f'Lỗi khi đọc file Excel: {exc}'}), 500


@inventory_bp.route('/api/ton-kho-vat-tu-tot-thuong-dung')
@login_required
def api_ton_kho_vat_tu_tot_thuong_dung():
    if not os.path.exists(INVENTORY_COMMON_GOOD_FILE):
        return jsonify({'error': 'File Excel không tồn tại'}), 404

    try:
        df = read_excel_sheet_cached(INVENTORY_COMMON_GOOD_FILE, 'th-ton-theo-don-vi')
        df = _apply_common_formatting(_select_columns(df, ['STT', 'Đơn Vị', 'Tên Kho'], _quantity_columns(df) + _money_columns(df)))
        return jsonify(
            {
                'file_info': get_file_metadata(INVENTORY_COMMON_GOOD_FILE),
                'columns': df.columns.tolist(),
                'data': df.fillna('').to_dict('records'),
            }
        )
    except Exception as exc:
        return jsonify({'error': f'Lỗi khi đọc file Excel: {exc}'}), 500


@inventory_bp.route('/xac-minh-tam-dung')
@login_required
def page_xac_minh_tam_dung():
    return render_template(
        'pages/xac_minh_tam_dung.html',
        current_user=_current_user(),
        active_page='xac_minh_tam_dung',
    )


@inventory_bp.route('/tam-dung-khoi-phuc')
@login_required
def page_tam_dung_khoi_phuc():
    return render_template(
        'pages/tam_dung_khoi_phuc.html',
        current_user=_current_user(),
        active_page='tam_dung_khoi_phuc',
    )


@inventory_bp.route('/api/tam-dung-khoi-phuc-data')
@login_required
def get_tam_dung_khoi_phuc_data():
    try:
        date_context = resolve_date_context(
            request.args.get('date'),
            [binding['report_code'] for binding in TAM_DUNG_KHOI_PHUC_DATE_BINDINGS],
        )
        if date_context['date_has_data']:
            dataframes = load_many_tables_by_date(
                TAM_DUNG_KHOI_PHUC_DATE_BINDINGS,
                date_context['selected_date'],
            )
        else:
            dataframes = {binding['key']: pd.DataFrame() for binding in TAM_DUNG_KHOI_PHUC_DATE_BINDINGS}

        return jsonify(
            {
                'file_info': get_file_metadata(current_app.config['REPORT_HISTORY_DB_PATH']),
                'selected_date': date_context['selected_date'],
                'latest_available_date': date_context['latest_available_date'],
                'available_dates': date_context['available_dates'],
                'date_has_data': date_context['date_has_data'],
                'sheets': OrderedDict(
                    [
                        (
                            'v_tam_dung_khoi_phuc_dich_vu_chi_tiet_combined_tong_hop_theo_to',
                            build_sheet_payload(
                                dataframes['v_tam_dung_khoi_phuc_dich_vu_chi_tiet_combined_tong_hop_theo_to']
                            ),
                        ),
                        (
                            'v_tam_dung_khoi_phuc_dich_vu_chi_tiet_combined_tong_hop_theo_nvkt',
                            build_sheet_payload(
                                dataframes['v_tam_dung_khoi_phuc_dich_vu_chi_tiet_combined_tong_hop_theo_nvkt']
                            ),
                        ),
                    ]
                ),
            }
        )
    except Exception as exc:
        return jsonify({'error': f'Lỗi khi đọc dữ liệu tạm dừng khôi phục từ DB: {exc}'}), 500


@inventory_bp.route('/api/xac-minh-tam-dung-data')
@login_required
def get_xac_minh_tam_dung_data():
    try:
        df = load_xac_minh_chi_tiet_df('Trung tâm Viễn thông Sơn Tây')
        sheets = {}
        doi_vt_col = 'Đội VT'
        service_col = 'Dịch vụ'

        if doi_vt_col in df.columns:
            for doi_vt in df[doi_vt_col].dropna().unique():
                if not str(doi_vt).strip():
                    continue
                df_dv = df[df[doi_vt_col] == doi_vt].reset_index(drop=True)
                tab_name = str(doi_vt).replace('Tổ Kỹ thuật Địa bàn ', '')
                fiber_df = df_dv[df_dv[service_col].astype(str).str.contains('Fiber', case=False, na=False)].reset_index(drop=True)
                mytv_df = df_dv[df_dv[service_col].astype(str).str.contains('MyTV', case=False, na=False)].reset_index(drop=True)
                sheets[tab_name] = {
                    'fiber': build_sheet_payload(fiber_df),
                    'mytv': build_sheet_payload(mytv_df),
                }

        return jsonify({'file_info': get_file_metadata(current_app.config['REPORT_HISTORY_DB_PATH']), 'sheets': sheets})
    except Exception as exc:
        return jsonify({'error': f'Lỗi khi đọc dữ liệu xác minh tạm dừng từ DB: {exc}'}), 500


def _make_team_page(config):
    @login_required
    def view():
        return render_template(config['template'], current_user=_current_user(), active_page=config['active_page'])

    return view


def _make_team_download(config, key):
    download_name_key = 'good_download_name' if key == 'good_file' else 'base_download_name'

    @login_required
    def view():
        return safe_file_response(
            config[key],
            as_attachment=True,
            download_name=config[download_name_key],
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        )

    return view


def _make_team_api(config):
    @login_required
    def view():
        if not os.path.exists(config['base_file']):
            return jsonify({'error': 'File Excel không tồn tại'}), 404
        try:
            return jsonify(_team_inventory_payload(config))
        except Exception as exc:
            return jsonify({'error': f'Lỗi khi đọc file Excel: {exc}'}), 500

    return view


def _make_team_good_api(config):
    @login_required
    def view():
        if not os.path.exists(config['good_file']):
            return jsonify({'error': 'File Excel không tồn tại'}), 404
        try:
            return jsonify(_team_good_inventory_payload(config))
        except Exception as exc:
            return jsonify({'error': f'Lỗi khi đọc file Excel: {exc}'}), 500

    return view


for team_key, team_config in INVENTORY_TEAM_CONFIGS.items():
    inventory_bp.add_url_rule(
        f'/ton-kho-vat-tu-{team_key}',
        endpoint=f'page_ton_kho_vat_tu_{team_key}',
        view_func=_make_team_page(team_config),
    )
    inventory_bp.add_url_rule(
        f'/download/ton-kho-vat-tu-{team_key}',
        endpoint=f'download_ton_kho_vat_tu_{team_key}',
        view_func=_make_team_download(team_config, 'base_file'),
    )
    inventory_bp.add_url_rule(
        f'/api/ton-kho-vat-tu-{team_key}',
        endpoint=f'api_ton_kho_vat_tu_{team_key}',
        view_func=_make_team_api(team_config),
    )
    inventory_bp.add_url_rule(
        f'/download/ton-kho-vat-tu-{team_key}-tot-thuong-dung',
        endpoint=f'download_ton_kho_vat_tu_{team_key}_tot_thuong_dung',
        view_func=_make_team_download(team_config, 'good_file'),
    )
    inventory_bp.add_url_rule(
        f'/api/ton-kho-vat-tu-{team_key}-tot-thuong-dung',
        endpoint=f'api_ton_kho_vat_tu_{team_key}_tot_thuong_dung',
        view_func=_make_team_good_api(team_config),
    )
