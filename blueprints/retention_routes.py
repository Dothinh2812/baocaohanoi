import os
from collections import OrderedDict

import pandas as pd
from flask import current_app
from flask import Blueprint, jsonify, render_template, request, session

from app_helpers import (
    build_file_info,
    build_sheet_payload,
)
from auth import get_user_by_username, login_required
from config import UNIT_NAME
from repositories import (
    load_table_by_date,
    resolve_date_context,
    load_thu_hoi_chi_tiet_df,
    load_thu_hoi_tong_hop_df,
)


retention_bp = Blueprint('retention', __name__)


GIA_HAN_REPORT_CODES = (
    'ghtt_ghtt_hni_report',
    'ghtt_ghtt_sontay_report',
    'ghtt_ghtt_nvktdb_report',
)


def _current_user():
    return get_user_by_username(session['username'])


def _resolve_gia_han_context():
    return resolve_date_context(request.args.get('date'), GIA_HAN_REPORT_CODES)


def _select_columns_transform(columns, *, datetime_format='%Y-%m-%d %H:%M:%S'):
    def transform(df, _sheet_name):
        filtered_df = df[[col for col in columns if col in df.columns]].copy()
        return build_sheet_payload(filtered_df, datetime_format=datetime_format)

    return transform


def _rename_all_sheet(payload, *, source_name='Tất cả', target_name='TTVT Sơn Tây'):
    sheets = payload.get('sheets', {})
    if source_name not in sheets:
        return payload

    ordered_sheets = OrderedDict([(target_name, sheets[source_name])])
    for sheet_name, sheet_data in sheets.items():
        if sheet_name != source_name:
            ordered_sheets[sheet_name] = sheet_data
    payload['sheets'] = ordered_sheets
    return payload


def _group_dataframe_sheets(
    df,
    group_column,
    *,
    default_sheet_name='data',
    drop_group_column=False,
    datetime_format='%Y-%m-%d %H:%M:%S',
):
    if group_column not in df.columns:
        return OrderedDict([(default_sheet_name, build_sheet_payload(df, datetime_format=datetime_format))])

    sheets = OrderedDict()
    for group_value in df[group_column].dropna().unique():
        if not str(group_value).strip():
            continue
        group_df = df[df[group_column] == group_value].reset_index(drop=True)
        if drop_group_column and group_column in group_df.columns:
            group_df = group_df.drop(columns=[group_column])
        sheets[str(group_value)] = build_sheet_payload(group_df, datetime_format=datetime_format)

    return sheets


@retention_bp.route('/giahan')
@login_required
def page_giahan():
    return render_template('pages/giahan.html', current_user=_current_user(), active_page='giahan')


@retention_bp.route('/thuhoi')
@login_required
def page_thuhoi():
    return render_template('pages/thuhoi.html', current_user=_current_user(), active_page='thuhoi')


@retention_bp.route('/api/giahan-ghtt-hni')
def get_giahan_ghtt_hni():
    try:
        date_context = _resolve_gia_han_context()
        df = (
            load_table_by_date(
                'ghtt_ghtt_hni_report',
                'ghtt_ghtt_hni_report_kq_hni',
                date_context['selected_date'],
                order_by='t."Đơn vị"',
            )
            if date_context['date_has_data']
            else pd.DataFrame()
        )
        return jsonify(_build_ghtt_multi_header_payload(df, date_context=date_context))
    except Exception as exc:
        current_app.logger.exception('Khong the tai du lieu GHTT HNI tu SQLite: %s', exc)
        return jsonify({'error': f'Lỗi khi đọc dữ liệu GHTT HNI từ SQLite: {exc}'}), 500


def _build_ghtt_multi_header_payload(df, *, date_context=None):
    header1 = [
        'Đơn vị',
        'Tháng T', 'Tháng T', 'Tháng T',
        'Tháng T+1', 'Tháng T+1', 'Tháng T+1',
        'GHTT > 6 tháng', 'GHTT > 6 tháng', 'GHTT > 6 tháng',
        'Tổng',
    ]
    header2 = [
        'Đơn vị',
        'Hoàn thành T', 'Giao T', 'Tỷ lệ T',
        'Hoàn thành T+1', 'Giao T+1', 'Tỷ lệ T+1',
        'Số lượng', 'Hoàn thành T+1', 'Tỷ lệ > 6 tháng',
        'Tỷ lệ tổng',
    ]
    merges = [
        {'min_col': 1, 'max_col': 1, 'min_row': 1, 'max_row': 2},
        {'min_col': 2, 'max_col': 4, 'min_row': 1, 'max_row': 1},
        {'min_col': 5, 'max_col': 7, 'min_row': 1, 'max_row': 1},
        {'min_col': 8, 'max_col': 10, 'min_row': 1, 'max_row': 1},
        {'min_col': 11, 'max_col': 11, 'min_row': 1, 'max_row': 2},
    ]

    rows = []
    for _, row in df.iterrows():
        rows.append([
            row.get('Đơn vị', ''),
            row.get('Hoàn thành T', ''),
            row.get('Giao T', ''),
            row.get('Tỷ lệ T', ''),
            row.get('Hoàn thành T+1', ''),
            row.get('Giao T+1', ''),
            row.get('Tỷ lệ T+1', ''),
            row.get('Số lượng GHTT > 6 tháng', ''),
            row.get('Hoàn thành > 6 tháng T+1', ''),
            row.get('Tỷ lệ > 6 tháng T+1', ''),
            row.get('Tỷ lệ tổng', ''),
        ])

    return {
        'file_info': build_file_info(current_app.config['REPORT_HISTORY_DB_PATH'], include_name=True),
        'selected_date': date_context['selected_date'] if date_context else None,
        'latest_available_date': date_context['latest_available_date'] if date_context else None,
        'available_dates': date_context['available_dates'] if date_context else [],
        'date_has_data': date_context['date_has_data'] if date_context else True,
        'header1': header1,
        'header2': header2,
        'merges': merges,
        'data': rows,
        'total_cols': len(header1),
    }


@retention_bp.route('/api/giahan-ghtt-sty')
def get_giahan_ghtt_sty():
    try:
        date_context = _resolve_gia_han_context()
        df = (
            load_table_by_date(
                'ghtt_ghtt_sontay_report',
                'ghtt_ghtt_sontay_report_kq_sontay',
                date_context['selected_date'],
                order_by='t."Đơn vị"',
            )
            if date_context['date_has_data']
            else pd.DataFrame()
        )
        return jsonify(_build_ghtt_multi_header_payload(df, date_context=date_context))
    except Exception as exc:
        current_app.logger.exception('Khong the tai du lieu GHTT don vi tu SQLite: %s', exc)
        return jsonify({'error': f'Lỗi khi đọc dữ liệu GHTT từ SQLite: {exc}'}), 500


@retention_bp.route('/api/giahan-ghtt-nvktdb')
def get_giahan_ghtt_nvktdb():
    try:
        date_context = _resolve_gia_han_context()
        df = (
            load_table_by_date(
                'ghtt_ghtt_nvktdb_report',
                'ghtt_ghtt_nvktdb_report_kq_nvktdb',
                date_context['selected_date'],
                order_by='t."Đơn vị", t."NVKT"',
            )
            if date_context['date_has_data']
            else pd.DataFrame()
        )
        grouped_sheets = _group_dataframe_sheets(
            df,
            'Đơn vị',
            default_sheet_name='kq_nvktdb',
            drop_group_column=True,
        )
        return jsonify({
            'file_info': build_file_info(current_app.config['REPORT_HISTORY_DB_PATH'], include_name=True),
            'selected_date': date_context['selected_date'],
            'latest_available_date': date_context['latest_available_date'],
            'available_dates': date_context['available_dates'],
            'date_has_data': date_context['date_has_data'],
            'sheets': OrderedDict(sorted(grouped_sheets.items())),
        })
    except Exception as exc:
        current_app.logger.exception('Khong the tai du lieu GHTT NVKT tu SQLite: %s', exc)
        return jsonify({'error': f'Lỗi khi đọc dữ liệu GHTT NVKT từ SQLite: {exc}'}), 500


@retention_bp.route('/api/giahan-data')
def get_giahan_data():
    return jsonify({
        'error': 'legacy_endpoint_disabled',
        'title': 'Gia hạn KR6',
        'reason': 'API Excel KR6 cũ đã bị thay thế bằng dữ liệu GHTT từ report_history.db.',
        'required_display_contract': {'replacement': ['/api/giahan-ghtt-sty', '/api/giahan-ghtt-nvktdb']},
    }), 501


@retention_bp.route('/api/giahan-data-to')
def get_giahan_data_to():
    return jsonify({
        'error': 'legacy_endpoint_disabled',
        'title': 'Gia hạn KR6 theo tổ',
        'reason': 'API Excel KR6 cũ đã bị thay thế bằng dữ liệu GHTT từ report_history.db.',
        'required_display_contract': {'replacement': ['/api/giahan-ghtt-sty']},
    }), 501


@retention_bp.route('/api/giahan-kr7-data')
def get_giahan_kr7_data():
    return jsonify({
        'error': 'legacy_endpoint_disabled',
        'title': 'Gia hạn KR7',
        'reason': 'API Excel KR7 cũ đã bị thay thế bằng dữ liệu GHTT từ report_history.db.',
        'required_display_contract': {'replacement': ['/api/giahan-ghtt-nvktdb']},
    }), 501


@retention_bp.route('/api/giahan-kr7-data-to')
def get_giahan_kr7_data_to():
    return jsonify({
        'error': 'legacy_endpoint_disabled',
        'title': 'Gia hạn KR7 theo tổ',
        'reason': 'API Excel KR7 cũ đã bị thay thế bằng dữ liệu GHTT từ report_history.db.',
        'required_display_contract': {'replacement': ['/api/giahan-ghtt-sty']},
    }), 501


@retention_bp.route('/api/thu-hoi-data')
def get_thu_hoi_data():
    try:
        db_path = current_app.config['REPORT_HISTORY_DB_PATH']
        return jsonify({
            'file_info': build_file_info(db_path, include_name=True),
            'tong_hop': build_sheet_payload(load_thu_hoi_tong_hop_df()),
            'chi_tiet': build_sheet_payload(load_thu_hoi_chi_tiet_df()),
        })
    except FileNotFoundError as exc:
        return jsonify({'error': str(exc)}), 404
    except Exception as exc:
        current_app.logger.exception('Khong the tai du lieu thu hoi tu SQLite: %s', exc)
        return jsonify({'error': f'Lỗi khi đọc dữ liệu Thu hồi từ SQLite: {exc}'}), 500
