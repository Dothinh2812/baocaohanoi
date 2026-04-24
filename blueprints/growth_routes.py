import os
from collections import OrderedDict
import pandas as pd
from flask import Blueprint, current_app, jsonify, render_template, request, session

from app_helpers import build_file_info, build_sheet_payload
from auth import get_user_by_username, login_required
from config import UNIT_NAME
from repositories import (
    load_ket_qua_tiep_thi_nv_df,
    load_many_tables_by_date,
    resolve_date_context,
)


growth_bp = Blueprint('growth', __name__)


TIEP_THI_DATE_BINDINGS = [
    {
        'key': 'tong_hop',
        'report_code': 'kq_tiep_thi_kq_tiep_thi_report',
        'table_name': 'kq_tiep_thi_kq_tiep_thi_report_kq_th',
        'order_by': 't."STT"',
    },
    {
        'key': 'chi_tiet',
        'report_code': 'kq_tiep_thi_kq_tiep_thi_report',
        'table_name': 'kq_tiep_thi_kq_tiep_thi_report_kq_tiep_thi',
        'order_by': 't."Đơn vị", t."Mã NV"',
    },
]


NGUNG_PSC_DATE_BINDINGS = [
    {
        'key': 'v_ngung_psc_fiber_thang_t_1_cap_ttvt',
        'report_code': 'tam_dung_khoi_phuc_dich_vu_ngung_psc_fiber_thang_t_1_cap_ttvt',
        'table_name': 'tam_dung_khoi_phuc_dich_vu_ngung_psc_fiber_thang_t_1_cap_ttvt_th_ngung_psc_thang_t_1',
        'order_by': 't."Đơn vị/Nhân viên KT"',
    },
    {
        'key': 'v_ngung_psc_mytv_thang_t_1_cap_ttvt',
        'report_code': 'tam_dung_khoi_phuc_dich_vu_ngung_psc_mytv_thang_t_1_cap_ttvt',
        'table_name': 'tam_dung_khoi_phuc_dich_vu_ngung_psc_mytv_thang_t_1_cap_ttvt_th_ngung_psc_thang_t_1',
        'order_by': 't."Đơn vị/Nhân viên KT"',
    },
]


def _resolve_ngung_psc_context():
    return resolve_date_context(
        request.args.get('date'),
        [binding['report_code'] for binding in NGUNG_PSC_DATE_BINDINGS],
    )


def _resolve_tiep_thi_context():
    return resolve_date_context(
        request.args.get('date'),
        [binding['report_code'] for binding in TIEP_THI_DATE_BINDINGS],
    )


def _load_ngung_psc_tables(date_context):
    if not date_context['date_has_data']:
        return {binding['key']: pd.DataFrame() for binding in NGUNG_PSC_DATE_BINDINGS}
    return load_many_tables_by_date(
        NGUNG_PSC_DATE_BINDINGS,
        date_context['selected_date'],
    )


def _current_user():
    return get_user_by_username(session['username'])


@growth_bp.route('/tiepthi')
@login_required
def page_tiepthi():
    return render_template(
        'pages/tiepthi.html',
        current_user=_current_user(),
        active_page='tiepthi',
        kq_th_data=[],
        kq_th_columns=[],
    )


@growth_bp.route('/thuc-tang-ngung-psc')
@login_required
def page_ngungpsc():
    return render_template('pages/ngungpsc.html', current_user=_current_user(), active_page='ngungpsc')


@growth_bp.route('/api/tiepthi-data')
def get_tiepthi_data():
    try:
        date_context = _resolve_tiep_thi_context()
        if date_context['date_has_data']:
            dataframes = load_many_tables_by_date(TIEP_THI_DATE_BINDINGS, date_context['selected_date'])
            tong_hop_df = dataframes['tong_hop']
            df = dataframes['chi_tiet']
        else:
            tong_hop_df = pd.DataFrame()
            df = pd.DataFrame()

        if df.empty:
            data = {
                'file_info': build_file_info(current_app.config['REPORT_HISTORY_DB_PATH'], include_name=True),
                'selected_date': date_context['selected_date'],
                'latest_available_date': date_context['latest_available_date'],
                'available_dates': date_context['available_dates'],
                'date_has_data': date_context['date_has_data'],
                'tong_hop': build_sheet_payload(tong_hop_df),
                'sheets': OrderedDict({'Tất cả': build_sheet_payload(pd.DataFrame())}),
            }
            return jsonify(data)

        sheets = OrderedDict()
        sheets['Tất cả'] = build_sheet_payload(df)
        for don_vi in df['Đơn vị'].dropna().unique():
            unit_df = df[df['Đơn vị'] == don_vi].reset_index(drop=True)
            sheets[str(don_vi)] = build_sheet_payload(unit_df)

        data = {
            'file_info': build_file_info(current_app.config['REPORT_HISTORY_DB_PATH'], include_name=True),
            'selected_date': date_context['selected_date'],
            'latest_available_date': date_context['latest_available_date'],
            'available_dates': date_context['available_dates'],
            'date_has_data': date_context['date_has_data'],
            'tong_hop': build_sheet_payload(tong_hop_df),
            'sheets': sheets,
        }
        return jsonify(data)
    except FileNotFoundError as exc:
        return jsonify({'error': str(exc)}), 404
    except Exception as exc:
        current_app.logger.exception('Khong the tai du lieu tiep thi tu SQLite: %s', exc)
        return jsonify({'error': f'Lỗi khi đọc dữ liệu Tiếp thị từ SQLite: {exc}'}), 500


@growth_bp.route('/api/ngungpsc-data')
def get_ngungpsc_data():
    try:
        date_context = _resolve_ngung_psc_context()
        dataframes = _load_ngung_psc_tables(date_context)
        return jsonify({
            'file_info': build_file_info(current_app.config['REPORT_HISTORY_DB_PATH'], include_name=True),
            'selected_date': date_context['selected_date'],
            'latest_available_date': date_context['latest_available_date'],
            'available_dates': date_context['available_dates'],
            'date_has_data': date_context['date_has_data'],
            'sheets': OrderedDict({
                'v_ngung_psc_fiber_thang_t_1_cap_ttvt': build_sheet_payload(
                    dataframes['v_ngung_psc_fiber_thang_t_1_cap_ttvt']
                )
            }),
        })
    except Exception as exc:
        current_app.logger.exception('Khong the tai du lieu ngung PSC Fiber tu SQLite: %s', exc)
        return jsonify({'error': f'Lỗi khi đọc dữ liệu Ngưng PSC Fiber từ SQLite: {exc}'}), 500


@growth_bp.route('/api/ngungpsc-mytv-data')
def get_ngungpsc_mytv_data():
    try:
        date_context = _resolve_ngung_psc_context()
        dataframes = _load_ngung_psc_tables(date_context)
        return jsonify({
            'file_info': build_file_info(current_app.config['REPORT_HISTORY_DB_PATH'], include_name=True),
            'selected_date': date_context['selected_date'],
            'latest_available_date': date_context['latest_available_date'],
            'available_dates': date_context['available_dates'],
            'date_has_data': date_context['date_has_data'],
            'sheets': OrderedDict({
                'v_ngung_psc_mytv_thang_t_1_cap_ttvt': build_sheet_payload(
                    dataframes['v_ngung_psc_mytv_thang_t_1_cap_ttvt']
                )
            }),
        })
    except Exception as exc:
        current_app.logger.exception('Khong the tai du lieu ngung PSC MyTV tu SQLite: %s', exc)
        return jsonify({'error': f'Lỗi khi đọc dữ liệu Ngưng PSC MyTV từ SQLite: {exc}'}), 500
