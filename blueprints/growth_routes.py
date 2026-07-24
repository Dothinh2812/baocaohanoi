import os
from collections import OrderedDict
import pandas as pd
from flask import Blueprint, current_app, jsonify, render_template, request, session

from app_helpers import build_file_info, build_sheet_payload
from auth import get_user_by_username, login_required
from config import UNIT_NAME
from repositories import (
    load_many_tables_by_date,
    resolve_date_context,
)
from repositories.tiep_thi_source import (
    build_tiep_thi_marketer_service_summary_df,
    build_tiep_thi_service_summary_df,
    build_tiep_thi_summary_df,
    build_tiep_thi_unit_summary_df,
    build_tiep_thi_unit_service_summary_df,
    load_tiep_thi_detail_df,
    load_tiep_thi_quarterly_summary_df,
    load_tiep_thi_yearly_summary_df,
    resolve_tiep_thi_date_context,
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


def _build_empty_tiep_thi_payload(db_path, date_context):
    return {
        'file_info': build_file_info(db_path, include_name=True),
        'selected_date': date_context['selected_date'],
        'latest_available_date': date_context['latest_available_date'],
        'available_dates': date_context['available_dates'],
        'date_has_data': date_context['date_has_data'],
        'tong_hop_don_vi': build_sheet_payload(pd.DataFrame()),
        'tong_hop_quy': build_sheet_payload(pd.DataFrame()),
        'tong_hop_nam': build_sheet_payload(pd.DataFrame()),
        'tong_hop': build_sheet_payload(pd.DataFrame()),
        'tong_hop_loai_dich_vu': build_sheet_payload(pd.DataFrame()),
        'tong_hop_don_vi_loai_dich_vu': build_sheet_payload(pd.DataFrame()),
        'tong_hop_nguoi_tiep_thi_loai_dich_vu': build_sheet_payload(pd.DataFrame()),
        'sheets': OrderedDict({'Tất cả': build_sheet_payload(pd.DataFrame())}),
    }


def _build_tiep_thi_payload(
    db_path,
    date_context,
    tong_hop_don_vi_df,
    tong_hop_quy_df,
    tong_hop_nam_df,
    tong_hop_df,
    tong_hop_loai_dich_vu_df,
    tong_hop_don_vi_loai_dich_vu_df,
    tong_hop_nguoi_tiep_thi_loai_dich_vu_df,
    df,
):
    if df.empty:
        return _build_empty_tiep_thi_payload(db_path, date_context)

    sheets = OrderedDict()
    sheets['Tất cả'] = build_sheet_payload(df)
    for don_vi in df['Đơn vị'].dropna().unique():
        unit_df = df[df['Đơn vị'] == don_vi].reset_index(drop=True)
        sheets[str(don_vi)] = build_sheet_payload(unit_df)

    return {
        'file_info': build_file_info(db_path, include_name=True),
        'selected_date': date_context['selected_date'],
        'latest_available_date': date_context['latest_available_date'],
        'available_dates': date_context['available_dates'],
        'date_has_data': date_context['date_has_data'],
        'tong_hop_don_vi': build_sheet_payload(tong_hop_don_vi_df),
        'tong_hop_quy': build_sheet_payload(tong_hop_quy_df),
        'tong_hop_nam': build_sheet_payload(tong_hop_nam_df),
        'tong_hop': build_sheet_payload(tong_hop_df),
        'tong_hop_loai_dich_vu': build_sheet_payload(tong_hop_loai_dich_vu_df),
        'tong_hop_don_vi_loai_dich_vu': build_sheet_payload(tong_hop_don_vi_loai_dich_vu_df),
        'tong_hop_nguoi_tiep_thi_loai_dich_vu': build_sheet_payload(tong_hop_nguoi_tiep_thi_loai_dich_vu_df),
        'sheets': sheets,
    }


def _get_tiepthi_data_from_configured_db(db_path):
    date_context = resolve_tiep_thi_date_context(db_path, request.args.get('date'))
    if not date_context['date_has_data']:
        return _build_empty_tiep_thi_payload(db_path, date_context)

    df = load_tiep_thi_detail_df(db_path, date_context['selected_date'])
    tong_hop_don_vi_df = build_tiep_thi_unit_summary_df(df)
    tong_hop_quy_df = load_tiep_thi_quarterly_summary_df(db_path, date_context['selected_date'])
    tong_hop_nam_df = load_tiep_thi_yearly_summary_df(db_path, date_context['selected_date'])
    tong_hop_df = build_tiep_thi_summary_df(df)
    tong_hop_loai_dich_vu_df = build_tiep_thi_service_summary_df(df)
    tong_hop_don_vi_loai_dich_vu_df = build_tiep_thi_unit_service_summary_df(df)
    tong_hop_nguoi_tiep_thi_loai_dich_vu_df = build_tiep_thi_marketer_service_summary_df(df)
    return _build_tiep_thi_payload(
        db_path,
        date_context,
        tong_hop_don_vi_df,
        tong_hop_quy_df,
        tong_hop_nam_df,
        tong_hop_df,
        tong_hop_loai_dich_vu_df,
        tong_hop_don_vi_loai_dich_vu_df,
        tong_hop_nguoi_tiep_thi_loai_dich_vu_df,
        df,
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
        tiep_thi_db_path = current_app.config.get('TIEP_THI_DB_PATH')
        if tiep_thi_db_path:
            return jsonify(_get_tiepthi_data_from_configured_db(tiep_thi_db_path))

        date_context = _resolve_tiep_thi_context()
        if date_context['date_has_data']:
            dataframes = load_many_tables_by_date(TIEP_THI_DATE_BINDINGS, date_context['selected_date'])
            tong_hop_df = dataframes['tong_hop']
            df = dataframes['chi_tiet']
        else:
            tong_hop_df = pd.DataFrame()
            df = pd.DataFrame()

        if df.empty:
            return jsonify(_build_empty_tiep_thi_payload(current_app.config['REPORT_HISTORY_DB_PATH'], date_context))

        return jsonify(_build_tiep_thi_payload(
            current_app.config['REPORT_HISTORY_DB_PATH'],
            date_context,
            pd.DataFrame(),
            pd.DataFrame(),
            pd.DataFrame(),
            tong_hop_df,
            pd.DataFrame(),
            pd.DataFrame(),
            pd.DataFrame(),
            df,
        ))
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
