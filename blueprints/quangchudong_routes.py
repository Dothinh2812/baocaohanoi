import io
import re
import unicodedata
from datetime import datetime

import pandas as pd
from flask import Blueprint, jsonify, render_template, send_file, session

from auth import get_user_by_username, login_required
from services import get_quangchudong_cache


quangchudong_bp = Blueprint('quangchudong', __name__)


def _slugify_nvkt_name(value):
    normalized = unicodedata.normalize('NFD', str(value or '').strip())
    without_marks = ''.join(char for char in normalized if unicodedata.category(char) != 'Mn')
    without_marks = without_marks.replace('Đ', 'D').replace('đ', 'd')
    slug = re.sub(r'[^a-zA-Z0-9]+', '-', without_marks.lower()).strip('-')
    return slug


def _parse_alert_time(value):
    normalized = str(value or '').strip().replace('T', ' ')
    if not normalized:
        return None

    for time_format in ('%Y-%m-%d %H:%M:%S.%f', '%Y-%m-%d %H:%M:%S'):
        try:
            return datetime.strptime(normalized, time_format)
        except ValueError:
            continue
    return None


def _get_event_time(row):
    return _parse_alert_time(row.get('first_off_time') or row.get('alert_time'))


def _filter_off_today_rows(rows, cutoff_time):
    filtered = []
    for row in rows:
        event_time = _get_event_time(row)
        if event_time is not None and event_time >= cutoff_time:
            filtered.append(row)
    filtered.sort(key=lambda row: _get_event_time(row) or datetime.min, reverse=True)
    return filtered


def _with_sleep_days(row, now):
    enriched = dict(row)
    event_time = _get_event_time(row)
    if event_time is None:
        enriched['thoi_gian_ngu'] = ''
        return enriched
    diff_days = max(1, int((now - event_time).total_seconds() // 86400))
    enriched['thoi_gian_ngu'] = diff_days
    return enriched


def _filter_sleeping_rows(rows, cutoff_time, now):
    filtered = []
    for row in rows:
        event_time = _get_event_time(row)
        if event_time is not None and event_time < cutoff_time:
            filtered.append(_with_sleep_days(row, now))
    filtered.sort(key=lambda row: row.get('thoi_gian_ngu') or 0, reverse=True)
    return filtered


def _today_six_am():
    return datetime.now().replace(hour=6, minute=0, second=0, microsecond=0)


def _filter_nvkt_alerts(alerts, nvkt_slug, cutoff_time):
    rows = []
    display_name = ''

    for alert in alerts:
        nvkt_name = str(alert.get('ten_nvkt_db') or '').strip()
        if _slugify_nvkt_name(nvkt_name) != nvkt_slug:
            continue

        if not display_name and nvkt_name:
            display_name = nvkt_name

        event_time = _parse_alert_time(alert.get('first_off_time') or alert.get('alert_time'))
        if event_time is None or event_time < cutoff_time:
            continue

        rows.append(alert)

    rows.sort(
        key=lambda row: _parse_alert_time(row.get('first_off_time') or row.get('alert_time')) or datetime.min,
        reverse=True,
    )
    return rows, display_name


@quangchudong_bp.route('/quangchudong')
@login_required
def page_quangchudong():
    user = get_user_by_username(session['username'])
    return render_template('quangchudong.html', current_user=user, active_page='quangchudong')


@quangchudong_bp.route('/quangchudong/<nvkt_slug>')
@login_required
def page_quangchudong_nvkt(nvkt_slug):
    user = get_user_by_username(session['username'])
    return render_template(
        'quangchudong_nvkt_mobile.html',
        current_user=user,
        nvkt_slug=nvkt_slug,
    )


@quangchudong_bp.route('/api/quangchudong/active')
@login_required
def get_active_outages():
    return jsonify(get_quangchudong_cache().get_active_payload())


@quangchudong_bp.route('/api/quangchudong/dashboard')
@login_required
def get_dashboard_payload():
    return jsonify(get_quangchudong_cache().get_dashboard_payload())


@quangchudong_bp.route('/api/quangchudong/nvkt/<nvkt_slug>/alerts')
@login_required
def get_nvkt_alerts(nvkt_slug):
    normalized_slug = _slugify_nvkt_name(nvkt_slug)
    cutoff_time = _today_six_am()
    payload = get_quangchudong_cache().get_dashboard_payload()
    active = payload.get('active') or {}
    alerts, display_name = _filter_nvkt_alerts(
        active.get('alerts') or [],
        normalized_slug,
        cutoff_time,
    )

    return jsonify({
        'alerts': alerts,
        'cache': active.get('cache') or {},
        'count': len(alerts),
        'cutoff_time': cutoff_time.strftime('%Y-%m-%d %H:%M:%S'),
        'nvkt_name': display_name,
        'nvkt_slug': normalized_slug,
    })


@quangchudong_bp.route('/api/quangchudong/recovered')
@login_required
def get_recovered_alerts():
    return jsonify(get_quangchudong_cache().get_recovered_payload())


@quangchudong_bp.route('/api/quangchudong/wide-area')
@login_required
def get_wide_area_outages():
    return jsonify(get_quangchudong_cache().get_wide_area_payload())


@quangchudong_bp.route('/api/quangchudong/wide-area-groups')
@login_required
def get_wide_area_groups():
    return jsonify(get_quangchudong_cache().get_wide_area_groups_payload())


@quangchudong_bp.route('/api/quangchudong/exclusion-list')
@login_required
def get_exclusion_list():
    return jsonify(get_quangchudong_cache().get_exclusion_list_payload())


@quangchudong_bp.route('/api/quangchudong/pattern-exclusions')
@login_required
def get_pattern_exclusions():
    return jsonify(get_quangchudong_cache().get_pattern_exclusions_payload())


@quangchudong_bp.route('/api/quangchudong/port-down-groups')
@login_required
def get_port_down_groups():
    return jsonify(get_quangchudong_cache().get_port_down_groups_payload())


@quangchudong_bp.route('/api/quangchudong/stats')
@login_required
def get_outage_stats():
    return jsonify(get_quangchudong_cache().get_stats_payload())


_SUBSCRIBER_EXPORT_COLUMNS = [
    ('ma_tb', 'Mã TB'),
    ('ten_tb', 'Tên TB'),
    ('diachi_ld', 'Địa chỉ'),
    ('dienthoai_lh', 'SĐT'),
    ('ten_nvkt_db', 'NVKT'),
    ('doi_vt', 'Đội VT'),
    ('olt_name', 'OLT'),
    ('sa', 'SA'),
    ('first_off_time', 'Thời gian OFF'),
    ('duration_text', 'Thời lượng'),
    ('ngay_bh', 'Phiếu BH Lúc'),
]

_GROUP_EXPORT_COLUMNS = [
    ('parent_port_key', 'Parent port'),
    ('olt_name', 'OLT'),
    ('sa', 'SA'),
    ('down_count', 'Số thuê bao'),
    ('start_time', 'Thời điểm bắt đầu'),
]


def _rows_to_export_frame(rows, columns):
    records = []
    for row in rows:
        records.append({label: row.get(key, '') for key, label in columns})
    return pd.DataFrame(records, columns=[label for _, label in columns])


def _subscriber_export_frame(rows, *, include_sleep_days=False):
    columns = list(_SUBSCRIBER_EXPORT_COLUMNS)
    if include_sleep_days:
        columns.append(('thoi_gian_ngu', 'Thời gian ngủ'))
    return _rows_to_export_frame(rows, columns)


@quangchudong_bp.route('/download/quangchudong-report')
@login_required
def download_quangchudong_report():
    payload = get_quangchudong_cache().get_dashboard_payload()
    active = payload.get('active') or {}
    alerts = active.get('alerts') or []
    cutoff_time = _today_six_am()
    now = datetime.now()

    sheets = [
        ('Canh_bao_DOWN', _subscriber_export_frame(alerts)),
        ('OFF_trong_ngay', _subscriber_export_frame(_filter_off_today_rows(alerts, cutoff_time))),
        ('Thue_bao_ngu', _subscriber_export_frame(
            _filter_sleeping_rows(alerts, cutoff_time, now),
            include_sleep_days=True,
        )),
        ('Su_co_dien_rong', _rows_to_export_frame(payload.get('wide_area_groups') or [], _GROUP_EXPORT_COLUMNS)),
        ('Pattern_exclusion', _subscriber_export_frame(payload.get('pattern_exclusions') or [])),
        ('Port_OLT_Down', _rows_to_export_frame(payload.get('port_down_groups') or [], _GROUP_EXPORT_COLUMNS)),
    ]

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for sheet_name, frame in sheets:
            frame.to_excel(writer, sheet_name=sheet_name, index=False)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name=f'quang_chu_dong_{datetime.now():%Y%m%d_%H%M}.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )
