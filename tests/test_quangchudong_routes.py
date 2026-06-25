from pathlib import Path
import sys
from datetime import datetime
import io
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from dashboard import app
from blueprints import quangchudong_routes


class _FakeQuangChuDongCache:
    def get_dashboard_payload(self):
        return {
            'active': {
                'alerts': [{'ma_tb': 'TB1'}],
                'excluded': [],
                'sources': [],
                'cache': {'refreshed_at': '13/05/2026 07:30:00'},
            },
            'wide_area_groups': [{'parent_port_key': 'OLT1'}],
            'pattern_exclusions': [{'ma_tb': 'TB2'}],
            'port_down_groups': [{'parent_port_key': 'OLT2'}],
        }


def test_quangchudong_dashboard_api_returns_combined_payload(monkeypatch):
    monkeypatch.setattr(
        quangchudong_routes,
        'get_quangchudong_cache',
        lambda: _FakeQuangChuDongCache(),
    )

    client = app.test_client()
    with client.session_transaction() as session:
        session['username'] = 'test-user'

    response = client.get('/api/quangchudong/dashboard')

    assert response.status_code == 200
    assert response.get_json() == {
        'active': {
            'alerts': [{'ma_tb': 'TB1'}],
            'excluded': [],
            'sources': [],
            'cache': {'refreshed_at': '13/05/2026 07:30:00'},
        },
        'wide_area_groups': [{'parent_port_key': 'OLT1'}],
        'pattern_exclusions': [{'ma_tb': 'TB2'}],
        'port_down_groups': [{'parent_port_key': 'OLT2'}],
    }


def test_quangchudong_page_has_off_today_6am_duration_filter():
    client = app.test_client()
    with client.session_transaction() as session:
        session['username'] = 'test-user'

    response = client.get('/quangchudong')

    assert response.status_code == 200
    html = response.data.decode('utf-8')
    assert '<option value="off-today-6am">OFF từ 06h hôm nay</option>' in html
    assert 'data-event-time="${row.first_off_time || row.alert_time || \'\'}"' in html
    assert 'parseLocalDateTime' in html
    assert 'ticket-today-6am' not in html


class _FakeNvktQuangChuDongCache:
    def get_dashboard_payload(self):
        return {
            'active': {
                'alerts': [
                    {
                        'ma_tb': 'AFTER_FIRST_OFF',
                        'ten_nvkt_db': 'Lê Văn Tuấn',
                        'first_off_time': '2026-05-13 06:00:00',
                        'alert_time': '2026-05-13 05:55:00',
                    },
                    {
                        'ma_tb': 'AFTER_ALERT',
                        'ten_nvkt_db': 'Le Van Tuan',
                        'first_off_time': '',
                        'alert_time': '2026-05-13 07:15:00',
                    },
                    {
                        'ma_tb': 'BEFORE_CUTOFF',
                        'ten_nvkt_db': 'Lê Văn Tuấn',
                        'first_off_time': '2026-05-13 05:59:59',
                    },
                    {
                        'ma_tb': 'OTHER_NVKT',
                        'ten_nvkt_db': 'Nguyễn Văn A',
                        'first_off_time': '2026-05-13 08:00:00',
                    },
                    {
                        'ma_tb': 'YESTERDAY',
                        'ten_nvkt_db': 'Lê Văn Tuấn',
                        'first_off_time': '2026-05-12 23:00:00',
                    },
                ],
                'excluded': [],
                'sources': [],
                'cache': {'refreshed_at': '13/05/2026 08:00:00'},
            },
            'wide_area_groups': [],
            'pattern_exclusions': [],
            'port_down_groups': [],
        }


def test_quangchudong_nvkt_api_filters_by_slug_and_today_6am(monkeypatch):
    monkeypatch.setattr(
        quangchudong_routes,
        'get_quangchudong_cache',
        lambda: _FakeNvktQuangChuDongCache(),
    )
    monkeypatch.setattr(quangchudong_routes, 'datetime', _FixedDateTime)

    client = app.test_client()
    with client.session_transaction() as session:
        session['username'] = 'test-user'

    response = client.get('/api/quangchudong/nvkt/le-van-tuan/alerts')

    assert response.status_code == 200
    payload = response.get_json()
    assert payload['nvkt_slug'] == 'le-van-tuan'
    assert payload['nvkt_name'] == 'Lê Văn Tuấn'
    assert payload['cutoff_time'] == '2026-05-13 06:00:00'
    assert [row['ma_tb'] for row in payload['alerts']] == ['AFTER_ALERT', 'AFTER_FIRST_OFF']
    assert payload['count'] == 2
    assert payload['cache']['refreshed_at'] == '13/05/2026 08:00:00'


def test_quangchudong_nvkt_mobile_page_is_standalone(monkeypatch):
    monkeypatch.setattr(quangchudong_routes, 'datetime', _FixedDateTime)

    client = app.test_client()
    with client.session_transaction() as session:
        session['username'] = 'test-user'

    response = client.get('/quangchudong/le-van-tuan')

    assert response.status_code == 200
    html = response.data.decode('utf-8')
    assert '/api/quangchudong/nvkt/le-van-tuan/alerts' in html
    assert 'data-nvkt-slug="le-van-tuan"' in html
    assert 'href="tel:${escapeHtml(href)}"' in html
    assert 'sidebar' not in html.lower()


def test_quangchudong_page_has_new_off_today_and_sleeping_tabs():
    client = app.test_client()
    with client.session_transaction() as session:
        session['username'] = 'test-user'

    response = client.get('/quangchudong')

    assert response.status_code == 200
    html = response.data.decode('utf-8')
    assert 'Cảnh báo OFF trong ngày' in html
    assert 'Thuê bao ngủ theo ngày' in html
    assert 'id="off-today-table"' in html
    assert 'id="sleeping-table"' in html
    assert 'Thời gian ngủ' in html
    assert 'renderOffTodayAlerts' in html
    assert 'renderSleepingSubscribers' in html


def test_quangchudong_helpers_split_off_today_and_sleeping_rows():
    rows = [
        {'ma_tb': 'AFTER', 'first_off_time': '2026-05-13T06:00:00'},
        {'ma_tb': 'ALERT_FALLBACK', 'first_off_time': '', 'alert_time': '2026-05-13 07:15:00'},
        {'ma_tb': 'BEFORE', 'first_off_time': '2026-05-12 05:30:00'},
        {'ma_tb': 'BAD_TIME', 'first_off_time': 'not-a-date'},
    ]
    cutoff = datetime(2026, 5, 13, 6, 0, 0)
    now = datetime(2026, 5, 15, 8, 30, 0)

    off_today = quangchudong_routes._filter_off_today_rows(rows, cutoff)
    sleeping = quangchudong_routes._filter_sleeping_rows(rows, cutoff, now)

    assert [row['ma_tb'] for row in off_today] == ['ALERT_FALLBACK', 'AFTER']
    assert [row['ma_tb'] for row in sleeping] == ['BEFORE']
    assert sleeping[0]['thoi_gian_ngu'] == 3


class _FakeExportQuangChuDongCache:
    def get_dashboard_payload(self):
        return {
            'active': {
                'alerts': [
                    {
                        'ma_tb': 'TODAY',
                        'ten_tb': 'Trong ngày',
                        'first_off_time': '2026-05-13 07:00:00',
                        'doi_vt': 'Tổ A',
                    },
                    {
                        'ma_tb': 'SLEEP',
                        'ten_tb': 'Ngủ dài',
                        'first_off_time': '2026-05-10 05:00:00',
                        'doi_vt': 'Tổ B',
                    },
                ],
                'excluded': [],
                'sources': [],
                'cache': {'refreshed_at': '13/05/2026 08:00:00'},
            },
            'wide_area_groups': [{'parent_port_key': 'PARENT1', 'down_count': 3}],
            'pattern_exclusions': [{'ma_tb': 'PATTERN1', 'exclude_reason': 'Trong danh sách tắt chủ động'}],
            'port_down_groups': [{'parent_port_key': 'PORT1', 'down_count': 5}],
        }


def test_quangchudong_download_excel_contains_all_tab_sheets(monkeypatch):
    monkeypatch.setattr(
        quangchudong_routes,
        'get_quangchudong_cache',
        lambda: _FakeExportQuangChuDongCache(),
    )
    monkeypatch.setattr(quangchudong_routes, 'datetime', _FixedDateTime)

    client = app.test_client()
    with client.session_transaction() as session:
        session['username'] = 'test-user'

    response = client.get('/download/quangchudong-report')

    assert response.status_code == 200
    assert response.mimetype == 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    assert response.headers['Content-Disposition'].startswith('attachment;')

    workbook = pd.ExcelFile(io.BytesIO(response.data))
    assert workbook.sheet_names == [
        'Canh_bao_DOWN',
        'OFF_trong_ngay',
        'Thue_bao_ngu',
        'Su_co_dien_rong',
        'Pattern_exclusion',
        'Port_OLT_Down',
    ]

    off_today = pd.read_excel(workbook, sheet_name='OFF_trong_ngay')
    sleeping = pd.read_excel(workbook, sheet_name='Thue_bao_ngu')
    assert off_today['Mã TB'].tolist() == ['TODAY']
    assert sleeping['Mã TB'].tolist() == ['SLEEP']
    assert sleeping['Thời gian ngủ'].tolist() == [3]


class _FixedDateTime(datetime):
    @classmethod
    def now(cls, tz=None):
        return cls(2026, 5, 13, 8, 30, 0)
