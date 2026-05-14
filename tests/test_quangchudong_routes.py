from pathlib import Path
import sys
from datetime import datetime

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


class _FixedDateTime(datetime):
    @classmethod
    def now(cls, tz=None):
        return cls(2026, 5, 13, 8, 30, 0)
