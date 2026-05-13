from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from dashboard import app
from blueprints import quangchudong_routes


class _FakeQuangChuDongCache:
    def get_dashboard_payload(self):
        return {
            'active': {'alerts': [{'ma_tb': 'TB1'}], 'excluded': [], 'sources': []},
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
        'active': {'alerts': [{'ma_tb': 'TB1'}], 'excluded': [], 'sources': []},
        'wide_area_groups': [{'parent_port_key': 'OLT1'}],
        'pattern_exclusions': [{'ma_tb': 'TB2'}],
        'port_down_groups': [{'parent_port_key': 'OLT2'}],
    }
