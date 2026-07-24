import sqlite3
from datetime import datetime
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from dashboard import app
from blueprints import sa_outage_routes


SCHEMA = """
CREATE TABLE sa_outage_incidents (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    incident_code TEXT NOT NULL UNIQUE,
    sa_code TEXT NOT NULL,
    doi_vt TEXT,
    status TEXT NOT NULL,
    started_at TEXT NOT NULL,
    last_seen_at TEXT NOT NULL,
    ended_at TEXT,
    started_cycle_id TEXT,
    last_cycle_id TEXT,
    ended_cycle_id TEXT,
    start_off_count INTEGER NOT NULL,
    latest_off_count INTEGER NOT NULL,
    max_off_count INTEGER NOT NULL,
    latest_total_count INTEGER NOT NULL,
    latest_valid_count INTEGER NOT NULL,
    latest_valid_ratio REAL NOT NULL,
    threshold INTEGER NOT NULL,
    duration_minutes INTEGER,
    recovery_reason TEXT,
    notes TEXT,
    created_at TEXT NOT NULL DEFAULT (datetime('now')),
    updated_at TEXT NOT NULL DEFAULT (datetime('now'))
);
"""


class _FixedDateTime(datetime):
    @classmethod
    def now(cls, tz=None):
        return cls(2026, 5, 16, 10, 0, 0)


def _insert_incident(conn, code, sa, status, started_at, ended_at, max_off_count):
    conn.execute(
        """
        INSERT INTO sa_outage_incidents (
            incident_code, sa_code, doi_vt, status, started_at, last_seen_at, ended_at,
            start_off_count, latest_off_count, max_off_count, latest_total_count,
            latest_valid_count, latest_valid_ratio, threshold, duration_minutes,
            recovery_reason, notes
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            code,
            sa,
            'ToKT_SonTay',
            status,
            started_at,
            ended_at or started_at,
            ended_at,
            3,
            1,
            max_off_count,
            20,
            20,
            1.0,
            5,
            90 if ended_at else None,
            'auto clear' if ended_at else None,
            '',
        ),
    )


def _prepare_sa_outage_db(tmp_path, monkeypatch):
    db_path = tmp_path / 'sa_outage.db'
    conn = sqlite3.connect(db_path)
    conn.executescript(SCHEMA)
    _insert_incident(
        conn,
        'SA-MAY-1',
        'STY.G51_5/2',
        'CLOSED',
        '2026-05-10 08:00:00',
        '2026-05-10 09:30:00',
        12,
    )
    _insert_incident(
        conn,
        'SA-MAY-2',
        'SLC.G21_0/14',
        'OPEN',
        '2026-05-11 07:00:00',
        None,
        8,
    )
    _insert_incident(
        conn,
        'SA-APR-1',
        'BVI.G43_3/10',
        'CLOSED',
        '2026-04-20 12:00:00',
        '2026-04-20 13:00:00',
        30,
    )
    conn.commit()
    conn.close()

    monkeypatch.setattr(sa_outage_routes, 'SA_OUTAGE_DB_PATH', str(db_path))
    monkeypatch.setattr(sa_outage_routes, 'datetime', _FixedDateTime)
    sa_outage_routes._sa_outage_cache = None
    sa_outage_routes._sa_outage_cache_ts = 0
    return db_path


def _logged_in_client():
    client = app.test_client()
    with client.session_transaction() as session:
        session['username'] = 'test-user'
    return client


def test_su_co_sa_api_returns_default_monthly_report(tmp_path, monkeypatch):
    _prepare_sa_outage_db(tmp_path, monkeypatch)

    response = _logged_in_client().get('/api/su-co-sa/data')

    assert response.status_code == 200
    payload = response.get_json()
    report = payload['monthly_report']
    assert report['selected_month'] == '2026-05'
    assert report['summary'] == {
        'incident_count': 2,
        'total_max_off_count': 20,
    }
    assert [row['incident_code'] for row in report['incidents']] == ['SA-MAY-2', 'SA-MAY-1']
    assert report['incidents'][0]['affected_subscribers'] == 8
    assert report['incidents'][0]['incident_date'] == '2026-05-11'
    assert report['incidents'][0]['user_notes'] == ''


def test_su_co_sa_api_filters_monthly_report_by_query_month(tmp_path, monkeypatch):
    _prepare_sa_outage_db(tmp_path, monkeypatch)

    response = _logged_in_client().get('/api/su-co-sa/data?month=2026-04')

    assert response.status_code == 200
    report = response.get_json()['monthly_report']
    assert report['selected_month'] == '2026-04'
    assert report['summary'] == {
        'incident_count': 1,
        'total_max_off_count': 30,
    }
    assert [row['incident_code'] for row in report['incidents']] == ['SA-APR-1']


def test_su_co_sa_api_uses_current_month_for_invalid_query_month(tmp_path, monkeypatch):
    _prepare_sa_outage_db(tmp_path, monkeypatch)

    response = _logged_in_client().get('/api/su-co-sa/data?month=2026-13')

    assert response.status_code == 200
    report = response.get_json()['monthly_report']
    assert report['selected_month'] == '2026-05'
    assert [row['incident_code'] for row in report['incidents']] == ['SA-MAY-2', 'SA-MAY-1']


def test_su_co_sa_page_has_monthly_report_controls():
    response = _logged_in_client().get('/su_co_sa')

    assert response.status_code == 200
    html = response.data.decode('utf-8')
    assert 'id="monthReportFilter"' in html
    assert 'id="sectionMonthlyReport"' in html
    assert 'Thống kê tất cả sự cố theo tháng' in html
    assert 'Ghi chú xử lý' in html
    assert 'save-note-btn' in html


def test_su_co_sa_api_saves_user_notes_without_existing_schema_columns(tmp_path, monkeypatch):
    db_path = _prepare_sa_outage_db(tmp_path, monkeypatch)

    response = _logged_in_client().post(
        '/api/su-co-sa/incidents/1/notes',
        json={'user_notes': 'Nguyên nhân: mất điện. Đang phối hợp xử lý.'},
    )

    assert response.status_code == 200
    assert response.get_json() == {
        'ok': True,
        'incident_id': 1,
        'user_notes': 'Nguyên nhân: mất điện. Đang phối hợp xử lý.',
    }

    conn = sqlite3.connect(db_path)
    columns = {row[1] for row in conn.execute('PRAGMA table_info(sa_outage_incidents)')}
    row = conn.execute(
        'SELECT notes, user_notes, user_notes_updated_by FROM sa_outage_incidents WHERE id = 1'
    ).fetchone()
    conn.close()

    assert {'user_notes', 'user_notes_updated_at', 'user_notes_updated_by'} <= columns
    assert row == ('', 'Nguyên nhân: mất điện. Đang phối hợp xử lý.', 'test-user')

    monthly_response = _logged_in_client().get('/api/su-co-sa/data?month=2026-05')
    incident = monthly_response.get_json()['monthly_report']['incidents'][1]
    assert incident['id'] == 1
    assert incident['user_notes'] == 'Nguyên nhân: mất điện. Đang phối hợp xử lý.'
