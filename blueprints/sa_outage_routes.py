import os
import sqlite3
import time
from datetime import datetime
from threading import Lock

from flask import Blueprint, jsonify, render_template, session

from auth import get_user_by_username, login_required

SA_OUTAGE_DB_PATH = '/home/vtst/1bss/runtime/default/sqlite/sa_outage.db'
SA_OUTAGE_CACHE_SECONDS = 30

sa_outage_bp = Blueprint('sa_outage', __name__)

_sa_outage_cache_lock = Lock()
_sa_outage_cache = None
_sa_outage_cache_ts = 0


def _current_user():
    username = session.get('username')
    return get_user_by_username(username) if username else None


def _readonly_uri(db_path):
    return f'file:{os.path.abspath(db_path)}?mode=ro&immutable=1'


def _query_db(sql, params=()):
    """Kết nối SQLite read-only và trả về list of dict."""
    conn = sqlite3.connect(_readonly_uri(SA_OUTAGE_DB_PATH), uri=True)
    conn.row_factory = sqlite3.Row
    try:
        cur = conn.execute(sql, params)
        rows = [dict(r) for r in cur.fetchall()]
    finally:
        conn.close()
    return rows


def _get_sa_outage_data():
    global _sa_outage_cache, _sa_outage_cache_ts
    now = time.time()
    with _sa_outage_cache_lock:
        if _sa_outage_cache is not None and (now - _sa_outage_cache_ts) < SA_OUTAGE_CACHE_SECONDS:
            return _sa_outage_cache

    dang_ton = _query_db("""
        SELECT
            incident_code,
            sa_code,
            doi_vt,
            CASE
                WHEN status = 'OPEN' THEN 'Đang tồn'
                ELSE 'Đã kết thúc'
            END AS trang_thai,
            started_at,
            last_seen_at,
            ended_at,
            CASE
                WHEN ended_at IS NULL THEN CAST((julianday('now', 'localtime') - julianday(started_at)) * 24 * 60 AS INTEGER)
                ELSE duration_minutes
            END AS thoi_gian_keo_dai_phut,
            start_off_count,
            latest_off_count,
            max_off_count,
            latest_total_count,
            ROUND(
                CASE
                    WHEN latest_total_count > 0
                        THEN (latest_off_count * 100.0) / latest_total_count
                    ELSE 0
                END,
                1
            ) AS latest_off_percentage,
            recovery_reason,
            notes,
            'runtime' AS source,
            'exact' AS confidence
        FROM sa_outage_incidents
        WHERE ended_at IS NULL
        ORDER BY started_at DESC
    """)

    da_clear = _query_db("""
        SELECT
            incident_code,
            sa_code,
            doi_vt,
            status,
            started_at,
            ended_at,
            duration_minutes,
            recovery_reason,
            notes,
            'runtime' AS source,
            'exact' AS confidence
        FROM sa_outage_incidents
        WHERE status = 'CLOSED'
          AND date(ended_at, 'localtime') = date('now', 'localtime')
        ORDER BY ended_at DESC
    """)

    result = {
        'ok': True,
        'generated_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'dang_ton': dang_ton,
        'da_clear_hom_nay': da_clear,
    }

    with _sa_outage_cache_lock:
        _sa_outage_cache = result
        _sa_outage_cache_ts = now

    return result


@sa_outage_bp.route('/su_co_sa')
@login_required
def page_su_co_sa():
    return render_template(
        'pages/su_co_sa.html',
        current_user=_current_user(),
        active_page='su_co_sa',
    )


@sa_outage_bp.route('/api/su-co-sa/data')
@login_required
def api_su_co_sa_data():
    try:
        return jsonify(_get_sa_outage_data())
    except Exception as exc:
        return jsonify({'ok': False, 'error': str(exc)}), 500
