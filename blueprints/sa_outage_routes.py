import os
import sqlite3
import time
from datetime import datetime
from threading import Lock

from flask import Blueprint, jsonify, render_template, request, session

from auth import get_user_by_username, login_required

SA_OUTAGE_DB_PATH = '/home/vtst/1bss/runtime/default/sqlite/sa_outage.db'
SA_OUTAGE_CACHE_SECONDS = 30

sa_outage_bp = Blueprint('sa_outage', __name__)

_sa_outage_cache_lock = Lock()
_sa_outage_cache = None
_sa_outage_cache_ts = 0
_sa_outage_schema_lock = Lock()
_sa_outage_schema_ready_path = None


def _current_user():
    username = session.get('username')
    return get_user_by_username(username) if username else None


def _readonly_uri(db_path):
    return f'file:{os.path.abspath(db_path)}?mode=ro'


def _query_db(sql, params=()):
    """Kết nối SQLite read-only và trả về list of dict."""
    conn = sqlite3.connect(_readonly_uri(SA_OUTAGE_DB_PATH), uri=True, timeout=5)
    conn.execute('PRAGMA busy_timeout=5000')
    conn.row_factory = sqlite3.Row
    try:
        cur = conn.execute(sql, params)
        rows = [dict(r) for r in cur.fetchall()]
    finally:
        conn.close()
    return rows


def _write_connection():
    conn = sqlite3.connect(SA_OUTAGE_DB_PATH, timeout=5)
    conn.execute('PRAGMA journal_mode=WAL')
    conn.execute('PRAGMA busy_timeout=5000')
    conn.row_factory = sqlite3.Row
    return conn


def _ensure_user_notes_schema():
    global _sa_outage_schema_ready_path
    db_path = os.path.abspath(SA_OUTAGE_DB_PATH)
    with _sa_outage_schema_lock:
        if _sa_outage_schema_ready_path == db_path:
            return
        with _write_connection() as conn:
            columns = {
                row['name']
                for row in conn.execute('PRAGMA table_info(sa_outage_incidents)').fetchall()
            }
            if 'user_notes' not in columns:
                _add_column_if_missing(conn, 'user_notes', 'TEXT')
            if 'user_notes_updated_at' not in columns:
                _add_column_if_missing(conn, 'user_notes_updated_at', 'TEXT')
            if 'user_notes_updated_by' not in columns:
                _add_column_if_missing(conn, 'user_notes_updated_by', 'TEXT')
        _sa_outage_schema_ready_path = db_path


def _add_column_if_missing(conn, column_name, column_type):
    try:
        conn.execute(f'ALTER TABLE sa_outage_incidents ADD COLUMN {column_name} {column_type}')
    except sqlite3.OperationalError as exc:
        if 'duplicate column name' not in str(exc).lower():
            raise


def _invalidate_sa_outage_cache():
    global _sa_outage_cache, _sa_outage_cache_ts
    with _sa_outage_cache_lock:
        _sa_outage_cache = None
        _sa_outage_cache_ts = 0


def _default_month():
    return datetime.now().strftime('%Y-%m')


def _normalize_month(month):
    if not month:
        return _default_month()
    try:
        parsed = datetime.strptime(month, '%Y-%m')
    except ValueError:
        return _default_month()
    return parsed.strftime('%Y-%m')


def _get_sa_outage_data(month=None):
    global _sa_outage_cache, _sa_outage_cache_ts
    selected_month = _normalize_month(month)
    cache_key = selected_month
    now = time.time()
    with _sa_outage_cache_lock:
        if (
            _sa_outage_cache is not None
            and _sa_outage_cache.get('cache_key') == cache_key
            and (now - _sa_outage_cache_ts) < SA_OUTAGE_CACHE_SECONDS
        ):
            return _sa_outage_cache

    _ensure_user_notes_schema()

    dang_ton = _query_db("""
        SELECT
            id,
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
            COALESCE(user_notes, '') AS user_notes,
            user_notes_updated_at,
            user_notes_updated_by,
            'runtime' AS source,
            'exact' AS confidence
        FROM sa_outage_incidents
        WHERE ended_at IS NULL
        ORDER BY started_at DESC
    """)

    da_clear = _query_db("""
        SELECT
            id,
            incident_code,
            sa_code,
            doi_vt,
            status,
            started_at,
            ended_at,
            duration_minutes,
            recovery_reason,
            notes,
            COALESCE(user_notes, '') AS user_notes,
            user_notes_updated_at,
            user_notes_updated_by,
            'runtime' AS source,
            'exact' AS confidence
        FROM sa_outage_incidents
        WHERE status = 'CLOSED'
          AND date(ended_at, 'localtime') = date('now', 'localtime')
        ORDER BY ended_at DESC
    """)

    monthly_incidents = _query_db("""
        SELECT
            id,
            incident_code,
            sa_code,
            doi_vt,
            CASE
                WHEN status = 'OPEN' THEN 'Đang tồn'
                ELSE 'Đã kết thúc'
            END AS trang_thai,
            status,
            started_at,
            ended_at,
            date(started_at, 'localtime') AS incident_date,
            max_off_count AS affected_subscribers,
            duration_minutes,
            recovery_reason,
            notes,
            COALESCE(user_notes, '') AS user_notes,
            user_notes_updated_at,
            user_notes_updated_by
        FROM sa_outage_incidents
        WHERE strftime('%Y-%m', started_at) = ?
        ORDER BY started_at DESC
    """, (selected_month,))

    monthly_summary = {
        'incident_count': len(monthly_incidents),
        'total_max_off_count': sum(row.get('affected_subscribers') or 0 for row in monthly_incidents),
    }

    result = {
        'ok': True,
        'generated_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'cache_key': cache_key,
        'dang_ton': dang_ton,
        'da_clear_hom_nay': da_clear,
        'monthly_report': {
            'selected_month': selected_month,
            'summary': monthly_summary,
            'incidents': monthly_incidents,
        },
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
        return jsonify(_get_sa_outage_data(request.args.get('month')))
    except Exception as exc:
        return jsonify({'ok': False, 'error': str(exc)}), 500


@sa_outage_bp.route('/api/su-co-sa/incidents/<int:incident_id>/notes', methods=['POST'])
@login_required
def api_update_su_co_sa_user_notes(incident_id):
    payload = request.get_json(silent=True) or {}
    user_notes = str(payload.get('user_notes') or '').strip()
    if len(user_notes) > 2000:
        return jsonify({'ok': False, 'error': 'Ghi chú không được vượt quá 2000 ký tự'}), 400

    try:
        _ensure_user_notes_schema()
        with _write_connection() as conn:
            cursor = conn.execute(
                """
                UPDATE sa_outage_incidents
                SET user_notes = ?,
                    user_notes_updated_at = datetime('now', 'localtime'),
                    user_notes_updated_by = ?,
                    updated_at = datetime('now', 'localtime')
                WHERE id = ?
                """,
                (user_notes, session.get('username') or '', incident_id),
            )
        if cursor.rowcount == 0:
            return jsonify({'ok': False, 'error': 'Không tìm thấy sự cố SA'}), 404
        _invalidate_sa_outage_cache()
        return jsonify({'ok': True, 'incident_id': incident_id, 'user_notes': user_notes})
    except Exception as exc:
        return jsonify({'ok': False, 'error': str(exc)}), 500
