import sqlite3
from datetime import datetime

from flask import Blueprint, jsonify, render_template, session

from auth import get_user_by_username, login_required

SA_OUTAGE_DB_PATH = '/home/vtst/one-suachua/database/sa_outate_db.db'

sa_outage_bp = Blueprint('sa_outage', __name__)


def _current_user():
    username = session.get('username')
    return get_user_by_username(username) if username else None


def _query_db(sql, params=()):
    """Kết nối SQLite và trả về list of dict."""
    conn = sqlite3.connect(SA_OUTAGE_DB_PATH)
    conn.row_factory = sqlite3.Row
    try:
        cur = conn.execute(sql, params)
        rows = [dict(r) for r in cur.fetchall()]
    finally:
        conn.close()
    return rows


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
        # ── 1. Sự cố đang tồn ──────────────────────────────────────────────
        dang_ton = _query_db("""
            SELECT
                incident_code,
                sa_code,
                doi_vt,
                trang_thai,
                started_at,
                last_seen_at,
                ended_at,
                thoi_gian_keo_dai_phut,
                start_off_count,
                latest_off_count,
                max_off_count,
                latest_total_count,
                latest_off_percentage,
                recovery_reason,
                notes,
                source,
                confidence
            FROM v_sa_outage_monitoring
            WHERE ended_at IS NULL
            ORDER BY started_at DESC
        """)

        # ── 2. Sự cố đã clear hôm nay ──────────────────────────────────────
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
                source,
                confidence
            FROM sa_outage_incidents
            WHERE status = 'CLOSED'
              AND date(ended_at, 'localtime') = date('now', 'localtime')
            ORDER BY ended_at DESC
        """)

        return jsonify({
            'ok': True,
            'generated_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'dang_ton': dang_ton,
            'da_clear_hom_nay': da_clear,
        })

    except Exception as exc:
        return jsonify({'ok': False, 'error': str(exc)}), 500
