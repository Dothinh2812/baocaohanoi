import os
import sqlite3
from datetime import datetime

from config import (
    SHC_PROCESSING_RESULTS_DB_PATH,
    SHC_SOURCE_K1_DB_PATH,
    SHC_SOURCE_K2_DB_PATH,
)


UNKNOWN_VALUE = 'Chua xac dinh'


def _connect(db_path):
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    return conn


def _require_db(db_path, label):
    if not db_path or not os.path.exists(db_path):
        raise FileNotFoundError(f'Không tìm thấy {label}: {db_path or "chưa cấu hình"}')


def _row_to_counts(row):
    if not row:
        return {'k1_count': 0, 'k2_count': 0, 'total_count': 0}
    return {
        'k1_count': int(row['k1_count'] or 0),
        'k2_count': int(row['k2_count'] or 0),
        'total_count': int(row['total_count'] or 0),
    }


def _normalize_team_filter(team_filter):
    return str(team_filter or '').strip()


def _max_date_from_db(db_path, table_name, column_name):
    if not db_path or not os.path.exists(db_path):
        return ''

    conn = _connect(db_path)
    try:
        row = conn.execute(f'SELECT MAX({column_name}) AS max_date FROM {table_name}').fetchone()
        return str(row['max_date'] or '')
    finally:
        conn.close()


def get_default_report_date():
    candidates = [
        _max_date_from_db(SHC_PROCESSING_RESULTS_DB_PATH, 'processing_results', 'processed_date'),
        _max_date_from_db(SHC_SOURCE_K1_DB_PATH, 'suy_hao_snapshots', 'ngay_bao_cao'),
        _max_date_from_db(SHC_SOURCE_K2_DB_PATH, 'suy_hao_snapshots', 'ngay_bao_cao'),
    ]
    candidates = [value for value in candidates if value]
    if candidates:
        return max(candidates)
    return datetime.now().strftime('%Y-%m-%d')


def _load_required_rows(db_path, report_date, ticket_type):
    conn = _connect(db_path)
    try:
        rows = conn.execute(
            """
            SELECT
                COALESCE(NULLIF(TRIM(doi_one), ''), ?) AS doi_one,
                COALESCE(NULLIF(TRIM(nvkt_db_normalized), ''), ?) AS user_name,
                COUNT(*) AS required_count
            FROM suy_hao_snapshots
            WHERE ngay_bao_cao = ?
            GROUP BY doi_one, nvkt_db_normalized
            ORDER BY doi_one ASC, user_name ASC
            """,
            (UNKNOWN_VALUE, UNKNOWN_VALUE, report_date),
        ).fetchall()

        return [
            {
                'doi_one': row['doi_one'],
                'user_name': row['user_name'],
                'ticket_type': ticket_type,
                'required_count': int(row['required_count'] or 0),
            }
            for row in rows
        ]
    finally:
        conn.close()


def _load_source_required_counts(report_date):
    _require_db(SHC_SOURCE_K1_DB_PATH, 'DB nguồn suy hao K1')
    _require_db(SHC_SOURCE_K2_DB_PATH, 'DB nguồn suy hao K2')

    k1_rows = _load_required_rows(SHC_SOURCE_K1_DB_PATH, report_date, 'K1')
    k2_rows = _load_required_rows(SHC_SOURCE_K2_DB_PATH, report_date, 'K2')

    user_map = {}
    team_map = {}
    summary = {
        'k1_required_count': 0,
        'k2_required_count': 0,
        'total_required_count': 0,
    }

    for row in k1_rows + k2_rows:
        key = (str(row['doi_one']), str(row['user_name']))
        if key not in user_map:
            user_map[key] = {
                'doi_one': key[0],
                'user_name': key[1],
                'k1_required_count': 0,
                'k2_required_count': 0,
                'total_required_count': 0,
            }
        if key[0] not in team_map:
            team_map[key[0]] = {
                'doi_one': key[0],
                'k1_required_count': 0,
                'k2_required_count': 0,
                'total_required_count': 0,
            }

        required_count = int(row['required_count'] or 0)
        if row['ticket_type'] == 'K1':
            user_map[key]['k1_required_count'] += required_count
            team_map[key[0]]['k1_required_count'] += required_count
            summary['k1_required_count'] += required_count
        else:
            user_map[key]['k2_required_count'] += required_count
            team_map[key[0]]['k2_required_count'] += required_count
            summary['k2_required_count'] += required_count

        user_map[key]['total_required_count'] += required_count
        team_map[key[0]]['total_required_count'] += required_count
        summary['total_required_count'] += required_count

    return {
        'summary': summary,
        'user_map': user_map,
        'team_map': team_map,
    }


def get_shc_processing_report(report_date=None, team_filter=''):
    _require_db(SHC_PROCESSING_RESULTS_DB_PATH, 'DB nội bộ kết quả xử lý SHC')

    resolved_report_date = str(report_date or '').strip() or get_default_report_date()
    resolved_report_month = resolved_report_date[:7]
    resolved_team_filter = _normalize_team_filter(team_filter)

    conn = _connect(SHC_PROCESSING_RESULTS_DB_PATH)
    try:
        source_required_counts = _load_source_required_counts(resolved_report_date)

        daily_summary_row = conn.execute(
            """
            SELECT
                COALESCE(SUM(CASE WHEN k1_status = 'dat' THEN 1 ELSE 0 END), 0) AS k1_count,
                COALESCE(SUM(CASE WHEN k2_status = 'dat' THEN 1 ELSE 0 END), 0) AS k2_count,
                COALESCE(
                    SUM(CASE WHEN k1_status = 'dat' THEN 1 ELSE 0 END) +
                    SUM(CASE WHEN k2_status = 'dat' THEN 1 ELSE 0 END),
                    0
                ) AS total_count
            FROM processing_results
            WHERE processed_date = ?
            """,
            (resolved_report_date,),
        ).fetchone()

        monthly_summary_row = conn.execute(
            """
            SELECT
                COALESCE(SUM(CASE WHEN k1_status = 'dat' THEN 1 ELSE 0 END), 0) AS k1_count,
                COALESCE(SUM(CASE WHEN k2_status = 'dat' THEN 1 ELSE 0 END), 0) AS k2_count,
                COALESCE(
                    SUM(CASE WHEN k1_status = 'dat' THEN 1 ELSE 0 END) +
                    SUM(CASE WHEN k2_status = 'dat' THEN 1 ELSE 0 END),
                    0
                ) AS total_count
            FROM processing_results
            WHERE SUBSTR(processed_date, 1, 7) = ?
            """,
            (resolved_report_month,),
        ).fetchone()

        daily_user_rows = conn.execute(
            """
            SELECT
                COALESCE(NULLIF(TRIM(user_name), ''), ?) AS user_name,
                COALESCE(NULLIF(TRIM(doi_one), ''), ?) AS doi_one,
                COALESCE(SUM(CASE WHEN k1_status = 'dat' THEN 1 ELSE 0 END), 0) AS k1_count,
                COALESCE(SUM(CASE WHEN k2_status = 'dat' THEN 1 ELSE 0 END), 0) AS k2_count,
                COALESCE(
                    SUM(CASE WHEN k1_status = 'dat' THEN 1 ELSE 0 END) +
                    SUM(CASE WHEN k2_status = 'dat' THEN 1 ELSE 0 END),
                    0
                ) AS total_count
            FROM processing_results
            WHERE processed_date = ?
              AND (k1_status = 'dat' OR k2_status = 'dat')
            GROUP BY user_name, doi_one
            ORDER BY doi_one ASC, total_count DESC, k1_count DESC, k2_count DESC, user_name ASC
            """,
            (UNKNOWN_VALUE, UNKNOWN_VALUE, resolved_report_date),
        ).fetchall()

        daily_team_rows = conn.execute(
            """
            SELECT
                COALESCE(NULLIF(TRIM(doi_one), ''), ?) AS doi_one,
                COALESCE(SUM(CASE WHEN k1_status = 'dat' THEN 1 ELSE 0 END), 0) AS k1_count,
                COALESCE(SUM(CASE WHEN k2_status = 'dat' THEN 1 ELSE 0 END), 0) AS k2_count,
                COALESCE(
                    SUM(CASE WHEN k1_status = 'dat' THEN 1 ELSE 0 END) +
                    SUM(CASE WHEN k2_status = 'dat' THEN 1 ELSE 0 END),
                    0
                ) AS total_count
            FROM processing_results
            WHERE processed_date = ?
              AND (k1_status = 'dat' OR k2_status = 'dat')
            GROUP BY doi_one
            ORDER BY total_count DESC, k1_count DESC, k2_count DESC, doi_one ASC
            """,
            (UNKNOWN_VALUE, resolved_report_date),
        ).fetchall()

        monthly_team_rows = conn.execute(
            """
            SELECT
                COALESCE(NULLIF(TRIM(doi_one), ''), ?) AS doi_one,
                COALESCE(SUM(CASE WHEN k1_status = 'dat' THEN 1 ELSE 0 END), 0) AS k1_count,
                COALESCE(SUM(CASE WHEN k2_status = 'dat' THEN 1 ELSE 0 END), 0) AS k2_count,
                COALESCE(
                    SUM(CASE WHEN k1_status = 'dat' THEN 1 ELSE 0 END) +
                    SUM(CASE WHEN k2_status = 'dat' THEN 1 ELSE 0 END),
                    0
                ) AS total_count
            FROM processing_results
            WHERE SUBSTR(processed_date, 1, 7) = ?
              AND (k1_status = 'dat' OR k2_status = 'dat')
            GROUP BY doi_one
            ORDER BY total_count DESC, k1_count DESC, k2_count DESC, doi_one ASC
            """,
            (UNKNOWN_VALUE, resolved_report_month),
        ).fetchall()

        monthly_day_rows = conn.execute(
            """
            SELECT
                processed_date,
                COALESCE(SUM(CASE WHEN k1_status = 'dat' THEN 1 ELSE 0 END), 0) AS k1_count,
                COALESCE(SUM(CASE WHEN k2_status = 'dat' THEN 1 ELSE 0 END), 0) AS k2_count,
                COALESCE(
                    SUM(CASE WHEN k1_status = 'dat' THEN 1 ELSE 0 END) +
                    SUM(CASE WHEN k2_status = 'dat' THEN 1 ELSE 0 END),
                    0
                ) AS total_count
            FROM processing_results
            WHERE SUBSTR(processed_date, 1, 7) = ?
              AND (k1_status = 'dat' OR k2_status = 'dat')
            GROUP BY processed_date
            ORDER BY processed_date DESC
            """,
            (resolved_report_month,),
        ).fetchall()

        merged_daily_user_rows = []
        source_user_map = dict(source_required_counts['user_map'])
        processed_user_map = {
            (str(row['doi_one']), str(row['user_name'])): dict(row)
            for row in daily_user_rows
        }
        for key in sorted(set(source_user_map.keys()) | set(processed_user_map.keys()), key=lambda item: (item[0], item[1])):
            processed_row = processed_user_map.get(key, {})
            required_row = source_user_map.get(
                key,
                {
                    'doi_one': key[0],
                    'user_name': key[1],
                    'k1_required_count': 0,
                    'k2_required_count': 0,
                    'total_required_count': 0,
                },
            )
            merged_row = {
                'doi_one': key[0],
                'user_name': key[1],
                'k1_count': int(processed_row.get('k1_count', 0) or 0),
                'k2_count': int(processed_row.get('k2_count', 0) or 0),
                'total_count': int(processed_row.get('total_count', 0) or 0),
                'k1_required_count': int(required_row.get('k1_required_count', 0) or 0),
                'k2_required_count': int(required_row.get('k2_required_count', 0) or 0),
                'total_required_count': int(required_row.get('total_required_count', 0) or 0),
            }
            merged_row['highlight_warning'] = (
                merged_row['total_count'] < 3 and merged_row['total_required_count'] >= 3
            )
            merged_daily_user_rows.append(merged_row)

        merged_daily_team_rows = []
        source_team_map = dict(source_required_counts['team_map'])
        processed_team_map = {
            str(row['doi_one']): dict(row)
            for row in daily_team_rows
        }
        for team_name in sorted(set(source_team_map.keys()) | set(processed_team_map.keys())):
            processed_row = processed_team_map.get(team_name, {})
            required_row = source_team_map.get(
                team_name,
                {
                    'doi_one': team_name,
                    'k1_required_count': 0,
                    'k2_required_count': 0,
                    'total_required_count': 0,
                },
            )
            merged_row = {
                'doi_one': team_name,
                'k1_count': int(processed_row.get('k1_count', 0) or 0),
                'k2_count': int(processed_row.get('k2_count', 0) or 0),
                'total_count': int(processed_row.get('total_count', 0) or 0),
                'k1_required_count': int(required_row.get('k1_required_count', 0) or 0),
                'k2_required_count': int(required_row.get('k2_required_count', 0) or 0),
                'total_required_count': int(required_row.get('total_required_count', 0) or 0),
            }
            merged_row['highlight_warning'] = (
                merged_row['total_count'] < 3 and merged_row['total_required_count'] >= 3
            )
            merged_daily_team_rows.append(merged_row)

        team_options = sorted(
            {
                row['doi_one']
                for row in merged_daily_team_rows
                if str(row.get('doi_one') or '').strip()
            }
        )

        if resolved_team_filter:
            merged_daily_user_rows = [
                row for row in merged_daily_user_rows
                if row['doi_one'] == resolved_team_filter
            ]
            merged_daily_team_rows = [
                row for row in merged_daily_team_rows
                if row['doi_one'] == resolved_team_filter
            ]
            monthly_team_rows = [
                dict(row) for row in monthly_team_rows
                if dict(row).get('doi_one') == resolved_team_filter
            ]
        else:
            monthly_team_rows = [dict(row) for row in monthly_team_rows]

        return {
            'selected_date': resolved_report_date,
            'selected_month': resolved_report_month,
            'selected_team_filter': resolved_team_filter,
            'team_options': team_options,
            'daily_summary': _row_to_counts(daily_summary_row),
            'daily_required_summary': source_required_counts['summary'],
            'monthly_summary': _row_to_counts(monthly_summary_row),
            'daily_user_rows': merged_daily_user_rows,
            'daily_team_rows': merged_daily_team_rows,
            'monthly_team_rows': monthly_team_rows,
            'monthly_day_rows': [dict(row) for row in monthly_day_rows],
        }
    finally:
        conn.close()
