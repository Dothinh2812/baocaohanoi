from __future__ import annotations

from datetime import date as date_cls

import pandas as pd

from .sqlite_runtime import read_sql_dataframe, read_sql_rows


_SUCCESS_IMPORT_STATUSES = ('thanh_cong', 'khong_co_sheet_tong_hop')


def _validate_iso_date(value):
    if not value:
        return None
    try:
        return date_cls.fromisoformat(str(value)).isoformat()
    except ValueError as exc:
        raise ValueError('Ngày không hợp lệ, cần theo định dạng YYYY-MM-DD') from exc


def get_available_dates(report_codes):
    report_codes = tuple(dict.fromkeys(code for code in report_codes if code))
    if not report_codes:
        return []

    placeholders = ', '.join('?' for _ in report_codes)
    status_placeholders = ', '.join('?' for _ in _SUCCESS_IMPORT_STATUSES)

    rows = read_sql_rows(
        f'''
        SELECT ngay_du_lieu
        FROM bao_cao_tong_hop_ngay
        WHERE ma_bao_cao IN ({placeholders})
          AND trang_thai_nap IN ({status_placeholders})
        GROUP BY ngay_du_lieu
        HAVING COUNT(DISTINCT ma_bao_cao) = ?
        ORDER BY ngay_du_lieu DESC
        ''',
        report_codes + _SUCCESS_IMPORT_STATUSES + (len(report_codes),),
    )
    return [row['ngay_du_lieu'] for row in rows if row.get('ngay_du_lieu')]


def resolve_date_context(requested_date, report_codes):
    normalized_requested_date = _validate_iso_date(requested_date)
    available_dates = get_available_dates(report_codes)
    latest_available_date = available_dates[0] if available_dates else None

    if normalized_requested_date:
        return {
            'selected_date': normalized_requested_date,
            'latest_available_date': latest_available_date,
            'available_dates': available_dates,
            'date_has_data': normalized_requested_date in available_dates,
        }

    return {
        'selected_date': latest_available_date,
        'latest_available_date': latest_available_date,
        'available_dates': available_dates,
        'date_has_data': bool(latest_available_date),
    }


def load_table_by_date(report_code, table_name, selected_date, *, order_by=None):
    normalized_date = _validate_iso_date(selected_date)
    if not normalized_date:
        return pd.DataFrame()

    order_clause = f' ORDER BY {order_by}' if order_by else ''
    return read_sql_dataframe(
        f'''
        SELECT t.*
        FROM "{table_name}" t
        JOIN sheet_bao_cao_tong_hop s
          ON s.id = t.__sheet_id
        JOIN bao_cao_tong_hop_ngay b
          ON b.id = s.bao_cao_tong_hop_ngay_id
        WHERE b.ma_bao_cao = ?
          AND s.ten_bang_du_lieu = ?
          AND b.ngay_du_lieu = ?
          AND b.trang_thai_nap IN ({', '.join('?' for _ in _SUCCESS_IMPORT_STATUSES)})
        {order_clause}
        ''',
        (report_code, table_name, normalized_date, *_SUCCESS_IMPORT_STATUSES),
    )


def load_many_tables_by_date(bindings, selected_date):
    result = {}
    for binding in bindings:
        result[binding['key']] = load_table_by_date(
            binding['report_code'],
            binding['table_name'],
            selected_date,
            order_by=binding.get('order_by'),
        )
    return result
