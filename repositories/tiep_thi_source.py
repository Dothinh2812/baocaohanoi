import os
import sqlite3
from contextlib import contextmanager
from datetime import date as date_cls

import pandas as pd


DETAIL_COLUMNS = [
    'STT',
    'Ngày',
    'Mã TB',
    'Mã GD',
    'Phòng tiếp thị',
    'Đơn vị',
    'Người tiếp thị',
    'Loại dịch vụ',
]
SUMMARY_COLUMNS = ['STT', 'Đơn vị', 'Người tiếp thị', 'Số mã tiếp thị']
UNIT_SUMMARY_COLUMNS = ['STT', 'Đơn vị', 'Số mã tiếp thị']
SERVICE_SUMMARY_COLUMNS = ['STT', 'Loại dịch vụ', 'Số mã tiếp thị']
UNIT_SERVICE_SUMMARY_COLUMNS = ['STT', 'Đơn vị', 'Loại dịch vụ', 'Số mã tiếp thị']
MARKETER_SERVICE_SUMMARY_COLUMNS = ['STT', 'Đơn vị', 'Người tiếp thị', 'Loại dịch vụ', 'Số mã tiếp thị']
QUARTERLY_SUMMARY_COLUMNS = ['STT', 'Năm', 'Quý', 'Đơn vị', 'Người tiếp thị', 'Số mã tiếp thị']
YEARLY_SUMMARY_COLUMNS = ['STT', 'Năm', 'Đơn vị', 'Người tiếp thị', 'Số mã tiếp thị']


def _readonly_uri(db_path):
    return f'file:{os.path.abspath(db_path)}?mode=ro&immutable=1'


def _validate_iso_date(value):
    if not value:
        return None
    try:
        return date_cls.fromisoformat(str(value)).isoformat()
    except ValueError as exc:
        raise ValueError('Ngày không hợp lệ, cần theo định dạng YYYY-MM-DD') from exc


@contextmanager
def _connect(db_path):
    if not db_path or not os.path.exists(db_path):
        raise FileNotFoundError('Không tìm thấy bao_cao_tiep_thi.db cho dashv4')

    conn = sqlite3.connect(_readonly_uri(db_path), uri=True)
    conn.row_factory = sqlite3.Row
    try:
        yield conn
    finally:
        conn.close()


def get_tiep_thi_available_dates(db_path):
    with _connect(db_path) as conn:
        rows = conn.execute(
            '''
            SELECT report_date
            FROM tiep_thi_daily
            GROUP BY report_date
            ORDER BY report_date DESC
            '''
        ).fetchall()
    return [row['report_date'] for row in rows if row['report_date']]


def resolve_tiep_thi_date_context(db_path, requested_date):
    normalized_requested_date = _validate_iso_date(requested_date)
    available_dates = get_tiep_thi_available_dates(db_path)
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


def load_tiep_thi_detail_df(db_path, selected_date):
    normalized_date = _validate_iso_date(selected_date)
    if not normalized_date:
        return pd.DataFrame(columns=DETAIL_COLUMNS)

    selected_month = normalized_date[:7]
    with _connect(db_path) as conn:
        df = pd.read_sql_query(
            '''
            SELECT
                report_date AS "Ngày",
                ma_tb AS "Mã TB",
                ma_gd AS "Mã GD",
                ten_phong_tthi AS "Phòng tiếp thị",
                ten_to AS "Đơn vị",
                ten_nguoi_tiepthi AS "Người tiếp thị",
                COALESCE(NULLIF(TRIM(loaihinh_tb), ''), '(trống)') AS "Loại dịch vụ"
            FROM tiep_thi_daily
            WHERE substr(report_date, 1, 7) = ?
              AND report_date <= ?
            ORDER BY report_date DESC, ten_to, ten_nguoi_tiepthi, ma_gd, ma_tb
            ''',
            conn,
            params=(selected_month, normalized_date),
        )

    if df.empty:
        return pd.DataFrame(columns=DETAIL_COLUMNS)

    df.insert(0, 'STT', range(1, len(df) + 1))
    return df


def build_tiep_thi_summary_df(detail_df):
    if detail_df.empty:
        return pd.DataFrame(columns=SUMMARY_COLUMNS)

    summary_df = (
        detail_df.groupby(['Đơn vị', 'Người tiếp thị'], dropna=False)
        .size()
        .reset_index(name='Số mã tiếp thị')
        .sort_values(['Đơn vị', 'Người tiếp thị'], kind='stable')
        .reset_index(drop=True)
    )
    summary_df.insert(0, 'STT', range(1, len(summary_df) + 1))

    total_row = pd.DataFrame(
        [
            {
                'STT': '',
                'Đơn vị': 'TỔNG CỘNG',
                'Người tiếp thị': '',
                'Số mã tiếp thị': int(summary_df['Số mã tiếp thị'].sum()),
            }
        ]
    )
    return pd.concat([summary_df, total_row], ignore_index=True)


def _build_grouped_summary_df(detail_df, group_columns, output_columns, total_values):
    if detail_df.empty:
        return pd.DataFrame(columns=output_columns)

    summary_df = (
        detail_df.groupby(group_columns, dropna=False)
        .size()
        .reset_index(name='Số mã tiếp thị')
        .reset_index(drop=True)
    )
    leading_groups = group_columns[:-1]
    sort_columns = leading_groups + ['Số mã tiếp thị', group_columns[-1]]
    ascending = [True] * len(leading_groups) + [False, True]
    summary_df = summary_df.sort_values(
        sort_columns,
        ascending=ascending,
        kind='stable',
    ).reset_index(drop=True)
    summary_df.insert(0, 'STT', range(1, len(summary_df) + 1))

    total_row = {
        'STT': '',
        'Số mã tiếp thị': int(summary_df['Số mã tiếp thị'].sum()),
    }
    total_row.update(total_values)
    return pd.concat([summary_df, pd.DataFrame([total_row])], ignore_index=True)


def build_tiep_thi_service_summary_df(detail_df):
    return _build_grouped_summary_df(
        detail_df,
        ['Loại dịch vụ'],
        SERVICE_SUMMARY_COLUMNS,
        {'Loại dịch vụ': 'TỔNG CỘNG'},
    )


def build_tiep_thi_unit_service_summary_df(detail_df):
    return _build_grouped_summary_df(
        detail_df,
        ['Đơn vị', 'Loại dịch vụ'],
        UNIT_SERVICE_SUMMARY_COLUMNS,
        {'Đơn vị': 'TỔNG CỘNG', 'Loại dịch vụ': ''},
    )


def build_tiep_thi_marketer_service_summary_df(detail_df):
    return _build_grouped_summary_df(
        detail_df,
        ['Đơn vị', 'Người tiếp thị', 'Loại dịch vụ'],
        MARKETER_SERVICE_SUMMARY_COLUMNS,
        {'Đơn vị': 'TỔNG CỘNG', 'Người tiếp thị': '', 'Loại dịch vụ': ''},
    )


def build_tiep_thi_unit_summary_df(detail_df):
    if detail_df.empty:
        return pd.DataFrame(columns=UNIT_SUMMARY_COLUMNS)

    summary_df = (
        detail_df.groupby(['Đơn vị'], dropna=False)
        .size()
        .reset_index(name='Số mã tiếp thị')
        .sort_values(['Đơn vị'], kind='stable')
        .reset_index(drop=True)
    )
    summary_df.insert(0, 'STT', range(1, len(summary_df) + 1))

    total_row = pd.DataFrame(
        [
            {
                'STT': '',
                'Đơn vị': 'TỔNG CỘNG',
                'Số mã tiếp thị': int(summary_df['Số mã tiếp thị'].sum()),
            }
        ]
    )
    return pd.concat([summary_df, total_row], ignore_index=True)


def _selected_quarter(normalized_date):
    month = int(normalized_date[5:7])
    return (month - 1) // 3 + 1


def load_tiep_thi_quarterly_summary_df(db_path, selected_date):
    normalized_date = _validate_iso_date(selected_date)
    if not normalized_date:
        return pd.DataFrame(columns=QUARTERLY_SUMMARY_COLUMNS)

    selected_year = normalized_date[:4]
    selected_quarter = _selected_quarter(normalized_date)
    with _connect(db_path) as conn:
        df = pd.read_sql_query(
            '''
            SELECT
                year AS "Năm",
                quarter AS "Quý",
                ten_to AS "Đơn vị",
                ten_nguoi_tiepthi AS "Người tiếp thị",
                so_ma_tiep_thi AS "Số mã tiếp thị"
            FROM v_tiep_thi_quarterly_summary
            WHERE year = ?
              AND quarter = ?
            ORDER BY ten_to, ten_nguoi_tiepthi
            ''',
            conn,
            params=(selected_year, selected_quarter),
        )

    if df.empty:
        return pd.DataFrame(columns=QUARTERLY_SUMMARY_COLUMNS)

    df.insert(0, 'STT', range(1, len(df) + 1))
    total_row = pd.DataFrame(
        [
            {
                'STT': '',
                'Năm': selected_year,
                'Quý': selected_quarter,
                'Đơn vị': 'TỔNG CỘNG',
                'Người tiếp thị': '',
                'Số mã tiếp thị': int(df['Số mã tiếp thị'].sum()),
            }
        ]
    )
    return pd.concat([df, total_row], ignore_index=True)


def load_tiep_thi_yearly_summary_df(db_path, selected_date):
    normalized_date = _validate_iso_date(selected_date)
    if not normalized_date:
        return pd.DataFrame(columns=YEARLY_SUMMARY_COLUMNS)

    selected_year = normalized_date[:4]
    with _connect(db_path) as conn:
        df = pd.read_sql_query(
            '''
            SELECT
                year AS "Năm",
                ten_to AS "Đơn vị",
                ten_nguoi_tiepthi AS "Người tiếp thị",
                so_ma_tiep_thi AS "Số mã tiếp thị"
            FROM v_tiep_thi_yearly_summary
            WHERE year = ?
            ORDER BY ten_to, ten_nguoi_tiepthi
            ''',
            conn,
            params=(selected_year,),
        )

    if df.empty:
        return pd.DataFrame(columns=YEARLY_SUMMARY_COLUMNS)

    df.insert(0, 'STT', range(1, len(df) + 1))
    total_row = pd.DataFrame(
        [
            {
                'STT': '',
                'Năm': selected_year,
                'Đơn vị': 'TỔNG CỘNG',
                'Người tiếp thị': '',
                'Số mã tiếp thị': int(df['Số mã tiếp thị'].sum()),
            }
        ]
    )
    return pd.concat([df, total_row], ignore_index=True)
