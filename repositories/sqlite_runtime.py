import os
import sqlite3
from contextlib import contextmanager

import pandas as pd

from config import REPORT_HISTORY_DB_PATH


def _readonly_uri(db_path):
    # report_history.db được app chỉ đọc; immutable=1 tránh lỗi journal/WAL
    # khi mở từ dashboard read-only và vẫn cho phép mỗi request mở lại file mới.
    return f'file:{os.path.abspath(db_path)}?mode=ro&immutable=1'


@contextmanager
def get_report_history_connection():
    if not REPORT_HISTORY_DB_PATH or not os.path.exists(REPORT_HISTORY_DB_PATH):
        raise FileNotFoundError('Không tìm thấy report_history.db cho dashv4')

    conn = sqlite3.connect(_readonly_uri(REPORT_HISTORY_DB_PATH), uri=True)
    conn.row_factory = sqlite3.Row
    try:
        yield conn
    finally:
        conn.close()


def read_sql_dataframe(query, params=()):
    with get_report_history_connection() as conn:
        return pd.read_sql_query(query, conn, params=params)


def read_sql_rows(query, params=()):
    with get_report_history_connection() as conn:
        rows = conn.execute(query, params).fetchall()
    return [dict(row) for row in rows]
