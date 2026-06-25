#!/usr/bin/env python3
"""Cron entry point — sync tiến độ SHC CTS trong ngày vào shc_cts_tien_do.

Cron (mỗi giờ + chốt 20:00):
  0 * * * * DASHV4_UNIT_CODE=son_tay python3 scripts/sync_shc_cts_tien_do.py
  0 20 * * * DASHV4_UNIT_CODE=son_tay python3 scripts/sync_shc_cts_tien_do.py
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import runtime_limits  # noqa: F401, E402
from blueprints.quality_routes import _sync_shc_cts_tien_do_to_db  # noqa: E402


def main():
    result = _sync_shc_cts_tien_do_to_db()
    print(result)
    if result.get('skipped') and result.get('reason') not in (
            'intraday_missing', 'excel_unreadable', ''):
        sys.exit(1)


if __name__ == '__main__':
    main()
