#!/usr/bin/env python3
"""Cron entry point — sync Vũ trụ tổng PTTB vào pttb_phieu.
Cron: 0 * * * * DASHV4_UNIT_CODE=son_tay python3 scripts/sync_pttb_phieu.py
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import runtime_limits  # noqa: F401, E402
from blueprints.operations_routes import _sync_pttb_phieu_to_db  # noqa: E402


def main():
    result = _sync_pttb_phieu_to_db()
    print(result)
    if result.get('skipped') and result.get('reason') not in ('excel_missing', ''):
        sys.exit(1)


if __name__ == '__main__':
    main()
