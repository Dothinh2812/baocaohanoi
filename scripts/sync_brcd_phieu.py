#!/usr/bin/env python3
"""Cron entry point — sync Vũ trụ tổng BRCD vào brcd_phieu.

Config qua DASHV4_* env vars (như các entrypoint khác). Không qua HTTP.

Crontab (mỗi instance 1 dòng):
  0 * * * * DASHV4_UNIT_CODE=<unit> /home/vtst/dashv4/venv/bin/python3 \
      /home/vtst/dashv4/scripts/sync_brcd_phieu.py \
      >> /home/vtst/dashv4/logs/brcd_phieu_sync.log 2>&1

Exit codes:
  0 — sync thành công, hoặc Excel thiếu (lỡ nhịp refresh, không lỗi)
  1 — lỗi khác (DB lock, Excel corrupt, ...)
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import runtime_limits  # noqa: F401, E402

from blueprints.operations_routes import _sync_brcd_phieu_to_db  # noqa: E402


def main():
    result = _sync_brcd_phieu_to_db()
    print(result)
    if result.get('skipped') and result.get('reason') not in ('excel_missing', ''):
        sys.exit(1)


if __name__ == '__main__':
    main()
