#!/usr/bin/env python3
"""Dọn dẹp Flask session files hết hạn.

Chạy định kỳ qua cron:
    */30 * * * * python3 /home/vtst/dashv4/scripts/cleanup_sessions.py

Hoặc chạy thủ công:
    python3 /home/vtst/dashv4/scripts/cleanup_sessions.py
"""
import os
import sys
import time
from datetime import datetime, timedelta

SESSION_LIFETIME_HOURS = 12
SESSION_FILE_PREFIX = 'session_'


def cleanup_sessions(session_dir):
    if not os.path.isdir(session_dir):
        print(f'Thu muc session khong ton tai: {session_dir}')
        return 0

    cutoff = time.time() - SESSION_LIFETIME_HOURS * 3600
    removed = 0
    kept = 0

    for filename in os.listdir(session_dir):
        if not filename.startswith(SESSION_FILE_PREFIX):
            continue

        filepath = os.path.join(session_dir, filename)
        try:
            if os.path.getmtime(filepath) < cutoff:
                os.remove(filepath)
                removed += 1
            else:
                kept += 1
        except OSError:
            kept += 1

    print(f'[{datetime.now().strftime("%Y-%m-%d %H:%M:%S")}] '
          f'Session cleanup: removed={removed}, kept={kept}')
    return removed


def main():
    sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))
    from config import INSTANCE_RUNTIME_DIR, SESSION_FILE_DIR

    session_dir = SESSION_FILE_DIR
    cutoff_dt = datetime.now() - timedelta(hours=SESSION_LIFETIME_HOURS)
    print(f'Cleaning session files older than {SESSION_LIFETIME_HOURS}h '
          f'(before {cutoff_dt.strftime("%Y-%m-%d %H:%M:%S")})')
    print(f'Session directory: {session_dir}')

    removed = cleanup_sessions(session_dir)

    if removed > 0:
        print(f'Da xoa {removed} session file het han.')
    else:
        print('Khong co session file nao het han.')

    return removed


if __name__ == '__main__':
    main()