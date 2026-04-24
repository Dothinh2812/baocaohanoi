#!/bin/bash
# Script khởi động dashv4 từ thư mục /home/vtst/dashv4

cd /home/vtst/dashv4

echo "==================================="
echo "  Starting Dashv4 Application"
echo "==================================="
echo ""
echo "Dashboard directory: $(pwd)"
echo "DB path: ${DASHV4_DB_PATH:-/home/vtst/baocaohanoi/api_transition/runtime/son_tay/sqlite_history/report_history.db}"
echo "Timezone: ${DASHV4_TIMEZONE:-Asia/Ho_Chi_Minh}"
echo ""
echo "Accessible at: http://0.0.0.0:${DASHV4_PORT:-5010}"
echo ""
echo "Press Ctrl+C to stop"
echo "==================================="
echo ""

DASHV4_TIMEZONE="${DASHV4_TIMEZONE:-Asia/Ho_Chi_Minh}" \
TZ="${DASHV4_TIMEZONE:-Asia/Ho_Chi_Minh}" \
DASHV4_PORT="${DASHV4_PORT:-5010}" \
DASHV4_DB_PATH="${DASHV4_DB_PATH:-/home/vtst/baocaohanoi/api_transition/runtime/son_tay/sqlite_history/report_history.db}" \
python3 dashboard.py
