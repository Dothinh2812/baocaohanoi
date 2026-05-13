#!/bin/bash
# Script khởi động dashv4 từ thư mục /home/vtst/dashv4

cd /home/vtst/dashv4

echo "==================================="
echo "  Starting Dashv4 Application"
echo "==================================="
echo ""
echo "Dashboard directory: $(pwd)"
echo "DB path: ${DASHV4_DB_PATH:-/home/vtst/bchn/runtime/son_tay/sqlite_history/report_history.db}"
echo "Timezone: ${DASHV4_TIMEZONE:-Asia/Ho_Chi_Minh}"
echo ""
echo "Accessible at: http://0.0.0.0:${DASHV4_PORT:-5011}"
echo ""
echo "Press Ctrl+C to stop"
echo "==================================="
echo ""

export DASHV4_TIMEZONE="${DASHV4_TIMEZONE:-Asia/Ho_Chi_Minh}"
export TZ="${DASHV4_TIMEZONE}"
export DASHV4_NATIVE_THREADS="${DASHV4_NATIVE_THREADS:-1}"
export OMP_NUM_THREADS="${OMP_NUM_THREADS:-${DASHV4_NATIVE_THREADS}}"
export OPENBLAS_NUM_THREADS="${OPENBLAS_NUM_THREADS:-${DASHV4_NATIVE_THREADS}}"
export MKL_NUM_THREADS="${MKL_NUM_THREADS:-${DASHV4_NATIVE_THREADS}}"
export NUMEXPR_NUM_THREADS="${NUMEXPR_NUM_THREADS:-${DASHV4_NATIVE_THREADS}}"
export VECLIB_MAXIMUM_THREADS="${VECLIB_MAXIMUM_THREADS:-${DASHV4_NATIVE_THREADS}}"
export BLIS_NUM_THREADS="${BLIS_NUM_THREADS:-${DASHV4_NATIVE_THREADS}}"
export DASHV4_PORT="${DASHV4_PORT:-5011}"
export DASHV4_DB_PATH="${DASHV4_DB_PATH:-/home/vtst/bchn/runtime/son_tay/sqlite_history/report_history.db}"
export DASHV4_WORKERS="${DASHV4_WORKERS:-2}"
export DASHV4_GUNICORN_TIMEOUT="${DASHV4_GUNICORN_TIMEOUT:-30}"
export DASHV4_LOG_REQUESTS="${DASHV4_LOG_REQUESTS:-1}"

GUNICORN_BIN="${DASHV4_GUNICORN_BIN:-$(command -v gunicorn || true)}"
if [ -z "$GUNICORN_BIN" ] && [ -x /home/vtst/.local/bin/gunicorn ]; then
    GUNICORN_BIN=/home/vtst/.local/bin/gunicorn
fi
if [ -z "$GUNICORN_BIN" ]; then
    echo "Không tìm thấy gunicorn. Cài gunicorn hoặc đặt DASHV4_GUNICORN_BIN." >&2
    exit 1
fi

exec "$GUNICORN_BIN" -c gunicorn_config.py dashboard:app
