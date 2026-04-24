# Gunicorn configuration file for dashv4 application
# Sử dụng: gunicorn -c gunicorn_config.py dashboard:app

import multiprocessing
import os

from config import DashboardConfig

# Server socket
bind = f"{DashboardConfig.SERVER_HOST}:{DashboardConfig.SERVER_PORT}"
backlog = 2048

# Worker processes
workers = multiprocessing.cpu_count() * 2 + 1  # Công thức khuyến nghị: (2 x $num_cores) + 1
worker_class = 'sync'  # Dùng sync vì Flask app không async
worker_connections = 1000
max_requests = 1000  # Restart worker sau 1000 requests để tránh memory leak
max_requests_jitter = 50
timeout = 120  # Timeout 120 giây cho các request chậm (đọc Excel lớn)
keepalive = 5

# Process naming
proc_name = 'dashv4_vnpt'

# Server mechanics
daemon = False  # Không chạy background, để supervisor/systemd quản lý
pidfile = '/tmp/dashv4.pid'
umask = 0
user = None
group = None
tmp_upload_dir = None

# Logging
accesslog = 'logs/gunicorn_access.log'
errorlog = 'logs/gunicorn_error.log'
loglevel = 'info'
access_log_format = '%(h)s %(l)s %(u)s %(t)s "%(r)s" %(s)s %(b)s "%(f)s" "%(a)s" %(D)s'

# Tạo thư mục logs nếu chưa có
os.makedirs('logs', exist_ok=True)

# SSL (nếu cần)
# keyfile = None
# certfile = None

# Server hooks
def on_starting(server):
    """
    Hook được gọi khi server bắt đầu khởi động
    """
    print("=" * 60)
    print("Dashv4 - Starting...")
    print(f"Workers: {workers}")
    print(f"Bind: {bind}")
    print("=" * 60)


def on_reload(server):
    """
    Hook được gọi khi server reload
    """
    print("Server reloading...")


def worker_int(worker):
    """
    Hook được gọi khi worker nhận SIGINT hoặc SIGQUIT
    """
    print(f"Worker {worker.pid} received interrupt signal")


def post_fork(server, worker):
    """
    Hook được gọi sau khi worker được fork
    """
    server.log.info(f"Worker spawned (pid: {worker.pid})")
    from dashboard import initialize_runtime_services
    initialize_runtime_services(warm=True)


def pre_fork(server, worker):
    """
    Hook được gọi trước khi worker được fork
    """
    pass


def pre_exec(server):
    """
    Hook được gọi trước khi server exec()
    """
    server.log.info("Forked child, re-executing.")


def when_ready(server):
    """
    Hook được gọi khi server đã sẵn sàng nhận requests
    """
    server.log.info("Server is ready. Spawning workers")


def worker_abort(worker):
    """
    Hook được gọi khi worker bị abort
    """
    server.log.info(f"Worker {worker.pid} aborted")


# Preload app để chia sẻ memory giữa workers
preload_app = True

# Graceful timeout
graceful_timeout = 30

# Environment variables
raw_env = [
    'FLASK_ENV=production',
    f'TZ={DashboardConfig.TIMEZONE}',
    f'DASHV4_TIMEZONE={DashboardConfig.TIMEZONE}',
]
