import os
import time
import logging

import runtime_limits  # noqa: F401  # must run before pandas/numpy imports
from flask import Flask, flash, g, jsonify, redirect, render_template, request, session, url_for
from flask_session import Session

from app_helpers import configure_app
from auth import get_user_by_username
from blueprints import (
    auth_bp,
    growth_bp,
    inventory_bp,
    operations_bp,
    quality_bp,
    quangchudong_bp,
    retention_bp,
    sa_outage_bp,
    statistics_bp,
)
from config import DISABLED_NONPAGE_ENDPOINTS, DISABLED_PAGE_ENDPOINTS, PUBLIC_ENDPOINTS
from services import initialize_background_services

app = Flask(__name__)
configure_app(app)
if app.config.get('LOG_REQUESTS'):
    app.logger.setLevel(logging.INFO)

# Khởi tạo Session
Session(app)

app.register_blueprint(auth_bp)
app.register_blueprint(growth_bp)
app.register_blueprint(quangchudong_bp)
app.register_blueprint(inventory_bp)
app.register_blueprint(operations_bp)
app.register_blueprint(quality_bp)
app.register_blueprint(retention_bp)
app.register_blueprint(sa_outage_bp)
app.register_blueprint(statistics_bp)


@app.before_request
def log_request_start():
    if not app.config.get('LOG_REQUESTS'):
        return None
    g.request_started_at = time.perf_counter()
    app.logger.info('request start pid=%s method=%s path=%s', os.getpid(), request.method, request.path)
    return None


@app.after_request
def log_request_end(response):
    if app.config.get('LOG_REQUESTS') and hasattr(g, 'request_started_at'):
        duration_ms = int((time.perf_counter() - g.request_started_at) * 1000)
        app.logger.info(
            'request end pid=%s status=%s duration_ms=%s method=%s path=%s',
            os.getpid(),
            response.status_code,
            duration_ms,
            request.method,
            request.path,
        )
    return response


@app.before_request
def enforce_auth_policy():
    if request.endpoint in PUBLIC_ENDPOINTS or request.endpoint is None:
        return None

    if 'username' in session:
        if request.endpoint in DISABLED_PAGE_ENDPOINTS:
            feature = DISABLED_PAGE_ENDPOINTS[request.endpoint]
            return render_template(
                'pages/pending_feature.html',
                current_user=get_user_by_username(session.get('username')),
                active_page=feature['active_page'],
                feature_title=feature['title'],
                feature_reason=feature['reason'],
            ), 501
        if request.endpoint in DISABLED_NONPAGE_ENDPOINTS:
            feature = DISABLED_NONPAGE_ENDPOINTS[request.endpoint]
            payload = {
                'error': 'legacy_endpoint_disabled',
                'title': feature['title'],
                'reason': feature['reason'],
                'required_display_contract': feature.get('required_display_contract', {}),
            }
            if request.path.startswith('/api/'):
                return jsonify(payload), 501
            return jsonify(payload), 501
        return None

    if request.path.startswith('/api/'):
        return jsonify({'error': 'Vui lòng đăng nhập để tiếp tục'}), 401

    flash('Vui lòng đăng nhập để tiếp tục', 'warning')
    return redirect(url_for('auth.login'))


@app.route('/')
def index():
    return redirect(url_for('operations.page_tong_hop_bsc_kpi'))


def initialize_runtime_services(*, warm=False):
    if app.config.get('ENABLE_BACKGROUND_SERVICES'):
        initialize_background_services(app, warm=warm)


def run_development_server():
    app.run(
        debug=app.config['DEBUG'],
        port=app.config['SERVER_PORT'],
        host=app.config['SERVER_HOST'],
        threaded=app.config['DEV_SERVER_THREADED'],
        use_reloader=False,
    )


if __name__ == '__main__':
    if not app.config['DEBUG'] or os.environ.get('WERKZEUG_RUN_MAIN') == 'true':
        initialize_runtime_services(warm=True)
    # Khi chạy trực tiếp bằng python (development only)
    # Production nên dùng: gunicorn -c gunicorn_config.py dashboard:app
    run_development_server()
