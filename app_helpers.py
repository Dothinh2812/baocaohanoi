import glob
import hmac
import os
import secrets
import time
from collections import OrderedDict
from datetime import datetime
from functools import wraps
from threading import Lock

import pandas as pd
from flask import current_app, flash, jsonify, redirect, request, send_file, session, url_for
from werkzeug.utils import safe_join

import config
from route_policy import is_endpoint_enabled


_EXCEL_CACHE_MAX_ENTRIES = 50
_EXCEL_CACHE = OrderedDict()
_EXCEL_CACHE_LOCK = Lock()
_HIDDEN_DASHBOARD_COLUMNS = {
    'id',
    '__snapshot_id',
    '__sheet_id',
    '__row_num',
    '__row_hash',
    '__imported_at',
}


def configure_app(app):
    app.config.from_object(config.DashboardConfig)
    app.template_folder = 'templates'
    os.makedirs(app.config['SESSION_FILE_DIR'], exist_ok=True)
    register_template_globals(app)


def register_template_globals(app):
    @app.context_processor
    def inject_template_helpers():
        return {
            'csrf_token': get_csrf_token,
            'is_endpoint_enabled': is_endpoint_enabled,
        }


def get_csrf_token():
    token = session.get('_csrf_token')
    if not token:
        token = secrets.token_urlsafe(32)
        session['_csrf_token'] = token
    return token


def validate_csrf_token(token):
    session_token = session.get('_csrf_token')
    return bool(session_token and token and hmac.compare_digest(session_token, token))


def csrf_protect(view_func):
    @wraps(view_func)
    def wrapped(*args, **kwargs):
        if request.method in {'POST', 'PUT', 'PATCH', 'DELETE'}:
            token = request.form.get('csrf_token') or request.headers.get('X-CSRF-Token')
            if not validate_csrf_token(token):
                if request.path.startswith('/api/'):
                    return jsonify({'error': 'CSRF token không hợp lệ'}), 400
                flash('Phiên làm việc không hợp lệ. Vui lòng thử lại.', 'danger')
                return redirect(request.referrer or url_for('auth.login'))
        return view_func(*args, **kwargs)

    return wrapped


def json_error(message, status=400):
    return jsonify({'error': message}), status


def add_no_cache_headers(response):
    response.headers['Cache-Control'] = 'no-cache, no-store, must-revalidate'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '0'
    return response


def safe_file_response(file_path, *, as_attachment=False, download_name=None, mimetype=None, no_cache=False):
    if not os.path.isfile(file_path):
        return json_error('File không tồn tại', 404)

    response = send_file(
        file_path,
        as_attachment=as_attachment,
        download_name=download_name,
        mimetype=mimetype,
    )
    if no_cache:
        add_no_cache_headers(response)
    return response


def safe_directory_response(directory, relative_path, *, no_cache=False):
    safe_path = safe_join(directory, relative_path)
    if safe_path is None or not os.path.isfile(safe_path):
        return json_error('Đường dẫn file không hợp lệ', 404)

    response = send_file(safe_path)
    if no_cache:
        add_no_cache_headers(response)
    return response


def get_file_metadata(file_path):
    stat = os.stat(file_path)
    return {
        'modified': datetime.fromtimestamp(stat.st_mtime).strftime('%d/%m/%Y %H:%M:%S'),
        'size': stat.st_size,
        'name': os.path.basename(file_path),
    }


def build_file_info(file_path, *, include_name=False):
    metadata = get_file_metadata(file_path)
    file_info = {
        'modified': metadata['modified'],
        'size': metadata['size'],
    }
    if include_name:
        file_info['name'] = metadata['name']
    return file_info


def latest_matching_file(pattern):
    files = [path for path in glob.glob(pattern) if not os.path.basename(path).startswith('~$')]
    if not files:
        return None
    return max(files, key=os.path.getmtime)


def _freeze_cache_value(value):
    if isinstance(value, dict):
        return tuple((key, _freeze_cache_value(val)) for key, val in sorted(value.items()))
    if isinstance(value, (list, tuple, set)):
        return tuple(_freeze_cache_value(item) for item in value)
    return value


def _cache_key(file_path, sheet_name, kwargs):
    stat = os.stat(file_path)
    return (
        os.path.abspath(file_path),
        stat.st_mtime_ns,
        stat.st_size,
        sheet_name,
        _freeze_cache_value(kwargs),
    )


def _evict_excel_cache():
    stale_keys = [k for k in _EXCEL_CACHE if k[0] not in _active_file_paths]
    for key in stale_keys:
        _EXCEL_CACHE.pop(key, None)
    while len(_EXCEL_CACHE) > _EXCEL_CACHE_MAX_ENTRIES:
        _EXCEL_CACHE.popitem(last=False)


_active_file_paths = set()


def read_excel_sheet_cached(file_path, sheet_name, **kwargs):
    cache_key = _cache_key(file_path, sheet_name, kwargs)
    abs_path = cache_key[0]

    with _EXCEL_CACHE_LOCK:
        for stale_key in list(_EXCEL_CACHE):
            if stale_key[0] == abs_path and stale_key[1:3] != cache_key[1:3]:
                _EXCEL_CACHE.pop(stale_key, None)

        cached_df = _EXCEL_CACHE.get(cache_key)
        if cached_df is not None:
            _EXCEL_CACHE.move_to_end(cache_key)
            return cached_df.copy()

        _active_file_paths.add(abs_path)
        cached_df = pd.read_excel(file_path, sheet_name=sheet_name, **kwargs)
        _EXCEL_CACHE[cache_key] = cached_df
        _EXCEL_CACHE.move_to_end(cache_key)

        while len(_EXCEL_CACHE) > _EXCEL_CACHE_MAX_ENTRIES:
            evicted_key, _ = _EXCEL_CACHE.popitem(last=False)
            _active_file_paths.discard(evicted_key[0])

    return cached_df.copy()


def serialize_dataframe(df, datetime_format='%Y-%m-%d %H:%M:%S'):
    serialized = df.copy()
    hidden_cols = [col for col in serialized.columns if str(col) in _HIDDEN_DASHBOARD_COLUMNS]
    if hidden_cols:
        serialized = serialized.drop(columns=hidden_cols)
    for col in serialized.columns:
        if pd.api.types.is_datetime64_any_dtype(serialized[col]):
            serialized[col] = serialized[col].dt.strftime(datetime_format)
        elif pd.api.types.is_period_dtype(serialized[col]):
            serialized[col] = serialized[col].astype(str)
    return serialized.fillna('')


def build_sheet_payload(df, *, datetime_format='%Y-%m-%d %H:%M:%S'):
    serialized = serialize_dataframe(df, datetime_format=datetime_format)
    return {
        'columns': serialized.columns.tolist(),
        'data': serialized.to_dict('records'),
    }


def build_multi_sheet_payload(
    file_path,
    *,
    sheet_names=None,
    include_file_name=False,
    datetime_format='%Y-%m-%d %H:%M:%S',
    include_sheet_errors=False,
    transform=None,
):
    payload = {
        'file_info': build_file_info(file_path, include_name=include_file_name),
        'sheets': {},
    }

    if sheet_names is None:
        sheet_names = pd.ExcelFile(file_path).sheet_names

    for sheet_name in sheet_names:
        try:
            df = read_excel_sheet_cached(file_path, sheet_name)
            result = transform(df.copy(), sheet_name) if transform else df
            if result is None:
                result = df
            if isinstance(result, pd.DataFrame):
                result = build_sheet_payload(result, datetime_format=datetime_format)
            payload['sheets'][sheet_name] = result
        except Exception as exc:
            if not include_sheet_errors:
                raise
            payload['sheets'][sheet_name] = {'error': f'Không thể đọc sheet: {exc}'}

    return payload


def log_background_exception(message, exc):
    current_app.logger.warning('%s: %s', message, exc)
