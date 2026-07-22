# Checklist hằng ngày cho tổ trưởng tổ KTĐB — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Thêm trang `/checklist` mới cho tổ trưởng tổ KTĐB cập nhật trạng thái + ghi chú + ảnh minh chứng theo mẫu 3 giai đoạn × 15 mục cố định hằng ngày, per-đội, với auto-source metric từ `brcd_kiemsoat.db`.

**Architecture:** New blueprint `checklist_bp` trong `blueprints/checklist_routes.py` (Approach A). DB per-instance `<INSTANCE_RUNTIME_DIR>/checklist.db` với 3 bảng (`checklist_state`, `checklist_photos`, `checklist_audit_log`). Photo upload hạ tầng mới (multipart + CSRF + secure_filename + UUID filename + partition tháng). Mở rộng `auth.py` thêm role `to_truong` + cột `doi` lazy migration. 5 metric fetcher đọc read-only từ DBs có sẵn.

**Tech Stack:** Python 3.10, Flask 3.1, pandas 2.2, sqlite3 stdlib, openpyxl, werkzeug `secure_filename`, HTML/JS thuần.

**Spec:** `docs/superpowers/specs/2026-07-22-checklist-to-truong-design.md`

## Global Constraints

- Chạy mọi lệnh từ repo root `/home/vtst/dashv4`. Cây module phẳng, import dạng top-level.
- KHÔNG có lint/typecheck/formatter. Xác minh cú pháp bằng `python3 -m py_compile <file>`.
- Test dùng pytest HỆ THỐNG: `python3 -m pytest tests/`. KHÔNG dùng `venv/bin/pytest`.
- KHÔNG thêm comment vào code trừ khi task yêu cầu.
- UI/comment dùng tiếng Việt.
- Endpoint prefix `checklist.` (vd `checklist.page_checklist`, `checklist.api_checklist_data`).
- POST JSON endpoint **không dùng `@csrf_protect`** (convention kiemsoat). POST multipart (upload-photo) **bắt buộc** `@csrf_protect`.
- Doc-sync rule (AGENTS.md): task cuối phải cập nhật `docs/00`, `docs/04`, `docs/08`, `AGENTS.md`, tạo mới `docs/12-checklist-to-truong.md`.

## File Structure

**Tạo mới:**
- `blueprints/checklist_routes.py` — toàn bộ logic checklist.
- `templates/pages/checklist.html`, `static/js/pages/checklist.js`, `static/css/checklist.css`.
- `tests/test_auth_to_truong.py`, `tests/test_checklist_state.py`, `tests/test_checklist_photos.py`, `tests/test_checklist_metric_fetchers.py`, `tests/test_checklist_routes.py`.
- `docs/12-checklist-to-truong.md`.

**Sửa:**
- `config.py` — thêm 4 biến checklist.
- `auth.py` — role `to_truong` + cột `doi` lazy + decorator `to_truong_or_admin_required`.
- `auth_routes.py` — set session `doi` + endpoint `/admin/set-user-doi`.
- `templates/admin_users.html` — thêm cột Đội + dropdown Role 3 giá trị.
- `blueprints/__init__.py`, `dashboard.py`, `route_policy.py`, `templates/base.html` — wiring.
- `static/js/api.js` — thêm 5 methods.
- `templates/base.html` — CSRF meta tag cho JS fetch.
- `docs/00-doc-index.md`, `docs/04-mapping-route-va-du-lieu.md`, `docs/08-trang-thai-thuc-thi.md`, `AGENTS.md` — doc-sync.

---

## Task 1: Cấu hình checklist trong config.py

**Files:**
- Modify: `config.py` (ngay sau khối `BRCD_KIEMSOAT_DB_PATH` dòng ~117-124)

**Interfaces:**
- Produces: `CHECKLIST_DB_PATH`, `CHECKLIST_UPLOAD_DIR`, `CHECKLIST_PHOTO_MAX_BYTES`, `CHECKLIST_PHOTO_MAX_PER_ITEM`.

- [ ] **Step 1: Thêm biến config**

Mở `config.py`, tìm khối `BRCD_KIEMSOAT_DB_PATH = _first_existing_path(...)`. Chèn ngay sau đó:

```python
CHECKLIST_DB_PATH = _first_existing_path(
    os.getenv('DASHV4_CHECKLIST_DB_PATH'),
    os.path.join(INSTANCE_RUNTIME_DIR, 'checklist.db'),
)
CHECKLIST_UPLOAD_DIR = os.getenv(
    'DASHV4_CHECKLIST_UPLOAD_DIR',
    os.path.join(INSTANCE_RUNTIME_DIR, 'checklist_uploads'),
)
CHECKLIST_PHOTO_MAX_BYTES = int(os.getenv('DASHV4_CHECKLIST_PHOTO_MAX_BYTES', str(5 * 1024 * 1024)))
CHECKLIST_PHOTO_MAX_PER_ITEM = int(os.getenv('DASHV4_CHECKLIST_PHOTO_MAX_PER_ITEM', '5'))
```

- [ ] **Step 2: Xác minh import**

Run: `python3 -c "from config import CHECKLIST_DB_PATH, CHECKLIST_UPLOAD_DIR, CHECKLIST_PHOTO_MAX_BYTES, CHECKLIST_PHOTO_MAX_PER_ITEM; print(CHECKLIST_DB_PATH)"`
Expected: in đường dẫn kết thúc `checklist.db`.

- [ ] **Step 3: Commit**

```bash
git add config.py
git commit -m "config: thêm biến CHECKLIST_DB_PATH và upload config"
```

---

## Task 2: Mở rộng auth.py với role to_truong + cột doi (lazy migration)

**Files:**
- Modify: `auth.py` (helpers + `get_all_users` + `update_user_role` + thêm `update_user_doi`).
- Test: `tests/test_auth_to_truong.py` (create new).

**Interfaces:**
- Produces: `_safe_get_doi(row)`, `update_user_doi(username, new_doi)`.
- Modifies: `update_user_role()` chấp nhận `'to_truong'`; `get_all_users()` đảm bảo có cột `doi`.

- [ ] **Step 1: Viết test fail**

Tạo `tests/test_auth_to_truong.py`:

```python
import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from auth import _safe_get_doi, get_all_users, update_user_role


def _write_users_xlsx(tmp_path, rows):
    path = tmp_path / 'username.xlsx'
    pd.DataFrame(rows).to_excel(path, index=False)
    return str(path)


def test_safe_get_doi_returns_value_when_present():
    assert _safe_get_doi({'username': 'a', 'doi': 'Đội 1'}) == 'Đội 1'


def test_safe_get_doi_returns_none_when_missing():
    assert _safe_get_doi({'username': 'a'}) is None


def test_get_all_users_includes_doi_key(tmp_path, monkeypatch):
    path = _write_users_xlsx(tmp_path, [
        {'username': 'u1', 'name': 'User 1', 'password': 'x', 'role': 'to_truong',
         'is_first_login': False, 'is_active': True, 'doi': 'Đội 1'},
        {'username': 'u2', 'name': 'User 2', 'password': 'x', 'role': 'user',
         'is_first_login': False, 'is_active': True},
    ])
    monkeypatch.setattr('auth.EXCEL_FILE', path)
    from auth import invalidate_users_cache
    invalidate_users_cache()

    df = get_all_users()
    rows = df.to_dict('records')

    assert rows[0]['doi'] == 'Đội 1'
    assert pd.isna(rows[1]['doi'])


def test_update_user_role_accepts_to_truong(tmp_path, monkeypatch):
    path = _write_users_xlsx(tmp_path, [
        {'username': 'u1', 'name': 'User 1', 'password': 'x', 'role': 'user',
         'is_first_login': False, 'is_active': True},
    ])
    monkeypatch.setattr('auth.EXCEL_FILE', path)
    from auth import invalidate_users_cache
    invalidate_users_cache()

    ok, msg = update_user_role('u1', 'to_truong')
    assert ok is True
    df = pd.read_excel(path)
    assert df.iloc[0]['role'] == 'to_truong'


def test_update_user_role_rejects_invalid_role(tmp_path, monkeypatch):
    path = _write_users_xlsx(tmp_path, [
        {'username': 'u1', 'name': 'User 1', 'password': 'x', 'role': 'user',
         'is_first_login': False, 'is_active': True},
    ])
    monkeypatch.setattr('auth.EXCEL_FILE', path)
    from auth import invalidate_users_cache
    invalidate_users_cache()

    ok, msg = update_user_role('u1', 'superadmin')
    assert ok is False
    assert 'không hợp lệ' in msg.lower()
```

- [ ] **Step 2: Run test fail**

Run: `python3 -m pytest tests/test_auth_to_truong.py -v`
Expected: FAIL `ImportError: cannot import name '_safe_get_doi'`.

- [ ] **Step 3: Implement**

Trong `auth.py`:

3a. Thêm helper `_safe_get_doi` ngay sau imports (trước `EXCEL_FILE`):

```python
def _safe_get_doi(row):
    """Trả giá trị cột 'doi' nếu có, None nếu thiếu/NaN. Lazy migration users.xlsx cũ."""
    if not isinstance(row, dict):
        try:
            row = dict(row)
        except Exception:
            return None
    val = row.get('doi')
    if val is None:
        return None
    try:
        if pd.isna(val):
            return None
    except (TypeError, ValueError):
        pass
    return val
```

3b. Sửa `update_user_role()` (~line 171-181) — đổi enum check:

```python
def update_user_role(username, new_role):
    """Cập nhật role của user (admin/user/to_truong)"""
    if new_role not in ['admin', 'user', 'to_truong']:
        return False, "Role không hợp lệ"

    updates = {'role': new_role}

    if update_user(username, updates):
        return True, f"Đã đổi role của {username} thành {new_role}"
    else:
        return False, "Lỗi khi cập nhật role"
```

3c. Thêm `update_user_doi` ngay sau `update_user_role`:

```python
def update_user_doi(username, new_doi):
    """Cập nhật đội của user (cho role to_truong). new_doi có thể là None/rỗng."""
    if new_doi is not None and not isinstance(new_doi, str):
        return False, "Đội phải là chuỗi hoặc rỗng"

    updates = {'doi': new_doi if new_doi else None}

    if update_user(username, updates):
        return True, f"Đã cập nhật đội của {username} thành {new_doi or '—'}"
    else:
        return False, "Lỗi khi cập nhật đội"
```

3d. Sửa `get_all_users()` (~line 54-58) — đảm bảo có cột `doi`:

Tìm:
```python
    with _users_cache_lock:
        _users_cache = df
        _users_cache_mtime = current_mtime

    return df
```

Sửa thành:
```python
    if 'doi' not in df.columns:
        df = df.assign(doi=None)

    with _users_cache_lock:
        _users_cache = df
        _users_cache_mtime = current_mtime

    return df
```

- [ ] **Step 4: Run test pass**

Run: `python3 -m pytest tests/test_auth_to_truong.py -v`
Expected: 5 PASS.

- [ ] **Step 5: Commit**

```bash
git add auth.py tests/test_auth_to_truong.py
git commit -m "auth: thêm role to_truong + cột doi (lazy migration)"
```

---

## Task 3: Decorator to_truong_or_admin_required

**Files:**
- Modify: `auth.py` (thêm decorator sau `admin_required`).
- Test: `tests/test_auth_to_truong.py` (append).

**Interfaces:**
- Produces: `to_truong_or_admin_required(f)`.

- [ ] **Step 1: Viết test fail**

Append vào `tests/test_auth_to_truong.py`:

```python
from unittest.mock import patch
from flask import Flask


def _make_test_app():
    app = Flask(__name__)
    app.config['SECRET_KEY'] = 'test'
    app.config['TESTING'] = True
    return app


def _run_decorator(role, doi):
    from auth import to_truong_or_admin_required
    from flask import session

    app = _make_test_app()

    @app.route('/api/test')
    @to_truong_or_admin_required
    def handler():
        return {'ok': 1}

    with app.test_request_context('/api/test'):
        with patch.dict(session, {'username': 'u1', 'role': role, 'doi': doi}):
            try:
                return handler()
            except Exception as e:
                return ('exc', e)


def test_decorator_blocks_anonymous_user():
    from auth import to_truong_or_admin_required
    app = _make_test_app()

    @app.route('/api/test')
    @to_truong_or_admin_required
    def handler():
        return {'ok': 1}

    with app.test_request_context('/api/test'):
        result = handler()
        assert result[1] == 401


def test_decorator_blocks_role_user():
    result = _run_decorator('user', None)
    assert result[1] == 403


def test_decorator_allows_to_truong_with_doi():
    result = _run_decorator('to_truong', 'Đội 1')
    assert result == {'ok': 1}


def test_decorator_allows_admin():
    result = _run_decorator('admin', None)
    assert result == {'ok': 1}


def test_decorator_blocks_to_truong_without_doi():
    result = _run_decorator('to_truong', None)
    assert result[1] == 403
```

- [ ] **Step 2: Run test fail**

Run: `python3 -m pytest tests/test_auth_to_truong.py -k decorator -v`
Expected: FAIL `ImportError`.

- [ ] **Step 3: Implement**

Trong `auth.py`, ngay sau `admin_required` (~line 226), thêm:

```python
def to_truong_or_admin_required(f):
    """Decorator yêu cầu role admin hoặc to_truong (có doi). User thường bị chặn."""
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'username' not in session:
            if request.path.startswith('/api/'):
                return jsonify({'error': 'Vui lòng đăng nhập để tiếp tục'}), 401
            flash('Vui lòng đăng nhập để tiếp tục', 'warning')
            return redirect(url_for('auth.login'))

        role = session.get('role')
        doi = session.get('doi')

        if role == 'admin':
            return f(*args, **kwargs)
        if role == 'to_truong' and doi:
            return f(*args, **kwargs)

        if request.path.startswith('/api/'):
            return jsonify({'error': 'Yêu cầu role to_truong hoặc admin'}), 403
        flash('Bạn không có quyền truy cập tính năng này', 'danger')
        return redirect(url_for('index'))
    return decorated_function
```

- [ ] **Step 4: Run test pass**

Run: `python3 -m pytest tests/test_auth_to_truong.py -v`
Expected: 10 PASS (5 cũ + 5 mới).

- [ ] **Step 5: Commit**

```bash
git add auth.py tests/test_auth_to_truong.py
git commit -m "auth: thêm decorator to_truong_or_admin_required"
```

---

## Task 4: auth_routes set session doi + endpoint set-user-doi

**Files:**
- Modify: `auth_routes.py` (login handler set session + thêm endpoint).
- Test: `tests/test_auth_to_truong.py` (append).

**Interfaces:**
- Produces: endpoint `auth.set_user_doi`.

- [ ] **Step 1: Kiểm tra cấu trúc hiện có của auth_routes.py**

Run: `head -50 blueprints/auth_routes.py`

Expected: file có `from app_helpers import csrf_protect` (line 5) và `from auth import (...)` (lines 6-16). Login handler set session ở dòng 39-41: `session['username']`, `session['name']`, `session['role']`. Endpoint `auth.admin_users` đã tồn tại (line 142).

- [ ] **Step 2: Viết test fail**

Append vào `tests/test_auth_to_truong.py`:

```python
def test_login_sets_session_doi(tmp_path, monkeypatch):
    """Login phải set session['doi'] nếu user có cột doi."""
    import pandas as pd
    from werkzeug.security import generate_password_hash

    path = tmp_path / 'username.xlsx'
    pd.DataFrame([
        {'username': 'tt1', 'name': 'Tổ trưởng 1', 'password': generate_password_hash('pass'),
         'role': 'to_truong', 'is_first_login': False, 'is_active': True, 'doi': 'Đội 1'},
    ]).to_excel(path, index=False)
    monkeypatch.setattr('auth.EXCEL_FILE', str(path))
    from auth import invalidate_users_cache
    invalidate_users_cache()

    from dashboard import app
    client = app.test_client()
    client.post('/login', data={'username': 'tt1', 'password': 'pass'})

    with client.session_transaction() as sess:
        assert sess.get('doi') == 'Đội 1'
        assert sess.get('role') == 'to_truong'


def test_admin_set_user_doi(tmp_path, monkeypatch):
    import pandas as pd

    path = tmp_path / 'username.xlsx'
    pd.DataFrame([
        {'username': 'u1', 'name': 'User 1', 'password': 'x', 'role': 'user',
         'is_first_login': False, 'is_active': True},
    ]).to_excel(path, index=False)
    monkeypatch.setattr('auth.EXCEL_FILE', str(path))
    from auth import invalidate_users_cache
    invalidate_users_cache()

    from dashboard import app
    client = app.test_client()
    with client.session_transaction() as sess:
        sess['username'] = 'admin1'
        sess['role'] = 'admin'
        sess['_csrf_token'] = 'test-csrf'

    response = client.post('/admin/set-user-doi', data={
        'username': 'u1', 'doi': 'Đội 2', 'csrf_token': 'test-csrf',
    })

    assert response.status_code == 302
    df2 = pd.read_excel(path)
    assert df2.iloc[0]['doi'] == 'Đội 2'
```

- [ ] **Step 3: Run test fail**

Run: `python3 -m pytest tests/test_auth_to_truong.py -k "test_login_sets_session_doi or test_admin_set_user_doi" -v`
Expected: FAIL.

- [ ] **Step 4: Implement**

4a. Trong `blueprints/auth_routes.py`, sửa khối `from auth import (...)` (lines 6-16) để thêm `_safe_get_doi` và `update_user_doi`:

```python
from auth import (
    _safe_get_doi,
    admin_required,
    change_user_password,
    get_all_users,
    get_user_by_username,
    log_login_activity,
    login_required,
    reset_user_password,
    toggle_user_active,
    update_user_doi,
    verify_password,
)
```

4b. Sửa login handler (sau dòng 41 `session['role'] = user.get('role', 'user')`), thêm dòng:

```python
            session['doi'] = _safe_get_doi(user)
```

4c. Thêm endpoint cuối file `auth_routes.py`:

```python
@auth_bp.route('/admin/set-user-doi', methods=['POST'])
@admin_required
@csrf_protect
def set_user_doi():
    username = request.form.get('username', '').strip()
    new_doi = request.form.get('doi', '').strip()

    if not username:
        flash('Thiếu username', 'danger')
        return redirect(url_for('auth.admin_users'))

    ok, msg = update_user_doi(username, new_doi or None)
    flash(msg, 'success' if ok else 'danger')
    return redirect(url_for('auth.admin_users'))
```

- [ ] **Step 5: Run test pass**

Run: `python3 -m pytest tests/test_auth_to_truong.py -v`
Expected: 12 PASS.

- [ ] **Step 6: Commit**

```bash
git add blueprints/auth_routes.py tests/test_auth_to_truong.py
git commit -m "auth_routes: set session doi + endpoint admin set-user-doi"
```

---

## Task 5: Tạo blueprint checklist_bp + schema setup

**Files:**
- Create: `blueprints/checklist_routes.py` (skeleton + schema).
- Modify: `blueprints/__init__.py`, `dashboard.py`.
- Test: `tests/test_checklist_state.py` (create new).

**Interfaces:**
- Produces: `checklist_bp`, `_ensure_checklist_schema()`, `_checklist_write_connection()`, `_checklist_read_connection()`, `_checklist_schema_ready_path` (module global).

- [ ] **Step 1: Viết test fail**

Tạo `tests/test_checklist_state.py`:

```python
import sqlite3
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from blueprints import checklist_routes


def _prepare_db(tmp_path, monkeypatch):
    db_path = tmp_path / 'checklist.db'
    upload_dir = tmp_path / 'uploads'
    upload_dir.mkdir()
    monkeypatch.setattr(checklist_routes, 'CHECKLIST_DB_PATH', str(db_path))
    monkeypatch.setattr(checklist_routes, 'CHECKLIST_UPLOAD_DIR', str(upload_dir))
    checklist_routes._checklist_schema_ready_path = None
    checklist_routes._ensure_checklist_schema()
    return db_path


def test_ensure_schema_creates_3_tables(tmp_path, monkeypatch):
    db_path = _prepare_db(tmp_path, monkeypatch)

    conn = sqlite3.connect(db_path)
    tables = {row[0] for row in conn.execute(
        "SELECT name FROM sqlite_master WHERE type='table'"
    )}
    conn.close()

    assert {'checklist_state', 'checklist_photos', 'checklist_audit_log'} <= tables


def test_ensure_schema_is_idempotent(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    checklist_routes._ensure_checklist_schema()
    checklist_routes._ensure_checklist_schema()


def test_state_unique_constraint(tmp_path, monkeypatch):
    db_path = _prepare_db(tmp_path, monkeypatch)

    conn = sqlite3.connect(db_path)
    conn.execute(
        "INSERT INTO checklist_state (ngay, doi, phase_key, item_key, trang_thai) "
        "VALUES ('2026-07-22', 'Đội 1', 'sang', 'item1', 'chua_lam')"
    )
    conn.commit()

    raised = False
    try:
        conn.execute(
            "INSERT INTO checklist_state (ngay, doi, phase_key, item_key, trang_thai) "
            "VALUES ('2026-07-22', 'Đội 1', 'sang', 'item1', 'xong')"
        )
        conn.commit()
    except sqlite3.IntegrityError:
        raised = True
    finally:
        conn.close()

    assert raised
```

- [ ] **Step 2: Run test fail**

Run: `python3 -m pytest tests/test_checklist_state.py -v`
Expected: FAIL `ImportError: No module named 'blueprints.checklist_routes'`.

- [ ] **Step 3: Tạo blueprint skeleton**

Tạo `blueprints/checklist_routes.py`:

```python
import os
import sqlite3
from threading import Lock

from flask import Blueprint

from config import (
    CHECKLIST_DB_PATH,
    CHECKLIST_UPLOAD_DIR,
    CHECKLIST_PHOTO_MAX_BYTES,
    CHECKLIST_PHOTO_MAX_PER_ITEM,
)

checklist_bp = Blueprint('checklist', __name__)

_checklist_schema_lock = Lock()
_checklist_schema_ready_path = None


def _checklist_write_connection():
    conn = sqlite3.connect(CHECKLIST_DB_PATH, timeout=5)
    conn.row_factory = sqlite3.Row
    conn.execute('PRAGMA journal_mode=WAL')
    conn.execute('PRAGMA busy_timeout=5000')
    return conn


def _checklist_read_connection():
    conn = sqlite3.connect(
        f'file:{CHECKLIST_DB_PATH}?mode=ro',
        uri=True,
        timeout=5,
    )
    conn.row_factory = sqlite3.Row
    return conn


def _ensure_checklist_schema():
    global _checklist_schema_ready_path
    if _checklist_schema_ready_path == CHECKLIST_DB_PATH:
        return

    db_dir = os.path.dirname(CHECKLIST_DB_PATH)
    if db_dir:
        os.makedirs(db_dir, exist_ok=True)
    os.makedirs(CHECKLIST_UPLOAD_DIR, exist_ok=True)

    with _checklist_schema_lock:
        if _checklist_schema_ready_path == CHECKLIST_DB_PATH:
            return
        with _checklist_write_connection() as conn:
            conn.executescript('''
                CREATE TABLE IF NOT EXISTS checklist_state (
                    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
                    ngay                TEXT NOT NULL,
                    doi                 TEXT NOT NULL,
                    phase_key           TEXT NOT NULL,
                    item_key            TEXT NOT NULL,
                    trang_thai          TEXT NOT NULL DEFAULT 'chua_lam',
                    ghi_chu             TEXT,
                    nguoi_nhap          TEXT,
                    nguoi_nhap_name     TEXT,
                    thoi_diem_nhap      TEXT,
                    thoi_diem_cap_nhat  TEXT,
                    UNIQUE(ngay, doi, phase_key, item_key)
                );
                CREATE INDEX IF NOT EXISTS idx_checklist_state_ngay_doi
                    ON checklist_state(ngay, doi);

                CREATE TABLE IF NOT EXISTS checklist_photos (
                    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
                    ngay                TEXT NOT NULL,
                    doi                 TEXT NOT NULL,
                    phase_key           TEXT NOT NULL,
                    item_key            TEXT NOT NULL,
                    filename_stored     TEXT NOT NULL,
                    filename_original   TEXT,
                    mime_type           TEXT,
                    size_bytes          INTEGER,
                    nguoi_nhap          TEXT,
                    thoi_diem_nhap      TEXT
                );
                CREATE INDEX IF NOT EXISTS idx_checklist_photos_item
                    ON checklist_photos(ngay, doi, phase_key, item_key);

                CREATE TABLE IF NOT EXISTS checklist_audit_log (
                    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
                    ngay                TEXT NOT NULL,
                    doi                 TEXT NOT NULL,
                    phase_key           TEXT NOT NULL,
                    item_key            TEXT NOT NULL,
                    action              TEXT NOT NULL,
                    payload             TEXT,
                    nguoi_thuc_hien     TEXT,
                    thoi_diem           TEXT NOT NULL
                );
                CREATE INDEX IF NOT EXISTS idx_checklist_audit_ngay_doi
                    ON checklist_audit_log(ngay, doi);
            ''')
        _checklist_schema_ready_path = CHECKLIST_DB_PATH
```

- [ ] **Step 4: Đăng ký blueprint**

4a. Mở `blueprints/__init__.py`. Thêm `from .checklist_routes import checklist_bp` + thêm `'checklist_bp'` vào `__all__`.

4b. Mở `dashboard.py`. Thêm `checklist_bp` vào import từ `blueprints`, rồi `app.register_blueprint(checklist_bp)` ở khu vực register hiện có.

- [ ] **Step 5: Run test pass**

Run: `python3 -m pytest tests/test_checklist_state.py -v`
Expected: 3 PASS.

- [ ] **Step 6: Smoke test app start**

Run: `python3 -c "from dashboard import app; print('OK')"`
Expected: in `OK`, không có ImportError.

- [ ] **Step 7: Commit**

```bash
git add blueprints/checklist_routes.py blueprints/__init__.py dashboard.py tests/test_checklist_state.py
git commit -m "checklist: tạo blueprint + schema 3 bảng (state/photos/audit_log)"
```

---

## Task 6: CHECKLIST_TEMPLATE constant

**Files:**
- Modify: `blueprints/checklist_routes.py` (thêm constant).
- Test: `tests/test_checklist_state.py` (append).

**Interfaces:**
- Produces: hằng số `CHECKLIST_TEMPLATE` (list 3 phases × items).

- [ ] **Step 1: Viết test fail**

Append vào `tests/test_checklist_state.py`:

```python
def test_template_has_3_phases_15_items():
    template = checklist_routes.CHECKLIST_TEMPLATE

    assert len(template) == 3
    total_items = sum(len(p['items']) for p in template)
    assert total_items == 15


def test_template_phase_keys_unique():
    template = checklist_routes.CHECKLIST_TEMPLATE
    keys = [p['phase_key'] for p in template]
    assert len(keys) == len(set(keys))


def test_template_item_keys_unique_per_phase():
    template = checklist_routes.CHECKLIST_TEMPLATE
    for phase in template:
        keys = [i['item_key'] for i in phase['items']]
        assert len(keys) == len(set(keys))


def test_template_has_5_metric_sources():
    template = checklist_routes.CHECKLIST_TEMPLATE
    metric_items = [
        i for p in template for i in p['items'] if i.get('metric_source')
    ]
    assert len(metric_items) == 5
```

- [ ] **Step 2: Run test fail**

Run: `python3 -m pytest tests/test_checklist_state.py -k template -v`
Expected: FAIL `AttributeError`.

- [ ] **Step 3: Implement**

Thêm vào `blueprints/checklist_routes.py` ngay sau `_ensure_checklist_schema`:

```python
CHECKLIST_TEMPLATE = [
    {
        'phase_key': 'sang_giao_viec',
        'phase_name': 'Đầu giờ sáng (07:15-07:45): Giao việc',
        'time_window': '07:15 - 07:45',
        'items': [
            {'item_key': 'atvsld_trang_phuc_ccdc',
             'name': 'Đánh giá ATVSLĐ, trang phục, CCDC toàn tổ', 'metric_source': None},
            {'item_key': 'pho_bien_vb_moi',
             'name': 'Phổ biến văn bản mới VNPT HNi / TTVT', 'metric_source': None},
            {'item_key': 'giai_dap_vuong_mac',
             'name': 'Giải đáp hướng dẫn vướng mắc NVKT', 'metric_source': None},
            {'item_key': 'doichieu_vattu_hoancong',
             'name': 'Đối chiếu vật tư tồn kho, hoàn công/treỏo', 'metric_source': 'inventory'},
            {'item_key': 'giao_chitieu_catondong',
             'name': 'Giao chỉ tiêu xử lý dứt điểm ca tồn đọng',
             'metric_source': 'brcd_pttb_overdue'},
            {'item_key': 'phancong_hotro_cheo',
             'name': 'Phân công nhân sự hỗ trợ chéo', 'metric_source': None},
            {'item_key': 'giao_chitieu_nangsuat',
             'name': 'Giao chỉ tiêu năng suất ngày (suy hao/chạm cước)', 'metric_source': None},
        ],
    },
    {
        'phase_key': 'trong_ngay_dieu_phoi',
        'phase_name': 'Trong ngày (07:45-18:00): Điều phối',
        'time_window': '07:45 - 18:00',
        'items': [
            {'item_key': 'giamsat_dieu_huong',
             'name': 'Giám sát điều hướng ưu tiên (sự cố > lắp mới > dịch chuyển > suy hao > gia hạn)',
             'metric_source': 'brcd_pttb_priority'},
            {'item_key': 'capphat_vattu_s2s1',
             'name': 'Cấp phát vật tư bổ sung, hỗ trợ S2/S1', 'metric_source': None},
            {'item_key': 'bori_lich_thucdia',
             'name': 'Bố trí lịch thực địa (1-2 buổi/tuần)', 'metric_source': None},
            {'item_key': 'trieu_tap_dot_xuat',
             'name': 'Triệu tập đột xuất 13h15 NV năng suất thấp',
             'metric_source': 'brcd_nvkt_low_productivity'},
        ],
    },
    {
        'phase_key': 'cuoi_ngay_danh_gia',
        'phase_name': 'Cuối ngày (trước 18:00): Đánh giá kỷ luật',
        'time_window': 'trước 18:00',
        'items': [
            {'item_key': 'tiepnhan_thietbi_thuhoi',
             'name': 'Tiếp nhận thiết bị thu hồi, vật tư thừa trả', 'metric_source': None},
            {'item_key': 'rasoat_nguyennhan_ton',
             'name': 'Rà soát 100% NV cập nhật nguyên nhân tồn',
             'metric_source': 'kiemsoat_completion'},
            {'item_key': 'lap_bienban_vipham',
             'name': 'Lập biên bản vi phạm kỷ luật', 'metric_source': None},
            {'item_key': 'tonghop_dutru_vattu',
             'name': 'Tổng hợp dự trù vật tư, kế hoạch nhân sự ngày sau', 'metric_source': None},
        ],
    },
]
```

- [ ] **Step 4: Run test pass**

Run: `python3 -m pytest tests/test_checklist_state.py -v`
Expected: 7 PASS.

- [ ] **Step 5: Commit**

```bash
git add blueprints/checklist_routes.py tests/test_checklist_state.py
git commit -m "checklist: thêm CHECKLIST_TEMPLATE (3 phases × 15 items)"
```

---

## Task 7: Endpoint POST /api/checklist/luu (UPSERT state + audit log)

**Files:**
- Modify: `blueprints/checklist_routes.py` (thêm helpers + endpoint).
- Test: `tests/test_checklist_state.py` (append).

**Interfaces:**
- Produces: endpoint `checklist.api_checklist_luu`, helpers `_log_audit(...)`, `_save_state_record(...)`.

- [ ] **Step 1: Viết test fail**

Append vào `tests/test_checklist_state.py`:

```python
def _logged_in_to_truong_client():
    from dashboard import app
    client = app.test_client()
    with client.session_transaction() as sess:
        sess['username'] = 'tt1'
        sess['name'] = 'Tổ trưởng 1'
        sess['role'] = 'to_truong'
        sess['doi'] = 'Đội 1'
    return client


def test_luu_creates_new_state(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    client = _logged_in_to_truong_client()

    response = client.post('/api/checklist/luu', json={
        'date': '2026-07-22', 'doi': 'Đội 1',
        'phase_key': 'sang_giao_viec', 'item_key': 'atvsld_trang_phuc_ccdc',
        'trang_thai': 'xong', 'ghi_chu': 'Đã kiểm tra 8/8 NV',
    })

    assert response.status_code == 200
    assert response.get_json()['ok'] == 1

    import sqlite3 as s3
    conn = s3.connect(tmp_path / 'checklist.db')
    row = conn.execute(
        "SELECT trang_thai, ghi_chu, nguoi_nhap FROM checklist_state "
        "WHERE ngay=? AND doi=? AND phase_key=? AND item_key=?",
        ('2026-07-22', 'Đội 1', 'sang_giao_viec', 'atvsld_trang_phuc_ccdc')
    ).fetchone()
    conn.close()
    assert row[0] == 'xong'
    assert row[1] == 'Đã kiểm tra 8/8 NV'
    assert row[2] == 'tt1'


def test_luu_upserts_existing_state(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    client = _logged_in_to_truong_client()

    for status in ['dang_lam', 'xong']:
        client.post('/api/checklist/luu', json={
            'date': '2026-07-22', 'doi': 'Đội 1',
            'phase_key': 'sang_giao_viec', 'item_key': 'atvsld_trang_phuc_ccdc',
            'trang_thai': status, 'ghi_chu': status,
        })

    import sqlite3 as s3
    conn = s3.connect(tmp_path / 'checklist.db')
    count = conn.execute(
        "SELECT COUNT(*) FROM checklist_state WHERE ngay=? AND doi=? AND phase_key=? AND item_key=?",
        ('2026-07-22', 'Đội 1', 'sang_giao_viec', 'atvsld_trang_phuc_ccdc')
    ).fetchone()[0]
    conn.close()
    assert count == 1


def test_luu_deletes_when_empty_and_chua_lam(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    client = _logged_in_to_truong_client()

    client.post('/api/checklist/luu', json={
        'date': '2026-07-22', 'doi': 'Đội 1',
        'phase_key': 'sang_giao_viec', 'item_key': 'atvsld_trang_phuc_ccdc',
        'trang_thai': 'xong', 'ghi_chu': 'Done',
    })
    client.post('/api/checklist/luu', json={
        'date': '2026-07-22', 'doi': 'Đội 1',
        'phase_key': 'sang_giao_viec', 'item_key': 'atvsld_trang_phuc_ccdc',
        'trang_thai': 'chua_lam', 'ghi_chu': '',
    })

    import sqlite3 as s3
    conn = s3.connect(tmp_path / 'checklist.db')
    count = conn.execute(
        "SELECT COUNT(*) FROM checklist_state WHERE ngay=? AND doi=? AND phase_key=? AND item_key=?",
        ('2026-07-22', 'Đội 1', 'sang_giao_viec', 'atvsld_trang_phuc_ccdc')
    ).fetchone()[0]
    conn.close()
    assert count == 0


def test_luu_rejects_doi_mismatch(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    client = _logged_in_to_truong_client()

    response = client.post('/api/checklist/luu', json={
        'date': '2026-07-22', 'doi': 'Đội 99',
        'phase_key': 'sang_giao_viec', 'item_key': 'atvsld_trang_phuc_ccdc',
        'trang_thai': 'xong', 'ghi_chu': '',
    })
    assert response.status_code == 403


def test_luu_writes_audit_log(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    client = _logged_in_to_truong_client()

    client.post('/api/checklist/luu', json={
        'date': '2026-07-22', 'doi': 'Đội 1',
        'phase_key': 'sang_giao_viec', 'item_key': 'atvsld_trang_phuc_ccdc',
        'trang_thai': 'xong', 'ghi_chu': 'Done',
    })

    import sqlite3 as s3
    conn = s3.connect(tmp_path / 'checklist.db')
    count = conn.execute(
        "SELECT COUNT(*) FROM checklist_audit_log WHERE action='upsert_state'"
    ).fetchone()[0]
    conn.close()
    assert count >= 1
```

- [ ] **Step 2: Run test fail**

Run: `python3 -m pytest tests/test_checklist_state.py -k luu -v`
Expected: FAIL 404.

- [ ] **Step 3: Implement**

Trong `blueprints/checklist_routes.py`:

3a. Sửa khối imports đầu file thành:

```python
import json
import os
import sqlite3
from datetime import datetime
from threading import Lock

from flask import Blueprint, jsonify, request, session

from auth import login_required, to_truong_or_admin_required
from config import (
    CHECKLIST_DB_PATH,
    CHECKLIST_UPLOAD_DIR,
    CHECKLIST_PHOTO_MAX_BYTES,
    CHECKLIST_PHOTO_MAX_PER_ITEM,
)
```

3b. Thêm helper `_log_audit` (sau `_ensure_checklist_schema`, trước `CHECKLIST_TEMPLATE`):

```python
def _log_audit(ngay, doi, phase_key, item_key, action, payload, nguoi_thuc_hien):
    """Append-only audit log. Best-effort — bỏ qua lỗi."""
    try:
        with _checklist_write_connection() as conn:
            conn.execute(
                "INSERT INTO checklist_audit_log "
                "(ngay, doi, phase_key, item_key, action, payload, nguoi_thuc_hien, thoi_diem) "
                "VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                (ngay, doi, phase_key, item_key, action,
                 json.dumps(payload, ensure_ascii=False) if payload else None,
                 nguoi_thuc_hien, datetime.now().isoformat(timespec='seconds'))
            )
    except Exception:
        pass
```

3c. Thêm helper `_save_state_record`:

```python
def _save_state_record(ngay, doi, phase_key, item_key, trang_thai, ghi_chu,
                       username, name):
    """UPSERT hoặc DELETE row checklist_state + audit log. Trả dict new_state (hoặc None)."""
    is_delete = (trang_thai == 'chua_lam') and (not ghi_chu or not ghi_chu.strip())

    with _checklist_write_connection() as conn:
        existing = conn.execute(
            "SELECT * FROM checklist_state WHERE ngay=? AND doi=? AND phase_key=? AND item_key=?",
            (ngay, doi, phase_key, item_key)
        ).fetchone()

        if is_delete:
            if existing:
                conn.execute(
                    "DELETE FROM checklist_state WHERE ngay=? AND doi=? AND phase_key=? AND item_key=?",
                    (ngay, doi, phase_key, item_key)
                )
            _log_audit(ngay, doi, phase_key, item_key, 'delete_state',
                       {'old': dict(existing) if existing else None, 'new': None}, username)
            return None

        now = datetime.now().isoformat(timespec='seconds')
        if existing:
            conn.execute(
                "UPDATE checklist_state SET trang_thai=?, ghi_chu=?, nguoi_nhap=?, "
                "nguoi_nhap_name=?, thoi_diem_cap_nhat=? "
                "WHERE ngay=? AND doi=? AND phase_key=? AND item_key=?",
                (trang_thai, ghi_chu, username, name, now,
                 ngay, doi, phase_key, item_key)
            )
            new_state = {'trang_thai': trang_thai, 'ghi_chu': ghi_chu,
                         'nguoi_nhap': username, 'thoi_diem_cap_nhat': now}
        else:
            conn.execute(
                "INSERT INTO checklist_state "
                "(ngay, doi, phase_key, item_key, trang_thai, ghi_chu, "
                "nguoi_nhap, nguoi_nhap_name, thoi_diem_nhap, thoi_diem_cap_nhat) "
                "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                (ngay, doi, phase_key, item_key, trang_thai, ghi_chu,
                 username, name, now, now)
            )
            new_state = {'trang_thai': trang_thai, 'ghi_chu': ghi_chu,
                         'nguoi_nhap': username, 'thoi_diem_nhap': now,
                         'thoi_diem_cap_nhat': now}

    _log_audit(ngay, doi, phase_key, item_key, 'upsert_state',
               {'old': dict(existing) if existing else None, 'new': new_state}, username)
    return new_state
```

3d. Thêm endpoint:

```python
@checklist_bp.route('/api/checklist/luu', methods=['POST'])
@login_required
@to_truong_or_admin_required
def api_checklist_luu():
    data = request.get_json(silent=True) or {}

    ngay = (data.get('date') or '').strip()
    doi = (data.get('doi') or '').strip()
    phase_key = (data.get('phase_key') or '').strip()
    item_key = (data.get('item_key') or '').strip()
    trang_thai = (data.get('trang_thai') or 'chua_lam').strip()
    ghi_chu = (data.get('ghi_chu') or '').strip()

    if not (ngay and doi and phase_key and item_key):
        return jsonify({'error': 'Thiếu date/doi/phase_key/item_key'}), 400
    if trang_thai not in ('chua_lam', 'dang_lam', 'xong', 'bo_qua'):
        return jsonify({'error': 'trang_thai không hợp lệ'}), 400
    if len(ghi_chu) > 2000:
        return jsonify({'error': 'ghi_chu vượt 2000 ký tự'}), 400

    if session.get('role') != 'admin' and doi != session.get('doi'):
        return jsonify({'error': 'Bạn chỉ sửa được đội của mình'}), 403

    _ensure_checklist_schema()
    try:
        new_state = _save_state_record(
            ngay, doi, phase_key, item_key, trang_thai, ghi_chu,
            session.get('username'), session.get('name')
        )
    except sqlite3.OperationalError as exc:
        if 'locked' in str(exc).lower() or 'busy' in str(exc).lower():
            return jsonify({'error': 'db_locked'}), 503
        raise

    return jsonify({'ok': 1, 'state': new_state})
```

- [ ] **Step 4: Run test pass**

Run: `python3 -m pytest tests/test_checklist_state.py -v`
Expected: 12 PASS.

- [ ] **Step 5: Commit**

```bash
git add blueprints/checklist_routes.py tests/test_checklist_state.py
git commit -m "checklist: endpoint POST /api/checklist/luu (UPSERT state + audit log)"
```

---

## Task 8: Endpoint GET /api/checklist/data

**Files:**
- Modify: `blueprints/checklist_routes.py` (thêm helpers + endpoint).
- Test: `tests/test_checklist_routes.py` (create new).

**Interfaces:**
- Produces: endpoint `checklist.api_checklist_data`, helpers `_load_state_map`, `_load_photos_map`, `_compute_can_edit`, `_list_available_dois`.

- [ ] **Step 1: Viết test fail**

Tạo `tests/test_checklist_routes.py`:

```python
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from blueprints import checklist_routes


def _prepare_db(tmp_path, monkeypatch):
    db_path = tmp_path / 'checklist.db'
    upload_dir = tmp_path / 'uploads'
    upload_dir.mkdir()
    monkeypatch.setattr(checklist_routes, 'CHECKLIST_DB_PATH', str(db_path))
    monkeypatch.setattr(checklist_routes, 'CHECKLIST_UPLOAD_DIR', str(upload_dir))
    checklist_routes._checklist_schema_ready_path = None
    checklist_routes._ensure_checklist_schema()
    return db_path


def _logged_in_to_truong_client(doi='Đội 1'):
    from dashboard import app
    client = app.test_client()
    with client.session_transaction() as sess:
        sess['username'] = 'tt1'
        sess['name'] = 'Tổ trưởng 1'
        sess['role'] = 'to_truong'
        sess['doi'] = doi
    return client


def test_data_returns_template_with_3_phases(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    client = _logged_in_to_truong_client()

    response = client.get('/api/checklist/data?date=2026-07-22&doi=Đội 1')

    assert response.status_code == 200
    body = response.get_json()
    assert body['date'] == '2026-07-22'
    assert body['doi'] == 'Đội 1'
    assert len(body['phases']) == 3
    assert body['can_edit'] is True


def test_data_returns_can_edit_false_for_other_doi(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    client = _logged_in_to_truong_client(doi='Đội 1')

    response = client.get('/api/checklist/data?date=2026-07-22&doi=Đội 2')
    body = response.get_json()
    assert body['can_edit'] is False


def test_data_returns_state_and_photos_per_item(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    client = _logged_in_to_truong_client()

    client.post('/api/checklist/luu', json={
        'date': '2026-07-22', 'doi': 'Đội 1',
        'phase_key': 'sang_giao_viec', 'item_key': 'atvsld_trang_phuc_ccdc',
        'trang_thai': 'xong', 'ghi_chu': 'Done',
    })

    response = client.get('/api/checklist/data?date=2026-07-22&doi=Đội 1')
    body = response.get_json()
    item = body['phases'][0]['items'][0]
    assert item['item_key'] == 'atvsld_trang_phuc_ccdc'
    assert item['state']['trang_thai'] == 'xong'
    assert item['state']['ghi_chu'] == 'Done'
    assert item['photos'] == []


def test_data_uses_today_when_no_date(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    client = _logged_in_to_truong_client()

    response = client.get('/api/checklist/data')
    body = response.get_json()
    from datetime import date
    assert body['date'] == date.today().isoformat()


def test_data_uses_session_doi_when_no_doi(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    client = _logged_in_to_truong_client(doi='Đội 5')

    response = client.get('/api/checklist/data?date=2026-07-22')
    body = response.get_json()
    assert body['doi'] == 'Đội 5'
```

- [ ] **Step 2: Run test fail**

Run: `python3 -m pytest tests/test_checklist_routes.py -v`
Expected: FAIL 404.

- [ ] **Step 3: Implement**

Thêm vào `blueprints/checklist_routes.py` (sau `api_checklist_luu`):

```python
def _load_state_map(ngay, doi):
    """Trả dict {(phase_key, item_key): state_dict} cho 1 (ngay, doi)."""
    try:
        with _checklist_read_connection() as conn:
            rows = conn.execute(
                "SELECT phase_key, item_key, trang_thai, ghi_chu, nguoi_nhap, "
                "nguoi_nhap_name, thoi_diem_nhap, thoi_diem_cap_nhat "
                "FROM checklist_state WHERE ngay=? AND doi=?",
                (ngay, doi)
            ).fetchall()
    except sqlite3.OperationalError:
        return {}

    return {
        (r['phase_key'], r['item_key']): {
            'trang_thai': r['trang_thai'],
            'ghi_chu': r['ghi_chu'],
            'nguoi_nhap': r['nguoi_nhap'],
            'nguoi_nhap_name': r['nguoi_nhap_name'],
            'thoi_diem_nhap': r['thoi_diem_nhap'],
            'thoi_diem_cap_nhat': r['thoi_diem_cap_nhat'],
        }
        for r in rows
    }


def _load_photos_map(ngay, doi):
    """Trả dict {(phase_key, item_key): [photo_dict, ...]} cho 1 (ngay, doi)."""
    try:
        with _checklist_read_connection() as conn:
            rows = conn.execute(
                "SELECT id, phase_key, item_key, filename_stored, filename_original, "
                "size_bytes, nguoi_nhap, thoi_diem_nhap "
                "FROM checklist_photos WHERE ngay=? AND doi=? ORDER BY id",
                (ngay, doi)
            ).fetchall()
    except sqlite3.OperationalError:
        return {}

    result = {}
    for r in rows:
        key = (r['phase_key'], r['item_key'])
        result.setdefault(key, []).append({
            'id': r['id'],
            'url': f"/checklist/photo/{r['filename_stored']}",
            'filename_original': r['filename_original'],
            'size_bytes': r['size_bytes'],
            'nguoi_nhap': r['nguoi_nhap'],
            'thoi_diem_nhap': r['thoi_diem_nhap'],
        })
    return result


def _compute_can_edit(doi):
    """True nếu user hiện tại được sửa đội này."""
    role = session.get('role')
    if role == 'admin':
        return True
    if role == 'to_truong' and session.get('doi') == doi:
        return True
    return False


def _list_available_dois():
    """SELECT DISTINCT doi_vt FROM brcd_phieu. Fallback []."""
    try:
        from config import BRCD_KIEMSOAT_DB_PATH
        if not os.path.exists(BRCD_KIEMSOAT_DB_PATH):
            return []
        conn = sqlite3.connect(
            f'file:{BRCD_KIEMSOAT_DB_PATH}?mode=ro', uri=True, timeout=2
        )
        rows = conn.execute(
            "SELECT DISTINCT doi_vt FROM brcd_phieu WHERE doi_vt IS NOT NULL ORDER BY 1"
        ).fetchall()
        conn.close()
        return [r[0] for r in rows if r[0]]
    except Exception:
        return []


@checklist_bp.route('/api/checklist/data')
@login_required
def api_checklist_data():
    from datetime import date
    import copy

    ngay = (request.args.get('date') or '').strip() or date.today().isoformat()
    doi = (request.args.get('doi') or '').strip() or session.get('doi', '')

    if not doi:
        return jsonify({'error': 'Thiếu doi và session không có doi'}), 400

    _ensure_checklist_schema()
    state_map = _load_state_map(ngay, doi)
    photos_map = _load_photos_map(ngay, doi)

    template = copy.deepcopy(CHECKLIST_TEMPLATE)
    for phase in template:
        for item in phase['items']:
            key = (phase['phase_key'], item['item_key'])
            item['state'] = state_map.get(key)
            item['photos'] = photos_map.get(key, [])
            item['metric'] = None

    return jsonify({
        'date': ngay,
        'doi': doi,
        'available_dois': _list_available_dois(),
        'can_edit': _compute_can_edit(doi),
        'phases': template,
    })
```

- [ ] **Step 4: Run test pass**

Run: `python3 -m pytest tests/test_checklist_routes.py -v`
Expected: 5 PASS.

- [ ] **Step 5: Commit**

```bash
git add blueprints/checklist_routes.py tests/test_checklist_routes.py
git commit -m "checklist: endpoint GET /api/checklist/data (template + state + photos)"
```

---

## Task 9: Page shell + sidebar + route_policy

**Files:**
- Create: `templates/pages/checklist.html`.
- Modify: `templates/base.html`, `route_policy.py`.
- Modify: `blueprints/checklist_routes.py` (thêm `page_checklist`).
- Test: `tests/test_checklist_routes.py` (append).

**Interfaces:**
- Produces: endpoint `checklist.page_checklist`, `active_page='checklist'`.

- [ ] **Step 1: Viết test fail**

Append vào `tests/test_checklist_routes.py`:

```python
def test_page_checklist_returns_200(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    client = _logged_in_to_truong_client()

    response = client.get('/checklist')

    assert response.status_code == 200
    assert b'checklist' in response.data.lower()
```

- [ ] **Step 2: Run test fail**

Run: `python3 -m pytest tests/test_checklist_routes.py -k test_page_checklist -v`
Expected: FAIL 404.

- [ ] **Step 3: Implement**

3a. Thêm endpoint + import `_current_user`:

Trong `blueprints/checklist_routes.py`, thêm vào imports:

```python
from flask import Blueprint, jsonify, render_template, request, session
from app_helpers import _current_user
```

(Sửa dòng import `from flask import` hiện có để thêm `render_template`.)

Thêm endpoint:

```python
@checklist_bp.route('/checklist')
@login_required
def page_checklist():
    _ensure_checklist_schema()
    return render_template(
        'pages/checklist.html',
        current_user=_current_user(),
        active_page='checklist',
    )
```

3b. Tạo `templates/pages/checklist.html`:

```html
{% extends "base.html" %}

{% block content %}
<div class="dashboard-header">
    <h1><i class="fas fa-clipboard-check"></i> Checklist tổ trưởng tổ KTĐB</h1>
    <div id="checklist-unit-name">Đơn vị: {{ config.UNIT_NAME }}</div>
</div>

<div class="checklist-toolbar">
    <div class="checklist-filters">
        <label>Ngày:
            <button id="checklist-prev-day" type="button">&lt;</button>
            <input type="date" id="checklist-date-input">
            <button id="checklist-next-day" type="button">&gt;</button>
            <button id="checklist-today" type="button">Hôm nay</button>
        </label>
        <label>Đội:
            <select id="checklist-doi-select"></select>
        </label>
        <div class="checklist-progress">
            Tiến độ: <span id="checklist-progress-text">—</span>
        </div>
    </div>
    <button id="checklist-export-btn" type="button" class="btn-excel">
        <i class="fas fa-file-excel"></i> Xuất Excel
    </button>
</div>

<div id="checklist-phases-container">
    <p class="checklist-loading">Đang tải...</p>
</div>
{% endblock %}

{% block extra_js %}
<link rel="stylesheet" href="{{ url_for('static', filename='css/checklist.css') }}">
<script src="{{ url_for('static', filename='js/pages/checklist.js') }}"></script>
{% endblock %}
```

3c. Sửa `route_policy.py`: thêm entry `'checklist.page_checklist': 'checklist'` vào dict `PAGE_ACTIVE_KEYS` (xem dòng ~8-34).

3d. Sửa `templates/base.html`: tìm sidebar menu `/shc-cts` (khoảng dòng 182-185). Thêm sau đó (sau nhóm kiemsoat):

```html
{% if is_endpoint_enabled('checklist.page_checklist') %}
<a href="{{ url_for('checklist.page_checklist') }}"
   class="menu-item {% if active_page == 'checklist' %}active{% endif %}">
    <i class="fas fa-clipboard-check"></i>
    <span>14. Checklist tổ trưởng</span>
</a>
{% endif %}
```

- [ ] **Step 4: Run test pass**

Run: `python3 -m pytest tests/test_checklist_routes.py -v`
Expected: 6 PASS.

- [ ] **Step 5: Commit**

```bash
git add blueprints/checklist_routes.py templates/pages/checklist.html templates/base.html route_policy.py tests/test_checklist_routes.py
git commit -m "checklist: page shell + sidebar menu + route_policy"
```

---

## Task 10: Photo upload endpoint + validation

**Files:**
- Modify: `blueprints/checklist_routes.py`.
- Test: `tests/test_checklist_photos.py` (create new).

**Interfaces:**
- Produces: endpoint `checklist.api_checklist_upload_photo`, helper `_validate_photo`.

- [ ] **Step 1: Viết test fail**

Tạo `tests/test_checklist_photos.py`:

```python
import io
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from blueprints import checklist_routes


def _prepare_db(tmp_path, monkeypatch):
    db_path = tmp_path / 'checklist.db'
    upload_dir = tmp_path / 'uploads'
    upload_dir.mkdir()
    monkeypatch.setattr(checklist_routes, 'CHECKLIST_DB_PATH', str(db_path))
    monkeypatch.setattr(checklist_routes, 'CHECKLIST_UPLOAD_DIR', str(upload_dir))
    checklist_routes._checklist_schema_ready_path = None
    checklist_routes._ensure_checklist_schema()
    return db_path, upload_dir


def _logged_in_to_truong_client(doi='Đội 1'):
    from dashboard import app
    client = app.test_client()
    with client.session_transaction() as sess:
        sess['username'] = 'tt1'
        sess['name'] = 'Tổ trưởng 1'
        sess['role'] = 'to_truong'
        sess['doi'] = doi
        sess['_csrf_token'] = 'test-csrf'
    return client


def test_upload_photo_success(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    client = _logged_in_to_truong_client()

    response = client.post('/api/checklist/upload-photo', data={
        'date': '2026-07-22', 'doi': 'Đội 1',
        'phase_key': 'sang_giao_viec', 'item_key': 'atvsld_trang_phuc_ccdc',
        'photo': (io.BytesIO(b'\xFF\xD8\xFF\xE0' + b'\0' * 100), 'test.jpg'),
        'csrf_token': 'test-csrf',
    }, content_type='multipart/form-data')

    assert response.status_code == 200
    body = response.get_json()
    assert body['ok'] == 1
    assert body['photo']['filename_original'] == 'test.jpg'
    assert body['photo']['url'].startswith('/checklist/photo/')


def test_upload_photo_rejects_wrong_extension(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    client = _logged_in_to_truong_client()

    response = client.post('/api/checklist/upload-photo', data={
        'date': '2026-07-22', 'doi': 'Đội 1',
        'phase_key': 'sang_giao_viec', 'item_key': 'atvsld_trang_phuc_ccdc',
        'photo': (io.BytesIO(b'x'), 'malicious.exe'),
        'csrf_token': 'test-csrf',
    }, content_type='multipart/form-data')

    assert response.status_code == 400


def test_upload_photo_rejects_doi_mismatch(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    client = _logged_in_to_truong_client(doi='Đội 1')

    response = client.post('/api/checklist/upload-photo', data={
        'date': '2026-07-22', 'doi': 'Đội 99',
        'phase_key': 'sang_giao_viec', 'item_key': 'atvsld_trang_phuc_ccdc',
        'photo': (io.BytesIO(b'\xFF\xD8\xFF\xE0' + b'\0' * 100), 'test.jpg'),
        'csrf_token': 'test-csrf',
    }, content_type='multipart/form-data')

    assert response.status_code == 403


def test_upload_photo_rejects_too_many(tmp_path, monkeypatch):
    monkeypatch.setattr(checklist_routes, 'CHECKLIST_PHOTO_MAX_PER_ITEM', 2)
    _prepare_db(tmp_path, monkeypatch)
    client = _logged_in_to_truong_client()

    for i in range(2):
        r = client.post('/api/checklist/upload-photo', data={
            'date': '2026-07-22', 'doi': 'Đội 1',
            'phase_key': 'sang_giao_viec', 'item_key': 'atvsld_trang_phuc_ccdc',
            'photo': (io.BytesIO(b'\xFF\xD8\xFF\xE0' + b'\0' * 100), f't{i}.jpg'),
            'csrf_token': 'test-csrf',
        }, content_type='multipart/form-data')
        assert r.status_code == 200

    r3 = client.post('/api/checklist/upload-photo', data={
        'date': '2026-07-22', 'doi': 'Đội 1',
        'phase_key': 'sang_giao_viec', 'item_key': 'atvsld_trang_phuc_ccdc',
        'photo': (io.BytesIO(b'\xFF\xD8\xFF\xE0' + b'\0' * 100), 't3.jpg'),
        'csrf_token': 'test-csrf',
    }, content_type='multipart/form-data')

    assert r3.status_code == 400
    assert 'giới hạn' in r3.get_json()['error'].lower()


def test_upload_photo_stores_in_month_dir(tmp_path, monkeypatch):
    _, upload_dir = _prepare_db(tmp_path, monkeypatch)
    client = _logged_in_to_truong_client()

    client.post('/api/checklist/upload-photo', data={
        'date': '2026-07-22', 'doi': 'Đội 1',
        'phase_key': 'sang_giao_viec', 'item_key': 'atvsld_trang_phuc_ccdc',
        'photo': (io.BytesIO(b'\xFF\xD8\xFF\xE0' + b'\0' * 100), 'test.jpg'),
        'csrf_token': 'test-csrf',
    }, content_type='multipart/form-data')

    month_dir = upload_dir / '2026-07'
    assert month_dir.is_dir()
    files = list(month_dir.iterdir())
    assert len(files) == 1
    assert files[0].suffix == '.jpg'
```

- [ ] **Step 2: Run test fail**

Run: `python3 -m pytest tests/test_checklist_photos.py -v`
Expected: FAIL 404.

- [ ] **Step 3: Implement**

Trong `blueprints/checklist_routes.py`:

3a. Thêm imports:

```python
import re
import uuid
from werkzeug.utils import secure_filename
from app_helpers import csrf_protect
```

3b. Helper `_validate_photo`:

```python
_PHOTO_ALLOWED_EXT = {'.jpg', '.jpeg', '.png', '.webp'}
_PHOTO_ALLOWED_MIME = {'image/jpeg', 'image/png', 'image/webp'}


def _validate_photo(photo):
    """Trả (ext, error_msg). ext=None khi invalid."""
    if photo is None:
        return None, 'Thiếu file'
    if not photo.filename:
        return None, 'File rỗng'

    ext = os.path.splitext(photo.filename)[1].lower()
    if ext not in _PHOTO_ALLOWED_EXT:
        return None, f'Định dạng không hợp lệ: {ext}'

    if photo.mimetype not in _PHOTO_ALLOWED_MIME:
        return None, f'MIME không hợp lệ: {photo.mimetype}'

    return ext, None
```

3c. Endpoint:

```python
@checklist_bp.route('/api/checklist/upload-photo', methods=['POST'])
@login_required
@to_truong_or_admin_required
@csrf_protect
def api_checklist_upload_photo():
    ngay = (request.form.get('date') or '').strip()
    doi = (request.form.get('doi') or '').strip()
    phase_key = (request.form.get('phase_key') or '').strip()
    item_key = (request.form.get('item_key') or '').strip()

    if not (ngay and doi and phase_key and item_key):
        return jsonify({'error': 'Thiếu date/doi/phase_key/item_key'}), 400

    if session.get('role') != 'admin' and doi != session.get('doi'):
        return jsonify({'error': 'Bạn chỉ upload được cho đội của mình'}), 403

    if request.content_length and request.content_length > CHECKLIST_PHOTO_MAX_BYTES + 1024:
        return jsonify({'error': 'File quá lớn'}), 413

    photo = request.files.get('photo')
    ext, err = _validate_photo(photo)
    if err:
        return jsonify({'error': err}), 400

    photo.seek(0, 2)
    size = photo.tell()
    photo.seek(0)
    if size > CHECKLIST_PHOTO_MAX_BYTES:
        return jsonify({'error': f'File vượt giới hạn {CHECKLIST_PHOTO_MAX_BYTES} bytes'}), 413
    if size == 0:
        return jsonify({'error': 'File rỗng'}), 400

    _ensure_checklist_schema()

    try:
        with _checklist_write_connection() as conn:
            count = conn.execute(
                "SELECT COUNT(*) FROM checklist_photos "
                "WHERE ngay=? AND doi=? AND phase_key=? AND item_key=?",
                (ngay, doi, phase_key, item_key)
            ).fetchone()[0]
            if count >= CHECKLIST_PHOTO_MAX_PER_ITEM:
                return jsonify({
                    'error': f'Đạt giới hạn {CHECKLIST_PHOTO_MAX_PER_ITEM} ảnh/mục'
                }), 400

            stored_filename = f"{uuid.uuid4().hex}{ext}"
            month_dir = os.path.join(CHECKLIST_UPLOAD_DIR, ngay[:7])
            os.makedirs(month_dir, exist_ok=True)
            photo.save(os.path.join(month_dir, stored_filename))

            original_secure = secure_filename(photo.filename) or f'photo{ext}'
            now = datetime.now().isoformat(timespec='seconds')

            cur = conn.execute(
                "INSERT INTO checklist_photos "
                "(ngay, doi, phase_key, item_key, filename_stored, filename_original, "
                "mime_type, size_bytes, nguoi_nhap, thoi_diem_nhap) "
                "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                (ngay, doi, phase_key, item_key, stored_filename, original_secure,
                 photo.mimetype, size, session.get('username'), now)
            )
            photo_id = cur.lastrowid
    except sqlite3.OperationalError as exc:
        if 'locked' in str(exc).lower() or 'busy' in str(exc).lower():
            return jsonify({'error': 'db_locked'}), 503
        raise

    _log_audit(ngay, doi, phase_key, item_key, 'upload_photo',
               {'photo_id': photo_id, 'filename': stored_filename, 'size': size},
               session.get('username'))

    return jsonify({
        'ok': 1,
        'photo': {
            'id': photo_id,
            'url': f"/checklist/photo/{stored_filename}",
            'thumbnail_url': f"/checklist/photo/{stored_filename}",
            'filename_original': original_secure,
            'size_bytes': size,
        }
    })
```

- [ ] **Step 4: Run test pass**

Run: `python3 -m pytest tests/test_checklist_photos.py -v`
Expected: 5 PASS.

- [ ] **Step 5: Commit**

```bash
git add blueprints/checklist_routes.py tests/test_checklist_photos.py
git commit -m "checklist: endpoint upload-photo với validation chain"
```

---

## Task 11: Photo delete + serve endpoints

**Files:**
- Modify: `blueprints/checklist_routes.py`.
- Test: `tests/test_checklist_photos.py` (append).

**Interfaces:**
- Produces: endpoint `checklist.api_checklist_delete_photo`, `checklist.serve_photo`.

- [ ] **Step 1: Viết test fail**

Append vào `tests/test_checklist_photos.py`:

```python
def _upload_one_photo(client):
    response = client.post('/api/checklist/upload-photo', data={
        'date': '2026-07-22', 'doi': 'Đội 1',
        'phase_key': 'sang_giao_viec', 'item_key': 'atvsld_trang_phuc_ccdc',
        'photo': (io.BytesIO(b'\xFF\xD8\xFF\xE0' + b'\0' * 100), 'test.jpg'),
        'csrf_token': 'test-csrf',
    }, content_type='multipart/form-data')
    return response.get_json()['photo']


def test_delete_photo_success(tmp_path, monkeypatch):
    db_path, upload_dir = _prepare_db(tmp_path, monkeypatch)
    client = _logged_in_to_truong_client()

    photo = _upload_one_photo(client)
    photo_id = photo['id']
    filename = photo['url'].split('/')[-1]

    response = client.post('/api/checklist/delete-photo', json={'photo_id': photo_id})

    assert response.status_code == 200
    assert response.get_json()['ok'] == 1

    import sqlite3 as s3
    conn = s3.connect(db_path)
    count = conn.execute(
        "SELECT COUNT(*) FROM checklist_photos WHERE id=?", (photo_id,)
    ).fetchone()[0]
    conn.close()
    assert count == 0
    assert not (upload_dir / '2026-07' / filename).exists()


def test_delete_photo_rejects_not_owner(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    client = _logged_in_to_truong_client(doi='Đội 1')
    photo = _upload_one_photo(client)

    with client.session_transaction() as sess:
        sess['doi'] = 'Đội 99'

    response = client.post('/api/checklist/delete-photo', json={'photo_id': photo['id']})
    assert response.status_code == 403


def test_delete_photo_404_when_not_found(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    client = _logged_in_to_truong_client()

    response = client.post('/api/checklist/delete-photo', json={'photo_id': 99999})
    assert response.status_code == 404


def test_serve_photo_returns_image(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    client = _logged_in_to_truong_client()
    photo = _upload_one_photo(client)

    response = client.get(photo['url'])

    assert response.status_code == 200
    assert response.content_type.startswith('image/')


def test_serve_photo_rejects_bad_filename(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    client = _logged_in_to_truong_client()

    response = client.get('/checklist/photo/../../../etc/passwd')
    assert response.status_code in (400, 404)


def test_serve_photo_404_when_missing(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    client = _logged_in_to_truong_client()

    response = client.get('/checklist/photo/abcdef0123456789abcdef0123456789.jpg')
    assert response.status_code == 404
```

- [ ] **Step 2: Run test fail**

Run: `python3 -m pytest tests/test_checklist_photos.py -k "delete or serve" -v`
Expected: FAIL 404.

- [ ] **Step 3: Implement**

Trong `blueprints/checklist_routes.py`:

3a. Sửa imports (thêm `abort`, `current_app`, `send_file`):

```python
from flask import (
    Blueprint, abort, current_app, jsonify, render_template,
    request, send_file, session,
)
```

3b. Endpoint delete:

```python
@checklist_bp.route('/api/checklist/delete-photo', methods=['POST'])
@login_required
@to_truong_or_admin_required
def api_checklist_delete_photo():
    data = request.get_json(silent=True) or {}
    photo_id = data.get('photo_id')
    try:
        photo_id = int(photo_id)
    except (TypeError, ValueError):
        return jsonify({'error': 'photo_id không hợp lệ'}), 400

    _ensure_checklist_schema()
    try:
        with _checklist_write_connection() as conn:
            row = conn.execute(
                "SELECT * FROM checklist_photos WHERE id=?", (photo_id,)
            ).fetchone()
            if not row:
                return jsonify({'error': 'Không tìm thấy ảnh'}), 404

            if session.get('role') != 'admin' and row['doi'] != session.get('doi'):
                return jsonify({'error': 'Bạn chỉ xóa được ảnh đội mình'}), 403

            month = row['ngay'][:7]
            file_path = os.path.join(CHECKLIST_UPLOAD_DIR, month, row['filename_stored'])
            try:
                os.remove(file_path)
            except FileNotFoundError:
                pass
            except OSError as exc:
                current_app.logger.warning(f"Cannot remove {file_path}: {exc}")

            conn.execute("DELETE FROM checklist_photos WHERE id=?", (photo_id,))
    except sqlite3.OperationalError as exc:
        if 'locked' in str(exc).lower() or 'busy' in str(exc).lower():
            return jsonify({'error': 'db_locked'}), 503
        raise

    _log_audit(row['ngay'], row['doi'], row['phase_key'], row['item_key'],
               'delete_photo', {'photo_id': photo_id, 'filename': row['filename_stored']},
               session.get('username'))

    return jsonify({'ok': 1})
```

3c. Endpoint serve + regex guard:

```python
_PHOTO_FILENAME_RE = re.compile(r'^[a-f0-9]{32}\.(jpg|jpeg|png|webp)$')


@checklist_bp.route('/checklist/photo/<path:filename>')
@login_required
def serve_photo(filename):
    if not _PHOTO_FILENAME_RE.match(filename):
        abort(400)

    _ensure_checklist_schema()
    try:
        with _checklist_read_connection() as conn:
            row = conn.execute(
                "SELECT ngay, mime_type FROM checklist_photos WHERE filename_stored=?",
                (filename,)
            ).fetchone()
    except sqlite3.OperationalError:
        abort(500)

    if not row:
        abort(404)

    month = row['ngay'][:7]
    file_path = os.path.join(CHECKLIST_UPLOAD_DIR, month, filename)
    if not os.path.exists(file_path):
        abort(404)

    response = send_file(file_path, mimetype=row['mime_type'] or 'image/jpeg')
    response.headers['Cache-Control'] = 'public, max-age=86400'
    return response
```

- [ ] **Step 4: Run test pass**

Run: `python3 -m pytest tests/test_checklist_photos.py -v`
Expected: 11 PASS.

- [ ] **Step 5: Commit**

```bash
git add blueprints/checklist_routes.py tests/test_checklist_photos.py
git commit -m "checklist: endpoint delete-photo + serve-photo với path traversal guard"
```

---

## Task 12: Metric fetchers + registry

**Files:**
- Modify: `blueprints/checklist_routes.py`.
- Test: `tests/test_checklist_metric_fetchers.py` (create new).

**Interfaces:**
- Produces: `METRIC_FETCHERS` dict, `_safe_fetch`, 5 fetcher functions.

- [ ] **Step 1: Viết test fail**

Tạo `tests/test_checklist_metric_fetchers.py`:

```python
import sqlite3
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from blueprints import checklist_routes


def _seed_brcd_kiemsoat_db(tmp_path, rows_brcd, rows_kiemsoat=None):
    """Tạo brcd_kiemsoat.db mẫu với brcd_phieu + brcd_kiemsoat rows."""
    db_path = tmp_path / 'brcd_kiemsoat.db'
    conn = sqlite3.connect(db_path)
    conn.executescript('''
        CREATE TABLE brcd_phieu (
            baohong_id INTEGER PRIMARY KEY,
            doi_vt TEXT, nvkt TEXT, trang_thai_cong TEXT,
            gio_con_lai_thuc REAL, sa TEXT, first_seen TEXT, last_seen TEXT
        );
        CREATE TABLE brcd_kiemsoat (
            baohong_id INTEGER PRIMARY KEY,
            doi_vt TEXT, nvkt TEXT, noi_dung_kiem_soat TEXT,
            nguoi_nhap TEXT
        );
    ''')
    conn.executemany(
        "INSERT INTO brcd_phieu (baohong_id, doi_vt, nvkt, trang_thai_cong, "
        "gio_con_lai_thuc, sa, first_seen, last_seen) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
        rows_brcd
    )
    if rows_kiemsoat:
        conn.executemany(
            "INSERT INTO brcd_kiemsoat (baohong_id, doi_vt, nvkt, noi_dung_kiem_soat, nguoi_nhap) "
            "VALUES (?, ?, ?, ?, ?)",
            rows_kiemsoat
        )
    conn.commit()
    conn.close()
    return db_path


def test_fetch_brcd_pttb_overdue(tmp_path, monkeypatch):
    db_path = _seed_brcd_kiemsoat_db(tmp_path, [
        (1, 'Đội 1', 'nv1', 'Đang xử lý', -2.0, 'SC', '2026-07-22', '2026-07-22'),
        (2, 'Đội 1', 'nv2', 'Đang xử lý', 5.0, 'SC', '2026-07-22', '2026-07-22'),
        (3, 'Đội 2', 'nv3', 'Đang xử lý', -3.0, 'SC', '2026-07-22', '2026-07-22'),
    ])
    monkeypatch.setattr('config.BRCD_KIEMSOAT_DB_PATH', str(db_path))

    result = checklist_routes._fetch_brcd_pttb_overdue('Đội 1', '2026-07-22')

    assert result is not None
    assert result['sc'] == 1
    assert 'source_url' in result


def test_fetch_kiemsoat_completion(tmp_path, monkeypatch):
    db_path = _seed_brcd_kiemsoat_db(
        tmp_path,
        rows_brcd=[
            (1, 'Đội 1', 'nv1', 'Đang xử lý', 1.0, 'SC', '2026-07-22', '2026-07-22'),
            (2, 'Đội 1', 'nv2', 'Đang xử lý', 1.0, 'SC', '2026-07-22', '2026-07-22'),
        ],
        rows_kiemsoat=[
            (1, 'Đội 1', 'nv1', 'Đã liên hệ', 'tt1'),
        ]
    )
    monkeypatch.setattr('config.BRCD_KIEMSOAT_DB_PATH', str(db_path))

    result = checklist_routes._fetch_kiemsoat_completion('Đội 1', '2026-07-22')

    assert result is not None
    assert result['total_phieu'] == 2
    assert result['da_ks'] == 1
    assert 0 <= result['rate'] <= 1


def test_safe_fetch_returns_none_on_missing_db(tmp_path, monkeypatch):
    monkeypatch.setattr('config.BRCD_KIEMSOAT_DB_PATH', str(tmp_path / 'nonexistent.db'))

    result = checklist_routes._safe_fetch('brcd_pttb_overdue', 'Đội 1', '2026-07-22')

    assert result is None


def test_safe_fetch_returns_none_on_unknown_source():
    result = checklist_routes._safe_fetch('unknown_source', 'Đội 1', '2026-07-22')
    assert result is None


def test_metric_fetchers_registry_has_5_entries():
    assert len(checklist_routes.METRIC_FETCHERS) == 5
    expected_keys = {
        'inventory', 'brcd_pttb_overdue', 'brcd_pttb_priority',
        'brcd_nvkt_low_productivity', 'kiemsoat_completion',
    }
    assert set(checklist_routes.METRIC_FETCHERS.keys()) == expected_keys
```

- [ ] **Step 2: Run test fail**

Run: `python3 -m pytest tests/test_checklist_metric_fetchers.py -v`
Expected: FAIL `AttributeError`.

- [ ] **Step 3: Implement**

Thêm vào `blueprints/checklist_routes.py` (cuối file):

```python
def _brcd_kiemsoat_read_connection():
    from config import BRCD_KIEMSOAT_DB_PATH
    return sqlite3.connect(
        f'file:{BRCD_KIEMSOAT_DB_PATH}?mode=ro', uri=True, timeout=2
    )


def _fetch_brcd_pttb_overdue(doi, ngay):
    """Đếm SC quá hạn (gio_con_lai_thuc < 0) cho đội."""
    from config import BRCD_KIEMSOAT_DB_PATH
    if not os.path.exists(BRCD_KIEMSOAT_DB_PATH):
        return None
    with _brcd_kiemsoat_read_connection() as conn:
        sc = conn.execute(
            "SELECT COUNT(*) FROM brcd_phieu WHERE doi_vt=? AND gio_con_lai_thuc < 0",
            (doi,)
        ).fetchone()[0]
    return {
        'label': 'SC/PT quá hạn',
        'value': f'{sc} phiếu',
        'detail': {'sc': sc},
        'fetched_at': datetime.now().isoformat(timespec='seconds'),
        'source_url': '/brcd',
        'source_label': 'Xem /brcd',
    }


def _fetch_brcd_pttb_priority(doi, ngay):
    """Group brcd_phieu theo sa cho đội."""
    from config import BRCD_KIEMSOAT_DB_PATH
    if not os.path.exists(BRCD_KIEMSOAT_DB_PATH):
        return None
    with _brcd_kiemsoat_read_connection() as conn:
        rows = conn.execute(
            "SELECT sa, COUNT(*) FROM brcd_phieu WHERE doi_vt=? GROUP BY sa",
            (doi,)
        ).fetchall()
    breakdown = {r[0] or 'Khác': r[1] for r in rows}
    total = sum(breakdown.values())
    return {
        'label': 'Phiếu theo loại',
        'value': f'{total} phiếu',
        'detail': breakdown,
        'fetched_at': datetime.now().isoformat(timespec='seconds'),
        'source_url': '/brcd',
        'source_label': 'Xem /brcd',
    }


def _fetch_brcd_nvkt_low_productivity(doi, ngay, threshold=0.5):
    """Tìm NVKT có completion_rate < threshold trong ngày."""
    from config import BRCD_KIEMSOAT_DB_PATH
    if not os.path.exists(BRCD_KIEMSOAT_DB_PATH):
        return None
    with _brcd_kiemsoat_read_connection() as conn:
        rows = conn.execute(
            "SELECT nvkt, COUNT(*) as total, "
            "SUM(CASE WHEN trang_thai_cong IN ('Đã xong', 'Hoàn thành') THEN 1 ELSE 0 END) as done "
            "FROM brcd_phieu WHERE doi_vt=? AND DATE(first_seen)=? "
            "GROUP BY nvkt",
            (doi, ngay)
        ).fetchall()
    low_nvkt = []
    for r in rows:
        nvkt, total, done = r[0], r[1], r[2] or 0
        rate = done / total if total else 1.0
        if rate < threshold:
            low_nvkt.append({'nvkt': nvkt, 'total': total, 'done': done, 'rate': round(rate, 2)})
    return {
        'label': 'NV năng suất thấp',
        'value': f'{len(low_nvkt)} NV',
        'detail': {'low_nvkt': low_nvkt},
        'fetched_at': datetime.now().isoformat(timespec='seconds'),
        'source_url': '/brcd',
        'source_label': 'Xem /brcd',
    }


def _fetch_kiemsoat_completion(doi, ngay):
    """Tỷ lệ phiếu đã có annotation kiemsoat."""
    from config import BRCD_KIEMSOAT_DB_PATH
    if not os.path.exists(BRCD_KIEMSOAT_DB_PATH):
        return None
    with _brcd_kiemsoat_read_connection() as conn:
        total = conn.execute(
            "SELECT COUNT(*) FROM brcd_phieu WHERE doi_vt=?", (doi,)
        ).fetchone()[0]
        annotated = conn.execute(
            "SELECT COUNT(*) FROM brcd_phieu p "
            "JOIN brcd_kiemsoat k ON p.baohong_id = k.baohong_id "
            "WHERE p.doi_vt=? AND k.noi_dung_kiem_soat IS NOT NULL",
            (doi,)
        ).fetchone()[0]
    rate = annotated / total if total else 0.0
    return {
        'label': 'Tỷ lệ đã KS',
        'value': f'{annotated}/{total}',
        'detail': {'total_phieu': total, 'da_ks': annotated,
                   'chua_ks': total - annotated, 'rate': round(rate, 2)},
        'fetched_at': datetime.now().isoformat(timespec='seconds'),
        'source_url': '/brcd',
        'source_label': 'Xem /brcd',
    }


def _fetch_inventory(doi, ngay):
    """Placeholder MVP: chưa có data source cố định. Trả None."""
    current_app.logger.info('_fetch_inventory not yet implemented; returning None')
    return None


METRIC_FETCHERS = {
    'inventory': _fetch_inventory,
    'brcd_pttb_overdue': _fetch_brcd_pttb_overdue,
    'brcd_pttb_priority': _fetch_brcd_pttb_priority,
    'brcd_nvkt_low_productivity': _fetch_brcd_nvkt_low_productivity,
    'kiemsoat_completion': _fetch_kiemsoat_completion,
}


def _safe_fetch(source_key, doi, ngay):
    fn = METRIC_FETCHERS.get(source_key)
    if not fn:
        return None
    try:
        return fn(doi, ngay)
    except sqlite3.OperationalError as exc:
        msg = str(exc).lower()
        if 'locked' in msg or 'busy' in msg or 'no such table' in msg:
            current_app.logger.warning(f"checklist metric {source_key}: {exc}")
            return None
        current_app.logger.exception(f"checklist metric {source_key} failed")
        return None
    except Exception:
        current_app.logger.exception(f"checklist metric {source_key} failed")
        return None
```

- [ ] **Step 4: Run test pass**

Run: `python3 -m pytest tests/test_checklist_metric_fetchers.py -v`
Expected: 5 PASS.

- [ ] **Step 5: Commit**

```bash
git add blueprints/checklist_routes.py tests/test_checklist_metric_fetchers.py
git commit -m "checklist: 5 metric fetcher + registry + safe_fetch fail-safe"
```

---

## Task 13: Wire metrics vào api_checklist_data

**Files:**
- Modify: `blueprints/checklist_routes.py` (api_checklist_data thêm metric).
- Test: `tests/test_checklist_routes.py` (append).

- [ ] **Step 1: Viết test fail**

Append vào `tests/test_checklist_routes.py`:

```python
def test_data_includes_metric_for_items_with_source(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)

    monkeypatch.setattr('config.BRCD_KIEMSOAT_DB_PATH',
                        str(tmp_path / 'nonexistent.db'))

    client = _logged_in_to_truong_client()
    response = client.get('/api/checklist/data?date=2026-07-22&doi=Đội 1')
    body = response.get_json()

    items_with_metric_source = [
        item for phase in body['phases']
        for item in phase['items']
        if item.get('metric_source')
    ]
    assert len(items_with_metric_source) == 5
    for item in items_with_metric_source:
        assert 'metric' in item
```

- [ ] **Step 2: Run test fail**

Run: `python3 -m pytest tests/test_checklist_routes.py -k "test_data_includes_metric" -v`
Expected: FAIL (metric field null hoặc missing).

- [ ] **Step 3: Implement**

Trong `blueprints/checklist_routes.py`, sửa `api_checklist_data`. Tìm đoạn:

```python
    template = copy.deepcopy(CHECKLIST_TEMPLATE)
    for phase in template:
        for item in phase['items']:
            key = (phase['phase_key'], item['item_key'])
            item['state'] = state_map.get(key)
            item['photos'] = photos_map.get(key, [])
            item['metric'] = None
```

Sửa thành:

```python
    template = copy.deepcopy(CHECKLIST_TEMPLATE)
    for phase in template:
        for item in phase['items']:
            key = (phase['phase_key'], item['item_key'])
            item['state'] = state_map.get(key)
            item['photos'] = photos_map.get(key, [])
            if item.get('metric_source'):
                item['metric'] = _safe_fetch(item['metric_source'], doi, ngay)
            else:
                item['metric'] = None
```

- [ ] **Step 4: Run test pass**

Run: `python3 -m pytest tests/test_checklist_routes.py tests/test_checklist_metric_fetchers.py -v`
Expected: PASS tất cả.

- [ ] **Step 5: Commit**

```bash
git add blueprints/checklist_routes.py tests/test_checklist_routes.py
git commit -m "checklist: wire metric fetcher vào api_checklist_data"
```

---

## Task 14: Frontend — JS + CSS + CSRF meta

**Files:**
- Create: `static/css/checklist.css`, `static/js/pages/checklist.js`.
- Modify: `static/js/api.js` (thêm 5 methods), `templates/base.html` (CSRF meta tag).

**Note:** Task UI thuần. Verify bằng `python3 -m pytest tests/ -q` (không regression) + manual smoke.

- [ ] **Step 1: Thêm CSRF meta tag vào base.html**

Mở `templates/base.html`, tìm thẻ `<head>`. Thêm bên trong `<head>`:

```html
<meta name="csrf-token" content="{{ csrf_token() }}">
```

- [ ] **Step 2: Thêm API methods vào static/js/api.js**

Mở `static/js/api.js`, append cuối file (sau methods kiemsoat hiện có):

```javascript
window.CHECKLIST_API = {
    getData(date, doi) {
        const params = new URLSearchParams();
        if (date) params.set('date', date);
        if (doi) params.set('doi', doi);
        return fetch(`/api/checklist/data?${params}`).then(r => r.json());
    },

    saveState(date, doi, phaseKey, itemKey, trangThai, ghiChu) {
        return fetch('/api/checklist/luu', {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify({
                date, doi, phase_key: phaseKey, item_key: itemKey,
                trang_thai: trangThai, ghi_chu: ghiChu || '',
            }),
        }).then(r => r.json());
    },

    uploadPhoto(date, doi, phaseKey, itemKey, file, onProgress) {
        const formData = new FormData();
        const csrf = document.querySelector('meta[name="csrf-token"]');
        if (csrf) formData.append('csrf_token', csrf.content);
        formData.append('date', date);
        formData.append('doi', doi);
        formData.append('phase_key', phaseKey);
        formData.append('item_key', itemKey);
        formData.append('photo', file);

        return new Promise((resolve, reject) => {
            const xhr = new XMLHttpRequest();
            xhr.open('POST', '/api/checklist/upload-photo');
            if (onProgress && xhr.upload) {
                xhr.upload.onprogress = e => onProgress(e.loaded / e.total);
            }
            xhr.onload = () => {
                try { resolve(JSON.parse(xhr.responseText)); }
                catch (err) { reject(err); }
            };
            xhr.onerror = () => reject(new Error('Upload failed'));
            xhr.send(formData);
        });
    },

    deletePhoto(photoId) {
        return fetch('/api/checklist/delete-photo', {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify({photo_id: photoId}),
        }).then(r => r.json());
    },

    downloadReportUrl(from, to, doi) {
        const params = new URLSearchParams();
        if (from) params.set('from', from);
        if (to) params.set('to', to);
        if (doi) params.set('doi', doi);
        return `/download/checklist-report?${params}`;
    },
};
```

- [ ] **Step 3: Tạo static/css/checklist.css**

```css
.checklist-toolbar {
    display: flex; justify-content: space-between; align-items: center;
    margin-bottom: 1rem; padding: 0.75rem 1rem;
    background: #fff; border-radius: 8px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.1);
}
.checklist-filters { display: flex; gap: 1rem; align-items: center; }
.checklist-filters label {
    display: flex; align-items: center; gap: 0.5rem; font-weight: 500;
}
.checklist-progress { font-weight: bold; color: #2c5282; }

.btn-excel {
    background: #1e7e34; color: white; border: none;
    padding: 0.5rem 1rem; border-radius: 4px; cursor: pointer;
}
.btn-excel:hover { background: #155724; }

.checklist-phase {
    background: #fff; border-radius: 8px; margin-bottom: 1rem;
    box-shadow: 0 1px 3px rgba(0,0,0,0.1); overflow: hidden;
}
.checklist-phase-header {
    display: flex; justify-content: space-between; align-items: center;
    padding: 0.75rem 1rem; background: #f7fafc;
    border-bottom: 1px solid #e2e8f0; cursor: pointer;
}
.checklist-phase-title { font-weight: bold; color: #2d3748; }
.checklist-phase-time { color: #718096; font-size: 0.9em; }
.checklist-phase-progress { font-weight: 500; }

.checklist-item { border-bottom: 1px solid #edf2f7; }
.checklist-item-header {
    display: flex; align-items: center; gap: 0.75rem;
    padding: 0.75rem 1rem; cursor: pointer;
}
.checklist-item-header:hover { background: #f7fafc; }

.status-pill {
    display: inline-block; padding: 0.25rem 0.75rem;
    border-radius: 12px; font-size: 0.85em; font-weight: 500;
    min-width: 80px; text-align: center;
}
.status-chua-lam { background: #e2e8f0; color: #4a5568; }
.status-dang-lam { background: #bee3f8; color: #2c5282; }
.status-xong { background: #c6f6d5; color: #22543d; }
.status-bo-qua { background: #feebc8; color: #744210; }

.status-select {
    border: 1px solid #cbd5e0; border-radius: 12px;
    padding: 0.25rem 0.5rem; font-size: 0.85em;
}

.checklist-item-name { flex: 1; color: #2d3748; }

.metric-badge {
    background: #ebf8ff; color: #2c5282;
    padding: 0.15rem 0.5rem; border-radius: 4px;
    font-size: 0.85em; text-decoration: none;
}
.metric-badge:hover { background: #bee3f8; }

.checklist-item-meta { color: #718096; font-size: 0.85em; }

.btn-expand {
    background: transparent; border: none; color: #4a5568;
    cursor: pointer; font-size: 0.8em; transition: transform 0.2s;
}
.btn-expand.expanded { transform: rotate(180deg); }

.checklist-item-expanded {
    padding: 1rem; background: #fafbfc; display: none;
}
.checklist-item-expanded.visible { display: block; }

.checklist-notes {
    width: 100%; min-height: 80px; padding: 0.5rem;
    border: 1px solid #cbd5e0; border-radius: 4px;
    resize: vertical; font-family: inherit;
}
.checklist-notes-save-bar {
    display: flex; justify-content: flex-end;
    margin-top: 0.5rem; gap: 0.5rem;
}

.btn-save {
    background: #2c5282; color: white; border: none;
    padding: 0.4rem 1rem; border-radius: 4px; cursor: pointer;
}
.btn-save:disabled { opacity: 0.5; cursor: default; }

.checklist-photos {
    margin-top: 1rem; display: flex; flex-wrap: wrap; gap: 0.5rem;
}

.photo-thumb {
    position: relative; width: 80px; height: 80px;
    border: 1px solid #cbd5e0; border-radius: 4px; overflow: hidden;
}
.photo-thumb img {
    width: 100%; height: 100%; object-fit: cover; cursor: pointer;
}
.photo-thumb .photo-delete {
    position: absolute; top: 2px; right: 2px;
    background: rgba(0,0,0,0.6); color: white; border: none;
    border-radius: 50%; width: 20px; height: 20px;
    cursor: pointer; font-size: 0.7em;
    display: flex; align-items: center; justify-content: center;
}

.photo-upload-zone {
    width: 80px; height: 80px;
    border: 2px dashed #cbd5e0; border-radius: 4px;
    display: flex; align-items: center; justify-content: center;
    cursor: pointer; color: #718096;
    flex-direction: column; font-size: 0.7em; text-align: center;
}
.photo-upload-zone.drag-over { border-color: #2c5282; background: #ebf8ff; }

.checklist-viewonly-banner {
    background: #feebc8; color: #744210;
    padding: 0.5rem 1rem; border-radius: 4px;
    margin-bottom: 1rem; font-size: 0.9em;
}
.checklist-loading { color: #718096; padding: 1rem; }

.checklist-toast {
    position: fixed; bottom: 1rem; right: 1rem;
    padding: 0.75rem 1rem; border-radius: 4px;
    color: white; z-index: 9999; transition: opacity 0.3s;
}
.checklist-toast.success { background: #1e7e34; }
.checklist-toast.error { background: #c82333; }
```

- [ ] **Step 4: Tạo static/js/pages/checklist.js**

```javascript
(function() {
    'use strict';

    let state = { date: null, doi: null, data: null };

    function init() {
        const params = new URLSearchParams(window.location.search);
        state.date = params.get('date') || new Date().toISOString().slice(0, 10);
        state.doi = params.get('doi') || null;

        document.getElementById('checklist-date-input').value = state.date;
        document.getElementById('checklist-prev-day').addEventListener('click', () => shiftDay(-1));
        document.getElementById('checklist-next-day').addEventListener('click', () => shiftDay(1));
        document.getElementById('checklist-today').addEventListener('click', goToToday);
        document.getElementById('checklist-date-input').addEventListener('change', e => {
            state.date = e.target.value; updateUrl(); loadData();
        });
        document.getElementById('checklist-doi-select').addEventListener('change', e => {
            state.doi = e.target.value; updateUrl(); loadData();
        });
        document.getElementById('checklist-export-btn').addEventListener('click', exportExcel);

        loadData();
    }

    async function loadData() {
        const container = document.getElementById('checklist-phases-container');
        container.innerHTML = '<p class="checklist-loading">Đang tải...</p>';
        try {
            const data = await window.CHECKLIST_API.getData(state.date, state.doi);
            state.data = data;
            state.date = data.date;
            state.doi = data.doi;
            document.getElementById('checklist-date-input').value = state.date;
            populateDoiSelect(data.available_dois, state.doi);
            renderPhases(data);
        } catch (err) {
            container.innerHTML = `<p class="checklist-loading" style="color:#c82333">Lỗi: ${err.message}</p>`;
        }
    }

    function populateDoiSelect(availableDois, currentDoi) {
        const select = document.getElementById('checklist-doi-select');
        const options = availableDois && availableDois.length ? availableDois : [currentDoi];
        const all = [...new Set([...options, currentDoi].filter(Boolean))];
        select.innerHTML = all.map(d =>
            `<option value="${d}" ${d === currentDoi ? 'selected' : ''}>${d}</option>`
        ).join('');
    }

    function renderPhases(data) {
        const container = document.getElementById('checklist-phases-container');
        let html = '';

        if (!data.can_edit) {
            html += '<div class="checklist-viewonly-banner">Chỉ xem — bạn không phải tổ trưởng đội này.</div>';
        }

        const totalItems = data.phases.reduce((s, p) => s + p.items.length, 0);
        const doneItems = data.phases.reduce((s, p) =>
            s + p.items.filter(i => i.state && i.state.trang_thai === 'xong').length, 0);
        document.getElementById('checklist-progress-text').textContent =
            `${doneItems}/${totalItems} (${Math.round(doneItems / totalItems * 100)}%)`;

        data.phases.forEach((phase, idx) => {
            const done = phase.items.filter(i => i.state && i.state.trang_thai === 'xong').length;
            html += `
                <div class="checklist-phase">
                    <div class="checklist-phase-header">
                        <span class="checklist-phase-title">${phase.phase_name}</span>
                        <span class="checklist-phase-time">⏰ ${phase.time_window || ''}</span>
                        <span class="checklist-phase-progress">${done}/${phase.items.length}</span>
                    </div>
                    <div class="checklist-items-container">
                        ${phase.items.map(item => renderItem(item, phase, data.can_edit)).join('')}
                    </div>
                </div>`;
        });
        container.innerHTML = html;
        bindItemEvents();
    }

    function renderItem(item, phase, canEdit) {
        const status = item.state ? item.state.trang_thai : 'chua_lam';
        const note = item.state ? (item.state.ghi_chu || '') : '';
        const photos = item.photos || [];

        const photosHtml = photos.map(p => `
            <div class="photo-thumb" data-photo-id="${p.id}">
                <img src="${p.url}" alt="${p.filename_original || ''}" onclick="window.open('${p.url}', '_blank')">
                ${canEdit ? `<button class="photo-delete" onclick="checklistDeletePhoto(${p.id}, this)">✕</button>` : ''}
            </div>
        `).join('');

        const statusPill = canEdit
            ? `<select class="status-select" data-phase-key="${phase.phase_key}" data-item-key="${item.item_key}">
                <option value="chua_lam" ${status === 'chua_lam' ? 'selected' : ''}>Chưa làm</option>
                <option value="dang_lam" ${status === 'dang_lam' ? 'selected' : ''}>Đang làm</option>
                <option value="xong" ${status === 'xong' ? 'selected' : ''}>Xong</option>
                <option value="bo_qua" ${status === 'bo_qua' ? 'selected' : ''}>Bỏ qua</option>
              </select>`
            : `<span class="status-pill status-${status}">${statusLabel(status)}</span>`;

        const metricBadge = item.metric
            ? `<a class="metric-badge" href="${item.metric.source_url}" target="_blank">${item.metric.label}: ${item.metric.value}</a>`
            : '';

        const photoCount = photos.length > 0 ? `· 📷 ${photos.length}` : '';
        const noteCount = note ? `· 📝 ${note.split('\n').length} dòng` : '';

        return `
            <div class="checklist-item" data-phase-key="${phase.phase_key}" data-item-key="${item.item_key}">
                <div class="checklist-item-header">
                    ${statusPill}
                    <span class="checklist-item-name">${item.name}</span>
                    ${metricBadge}
                    <span class="checklist-item-meta">${noteCount} ${photoCount}</span>
                    <button class="btn-expand" type="button">▼</button>
                </div>
                <div class="checklist-item-expanded">
                    <textarea class="checklist-notes" ${canEdit ? '' : 'disabled'}>${escapeHtml(note)}</textarea>
                    <div class="checklist-notes-save-bar">
                        <button class="btn-save" type="button" ${canEdit ? '' : 'disabled'}>💾 Lưu</button>
                    </div>
                    <div class="checklist-photos">
                        ${photosHtml}
                        ${canEdit ? `
                            <label class="photo-upload-zone">
                                + Thêm ảnh<br>(kéo-thả)
                                <input type="file" accept="image/jpeg,image/png,image/webp" hidden>
                            </label>` : ''}
                    </div>
                    ${item.state && item.state.thoi_diem_cap_nhat
                        ? `<div class="checklist-item-meta" style="margin-top:0.5rem">Cập nhật ${item.state.thoi_diem_cap_nhat} bởi ${item.state.nguoi_nhap || ''}</div>`
                        : ''}
                </div>
            </div>`;
    }

    function statusLabel(s) {
        return {'chua_lam': 'Chưa làm', 'dang_lam': 'Đang làm', 'xong': 'Xong', 'bo_qua': 'Bỏ qua'}[s] || s;
    }

    function escapeHtml(s) {
        const d = document.createElement('div');
        d.textContent = s || '';
        return d.innerHTML;
    }

    function bindItemEvents() {
        document.querySelectorAll('.checklist-item-header').forEach(header => {
            header.addEventListener('click', e => {
                if (e.target.tagName === 'SELECT' || e.target.tagName === 'A' || e.target.tagName === 'BUTTON') return;
                const item = header.parentElement;
                const expanded = item.querySelector('.checklist-item-expanded');
                const btn = header.querySelector('.btn-expand');
                expanded.classList.toggle('visible');
                btn.classList.toggle('expanded');
            });
        });

        document.querySelectorAll('.status-select').forEach(sel => {
            sel.addEventListener('change', e => {
                const item = e.target.closest('.checklist-item');
                const note = item.querySelector('.checklist-notes').value;
                saveItemState(e.target.dataset.phaseKey, e.target.dataset.itemKey, e.target.value, note);
            });
        });

        document.querySelectorAll('.btn-save').forEach(btn => {
            btn.addEventListener('click', e => {
                const item = e.target.closest('.checklist-item');
                const status = item.querySelector('.status-select')?.value || 'chua_lam';
                const note = item.querySelector('.checklist-notes').value;
                saveItemState(item.dataset.phaseKey, item.dataset.itemKey, status, note);
            });
        });

        document.querySelectorAll('.checklist-notes').forEach(ta => {
            ta.addEventListener('blur', e => {
                const item = e.target.closest('.checklist-item');
                const status = item.querySelector('.status-select')?.value || 'chua_lam';
                saveItemState(item.dataset.phaseKey, item.dataset.itemKey, status, e.target.value);
            });
        });

        document.querySelectorAll('.photo-upload-zone input[type=file]').forEach(input => {
            input.addEventListener('change', e => {
                if (e.target.files[0]) handleUpload(e.target.closest('.checklist-item'), e.target.files[0]);
            });
        });

        document.querySelectorAll('.photo-upload-zone').forEach(zone => {
            zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('drag-over'); });
            zone.addEventListener('dragleave', () => zone.classList.remove('drag-over'));
            zone.addEventListener('drop', e => {
                e.preventDefault();
                zone.classList.remove('drag-over');
                const file = e.dataTransfer.files[0];
                if (file) handleUpload(zone.closest('.checklist-item'), file);
            });
        });
    }

    async function saveItemState(phaseKey, itemKey, trangThai, ghiChu) {
        try {
            const r = await window.CHECKLIST_API.saveState(
                state.date, state.doi, phaseKey, itemKey, trangThai, ghiChu
            );
            if (r.ok) showToast('Đã lưu ✓', 'success');
            else showToast(`Lỗi: ${r.error}`, 'error');
        } catch (err) {
            showToast(`Lỗi: ${err.message}`, 'error');
        }
    }

    async function handleUpload(item, file) {
        const phaseKey = item.dataset.phaseKey;
        const itemKey = item.dataset.itemKey;
        try {
            const r = await window.CHECKLIST_API.uploadPhoto(
                state.date, state.doi, phaseKey, itemKey, file
            );
            if (r.ok) { showToast('Đã upload ✓', 'success'); loadData(); }
            else showToast(`Lỗi: ${r.error}`, 'error');
        } catch (err) {
            showToast(`Lỗi upload: ${err.message}`, 'error');
        }
    }

    window.checklistDeletePhoto = async function(photoId, btn) {
        if (!confirm('Xóa ảnh này?')) return;
        try {
            const r = await window.CHECKLIST_API.deletePhoto(photoId);
            if (r.ok) { showToast('Đã xóa ✓', 'success'); loadData(); }
            else showToast(`Lỗi: ${r.error}`, 'error');
        } catch (err) {
            showToast(`Lỗi: ${err.message}`, 'error');
        }
    };

    function showToast(msg, kind) {
        const existing = document.querySelector('.checklist-toast');
        if (existing) existing.remove();
        const t = document.createElement('div');
        t.className = `checklist-toast ${kind}`;
        t.textContent = msg;
        document.body.appendChild(t);
        setTimeout(() => t.style.opacity = '0', 2500);
        setTimeout(() => t.remove(), 3000);
    }

    function shiftDay(delta) {
        const d = new Date(state.date);
        d.setDate(d.getDate() + delta);
        state.date = d.toISOString().slice(0, 10);
        document.getElementById('checklist-date-input').value = state.date;
        updateUrl(); loadData();
    }

    function goToToday() {
        state.date = new Date().toISOString().slice(0, 10);
        document.getElementById('checklist-date-input').value = state.date;
        updateUrl(); loadData();
    }

    function updateUrl() {
        const params = new URLSearchParams();
        if (state.date) params.set('date', state.date);
        if (state.doi) params.set('doi', state.doi);
        window.history.replaceState({}, '', `${window.location.pathname}?${params}`);
    }

    function exportExcel() {
        const from = prompt('Từ ngày (YYYY-MM-DD):', state.date);
        if (!from) return;
        const to = prompt('Đến ngày (YYYY-MM-DD):', state.date);
        if (!to) return;
        window.location.href = window.CHECKLIST_API.downloadReportUrl(from, to, state.doi);
    }

    document.addEventListener('DOMContentLoaded', init);
})();
```

- [ ] **Step 5: Run full test suite**

Run: `python3 -m pytest tests/ -q`
Expected: tất cả tests PASS (không regression).

- [ ] **Step 6: Smoke test thủ công**

Run: `python3 dashboard.py &` → mở `http://localhost:5011/checklist`. Verify (login to_truong thủ công):
- Trang load, 3 phases hiển thị
- Click pill → dropdown, đổi trạng thái → toast "Đã lưu ✓"
- Expand item → nhập ghi chú → blur → toast
- Upload 1 ảnh → thumbnail hiện
- Xóa ảnh → mất
- Đổi đội (admin) → view-only banner nếu đội khác
- Kill server: `pkill -f dashboard.py`

- [ ] **Step 7: Commit**

```bash
git add static/css/checklist.css static/js/pages/checklist.js static/js/api.js templates/base.html
git commit -m "checklist: frontend JS/CSS hoàn chỉnh + CSRF meta tag"
```

---

## Task 15: Excel export endpoint

**Files:**
- Modify: `blueprints/checklist_routes.py`.
- Test: `tests/test_checklist_routes.py` (append).

**Interfaces:**
- Produces: endpoint `checklist.download_checklist_report`.

- [ ] **Step 1: Viết test fail**

Append vào `tests/test_checklist_routes.py`:

```python
def test_download_report_returns_xlsx(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    client = _logged_in_to_truong_client()

    client.post('/api/checklist/luu', json={
        'date': '2026-07-22', 'doi': 'Đội 1',
        'phase_key': 'sang_giao_viec', 'item_key': 'atvsld_trang_phuc_ccdc',
        'trang_thai': 'xong', 'ghi_chu': 'Done',
    })

    response = client.get('/download/checklist-report?from=2026-07-22&to=2026-07-22&doi=Đội 1')

    assert response.status_code == 200
    assert 'spreadsheet' in response.content_type


def test_download_report_requires_login(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    from dashboard import app
    client = app.test_client()

    response = client.get('/download/checklist-report?from=2026-07-22&to=2026-07-22')

    assert response.status_code == 302
    assert '/login' in response.headers.get('Location', '')
```

- [ ] **Step 2: Run test fail**

Run: `python3 -m pytest tests/test_checklist_routes.py -k download -v`
Expected: FAIL 404.

- [ ] **Step 3: Implement**

Thêm vào `blueprints/checklist_routes.py` (đảm bảo imports đầu file có `io`, `pandas`, `send_file`):

```python
import io
import pandas as pd
from flask import send_file
```

(Sửa khối `from flask import` để thêm `send_file` nếu chưa có.)

Endpoint:

```python
@checklist_bp.route('/download/checklist-report')
@login_required
def download_checklist_report():
    from_param = (request.args.get('from') or '').strip()
    to_param = (request.args.get('to') or '').strip()
    doi_param = (request.args.get('doi') or '').strip()

    if not (from_param and to_param):
        return jsonify({'error': 'Thiếu from/to'}), 400

    _ensure_checklist_schema()

    state_rows = []
    photo_rows = []
    try:
        with _checklist_read_connection() as conn:
            state_sql = (
                "SELECT ngay, doi, phase_key, item_key, trang_thai, ghi_chu, "
                "nguoi_nhap, nguoi_nhap_name, thoi_diem_nhap, thoi_diem_cap_nhat "
                "FROM checklist_state WHERE ngay BETWEEN ? AND ?"
            )
            state_params = [from_param, to_param]
            if doi_param:
                state_sql += " AND doi=?"
                state_params.append(doi_param)
            state_rows = conn.execute(state_sql, state_params).fetchall()

            photo_sql = (
                "SELECT ngay, doi, phase_key, item_key, filename_stored, filename_original, "
                "size_bytes, nguoi_nhap, thoi_diem_nhap "
                "FROM checklist_photos WHERE ngay BETWEEN ? AND ?"
            )
            photo_params = [from_param, to_param]
            if doi_param:
                photo_sql += " AND doi=?"
                photo_params.append(doi_param)
            photo_rows = conn.execute(photo_sql, photo_params).fetchall()
    except sqlite3.OperationalError:
        pass

    state_cols = ['ngay', 'doi', 'phase_key', 'item_key', 'trang_thai', 'ghi_chu',
                  'nguoi_nhap', 'nguoi_nhap_name', 'thoi_diem_nhap', 'thoi_diem_cap_nhat']
    state_df = pd.DataFrame([dict(r) for r in state_rows]) if state_rows else pd.DataFrame(columns=state_cols)

    photo_cols = ['ngay', 'doi', 'phase_key', 'item_key', 'filename_stored',
                  'filename_original', 'size_bytes', 'nguoi_nhap', 'thoi_diem_nhap']
    photo_df = pd.DataFrame([dict(r) for r in photo_rows]) if photo_rows else pd.DataFrame(columns=photo_cols)
    if not photo_df.empty:
        photo_df['url'] = photo_df['filename_stored'].apply(lambda f: f"/checklist/photo/{f}")

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        state_df.to_excel(writer, sheet_name='Checklist State', index=False)
        photo_df.to_excel(writer, sheet_name='Photos', index=False)
    output.seek(0)

    download_name = f"checklist_{from_param}_{to_param}{('_' + doi_param) if doi_param else ''}.xlsx"
    return send_file(
        output,
        as_attachment=True,
        download_name=download_name,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )
```

- [ ] **Step 4: Run test pass**

Run: `python3 -m pytest tests/test_checklist_routes.py -v`
Expected: 8 PASS.

- [ ] **Step 5: Commit**

```bash
git add blueprints/checklist_routes.py tests/test_checklist_routes.py
git commit -m "checklist: endpoint download Excel report (2 sheet state + photos)"
```

---

## Task 16: Admin UI — admin_users.html thêm cột Đội + Role dropdown

**Files:**
- Modify: `templates/admin_users.html`.

**Note:** UI thuần. Verify thủ công.

- [ ] **Step 1: Thêm cột "Đội" vào header bảng**

Trong `templates/admin_users.html`, tìm khối `<thead>` (khoảng dòng 299-307). Thêm `<th>Đội</th>` ngay sau `<th>Role</th>`:

```html
<tr>
    <th>STT</th>
    <th>Username</th>
    <th>Họ tên</th>
    <th>Role</th>
    <th>Đội</th>
    <th>Trạng thái</th>
    <th>Lần đầu login</th>
    <th>Hành động</th>
</tr>
```

- [ ] **Step 2: Thêm cột Đội vào body + mở rộng role badge**

Tìm `<td>` render Role (khoảng dòng 315-321). Sửa để hỗ trợ `to_truong` và thêm ô Đội ngay sau:

```html
<td>
    {% if user.role == 'admin' %}
        <span class="badge badge-admin">ADMIN</span>
    {% elif user.role == 'to_truong' %}
        <span class="badge badge-to-truong">TỔ TRƯỞNG</span>
    {% else %}
        <span class="badge badge-user">USER</span>
    {% endif %}
</td>
<td>
    {% if user.doi %}{{ user.doi }}{% else %}—{% endif %}
</td>
```

- [ ] **Step 3: Thêm form "Sửa đội" vào cột Hành động**

Tìm `<div class="action-buttons">` (khoảng dòng 337). Thêm form mới NGAY SAU thẻ mở div (trước form Reset PW):

```html
<div class="action-buttons">
    <form method="POST" action="{{ url_for('auth.set_user_doi') }}" style="display:inline;">
        <input type="hidden" name="username" value="{{ user.username }}">
        <input type="hidden" name="csrf_token" value="{{ csrf_token() }}">
        <input type="text" name="doi" value="{{ user.doi or '' }}"
               placeholder="Đội..." style="width: 90px; padding: 4px;">
        <button type="submit" class="btn btn-secondary btn-sm">Sửa đội</button>
    </form>
    <!-- form Reset PW hiện có ở đây -->
```

- [ ] **Step 4: Thêm CSS cho badge to_truong (optional)**

Nếu file CSS có style cho `.badge-admin` / `.badge-user`, thêm `.badge-to-truong` (vd màu xanh dương) cùng khu vực. Skip nếu không quan tâm màu.

- [ ] **Step 5: Smoke test thủ công**

Run: `python3 dashboard.py &` → login admin → mở `/admin/users`. Verify:
- Bảng có thêm cột "Đội"
- Set "Đội 1" cho 1 user → reload → giá trị còn đó
- Đổi role sang `to_truong` → reload → role cập nhật
- Kill server.

- [ ] **Step 6: Commit**

```bash
git add templates/admin_users.html
git commit -m "admin: thêm cột Đội + dropdown Role 3 giá trị"
```

---

## Task 17: Cập nhật docs (00, 04, 08, AGENTS, docs/12 mới)

**Files:**
- Modify: `docs/00-doc-index.md`, `docs/04-mapping-route-va-du-lieu.md`, `docs/08-trang-thai-thuc-thi.md`, `AGENTS.md`.
- Create: `docs/12-checklist-to-truong.md`.

- [ ] **Step 1: Update docs/00-doc-index.md**

Thêm reference tới `docs/12-checklist-to-truong.md` trong section "Khi cần làm gì thì đọc gì" + mục "Kiểm soát tổ trưởng / checklist hằng ngày". Thêm entry mới dưới entry docs/11 hiện có.

- [ ] **Step 2: Update docs/04-mapping-route-va-du-lieu.md**

Thêm 6 dòng vào bảng mapping chính:

| Route | Endpoint | ma_bao_cao | ten_bang_du_lieu | supports_date | status |
|-------|----------|------------|------------------|---------------|--------|
| `/checklist` | `checklist.page_checklist` | — | checklist.db::checklist_state/photos/audit_log | n/a (writable) | migrated |
| `/api/checklist/data` | `checklist.api_checklist_data` | — | (sourced live) | n/a | migrated |
| `/api/checklist/luu` | `checklist.api_checklist_luu` | — | checklist.db::checklist_state + audit_log | n/a (writable) | migrated |
| `/api/checklist/upload-photo` | `checklist.api_checklist_upload_photo` | — | checklist.db::checklist_photos | n/a | migrated |
| `/api/checklist/delete-photo` | `checklist.api_checklist_delete_photo` | — | checklist.db::checklist_photos | n/a | migrated |
| `/download/checklist-report` | `checklist.download_checklist_report` | — | checklist.db (cross-table) | n/a | migrated |

- [ ] **Step 3: Update docs/08-trang-thai-thuc-thi.md**

Thêm section "9. KT tổ trưởng — Checklist hằng ngày" sau section 8:

```markdown
### 9. KT tổ trưởng — Checklist hằng ngày

Trang `/checklist` (blueprint `checklist_bp`) cho tổ trưởng cập nhật trạng thái + ghi chú
+ ảnh theo mẫu 3 giai đoạn × 15 mục cố định, per-đội. Tách biệt với 3 trang kiemsoat hiện
có (/brcd, /pttb, /shc-cts) — không ghi ngược vào brcd_kiemsoat.db.

**Endpoints (đã migrated, all live):**
- GET `/checklist` — page shell
- GET `/api/checklist/data?date=&doi=` — trả template + state + photos + metric live
- POST `/api/checklist/luu` — UPSERT state + audit log
- POST `/api/checklist/upload-photo` — multipart upload ảnh (CSRF)
- POST `/api/checklist/delete-photo` — xóa ảnh
- GET `/checklist/photo/<filename>` — serve ảnh
- GET `/download/checklist-report` — Excel export

**Data**: per-instance SQLite `<INSTANCE_RUNTIME_DIR>/checklist.db` (3 bảng: state,
photos, audit_log). Ảnh lưu `<INSTANCE_RUNTIME_DIR>/checklist_uploads/<YYYY-MM>/<uuid>.<ext>`.

**Auth**: role mới `to_truong` + cột mới `doi` trong username.xlsx (lazy migration).
Tổ trưởng xem tất cả đội, chỉ sửa đội mình.

**Metric**: 5 fetcher đọc read-only từ brcd_kiemsoat.db + inventory placeholder.
Fail-safe.

Xem `docs/12-checklist-to-truong.md` cho operational runbook.
```

- [ ] **Step 4: Update AGENTS.md**

Trong đoạn "Architecture", thêm dòng mô tả `checklist_routes.py`:

```markdown
- `blueprints/checklist_routes.py` — checklist hằng ngày cho tổ trưởng (3 phases × 15 mục
  hard-code, per-đội, photo upload hạ tầng mới, 5 metric fetcher). Endpoint prefix
  `checklist.`. DB riêng `<INSTANCE_RUNTIME_DIR>/checklist.db`.
```

Thêm 1 dòng trong "Critical import ordering":

```markdown
- Photo upload trong `checklist_routes.py` dùng `werkzeug.utils.secure_filename` + CSRF
  qua `app_helpers.csrf_protect` (multipart form). Path traversal guard bằng regex
  `^[a-f0-9]{32}\.(jpg|jpeg|png|webp)$`.
```

- [ ] **Step 5: Tạo docs/12-checklist-to-truong.md**

Tạo file mới ~250 dòng theo cấu trúc mirror `docs/11-kiemsoat-to-truong-van-hanh.md`. Outline các section:

1. **Tổng quan** — mục đích, khác gì với kiemsoat.
2. **Nguyên lý thiết kế** — hard-code template, write-aside, per-đội, lazy migration.
3. **CSDL** — đường dẫn DB + schema 3 bảng (tham chiếu spec cho DDL đầy đủ).
4. **Biến môi trường** — bảng 4 biến `DASHV4_CHECKLIST_*`.
5. **Nguồn metric** — 5 fetcher + tham chiếu SQL.
6. **Photo upload** — validation chain, storage structure, serving, multi-worker safety.
7. **API endpoints** — bảng 7 endpoints + payload/response.
8. **Backup & Restore** — `sqlite3 ... ".backup"` cho DB + `rsync -a` cho uploads dir.
9. **Troubleshooting** — ảnh không upload, metric lỗi, role sai, quá giới hạn.
10. **Cạm bẫy** — case-sensitivity `doi_vt`, lazy migration, max photo limit.
11. **Test** — `python3 -m pytest tests/test_checklist_*.py tests/test_auth_to_truong.py -v`.
12. **File tham chiếu nhanh** — bảng các file quan trọng + vai trò.
13. **Checklist bảo dưỡng định kỳ** — backup, disk size, test, smoke.

- [ ] **Step 6: Commit**

```bash
git add docs/00-doc-index.md docs/04-mapping-route-va-du-lieu.md docs/08-trang-thai-thuc-thi.md docs/12-checklist-to-truong.md AGENTS.md
git commit -m "docs: cập nhật doc-sync + tạo docs/12-checklist-to-truong.md"
```

---

## Hoàn thành

Sau Task 17, chạy final verification:

- [ ] `python3 -m py_compile blueprints/checklist_routes.py auth.py`
- [ ] `python3 -m pytest tests/ -q` (~93 tests, all green)
- [ ] `python3 dashboard.py &` → smoke `/checklist`, login to_truong, đổi trạng thái, upload ảnh, reload. `pkill -f dashboard.py`
- [ ] Kiểm tra git log: 17 commits liên tiếp về checklist.

Plan hoàn tất.
