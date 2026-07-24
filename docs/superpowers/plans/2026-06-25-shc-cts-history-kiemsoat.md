# SHC CTS — Lịch sử tiến độ theo ngày + Kiểm soát tổ trưởng Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Lưu lại snapshot "Tiến độ xử lý SHC trong ngày" (đọc từ Excel `Bao_cao_tien_trinh_*.xlsx`) vào DB nội bộ của dashboard theo từng ngày, cho phép xem lại lịch sử theo ngày và cho tổ trưởng nhập ghi chú kiểm soát theo từng NVKT — mô phỏng cơ chế của `/brcd` và `/pttb`.

**Architecture:** Thêm DB nội bộ per-instance `INSTANCE_RUNTIME_DIR/shc_cts.db` với 2 bảng: `shc_cts_tien_do` (snapshot lịch sử, grain = ngày × NVKT) và `shc_cts_kiemsoat` (ghi chú tổ trưởng, cùng grain). Một hàm sync idempotent `_sync_shc_cts_tien_do_to_db()` đọc file intraday mới nhất, tách ngày từ tên file (`Bao_cao_tien_trinh_YYYYMMDD.xlsx`) và upsert. Sync chạy qua cron + lazy trên endpoint detail. Bổ sung 4 endpoint mới trong `quality_routes.py` và 1 section UI có datepicker + cột "Kiểm soát" editable (giống `pttb.js`).

**Tech Stack:** Python 3.10, Flask 3.1, pandas 2.2, sqlite3 (stdlib), openpyxl, xlsxwriter, Chart.js, HTML/JS thuần (không framework FE).

## Global Constraints

- Chạy mọi lệnh từ repo root `/home/vtst/dashv4`. Cây module phẳng, import dạng top-level (`from config import ...`, `from blueprints ... import ...`).
- KHÔNG có lint/typecheck/formatter. Xác minh cú pháp Python bằng `python3 -m py_compile <file>`.
- Test dùng pytest HỆ THỐNG (`~/.local/bin/pytest`, 9.x): `python3 -m pytest tests/`. KHÔNG dùng `venv/bin/pytest`.
- KHÔNG thêm comment vào code trừ khi task yêu cầu.
- UI/comment dùng tiếng Việt.
- Bảng intraday "Theo NVKT" có đúng 9 cột (theo thứ tự): `Đơn vị`, `NVKT_DB`, `Tổng số`, `Đạt baseline`, `Đã xử lý trong ngày`, `Tổng đã đạt`, `Chưa đạt`, `OFF/Lỗi`, `% đạt`.
- File intraday đặt tên `Bao_cao_tien_trinh_YYYYMMDD.xlsx` (ngày ở tên file). Sheet mục tiêu tên `Theo NVKT`.
- Grain kiểm soát = `(ngay_xu_ly, nvkt_db)`. Grain snapshot = `(ngay_xu_ly, nvkt_db)`, upsert giữ 1 dòng cuối mỗi ngày.
- Các endpoint mới KHÔNG cần đăng ký trong `route_policy.py` (chỉ PAGE endpoint mới cần vào `PAGE_ACTIVE_KEYS`; các endpoint API kiemsoat của brcd/pttb cũng không đăng ký). Page `quality.page_shc_cts` đã có sẵn.
- Tuân thủ doc-sync: task cuối phải cập nhật `docs/08` và `docs/04`.
- Tiếng key DB dùng snake_case; tên cột hiển thị giữ tiếng Việt có dấu y hệt Excel.

## File Structure

- **Modify** `config.py` — thêm `SHC_CTS_HISTORY_DB_PATH`.
- **Modify** `blueprints/quality_routes.py` — thêm helpers DB, hàm sync, 4 endpoint mới.
- **Create** `scripts/sync_shc_cts_tien_do.py` — cron entry point.
- **Modify** `static/js/api.js` — thêm 4 method API cho SHC CTS kiemsoat.
- **Modify** `static/js/pages/shc_cts.js` — thêm logic section kiểm soát (datepicker, bảng editable, thongke).
- **Modify** `templates/pages/shc_cts.html` — thêm markup section kiểm soát + CSS.
- **Create** `tests/test_shc_cts_kiemsoat.py` — test backend (schema, sync, 4 endpoint).
- **Modify** `docs/08-trang-thai-thuc-thi.md` và `docs/04-mapping-route-va-du-lieu.md` — doc-sync.

---

## Task 1: Thêm đường dẫn DB lịch sử vào config

**Files:**
- Modify: `config.py` (ngay sau khối `BRCD_KIEMSOAT_DB_PATH` ở dòng ~117-120)

**Interfaces:**
- Produces: biến module `SHC_CTS_HISTORY_DB_PATH` (str) — import được từ `blueprints.quality_routes`.

- [ ] **Step 1: Thêm biến config**

Mở `config.py`, tìm đoạn:

```python
BRCD_KIEMSOAT_DB_PATH = _first_existing_path(
    os.getenv('DASHV4_BRCD_KIEMSOAT_DB_PATH'),
    os.path.join(INSTANCE_RUNTIME_DIR, 'brcd_kiemsoat.db'),
)
```

Chèn ngay sau đó:

```python
SHC_CTS_HISTORY_DB_PATH = _first_existing_path(
    os.getenv('DASHV4_SHC_CTS_HISTORY_DB_PATH'),
    os.path.join(INSTANCE_RUNTIME_DIR, 'shc_cts.db'),
)
```

- [ ] **Step 2: Xác minh import được**

Run: `python3 -c "from config import SHC_CTS_HISTORY_DB_PATH; print(SHC_CTS_HISTORY_DB_PATH)"`
Expected: in ra đường dẫn kết thúc bằng `shc_cts.db` (trong runtime_app/<unit>/).

- [ ] **Step 3: Commit**

```bash
git add config.py
git commit -m "config: thêm SHC_CTS_HISTORY_DB_PATH cho lịch sử tiến độ SHC CTS"
```

---

## Task 2: Schema + connection helpers trong quality_routes

**Files:**
- Modify: `blueprints/quality_routes.py` (imports đầu file + khu vực sau các hàm `_add_progress_timestamp_column`, dòng ~557)

**Interfaces:**
- Produces: `_ensure_shc_cts_schema()`, `_shc_cts_read_connection()`, `_shc_cts_write_connection()` (context manager trả `sqlite3.Connection` với `row_factory=sqlite3.Row`).
- Produces: biến module `SHC_CTS_HISTORY_DB_PATH` (sau khi import), `SHC_CTS_KIEMSOAT_NOI_DUNG_MAX = 2000`.

- [ ] **Step 1: Thêm imports**

Trong `blueprints/quality_routes.py`, sửa khối import đầu file. Tìm:

```python
import io
import os
import unicodedata
from collections import OrderedDict
from datetime import datetime

import pandas as pd
from flask import Blueprint, current_app, jsonify, render_template, request, send_file, session
```

Thay bằng:

```python
import io
import os
import sqlite3
import unicodedata
from collections import OrderedDict
from datetime import datetime
from threading import Lock

import pandas as pd
from flask import Blueprint, current_app, jsonify, render_template, request, send_file, session
```

Sau đó, trong câu `from config import ...` (dòng ~20), tìm:

```python
from config import BAOCAO_HANOI_DOWNLOADS_DIR, SHC_NVKT_DETAIL_REPORTS, SHC_NVKT_TEAMS
```

Thay bằng:

```python
from config import (
    BAOCAO_HANOI_DOWNLOADS_DIR,
    SHC_CTS_HISTORY_DB_PATH,
    SHC_NVKT_DETAIL_REPORTS,
    SHC_NVKT_TEAMS,
)
```

- [ ] **Step 2: Thêm constants + helpers**

Tìm hàm `_add_progress_timestamp_column` (khoảng dòng 551-557). Ngay SAU hàm đó (trước `def _shc_cts_payload_from_excel():`), chèn khối:

```python
SHC_CTS_KIEMSOAT_NOI_DUNG_MAX = 2000
SHC_CTS_INTRADAY_PROGRESS_COLUMNS = [
    'Đơn vị',
    'NVKT_DB',
    'Tổng số',
    'Đạt baseline',
    'Đã xử lý trong ngày',
    'Tổng đã đạt',
    'Chưa đạt',
    'OFF/Lỗi',
    '% đạt',
]

_shc_cts_schema_lock = Lock()
_shc_cts_schema_ready_path = None


def _shc_cts_write_connection():
    conn = sqlite3.connect(SHC_CTS_HISTORY_DB_PATH, timeout=5)
    conn.execute('PRAGMA journal_mode=WAL')
    conn.execute('PRAGMA busy_timeout=5000')
    conn.row_factory = sqlite3.Row
    return conn


def _shc_cts_read_connection():
    conn = sqlite3.connect(
        f'file:{os.path.abspath(SHC_CTS_HISTORY_DB_PATH)}?mode=ro',
        uri=True,
        timeout=5,
    )
    conn.row_factory = sqlite3.Row
    return conn


def _ensure_shc_cts_schema():
    global _shc_cts_schema_ready_path
    db_path = os.path.abspath(SHC_CTS_HISTORY_DB_PATH)
    with _shc_cts_schema_lock:
        if _shc_cts_schema_ready_path == db_path:
            return
        db_dir = os.path.dirname(db_path)
        if db_dir:
            os.makedirs(db_dir, exist_ok=True)
        with _shc_cts_write_connection() as conn:
            conn.execute(
                '''
                CREATE TABLE IF NOT EXISTS shc_cts_tien_do (
                    ngay_xu_ly        TEXT NOT NULL,
                    don_vi            TEXT,
                    nvkt_db           TEXT NOT NULL,
                    tong_so           INTEGER,
                    dat_baseline      INTEGER,
                    da_xu_ly_ngay     INTEGER,
                    tong_dat          INTEGER,
                    chua_dat          INTEGER,
                    off_loi           INTEGER,
                    ty_le_dat         REAL,
                    captured_at       TEXT,
                    PRIMARY KEY (ngay_xu_ly, nvkt_db)
                )
                '''
            )
            conn.execute(
                'CREATE INDEX IF NOT EXISTS idx_shc_cts_tien_do_ngay ON shc_cts_tien_do(ngay_xu_ly)'
            )
            conn.execute(
                'CREATE INDEX IF NOT EXISTS idx_shc_cts_tien_do_don_vi ON shc_cts_tien_do(don_vi)'
            )
            conn.execute(
                '''
                CREATE TABLE IF NOT EXISTS shc_cts_kiemsoat (
                    ngay_xu_ly         TEXT NOT NULL,
                    nvkt_db            TEXT NOT NULL,
                    don_vi             TEXT,
                    noi_dung_kiem_soat TEXT NOT NULL DEFAULT '',
                    nguoi_nhap         TEXT,
                    thoi_diem_nhap     TEXT,
                    thoi_diem_cap_nhat TEXT,
                    PRIMARY KEY (ngay_xu_ly, nvkt_db)
                )
                '''
            )
        _shc_cts_schema_ready_path = db_path
```

- [ ] **Step 3: Xác minh cú pháp**

Run: `python3 -m py_compile blueprints/quality_routes.py`
Expected: không có output (thành công).

- [ ] **Step 4: Commit**

```bash
git add blueprints/quality_routes.py
git commit -m "quality: thêm schema + connection helpers cho lịch sử SHC CTS"
```

---

## Task 3: Hàm parse ngày + sync snapshot vào DB

**Files:**
- Modify: `blueprints/quality_routes.py` (chèn sau khối helpers Task 2)
- Create: `tests/test_shc_cts_kiemsoat.py`

**Interfaces:**
- Produces: `_parse_shc_cts_intraday_date(filename) -> str | None` (trả `'YYYY-MM-DD'` từ `Bao_cao_tien_trinh_YYYYMMDD.xlsx`, `None` nếu không khớp).
- Produces: `_sync_shc_cts_tien_do_to_db() -> dict` với shape `{'synced': N, 'new': M, 'updated': K, 'skipped': 0|1, 'reason': str, 'ngay_xu_ly': str}`.

- [ ] **Step 1: Viết test (TDD)**

Tạo file `tests/test_shc_cts_kiemsoat.py` với nội dung:

```python
import sqlite3
import sys
import time
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from dashboard import app
from blueprints import quality_routes


_PROGRESS_COLUMNS = [
    'Đơn vị', 'NVKT_DB', 'Tổng số', 'Đạt baseline',
    'Đã xử lý trong ngày', 'Tổng đã đạt', 'Chưa đạt', 'OFF/Lỗi', '% đạt',
]


def _prepare_db(tmp_path, monkeypatch):
    db_path = tmp_path / 'shc_cts.db'
    monkeypatch.setattr(quality_routes, 'SHC_CTS_HISTORY_DB_PATH', str(db_path))
    quality_routes._shc_cts_schema_ready_path = None
    quality_routes._ensure_shc_cts_schema()
    return db_path


def _logged_in_client():
    client = app.test_client()
    with client.session_transaction() as sess:
        sess['username'] = 'test-user'
    return client


def _write_intraday(tmp_path, monkeypatch, rows, filename='Bao_cao_tien_trinh_20260625.xlsx'):
    intraday_path = tmp_path / filename
    df = pd.DataFrame(rows, columns=_PROGRESS_COLUMNS)
    with pd.ExcelWriter(intraday_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Theo NVKT', index=False)
    monkeypatch.setattr(quality_routes, 'SHC_CTS_INTRADAY_REPORT_DIR', str(tmp_path))
    return intraday_path


def _progress_rows():
    return [
        {
            'Đơn vị': 'Tổ Kỹ thuật Địa bàn Phúc Thọ', 'NVKT_DB': 'Nguyễn Văn A',
            'Tổng số': 10, 'Đạt baseline': 5, 'Đã xử lý trong ngày': 3,
            'Tổng đã đạt': 6, 'Chưa đạt': 4, 'OFF/Lỗi': 0, '% đạt': 60,
        },
        {
            'Đơn vị': 'Tổ Kỹ thuật Địa bàn Phúc Thọ', 'NVKT_DB': 'Trần Văn B',
            'Tổng số': 8, 'Đạt baseline': 4, 'Đã xử lý trong ngày': 2,
            'Tổng đã đạt': 5, 'Chưa đạt': 3, 'OFF/Lỗi': 1, '% đạt': 62.5,
        },
    ]


# --- Parse ngày ---

def test_parse_intraday_date_ok():
    assert quality_routes._parse_shc_cts_intraday_date(
        'Bao_cao_tien_trinh_20260625.xlsx') == '2026-06-25'


def test_parse_intraday_date_invalid_returns_none():
    assert quality_routes._parse_shc_cts_intraday_date('Bao_cao_chot_ngay_20260625.xlsx') is None
    assert quality_routes._parse_shc_cts_intraday_date('Bao_cao_tien_trinh_abc.xlsx') is None


# --- Schema ---

def test_ensure_schema_creates_shc_cts_tables(tmp_path, monkeypatch):
    db_path = _prepare_db(tmp_path, monkeypatch)
    conn = sqlite3.connect(db_path)
    cols_td = {r[1] for r in conn.execute('PRAGMA table_info(shc_cts_tien_do)')}
    cols_ks = {r[1] for r in conn.execute('PRAGMA table_info(shc_cts_kiemsoat)')}
    conn.close()
    assert {'ngay_xu_ly', 'don_vi', 'nvkt_db', 'tong_so', 'dat_baseline',
            'da_xu_ly_ngay', 'tong_dat', 'chua_dat', 'off_loi', 'ty_le_dat',
            'captured_at'} <= cols_td
    assert {'ngay_xu_ly', 'nvkt_db', 'don_vi', 'noi_dung_kiem_soat',
            'nguoi_nhap', 'thoi_diem_nhap', 'thoi_diem_cap_nhat'} <= cols_ks


# --- Sync ---

def test_sync_first_time(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    _write_intraday(tmp_path, monkeypatch, _progress_rows(),
                    'Bao_cao_tien_trinh_20260625.xlsx')
    result = quality_routes._sync_shc_cts_tien_do_to_db()
    assert result['synced'] == 2
    assert result['new'] == 2
    assert result['updated'] == 0
    assert result['skipped'] == 0
    assert result['ngay_xu_ly'] == '2026-06-25'


def test_sync_upsert_keeps_latest_values(tmp_path, monkeypatch):
    db_path = _prepare_db(tmp_path, monkeypatch)
    rows = _progress_rows()
    _write_intraday(tmp_path, monkeypatch, rows, 'Bao_cao_tien_trinh_20260625.xlsx')
    quality_routes._sync_shc_cts_tien_do_to_db()

    # Ghi đè file cùng ngày với giá trị mới
    rows[0] = {**rows[0], 'Tổng đã đạt': 9, 'Chưa đạt': 1}
    _write_intraday(tmp_path, monkeypatch, rows, 'Bao_cao_tien_trinh_20260625.xlsx')
    result = quality_routes._sync_shc_cts_tien_do_to_db()

    assert result['synced'] == 2
    assert result['new'] == 0
    assert result['updated'] == 2
    conn = sqlite3.connect(db_path)
    row = conn.execute(
        "SELECT tong_dat, chua_dat FROM shc_cts_tien_do WHERE nvkt_db='Nguyễn Văn A'"
    ).fetchone()
    conn.close()
    assert row[0] == 9
    assert row[1] == 1


def test_sync_skips_when_intraday_missing(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    monkeypatch.setattr(quality_routes, 'SHC_CTS_INTRADAY_REPORT_DIR', str(tmp_path))
    result = quality_routes._sync_shc_cts_tien_do_to_db()
    assert result['skipped'] == 1
    assert result['reason'] == 'intraday_missing'
```

- [ ] **Step 2: Chạy test để xác nhận FAIL**

Run: `python3 -m pytest tests/test_shc_cts_kiemsoat.py -q`
Expected: FAIL — `AttributeError: module 'blueprints.quality_routes' has no attribute '_parse_shc_cts_intraday_date'` (và các hàm sync).

- [ ] **Step 3: Implement parse + sync**

Trong `blueprints/quality_routes.py`, chèn ngay SAU khối `_ensure_shc_cts_schema()` (cuối Task 2):

```python
def _parse_shc_cts_intraday_date(filename):
    name = os.path.basename(str(filename))
    prefix = 'Bao_cao_tien_trinh_'
    if not name.startswith(prefix) or not name.lower().endswith('.xlsx'):
        return None
    date_text = name[len(prefix):-5]
    try:
        parsed = datetime.strptime(date_text, '%Y%m%d')
    except ValueError:
        return None
    return parsed.strftime('%Y-%m-%d')


def _shc_cts_intraday_row_to_params(row, ngay_xu_ly, now_iso):
    def _num(col):
        v = row.get(col)
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return None
        try:
            return int(v)
        except (TypeError, ValueError):
            return None

    def _real(col):
        v = row.get(col)
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return None
        try:
            return float(v)
        except (TypeError, ValueError):
            return None

    return (
        ngay_xu_ly,
        str(row.get('Đơn vị') or '').strip(),
        str(row.get('NVKT_DB') or '').strip(),
        _num('Tổng số'),
        _num('Đạt baseline'),
        _num('Đã xử lý trong ngày'),
        _num('Tổng đã đạt'),
        _num('Chưa đạt'),
        _num('OFF/Lỗi'),
        _real('% đạt'),
        now_iso,
    )


_SHC_CTS_TIEN_DO_UPSERT_SQL = """
    INSERT INTO shc_cts_tien_do (
        ngay_xu_ly, don_vi, nvkt_db, tong_so, dat_baseline, da_xu_ly_ngay,
        tong_dat, chua_dat, off_loi, ty_le_dat, captured_at
    )
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ON CONFLICT(ngay_xu_ly, nvkt_db) DO UPDATE SET
        don_vi = excluded.don_vi,
        tong_so = excluded.tong_so,
        dat_baseline = excluded.dat_baseline,
        da_xu_ly_ngay = excluded.da_xu_ly_ngay,
        tong_dat = excluded.tong_dat,
        chua_dat = excluded.chua_dat,
        off_loi = excluded.off_loi,
        ty_le_dat = excluded.ty_le_dat,
        captured_at = excluded.captured_at
"""


def _sync_shc_cts_tien_do_to_db():
    """Đọc file intraday SHC CTS mới nhất, upsert snapshot vào shc_cts_tien_do.

    Idempotent. Không raise khi file thiếu / DB lock — trả dict với skipped=1.
    Trả: {'synced': N, 'new': M, 'updated': K, 'skipped': 0|1, 'reason': str,
          'ngay_xu_ly': str}
    """
    intraday_pattern = os.path.join(SHC_CTS_INTRADAY_REPORT_DIR, SHC_CTS_INTRADAY_REPORT_PATTERN)
    intraday_path = latest_matching_file(intraday_pattern)
    if not intraday_path:
        return {'synced': 0, 'new': 0, 'updated': 0, 'skipped': 1,
                'reason': 'intraday_missing', 'ngay_xu_ly': ''}
    ngay_xu_ly = _parse_shc_cts_intraday_date(os.path.basename(intraday_path))
    if not ngay_xu_ly:
        return {'synced': 0, 'new': 0, 'updated': 0, 'skipped': 1,
                'reason': 'date_unparseable', 'ngay_xu_ly': ''}

    _ensure_shc_cts_schema()
    try:
        df = read_excel_sheet_cached(intraday_path, SHC_CTS_INTRADAY_PROGRESS_SHEET)
    except Exception:
        return {'synced': 0, 'new': 0, 'updated': 0, 'skipped': 1,
                'reason': 'excel_unreadable', 'ngay_xu_ly': ngay_xu_ly}

    df = df.copy()
    df['NVKT_DB'] = df['NVKT_DB'].astype(str).str.strip()
    df = df[df['NVKT_DB'].str.len() > 0]
    df = df[~df['NVKT_DB'].str.upper().eq('TỔNG')]
    if df.empty:
        return {'synced': 0, 'new': 0, 'updated': 0, 'skipped': 0,
                'reason': '', 'ngay_xu_ly': ngay_xu_ly}

    all_keys = [(ngay_xu_ly, nvkt) for nvkt in df['NVKT_DB'].tolist()]
    now_iso = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    try:
        with _shc_cts_write_connection() as conn:
            placeholders = ', '.join('?' for _ in df['NVKT_DB'])
            existing_rows = conn.execute(
                f'SELECT nvkt_db FROM shc_cts_tien_do '
                f'WHERE ngay_xu_ly = ? AND nvkt_db IN ({placeholders})',
                [ngay_xu_ly, *df['NVKT_DB'].tolist()],
            ).fetchall()
            existing = {r['nvkt_db'] for r in existing_rows}

            for _, row in df.iterrows():
                params = _shc_cts_intraday_row_to_params(row, ngay_xu_ly, now_iso)
                conn.execute(_SHC_CTS_TIEN_DO_UPSERT_SQL, params)

        new_count = sum(1 for _, nvkt in all_keys if nvkt not in existing)
        updated_count = len(all_keys) - new_count
        return {'synced': len(all_keys), 'new': new_count, 'updated': updated_count,
                'skipped': 0, 'reason': '', 'ngay_xu_ly': ngay_xu_ly}
    except sqlite3.OperationalError as exc:
        msg = str(exc).lower()
        if 'locked' in msg or 'busy' in msg:
            return {'synced': 0, 'new': 0, 'updated': 0, 'skipped': 1,
                    'reason': 'db_locked', 'ngay_xu_ly': ngay_xu_ly}
        raise
```

- [ ] **Step 4: Chạy test để xác nhận PASS**

Run: `python3 -m pytest tests/test_shc_cts_kiemsoat.py -q`
Expected: 7 passed.

- [ ] **Step 5: Commit**

```bash
git add blueprints/quality_routes.py tests/test_shc_cts_kiemsoat.py
git commit -m "quality: thêm sync snapshot tiến độ SHC CTS vào DB nội bộ"
```

---

## Task 4: Endpoint POST `/api/shc-cts-kiemsoat/luu`

**Files:**
- Modify: `blueprints/quality_routes.py` (chèn sau sync, Task 3)
- Modify: `tests/test_shc_cts_kiemsoat.py`

**Interfaces:**
- Produces: route `POST /api/shc-cts-kiemsoat/luu` — nhận JSON `{ngay_xu_ly, nvkt_db, don_vi, noi_dung}`, trả `{ok, ngay_xu_ly, nvkt_db, noi_dung, nguoi_nhap}`. Nếu `noi_dung` rỗng → xóa dòng kiemsoat.

- [ ] **Step 1: Thêm test**

Bổ sung vào cuối `tests/test_shc_cts_kiemsoat.py`:

```python
# --- Luu kiem soat ---

def test_kiemsoat_luu_creates(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    response = _logged_in_client().post(
        '/api/shc-cts-kiemsoat/luu',
        json={'ngay_xu_ly': '2026-06-25', 'nvkt_db': 'Nguyễn Văn A',
              'don_vi': 'Tổ Phúc Thọ', 'noi_dung': 'Đã kiểm tra'},
    )
    assert response.status_code == 200
    body = response.get_json()
    assert body['ok'] is True
    assert body['nguoi_nhap'] == 'test-user'
    assert body['noi_dung'] == 'Đã kiểm tra'


def test_kiemsoat_luu_update_keeps_thoi_diem_nhap(tmp_path, monkeypatch):
    db_path = _prepare_db(tmp_path, monkeypatch)
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            "INSERT INTO shc_cts_kiemsoat (ngay_xu_ly, nvkt_db, noi_dung_kiem_soat, "
            "nguoi_nhap, thoi_diem_nhap, thoi_diem_cap_nhat) "
            "VALUES ('2026-06-25','NV1','cũ','u','2020-01-01 00:00:00','2020-01-01 00:00:00')"
        )
        conn.commit()
    _logged_in_client().post(
        '/api/shc-cts-kiemsoat/luu',
        json={'ngay_xu_ly': '2026-06-25', 'nvkt_db': 'NV1', 'noi_dung': 'mới'},
    )
    conn = sqlite3.connect(db_path)
    row = conn.execute(
        "SELECT thoi_diem_nhap, noi_dung_kiem_soat FROM shc_cts_kiemsoat "
        "WHERE ngay_xu_ly='2026-06-25' AND nvkt_db='NV1'"
    ).fetchone()
    conn.close()
    assert row[0] == '2020-01-01 00:00:00'
    assert row[1] == 'mới'


def test_kiemsoat_luu_empty_deletes(tmp_path, monkeypatch):
    db_path = _prepare_db(tmp_path, monkeypatch)
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            "INSERT INTO shc_cts_kiemsoat (ngay_xu_ly, nvkt_db, noi_dung_kiem_soat) "
            "VALUES ('2026-06-25','NV1','cũ')"
        )
        conn.commit()
    _logged_in_client().post(
        '/api/shc-cts-kiemsoat/luu',
        json={'ngay_xu_ly': '2026-06-25', 'nvkt_db': 'NV1', 'noi_dung': ''},
    )
    conn = sqlite3.connect(db_path)
    count = conn.execute(
        "SELECT COUNT(*) FROM shc_cts_kiemsoat "
        "WHERE ngay_xu_ly='2026-06-25' AND nvkt_db='NV1'"
    ).fetchone()[0]
    conn.close()
    assert count == 0


def test_kiemsoat_luu_rejects_missing_keys(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    response = _logged_in_client().post(
        '/api/shc-cts-kiemsoat/luu',
        json={'ngay_xu_ly': '2026-06-25', 'noi_dung': 'x'},
    )
    assert response.status_code == 400
```

- [ ] **Step 2: Chạy test để xác nhận FAIL**

Run: `python3 -m pytest tests/test_shc_cts_kiemsoat.py -k luu -q`
Expected: FAIL — 404 (route chưa tồn tại).

- [ ] **Step 3: Implement endpoint**

Trong `blueprints/quality_routes.py`, chèn ngay SAU `_sync_shc_cts_tien_do_to_db()`:

```python
@quality_bp.route('/api/shc-cts-kiemsoat/luu', methods=['POST'])
@login_required
def api_shc_cts_kiemsoat_luu():
    payload = request.get_json(silent=True) or {}
    ngay_xu_ly = str(payload.get('ngay_xu_ly') or '').strip()
    nvkt_db = str(payload.get('nvkt_db') or '').strip()
    if not ngay_xu_ly or not nvkt_db:
        return jsonify({'ok': False, 'error': 'ngay_xu_ly và nvkt_db là bắt buộc'}), 400

    noi_dung = str(payload.get('noi_dung') or '').strip()
    if len(noi_dung) > SHC_CTS_KIEMSOAT_NOI_DUNG_MAX:
        return jsonify({'ok': False, 'error': 'Nội dung không được vượt quá 2000 ký tự'}), 400

    don_vi = str(payload.get('don_vi') or '')
    username = session.get('username') or ''
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    _ensure_shc_cts_schema()
    try:
        with _shc_cts_write_connection() as conn:
            if not noi_dung:
                conn.execute(
                    'DELETE FROM shc_cts_kiemsoat WHERE ngay_xu_ly = ? AND nvkt_db = ?',
                    (ngay_xu_ly, nvkt_db),
                )
                return jsonify({'ok': True, 'ngay_xu_ly': ngay_xu_ly,
                                'nvkt_db': nvkt_db, 'noi_dung': '', 'nguoi_nhap': ''})

            existing = conn.execute(
                'SELECT thoi_diem_nhap FROM shc_cts_kiemsoat WHERE ngay_xu_ly = ? AND nvkt_db = ?',
                (ngay_xu_ly, nvkt_db),
            ).fetchone()
            thoi_diem_nhap = existing['thoi_diem_nhap'] if existing else now
            conn.execute(
                '''
                INSERT INTO shc_cts_kiemsoat
                    (ngay_xu_ly, nvkt_db, don_vi, noi_dung_kiem_soat,
                     nguoi_nhap, thoi_diem_nhap, thoi_diem_cap_nhat)
                VALUES (?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(ngay_xu_ly, nvkt_db) DO UPDATE SET
                    don_vi = excluded.don_vi,
                    noi_dung_kiem_soat = excluded.noi_dung_kiem_soat,
                    nguoi_nhap = excluded.nguoi_nhap,
                    thoi_diem_cap_nhat = excluded.thoi_diem_cap_nhat
                ''',
                (ngay_xu_ly, nvkt_db, don_vi, noi_dung, username, thoi_diem_nhap, now),
            )
    except sqlite3.Error as exc:
        return jsonify({'ok': False, 'error': str(exc)}), 500

    return jsonify({'ok': True, 'ngay_xu_ly': ngay_xu_ly, 'nvkt_db': nvkt_db,
                    'noi_dung': noi_dung, 'nguoi_nhap': username})
```

- [ ] **Step 4: Chạy test để xác nhận PASS**

Run: `python3 -m pytest tests/test_shc_cts_kiemsoat.py -k luu -q`
Expected: 4 passed.

- [ ] **Step 5: Commit**

```bash
git add blueprints/quality_routes.py tests/test_shc_cts_kiemsoat.py
git commit -m "quality: thêm endpoint lưu kiểm soát SHC CTS theo (ngày, NVKT)"
```

---

## Task 5: Endpoint GET `/api/shc-cts-kiemsoat/detail`

**Files:**
- Modify: `blueprints/quality_routes.py`
- Modify: `tests/test_shc_cts_kiemsoat.py`

**Interfaces:**
- Produces: route `GET /api/shc-cts-kiemsoat/detail?date=YYYY-MM-DD`.
- Produces: helper `_load_shc_cts_kiemsoat_df(ngay_xu_ly) -> pandas.DataFrame` (cột gốc + `kiemsoat_noi_dung`, `kiemsoat_nguoi_nhap`, `kiemsoat_thoi_diem`, `kiemsoat_da_nhap`).
- Response shape:
  ```json
  {
    "selected_date": "YYYY-MM-DD",
    "is_today_live": true,
    "available_dates": ["2026-06-25", "2026-06-24"],
    "sheets": {"<don_vi>": {"columns": [...], "data": [...]}}
  }
  ```
- Quy tắc chọn nguồn: nếu `date` rỗng HOẶC bằng ngày của file intraday mới nhất → đọc Excel trực tiếp (live) + lazy-sync; ngược lại (ngày quá khứ) → đọc từ `shc_cts_tien_do`.

- [ ] **Step 1: Thêm test**

Bổ sung vào cuối `tests/test_shc_cts_kiemsoat.py`:

```python
# --- Detail ---

def test_detail_today_reads_live_excel_and_syncs(tmp_path, monkeypatch):
    db_path = _prepare_db(tmp_path, monkeypatch)
    _write_intraday(tmp_path, monkeypatch, _progress_rows(),
                    'Bao_cao_tien_trinh_20260625.xlsx')
    response = _logged_in_client().get('/api/shc-cts-kiemsoat/detail')
    assert response.status_code == 200
    body = response.get_json()
    assert body['selected_date'] == '2026-06-25'
    assert body['is_today_live'] is True
    assert '2026-06-25' in body['available_dates']
    don_vi = 'Tổ Kỹ thuật Địa bàn Phúc Thọ'
    assert don_vi in body['sheets']
    assert body['sheets'][don_vi]['data'][0]['NVKT_DB'] == 'Nguyễn Văn A'
    # Lazy sync đã ghi vào DB
    conn = sqlite3.connect(db_path)
    n = conn.execute('SELECT COUNT(*) FROM shc_cts_tien_do').fetchone()[0]
    conn.close()
    assert n == 2


def test_detail_past_date_reads_from_db(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    _write_intraday(tmp_path, monkeypatch, _progress_rows(),
                    'Bao_cao_tien_trinh_20260625.xlsx')
    quality_routes._sync_shc_cts_tien_do_to_db()
    # Xóa file intraday để giả lập "không có file của ngày cũ"
    _write_intraday(tmp_path, monkeypatch, _progress_rows(),
                    'Bao_cao_tien_trinh_20260626.xlsx')

    response = _logged_in_client().get('/api/shc-cts-kiemsoat/detail?date=2026-06-25')
    assert response.status_code == 200
    body = response.get_json()
    assert body['selected_date'] == '2026-06-25'
    assert body['is_today_live'] is False
    don_vi = 'Tổ Kỹ thuật Địa bàn Phúc Thọ'
    assert body['sheets'][don_vi]['data'][0]['NVKT_DB'] == 'Nguyễn Văn A'


def test_detail_joins_kiemsoat_annotation(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    _write_intraday(tmp_path, monkeypatch, _progress_rows(),
                    'Bao_cao_tien_trinh_20260625.xlsx')
    _logged_in_client().post(
        '/api/shc-cts-kiemsoat/luu',
        json={'ngay_xu_ly': '2026-06-25', 'nvkt_db': 'Nguyễn Văn A',
              'don_vi': 'Tổ Kỹ thuật Địa bàn Phúc Thọ', 'noi_dung': 'Ghi chú KS'},
    )
    response = _logged_in_client().get('/api/shc-cts-kiemsoat/detail')
    body = response.get_json()
    don_vi = 'Tổ Kỹ thuật Địa bàn Phúc Thọ'
    row = body['sheets'][don_vi]['data'][0]
    assert row['kiemsoat_noi_dung'] == 'Ghi chú KS'
    assert row['kiemsoat_da_nhap'] is True


def test_detail_returns_404_when_no_data(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    monkeypatch.setattr(quality_routes, 'SHC_CTS_INTRADAY_REPORT_DIR', str(tmp_path))
    response = _logged_in_client().get('/api/shc-cts-kiemsoat/detail?date=2099-01-01')
    assert response.status_code == 404
```

- [ ] **Step 2: Chạy test để xác nhận FAIL**

Run: `python3 -m pytest tests/test_shc_cts_kiemsoat.py -k detail -q`
Expected: FAIL — 404.

- [ ] **Step 3: Implement helper + endpoint**

Trong `blueprints/quality_routes.py`, chèn SAU endpoint `/api/shc-cts-kiemsoat/luu`:

```python
def _resolve_shc_cts_selected_date(requested_date):
    """Trả (selected_date, is_today_live). today_live khi date rỗng hoặc == ngày file intraday mới nhất."""
    intraday_pattern = os.path.join(SHC_CTS_INTRADAY_REPORT_DIR, SHC_CTS_INTRADAY_REPORT_PATTERN)
    intraday_path = latest_matching_file(intraday_pattern)
    today_file_date = _parse_shc_cts_intraday_date(os.path.basename(intraday_path)) if intraday_path else None
    requested = (requested_date or '').strip()
    if not requested:
        return today_file_date or '', bool(today_file_date)
    return requested, (requested == today_file_date)


def _get_shc_cts_available_dates():
    _ensure_shc_cts_schema()
    with _shc_cts_read_connection() as conn:
        rows = conn.execute(
            'SELECT DISTINCT ngay_xu_ly FROM shc_cts_tien_do ORDER BY ngay_xu_ly DESC'
        ).fetchall()
    db_dates = [r['ngay_xu_ly'] for r in rows if r['ngay_xu_ly']]

    intraday_pattern = os.path.join(SHC_CTS_INTRADAY_REPORT_DIR, SHC_CTS_INTRADAY_REPORT_PATTERN)
    intraday_path = latest_matching_file(intraday_pattern)
    if intraday_path:
        file_date = _parse_shc_cts_intraday_date(os.path.basename(intraday_path))
        if file_date and file_date not in db_dates:
            db_dates.insert(0, file_date)
    return db_dates


def _apply_shc_cts_kiemsoat_annotation(df, ngay_xu_ly):
    """Join df (cột gốc có 'NVKT_DB') với shc_cts_kiemsoat. df phải có cột NVKT_DB str."""
    _ensure_shc_cts_schema()
    nvkt_list = [str(v).strip() for v in df['NVKT_DB'].tolist()]
    ks_map = {}
    if nvkt_list:
        placeholders = ', '.join('?' for _ in nvkt_list)
        with _shc_cts_read_connection() as conn:
            rows = conn.execute(
                f'''SELECT nvkt_db, noi_dung_kiem_soat, nguoi_nhap,
                           thoi_diem_cap_nhat, thoi_diem_nhap
                    FROM shc_cts_kiemsoat
                    WHERE ngay_xu_ly = ? AND nvkt_db IN ({placeholders})''',
                [ngay_xu_ly, *nvkt_list],
            ).fetchall()
        ks_map = {r['nvkt_db']: dict(r) for r in rows}
    df = df.copy()
    df['kiemsoat_noi_dung'] = df['NVKT_DB'].map(lambda k: ks_map.get(k, {}).get('noi_dung_kiem_soat', ''))
    df['kiemsoat_nguoi_nhap'] = df['NVKT_DB'].map(lambda k: ks_map.get(k, {}).get('nguoi_nhap', ''))
    df['kiemsoat_thoi_diem'] = df['NVKT_DB'].map(
        lambda k: ks_map.get(k, {}).get('thoi_diem_cap_nhat') or ks_map.get(k, {}).get('thoi_diem_nhap') or ''
    )
    df['kiemsoat_da_nhap'] = df['NVKT_DB'].isin(ks_map.keys())
    return df


def _load_shc_cts_kiemsoat_df(ngay_xu_ly):
    """Trả DataFrame tiến độ ngày đã chọn + annotation kiểm soát, hoặc None."""
    intraday_pattern = os.path.join(SHC_CTS_INTRADAY_REPORT_DIR, SHC_CTS_INTRADAY_REPORT_PATTERN)
    intraday_path = latest_matching_file(intraday_pattern)
    today_file_date = _parse_shc_cts_intraday_date(os.path.basename(intraday_path)) if intraday_path else None

    if ngay_xu_ly == today_file_date and intraday_path:
        try:
            _sync_shc_cts_tien_do_to_db()
        except Exception:
            current_app.logger.warning('shc_cts_tien_do sync thất bại trong detail', exc_info=True)
        try:
            df = read_excel_sheet_cached(intraday_path, SHC_CTS_INTRADAY_PROGRESS_SHEET)
        except Exception:
            return None
    else:
        _ensure_shc_cts_schema()
        with _shc_cts_read_connection() as conn:
            rows = conn.execute(
                '''SELECT don_vi AS "Đơn vị", nvkt_db AS "NVKT_DB",
                          tong_so AS "Tổng số", dat_baseline AS "Đạt baseline",
                          da_xu_ly_ngay AS "Đã xử lý trong ngày",
                          tong_dat AS "Tổng đã đạt", chua_dat AS "Chưa đạt",
                          off_loi AS "OFF/Lỗi", ty_le_dat AS "% đạt"
                   FROM shc_cts_tien_do WHERE ngay_xu_ly = ?
                   ORDER BY don_vi, nvkt_db''',
                (ngay_xu_ly,),
            ).fetchall()
        if not rows:
            return None
        df = pd.DataFrame([dict(r) for r in rows])

    df['NVKT_DB'] = df['NVKT_DB'].astype(str).str.strip()
    df = df[df['NVKT_DB'].str.len() > 0]
    df = df[~df['NVKT_DB'].str.upper().eq('TỔNG')]
    if df.empty:
        return None
    return _apply_shc_cts_kiemsoat_annotation(df, ngay_xu_ly)


@quality_bp.route('/api/shc-cts-kiemsoat/detail')
@login_required
def api_shc_cts_kiemsoat_detail():
    selected_date, is_today_live = _resolve_shc_cts_selected_date(request.args.get('date'))
    if not selected_date:
        return jsonify({'error': 'Chưa có dữ liệu tiến độ SHC CTS'}), 404

    df = _load_shc_cts_kiemsoat_df(selected_date)
    if df is None:
        return jsonify({'error': f'Không có dữ liệu tiến độ ngày {selected_date}'}), 404

    sheets = {}
    don_vi_col = 'Đơn vị' if 'Đơn vị' in df.columns else None
    if don_vi_col:
        for don_vi, group in df.groupby(don_vi_col, sort=False):
            sheets[str(don_vi)] = build_sheet_payload(group)
    else:
        sheets['Tất cả'] = build_sheet_payload(df)

    return jsonify({
        'selected_date': selected_date,
        'is_today_live': is_today_live,
        'available_dates': _get_shc_cts_available_dates(),
        'sheets': sheets,
    })
```

- [ ] **Step 4: Chạy test để xác nhận PASS**

Run: `python3 -m pytest tests/test_shc_cts_kiemsoat.py -k detail -q`
Expected: 4 passed.

- [ ] **Step 5: Commit**

```bash
git add blueprints/quality_routes.py tests/test_shc_cts_kiemsoat.py
git commit -m "quality: thêm endpoint detail kiểm soát SHC CTS (live hôm nay + DB ngày cũ)"
```

---

## Task 6: Endpoint GET `/api/shc-cts-kiemsoat/thongke`

**Files:**
- Modify: `blueprints/quality_routes.py`
- Modify: `tests/test_shc_cts_kiemsoat.py`

**Interfaces:**
- Produces: route `GET /api/shc-cts-kiemsoat/thongke?date=YYYY-MM-DD&don_vi=<optional>`.
- Response shape:
  ```json
  {
    "selected_date": "YYYY-MM-DD",
    "summary": {"tong_so": N, "tong_dat": N, "chua_dat": N, "da_xu_ly_ngay": N,
                "da_ks": N, "chua_ks": N, "ty_le_dat": F},
    "by_don_vi": [{"don_vi": "...", "tong_so": N, "tong_dat": N, "chua_dat": N,
                   "da_ks": N, "chua_ks": N, "ty_le_dat": F}],
    "lich_su": [{"ngay_xu_ly": "...", "tong_so": N, "tong_dat": N, "da_ks": N}]
  }
  ```

- [ ] **Step 1: Thêm test**

Bổ sung vào cuối `tests/test_shc_cts_kiemsoat.py`:

```python
# --- Thong ke ---

def test_thongke_returns_summary_and_by_don_vi(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    _write_intraday(tmp_path, monkeypatch, _progress_rows(),
                    'Bao_cao_tien_trinh_20260625.xlsx')
    response = _logged_in_client().get('/api/shc-cts-kiemsoat/thongke?date=2026-06-25')
    assert response.status_code == 200
    body = response.get_json()
    assert body['selected_date'] == '2026-06-25'
    assert body['summary']['tong_so'] == 18  # 10 + 8
    assert body['summary']['tong_dat'] == 11  # 6 + 5
    assert body['summary']['da_ks'] == 0
    assert len(body['by_don_vi']) == 1
    assert body['by_don_vi'][0]['don_vi'] == 'Tổ Kỹ thuật Địa bàn Phúc Thọ'


def test_thongke_filter_by_don_vi(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    rows = _progress_rows() + [
        {'Đơn vị': 'Tổ Kỹ thuật Địa bàn Sơn Tây', 'NVKT_DB': 'NV C',
         'Tổng số': 5, 'Đạt baseline': 2, 'Đã xử lý trong ngày': 1,
         'Tổng đã đạt': 3, 'Chưa đạt': 2, 'OFF/Lỗi': 0, '% đạt': 60},
    ]
    _write_intraday(tmp_path, monkeypatch, rows, 'Bao_cao_tien_trinh_20260625.xlsx')
    response = _logged_in_client().get(
        '/api/shc-cts-kiemsoat/thongke?date=2026-06-25&don_vi=Tổ Kỹ thuật Địa bàn Phúc Thọ')
    body = response.get_json()
    assert body['summary']['tong_so'] == 18
    assert len(body['by_don_vi']) == 1


def test_thongke_lich_su_multiple_days(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    _write_intraday(tmp_path, monkeypatch, _progress_rows(),
                    'Bao_cao_tien_trinh_20260624.xlsx')
    quality_routes._sync_shc_cts_tien_do_to_db()
    _write_intraday(tmp_path, monkeypatch, _progress_rows(),
                    'Bao_cao_tien_trinh_20260625.xlsx')
    quality_routes._sync_shc_cts_tien_do_to_db()

    response = _logged_in_client().get('/api/shc-cts-kiemsoat/thongke?date=2026-06-25')
    body = response.get_json()
    days = {item['ngay_xu_ly'] for item in body['lich_su']}
    assert {'2026-06-24', '2026-06-25'} <= days


def test_thongke_counts_kiemsoat(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    _write_intraday(tmp_path, monkeypatch, _progress_rows(),
                    'Bao_cao_tien_trinh_20260625.xlsx')
    _logged_in_client().post(
        '/api/shc-cts-kiemsoat/luu',
        json={'ngay_xu_ly': '2026-06-25', 'nvkt_db': 'Nguyễn Văn A',
              'don_vi': 'Tổ Kỹ thuật Địa bàn Phúc Thọ', 'noi_dung': 'KS xong'},
    )
    response = _logged_in_client().get('/api/shc-cts-kiemsoat/thongke?date=2026-06-25')
    body = response.get_json()
    assert body['summary']['da_ks'] == 1
    assert body['summary']['chua_ks'] == 1


def test_thongke_404_when_no_data(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    response = _logged_in_client().get('/api/shc-cts-kiemsoat/thongke?date=2099-01-01')
    assert response.status_code == 404
```

- [ ] **Step 2: Chạy test để xác nhận FAIL**

Run: `python3 -m pytest tests/test_shc_cts_kiemsoat.py -k thongke -q`
Expected: FAIL — 404.

- [ ] **Step 3: Implement endpoint**

Trong `blueprints/quality_routes.py`, chèn SAU endpoint detail:

```python
def _compute_shc_cts_thongke(df, ngay_xu_ly):
    """df: DataFrame có các cột gốc + kiemsoat_da_nhap. Trả dict summary + by_don_vi."""
    def _agg(group):
        return {
            'tong_so': int(group['Tổng số'].sum()) if 'Tổng số' in group else 0,
            'tong_dat': int(group['Tổng đã đạt'].sum()) if 'Tổng đã đạt' in group else 0,
            'chua_dat': int(group['Chưa đạt'].sum()) if 'Chưa đạt' in group else 0,
            'da_xu_ly_ngay': int(group['Đã xử lý trong ngày'].sum()) if 'Đã xử lý trong ngày' in group else 0,
            'da_ks': int(group['kiemsoat_da_nhap'].sum()) if 'kiemsoat_da_nhap' in group else 0,
            'chua_ks': int(len(group) - group['kiemsoat_da_nhap'].sum()) if 'kiemsoat_da_nhap' in group else len(group),
        }

    summary = _agg(df)
    tong_so = summary['tong_so']
    summary['ty_le_dat'] = round(summary['tong_dat'] * 100.0 / tong_so, 1) if tong_so else 0.0

    by_don_vi = []
    if 'Đơn vị' in df.columns:
        for don_vi, group in df.groupby('Đơn vị', sort=False):
            row = {'don_vi': str(don_vi), **_agg(group)}
            row['ty_le_dat'] = round(row['tong_dat'] * 100.0 / row['tong_so'], 1) if row['tong_so'] else 0.0
            by_don_vi.append(row)
        by_don_vi.sort(key=lambda r: r['tong_so'], reverse=True)

    return {'summary': summary, 'by_don_vi': by_don_vi}


def _compute_shc_cts_lich_su(ngay_xu_ly):
    """Trả list trend nhiều ngày: tong_so, tong_dat, da_ks theo ngay_xu_ly."""
    _ensure_shc_cts_schema()
    with _shc_cts_read_connection() as conn:
        rows = conn.execute(
            '''SELECT ngay_xu_ly,
                      SUM(COALESCE(tong_so, 0)) AS tong_so,
                      SUM(COALESCE(tong_dat, 0)) AS tong_dat
               FROM shc_cts_tien_do
               GROUP BY ngay_xu_ly
               ORDER BY ngay_xu_ly DESC
               LIMIT 60'''
        ).fetchall()
        ks_rows = []
        if rows:
            dates = [r['ngay_xu_ly'] for r in rows]
            placeholders = ', '.join('?' for _ in dates)
            ks_rows = conn.execute(
                f'''SELECT ngay_xu_ly, COUNT(*) AS da_ks
                    FROM shc_cts_kiemsoat
                    WHERE noi_dung_kiem_soat != ''
                      AND ngay_xu_ly IN ({placeholders})
                    GROUP BY ngay_xu_ly''',
                dates,
            ).fetchall()
        ks_map = {r['ngay_xu_ly']: r['da_ks'] for r in ks_rows}
    return [
        {
            'ngay_xu_ly': r['ngay_xu_ly'],
            'tong_so': r['tong_so'],
            'tong_dat': r['tong_dat'],
            'da_ks': ks_map.get(r['ngay_xu_ly'], 0),
        }
        for r in rows
    ]


@quality_bp.route('/api/shc-cts-kiemsoat/thongke')
@login_required
def api_shc_cts_kiemsoat_thongke():
    selected_date, _ = _resolve_shc_cts_selected_date(request.args.get('date'))
    if not selected_date:
        return jsonify({'error': 'Chưa có dữ liệu tiến độ SHC CTS'}), 404

    df = _load_shc_cts_kiemsoat_df(selected_date)
    if df is None:
        return jsonify({'error': f'Không có dữ liệu tiến độ ngày {selected_date}'}), 404

    don_vi_filter = (request.args.get('don_vi') or '').strip()
    if don_vi_filter and 'Đơn vị' in df.columns:
        df = df[df['Đơn vị'].astype(str) == don_vi_filter]

    thongke = _compute_shc_cts_thongke(df, selected_date)
    return jsonify({
        'selected_date': selected_date,
        'summary': thongke['summary'],
        'by_don_vi': thongke['by_don_vi'],
        'lich_su': _compute_shc_cts_lich_su(selected_date),
    })
```

- [ ] **Step 4: Chạy test để xác nhận PASS**

Run: `python3 -m pytest tests/test_shc_cts_kiemsoat.py -k thongke -q`
Expected: 5 passed.

- [ ] **Step 5: Commit**

```bash
git add blueprints/quality_routes.py tests/test_shc_cts_kiemsoat.py
git commit -m "quality: thêm endpoint thống kê kiểm soát SHC CTS (summary + by_don_vi + lịch sử)"
```

---

## Task 7: Endpoint GET `/download/shc-cts-kiemsoat-report`

**Files:**
- Modify: `blueprints/quality_routes.py`
- Modify: `tests/test_shc_cts_kiemsoat.py`

**Interfaces:**
- Produces: route `GET /download/shc-cts-kiemsoat-report?date=YYYY-MM-DD&don_vi=<optional>` — trả file `.xlsx`.

- [ ] **Step 1: Thêm test**

Bổ sung vào cuối `tests/test_shc_cts_kiemsoat.py`:

```python
# --- Download report ---

def test_report_download_returns_xlsx(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    _write_intraday(tmp_path, monkeypatch, _progress_rows(),
                    'Bao_cao_tien_trinh_20260625.xlsx')
    response = _logged_in_client().get('/download/shc-cts-kiemsoat-report?date=2026-06-25')
    assert response.status_code == 200
    assert 'spreadsheetml' in response.mimetype
    assert 'shc_cts_kiemsoat_2026-06-25' in response.headers['Content-Disposition']


def test_report_download_404_when_no_data(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    response = _logged_in_client().get('/download/shc-cts-kiemsoat-report?date=2099-01-01')
    assert response.status_code == 404
```

- [ ] **Step 2: Chạy test để xác nhận FAIL**

Run: `python3 -m pytest tests/test_shc_cts_kiemsoat.py -k report -q`
Expected: FAIL — 404.

- [ ] **Step 3: Implement endpoint**

Trong `blueprints/quality_routes.py`, chèn SAU endpoint thongke:

```python
@quality_bp.route('/download/shc-cts-kiemsoat-report')
@login_required
def download_shc_cts_kiemsoat_report():
    selected_date, _ = _resolve_shc_cts_selected_date(request.args.get('date'))
    if not selected_date:
        return jsonify({'error': 'Chưa có dữ liệu tiến độ SHC CTS'}), 404

    df = _load_shc_cts_kiemsoat_df(selected_date)
    if df is None:
        return jsonify({'error': f'Không có dữ liệu tiến độ ngày {selected_date}'}), 404

    don_vi_filter = (request.args.get('don_vi') or '').strip()
    if don_vi_filter and 'Đơn vị' in df.columns:
        df = df[df['Đơn vị'].astype(str) == don_vi_filter]

    export_df = serialize_dataframe(df)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        export_df.to_excel(writer, index=False, sheet_name='Kiem_soat_SHC_CTS')
    output.seek(0)

    suffix = f'_{don_vi_filter}' if don_vi_filter else ''
    download_name = f'shc_cts_kiemsoat_{selected_date}{suffix}.xlsx'
    return send_file(
        output,
        as_attachment=True,
        download_name=download_name,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )
```

- [ ] **Step 4: Chạy toàn bộ test file để xác nhận PASS**

Run: `python3 -m pytest tests/test_shc_cts_kiemsoat.py -q`
Expected: tất cả PASS.

- [ ] **Step 5: Chạy full test suite để chắc không phá test cũ**

Run: `python3 -m pytest tests/ -q`
Expected: tất cả PASS (không có regression).

- [ ] **Step 6: Commit**

```bash
git add blueprints/quality_routes.py tests/test_shc_cts_kiemsoat.py
git commit -m "quality: thêm endpoint xuất Excel báo cáo kiểm soát SHC CTS"
```

---

## Task 8: Cron script `scripts/sync_shc_cts_tien_do.py`

**Files:**
- Create: `scripts/sync_shc_cts_tien_do.py`
- Modify: `tests/test_shc_cts_kiemsoat.py`

**Interfaces:**
- Produces: script chạy được `python3 scripts/sync_shc_cts_tien_do.py` (env `DASHV4_UNIT_CODE=<unit>`).

- [ ] **Step 1: Thêm test**

Bổ sung vào cuối `tests/test_shc_cts_kiemsoat.py`:

```python
# --- Cron script ---

def test_sync_script_runs(tmp_path, monkeypatch):
    import subprocess
    _prepare_db(tmp_path, monkeypatch)
    _write_intraday(tmp_path, monkeypatch, _progress_rows(),
                    'Bao_cao_tien_trinh_20260625.xlsx')
    env = {
        **dict(__import__('os').environ),
        'DASHV4_RUNTIME_DIR': str(tmp_path),
        'DASHV4_SHC_CTS_HISTORY_DB_PATH': str(tmp_path / 'shc_cts.db'),
    }
    result = subprocess.run(
        [sys.executable, 'scripts/sync_shc_cts_tien_do.py'],
        cwd=str(Path(__file__).resolve().parents[1]),
        env=env, capture_output=True, text=True,
    )
    assert result.returncode == 0, result.stderr
```

- [ ] **Step 2: Tạo script**

Tạo `scripts/sync_shc_cts_tien_do.py`:

```python
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
```

- [ ] **Step 3: Cấp quyền thực thi + chạy test**

Run:
```bash
chmod +x scripts/sync_shc_cts_tien_do.py
python3 -m pytest tests/test_shc_cts_kiemsoat.py -k script -q
```
Expected: 1 passed.

- [ ] **Step 4: Commit**

```bash
git add scripts/sync_shc_cts_tien_do.py tests/test_shc_cts_kiemsoat.py
git commit -m "scripts: thêm cron sync snapshot tiến độ SHC CTS"
```

---

## Task 9: API client JS cho SHC CTS kiemsoat

**Files:**
- Modify: `static/js/api.js`

**Interfaces:**
- Produces: `API.getShcCtsKiemSoatDetail(query)`, `API.getShcCtsKiemSoatThongKe(query)`, `API.saveShcCtsKiemSoat(payload)`.

- [ ] **Step 1: Thêm methods**

Mở `static/js/api.js`. Tìm method `savePttbKiemSoat` (khoảng dòng 140-152) và kết thúc bằng `return response.json();\n    },`. Ngay SAU khối `savePttbKiemSoat` (sau dấu `},` đóng method), chèn:

```javascript

    /* ========================================
       SHC CTS APIs (Tiến độ + Kiểm soát)
       ======================================== */

    getShcCtsKiemSoatDetail(query = '') {
        const suffix = query ? `?${query}` : '';
        return this.fetchData(`/api/shc-cts-kiemsoat/detail${suffix}`);
    },

    getShcCtsKiemSoatThongKe(query = '') {
        const suffix = query ? `?${query}` : '';
        return this.fetchData(`/api/shc-cts-kiemsoat/thongke${suffix}`);
    },

    async saveShcCtsKiemSoat(payload) {
        const response = await fetch('/api/shc-cts-kiemsoat/luu', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload),
        });
        if (response.status === 401) {
            window.location.href = '/login';
            throw new Error('Phiên đăng nhập đã hết hạn');
        }
        return response.json();
    },
```

- [ ] **Step 2: Xác minh cú pháp JS**

Run: `node --check static/js/api.js`
Expected: không có output (thành công). Nếu môi trường không có node, bỏ qua bước này.

- [ ] **Step 3: Commit**

```bash
git add static/js/api.js
git commit -m "api.js: thêm client methods cho kiểm soát SHC CTS"
```

---

## Task 10: Markup section kiểm soát trong shc_cts.html

**Files:**
- Modify: `templates/pages/shc_cts.html`

**Interfaces:**
- Produces: section HTML có id `shc-cts-kiemsoat-section` với datepicker `shc-cts-ks-date`, dropdown tổ `shc-cts-ks-don-vi`, container stats `shc-cts-ks-stats`, container chi tiết `shc-cts-ks-chitiet`, container lịch sử `shc-cts-ks-lich-su`.

- [ ] **Step 1: Thêm CSS vào block extra_css**

Mở `templates/pages/shc_cts.html`. Trong block `{% block extra_css %}`, ngay TRƯỚC thẻ đóng `</style>`, chèn:

```css
    .shc-cts-ks-filters { display:flex; gap:10px; flex-wrap:wrap; align-items:center; margin-bottom:12px; }
    .shc-cts-ks-filters label { font-size:0.85rem; color:#3a4a5e; }
    .shc-cts-ks-filters select, .shc-cts-ks-filters input[type=date] {
        padding:4px 8px; border:1px solid #c4d0e0; border-radius:4px;
    }
    .shc-cts-ks-cards { display:flex; gap:12px; flex-wrap:wrap; }
    .shc-cts-ks-card {
        background:#fff; border:1px solid #dce8df; border-radius:10px;
        padding:12px 16px; min-width:130px; box-shadow:0 2px 8px rgba(31,76,48,0.06);
    }
    .shc-cts-ks-card .v { font-size:1.4rem; font-weight:700; color:#173d2b; }
    .shc-cts-ks-card .l { font-size:0.8rem; color:#607568; }
    .shc-cts-ks-cell textarea {
        width:100%; padding:4px; border:1px solid #c4d0e0; border-radius:4px; font-size:0.82rem;
    }
    .shc-cts-ks-save-btn {
        margin-top:3px; padding:2px 10px; font-size:0.78rem; cursor:pointer;
        background:#198754; color:#fff; border:none; border-radius:4px;
    }
    .shc-cts-ks-badge { font-size:0.74rem; }
    .shc-cts-ks-da { color:#198754; font-weight:600; }
    .shc-cts-ks-chua { color:#dc3545; }
    .shc-cts-ks-live-pill {
        display:inline-block; padding:2px 8px; font-size:0.72rem; border-radius:999px;
        background:#fff3cd; color:#664d03; margin-left:8px;
    }
    .shc-cts-ks-live-pill.db { background:#d1e7dd; color:#0f5132; }
```

- [ ] **Step 2: Thêm section HTML vào block content**

Tìm section "Tiến độ xử lý shc trong ngày" (khoảng dòng 179-193). Ngay SAU thẻ `</div>` đóng section đó (trước thẻ `</div>` cuối cùng đóng `#shc-cts-section`), chèn section mới:

```html
    <div class="section" id="shc-cts-kiemsoat-section">
        <h3 class="section-title">
            <i class="fas fa-clipboard-check"></i>
            Lịch sử tiến độ &amp; Kiểm soát tổ trưởng
        </h3>
        <div class="shc-cts-ks-filters">
            <label>Ngày xử lý:
                <input type="date" id="shc-cts-ks-date">
            </label>
            <label>Tổ:
                <select id="shc-cts-ks-don-vi">
                    <option value="">Tất cả</option>
                </select>
            </label>
            <button class="download-btn" onclick="exportShcCtsKiemSoat()">
                <i class="fas fa-file-excel"></i> Xuất Excel
            </button>
            <span id="shc-cts-ks-source-pill"></span>
        </div>
        <div id="shc-cts-ks-stats" class="shc-cts-ks-cards"></div>
        <h4 style="margin-top:16px;">Lịch sử theo ngày</h4>
        <div id="shc-cts-ks-lich-su"></div>
        <h4 style="margin-top:16px;">Chi tiết theo NVKT (nhập kiểm soát)</h4>
        <div id="shc-cts-ks-chitiet"></div>
    </div>
```

- [ ] **Step 3: Xác minh template render không lỗi**

Run: `python3 -c "from dashboard import app; c=app.test_client();
import sys
"`
Nếu lỗi cú pháp trên, dùng cách đơn giản hơn:
Run: `python3 -c "
from dashboard import app
with app.test_client() as c:
    with c.session_transaction() as s: s['username']='t'
    r = c.get('/shc-cts')
    print(r.status_code)
"`
Expected: in ra `200`.

- [ ] **Step 4: Commit**

```bash
git add templates/pages/shc_cts.html
git commit -m "shc_cts: thêm markup section lịch sử & kiểm soát tổ trưởng"
```

---

## Task 11: Logic JS section kiểm soát trong shc_cts.js

**Files:**
- Modify: `static/js/pages/shc_cts.js`

**Interfaces:**
- Consumes: `API.getShcCtsKiemSoatDetail`, `API.getShcCtsKiemSoatThongKe`, `API.saveShcCtsKiemSoat` (từ Task 9); section markup (Task 10).
- Produces: `initShcCtsKiemSoat()`, `loadShcCtsKiemSoatDetail()`, `loadShcCtsKiemSoatThongKe()`, `saveShcCtsKiemSoatRow(nvktDb)`, `exportShcCtsKiemSoat()`.

- [ ] **Step 1: Hook init vào DOMContentLoaded**

Mở `static/js/pages/shc_cts.js`. Tìm khối:

```javascript
document.addEventListener('DOMContentLoaded', async function () {
    await initSHCCTSDetailDownloader();
    await loadSHCCTSData();
});
```

Thay bằng:

```javascript
document.addEventListener('DOMContentLoaded', async function () {
    await initSHCCTSDetailDownloader();
    await loadSHCCTSData();
    initShcCtsKiemSoat();
});
```

- [ ] **Step 2: Thêm toàn bộ logic kiemsoat**

Chèn khối sau vào CUỐI file `static/js/pages/shc_cts.js`:

```javascript

/* ========================================
   LỊCH SỬ TIẾN ĐỘ & KIỂM SOÁT TỔ TRƯỞNG SHC CTS
   - Datepicker chọn ngày (hôm nay = live Excel, cũ = DB)
   - Bảng chi tiết theo NVKT + cột nhập kiểm soát
   ======================================== */

const SHC_CTS_KS_DISPLAY_COLS = [
    'NVKT_DB', 'Tổng số', 'Đạt baseline', 'Đã xử lý trong ngày',
    'Tổng đã đạt', 'Chưa đạt', 'OFF/Lỗi', '% đạt',
];
const SHC_CTS_KS_DISPLAY_LABELS = {
    'NVKT_DB': 'NVKT', 'Tổng số': 'Tổng', 'Đạt baseline': 'Đạt BL',
    'Đã xử lý trong ngày': 'XL trong ngày', 'Tổng đã đạt': 'Đã đạt',
    'Chưa đạt': 'Chưa đạt', 'OFF/Lỗi': 'OFF/Lỗi', '% đạt': '%',
};
let _shcCtsKsCurrent = { date: '', donVi: '', sheets: {}, sourcePill: '' };

function _shcCtsKsEscape(v) {
    return String(v == null ? '' : v)
        .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function _shcCtsKsThoiDiem(value) {
    if (!value) return '';
    const m = String(value).match(/^(\d{4})-(\d{2})-(\d{2}) (\d{2}):(\d{2})/);
    return m ? `${m[4]}:${m[5]} ${m[3]}/${m[2]}` : value;
}

function _shcCtsKsBadge(row) {
    if (row.kiemsoat_da_nhap) {
        const when = _shcCtsKsThoiDiem(row.kiemsoat_thoi_diem);
        const who = row.kiemsoat_nguoi_nhap || '';
        return `<span class="shc-cts-ks-badge shc-cts-ks-da">Đã KS${when ? ' ' + when : ''}${who ? ' — ' + _shcCtsKsEscape(who) : ''}</span>`;
    }
    return `<span class="shc-cts-ks-badge shc-cts-ks-chua">Chưa</span>`;
}

function initShcCtsKiemSoat() {
    const dateInput = document.getElementById('shc-cts-ks-date');
    const donViSelect = document.getElementById('shc-cts-ks-don-vi');
    if (!dateInput) return;
    dateInput.addEventListener('change', () => { _shcCtsKsCurrent.date = dateInput.value; loadShcCtsKiemSoatDetail(); });
    donViSelect.addEventListener('change', () => { _shcCtsKsCurrent.donVi = donViSelect.value; loadShcCtsKiemSoatThongKe(); });
    loadShcCtsKiemSoatDetail();
}
window.initShcCtsKiemSoat = initShcCtsKiemSoat;

async function loadShcCtsKiemSoatDetail() {
    const dateInput = document.getElementById('shc-cts-ks-date');
    const donViSelect = document.getElementById('shc-cts-ks-don-vi');
    const chiTietEl = document.getElementById('shc-cts-ks-chitiet');
    const pillEl = document.getElementById('shc-cts-ks-source-pill');
    if (!chiTietEl) return;
    chiTietEl.innerHTML = '<div class="loading"><i class="fas fa-spinner"></i><br>Đang tải...</div>';

    const params = new URLSearchParams();
    if (_shcCtsKsCurrent.date) params.set('date', _shcCtsKsCurrent.date);
    try {
        const data = await API.getShcCtsKiemSoatDetail(params.toString());
        _shcCtsKsCurrent.sheets = data.sheets || {};

        if (!dateInput.value) dateInput.value = data.selected_date;
        _shcCtsKsCurrent.date = data.selected_date;

        const donViNames = Object.keys(_shcCtsKsCurrent.sheets);
        const prevDonVi = donViSelect.value;
        donViSelect.innerHTML = '<option value="">Tất cả</option>' +
            donViNames.map(n => `<option value="${_shcCtsKsEscape(n)}"${n === prevDonVi ? ' selected' : ''}>${_shcCtsKsEscape(n)}</option>`).join('');
        if (prevDonVi && donViNames.includes(prevDonVi)) {
            donViSelect.value = prevDonVi;
            _shcCtsKsCurrent.donVi = prevDonVi;
        }

        pillEl.innerHTML = data.is_today_live
            ? '<span class="shc-cts-ks-live-pill">Hôm nay (Excel trực tiếp)</span>'
            : '<span class="shc-cts-ks-live-pill db">Lịch sử (DB)</span>';

        renderShcCtsKiemSoatDetail();
        loadShcCtsKiemSoatThongKe();
    } catch (error) {
        chiTietEl.innerHTML = `<div class="error">Không tải được dữ liệu: ${_shcCtsKsEscape(error.message)}</div>`;
    }
}

function renderShcCtsKiemSoatDetail() {
    const container = document.getElementById('shc-cts-ks-chitiet');
    if (!container) return;
    const sheets = _shcCtsKsCurrent.sheets || {};
    const donViFilter = _shcCtsKsCurrent.donVi;
    const targetSheets = donViFilter ? { [donViFilter]: sheets[donViFilter] } : sheets;

    const parts = [];
    for (const [donVi, sheet] of Object.entries(targetSheets)) {
        const rows = (sheet && sheet.data) || [];
        const headers = SHC_CTS_KS_DISPLAY_COLS.map(c => `<th>${SHC_CTS_KS_DISPLAY_LABELS[c] || c}</th>`).join('');
        const body = rows.map(row => {
            const cells = SHC_CTS_KS_DISPLAY_COLS.map(c => `<td>${row[c] != null ? row[c] : ''}</td>`).join('');
            const nvkt = _shcCtsKsEscape(row.NVKT_DB);
            const noiDung = _shcCtsKsEscape(row.kiemsoat_noi_dung || '');
            return `<tr>${cells}<td class="shc-cts-ks-cell">
                <textarea class="shc-cts-ks-input" rows="2" data-nvkt="${nvkt}" data-don_vi="${_shcCtsKsEscape(row['Đơn vị'] || donVi)}">${noiDung}</textarea>
                <button class="shc-cts-ks-save-btn" onclick="saveShcCtsKiemSoatRow('${nvkt}')">Lưu</button>
                <span class="shc-cts-ks-badge" id="shc-cts-ks-status-${nvkt}">${_shcCtsKsBadge(row)}</span>
            </td></tr>`;
        }).join('');
        parts.push(`<h5>${_shcCtsKsEscape(donVi)}</h5>
            <div class="excel-table-card"><div class="excel-table-body" style="max-height:520px;overflow:auto;">
            <table class="excel-table"><thead><tr>${headers}<th>Kiểm soát</th></tr></thead>
            <tbody>${body || '<tr><td colspan="99">Không có NVKT.</td></tr>'}</tbody></table>
            </div></div>`);
    }
    container.innerHTML = parts.join('') || '<div>Không có dữ liệu.</div>';
}

async function saveShcCtsKiemSoatRow(nvktDb) {
    const textarea = document.querySelector(`.shc-cts-ks-input[data-nvkt="${CSS.escape(nvktDb)}"]`);
    if (!textarea) return;
    const statusEl = document.getElementById(`shc-cts-ks-status-${nvktDb}`);
    try {
        const result = await API.saveShcCtsKiemSoat({
            ngay_xu_ly: _shcCtsKsCurrent.date,
            nvkt_db: nvktDb,
            don_vi: textarea.dataset.don_vi,
            noi_dung: textarea.value,
        });
        if (!result || result.ok === false) throw new Error((result && result.error) || 'Lỗi');
        if (statusEl) {
            statusEl.innerHTML = _shcCtsKsBadge({
                kiemsoat_da_nhap: !!result.noi_dung,
                kiemsoat_thoi_diem: new Date().toISOString().replace('T', ' ').substring(0, 19),
                kiemsoat_nguoi_nhap: result.nguoi_nhap,
            });
        }
        loadShcCtsKiemSoatThongKe();
    } catch (error) {
        alert('Không lưu được: ' + error.message);
    }
}
window.saveShcCtsKiemSoatRow = saveShcCtsKiemSoatRow;

async function loadShcCtsKiemSoatThongKe() {
    const statsEl = document.getElementById('shc-cts-ks-stats');
    const lichSuEl = document.getElementById('shc-cts-ks-lich-su');
    if (!statsEl) return;
    const params = new URLSearchParams();
    if (_shcCtsKsCurrent.date) params.set('date', _shcCtsKsCurrent.date);
    if (_shcCtsKsCurrent.donVi) params.set('don_vi', _shcCtsKsCurrent.donVi);
    try {
        const data = await API.getShcCtsKiemSoatThongKe(params.toString());
        const s = data.summary || {};
        const card = (v, l) => `<div class="shc-cts-ks-card"><div class="v">${v}</div><div class="l">${l}</div></div>`;
        statsEl.innerHTML =
            card(s.tong_so ?? 0, 'Tổng số') +
            card(s.tong_dat ?? 0, 'Tổng đã đạt') +
            card(s.chua_dat ?? 0, 'Chưa đạt') +
            card((s.ty_le_dat ?? 0) + '%', 'Tỷ lệ đạt') +
            card(s.da_ks ?? 0, 'Đã kiểm soát') +
            card(s.chua_ks ?? 0, 'Chưa KS');

        const lichSu = data.lich_su || [];
        lichSuEl.innerHTML = lichSu.length
            ? `<div class="excel-table-card"><div class="excel-table-body" style="max-height:320px;overflow:auto;">
               <table class="excel-table"><thead><tr><th>Ngày</th><th>Tổng số</th><th>Tổng đã đạt</th><th>Đã KS</th></tr></thead>
               <tbody>${lichSu.map(r => `<tr><td>${r.ngay_xu_ly}</td><td>${r.tong_so}</td><td>${r.tong_dat}</td><td>${r.da_ks}</td></tr>`).join('')}</tbody>
               </table></div></div>`
            : '<div>Chưa có lịch sử.</div>';
    } catch (error) {
        statsEl.innerHTML = `<div class="error">${_shcCtsKsEscape(error.message)}</div>`;
    }
}

function exportShcCtsKiemSoat() {
    const params = new URLSearchParams();
    if (_shcCtsKsCurrent.date) params.set('date', _shcCtsKsCurrent.date);
    if (_shcCtsKsCurrent.donVi) params.set('don_vi', _shcCtsKsCurrent.donVi);
    window.location.href = '/download/shc-cts-kiemsoat-report' + (params.toString() ? '?' + params.toString() : '');
}
window.exportShcCtsKiemSoat = exportShcCtsKiemSoat;
```

- [ ] **Step 3: Xác minh cú pháp JS**

Run: `node --check static/js/pages/shc_cts.js`
Expected: không có output. Nếu không có node, bỏ qua.

- [ ] **Step 4: Xác minh page render 200**

Run:
```bash
python3 -c "
from dashboard import app
with app.test_client() as c:
    with c.session_transaction() as s: s['username']='t'
    print(c.get('/shc-cts').status_code)
"
```
Expected: `200`.

- [ ] **Step 5: Commit**

```bash
git add static/js/pages/shc_cts.js
git commit -m "shc_cts.js: thêm logic lịch sử tiến độ + kiểm soát tổ trưởng"
```

---

## Task 12: Doc-sync docs/08 và docs/04

**Files:**
- Modify: `docs/08-trang-thai-thuc-thi.md`
- Modify: `docs/04-mapping-route-va-du-lieu.md`

- [ ] **Step 1: Đọc cấu trúc hiện tại của 2 file**

Run (đọc để biết format):
- `docs/08-trang-thai-thuc-thi.md` — tìm dòng nói về `quality.page_shc_cts`.
- `docs/04-mapping-route-va-du-lieu.md` — tìm dòng mapping cho `shc-cts`.

- [ ] **Step 2: Cập nhật docs/08**

Trong `docs/08-trang-thai-thuc-thi.md`, ở mục cho `quality.page_shc_cts`, bổ sung ghi chú (giữ nguyên style file):

- Trạng thái `shc_cts`: "Đã bổ sung lịch sử tiến độ theo ngày (DB nội bộ `shc_cts.db`, bảng `shc_cts_tien_do`) + kiểm soát tổ trưởng (bảng `shc_cts_kiemsoat`)."
- Liệt kê 4 endpoint mới: `/api/shc-cts-kiemsoat/luu` (POST), `/api/shc-cts-kiemsoat/detail`, `/api/shc-cts-kiemsoat/thongke`, `/download/shc-cts-kiemsoat-report`.

- [ ] **Step 3: Cập nhật docs/04**

Trong `docs/04-mapping-route-va-du-lieu.md`, thêm dòng mapping cho 4 endpoint mới theo format cột `route ↔ ma_bao_cao ↔ ten_bang_du_lieu ↔ supports_date`:

| route | ma_bao_cao | ten_bang_du_lieu | supports_date |
| --- | --- | --- | --- |
| `/api/shc-cts-kiemsoat/detail` | — (Excel intraday + DB nội bộ) | `shc_cts_tien_do` + `shc_cts_kiemsoat` (`shc_cts.db`) | Có (query `date`) |
| `/api/shc-cts-kiemsoat/thongke` | — | `shc_cts_tien_do` + `shc_cts_kiemsoat` | Có |
| `/api/shc-cts-kiemsoat/luu` | — | `shc_cts_kiemsoat` | — |
| `/download/shc-cts-kiemsoat-report` | — | `shc_cts_tien_do` + `shc_cts_kiemsoat` | Có |

Ghi chú rõ: nguồn chính "hôm nay" = file Excel `Bao_cao_tien_trinh_YYYYMMDD.xlsx` (sheet `Theo NVKT`); ngày quá khứ = DB nội bộ snapshot.

- [ ] **Step 4: Commit**

```bash
git add docs/08-trang-thai-thuc-thi.md docs/04-mapping-route-va-du-lieu.md
git commit -m "docs: cập nhật docs/08 + docs/04 cho lịch sử + kiểm soát SHC CTS"
```

---

## Verification cuối cùng (sau tất cả task)

- [ ] **Full test suite pass:** `python3 -m pytest tests/ -q` — toàn bộ PASS.
- [ ] **Py compile tất cả file Python đã sửa:**
  ```bash
  python3 -m py_compile config.py blueprints/quality_routes.py scripts/sync_shc_cts_tien_do.py
  ```
- [ ] **Smoke test thủ công (dev server):**
  ```bash
  python3 dashboard.py
  ```
  Mở `/shc-cts`, kiểm tra: section "Lịch sử tiến độ & Kiểm soát tổ trưởng" hiện; datepicker mặc định = hôm nay; bảng NVKT có cột "Kiểm soát" editable; nhập + Lưu → badge "Đã KS"; đổi sang ngày cũ → badge "Lịch sử (DB)"; nút Xuất Excel tải file.
- [ ] **Cron:** thêm vào crontab của unit (tham khảo `scripts/sync_pttb_phieu.py`):
  ```
  0 * * * * DASHV4_UNIT_CODE=son_tay python3 /home/vtst/dashv4/scripts/sync_shc_cts_tien_do.py >> /home/vtst/dashv4/logs/shc_cts_sync.log 2>&1
  ```
