# PTTB Kiểm soát tổ trưởng Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Thêm tính năng ghi chú kiểm soát tổ trưởng cho trang /pttb, sao chép logic hệt /brcd với nguồn dữ liệu PTTB (`baoCaoPTTB.xlsx`).

**Architecture:** Bảng `pttb_kiemsoat` (annotation) + `pttb_phieu` (snapshot) trong cùng DB `brcd_kiemsoat.db`. Khóa chính `ma_thue_bao` (TEXT, unique trong snapshot). 4 endpoint API + cron script + UI section mới trong pttb.html.

**Tech Stack:** Python 3.10, Flask, pandas, sqlite3, openpyxl/xlsxwriter. Test: pytest. Verify: `python3 -m py_compile`.

## Global Constraints

- Tests: `python3 -m pytest tests/` (system pytest).
- Verify syntax: `python3 -m py_compile <file>`, `node --check static/js/pages/pttb.js`.
- No lint/typecheck configured.
- `runtime_limits` must import before pandas.
- Schema idempotent: `CREATE TABLE IF NOT EXISTS`.
- DB path: `BRCD_KIEMSOAT_DB_PATH` (cùng DB với BRCD kiêm soát).
- Test pattern: monkeypatch `BRCD_KIEMSOAT_DB_PATH` + reset `_brcd_kiemsoat_schema_ready_path = None` + monkeypatch `PTTB_SUMMARY_FILE`.
- Do NOT commit unless user asks.
- Comments tiếng Việt.

## File Structure

| File | Vai trò | Tạo/Sửa |
|---|---|---|
| `blueprints/operations_routes.py` | Schema, sync, 4 endpoints, helpers | Sửa |
| `scripts/sync_pttb_phieu.py` | Cron entry point | Tạo |
| `static/js/api.js` | 3 API methods mới | Sửa |
| `static/js/pages/pttb.js` | Kiêm soát rendering + thongke | Sửa |
| `templates/pages/pttb.html` | Section kiêm soát + thongke | Sửa |
| `tests/test_pttb_kiemsoat.py` | Tests cho PTTB kiêm soát | Tạo |
| `docs/04-mapping-route-va-du-lieu.md` | Cập nhật row `/pttb` | Sửa |
| `docs/08-trang-thai-thuc-thi.md` | Thêm section PTTB kiêm soát | Sửa |

## PTTB Display Columns

```python
PTTB_KIEMSOAT_DISPLAY_COLUMNS = [
    'ma_thue_bao',       # TEXT, unique key
    'ten_thuebao',
    'diachi_lapdat',
    'loaihinh_tb',
    'nhanvien_tiepthi',
    'doi_vt',
    'ten_kv',
    'ngayhen_den',
    'noidung_hen',
    'chitieu_tg',
    'gio_conlai',
    'trang_thai',
]
```

## PTTB Sheet Names

```python
PTTB_TEAM_SHEETS = ['ToKT_SonTay', 'ToKT_SuoiHai', 'ToKT_QuangOai', 'ToKT_PhucTho']
```

---

## Task 1: Schema + sync function + POST save (TDD)

**Files:**
- Modify: `blueprints/operations_routes.py`
- Create: `tests/test_pttb_kiemsoat.py`

**Interfaces:**
- Produces:
  - `_pttb_kiemsoat_write_connection()` — reuse `_brcd_kiemsoat_write_connection()` (same DB)
  - `_pttb_kiemsoat_read_connection()` — reuse `_brcd_kiemsoat_read_connection()`
  - `_sync_pttb_phieu_to_db() -> dict` — same contract as `_sync_brcd_phieu_to_db()`
  - `get_pttb_kiemsoat_map(ma_thue_bao_list) -> dict` — same pattern as `get_brcd_kiemsoat_map()`
  - `POST /api/pttb-kiemsoat/luu` — same pattern as BRCD POST

- [ ] **Step 1: Write tests**

Tạo `tests/test_pttb_kiemsoat.py`:

```python
import sqlite3
import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from dashboard import app
from blueprints import operations_routes


def _prepare_db(tmp_path, monkeypatch):
    db_path = tmp_path / 'brcd_kiemsoat.db'
    monkeypatch.setattr(operations_routes, 'BRCD_KIEMSOAT_DB_PATH', str(db_path))
    operations_routes._brcd_kiemsoat_schema_ready_path = None
    operations_routes._ensure_brcd_kiemsoat_schema()
    return db_path


def _logged_in_client():
    client = app.test_client()
    with client.session_transaction() as session:
        session['username'] = 'test-user'
    return client


_PTTB_COLUMNS = [
    'ma_thue_bao', 'ten_thuebao', 'diachi_lapdat', 'loaihinh_tb',
    'nhanvien_tiepthi', 'doi_vt', 'ten_kv', 'ngayhen_den',
    'noidung_hen', 'chitieu_tg', 'gio_conlai', 'trang_thai',
]


def _write_fake_pttb(tmp_path, monkeypatch, rows):
    """rows: list[dict] cho 1 sheet ToKT_SonTay."""
    excel_path = tmp_path / 'baoCaoPTTB.xlsx'
    df = pd.DataFrame(rows, columns=_PTTB_COLUMNS)
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='ToKT_SonTay', index=False)
    monkeypatch.setattr(operations_routes, 'PTTB_SUMMARY_FILE', str(excel_path))
    return excel_path


def _pttb_rows():
    return [
        {
            'ma_thue_bao': 'fbr00ec1h', 'ten_thuebao': 'Nguyễn Văn A',
            'diachi_lapdat': 'DC1', 'loaihinh_tb': 'Fiber',
            'nhanvien_tiepthi': 'NV1', 'doi_vt': 'ToKT_SonTay',
            'ten_kv': 'KV1', 'ngayhen_den': '2026-06-25',
            'noidung_hen': 'Hẹn lắp', 'chitieu_tg': 24,
            'gio_conlai': -50.0, 'trang_thai': 'Quá giờ',
        },
        {
            'ma_thue_bao': 'cam00ai4o', 'ten_thuebao': 'Trần Văn B',
            'diachi_lapdat': 'DC2', 'loaihinh_tb': 'MyTV',
            'nhanvien_tiepthi': 'NV2', 'doi_vt': 'ToKT_SonTay',
            'ten_kv': 'KV2', 'ngayhen_den': '2026-06-26',
            'noidung_hen': 'Hẹn sửa', 'chitieu_tg': 24,
            'gio_conlai': 5.0, 'trang_thai': 'Bình thường',
        },
    ]


# ---------------------------------------------------------------------------
# Task 1: Schema + sync + POST save
# ---------------------------------------------------------------------------


def test_ensure_schema_creates_pttb_tables(tmp_path, monkeypatch):
    db_path = _prepare_db(tmp_path, monkeypatch)

    conn = sqlite3.connect(db_path)
    cols_ks = {row[1] for row in conn.execute('PRAGMA table_info(pttb_kiemsoat)')}
    cols_phieu = {row[1] for row in conn.execute('PRAGMA table_info(pttb_phieu)')}
    conn.close()

    assert {'ma_thue_bao', 'loaihinh_tb', 'doi_vt', 'nhanvien_tiepthi',
            'noi_dung_kiem_soat', 'nguoi_nhap', 'thoi_diem_nhap',
            'thoi_diem_cap_nhat'} <= cols_ks
    assert {'ma_thue_bao', 'ten_thuebao', 'diachi_lapdat', 'loaihinh_tb',
            'nhanvien_tiepthi', 'doi_vt', 'ten_kv', 'ngayhen_den',
            'noidung_hen', 'chitieu_tg', 'gio_conlai', 'trang_thai',
            'sheet', 'first_seen', 'last_seen'} <= cols_phieu


def test_sync_pttb_first_time(tmp_path, monkeypatch):
    db_path = _prepare_db(tmp_path, monkeypatch)
    _write_fake_pttb(tmp_path, monkeypatch, _pttb_rows())

    result = operations_routes._sync_pttb_phieu_to_db()

    assert result == {'synced': 2, 'new': 2, 'updated': 0, 'skipped': 0, 'reason': ''}
    conn = sqlite3.connect(db_path)
    count = conn.execute('SELECT COUNT(*) FROM pttb_phieu').fetchone()[0]
    conn.close()
    assert count == 2


def test_sync_pttb_preserves_first_seen(tmp_path, monkeypatch):
    db_path = _prepare_db(tmp_path, monkeypatch)
    _write_fake_pttb(tmp_path, monkeypatch, _pttb_rows())

    operations_routes._sync_pttb_phieu_to_db()
    conn = sqlite3.connect(db_path)
    conn.execute("UPDATE pttb_phieu SET first_seen = '2020-01-01 00:00:00'")
    conn.commit()
    conn.close()

    operations_routes._sync_pttb_phieu_to_db()
    conn = sqlite3.connect(db_path)
    rows = {r[0]: (r[1], r[2]) for r in conn.execute(
        'SELECT ma_thue_bao, first_seen, last_seen FROM pttb_phieu'
    ).fetchall()}
    conn.close()
    for _, (first, last) in rows.items():
        assert first == '2020-01-01 00:00:00'
        assert last != '2020-01-01 00:00:00'


def test_sync_pttb_skips_when_excel_missing(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    monkeypatch.setattr(operations_routes, 'PTTB_SUMMARY_FILE',
                        str(tmp_path / 'khong_co.xlsx'))

    result = operations_routes._sync_pttb_phieu_to_db()
    assert result['skipped'] == 1
    assert result['reason'] == 'excel_missing'


def test_pttb_kiemsoat_luu_creates(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)

    response = _logged_in_client().post(
        '/api/pttb-kiemsoat/luu',
        json={
            'ma_thue_bao': 'fbr00ec1h',
            'loaihinh_tb': 'Fiber',
            'doi_vt': 'ToKT_SonTay',
            'nhanvien_tiepthi': 'NV1',
            'noi_dung': 'KH đi vắng hẹn 18h',
        },
    )
    assert response.status_code == 200
    body = response.get_json()
    assert body == {
        'ok': True,
        'ma_thue_bao': 'fbr00ec1h',
        'noi_dung': 'KH đi vắng hẹn 18h',
        'nguoi_nhap': 'test-user',
    }


def test_pttb_kiemsoat_luu_update_keeps_thoi_diem_nhap(tmp_path, monkeypatch):
    db_path = _prepare_db(tmp_path, monkeypatch)
    conn = sqlite3.connect(db_path)
    conn.execute(
        '''INSERT INTO pttb_kiemsoat
            (ma_thue_bao, noi_dung_kiem_soat, nguoi_nhap, thoi_diem_nhap, thoi_diem_cap_nhat)
            VALUES (?, ?, ?, ?, ?)''',
        ('fbr00ec1h', 'cũ', 'old', '2020-01-01 00:00:00', '2020-01-01 00:00:00'),
    )
    conn.commit()
    conn.close()

    response = _logged_in_client().post(
        '/api/pttb-kiemsoat/luu',
        json={'ma_thue_bao': 'fbr00ec1h', 'noi_dung': 'mới'},
    )
    assert response.status_code == 200
    conn = sqlite3.connect(db_path)
    row = conn.execute(
        'SELECT thoi_diem_nhap, thoi_diem_cap_nhat FROM pttb_kiemsoat WHERE ma_thue_bao = ?',
        ('fbr00ec1h',),
    ).fetchone()
    conn.close()
    assert row[0] == '2020-01-01 00:00:00'
    assert row[1] != '2020-01-01 00:00:00'


def test_pttb_kiemsoat_luu_empty_deletes(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    conn = sqlite3.connect(str(tmp_path / 'brcd_kiemsoat.db'))
    conn.execute(
        'INSERT INTO pttb_kiemsoat (ma_thue_bao, noi_dung_kiem_soat) VALUES (?, ?)',
        ('fbr00ec1h', 'cũ'),
    )
    conn.commit()
    conn.close()

    response = _logged_in_client().post(
        '/api/pttb-kiemsoat/luu',
        json={'ma_thue_bao': 'fbr00ec1h', 'noi_dung': ''},
    )
    assert response.status_code == 200
    conn = sqlite3.connect(str(tmp_path / 'brcd_kiemsoat.db'))
    count = conn.execute(
        'SELECT COUNT(*) FROM pttb_kiemsoat WHERE ma_thue_bao = ?', ('fbr00ec1h',)
    ).fetchone()[0]
    conn.close()
    assert count == 0
```

- [ ] **Step 2: Run tests, verify fail**

Run: `python3 -m pytest tests/test_pttb_kiemsoat.py -q`
Expected: FAIL (tables/endpoints don't exist).

- [ ] **Step 3: Add PTTB schema + constants + sync + POST**

Trong `blueprints/operations_routes.py`:

a) Thêm constants sau block BRCD kiêm soát:

```python
# ---------------------------------------------------------------------------
# Kiểm soát tổ trưởng (PTTB) — logic hệt BRCD, nguồn PTTB
# ---------------------------------------------------------------------------
PTTB_TEAM_SHEETS = ['ToKT_SonTay', 'ToKT_SuoiHai', 'ToKT_QuangOai', 'ToKT_PhucTho']
PTTB_KIEMSOAT_DISPLAY_COLUMNS = [
    'ma_thue_bao',
    'ten_thuebao',
    'diachi_lapdat',
    'loaihinh_tb',
    'nhanvien_tiepthi',
    'doi_vt',
    'ten_kv',
    'ngayhen_den',
    'noidung_hen',
    'chitieu_tg',
    'gio_conlai',
    'trang_thai',
]
PTTB_KIEMSOAT_NOI_DUNG_MAX = 2000
```

b) Mở rộng `_ensure_brcd_kiemsoat_schema()` — thêm CREATE TABLE cho `pttb_kiemsoat` và `pttb_phieu` (giống pattern BRCD nhưng khóa `ma_thue_bao` TEXT).

c) Thêm `get_pttb_kiemsoat_map(ma_thue_bao_list)` — same pattern as `get_brcd_kiemsoat_map()`.

d) Thêm `_pttb_phieu_upsert_sql`, `_pttb_phieu_row_to_params`, `_sync_pttb_phieu_to_db()` — same pattern as BRCD sync.

e) Thêm `POST /api/pttb-kiemsoat/luu` — same pattern as BRCD POST but key `ma_thue_bao` (TEXT).

- [ ] **Step 4: Verify syntax**

Run: `python3 -m py_compile blueprints/operations_routes.py`

- [ ] **Step 5: Run tests, verify pass**

Run: `python3 -m pytest tests/test_pttb_kiemsoat.py -q`
Expected: 8 passed.

- [ ] **Step 6: Run full suite**

Run: `python3 -m pytest tests/ -q`
Expected: all pass.

---

## Task 2: Detail endpoint + on-load sync

**Files:**
- Modify: `blueprints/operations_routes.py`
- Modify: `tests/test_pttb_kiemsoat.py` (append)

**Interfaces:**
- Produces: `_load_pttb_kiemsoat_df() -> DataFrame|None`, `GET /api/pttb-kiemsoat/detail`

- [ ] **Step 1: Append tests for detail + on-load**

```python
def test_pttb_detail_joins_annotation(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    _write_fake_pttb(tmp_path, monkeypatch, _pttb_rows())

    _logged_in_client().post(
        '/api/pttb-kiemsoat/luu',
        json={'ma_thue_bao': 'fbr00ec1h', 'doi_vt': 'ToKT_SonTay',
              'nhanvien_tiepthi': 'NV1', 'noi_dung': 'Đã liên hệ'},
    )

    response = _logged_in_client().get('/api/pttb-kiemsoat/detail')
    assert response.status_code == 200
    payload = response.get_json()
    assert 'ToKT_SonTay' in payload['sheets']
    rows = payload['sheets']['ToKT_SonTay']['data']
    by_id = {r['ma_thue_bao']: r for r in rows}
    assert by_id['fbr00ec1h']['kiemsoat_da_nhap'] is True
    assert by_id['fbr00ec1h']['kiemsoat_noi_dung'] == 'Đã liên hệ'
    assert by_id['cam00ai4o']['kiemsoat_da_nhap'] is False


def test_pttb_load_triggers_sync(tmp_path, monkeypatch):
    db_path = _prepare_db(tmp_path, monkeypatch)
    _write_fake_pttb(tmp_path, monkeypatch, _pttb_rows())

    conn = sqlite3.connect(db_path)
    assert conn.execute('SELECT COUNT(*) FROM pttb_phieu').fetchone()[0] == 0
    conn.close()

    _logged_in_client().get('/api/pttb-kiemsoat/detail')
    conn = sqlite3.connect(db_path)
    assert conn.execute('SELECT COUNT(*) FROM pttb_phieu').fetchone()[0] == 2
    conn.close()
```

- [ ] **Step 2: Implement `_load_pttb_kiemsoat_df()` + detail endpoint**

Same pattern as BRCD but:
- Đọc `PTTB_TEAM_SHEETS` thay vì `ToKT_*`
- Key `ma_thue_bao` (TEXT) thay vì `baohong_id` (INT)
- Join `get_pttb_kiemsoat_map()` thay vì `get_brcd_kiemsoat_map()`
- Numeric cols: `chitieu_tg`, `gio_conlai`

- [ ] **Step 3: Run tests + verify**

---

## Task 3: Thongke + lich_su endpoints

**Files:**
- Modify: `blueprints/operations_routes.py`
- Modify: `tests/test_pttb_kiemsoat.py` (append)

- [ ] **Step 1: Append tests**

```python
def test_pttb_thongke_returns_summary(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    _write_fake_pttb(tmp_path, monkeypatch, _pttb_rows())
    _logged_in_client().post(
        '/api/pttb-kiemsoat/luu',
        json={'ma_thue_bao': 'fbr00ec1h', 'doi_vt': 'ToKT_SonTay',
              'nhanvien_tiepthi': 'NV1', 'noi_dung': 'done'},
    )

    response = _logged_in_client().get('/api/pttb-kiemsoat/thongke')
    assert response.status_code == 200
    payload = response.get_json()
    assert payload['summary'] == {'total': 2, 'da_kiem_soat': 1, 'chua': 1, 'ty_le': 50.0}
    assert 'lich_su' in payload
```

- [ ] **Step 2: Implement `_apply_pttb_kiemsoat_filters()`, `_compute_pttb_lich_su()`, endpoint**

Same pattern as BRCD but:
- `gio_conlai` thay vì `giờ còn lại thực`
- `NHANVIEN_TIEPTHI` thay vì `NVKT`
- `ma_thue_bao` (TEXT) thay vì `baohong_id` (INT)
- Snapshot table: `pttb_phieu`, annotation table: `pttb_kiemsoat`

- [ ] **Step 3: Run tests + verify**

---

## Task 4: Excel report download

**Files:**
- Modify: `blueprints/operations_routes.py`
- Modify: `tests/test_pttb_kiemsoat.py` (append)

- [ ] **Step 1: Append test**

```python
def test_pttb_report_download_returns_xlsx(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    _write_fake_pttb(tmp_path, monkeypatch, _pttb_rows())
    _logged_in_client().post(
        '/api/pttb-kiemsoat/luu',
        json={'ma_thue_bao': 'fbr00ec1h', 'doi_vt': 'ToKT_SonTay',
              'nhanvien_tiepthi': 'NV1', 'noi_dung': 'done'},
    )

    response = _logged_in_client().get('/download/pttb-kiemsoat-report')
    assert response.status_code == 200
    assert 'spreadsheetml' in response.headers['Content-Type']
```

- [ ] **Step 2: Implement endpoint `GET /download/pttb-kiemsoat-report`**

Same pattern as BRCD report.

- [ ] **Step 3: Run tests + verify**

---

## Task 5: Cron script

**Files:**
- Create: `scripts/sync_pttb_phieu.py`
- Modify: `tests/test_pttb_kiemsoat.py` (append)

- [ ] **Step 1: Append subprocess test**

```python
def test_sync_pttb_script_runs(tmp_path, monkeypatch):
    import subprocess, os
    db_path = tmp_path / 'brcd_kiemsoat.db'
    excel_path = tmp_path / 'baoCaoPTTB.xlsx'
    df = pd.DataFrame(_pttb_rows(), columns=_PTTB_COLUMNS)
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='ToKT_SonTay', index=False)

    wrapper = tmp_path / 'run_sync.py'
    wrapper.write_text(f'''
import sys
sys.path.insert(0, {repr(str(Path(__file__).resolve().parents[1]))})
import runtime_limits
from blueprints import operations_routes
operations_routes.PTTB_SUMMARY_FILE = {repr(str(excel_path))}
operations_routes._brcd_kiemsoat_schema_ready_path = None
result = operations_routes._sync_pttb_phieu_to_db()
print(result)
''')
    env = {**os.environ, 'DASHV4_BRCD_KIEMSOAT_DB_PATH': str(db_path)}
    completed = subprocess.run(['python3', str(wrapper)], env=env,
                               capture_output=True, text=True, timeout=30)
    assert completed.returncode == 0, completed.stderr
    assert "'synced': 2" in completed.stdout
```

- [ ] **Step 2: Create `scripts/sync_pttb_phieu.py`**

```python
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
```

- [ ] **Step 3: Run tests + verify**

---

## Task 6: UI — pttb.html + pttb.js

**Files:**
- Modify: `templates/pages/pttb.html`
- Modify: `static/js/pages/pttb.js`
- Modify: `static/js/api.js`
- Modify: `tests/test_pttb_kiemsoat.py` (append)

- [ ] **Step 1: Append HTML test**

```python
def test_pttb_page_has_kiemsoat_section():
    response = _logged_in_client().get('/pttb')
    assert response.status_code == 200
    html = response.data.decode('utf-8')
    assert 'id="pttb-kiemsoat-section"' in html
    assert 'id="pttb-kiemsoat-filter-khoang"' in html
    brcd_js = Path(__file__).resolve().parents[1].joinpath(
        'static', 'js', 'pages', 'pttb.js').read_text('utf-8')
    assert 'pttb_kiemsoat' in brcd_js or 'pttb-kiemsoat' in brcd_js
```

- [ ] **Step 2: Add API methods to `static/js/api.js`**

```javascript
getPttbKiemSoatDetail: (query) => API.get(`/api/pttb-kiemsoat/detail${query ? '?' + query : ''}`),
getPttbKiemSoatThongKe: (query) => API.get(`/api/pttb-kiemsoat/thongke${query ? '?' + query : ''}`),
savePttbKiemSoat: (data) => API.post('/api/pttb-kiemsoat/luu', data),
```

- [ ] **Step 3: Add section HTML vào `pttb.html`**

Thêm section "Kiểm soát tổ trưởng PTTB" trước section "Chi tiết tồn PTTB các tổ" (trước line 19), với:
- Filter dropdowns (Nhóm giờ, Trạng thái, Tổ, Khoảng thời gian)
- Stat cards (Tổng tồn, Đã KS, Chưa KS, Tỉ lệ, Rời tồn đã KS, Rời tồn chưa KS)
- Detail table container
- Nút Xuất báo cáo Excel

- [ ] **Step 4: Add JS handlers vào `pttb.js`**

Thêm functions: `_pttbKiemSoatQuery()`, `loadPttbKiemSoatThongKe()`, `renderPttbKiemSoatStats()`, `renderPttbKiemSoatChiTiet()`, `reloadPttbKiemSoatThongKe()`, `exportPttbKiemSoat()`. Pattern giống BRCD.

- [ ] **Step 5: Run tests + verify JS syntax**

---

## Task 7: Doc-sync + final verification

**Files:**
- Modify: `docs/04-mapping-route-va-du-lieu.md`
- Modify: `docs/08-trang-thai-thuc-thi.md`

- [ ] **Step 1: Update docs/04** — thêm row `/pttb` với note kiêm soát.
- [ ] **Step 2: Update docs/08** — thêm section "Kiểm soát tổ trưởng tại /pttb".
- [ ] **Step 3: Final full test run**

Run: `python3 -m pytest tests/ -q`
Expected: all pass.

- [ ] **Step 4: Syntax check**

Run: `python3 -m py_compile blueprints/operations_routes.py scripts/sync_pttb_phieu.py` + `node --check static/js/pages/pttb.js`
