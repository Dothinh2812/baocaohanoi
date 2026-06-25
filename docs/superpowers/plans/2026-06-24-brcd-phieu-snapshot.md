# BRCD Phiếu Snapshot Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Snapshot toàn bộ phiếu BRCD từng xuất hiện trong Vũ trụ tổng (Excel) vào bảng `brcd_phieu` trong DB, để tra cứu lịch sử và đánh giá kiểm soát tổ trưởng kể cả sau khi phiếu đã rời tồn.

**Architecture:** Bổ sung bảng `brcd_phieu` (cùng file `brcd_kiemsoat.db` hiện có) + function `_sync_brcd_kiemsoat_to_db()` upsert Vũ trụ tổng. Hai entry point: on-load trong `_load_brcd_kiemsoat_df()` + script cron `scripts/sync_brcd_phieu.py`. Mở rộng endpoint `/api/brcd-kiemsoat/thongke` trả thêm section `lich_su` với 2 metric "phiếu rời tồn đã/chưa kiểm soát".

**Tech Stack:** Python 3.10, Flask, pandas, sqlite3 (built-in), openpyxl/xlsxwriter. Test: pytest (system). Không có lint/typecheck — verify bằng `python3 -m py_compile`.

## Global Constraints

- **Tests run via system pytest:** `python3 -m pytest tests/` (KHÔNG dùng `venv/bin/pytest` — nó không có pytest).
- **Single test:** `python3 -m pytest tests/test_brcd_phieu_snapshot.py -q`.
- **Verify Python syntax:** `python3 -m py_compile <file>`.
- **No lint/typecheck configured** — đừng tạo ra.
- **Comment style:** Code hiện tại dùng comments tiếng Việt để giải thích. Tuân thủ pattern này.
- **Import ordering:** `runtime_limits` phải import TRƯỚC pandas. Mọi entrypoint mới phải follow.
- **Schema idempotent:** `CREATE TABLE IF NOT EXISTS` — không destructive migration.
- **DB path:** cùng file `BRCD_KIEMSOAT_DB_PATH` (env `DASHV4_BRCD_KIEMSOAT_DB_PATH`, default `runtime_app/<unit>/brcd_kiemsoat.db`).
- **Test pattern:** `monkeypatch.setattr(operations_routes, 'BRCD_KIEMSOAT_DB_PATH', str(db_path))` + reset `_brcd_kiemsoat_schema_ready_path = None`. Xem `tests/test_brcd_kiemsoat.py:14-19`.
- **KHÔNG commit** trừ khi user yêu cầu.

---

## File Structure

| File | Vai trò | Tạo/Sửa |
|---|---|---|
| `blueprints/operations_routes.py` | Thêm schema `brcd_phieu` vào `_ensure_brcd_kiemsoat_schema()`, thêm `_sync_brcd_phieu_to_db()`, wire on-load, mở rộng `/thongke` | Sửa |
| `scripts/sync_brcd_phieu.py` | Cron entry point, import trực tiếp function (không qua HTTP) | Tạo |
| `tests/test_brcd_phieu_snapshot.py` | Test sync function + on-load + script + thongke lich_su | Tạo |
| `templates/pages/brcd.html` | Thêm dropdown "Khoảng thời gian" + 2 thẻ stat mới | Sửa |
| `static/js/pages/brcd.js` | Gửi `khoang` param + render 2 thẻ mới | Sửa |
| `docs/04-mapping-route-va-du-lieu.md` | Cập nhật row `/brcd` với note snapshot | Sửa |
| `docs/08-trang-thai-thuc-thi.md` | Bổ sung subsection snapshot trong section 6 | Sửa |

---

## Task 1: Thêm schema `brcd_phieu` và helper sync (TDD)

**Files:**
- Modify: `blueprints/operations_routes.py:100-124` (mở rộng `_ensure_brcd_kiemsoat_schema`)
- Modify: `blueprints/operations_routes.py` (thêm `_BRCD_PHIEU_UPSERT_SQL`, `_BRCD_PHIEU_CREATE_SQL`, `_row_to_phieu_params`, `_sync_brcd_phieu_to_db`)
- Test: `tests/test_brcd_phieu_snapshot.py` (tạo file)

**Interfaces:**
- Consumes: `_brcd_kiemsoat_write_connection()`, `_ensure_brcd_kiemsoat_schema()`, `read_excel_sheet_cached()`, `BRCD_DETAIL_MAIN_FILE`, `BRCD_KIEMSOAT_DISPLAY_COLUMNS`, `BRCD_KIEMSOAT_DB_PATH` (đã có sẵn)
- Produces:
  - `_sync_brcd_phieu_to_db() -> dict` — returns `{'synced': int, 'new': int, 'updated': int, 'skipped': int, 'reason': str}`. Idempotent, không raise khi Excel thiếu.

- [ ] **Step 1: Viết test schema mới**

Tạo `tests/test_brcd_phieu_snapshot.py`:

```python
import io
import os
import sqlite3
import sys
from datetime import date, timedelta
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


_DETAIL_COLUMNS = [
    'baohong_id', 'ma_tb', 'TEN_TB', 'DIACHI_LD', 'LOAIHINH_TB', 'GHICHU_HONG',
    'NVKT', 'DOI_VT', 'ngay_bh', 'Trạng thái cổng', 'ttvt_ton', 'chitieu_tg',
    'thời gian tồn thực', 'giờ còn lại thực',
]


def _write_fake_brcd_detail(tmp_path, monkeypatch, rows, sheet_name='ToKT_SonTay'):
    """rows: list[dict] cho 1 sheet ToKT_<...>."""
    excel_path = tmp_path / 'chiTietBrcd5Doi.xlsx'
    df = pd.DataFrame(rows, columns=_DETAIL_COLUMNS)
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        df.drop(columns=['baohong_id']).to_excel(writer, sheet_name=sheet_name + '_rut_gon', index=False)
    monkeypatch.setattr(operations_routes, 'BRCD_DETAIL_MAIN_FILE', str(excel_path))
    return excel_path


def _detail_rows():
    return [
        {
            'baohong_id': 11795318, 'ma_tb': 'tb1', 'TEN_TB': 'NV A', 'DIACHI_LD': 'DC1',
            'LOAIHINH_TB': 'Fiber', 'GHICHU_HONG': 'hong1', 'NVKT': 'NV1', 'DOI_VT': 'ToKT_SonTay',
            'ngay_bh': '2026-06-23 10:59', 'Trạng thái cổng': 'ON', 'ttvt_ton': 'KH đi vắng',
            'chitieu_tg': 8, 'thời gian tồn thực': 21.6, 'giờ còn lại thực': -13.6,
        },
        {
            'baohong_id': 11796115, 'ma_tb': 'tb2', 'TEN_TB': 'NV B', 'DIACHI_LD': 'DC2',
            'LOAIHINH_TB': 'MyTV', 'GHICHU_HONG': 'hong2', 'NVKT': 'NV2', 'DOI_VT': 'ToKT_SonTay',
            'ngay_bh': '2026-06-23 17:41', 'Trạng thái cổng': 'OFF', 'ttvt_ton': 'Đứt cáp',
            'chitieu_tg': 8, 'thời gian tồn thực': 4.0, 'giờ còn lại thực': 4.0,
        },
    ]


# ---------------------------------------------------------------------------
# Task 1: schema + sync function
# ---------------------------------------------------------------------------


def test_ensure_schema_creates_brcd_phieu_table(tmp_path, monkeypatch):
    db_path = _prepare_db(tmp_path, monkeypatch)

    conn = sqlite3.connect(db_path)
    cols = {row[1] for row in conn.execute('PRAGMA table_info(brcd_phieu)')}
    conn.close()

    assert {'baohong_id', 'ma_tb', 'ten_tb', 'diachi_ld', 'loaihinh_tb',
            'ghichu_hong', 'nvkt', 'doi_vt', 'ngay_bh', 'trang_thai_cong',
            'ttvt_ton', 'chitieu_tg', 'thoi_gian_ton_thuc', 'gio_con_lai_thuc',
            'sa', 'sheet', 'first_seen', 'last_seen'} <= cols


def test_ensure_schema_is_idempotent(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)

    operations_routes._ensure_brcd_kiemsoat_schema()
    operations_routes._ensure_brcd_kiemsoat_schema()


def test_sync_first_time_inserts_with_first_seen_equal_last_seen(tmp_path, monkeypatch):
    db_path = _prepare_db(tmp_path, monkeypatch)
    _write_fake_brcd_detail(tmp_path, monkeypatch, _detail_rows())

    result = operations_routes._sync_brcd_phieu_to_db()

    assert result == {'synced': 2, 'new': 2, 'updated': 0, 'skipped': 0, 'reason': ''}

    conn = sqlite3.connect(db_path)
    rows = {r[0]: r for r in conn.execute(
        'SELECT baohong_id, ma_tb, nvkt, doi_vt, sheet, first_seen, last_seen FROM brcd_phieu'
    ).fetchall()}
    conn.close()

    assert set(rows.keys()) == {11795318, 11796115}
    for baohong_id, (_, ma_tb, nvkt, doi_vt, sheet, first_seen, last_seen) in rows.items():
        assert first_seen == last_seen  # lần đầu: first_seen = last_seen
        assert first_seen is not None
        assert sheet == 'ToKT_SonTay'


def test_sync_second_time_preserves_first_seen_updates_last_seen(tmp_path, monkeypatch):
    db_path = _prepare_db(tmp_path, monkeypatch)
    _write_fake_brcd_detail(tmp_path, monkeypatch, _detail_rows())

    operations_routes._sync_brcd_phieu_to_db()

    # Fake time: set first_seen thủ công về quá khứ
    conn = sqlite3.connect(db_path)
    conn.execute("UPDATE brcd_phieu SET first_seen = '2020-01-01 00:00:00'")
    conn.commit()
    conn.close()

    # Sync lại
    result = operations_routes._sync_brcd_phieu_to_db()
    assert result == {'synced': 2, 'new': 0, 'updated': 2, 'skipped': 0, 'reason': ''}

    conn = sqlite3.connect(db_path)
    rows = {r[0]: (r[1], r[2]) for r in conn.execute(
        'SELECT baohong_id, first_seen, last_seen FROM brcd_phieu'
    ).fetchall()}
    conn.close()

    for baohong_id, (first_seen, last_seen) in rows.items():
        assert first_seen == '2020-01-01 00:00:00'  # giữ nguyên
        assert last_seen != '2020-01-01 00:00:00'   # cập nhật mới
        assert last_seen.startswith('20')  # ISO timestamp


def test_sync_skips_when_excel_missing(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    monkeypatch.setattr(operations_routes, 'BRCD_DETAIL_MAIN_FILE',
                        str(tmp_path / 'khong_co.xlsx'))

    result = operations_routes._sync_brcd_phieu_to_db()

    assert result == {'synced': 0, 'new': 0, 'updated': 0, 'skipped': 1, 'reason': 'excel_missing'}


def test_sync_keeps_phieu_that_left_universe(tmp_path, monkeypatch):
    """Phiếu biến mất khỏi Excel → dòng DB vẫn giữ, không bị xóa."""
    db_path = _prepare_db(tmp_path, monkeypatch)
    _write_fake_brcd_detail(tmp_path, monkeypatch, _detail_rows())  # 2 phiếu

    operations_routes._sync_brcd_phieu_to_db()

    # Lần 2: Excel chỉ còn 1 phiếu (phiếu 11796115 đã xử lý xong)
    _write_fake_brcd_detail(tmp_path, monkeypatch, [_detail_rows()[0]])

    result = operations_routes._sync_brcd_phieu_to_db()
    assert result == {'synced': 1, 'new': 0, 'updated': 1, 'skipped': 0, 'reason': ''}

    conn = sqlite3.connect(db_path)
    count = conn.execute('SELECT COUNT(*) FROM brcd_phieu').fetchone()[0]
    ids = {r[0] for r in conn.execute('SELECT baohong_id FROM brcd_phieu').fetchall()}
    conn.close()

    assert count == 2  # vẫn 2 dòng, không xóa
    assert ids == {11795318, 11796115}  # phiếu đã rời tồn vẫn còn trong DB


def test_sync_updates_field_changes(tmp_path, monkeypatch):
    """Khi Excel update NVKT của 1 phiếu, snapshot cũng update."""
    db_path = _prepare_db(tmp_path, monkeypatch)
    _write_fake_brcd_detail(tmp_path, monkeypatch, _detail_rows())
    operations_routes._sync_brcd_phieu_to_db()

    # Excel giờ đổi NVKT của phiếu 11795318 từ NV1 → NV_NEW
    new_rows = _detail_rows()
    new_rows[0] = {**new_rows[0], 'NVKT': 'NV_NEW'}
    _write_fake_brcd_detail(tmp_path, monkeypatch, new_rows)

    operations_routes._sync_brcd_phieu_to_db()

    conn = sqlite3.connect(db_path)
    nvkt = conn.execute(
        'SELECT nvkt FROM brcd_phieu WHERE baohong_id = 11795318'
    ).fetchone()[0]
    conn.close()
    assert nvkt == 'NV_NEW'
```

- [ ] **Step 2: Chạy test để verify fail (function chưa tồn tại)**

Run: `python3 -m pytest tests/test_brcd_phieu_snapshot.py -q`
Expected: FAIL với `AttributeError: module 'blueprints.operations_routes' has no attribute '_sync_brcd_phieu_to_db'` và schema test fail vì `no such table: brcd_phieu`.

- [ ] **Step 3: Thêm schema `brcd_phieu` vào `_ensure_brcd_kiemsoat_schema`**

Trong `blueprints/operations_routes.py`, tìm block `with _brcd_kiemsoat_write_connection() as conn:` bên trong `_ensure_brcd_kiemsoat_schema()` (khoảng dòng 109-123). Thêm sau `CREATE TABLE IF NOT EXISTS brcd_kiemsoat` block:

```python
            conn.execute(
                '''
                CREATE TABLE IF NOT EXISTS brcd_phieu (
                    baohong_id         INTEGER PRIMARY KEY,
                    ma_tb              TEXT,
                    ten_tb             TEXT,
                    diachi_ld          TEXT,
                    loaihinh_tb        TEXT,
                    ghichu_hong        TEXT,
                    nvkt               TEXT,
                    doi_vt             TEXT,
                    ngay_bh            TEXT,
                    trang_thai_cong    TEXT,
                    ttvt_ton           TEXT,
                    chitieu_tg         REAL,
                    thoi_gian_ton_thuc REAL,
                    gio_con_lai_thuc   REAL,
                    sa                 TEXT,
                    sheet              TEXT,
                    first_seen         TEXT,
                    last_seen          TEXT
                )
                '''
            )
            conn.execute(
                'CREATE INDEX IF NOT EXISTS idx_brcd_phieu_last_seen ON brcd_phieu(last_seen)'
            )
            conn.execute(
                'CREATE INDEX IF NOT EXISTS idx_brcd_phieu_doi_vt ON brcd_phieu(doi_vt)'
            )
```

- [ ] **Step 4: Thêm UPSERT SQL constants + helpers + sync function**

Trong `blueprints/operations_routes.py`, thêm ngay sau hàm `get_brcd_kiemsoat_map` (khoảng dòng 145):

```python
# ---------------------------------------------------------------------------
# Snapshot Vũ trụ tổng vào brcd_phieu (lịch sử phiếu tồn)
# ---------------------------------------------------------------------------

_BRCD_PHIEU_UPSERT_SQL = """
    INSERT INTO brcd_phieu (
        baohong_id, ma_tb, ten_tb, diachi_ld, loaihinh_tb, ghichu_hong,
        nvkt, doi_vt, ngay_bh, trang_thai_cong, ttvt_ton,
        chitieu_tg, thoi_gian_ton_thuc, gio_con_lai_thuc, sa,
        sheet, first_seen, last_seen
    )
    VALUES (
        ?, ?, ?, ?, ?, ?,
        ?, ?, ?, ?, ?,
        ?, ?, ?, ?,
        ?, ?, ?
    )
    ON CONFLICT(baohong_id) DO UPDATE SET
        ma_tb = excluded.ma_tb,
        ten_tb = excluded.ten_tb,
        diachi_ld = excluded.diachi_ld,
        loaihinh_tb = excluded.loaihinh_tb,
        ghichu_hong = excluded.ghichu_hong,
        nvkt = excluded.nvkt,
        doi_vt = excluded.doi_vt,
        ngay_bh = excluded.ngay_bh,
        trang_thai_cong = excluded.trang_thai_cong,
        ttvt_ton = excluded.ttvt_ton,
        chitieu_tg = excluded.chitieu_tg,
        thoi_gian_ton_thuc = excluded.thoi_gian_ton_thuc,
        gio_con_lai_thuc = excluded.gio_con_lai_thuc,
        sa = excluded.sa,
        sheet = excluded.sheet,
        last_seen = excluded.last_seen
"""


def _phieu_row_to_params(row, sheet, now_iso):
    """Map 1 dòng Excel (Series) sang tuple params cho UPSERT."""
    def _val(col):
        v = row.get(col)
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return None
        return v
    return (
        int(row['baohong_id']),
        _val('ma_tb'),
        _val('TEN_TB'),
        _val('DIACHI_LD'),
        _val('LOAIHINH_TB'),
        _val('GHICHU_HONG'),
        _val('NVKT'),
        _val('DOI_VT'),
        _val('ngay_bh'),
        _val('Trạng thái cổng'),
        _val('ttvt_ton'),
        _val('chitieu_tg'),
        _val('thời gian tồn thực'),
        _val('giờ còn lại thực'),
        _val('SA'),
        sheet,
        now_iso,  # first_seen (chỉ tác dụng khi INSERT; UPSERT bỏ qua trên UPDATE)
        now_iso,  # last_seen (luôn update)
    )


def _sync_brcd_phieu_to_db():
    """Đọc Vũ trụ tổng hiện tại (Excel), upsert vào brcd_phieu.

    Idempotent. Không raise khi Excel thiếu hoặc DB lock — trả dict với skipped=1.
    Trả: {'synced': N, 'new': M, 'updated': K, 'skipped': 0|1, 'reason': str}
    """
    if not os.path.exists(BRCD_DETAIL_MAIN_FILE):
        return {'synced': 0, 'new': 0, 'updated': 0, 'skipped': 1, 'reason': 'excel_missing'}
    _ensure_brcd_kiemsoat_schema()

    try:
        all_sheets = pd.ExcelFile(BRCD_DETAIL_MAIN_FILE).sheet_names
        team_sheets = [s for s in all_sheets
                       if s.startswith('ToKT_') and not s.endswith('_rut_gon')]
    except Exception:
        return {'synced': 0, 'new': 0, 'updated': 0, 'skipped': 1, 'reason': 'excel_unreadable'}

    incoming = []  # list[tuple[sheet_name, DataFrame]]
    for sheet in team_sheets:
        df = read_excel_sheet_cached(BRCD_DETAIL_MAIN_FILE, sheet)
        cols = [c for c in BRCD_KIEMSOAT_DISPLAY_COLUMNS if c in df.columns]
        df = df[cols].copy()
        df['baohong_id'] = pd.to_numeric(df['baohong_id'], errors='coerce')
        df = df.dropna(subset=['baohong_id']).copy()
        df['baohong_id'] = df['baohong_id'].astype(int)
        if len(df):
            incoming.append((sheet, df))

    if not incoming:
        return {'synced': 0, 'new': 0, 'updated': 0, 'skipped': 0, 'reason': ''}

    all_ids = set()
    for _, df in incoming:
        all_ids.update(df['baohong_id'].astype(int).tolist())

    now_iso = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    try:
        with _brcd_kiemsoat_write_connection() as conn:
            # Tính new vs updated trước bằng 1 query IN
            placeholders = ', '.join('?' for _ in all_ids)
            existing_rows = conn.execute(
                f'SELECT baohong_id FROM brcd_phieu WHERE baohong_id IN ({placeholders})',
                list(all_ids),
            ).fetchall()
            existing_ids = {r['baohong_id'] for r in existing_rows}

            for sheet, df in incoming:
                for _, row in df.iterrows():
                    params = _phieu_row_to_params(row, sheet, now_iso)
                    conn.execute(_BRCD_PHIEU_UPSERT_SQL, params)

        new_count = len(all_ids - existing_ids)
        updated_count = len(all_ids & existing_ids)
        return {
            'synced': len(all_ids),
            'new': new_count,
            'updated': updated_count,
            'skipped': 0,
            'reason': '',
        }
    except sqlite3.OperationalError as e:
        msg = str(e).lower()
        if 'locked' in msg or 'busy' in msg:
            return {'synced': 0, 'new': 0, 'updated': 0, 'skipped': 1, 'reason': 'db_locked'}
        raise
```

- [ ] **Step 5: Verify Python syntax**

Run: `python3 -m py_compile blueprints/operations_routes.py`
Expected: no output (success).

- [ ] **Step 6: Chạy test Task 1, verify pass**

Run: `python3 -m pytest tests/test_brcd_phieu_snapshot.py -q`
Expected: 7 passed.

- [ ] **Step 7: Chạy full test suite để verify không regression**

Run: `python3 -m pytest tests/ -q`
Expected: tất cả pass (sẽ có 64 + 7 = 71 tests).

- [ ] **Step 8: Commit**

```bash
git add blueprints/operations_routes.py tests/test_brcd_phieu_snapshot.py
git commit -m "feat(brcd): thêm bảng brcd_phieu + _sync_brcd_phieu_to_db

Snapshot Vũ trụ tổng vào DB để tra cứu lịch sử phiếu sau khi rời tồn.
Upsert theo baohong_id; first_seen giữ nguyên qua các lần sync.
Bỏ qua im lặng khi Excel thiếu hoặc DB lock."
```

---

## Task 2: Wire on-load sync vào `_load_brcd_kiemsoat_df`

**Files:**
- Modify: `blueprints/operations_routes.py:679-683` (đầu `_load_brcd_kiemsoat_df`)
- Test: `tests/test_brcd_phieu_snapshot.py` (append)

**Interfaces:**
- Consumes: `_sync_brcd_phieu_to_db()` (từ Task 1)
- Produces: `_load_brcd_kiemsoat_df()` giờ có side-effect sync. Detail/thongke endpoint tự động trigger sync.

- [ ] **Step 1: Append test on-load sync**

Thêm vào cuối `tests/test_brcd_phieu_snapshot.py`:

```python
# ---------------------------------------------------------------------------
# Task 2: on-load sync
# ---------------------------------------------------------------------------


def test_load_brcd_kiemsoat_df_triggers_sync(tmp_path, monkeypatch):
    """Khi detail endpoint load, brcd_phieu tự sync từ Excel."""
    db_path = _prepare_db(tmp_path, monkeypatch)
    _write_fake_brcd_detail(tmp_path, monkeypatch, _detail_rows())

    # Trước khi gọi endpoint: brcd_phieu rỗng
    conn = sqlite3.connect(db_path)
    assert conn.execute('SELECT COUNT(*) FROM brcd_phieu').fetchone()[0] == 0
    conn.close()

    response = _logged_in_client().get('/api/brcd-kiemsoat/detail')

    assert response.status_code == 200
    conn = sqlite3.connect(db_path)
    assert conn.execute('SELECT COUNT(*) FROM brcd_phieu').fetchone()[0] == 2
    conn.close()


def test_load_brcd_kiemsoat_df_does_not_crash_when_excel_missing(tmp_path, monkeypatch):
    _prepare_db(tmp_path, monkeypatch)
    monkeypatch.setattr(operations_routes, 'BRCD_DETAIL_MAIN_FILE',
                        str(tmp_path / 'khong_co.xlsx'))

    # Detail endpoint trả 404 (Excel thiếu) — không crash
    response = _logged_in_client().get('/api/brcd-kiemsoat/detail')
    assert response.status_code == 404
```

- [ ] **Step 2: Run test, verify fail**

Run: `python3 -m pytest tests/test_brcd_phieu_snapshot.py::test_load_brcd_kiemsoat_df_triggers_sync -q`
Expected: FAIL với assertion `0 == 2` (vì sync chưa được wire).

- [ ] **Step 3: Wire sync vào `_load_brcd_kiemsoat_df`**

Trong `blueprints/operations_routes.py`, sửa đầu `_load_brcd_kiemsoat_df` (dòng 679-683) từ:

```python
def _load_brcd_kiemsoat_df():
    """Đọc sheet đầy đủ ToKT_<doi> + join annotation kiểm soát. Trả DataFrame hoặc None."""
    if not os.path.exists(BRCD_DETAIL_MAIN_FILE):
        return None
    _ensure_brcd_kiemsoat_schema()
```

Thành:

```python
def _load_brcd_kiemsoat_df():
    """Đọc sheet đầy đủ ToKT_<doi> + join annotation kiểm soát. Trả DataFrame hoặc None."""
    if not os.path.exists(BRCD_DETAIL_MAIN_FILE):
        return None
    _ensure_brcd_kiemsoat_schema()
    # Snapshot Vũ trụ tổng hiện tại vào brcd_phieu (best-effort, không raise).
    try:
        _sync_brcd_phieu_to_db()
    except Exception:
        app.logger.warning('brcd_phieu sync thất bại trong _load', exc_info=True)
```

- [ ] **Step 4: Verify syntax**

Run: `python3 -m py_compile blueprints/operations_routes.py`

- [ ] **Step 5: Run test, verify pass**

Run: `python3 -m pytest tests/test_brcd_phieu_snapshot.py -q`
Expected: 9 passed (7 từ Task 1 + 2 từ Task 2).

- [ ] **Step 6: Run full suite**

Run: `python3 -m pytest tests/ -q`
Expected: tất cả pass.

- [ ] **Step 7: Commit**

```bash
git add blueprints/operations_routes.py tests/test_brcd_phieu_snapshot.py
git commit -m "feat(brcd): wire _sync_brcd_phieu_to_db vào _load_brcd_kiemsoat_df

Mỗi lần user mở /brcd hoặc gọi API detail/thongke, Vũ trụ tổng
tự được snapshot vào brcd_phieu. Best-effort: exception bị bắt
và log, không crash request."
```

---

## Task 3: Script cron `scripts/sync_brcd_phieu.py`

**Files:**
- Create: `scripts/sync_brcd_phieu.py`
- Test: `tests/test_brcd_phieu_snapshot.py` (append)

**Interfaces:**
- Consumes: `_sync_brcd_phieu_to_db()` (từ Task 1), `config` (đọc `DASHV4_*` env)
- Produces: Executable script, exit 0 khi success/skipped-excel-missing, exit 1 khi lỗi khác.

- [ ] **Step 1: Append test script entry point**

Thêm vào cuối `tests/test_brcd_phieu_snapshot.py`:

```python
# ---------------------------------------------------------------------------
# Task 3: cron script
# ---------------------------------------------------------------------------


def test_sync_script_runs_via_subprocess(tmp_path, monkeypatch):
    """Script scripts/sync_brcd_phieu.py chạy được với env DASHV4_*."""
    import subprocess

    # Setup DB + Excel trong tmp dir, dùng 1 process con
    db_path = tmp_path / 'brcd_kiemsoat.db'
    excel_path = tmp_path / 'chiTietBrcd5Doi.xlsx'

    # Tạo Excel giả (2 phiếu)
    df = pd.DataFrame(_detail_rows(), columns=_DETAIL_COLUMNS)
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='ToKT_SonTay', index=False)

    # Tạo init script để seed config trước khi import operations_routes
    env = {
        **os.environ,
        'DASHV4_BRCD_KIEMSOAT_DB_PATH': str(db_path),
    }
    # Phải monkeypatch BRCD_DETAIL_MAIN_FILE — script không nhận env cho path này.
    # Dùng wrapper: tạo script tạm trong tmp_path gọi _sync_brcd_phieu_to_db
    # sau khi đã monkeypatch.
    wrapper = tmp_path / 'run_sync.py'
    wrapper.write_text(f'''
import sys
sys.path.insert(0, {repr(str(Path(__file__).resolve().parents[1]))})
import runtime_limits
from blueprints import operations_routes
operations_routes.BRCD_DETAIL_MAIN_FILE = {repr(str(excel_path))}
operations_routes._brcd_kiemsoat_schema_ready_path = None
result = operations_routes._sync_brcd_phieu_to_db()
print(result)
''')

    completed = subprocess.run(
        ['python3', str(wrapper)],
        env=env,
        capture_output=True,
        text=True,
        timeout=30,
    )

    assert completed.returncode == 0, completed.stderr
    assert "'synced': 2" in completed.stdout

    conn = sqlite3.connect(db_path)
    assert conn.execute('SELECT COUNT(*) FROM brcd_phieu').fetchone()[0] == 2
    conn.close()
```

**Lưu ý:** Nếu test trên phức tạp quá, có thể đơn giản hóa: chỉ test script có thể import và `_sync_brcd_phieu_to_db` chạy được khi DB+Excel đã setup sẵn (không subprocess). Subprocess test isolation tốt hơn nhưng chậm hơn — chọn 1 trong 2. Implementer có quyền đổi.

- [ ] **Step 2: Run test, verify fail**

Run: `python3 -m pytest tests/test_brcd_phieu_snapshot.py::test_sync_script_runs_via_subprocess -q`
Expected: FAIL với `FileNotFoundError` hoặc similar (script chưa tạo).

- [ ] **Step 3: Tạo script `scripts/sync_brcd_phieu.py`**

```python
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

# Phải import runtime_limits TRƯỚC pandas (theo AGENTS.md critical ordering).
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import runtime_limits  # noqa: F401, E402

from blueprints.operations_routes import _sync_brcd_phieu_to_db  # noqa: E402


def main():
    result = _sync_brcd_phieu_to_db()
    print(result)
    # excel_missing là điều kiện bình thường giữa các lần refresh, không phải lỗi
    if result.get('skipped') and result.get('reason') not in ('excel_missing', ''):
        sys.exit(1)


if __name__ == '__main__':
    main()
```

- [ ] **Step 4: Verify script syntax**

Run: `python3 -m py_compile scripts/sync_brcd_phieu.py`
Expected: no output.

- [ ] **Step 5: Run test, verify pass**

Run: `python3 -m pytest tests/test_brcd_phieu_snapshot.py -q`
Expected: 10 passed.

- [ ] **Step 6: Test script chạy thực bằng tay với db thật**

Run: `python3 scripts/sync_brcd_phieu.py`
Expected: in ra `{'synced': N, 'new': M, 'updated': K, 'skipped': 0, 'reason': ''}` (N = số phiếu trong chiTietBrcd5Doi.xlsx hiện tại). Nếu Excel chưa refresh: `{'synced': 0, ..., 'skipped': 1, 'reason': 'excel_missing'}`.

Verify DB có dữ liệu:
```bash
sqlite3 runtime_app/son_tay/brcd_kiemsoat.db "SELECT COUNT(*), MIN(first_seen), MAX(last_seen) FROM brcd_phieu;"
```

- [ ] **Step 7: Commit**

```bash
git add scripts/sync_brcd_phieu.py tests/test_brcd_phieu_snapshot.py
git commit -m "feat(brcd): thêm scripts/sync_brcd_phieu.py cho cron hằng giờ

Import trực tiếp _sync_brcd_phieu_to_db (không qua HTTP) để tránh
phức tạp auth/session. Exit 0 cho success hoặc excel_missing,
exit 1 cho lỗi khác."
```

---

## Task 4: Mở rộng `/api/brcd-kiemsoat/thongke` với section `lich_su`

**Files:**
- Modify: `blueprints/operations_routes.py:767-810` (endpoint `api_brcd_kiemsoat_thongke`)
- Modify: `blueprints/operations_routes.py` (thêm `_compute_lich_su` helper)
- Test: `tests/test_brcd_phieu_snapshot.py` (append)

**Interfaces:**
- Consumes: `_brcd_kiemsoat_read_connection()` (đã có), `_ensure_brcd_kiemsoat_schema()`, `brcd_phieu` table (từ Task 1)
- Produces: Endpoint `/api/brcd-kiemsoat/thongke` trả thêm key `lich_su: {tu_ngay, den_ngay, roi_da_ks, roi_chua_ks}`. Query param mới: `khoang` ∈ {`tuan_nay`, `thang_nay`, `nam_nay`, `tat_ca`} (default `thang_nay`).

- [ ] **Step 1: Append test cho lich_su**

Thêm vào cuối `tests/test_brcd_phieu_snapshot.py`:

```python
# ---------------------------------------------------------------------------
# Task 4: thongke lich_su
# ---------------------------------------------------------------------------


def _seed_phieu_with_history(tmp_path, monkeypatch, db_path):
    """Tạo kịch bản:
    - Phiếu 11795318: rời tồn trong tháng này, ĐÃ kiểm soát
    - Phiếu 11796115: rời tồn trong tháng này, CHƯA kiểm soát
    - Phiếu 11797000: vẫn còn trong tồn (current universe)
    """
    rows = _detail_rows() + [{
        'baohong_id': 11797000, 'ma_tb': 'tb3', 'TEN_TB': 'NV C', 'DIACHI_LD': 'DC3',
        'LOAIHINH_TB': 'Fiber', 'GHICHU_HONG': 'h3', 'NVKT': 'NV3', 'DOI_VT': 'ToKT_SonTay',
        'ngay_bh': '2026-06-24 08:00', 'Trạng thái cổng': 'ON', 'ttvt_ton': 'x',
        'chitieu_tg': 8, 'thời gian tồn thực': 1.0, 'giờ còn lại thực': 7.0,
    }]
    _write_fake_brcd_detail(tmp_path, monkeypatch, rows)

    # Trigger sync (tất cả 3 phiếu có first_seen=last_seen=now)
    operations_routes._sync_brcd_phieu_to_db()

    # Fake last_seen của 2 phiếu "rời tồn" về quá khứ (trong tháng này)
    today = date.today()
    past_iso = today.replace(day=max(1, today.day - 3)).strftime('%Y-%m-%d 10:00:00')
    conn = sqlite3.connect(db_path)
    conn.execute(
        "UPDATE brcd_phieu SET last_seen = ? WHERE baohong_id IN (?, ?)",
        (past_iso, 11795318, 11796115),
    )
    conn.commit()
    conn.close()

    # Excel "hiện tại" giờ chỉ còn phiếu 11797000 — mô phỏng 2 phiếu kia đã xử lý xong
    _write_fake_brcd_detail(tmp_path, monkeypatch, [rows[2]])

    # Trigger sync lại để cập nhật last_seen của 11797000 = now
    operations_routes._sync_brcd_phieu_to_db()

    # Add kiemsoat cho phiếu 11795318 (đã kiểm soát trước khi rời)
    _logged_in_client().post(
        '/api/brcd-kiemsoat/luu',
        json={'baohong_id': 11795318, 'doi_vt': 'ToKT_SonTay', 'nvkt': 'NV1', 'noi_dung': 'done'},
    )
    # Excel cần chứa 11795318 để POST luu ghi annotation được — viết lại Excel với cả 3
    _write_fake_brcd_detail(tmp_path, monkeypatch, rows)
    _logged_in_client().post(
        '/api/brcd-kiemsoat/luu',
        json={'baohong_id': 11795318, 'doi_vt': 'ToKT_SonTay', 'nvkt': 'NV1', 'noi_dung': 'done'},
    )

    # Cuối cùng: Excel chỉ còn 11797000 (mô phỏng real state)
    _write_fake_brcd_detail(tmp_path, monkeypatch, [rows[2]])


def test_thongke_returns_lich_su_default_thang_nay(tmp_path, monkeypatch):
    db_path = _prepare_db(tmp_path, monkeypatch)
    _seed_phieu_with_history(tmp_path, monkeypatch, db_path)

    response = _logged_in_client().get('/api/brcd-kiemsoat/thongke')

    assert response.status_code == 200
    payload = response.get_json()
    assert 'lich_su' in payload
    ls = payload['lich_su']
    assert ls['roi_da_ks'] == 1   # 11795318 đã KS rồi rời
    assert ls['roi_chua_ks'] == 1  # 11796115 rời mà chưa KS


def test_thongke_lich_su_filter_khoang_tat_ca(tmp_path, monkeypatch):
    db_path = _prepare_db(tmp_path, monkeypatch)
    _seed_phieu_with_history(tmp_path, monkeypatch, db_path)

    response = _logged_in_client().get('/api/brcd-kiemsoat/thongke?khoang=tat_ca')

    payload = response.get_json()
    ls = payload['lich_su']
    # tat_ca không giới hạn date → 2 phiếu rời tồn vẫn đếm
    assert ls['roi_da_ks'] + ls['roi_chua_ks'] == 2


def test_thongke_lich_su_filter_doi_applies(tmp_path, monkeypatch):
    db_path = _prepare_db(tmp_path, monkeypatch)
    _seed_phieu_with_history(tmp_path, monkeypatch, db_path)

    response = _logged_in_client().get('/api/brcd-kiemsoat/thongke?doi=ToKT_SonTay')

    payload = response.get_json()
    ls = payload['lich_su']
    assert ls['roi_da_ks'] + ls['roi_chua_ks'] == 2

    # Đội khác → 0
    response2 = _logged_in_client().get('/api/brcd-kiemsoat/thongke?doi=ToKT_PhucTho')
    ls2 = response2.get_json()['lich_su']
    assert ls2['roi_da_ks'] + ls2['roi_chua_ks'] == 0
```

- [ ] **Step 2: Run test, verify fail**

Run: `python3 -m pytest tests/test_brcd_phieu_snapshot.py -k lich_su -q`
Expected: FAIL vì payload chưa có `lich_su`.

- [ ] **Step 3: Thêm `_compute_lich_su` helper**

Trong `blueprints/operations_routes.py`, thêm ngay sau `_apply_brcd_kiemsoat_filters` (khoảng dòng 750):

```python
def _compute_lich_su(filtered_df, args):
    """Tính 'phiếu đã rời tồn' từ snapshot brcd_phieu trong khoảng thời gian.

    "Rời tồn" = có trong brcd_phieu (trong khoảng) nhưng KHÔNG có trong
    current universe (filtered_df). Filter doi/loaihinh áp dụng cho snapshot.
    """
    from datetime import date, timedelta

    khoang = (args.get('khoang') or 'thang_nay').strip()
    today = date.today()
    if khoang == 'tuan_nay':
        tu_ngay = today - timedelta(days=today.weekday())  # thứ 2
        den_ngay = today
    elif khoang == 'nam_nay':
        tu_ngay = today.replace(month=1, day=1)
        den_ngay = today
    elif khoang == 'tat_ca':
        tu_ngay = date(1970, 1, 1)
        den_ngay = today
    else:  # thang_nay (default)
        tu_ngay = today.replace(day=1)
        den_ngay = today

    # Current universe IDs (sau filter doi/loaihinh/nhom)
    current_ids = set()
    if filtered_df is not None and 'baohong_id' in filtered_df.columns:
        current_ids = set(filtered_df['baohong_id'].astype(int).tolist())

    sql = ("SELECT baohong_id FROM brcd_phieu "
           "WHERE DATE(last_seen) >= ? AND DATE(last_seen) <= ?")
    params = [tu_ngay.isoformat(), den_ngay.isoformat()]

    doi = args.get('doi')
    if doi:
        sql += " AND doi_vt = ?"
        params.append(doi)
    loaihinh = args.get('loaihinh')
    if loaihinh:
        sql += " AND loaihinh_tb = ?"
        params.append(loaihinh)

    _ensure_brcd_kiemsoat_schema()
    with _brcd_kiemsoat_read_connection() as conn:
        snap_rows = conn.execute(sql, params).fetchall()
        snap_ids = {r['baohong_id'] for r in snap_rows}

        ks_ids = set()
        if snap_ids:
            placeholders = ', '.join('?' for _ in snap_ids)
            ks_rows = conn.execute(
                f"SELECT baohong_id FROM brcd_kiemsoat "
                f"WHERE baohong_id IN ({placeholders}) "
                f"AND COALESCE(noi_dung_kiem_soat, '') != ''",
                list(snap_ids),
            ).fetchall()
            ks_ids = {r['baohong_id'] for r in ks_rows}

    roi_ids = snap_ids - current_ids
    roi_da_ks = len(roi_ids & ks_ids)
    roi_chua_ks = len(roi_ids - ks_ids)

    return {
        'tu_ngay': tu_ngay.isoformat(),
        'den_ngay': den_ngay.isoformat(),
        'roi_da_ks': roi_da_ks,
        'roi_chua_ks': roi_chua_ks,
    }
```

- [ ] **Step 4: Wire `lich_su` vào `api_brcd_kiemsoat_thongke`**

Sửa block `return jsonify({...})` cuối của `api_brcd_kiemsoat_thongke` (dòng 800-810) từ:

```python
    return jsonify({
        'summary': {
            'total': total,
            'da_kiem_soat': da,
            'chua': chua,
            'ty_le': ty_le,
        },
        'by_doi': _agg('DOI_VT'),
        'by_nvkt': _agg('NVKT'),
        'chi_tiet': chi_tiet,
    })
```

Thành:

```python
    return jsonify({
        'summary': {
            'total': total,
            'da_kiem_soat': da,
            'chua': chua,
            'ty_le': ty_le,
        },
        'by_doi': _agg('DOI_VT'),
        'by_nvkt': _agg('NVKT'),
        'chi_tiet': chi_tiet,
        'lich_su': _compute_lich_su(filtered, request.args),
    })
```

- [ ] **Step 5: Verify syntax**

Run: `python3 -m py_compile blueprints/operations_routes.py`

- [ ] **Step 6: Run test, verify pass**

Run: `python3 -m pytest tests/test_brcd_phieu_snapshot.py -q`
Expected: 13 passed (10 + 3).

- [ ] **Step 7: Run full suite**

Run: `python3 -m pytest tests/ -q`
Expected: tất cả pass.

- [ ] **Step 8: Commit**

```bash
git add blueprints/operations_routes.py tests/test_brcd_phieu_snapshot.py
git commit -m "feat(brcd): thongke trả thêm section lich_su

2 metric mới từ snapshot brcd_phieu: 'roi_da_ks' và 'roi_chua_ks'
(phiếu đã rời tồn trong khoảng thời gian, đã/chưa được kiểm soát).
Filter mới 'khoang' ∈ {tuan_nay, thang_nay, nam_nay, tat_ca},
default thang_nay. Filter doi/loaihinh áp dụng cho cả 2 metric."
```

---

## Task 5: UI — dropdown khoảng thời gian + 2 thẻ stat mới

**Files:**
- Modify: `templates/pages/brcd.html:78-102` (filters section)
- Modify: `static/js/pages/brcd.js:385-430` (`_kiemSoatQuery`, `renderKiemSoatStats`)
- Test: `tests/test_brcd_phieu_snapshot.py` (append HTML assertion)

**Interfaces:**
- Consumes: API `/api/brcd-kiemsoat/thongke` giờ trả `lich_su` (từ Task 4)
- Produces: UI section "Thống kê kiểm soát" có dropdown "Khoảng thời gian" + 2 thẻ mới "Rời tồn đã KS" / "Rời tồn chưa KS" (đỏ).

- [ ] **Step 1: Append HTML assertion test**

Thêm vào cuối `tests/test_brcd_phieu_snapshot.py`:

```python
# ---------------------------------------------------------------------------
# Task 5: UI
# ---------------------------------------------------------------------------


def test_brcd_page_has_khoang_dropdown_and_lich_su_cards():
    response = _logged_in_client().get('/brcd')

    assert response.status_code == 200
    html = response.data.decode('utf-8')
    # Dropdown khoảng thời gian
    assert 'id="kiemsoat-filter-khoang"' in html
    assert 'value="thang_nay"' in html
    assert 'value="tuan_nay"' in html
    assert 'value="nam_nay"' in html
    assert 'value="tat_ca"' in html

    # JS render 2 thẻ mới
    brcd_js = Path(__file__).resolve().parents[1].joinpath(
        'static', 'js', 'pages', 'brcd.js').read_text('utf-8')
    assert 'roi_da_ks' in brcd_js
    assert 'roi_chua_ks' in brcd_js
    assert 'Rời tồn' in brcd_js
    # _kiemSoatQuery gửi param khoang
    assert "'khoang'" in brcd_js or '"khoang"' in brcd_js
```

- [ ] **Step 2: Run test, verify fail**

Run: `python3 -m pytest tests/test_brcd_phieu_snapshot.py::test_brcd_page_has_khoang_dropdown_and_lich_su_cards -q`
Expected: FAIL vì dropdown chưa tồn tại.

- [ ] **Step 3: Thêm dropdown "Khoảng thời gian" vào HTML**

Trong `templates/pages/brcd.html`, tìm filter `<label>Đội:` block (khoảng dòng 93-97), thêm SAU label Đội và TRƯỚC `<button class="download-btn" onclick="exportBrcdKiemSoat()">`:

```html
            <label>Khoảng thời gian (lịch sử):
                <select id="kiemsoat-filter-khoang" onchange="reloadKiemSoatThongKe()">
                    <option value="tuan_nay">Tuần này</option>
                    <option value="thang_nay" selected>Tháng này</option>
                    <option value="nam_nay">Năm nay</option>
                    <option value="tat_ca">Tất cả</option>
                </select>
            </label>
```

- [ ] **Step 4: Mở rộng `_kiemSoatQuery` gửi `khoang`**

Trong `static/js/pages/brcd.js`, tìm function `_kiemSoatQuery` (dòng 385-394) và sửa thành:

```javascript
function _kiemSoatQuery() {
    const params = new URLSearchParams();
    const nhom = document.getElementById('kiemsoat-filter-nhom');
    const trangthai = document.getElementById('kiemsoat-filter-trangthai');
    const doi = document.getElementById('kiemsoat-filter-doi');
    const khoang = document.getElementById('kiemsoat-filter-khoang');
    if (nhom && nhom.value) params.set('nhom', nhom.value);
    if (trangthai && trangthai.value) params.set('trangthai', trangthai.value);
    if (doi && doi.value) params.set('doi', doi.value);
    if (khoang && khoang.value) params.set('khoang', khoang.value);
    return params.toString();
}
```

- [ ] **Step 5: Thêm 2 thẻ stat mới vào `renderKiemSoatStats`**

Sửa `renderKiemSoatStats` (dòng 409-430), block `statsEl.innerHTML = ...`:

```javascript
function renderKiemSoatStats(data, statsEl, byDoiEl) {
    if (!statsEl) return;
    const s = data.summary || {};
    const ls = data.lich_su || {};
    const card = (num, label, cls) => `<div class="ks-card ${cls || ''}"><div class="ks-num">${num}</div><div class="ks-label">${label}</div></div>`;
    statsEl.innerHTML = `
        <div class="kiemsoat-stats-cards">
            ${card(s.total ?? 0, 'Tổng tồn', 'ks-total')}
            ${card(s.da_kiem_soat ?? 0, 'Đã kiểm soát', 'ks-da-card')}
            ${card(s.chua ?? 0, 'Chưa kiểm soát', 'ks-chua-card')}
            ${card((s.ty_le ?? 0) + '%', 'Tỉ lệ', 'ks-ty-le')}
            ${card(ls.roi_da_ks ?? 0, 'Rời tồn đã KS', 'ks-roi-da')}
            ${card(ls.roi_chua_ks ?? 0, 'Rời tồn chưa KS', 'ks-roi-chua')}
        </div>`;

    if (byDoiEl) {
        const rows = (data.by_doi || []);
        if (rows.length === 0) { byDoiEl.innerHTML = ''; return; }
        byDoiEl.innerHTML = `
            <table class="excel-table summary-table">
                <thead><tr><th>Đội</th><th>Tổng</th><th>Đã KS</th><th>Chưa</th><th>Tỉ lệ</th></tr></thead>
                <tbody>${rows.map(r => `<tr><td>${r.DOI_VT || ''}</td><td>${r.total}</td><td>${r.da_kiem_soat}</td><td>${r.chua}</td><td>${r.ty_le}%</td></tr>`).join('')}</tbody>
            </table>`;
    }
}
```

- [ ] **Step 6: Thêm CSS cho 2 thẻ mới (đỏ cho chưa KS)**

Trong `templates/pages/brcd.html`, tìm block `<style>` ở đầu file (dòng 7-32区域). Thêm vào cuối block style:

```css
        .ks-card.ks-roi-da { background: #e8f5e9; border-left: 4px solid #2e7d32; }
        .ks-card.ks-roi-chua { background: #ffebee; border-left: 4px solid #c62828; }
```

(Nếu `.ks-card.ks-da-card` đã có style tương tự, copy pattern.)

- [ ] **Step 7: Run test, verify pass**

Run: `python3 -m pytest tests/test_brcd_phieu_snapshot.py -q`
Expected: 14 passed (13 + 1).

- [ ] **Step 8: Run full suite**

Run: `python3 -m pytest tests/ -q`
Expected: tất cả pass (78 tests = 64 cũ + 14 mới).

- [ ] **Step 9: Verify JS syntax**

Run: `node --check static/js/pages/brcd.js`
Expected: no output.

- [ ] **Step 10: Commit**

```bash
git add templates/pages/brcd.html static/js/pages/brcd.js tests/test_brcd_phieu_snapshot.py
git commit -m "feat(brcd): UI hiển thị 2 thẻ 'Rời tồn đã/chưa KS' + dropdown khoảng

Mở section 'Thống kê kiểm soát' với dropdown Khoảng thời gian
(tuần/tháng/năm/tất cả) và 2 thẻ stat mới từ snapshot brcd_phieu.
Thẻ 'Rời tồn chưa KS' màu đỏ để nhấn mạnh mất kiểm soát."
```

---

## Task 6: Doc-sync + final verification

**Files:**
- Modify: `docs/04-mapping-route-va-du-lieu.md` (row `/brcd`)
- Modify: `docs/08-trang-thai-thuc-thi.md` (section 6)

**Interfaces:**
- Consumes: full implementation từ Tasks 1-5.

- [ ] **Step 1: Cập nhật `docs/04`**

Mở `docs/04-mapping-route-va-du-lieu.md`, tìm row `/brcd`. Thêm note: "Đã có snapshot `brcd_phieu` hằng giờ qua `scripts/sync_brcd_phieu.py` + on-load sync. Endpoint `/api/brcd-kiemsoat/thongke` trả thêm `lich_su` với 2 metric 'rời tồn đã/chưa KS'."

(Implementation step: đọc file, tìm row, thêm note vào cột ghi chú.)

- [ ] **Step 2: Cập nhật `docs/08`**

Mở `docs/08-trang-thai-thuc-thi.md`, tìm section 6 "Kiểm soát tổ trưởng tại /brcd". Thêm subsection 6.1 "Snapshot lịch sử":

```markdown
### 6.1. Snapshot lịch sử phiếu (`brcd_phieu`)

- **Mục đích:** Tra cứu phiếu đã rời tồn; đánh giá kiểm soát tổ trưởng trên phiếu đã xử lý xong.
- **Cơ chế:** Upsert Vũ trụ tổng (Excel) vào bảng `brcd_phieu` mỗi lần load `/brcd` + mỗi giờ qua cron `scripts/sync_brcd_phieu.py`.
- **Schema:** 1 dòng mỗi `baohong_id`. `first_seen` giữ nguyên qua sync, `last_seen` update mỗi lần. Phiếu rời tồn → dòng giữ lại với `last_seen` cũ.
- **Metric mới trong `/api/brcd-kiemsoat/thongke`:** section `lich_su` với `roi_da_ks` (rời tồn + đã KS) và `roi_chua_ks` (rời tồn + chưa KS). Filter `khoang` ∈ tuan_nay/thang_nay/nam_nay/tat_ca.
- **Cron setup (per-instance):**
  ```
  0 * * * * DASHV4_UNIT_CODE=<unit> /home/vtst/dashv4/venv/bin/python3 \
      /home/vtst/dashv4/scripts/sync_brcd_phieu.py \
      >> /home/vtst/dashv4/logs/brcd_phieu_sync.log 2>&1
  ```
```

- [ ] **Step 3: Final full test run**

Run: `python3 -m pytest tests/ -q`
Expected: 78 passed, 0 failed.

- [ ] **Step 4: Final syntax check tất cả file đã sửa**

Run:
```bash
python3 -m py_compile blueprints/operations_routes.py scripts/sync_brcd_phieu.py
node --check static/js/pages/brcd.js
```
Expected: no output từ cả 2.

- [ ] **Step 5: Commit docs**

```bash
git add docs/04-mapping-route-va-du-lieu.md docs/08-trang-thai-thuc-thi.md
git commit -m "docs(brcd): cập nhật docs/04 và docs/08 cho snapshot brcd_phieu

Thêm note snapshot hằng giờ + 2 metric lich_su mới trong thongke.
Hướng dẫn setup cron per-instance."
```

---

## Self-Review Checklist (chạy sau khi viết xong plan)

**1. Spec coverage:**
- ✅ Schema `brcd_phieu` (15 fields + sheet + 2 timestamp) → Task 1
- ✅ `_sync_brcd_phieu_to_db()` upsert → Task 1
- ✅ On-load trigger → Task 2
- ✅ Cron script → Task 3
- ✅ `/thongke` returns `lich_su` → Task 4
- ✅ Filter `khoang` + `doi`/`loaihinh` áp dụng → Task 4
- ✅ 2 thẻ UI + dropdown → Task 5
- ✅ Doc-sync (docs/04, docs/08) → Task 6
- ✅ Tests (sync, on-load, script, lich_su, HTML) → rải rác Tasks 1-5

**2. Placeholder scan:** không có TODO/TBD. Code block đầy đủ trong mọi step code.

**3. Type consistency:**
- `_sync_brcd_phieu_to_db()` return type: `{'synced', 'new', 'updated', 'skipped', 'reason'}` — consistent giữa Task 1 (define), Task 2 (gọi), Task 3 (script gọi).
- `_compute_lich_su(filtered_df, args)` return: `{'tu_ngay', 'den_ngay', 'roi_da_ks', 'roi_chua_ks'}` — consistent giữa Task 4 (define) và Task 5 (UI đọc).
- `_phieu_row_to_params(row, sheet, now_iso)` — Task 1.
- Column names trong schema khớp giữa Task 1 (CREATE TABLE) và `_BRCD_PHIEU_UPSERT_SQL` (cùng task).

**Edge cases đã cover:**
- Excel thiếu → sync skip, không crash
- DB lock → sync skip, không crash
- Schema idempotent
- Phiếu rời tồn không bị xóa khỏi DB
- Field changes update
- Filter `doi`/`loaihinh` áp dụng cho cả `lich_su`
- First-run backfill (first_seen = last_seen)
