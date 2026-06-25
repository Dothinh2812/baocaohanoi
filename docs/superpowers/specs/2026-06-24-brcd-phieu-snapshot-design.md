# Thiết kế: Snapshot lịch sử phiếu BRCD vào DB

- **Ngày:** 2026-06-24
- **Liên quan:** `docs/superpowers/plans/2026-06-24-brcd-kiemsoat-totronuong.md` (đã triển khai)
- **Trạng thái:** Đã duyệt, chờ writing-plans

## Bối cảnh & vấn đề

Tính năng "Kiểm soát tổ trưởng" tại `/brcd` đã chạy: DB `brcd_kiemsoat.db` (bảng `brcd_kiemsoat`) lưu annotation do tổ trưởng nhập, keyed bằng `baohong_id`. Thống kê "đã/chưa nhập" được tính bằng cách left-join annotation với **Vũ trụ tổng** — toàn bộ phiếu đang tồn đọc từ sheet `ToKT_<doi>` trong `chiTietBrcd5Doi.xlsx`.

File Excel này bị 1bss **ghi đè mỗi giờ**. Khi một phiếu được xử lý xong, nó biến mất khỏi Vũ trụ tổng → mất luôn khả năng tra cứu. Hệ quả:

- Không audit được "phiếu này từng tồn bao lâu trước khi rời đi".
- Không đánh giá được tổ trưởng trên các phiếu đã xử lý xong: phiếu đã rời tồn mà chưa từng được kiểm soát = "mất kiểm soát", nhưng sau khi rời đi thì không còn biết tới sự tồn tại của nó.
- `first_seen` thực sự (lần đầu phiếu xuất hiện trong tồn) không có — chỉ có `thoi_diem_nhap` của annotation, không biết độ trễ từ khi phiếu xuất hiện tới khi được kiểm soát.

## Mục tiêu

Lưu snapshot toàn bộ phiếu từng xuất hiện trong Vũ trụ tổng vào DB, keyed bằng `baohong_id`, phục vụ:

1. Tra cứu lịch sử phiếu sau khi đã rời tồn.
2. Đánh giá kiểm soát tổ trưởng: phiếu rời tồn mà chưa được annotation = mất kiểm soát.
3. Đo độ trễ kiểm soát: từ `first_seen` (lần đầu xuất hiện) đến `thoi_diem_nhap` (annotation đầu tiên).

## Phạm vi

**Trong phạm vi:**
- Bảng mới `brcd_phieu` (cùng file DB `brcd_kiemsoat.db` hiện có).
- Function sync `_sync_brcd_phieu_to_db()` — upsert Vũ trụ tổng vào DB.
- 2 entry point: on-load (mỗi lần load `_load_brcd_kiemsoat_df`) + script cron `scripts/sync_brcd_phieu.py`.
- Mở rộng endpoint `/api/brcd-kiemsoat/thongke` trả thêm section `lich_su` với 2 metric mới.
- Mở rộng UI section "Thống kê kiểm soát" hiển thị 2 metric mới + filter khoảng thời gian.
- Test cho sync function + endpoint + script.

**Ngoài phạm vi (YAGNI):**
- Không track lịch sử thay đổi trường của từng phiếu (NVKT đổi, địa chỉ đổi giữa chừng) — đã chốt ở bước brainstorm. Chấp nhận mất thông tin này để giữ schema đơn giản.
- Không có UI view mới (tab "Lịch sử tồn") — chỉ mở rộng thống kê.
- Không backfill `first_seen` cho quá khứ — chỉ capture từ thời điểm deploy.
- Không thêm endpoint upload/sửa snapshot — snapshot là read-only từ hệ thống.

## Quyết định thiết kế

### Phương án đã chọn: Bảng riêng `brcd_phieu`

Loại trừ:
- **Mở rộng bảng `brcd_kiemsoat`**: ô nhiễm ngữ nghĩa (bảng "annotation" giờ chứa cả phiếu chưa annotate), phải giữ DEFAULT rỗng cho mọi cột mới để không phá dữ liệu cũ.
- **Hai bảng (current + history append-only)**: overkill vì đã chốt không track thay đổi trường.

Lý do chọn: tách trách nhiệm rõ (snapshot ≠ annotation), query trực quan, không phá dữ liệu hiện có.

### Chiến lược capture

- **Upsert + first_seen/last_seen** (đã chốt): 1 dòng mỗi `baohong_id`. `first_seen` chỉ set khi INSERT. `last_seen` update mỗi lần sync. Các trường snapshot khác bị ghi đè mỗi sync.
- **Sync trigger**: cả hai (cron hằng giờ + on-load). Cron đảm bảo history liên tục kể cả khi không ai mở dashboard; on-load đảm bảo UI luôn thấy dữ liệu mới nhất.
- **Dữ liệu dùng để**: archive + mở rộng thống kê (không thêm UI view mới).

## Thiết kế chi tiết

### 1. Schema — bảng `brcd_phieu`

Cùng file `brcd_kiemsoat.db` (env `DASHV4_BRCD_KIEMSOAT_DB_PATH`).

```sql
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
    sheet              TEXT,   -- tên sheet ToKT_<doi> lần cuối nhìn thấy
    first_seen         TEXT,   -- ISO timestamp, chỉ set khi INSERT
    last_seen          TEXT    -- ISO timestamp, update mỗi sync
);
CREATE INDEX IF NOT EXISTS idx_brcd_phieu_last_seen ON brcd_phieu(last_seen);
CREATE INDEX IF NOT EXISTS idx_brcd_phieu_doi_vt    ON brcd_phieu(doi_vt);
```

Mapping cột Excel → cột DB (theo `BRCD_KIEMSOAT_DISPLAY_COLUMNS` hiện có):

| Excel column           | DB column          |
|------------------------|--------------------|
| `baohong_id`           | `baohong_id` (PK)  |
| `ma_tb`                | `ma_tb`            |
| `TEN_TB`               | `ten_tb`           |
| `DIACHI_LD`            | `diachi_ld`        |
| `LOAIHINH_TB`          | `loaihinh_tb`      |
| `GHICHU_HONG`          | `ghichu_hong`      |
| `NVKT`                 | `nvkt`             |
| `DOI_VT`               | `doi_vt`           |
| `ngay_bh`              | `ngay_bh`          |
| `Trạng thái cổng`      | `trang_thai_cong`  |
| `ttvt_ton`             | `ttvt_ton`         |
| `chitieu_tg`           | `chitieu_tg`       |
| `thời gian tồn thực`   | `thoi_gian_ton_thuc` |
| `giờ còn lại thực`     | `gio_con_lai_thuc` |
| `SA`                   | `sa`               |

### 2. Function sync — `_sync_brcd_phieu_to_db()`

Vị trí: `blueprints/operations_routes.py`, gần các helper kiemsoat hiện có.

```python
def _sync_brcd_phieu_to_db() -> dict:
    """Đọc Vũ trụ tổng hiện tại, upsert vào brcd_phieu.
    Trả {'synced': N, 'new': M, 'updated': K, 'skipped': 0|1, 'reason': str}.
    Idempotent — chạy lại cho cùng dữ liệu không thay đổi DB.
    Không raise khi Excel thiếu — trả {'skipped': 1, 'reason': 'excel_missing'}.
    """
```

Hành vi:
- Nếu `BRCD_DETAIL_MAIN_FILE` không tồn tại → trả `{'skipped': 1, 'reason': 'excel_missing', 'synced': 0}`. Không log error (bình thường giữa các lần refresh).
- Đọc tất cả sheet `ToKT_*` (không `_rut_gon`) bằng `read_excel_sheet_cached`.
- Lọc cột theo `BRCD_KIEMSOAT_DISPLAY_COLUMNS`, dropna `baohong_id`, ép int.
- Với mỗi dòng: dùng SQLite UPSERT (SQLite ≥ 3.24, 2018):
  ```sql
  INSERT INTO brcd_phieu (baohong_id, ma_tb, ..., first_seen, last_seen)
  VALUES (?, ?, ..., ?, ?)
  ON CONFLICT(baohong_id) DO UPDATE SET
      ma_tb = excluded.ma_tb,
      ...,
      last_seen = excluded.last_seen;
      -- first_seen KHÔNG nằm trong SET list → giữ giá trị cũ
  ```
  Tức `first_seen` chỉ set khi INSERT (lần đầu), không bao giờ UPDATE.
- Đóng gói trong 1 transaction (`BEGIN IMMEDIATE` → `COMMIT`).
- Trả dict thống kê để log và test.

### 3. Entry points

**On-load (transparent):**
```python
def _load_brcd_kiemsoat_df():
    if not os.path.exists(BRCD_DETAIL_MAIN_FILE):
        return None
    _ensure_brcd_kiemsoat_schema()
    _sync_brcd_phieu_to_db()   # <-- thêm dòng này, fire-and-forget
    # ...phần còn lại giữ nguyên...
```
Không bắt exception từ sync trong on-load — nếu DB lock, sync bỏ qua, trang vẫn render bằng dữ liệu DB cũ. Sync chỉ in warning ra log, không raise.

**Cron script — `scripts/sync_brcd_phieu.py`:**
```python
#!/usr/bin/env python3
"""Cron entry point — sync BRCD phieu snapshot cho instance hiện tại.
Config từ DASHV4_* env vars. Không qua HTTP."""
import runtime_limits  # phải import trước pandas
from blueprints.operations_routes import _sync_brcd_phieu_to_db

if __name__ == '__main__':
    result = _sync_brcd_phieu_to_db()
    print(result)
    if result.get('skipped') and result.get('reason') != 'excel_missing':
        raise SystemExit(1)
```

Crontab (mỗi instance 1 dòng):
```
0 * * * * DASHV4_UNIT_CODE=son_tay /home/vtst/dashv4/venv/bin/python3 /home/vtst/dashv4/scripts/sync_brcd_phieu.py >> /home/vtst/dashv4/logs/brcd_phieu_sync.log 2>&1
```

Lý do script import trực tiếp chứ không gọi HTTP: tránh phức tạp auth/session (dashboard yêu cầu login). Script dùng cùng env với systemd service (`DASHV4_UNIT_CODE` → `config.py` → đúng `INSTANCE_RUNTIME_DIR` → đúng DB path).

### 4. Mở rộng endpoint `/api/brcd-kiemsoat/thongke`

Payload hiện có (giữ nguyên):
```json
{
  "summary": {"total": N, "da_nhap": A, "chua_nhap": B, "ty_le": 0.X},
  "by_doi": [...],
  "by_nvkt": [...],
  "chi_tiet": [...]
}
```

Thêm section mới `lich_su` (chỉ tính khi có DB `brcd_phieu` đã có dữ liệu):
```json
{
  "lich_su": {
    "tu_ngay": "2026-06-01",
    "den_ngay": "2026-06-24",
    "roi_da_ks": 23,         // phiếu rời tồn trong khoảng, ĐÃ được kiểm soát
    "roi_chua_ks": 7         // phiếu rời tồn trong khoảng, CHƯA được kiểm soát (mất KS)
  }
}
```

Query logic:
- "Rời tồn" = `baohong_id` có trong `brcd_phieu` NHƯNG không có trong Vũ trụ tổng hiện tại (computed bằng set difference khi đã có DataFrame `_load_brcd_kiemsoat_df()`).
- "Trong khoảng [tu_ngay, den_ngay]" = `DATE(last_seen) BETWEEN tu_ngay AND den_ngay`.
- "Đã kiểm soát" = JOIN `brcd_kiemsoat` bằng `baohong_id` tìm thấy annotation (không rỗng).
- Filter mới: query param `khoang` = `tuan_nay` | `thang_nay` | `nam_nay` | `tat_ca` (mặc định `thang_nay`).
- Filter `doi`, `loaihinh`, `nhom` (nếu có) cũng áp dụng cho 2 metric mới.

### 5. UI — mở rộng section "Thống kê kiểm soát"

Trong `templates/pages/brcd.html` và `static/js/pages/brcd.js`:

- Thêm dropdown "Khoảng thời gian (lịch sử)" với 4 options: Tuần này / Tháng này (default) / Năm nay / Tất cả.
- Thêm 2 thẻ stat mới cạnh các thẻ hiện tại: **"Rời tồn đã KS"** và **"Rời tồn chưa KS"** (with label màu đỏ cho chưa KS để nhấn mạnh mất kiểm soát).
- Filter `doi`, `loaihinh` hiện có áp dụng cho cả 2 metric mới.
- Không thêm bảng chi tiết mới (YAGNI) — chi tiết phiếu rời tồn query qua sqlite3 khi cần.

### 6. Error handling

| Trường hợp | Xử lý |
|---|---|
| Excel thiếu (giữa 2 lần refresh) | sync trả `{'skipped': 1, 'reason': 'excel_missing'}`. Trang render bằng DB snapshot cũ. Không log error. |
| DB lock | WAL + busy_timeout 5s đã có. Nếu vẫn lock (OperatingError) → bắt exception, trả `{'skipped': 1, 'reason': 'db_locked'}`, không raise ra on-load. |
| Cron script fail (lý do khác excel_missing) | Exit code 1, stderr in ra log file. Systemd/cron báo. |
| Lần đầu chạy | Backfill tất cả phiếu Vũ trụ tổng hiện tại với `first_seen = last_seen = now`. Chấp nhận mất first_seen lịch sử. |
| Schema migration | `CREATE TABLE IF NOT EXISTS` idempotent. Bảng thêm vào `_ensure_brcd_kiemsoat_schema()` (rename không hợp lý — giữ tên function, kiểm tra cả 2 bảng). |

### 7. Testing — `tests/test_brcd_phieu_snapshot.py`

Dùng openpyxl tạo file Excel giả trong `tmp_path`, monkeypatch `BRCD_DETAIL_MAIN_FILE` và `BRCD_KIEMSOAT_DB_PATH`. Pattern giống `tests/test_brcd_kiemsoat.py`.

**Test cases:**
1. Sync lần đầu — insert mới: tất cả dòng có `first_seen = last_seen`.
2. Sync lần 2 (Excel không đổi) — update last_seen, `first_seen` giữ nguyên, các trường snapshot giữ nguyên.
3. Sync lần 3 (Excel thay đổi: NVKT đổi, thêm dòng mới, bỏ 1 dòng) — dòng cũ update, dòng mới insert, dòng bị bỏ GIỮ LẠI trong DB (không xóa).
4. Sync khi Excel thiếu — trả `{'skipped': 1, 'reason': 'excel_missing'}`, DB không đổi.
5. Endpoint `/thongke?khoang=thang_nay` trả section `lich_su` đúng số liệu (chế fixture: 2 phiếu rời tồn đã KS + 1 chưa KS trong tháng).
6. Endpoint `/thongke?khoang=tat_ca` trả tất cả phiếu rời tồn (không filter ngày).
7. Filter `doi` áp dụng cho cả 2 metric `lich_su`.
8. Script `scripts/sync_brcd_phieu.py` chạy với `DASHV4_UNIT_CODE=test` + tmp DB → exit 0, DB có dữ liệu.
9. Schema idempotent — chạy `_ensure_brcd_kiemsoat_schema()` 2 lần không lỗi.

## Tác động phụ

- **DB size**: tăng nhẹ. Mỗi dòng ~200 bytes × vài nghìn phiếu ≈ < 1 MB. WAL mode vẫn áp dụng.
- **Page load perf**: thêm 1 upsert transaction mỗi load (~50ms cho ~200 phiếu). Acceptable — không thêm network round-trip, chỉ local SQLite.
- **Cron load**: 1 lần/h, mỗi lần ~1s (đọc Excel + upsert). Negligible.
- **Multi-instance**: mỗi instance có DB riêng, script cron riêng. Không xung đột.

## Doc-sync rule (bắt buộc theo AGENTS.md)

Cập nhật cùng change:
- `docs/04-mapping-route-va-du-lieu.md` — cập nhật row `/brcd`: thêm note "đã có snapshot `brcd_phieu` hằng giờ".
- `docs/08-trang-thai-thuc-thi.md` — section 6 "Kiểm soát tổ trưởng": bổ sung subsection về snapshot + 2 metric mới.

## Sequence triển khai (gợi ý, chi tiết sẽ do writing-plans quyết định)

1. Thêm schema `brcd_phieu` vào `_ensure_brcd_kiemsoat_schema()`.
2. Viết `_sync_brcd_phieu_to_db()` + unit test.
3. Wire on-load vào `_load_brcd_kiemsoat_df()`.
4. Viết script `scripts/sync_brcd_phieu.py`.
5. Mở rộng endpoint `/thongke` (section `lich_su` + filter `khoang`).
6. Mở rộng UI (2 thẻ + dropdown).
7. Doc-sync.
