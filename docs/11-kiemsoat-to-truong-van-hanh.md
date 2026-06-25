# Kiểm soát tổ trưởng — Hướng dẫn vận hành & bảo dưỡng

Tài liệu này mô tả đầy đủ lớp "kiểm soát tổ trưởng" tại `/brcd` và `/pttb`,
dành cho người vận hành và bảo dưỡng. Mọi thay đổi cấu trúc (thêm cột, đổi khóa,
đổi nguồn Excel, thêm instance) phải cập nhật tài liệu này cùng lúc.

Nguyên lý kiến trúc và lý do chọn thiết kế hiện tại nằm ở
[08-trang-thai-thuc-thi.md](/home/vtst/dashv4/docs/08-trang-thai-thuc-thi.md)
mục 6 và 7. Tài liệu này tập trung vào **vận hành**.

## 1. Tổng quan

Lớp kiểm soát cho phép tổ trưởng nhập "nội dung kiểm soát" (1 ô tự do) trên
từng phiếu tồn, kèm thống kê, lọc, xuất Excel và lưu lịch sử phiếu đã rời tồn.

Có hai trang, cấu trúc gần như giống hệt nhau:

| Trang | Route | Khóa phiếu | Bảng annotation | Bảng snapshot |
|-------|-------|-----------|-----------------|---------------|
| BRCD  | `/brcd` | `baohong_id` (INTEGER) | `brcd_kiemsoat` | `brcd_phieu` |
| PTTB  | `/pttb` | `MA_THUE_BAO` (TEXT)   | `pttb_kiemsoat` | `pttb_phieu` |

Cả 4 bảng nằm trong **cùng một file SQLite** per-instance: `brcd_kiemsoat.db`.

## 2. Nguyên lý thiết kế (để hiểu khi bảo dưỡng)

- **Write-aside, không ghi ngược Excel.** Tồn live đọc Excel read-only từ `1bss`
  (bị refresh ghi đè mỗi giờ). Annotation lưu trong SQLite riêng để không bị mất.
- **Khóa khác nhau.** BRCD dùng `baohong_id` (số nguyên, duy nhất trong tồn).
  PTTB dùng `MA_THUE_BAO` (text, duy nhất trong tồn). Đừng nhầm hai khóa.
- **Snapshot không xóa.** Khi phiếu rời tồn, dòng trong `*_phieu` **giữ lại**
  với `last_seen` cũ — dùng cho thống kê lịch sử và audit. Không bao giờ
  `DELETE` khỏi snapshot khi sync.
- **`first_seen` bất biến.** UPSERT chỉ set `first_seen` khi INSERT, bỏ qua khi
  UPDATE. `last_seen` update mỗi lần sync. Đừng đảo nhầm hai trường này.
- **SQLite ghi được, per-instance.** Mỗi instance gunicorn có DB riêng trong
  `INSTANCE_RUNTIME_DIR` (mặc định `runtime_app/<unit_code>/`). Không chia sẻ
  giữa các đơn vị.

## 3. Cơ sở dữ liệu

### 3.1. Đường dẫn DB

```python
# config.py
BRCD_KIEMSOAT_DB_PATH = _first_existing_path(
    os.getenv('DASHV4_BRCD_KIEMSOAT_DB_PATH'),           # override
    os.path.join(INSTANCE_RUNTIME_DIR, 'brcd_kiemsoat.db')  # mặc định
)
```

Kiểm tra đường dẫn thực tế của instance đang chạy:

```bash
# Ví dụ son_tay
ls -la /home/vtst/dashv4/runtime_app/son_tay/brcd_kiemsoat.db
```

> **Tên file giữ nguyên `brcd_kiemsoat.db`** dù chứa cả 4 bảng (kể cả PTTB).
> Đổi tên file sẽ phá vỡ env override hiện có trên các instance đã deploy.

### 3.2. Schema 4 bảng

Schema được tạo tự động (idempotent, `CREATE TABLE IF NOT EXISTS`) trong
`_ensure_brcd_kiemsoat_schema()` tại `blueprints/operations_routes.py:120`.

**`brcd_kiemsoat`** (annotation BRCD):

| Cột | Kiểu | Ghi chú |
|-----|------|---------|
| `baohong_id` | INTEGER PK | khóa phiếu |
| `ma_tb`, `doi_vt`, `nvkt` | TEXT | metadata chép từ tồn khi lưu |
| `noi_dung_kiem_soat` | TEXT | nội dung tổ trưởng nhập |
| `nguoi_nhap` | TEXT | username từ session |
| `thoi_diem_nhap` | TEXT | ISO timestamp lần đầu |
| `thoi_diem_cap_nhat` | TEXT | ISO timestamp lần cuối |

**`brcd_phieu`** (snapshot BRCD): cột = `baohong_id` PK + toàn bộ cột Excel
(`ma_tb`, `ten_tb`, `diachi_ld`, `loaihinh_tb`, `ghichu_hong`, `nvkt`, `doi_vt`,
`ngay_bh`, `trang_thai_cong`, `ttvt_ton`, `chitieu_tg`, `thoi_gian_ton_thuc`,
`gio_con_lai_thuc`, `sa`, `sheet`) + `first_seen` + `last_seen`. Index trên
`last_seen` và `doi_vt`.

**`pttb_kiemsoat`** (annotation PTTB): như `brcd_kiemsoat` nhưng PK là
`ma_thue_bao TEXT`, thay `nvkt` bằng `nhanvien_tiepthi`, thêm `loaihinh_tb`.

**`pttb_phieu`** (snapshot PTTB): cột = `ma_thue_bao` PK + các cột Excel
(`ten_thuebao`, `diachi_lapdat`, `loaihinh_tb`, `nhanvien_tiepthi`, `doi_vt`,
`ten_kv`, `ngayhen_den`, `noidung_hen`, `chitieu_tg`, `gio_conlai`,
`trang_thai`, `sheet`) + `first_seen` + `last_seen`. Index trên `last_seen`
và `doi_vt`.

> **Lưu ý case-sensitivity:** Schema DB dùng **chữ thường** (`ma_thue_bao`,
> `doi_vt`, ...). Nhưng Excel PTTB dùng **CHỮ HOA** (`MA_THUE_BAO`, `DOI_VT`,
> `LOAIHINH_TB`, `NHANVIEN_TIEPTHI`, ...). Code transform phải khớp đúng case
> của từng nguồn. Đây là điểm dễ sai nhất — xem mục 10.

## 4. Biến môi trường

| Biến | Mặc định | Mục đích |
|------|----------|---------|
| `DASHV4_BRCD_KIEMSOAT_DB_PATH` | `INSTANCE_RUNTIME_DIR/brcd_kiemsoat.db` | Đường dẫn DB ghi được |
| `DASHV4_RUNTIME_DIR` | `runtime_app/<UNIT_CODE>` | Thư mục runtime per-instance |
| `DASHV4_UNIT_CODE` | `son_tay` | Xác định instance (cho cron) |

Cron script cần `DASHV4_UNIT_CODE` để load đúng config (đường dẫn Excel, DB).

## 5. Nguồn dữ liệu Excel

Đường dẫn hardcode trong `blueprints/operations_routes.py:46-51`:

| Biến | Đường dẫn | Sheet |
|------|-----------|-------|
| `BRCD_DETAIL_MAIN_FILE` | `/home/vtst/1bss/runtime/default/downloads/chiaTheoDoi/chiTietBrcd5Doi.xlsx` | `ToKT_<doi>` |
| `PTTB_SUMMARY_FILE` | `/home/vtst/1bss/runtime/default/downloads/ton_pttb/baoCaoPTTB.xlsx` | `ToKT_<to>` |

Các file này được `1bss` refresh mỗi giờ. Khi file thiếu (lỡ nhịp), sync/cron
bỏ qua yên lặng (exit 0), không lỗi.

**Timestamp hiển thị** (chữ đỏ nổi bật phía trên bảng chi tiết) lấy từ
`file_info.modified` = thời gian sửa file Excel cuối cùng (`mtime`). Đây là
thời gian `1bss` tải file về, phản ánh độ mới của số liệu.

## 6. Cron — snapshot hằng giờ

Snapshot `*_phieu` được upsert qua 2 nhánh:

1. **On-load sync** — mỗi lần user mở `/brcd` hoặc `/pttb`, hàm
   `_sync_*_phieu_to_db()` chạy ngầm trong request handler.
2. **Cron hằng giờ** — độc lập với việc có ai mở trang không.

### 6.1. Setup cron (per-instance)

Mỗi instance cần 2 dòng crontab (BRCD + PTTB):

```cron
# BRCD — mỗi giờ
0 * * * * DASHV4_UNIT_CODE=<unit> /home/vtst/dashv4/venv/bin/python3 \
    /home/vtst/dashv4/scripts/sync_brcd_phieu.py \
    >> /home/vtst/dashv4/logs/brcd_phieu_sync.log 2>&1

# PTTB — mỗi giờ
0 * * * * DASHV4_UNIT_CODE=<unit> /home/vtst/dashv4/venv/bin/python3 \
    /home/vtst/dashv4/scripts/sync_pttb_phieu.py \
    >> /home/vtst/dashv4/logs/pttb_phieu_sync.log 2>&1
```

### 6.2. Kiểm tra cron đang chạy

```bash
# Xem crontab
crontab -l | grep sync_.*_phieu

# Xem log gần nhất
tail -20 /home/vtst/dashv4/logs/brcd_phieu_sync.log
tail -20 /home/vtst/dashv4/logs/pttb_phieu_sync.log

# Chạy thử thủ công
DASHV4_UNIT_CODE=son_tay /home/vtst/dashv4/venv/bin/python3 \
    /home/vtst/dashv4/scripts/sync_brcd_phieu.py
```

Output mong đợi: `{'synced': N, 'skipped': False}` hoặc
`{'synced': 0, 'skipped': True, 'reason': 'excel_missing'}`.

### 6.3. Exit codes

| Code | Ý nghĩa |
|------|---------|
| 0 | Sync thành công, hoặc Excel thiếu (lỡ nhịp refresh, không lỗi) |
| 1 | Lỗi khác (DB lock, Excel corrupt, ...) |

Excel missing exit 0 vì là tình trạng tạm thời, không cần báo động.

## 7. API endpoints

Tất cả endpoint nằm trong `blueprints/operations_routes.py`, đăng ký qua
blueprint `operations`.

### BRCD

| Method | Endpoint | Mục đích |
|--------|----------|---------|
| POST | `/api/brcd-kiemsoat/luu` | Lưu/noi_dung_kiem_soat (upsert hoặc xóa nếu rỗng) |
| GET | `/api/brcd-kiemsoat/detail` | Sheet `ToKT_<doi>` + JOIN annotation + `file_info` |
| GET | `/api/brcd-kiemsoat/thongke` | Thống kê + by_doi + by_nvkt + chi_tiet + lich_su |
| GET | `/download/brcd-kiemsoat-report` | Xuất Excel theo bộ lọc |

### PTTB

| Method | Endpoint | Mục đích |
|--------|----------|---------|
| POST | `/api/pttb-kiemsoat/luu` | Lưu (keyed `ma_thue_bao`) |
| GET | `/api/pttb-kiemsoat/detail` | Sheet `ToKT_<to>` + JOIN annotation + `file_info` |
| GET | `/api/pttb-kiemsoat/thongke` | Thống kê + lich_su |
| GET | `/download/pttb-kiemsoat-report` | Xuất Excel theo bộ lọc |

POST endpoint yêu cầu `@login_required` (JSON body, không dùng `@csrf_protect`).

### Tham số lọc thongke

- `trangthai` ∈ `tat_ca` / `da_ks` / `chua_ks`
- `quagio` ∈ `tat_ca` / `qua_gio` / `chua_qua`
- `doi` — tên đội (ví dụ `Đội 1`) hoặc `tat_ca`
- `khoang` (cho section `lich_su`) ∈ `tuan_nay` / `thang_nay` / `nam_nay` / `tat_ca`

## 8. Backup & Restore

DB `brcd_kiemsoat.db` chứa annotation do tổ trưởng nhập và lịch sử snapshot.
**Mất file = mất toàn bộ ghi chú kiểm soát.** Phải backup định kỳ.

### Backup

```bash
# Backup dùng SQLite online backup (an toàn khi app đang chạy)
sqlite3 /home/vtst/dashv4/runtime_app/<unit>/brcd_kiemsoat.db ".backup \
    /home/vtst/backup/dashv4/<unit>/brcd_kiemsoat_$(date +%Y%m%d).db"
```

> Dùng `.backup` thay vì `cp` vì `cp` file đang được ghi có thể tạo bản corrupt.
> Lệnh `.backup` dùng SQLite online backup API, an toàn khi app đang hoạt động.

### Restore

```bash
# Dừng instance trước (tránh ghi đè)
sudo systemctl stop dashv4@<unit>

# Đổi tên DB hiện tại (đừng xóa ngay)
mv /home/vtst/dashv4/runtime_app/<unit>/brcd_kiemsoat.db \
   /home/vtst/dashv4/runtime_app/<unit>/brcd_kiemsoat.db.broken

# Copy bản backup vào
cp /home/vtst/backup/dashv4/<unit>/brcd_kiemsoat_<date>.db \
   /home/vtst/dashv4/runtime_app/<unit>/brcd_kiemsoat.db

# Khởi động lại
sudo systemctl start dashv4@<unit>
```

## 9. Troubleshooting

### Timestamp chữ đỏ không hiện

Kiểm tra endpoint detail có trả `file_info`:

```bash
curl -s -b cookies.txt http://localhost:5011/api/brcd-kiemsoat/detail | python3 -c "import sys,json; print(json.load(sys.stdin).get('file_info'))"
```

Nếu `file_info` null → file Excel không tồn tại trên host (kiểm tra đường dẫn
mục 5). Nếu `file_info.modified` có giá trị nhưng UI không hiện → clear cache
trình duyệt (JS cũ).

### Annotation không lưu được

1. Kiểm tra response POST `/api/brcd-kiemsoat/luu` — có lỗi 500 không?
2. Kiểm tra DB file có ghi được không:
   ```bash
   ls -la /home/vtst/dashv4/runtime_app/<unit>/brcd_kiemsoat.db
   # quyền phải là user chạy gunicorn (thường vtst:vtst, rw)
   ```
3. Kiểm tra log gunicorn:
   ```bash
   tail -50 /home/vtst/dashv4/logs/gunicorn-error.log
   ```

### Thống kê `lich_su` trống

- Snapshot `*_phieu` chưa có dữ liệu (cron chưa chạy lần nào, hoặc Excel thiếu).
- Kiểm tra:
  ```bash
  sqlite3 /home/vtst/dashv4/runtime_app/<unit>/brcd_kiemsoat.db \
      "SELECT COUNT(*) FROM brcd_phieu; SELECT COUNT(*) FROM pttb_phieu;"
  ```
- Nếu 0 → chạy sync thủ công (mục 6.2).

### `sqlite3.OperationalError: too many SQL variables`

Xảy ra khi tồn quá lớn. Code đã batch IN-clause 500 giá trị/lô — nếu vẫn lỗi,
kiểm tra `_compute_lich_su` / `_compute_pttb_lich_su` có dùng batch không.

## 10. Cạm bẫy / lưu ý quan trọng

### 10.1. Case-sensitivity cột Excel PTTB

Excel PTTB dùng **CHỮ HOA**: `MA_THUE_BAO`, `TEN_THUEBAO`, `DIACHI_LAPDAT`,
`LOAIHINH_TB`, `NHANVIEN_TIEPTHI`, `DOI_VT`, `TEN_KV`, `NGAYHEN_DEN`,
`NOIDUNG_HEN`. Nhưng có vài cột thường: `chitieu_tg`, `gio_conlai`,
`trang_thai`.

BRCD Excel dùng **mixed case** khớp code trực tiếp (`baohong_id`, `ma_tb`,
`TEN_TB`, `DOI_VT`, `NVKT`, ...).

Khi thêm cột mới, **phải mở file Excel gốc để xác nhận case** trước khi code.

### 10.2. Đừng đổi tên file DB

Tên file là `brcd_kiemsoat.db` (có chữ "brcd") dù chứa cả bảng PTTB. Các
instance đã deploy có env `DASHV4_BRCD_KIEMSOAT_DB_PATH` trỏ đúng tên này.
Đổi tên sẽ phá vỡ production.

### 10.3. PTTB không dùng `MA_GIAO_DICH` làm khóa

`MA_GIAO_DICH` có duplicate (1 giao dịch bao nhiều thuê bao). Khóa PTTB là
`MA_THUE_BAO` (text, duy nhất). Đừng đổi sang `MA_GIAO_DICH`.

### 10.4. Runtime limits import trước pandas

Cron script (`scripts/sync_*_phieu.py`) phải `import runtime_limits` trước khi
import pandas/numpy, để cap BLAS thread count. Bỏ qua sẽ oversubscribe CPU khi
chạy nhiều instance song song. Pattern:

```python
import runtime_limits  # noqa: F401, E402
from blueprints.operations_routes import _sync_brcd_phieu_to_db
```

### 10.5. Schema cache per-worker

`_brcd_kiemsoat_schema_ready_path` cache đường dẫn đã init schema per-worker.
Khi test monkeypatch `BRCD_KIEMSOAT_DB_PATH`, phải reset biến này về `None`
sau test (xem pattern trong `tests/test_brcd_kiemsoat.py`).

## 11. Test

```bash
# Toàn bộ test kiemsoat
python3 -m pytest tests/test_brcd_kiemsoat.py tests/test_brcd_phieu_snapshot.py tests/test_pttb_kiemsoat.py -v

# Toàn bộ suite
python3 -m pytest tests/ -q
```

Test tự chứa: build SQLite trong `tmp_path`, monkeypatch config attrs, không
đọc DB thật. Sau monkeypatch `BRCD_KIEMSOAT_DB_PATH`, reset
`_brcd_kiemsoat_schema_ready_path = None` ở teardown.

| File test | Số test | Phủ |
|-----------|---------|-----|
| `tests/test_brcd_kiemsoat.py` | 17 | BRCD annotation CRUD + thongke |
| `tests/test_brcd_phieu_snapshot.py` | 14 | BRCD snapshot sync + lich_su |
| `tests/test_pttb_kiemsoat.py` | 20 | PTTB annotation + snapshot + thongke |

## 12. File tham chiếu nhanh

| File | Vai trò |
|------|---------|
| `blueprints/operations_routes.py` | Toàn bộ logic: schema, sync, 8 endpoint, `_compute_*_lich_su` |
| `config.py` | `BRCD_KIEMSOAT_DB_PATH`, `INSTANCE_RUNTIME_DIR` |
| `scripts/sync_brcd_phieu.py` | Cron entry point BRCD |
| `scripts/sync_pttb_phieu.py` | Cron entry point PTTB |
| `static/js/pages/brcd.js` | BRCD UI: bảng chi tiết (inline textarea), thongke, timestamp |
| `static/js/pages/pttb.js` | PTTB UI: tương tự BRCD |
| `static/js/api.js` | 6 API method (3 BRCD + 3 PTTB) |
| `templates/pages/brcd.html` | Template BRCD |
| `templates/pages/pttb.html` | Template PTTB |
| `tests/test_brcd_kiemsoat.py` | Test BRCD annotation |
| `tests/test_brcd_phieu_snapshot.py` | Test BRCD snapshot |
| `tests/test_pttb_kiemsoat.py` | Test PTTB |

## 13. Checklist bảo dưỡng định kỳ

- [ ] Kiểm tra cron đang chạy: `crontab -l | grep sync_.*_phieu`
- [ ] Kiểm tra log sync không có lỗi: `tail logs/*_phieu_sync.log`
- [ ] Backup DB: `sqlite3 .../brcd_kiemsoat.db ".backup ..."`
- [ ] Kiểm tra dung lượng DB không phình to bất thường:
  ```bash
  sqlite3 .../brcd_kiemsoat.db "SELECT COUNT(*) FROM brcd_phieu; SELECT COUNT(*) FROM pttb_phieu;"
  ```
- [ ] Chạy test: `python3 -m pytest tests/ -q`
- [ ] Mở `/brcd` và `/pttb` kiểm tra timestamp chữ đỏ hiển thị đúng
