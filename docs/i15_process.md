# Tài liệu Module i15_process.py

## Mô tả

Module xử lý báo cáo **I1.5 Suy hao cao (SHC)** với các tính năng nâng cao:
- Tracking lịch sử theo ngày
- Bảo vệ chống chạy trùng dữ liệu
- So sánh biến động với ngày trước (T so với T-1)
- Hỗ trợ cả báo cáo K1 và K2

---

## Cấu trúc File

```
i15_process.py
├── normalize_nvkt()                         # Chuẩn hóa tên NVKT
├── process_I15_report_with_tracking()       # Wrapper xử lý K1
├── process_I15_k2_report_with_tracking()    # Wrapper xử lý K2
└── _process_I15_generic_with_tracking()     # Hàm xử lý chính
```

---

## Các hàm chính

### 1. `normalize_nvkt(x)`

Chuẩn hóa tên NVKT bằng cách:
- Giữ phần sau dấu `-` (ví dụ: `"ABC - Nguyễn Văn A"` → `"Nguyễn Văn A"`)
- Loại bỏ phần trong ngoặc `()`

**Tham số:**
| Tham số | Kiểu | Mô tả |
|---------|------|-------|
| `x` | str | Tên NVKT cần chuẩn hóa |

**Trả về:** `str` - Tên NVKT đã chuẩn hóa

---

### 2. `process_I15_report_with_tracking(force_update=False)`

Wrapper function để xử lý báo cáo **K1**.

**Tham số:**
| Tham số | Kiểu | Mặc định | Mô tả |
|---------|------|----------|-------|
| `force_update` | bool | `False` | Ghi đè dữ liệu nếu đã tồn tại |

**File đầu vào:** `downloads/baocao_hanoi/I1.5 report.xlsx`

**Database:** `suy_hao_history.db`

---

### 3. `process_I15_k2_report_with_tracking(force_update=False)`

Wrapper function để xử lý báo cáo **K2**.

**Tham số:**
| Tham số | Kiểu | Mặc định | Mô tả |
|---------|------|----------|-------|
| `force_update` | bool | `False` | Ghi đè dữ liệu nếu đã tồn tại |

**File đầu vào:** `downloads/baocao_hanoi/I1.5_k2 report.xlsx`

**Database:** `suy_hao_history_k2.db`

---

### 4. `_process_I15_generic_with_tracking(input_file, k_suffix, history_db, force_update)`

Hàm xử lý chính cho cả K1 và K2.

**Tham số:**
| Tham số | Kiểu | Mô tả |
|---------|------|-------|
| `input_file` | str | Đường dẫn file Excel đầu vào |
| `k_suffix` | str | Hậu tố `"K1"` hoặc `"K2"` |
| `history_db` | str | Tên file database lịch sử |
| `force_update` | bool | Ghi đè dữ liệu nếu đã tồn tại |

**Trả về:** `bool` - `True` nếu thành công, `False` nếu lỗi

---

## Quy trình xử lý

```
┌─────────────────────────────────────────────────────────────┐
│  1. ĐỌC DỮ LIỆU                                             │
│     - Đọc file Excel đầu vào                                │
│     - Lấy ngày báo cáo từ cột NGAY_SUYHAO                   │
└─────────────────────────┬───────────────────────────────────┘
                          ▼
┌─────────────────────────────────────────────────────────────┐
│  2. TRA CỨU THÔNG TIN                                       │
│     - Merge với danhba.db để lấy THIETBI, SA, KETCUOI       │
│     - Đọc thông tin từ bảng thong_ke, thong_ke_theo_don_vi  │
└─────────────────────────┬───────────────────────────────────┘
                          ▼
┌─────────────────────────────────────────────────────────────┐
│  3. CHUẨN HÓA DỮ LIỆU                                       │
│     - Tạo cột NVKT_DB_NORMALIZED                            │
└─────────────────────────┬───────────────────────────────────┘
                          ▼
┌─────────────────────────────────────────────────────────────┐
│  4. KIỂM TRA TRÙNG LẶP                                      │
│     - Nếu đã có dữ liệu ngày này → đọc từ DB                │
│     - Nếu force_update=True → ghi đè                        │
└─────────────────────────┬───────────────────────────────────┘
                          ▼
┌─────────────────────────────────────────────────────────────┐
│  5. SO SÁNH BIẾN ĐỘNG (T so với T-1)                        │
│     - TĂNG MỚI: Thuê bao có hôm nay, không có hôm qua       │
│     - GIẢM/HẾT: Thuê bao có hôm qua, không có hôm nay       │
│     - VẪN CÒN: Thuê bao có cả hai ngày                      │
└─────────────────────────┬───────────────────────────────────┘
                          ▼
┌─────────────────────────────────────────────────────────────┐
│  6. LƯU VÀO DATABASE                                        │
│     - suy_hao_snapshots: Snapshot hàng ngày                 │
│     - suy_hao_tracking: Theo dõi số ngày liên tục           │
│     - suy_hao_daily_changes: Chi tiết biến động             │
│     - suy_hao_daily_summary: Tổng hợp theo NVKT             │
└─────────────────────────┬───────────────────────────────────┘
                          ▼
┌─────────────────────────────────────────────────────────────┐
│  7. TẠO CÁC SHEET EXCEL                                     │
│     - TH_SHC_I15: Tổng hợp theo NVKT + Đơn vị               │
│     - TH_SHC_theo_to: Tổng hợp theo đơn vị                  │
│     - shc_theo_SA: Thống kê theo SA                         │
│     - Bien_dong_tong_hop: Biến động ngày                    │
│     - Tang_moi, Giam_het, Van_con: Chi tiết biến động       │
│     - Sheet chi tiết cho từng NVKT                          │
└─────────────────────────────────────────────────────────────┘
```

---

## Database Schema

### Bảng `suy_hao_snapshots`
Lưu snapshot thuê bao suy hao mỗi ngày.

| Cột | Mô tả |
|-----|-------|
| ngay_bao_cao | Ngày báo cáo |
| account_cts | Số thuê bao |
| ten_tb_one | Tên thuê bao |
| doi_one | Đơn vị |
| nvkt_db_normalized | NVKT đã chuẩn hóa |
| sa | Mã SA |
| thietbi, ketcuoi | Thông tin thiết bị |

### Bảng `suy_hao_tracking`
Theo dõi số ngày liên tục thuê bao bị suy hao.

| Cột | Mô tả |
|-----|-------|
| account_cts | Số thuê bao |
| ngay_xuat_hien_dau_tien | Ngày đầu tiên phát hiện |
| ngay_thay_cuoi_cung | Ngày cuối cùng còn suy hao |
| so_ngay_lien_tuc | Số ngày suy hao liên tục |
| trang_thai | DANG_SUY_HAO / DA_HET_SUY_HAO |

### Bảng `suy_hao_daily_changes`
Chi tiết biến động hàng ngày.

| Cột | Mô tả |
|-----|-------|
| ngay_bao_cao | Ngày báo cáo |
| account_cts | Số thuê bao |
| loai_bien_dong | TANG_MOI / GIAM_HET / VAN_CON |

### Bảng `suy_hao_daily_summary`
Tổng hợp theo đơn vị và NVKT mỗi ngày.

| Cột | Mô tả |
|-----|-------|
| ngay_bao_cao | Ngày báo cáo |
| doi_one | Đơn vị |
| nvkt_db_normalized | NVKT |
| tong_so_hien_tai | Tổng số thuê bao suy hao |
| so_tang_moi | Số tăng mới |
| so_giam_het | Số giảm/hết |
| so_van_con | Số vẫn còn |
| so_tb_quan_ly | Tổng số TB quản lý |
| ty_le_shc | Tỉ lệ suy hao cao (%) |

---

## Cách sử dụng

### Chạy từ Command Line

```bash
# Chạy cả K1 và K2 (mặc định)
python i15_process.py

# Chỉ chạy K1
python i15_process.py --k1

# Chỉ chạy K2
python i15_process.py --k2

# Ghi đè dữ liệu đã tồn tại
python i15_process.py --force

# Kết hợp tùy chọn
python i15_process.py --k2 --force

# Xem trợ giúp
python i15_process.py --help
```

### Import và sử dụng trong code

```python
from i15_process import (
    process_I15_report_with_tracking,
    process_I15_k2_report_with_tracking
)

# Xử lý K1
process_I15_report_with_tracking()

# Xử lý K2 với force update  
process_I15_k2_report_with_tracking(force_update=True)
```

---

## Output Files

### File Excel đầu ra

Các sheet được tạo trong file input:

| Sheet | Mô tả |
|-------|-------|
| `Sheet1` | Dữ liệu gốc đã enrich |
| `TH_SHC_I15` | Tổng hợp theo NVKT + Đơn vị |
| `TH_SHC_theo_to` | Tổng hợp theo đơn vị |
| `shc_theo_SA` | Thống kê theo SA |
| `Bien_dong_tong_hop` | Biến động tổng hợp theo NVKT |
| `Tang_moi` | Danh sách thuê bao tăng mới |
| `Giam_het` | Danh sách thuê bao đã giảm/hết |
| `Van_con` | Danh sách thuê bao vẫn còn |
| `[Tên NVKT]` | Sheet chi tiết cho từng NVKT |

### Báo cáo so sánh (khi chạy standalone)

| File | Mô tả |
|------|-------|
| `So_sanh_SHC_theo_ngay_T-1.xlsx` | So sánh SHC K1 theo ngày |
| `So_sanh_SHC_k2_theo_ngay_T-1.xlsx` | So sánh SHC K2 theo ngày |

---

## Lưu ý quan trọng

> [!WARNING]
> - Nếu đã chạy trong ngày, hệ thống sẽ **bỏ qua lưu database** để tránh trùng lặp
> - Sử dụng `force_update=True` hoặc `--force` để ghi đè khi cần

> [!NOTE]
> - Database K1 và K2 được lưu **riêng biệt** để tránh xung đột dữ liệu
> - Cột `so_ngay_lien_tuc` cho biết thuê bao đã bị suy hao bao nhiêu ngày liên tiếp
