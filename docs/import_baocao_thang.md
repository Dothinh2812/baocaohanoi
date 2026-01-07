# Tài liệu Module import_baocao_thang.py

## Mô tả

Script import dữ liệu báo cáo **theo THÁNG** từ Excel vào SQLite database. Database riêng biệt để lưu trữ số liệu theo tháng (1 bản ghi/tháng/đơn vị).

---

## So sánh với import_baocao.py

| Tiêu chí | import_baocao.py | import_baocao_thang.py |
|----------|------------------|------------------------|
| Database | `baocao_hanoi.db` | `baocao_hanoi_thang.db` |
| Khóa thời gian | `ngay_bao_cao` (YYYY-MM-DD) | `thang_bao_cao` (YYYY-MM) |
| Dữ liệu | 1 bản ghi/ngày/đơn vị | 1 bản ghi/tháng/đơn vị |
| Hành vi trùng | INSERT OR REPLACE | ON CONFLICT DO UPDATE |
| Mục đích | Lưu snapshot hàng ngày | Theo dõi xu hướng theo tháng |

---

## Cấu trúc Module

```
import_baocao_thang.py
├── Hàm tiện ích
│   ├── parse_month()             # Parse định dạng tháng
│   ├── get_or_create_don_vi()    # Lấy/tạo đơn vị
│   └── get_or_create_nvkt()      # Lấy/tạo NVKT
├── Hàm khởi tạo
│   └── init_database()           # Tạo schema database
├── Hàm import theo đơn vị
│   ├── import_c11()              # Import C1.1
│   ├── import_c12()              # Import C1.2
│   ├── import_c13()              # Import C1.3
│   └── import_c14()              # Import C1.4
├── Hàm import theo NVKT
│   ├── import_c14_nvkt()         # Import C1.4 chi tiết NVKT
│   ├── import_sm1c12_hll()       # Import SM1-C12 hỏng lại
│   ├── import_sm4c11_chitiet()   # Import SM4-C11 chi tiết
│   └── import_sm4c11_18h()       # Import SM4-C11 18h
└── Hàm chính
    └── import_baocao_thang()     # Wrapper function
```

---

## Định dạng tháng hỗ trợ

Hàm `parse_month()` hỗ trợ nhiều định dạng:

| Input | Output |
|-------|--------|
| `"Tháng 10/2025"` | `"2025-10"` |
| `"Tháng 01/2026"` | `"2026-01"` |
| `"2025-10"` | `"2025-10"` |
| `"10/2025"` | `"2025-10"` |
| `None` | Tháng hiện tại |

---

## Database Schema

### Cấu trúc chung

Tất cả các bảng đều có:
- `thang_bao_cao` (TEXT): Format YYYY-MM
- `created_at` (DATETIME): Thời điểm tạo
- `updated_at` (DATETIME): Thời điểm cập nhật cuối

### Bảng danh mục

| Bảng | Mô tả |
|------|-------|
| `don_vi` | Danh sách đơn vị |
| `nhan_vien_kt` | Danh sách NVKT |

### Bảng báo cáo

| Bảng | Khóa unique | Mô tả |
|------|-------------|-------|
| `bao_cao_c11` | thang + don_vi | C1.1 theo đơn vị |
| `bao_cao_c12` | thang + don_vi | C1.2 theo đơn vị |
| `bao_cao_c13` | thang + don_vi | C1.3 theo đơn vị |
| `bao_cao_c14` | thang + don_vi | C1.4 theo đơn vị |
| `bao_cao_c14_nvkt` | thang + nvkt | C1.4 theo NVKT |
| `bao_cao_sm1c12_hll` | thang + nvkt | SM1-C12 hỏng lại |
| `bao_cao_sm4c11_chitiet` | thang + nvkt | SM4-C11 chi tiết |
| `bao_cao_sm4c11_18h` | thang + nvkt | SM4-C11 18h |

---

## Cách sử dụng

### Command Line

```bash
# Import với tháng chỉ định
python import_baocao_thang.py --month "Tháng 10/2025"
python import_baocao_thang.py --month "2025-10"
python import_baocao_thang.py -m "10/2025"

# Import tháng hiện tại
python import_baocao_thang.py

# Chỉ khởi tạo database
python import_baocao_thang.py --init
```

### Tham số

| Tham số | Viết tắt | Mô tả |
|---------|----------|-------|
| `--month` | `-m` | Tháng báo cáo (nhiều định dạng) |
| `--init` | | Chỉ tạo schema, không import |

### Import trong code

```python
from import_baocao_thang import import_baocao_thang

# Import tháng 10/2025
import_baocao_thang("Tháng 10/2025")

# Import tháng hiện tại
import_baocao_thang()
```

---

## Điểm đặc biệt

### 1. Xử lý đơn vị "Tổng"
Tự động đổi tên đơn vị "Tổng" thành "TTVT Sơn Tây" để thống nhất với đơn vị cha.

### 2. UPSERT Pattern
Sử dụng `ON CONFLICT DO UPDATE` để:
- Insert bản ghi mới nếu chưa tồn tại
- Update bản ghi cũ nếu đã tồn tại
- Tự động cập nhật `updated_at`

```sql
INSERT INTO bao_cao_c11 (...) VALUES (...)
ON CONFLICT(thang_bao_cao, don_vi_id) DO UPDATE SET
    sm1_cl_chu_dong = excluded.sm1_cl_chu_dong,
    ...
    updated_at = CURRENT_TIMESTAMP
```

### 3. Kiểm tra file tồn tại
Mỗi hàm import đều kiểm tra file tồn tại trước khi đọc:
```python
if not file_path.exists():
    print(f"  ⚠ C1.1: File không tồn tại: {file_path}")
    return
```

---

## Output

```
Database: /path/to/baocao_hanoi_thang.db
Report directory: /path/to/downloads/baocao_hanoi
✓ Database schema initialized

Importing data for month: 2025-10
  ✓ C1.1: 5 records imported
  ✓ C1.2: 5 records imported
  ✓ C1.3: 5 records imported
  ✓ C1.4: 5 records imported
  ✓ C1.4 NVKT: 36 records imported
  ✓ SM1-C12 HLL: 36 records imported
  ✓ SM4-C11 Chi tiết: 36 records imported
  ✓ SM4-C11 18h: 36 records imported

✓ All reports imported successfully for 2025-10
```

---

## Lưu ý

> [!NOTE]
> **UPSERT**: Import nhiều lần cùng tháng sẽ cập nhật dữ liệu mới nhất, không tạo bản ghi trùng

> [!TIP]
> **Theo dõi xu hướng**: Dùng database này để so sánh KPI giữa các tháng
