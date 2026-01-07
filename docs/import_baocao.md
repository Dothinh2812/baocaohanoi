# Tài liệu Module import_baocao.py

## Mô tả

Script import dữ liệu báo cáo từ file Excel vào SQLite database. Hỗ trợ lưu trữ dữ liệu hàng ngày để tạo báo cáo lịch sử.

---

## Cấu trúc Module

```
import_baocao.py
├── Hàm khởi tạo
│   └── init_database()           # Tạo schema database
├── Hàm tiện ích
│   ├── get_or_create_don_vi()    # Lấy/tạo đơn vị
│   └── get_or_create_nvkt()      # Lấy/tạo NVKT
├── Hàm import theo đơn vị
│   ├── import_c11()              # Import C1.1
│   ├── import_c12()              # Import C1.2
│   ├── import_c13()              # Import C1.3
│   └── import_c14()              # Import C1.4
└── Hàm import theo NVKT
    ├── import_c14_nvkt()         # Import C1.4 chi tiết NVKT
    ├── import_sm1c12_hll()       # Import SM1-C12 hỏng lại
    ├── import_sm4c11_chitiet()   # Import SM4-C11 chi tiết
    └── import_sm4c11_18h()       # Import SM4-C11 18h
```

---

## Cấu hình

| Biến | Giá trị | Mô tả |
|------|---------|-------|
| `DB_PATH` | `baocao_hanoi.db` | File SQLite database |
| `REPORT_DIR` | `downloads/baocao_hanoi` | Thư mục chứa file Excel |

### Mapping File Excel

| Key | File | Sheet |
|-----|------|-------|
| `c11` | c1.1 report.xlsx | TH_C1.1 |
| `c12` | c1.2 report.xlsx | TH_C1.2 |
| `c13` | c1.3 report.xlsx | TH_C1.3 |
| `c14` | c1.4 report.xlsx | TH_C1.4 |
| `c14_nvkt` | c1.4_chitiet_report.xlsx | TH_HL_NVKT |
| `sm1c12` | SM1-C12.xlsx | TH_SM1C12_HLL_Thang |
| `sm4c11_chitiet` | SM4-C11.xlsx | chi_tiet |
| `sm4c11_18h` | SM4-C11.xlsx | chi_tieu_ko_hen_18h |

---

## Database Schema

### Bảng danh mục

#### `don_vi`
| Cột | Kiểu | Mô tả |
|-----|------|-------|
| id | INTEGER | Primary key |
| ten_don_vi | TEXT | Tên đơn vị (unique) |

#### `nhan_vien_kt`
| Cột | Kiểu | Mô tả |
|-----|------|-------|
| id | INTEGER | Primary key |
| don_vi_id | INTEGER | FK → don_vi |
| ten_nvkt | TEXT | Tên NVKT |

### Bảng báo cáo theo đơn vị

#### `bao_cao_c11` - Tỷ lệ sửa chữa chất lượng
| Cột | Mô tả |
|-----|-------|
| ngay_bao_cao | Ngày báo cáo |
| sm1_cl_chu_dong, sm2_cl_chu_dong | Số liệu chủ động |
| ty_le_cl_chu_dong | Tỷ lệ chủ động |
| sm3_brcd, sm4_brcd | Số liệu BRCD |
| ty_le_brcd | Tỷ lệ BRCD |
| chi_tieu_bsc | Điểm BSC |

#### `bao_cao_c12` - Tỷ lệ sự cố dịch vụ
| Cột | Mô tả |
|-----|-------|
| sm1_lap_lai, sm2_lap_lai | Số liệu lặp lại |
| ty_le_lap_lai | Tỷ lệ lặp lại |
| sm3_su_co, sm4_su_co | Số liệu sự cố |
| ty_le_su_co | Tỷ lệ sự cố |

#### `bao_cao_c13` - Tỷ lệ sửa chữa Internet
| Cột | Mô tả |
|-----|-------|
| sm1_sua_chua, sm2_sua_chua | Số liệu sửa chữa |
| sm3_lap_lai, sm4_lap_lai | Số liệu lặp lại |
| sm5_su_co, sm6_su_co | Số liệu sự cố |

#### `bao_cao_c14` - Độ hài lòng khách hàng
| Cột | Mô tả |
|-----|-------|
| tong_phieu | Tổng số phiếu |
| sl_da_ks, sl_ks_thanh_cong | Số lượng khảo sát |
| sl_kh_hai_long | Số KH hài lòng |
| ty_le_kh_hai_long | Tỷ lệ hài lòng |
| diem_bsc | Điểm BSC |

### Bảng báo cáo theo NVKT

#### `bao_cao_c14_nvkt`
| Cột | Mô tả |
|-----|-------|
| nvkt_id | FK → nhan_vien_kt |
| tong_phieu_ks_thanh_cong | Số phiếu KS thành công |
| tong_phieu_khl | Số phiếu KHL |
| ty_le_hai_long | Tỷ lệ hài lòng |

#### `bao_cao_sm1c12_hll` - Hỏng lại theo NVKT
| Cột | Mô tả |
|-----|-------|
| nvkt_id | FK → nhan_vien_kt |
| so_phieu_hll | Số phiếu hỏng lại |
| so_phieu_bao_hong | Tổng phiếu báo hỏng |
| ty_le_hll | Tỷ lệ hỏng lại |

#### `bao_cao_sm4c11_chitiet` & `bao_cao_sm4c11_18h`
| Cột | Mô tả |
|-----|-------|
| nvkt_id | FK → nhan_vien_kt |
| tong_phieu | Tổng số phiếu |
| so_phieu_dat | Số phiếu đạt |
| ty_le_dat | Tỷ lệ đạt |

---

## Cách sử dụng

### Command Line

```bash
# Import với ngày hiện tại
python import_baocao.py

# Import với ngày chỉ định
python import_baocao.py --date 2024-12-30

# Chỉ khởi tạo database (không import)
python import_baocao.py --init

# Xem help
python import_baocao.py --help
```

### Tham số

| Tham số | Viết tắt | Mô tả |
|---------|----------|-------|
| `--date` | `-d` | Ngày báo cáo (YYYY-MM-DD) |
| `--init` | | Chỉ tạo schema, không import |

### Import trong code

```python
import sqlite3
from import_baocao import (
    init_database,
    import_c11,
    import_c12,
    import_c14_nvkt
)

conn = sqlite3.connect("baocao_hanoi.db")

# Khởi tạo database
init_database(conn)

# Import từng báo cáo
import_c11(conn, "2024-12-30")
import_c12(conn, "2024-12-30")
import_c14_nvkt(conn, "2024-12-30")

conn.close()
```

---

## Output

```
Database: /path/to/baocao_hanoi.db
Report directory: /path/to/downloads/baocao_hanoi
✓ Database schema initialized

Importing data for date: 2024-12-30
  ✓ C1.1: 5 records imported
  ✓ C1.2: 5 records imported
  ✓ C1.3: 5 records imported
  ✓ C1.4: 5 records imported
  ✓ C1.4 NVKT: 36 records imported
  ✓ SM1-C12 HLL: 36 records imported
  ✓ SM4-C11 Chi tiết: 36 records imported
  ✓ SM4-C11 18h: 36 records imported

✓ All reports imported successfully for 2024-12-30
```

---

## Lưu ý

> [!NOTE]
> **INSERT OR REPLACE**: Dữ liệu cùng ngày + đơn vị/NVKT sẽ được ghi đè khi import lại

> [!WARNING]
> **File Excel**: Đảm bảo các file Excel tồn tại và có đúng sheet name trước khi chạy

> [!TIP]
> **Lịch sử**: Mỗi lần import với ngày khác nhau sẽ tạo bản ghi mới, cho phép tra cứu lịch sử theo thời gian
