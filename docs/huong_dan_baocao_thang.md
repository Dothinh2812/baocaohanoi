# Hướng dẫn sử dụng Module Báo cáo Tháng

## Tổng quan

Hệ thống bao gồm 3 module chính để quản lý dữ liệu báo cáo theo tháng:

| Module | Mục đích | Database |
|--------|----------|----------|
| `import_baocao.py` | Import dữ liệu báo cáo theo **ngày** | `baocao_hanoi.db` |
| `import_baocao_thang.py` | Import dữ liệu báo cáo theo **tháng** | `baocao_hanoi_thang.db` |
| `xuat_baocao_xuhung.py` | Xuất báo cáo xu hướng qua các tháng | Đọc từ `baocao_hanoi_thang.db` |

---

## 1. import_baocao.py - Import dữ liệu theo ngày

### Mục đích
- Lưu trữ dữ liệu báo cáo **hàng ngày** để tạo lịch sử chi tiết
- Mỗi ngày import sẽ tạo một bản ghi mới

### Cách sử dụng

```bash
# Import với ngày hiện tại
python3 import_baocao.py

# Import với ngày chỉ định
python3 import_baocao.py --date 2025-12-31

# Chỉ khởi tạo database (không import)
python3 import_baocao.py --init
```

### Database: `baocao_hanoi.db`

| Bảng | Mô tả | UNIQUE key |
|------|-------|------------|
| `bao_cao_c11` | Báo cáo C1.1 | `(ngay_bao_cao, don_vi_id)` |
| `bao_cao_c12` | Báo cáo C1.2 | `(ngay_bao_cao, don_vi_id)` |
| `bao_cao_c13` | Báo cáo C1.3 | `(ngay_bao_cao, don_vi_id)` |
| `bao_cao_c14` | Báo cáo C1.4 | `(ngay_bao_cao, don_vi_id)` |
| `bao_cao_c14_nvkt` | C1.4 chi tiết NVKT | `(ngay_bao_cao, nvkt_id)` |
| `bao_cao_sm1c12_hll` | SM1-C12 Hỏng lại | `(ngay_bao_cao, nvkt_id)` |
| `bao_cao_sm4c11_chitiet` | SM4-C11 chi tiết | `(ngay_bao_cao, nvkt_id)` |
| `bao_cao_sm4c11_18h` | SM4-C11 18h | `(ngay_bao_cao, nvkt_id)` |

---

## 2. import_baocao_thang.py - Import dữ liệu theo tháng

### Mục đích
- Lưu trữ dữ liệu báo cáo **theo tháng** (1 bản ghi/tháng/đơn vị)
- Hỗ trợ so sánh xu hướng giữa các tháng
- Cơ chế UPSERT: import lại cùng tháng sẽ **cập nhật** dữ liệu

### Cách sử dụng

```bash
# Import tháng hiện tại
python3 import_baocao_thang.py

# Import tháng chỉ định (nhiều định dạng)
python3 import_baocao_thang.py --month "Tháng 10/2025"
python3 import_baocao_thang.py --month "2025-11"
python3 import_baocao_thang.py --month "12/2025"

# Chỉ khởi tạo database
python3 import_baocao_thang.py --init
```

### Gọi từ Python code

```python
from import_baocao_thang import import_baocao_thang

# Import với tháng từ cấu hình
import_baocao_thang("Tháng 10/2025")
```

### Database: `baocao_hanoi_thang.db`

| Bảng | Mô tả | UNIQUE key |
|------|-------|------------|
| `bao_cao_c11` | Báo cáo C1.1 | `(thang_bao_cao, don_vi_id)` |
| `bao_cao_c12` | Báo cáo C1.2 | `(thang_bao_cao, don_vi_id)` |
| `bao_cao_c13` | Báo cáo C1.3 | `(thang_bao_cao, don_vi_id)` |
| `bao_cao_c14` | Báo cáo C1.4 | `(thang_bao_cao, don_vi_id)` |
| `bao_cao_c14_nvkt` | C1.4 chi tiết NVKT | `(thang_bao_cao, nvkt_id)` |
| `bao_cao_sm1c12_hll` | SM1-C12 Hỏng lại | `(thang_bao_cao, nvkt_id)` |
| `bao_cao_sm4c11_chitiet` | SM4-C11 chi tiết | `(thang_bao_cao, nvkt_id)` |
| `bao_cao_sm4c11_18h` | SM4-C11 18h | `(thang_bao_cao, nvkt_id)` |

### Quy trình import tháng mới

```
┌─────────────────────────────────────────────────────────────────────┐
│ Bước 1: Cấu hình tháng trong baocaohanoi.py                        │
│   REPORT_MONTH = "Tháng 11/2025"                                   │
│   REPORT_DATE_CHITIET = "01/11/2025"                               │
│   END_DATE_CHITIET = "30/11/2025"                                  │
│                                                                     │
│ Bước 2: Chạy baocaohanoi.py để tải báo cáo                         │
│   python3 baocaohanoi.py                                           │
│                                                                     │
│ Bước 3: Import vào database với tháng tương ứng                    │
│   python3 import_baocao_thang.py --month "Tháng 11/2025"           │
└─────────────────────────────────────────────────────────────────────┘
```

### Lưu ý quan trọng

> ⚠️ **File Excel trong `downloads/baocao_hanoi/` bị ghi đè mỗi lần tải**
> 
> Bạn phải import vào database **NGAY SAU KHI tải** trước khi đổi tháng và tải lại.

---

## 3. xuat_baocao_xuhung.py - Xuất báo cáo xu hướng

### Mục đích
- Xuất file Excel thể hiện xu hướng các chỉ tiêu qua các tháng
- Tạo biểu đồ bar chart để visualization
- Hỗ trợ 2 góc nhìn: theo chỉ tiêu và theo đơn vị/NVKT

### Cách sử dụng

```bash
# Xuất với tên file tự động (timestamp)
python3 xuat_baocao_xuhung.py

# Xuất với tên file chỉ định
python3 xuat_baocao_xuhung.py --output "bao_cao_nam_2025.xlsx"
```

### Gọi từ Python code

```python
from xuat_baocao_xuhung import export_trend_report

export_trend_report()
# hoặc
export_trend_report("bao_cao_custom.xlsx")
```

### File output: `downloads/baocao_hanoi/bao_cao_xu_huong_<timestamp>.xlsx`

#### Các sheet theo chỉ tiêu (có biểu đồ):

| Sheet | Nội dung |
|-------|----------|
| `Tong_hop` | Tổng hợp các chỉ tiêu chính TTVT Sơn Tây |
| `C1.1_DonVi` | Xu hướng C1.1 theo đơn vị |
| `C1.2_DonVi` | Xu hướng C1.2 theo đơn vị |
| `C1.3_DonVi` | Xu hướng C1.3 theo đơn vị |
| `C1.4_DonVi` | Xu hướng C1.4 theo đơn vị |
| `C1.4_NVKT` | Xu hướng C1.4 theo NVKT |
| `SM1C12_HLL_NVKT` | Xu hướng Hỏng lại theo NVKT |
| `SM4C11_ChiTiet_NVKT` | Xu hướng BRCD chi tiết theo NVKT |
| `SM4C11_18h_NVKT` | Xu hướng BRCD 18h theo NVKT |

#### Các sheet theo đơn vị/NVKT (tất cả chỉ số):

| Sheet | Nội dung |
|-------|----------|
| `TH_DonVi_ChiTiet` | Mỗi dòng = 1 đơn vị + 1 tháng, tất cả chỉ số |
| `Pivot_DonVi` | Mỗi dòng = 1 đơn vị, cột = chỉ số theo tháng |
| `TH_NVKT_ChiTiet` | Mỗi dòng = 1 NVKT + 1 tháng, tất cả chỉ số |
| `Pivot_NVKT` | Mỗi dòng = 1 NVKT, cột = chỉ số theo tháng |

### Cột xu hướng (XH_TL_xxx)
- Giá trị = Tháng cuối - Tháng đầu
- **Dương** = Cải thiện (đối với C1.1 BRCD, C1.4 Hài lòng)
- **Âm** = Giảm sút (đối với C1.2 Lặp lại, SM1-C12 Hỏng lại)

---

## Quy trình làm việc hoàn chỉnh

```
┌─────────────────────────────────────────────────────────────────────┐
│                    QUY TRÌNH IMPORT HÀNG THÁNG                      │
├─────────────────────────────────────────────────────────────────────┤
│                                                                     │
│  Tháng 10/2025:                                                     │
│  1. Cấu hình REPORT_MONTH = "Tháng 10/2025"                         │
│  2. python3 baocaohanoi.py                                          │
│  3. python3 import_baocao_thang.py --month "Tháng 10/2025"          │
│                                                                     │
│  Tháng 11/2025:                                                     │
│  1. Cấu hình REPORT_MONTH = "Tháng 11/2025"                         │
│  2. python3 baocaohanoi.py                                          │
│  3. python3 import_baocao_thang.py --month "Tháng 11/2025"          │
│                                                                     │
│  Tháng 12/2025:                                                     │
│  1. Cấu hình REPORT_MONTH = "Tháng 12/2025"                         │
│  2. python3 baocaohanoi.py                                          │
│  3. python3 import_baocao_thang.py --month "Tháng 12/2025"          │
│                                                                     │
│  Xuất báo cáo xu hướng:                                             │
│  python3 xuat_baocao_xuhung.py                                      │
│                                                                     │
└─────────────────────────────────────────────────────────────────────┘
```

---

## Khởi tạo lại Database

### Xóa và khởi tạo lại database tháng:
```bash
rm baocao_hanoi_thang.db
python3 import_baocao_thang.py --init
```

### Xóa và khởi tạo lại database ngày:
```bash
rm baocao_hanoi.db
python3 import_baocao.py --init
```

---

## Các file liên quan

```
baocaohanoi/
├── baocaohanoi.py              # Script chính - tải báo cáo
├── import_baocao.py            # Import dữ liệu theo ngày
├── import_baocao_thang.py      # Import dữ liệu theo tháng
├── xuat_baocao_xuhung.py       # Xuất báo cáo xu hướng
├── baocao_hanoi.db             # Database lưu theo ngày
├── baocao_hanoi_thang.db       # Database lưu theo tháng
└── downloads/baocao_hanoi/     # Thư mục chứa file Excel
    ├── c1.1 report.xlsx
    ├── c1.2 report.xlsx
    ├── c1.3 report.xlsx
    ├── c1.4 report.xlsx
    ├── SM1-C12.xlsx
    ├── SM4-C11.xlsx
    └── bao_cao_xu_huong_*.xlsx  # File output xu hướng
```

---

## Troubleshooting

### Lỗi "File không tồn tại"
- Kiểm tra đã chạy `baocaohanoi.py` để tải báo cáo chưa
- Kiểm tra thư mục `downloads/baocao_hanoi/` có file Excel không

### Import lại cùng tháng
- Database tháng sử dụng cơ chế UPSERT (INSERT hoặc UPDATE)
- Dữ liệu mới sẽ ghi đè dữ liệu cũ của cùng tháng
- Cột `updated_at` sẽ được cập nhật

### Không có dữ liệu trong báo cáo xu hướng
- Kiểm tra đã import dữ liệu vào `baocao_hanoi_thang.db` chưa
- Kiểm tra `--month` parameter có khớp với tháng đã import không
