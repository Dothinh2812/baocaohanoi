# Tài liệu Module kpi_calculator.py

## Mô tả

Module tính điểm KPI cho NVKT theo BSC Q4/2025 VNPT Hà Nội.

**Các chỉ tiêu được tính:**
- **C1.1**: Tỷ lệ sửa chữa chất lượng (TP1 + TP2)
- **C1.2**: Tỷ lệ sự cố dịch vụ (TP1 + TP2)
- **C1.4**: Độ hài lòng khách hàng
- **C1.5**: Tỷ lệ thiết lập dịch vụ đúng thời gian

---

## Cấu trúc Module

```
kpi_calculator.py
├── Hàm tính điểm thành phần
│   ├── tinh_diem_C11_TP1()    # C1.1 TP1: Sửa chữa chủ động (30%)
│   ├── tinh_diem_C11_TP2()    # C1.1 TP2: Báo hỏng đúng quy định (70%)
│   ├── tinh_diem_C12_TP1()    # C1.2 TP1: Báo hỏng lặp lại (50%)
│   ├── tinh_diem_C12_TP2()    # C1.2 TP2: Sự cố BRCĐ (50%)
│   ├── tinh_diem_C14()        # C1.4: Độ hài lòng KH
│   └── tinh_diem_C15()        # C1.5: Thiết lập dịch vụ
├── Hàm tiện ích
│   ├── chuan_hoa_ty_le()      # Chuẩn hóa tỷ lệ về 0-1
│   └── chuan_hoa_ten()        # Chuẩn hóa tên NVKT
├── Hàm đọc dữ liệu gốc
│   ├── doc_C11_TP1()          # Đọc SM2-C11.xlsx
│   ├── doc_C11_TP2()          # Đọc SM4-C11.xlsx
│   ├── doc_C12_TP1()          # Đọc SM1-C12.xlsx
│   ├── doc_C12_TP2()          # Đọc SM4-C12-ti-le-su-co-dv-brcd.xlsx
│   ├── doc_C14()              # Đọc c1.4_chitiet_report.xlsx
│   └── doc_C15()              # Đọc c1.5_chitiet_report.xlsx
├── Hàm đọc dữ liệu sau giảm trừ
│   ├── doc_C11_TP1_sau_giam_tru()
│   ├── doc_C11_TP2_sau_giam_tru()
│   ├── doc_C12_TP1_sau_giam_tru()
│   └── doc_C12_TP2_sau_giam_tru()
└── Hàm tính KPI tổng hợp
    ├── tinh_diem_kpi_nvkt()           # Tính KPI trước giảm trừ
    ├── tinh_diem_kpi_nvkt_sau_giam_tru()  # Tính KPI sau giảm trừ
    ├── tao_bao_cao_kpi()              # Wrapper tạo báo cáo
    ├── tao_bao_cao_kpi_sau_giam_tru() # Wrapper tạo báo cáo sau GT
    └── tao_bao_cao_so_sanh_kpi()      # So sánh trước/sau giảm trừ
```

---

## Công thức tính điểm

### C1.1 - Tỷ lệ sửa chữa chất lượng

#### Thành phần 1 (30%): Sửa chữa chủ động

| Kết quả | Điểm |
|---------|------|
| ≥ 99% | 5 |
| 90% - 99% | 1 + 4 × (KQ - 90%) / 9% |
| ≤ 90% | 1 |

#### Thành phần 2 (70%): Báo hỏng đúng quy định

| Kết quả | Điểm |
|---------|------|
| ≥ 99.5% | 5 |
| 89.5% - 99.5% | 1 + 4 × (KQ - 89.5%) / 10% |
| ≤ 89.5% | 1 |

**Điểm C1.1 = 0.3 × TP1 + 0.7 × TP2**

---

### C1.2 - Tỷ lệ sự cố dịch vụ

> [!NOTE]
> Các chỉ tiêu C1.2 là **càng thấp càng tốt** (ngược với C1.1)

#### Thành phần 1 (50%): Báo hỏng lặp lại

| Kết quả | Điểm |
|---------|------|
| ≤ 3% | 5 |
| 3% - 6% | 5 - 4 × (KQ - 3%) / 3% |
| ≥ 6% | 1 |

#### Thành phần 2 (50%): Sự cố BRCĐ

| Kết quả | Điểm |
|---------|------|
| ≤ 1.8% | 5 |
| 1.8% - 2.8% | 5 - 4 × (KQ - 1.8%) / 1% |
| ≥ 2.8% | 1 |

**Điểm C1.2 = 0.5 × TP1 + 0.5 × TP2**

---

### C1.4 - Độ hài lòng khách hàng

| Kết quả | Điểm |
|---------|------|
| ≥ 99.5% | 5 |
| 89.5% - 99.5% | 1 + 4 × (KQ - 89.5%) / 10% |
| ≤ 89.5% | 1 |

---

### C1.5 - Tỷ lệ thiết lập dịch vụ đạt thời gian quy định

| Kết quả | Điểm |
|---------|------|
| ≥ 99.5% | 5 |
| 89.5% - 99.5% | 1 + 4 × (KQ - 89.5%) / 10% |
| ≤ 89.5% | 1 |

---

## File dữ liệu đầu vào

### Dữ liệu gốc (Trước giảm trừ)

| Chỉ tiêu | File | Sheet | Thư mục |
|----------|------|-------|---------|
| C1.1 TP1 | `SM2-C11.xlsx` | TH_SM2 | downloads/baocao_hanoi |
| C1.1 TP2 | `SM4-C11.xlsx` | chi_tiet | downloads/baocao_hanoi |
| C1.2 TP1 | `SM1-C12.xlsx` | TH_SM1C12_HLL_Thang | downloads/baocao_hanoi |
| C1.2 TP2 | `SM4-C12-ti-le-su-co-dv-brcd.xlsx` | TH_C12_TiLeBaoHong | downloads/baocao_hanoi |
| C1.4 | `c1.4_chitiet_report.xlsx` | TH_HL_NVKT | downloads/baocao_hanoi |
| C1.5 | `c1.5_chitiet_report.xlsx` | KQ_C15_chitiet | downloads/baocao_hanoi |

### Dữ liệu sau giảm trừ

| Chỉ tiêu | File | Sheet | Thư mục |
|----------|------|-------|---------|
| C1.1 TP1 | `So_sanh_C11_SM2.xlsx` | So_sanh_chi_tiet | kq_sau_giam_tru |
| C1.1 TP2 | `So_sanh_C11_SM4.xlsx` | So_sanh_chi_tiet | kq_sau_giam_tru |
| C1.2 TP1 | `So_sanh_C12_SM1.xlsx` | So_sanh_chi_tiet | kq_sau_giam_tru |
| C1.2 TP2 | `SM4-C12-ti-le-su-co-dv-brcd.xlsx` | So_sanh_chi_tiet | kq_sau_giam_tru |

---

## File kết quả đầu ra

| File | Mô tả |
|------|-------|
| `KPI_NVKT_ChiTiet.xlsx` | Chi tiết điểm từng thành phần |
| `KPI_NVKT_TomTat.xlsx` | Tổng hợp điểm các chỉ tiêu chính |
| `So_sanh_KPI_truoc_sau_giam_tru.xlsx` | So sánh trước/sau giảm trừ |

---

## Cách sử dụng

### Chạy từ Command Line

```bash
python kpi_calculator.py
```

Chạy mặc định sẽ:
1. Tính KPI **trước giảm trừ** → lưu vào thư mục `KPI`
2. Tính KPI **sau giảm trừ** → lưu vào thư mục `KPI`
3. Tạo báo cáo **so sánh** trước/sau giảm trừ

### Import và sử dụng trong code

```python
from kpi_calculator import (
    tinh_diem_kpi_nvkt,
    tinh_diem_kpi_nvkt_sau_giam_tru,
    tao_bao_cao_so_sanh_kpi
)

# Tính KPI trước giảm trừ
df_kpi = tinh_diem_kpi_nvkt(
    data_folder="downloads/baocao_hanoi",
    output_folder="KPI"
)

# Tính KPI sau giảm trừ
df_kpi_sau_gt = tinh_diem_kpi_nvkt_sau_giam_tru(
    exclusion_folder="kq_sau_giam_tru",
    original_data_folder="downloads/baocao_hanoi",
    output_folder="KPI"
)

# So sánh trước/sau giảm trừ
df_compare = tao_bao_cao_so_sanh_kpi(
    data_folder="downloads/baocao_hanoi",
    exclusion_folder="kq_sau_giam_tru",
    output_folder="KPI"
)
```

### Sử dụng từng hàm tính điểm

```python
from kpi_calculator import tinh_diem_C11_TP1, tinh_diem_C12_TP2

# Tính điểm C1.1 TP1 với tỷ lệ 95%
diem = tinh_diem_C11_TP1(0.95)  # Kết quả: 3.22

# Tính điểm C1.2 TP2 với tỷ lệ 2.5%
diem = tinh_diem_C12_TP2(0.025)  # Kết quả: 2.2
```

---

## Lưu ý quan trọng

> [!WARNING]
> **Định dạng tỷ lệ**: 
> - Nếu tỷ lệ > 1 → tự động chia 100 (ví dụ: 95 → 0.95)
> - Hàm `chuan_hoa_ty_le()` xử lý tự động

> [!IMPORTANT]
> **Tên NVKT**: 
> - Tự động chuẩn hóa về Title Case
> - `"NGUYỄN VĂN A"` → `"Nguyễn Văn A"`
> - Tránh trùng lặp do nhập khác nhau

> [!NOTE]
> **Mẫu số = 0**:
> - C1.4: Nếu không có phiếu khảo sát → mặc định 100%
> - Các chỉ tiêu khác: Bỏ qua tính toán

---

## Liên kết với exclusion_process.py

Module `kpi_calculator.py` **không import trực tiếp** `exclusion_process.py`, mà **đọc các file output** do `exclusion_process.py` tạo ra.

### Quy trình hoạt động

```
┌─────────────────────────────────────┐
│   exclusion_process.py             │
│   (Chạy trước)                     │
└─────────────────┬───────────────────┘
                  │ Tạo file output
                  ▼
┌─────────────────────────────────────┐
│  downloads/kq_sau_giam_tru/        │
│  ├── So_sanh_C11_SM2.xlsx          │
│  ├── So_sanh_C11_SM4.xlsx          │
│  ├── So_sanh_C12_SM1.xlsx          │
│  └── SM4-C12-ti-le-su-co-dv-brcd.xlsx │
└─────────────────┬───────────────────┘
                  │ Đọc file
                  ▼
┌─────────────────────────────────────┐
│   kpi_calculator.py                │
│   (Chạy sau)                       │
└─────────────────────────────────────┘
```

### Các hàm đọc file sau giảm trừ

| Hàm | File đọc | Chỉ tiêu |
|-----|----------|----------|
| `doc_C11_TP1_sau_giam_tru()` | So_sanh_C11_SM2.xlsx | C1.1 TP1 |
| `doc_C11_TP2_sau_giam_tru()` | So_sanh_C11_SM4.xlsx | C1.1 TP2 |
| `doc_C12_TP1_sau_giam_tru()` | So_sanh_C12_SM1.xlsx | C1.2 TP1 |
| `doc_C12_TP2_sau_giam_tru()` | SM4-C12-ti-le-su-co-dv-brcd.xlsx | C1.2 TP2 |

### Tóm tắt

| Thông tin | Giá trị |
|-----------|---------|
| Import trực tiếp? | ❌ Không |
| Liên kết qua? | 📁 File output trong `downloads/kq_sau_giam_tru/` |
| Thứ tự chạy | 1️⃣ exclusion_process.py → 2️⃣ kpi_calculator.py |
| Chỉ tiêu sử dụng | C1.1 (TP1, TP2), C1.2 (TP1, TP2) |
| Chỉ tiêu không đổi | C1.4, C1.5 (vẫn đọc từ file gốc) |

