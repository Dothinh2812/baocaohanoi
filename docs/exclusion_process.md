# Tài liệu Module exclusion_process.py

## Mô tả

Module xử lý **giảm trừ phiếu báo hỏng** và tạo báo cáo so sánh **trước/sau giảm trừ** cho các chỉ tiêu C1.1 và C1.2.

**Mục đích**: Khi có các phiếu báo hỏng cần loại trừ khỏi tính toán KPI (ví dụ: sự cố bất khả kháng, lỗi hệ thống...), module này tính toán lại các chỉ tiêu và so sánh kết quả trước/sau giảm trừ.

---

## Cấu trúc Module

```
exclusion_process.py (1278 dòng)
├── Hàm tiện ích
│   ├── load_exclusion_list()        # Đọc DS phiếu loại trừ
│   ├── extract_nvkt_name()          # Trích xuất tên NVKT từ TEN_KV
│   └── calculate_statistics()       # Tính thống kê theo NVKT
├── Báo cáo so sánh C1.1
│   ├── create_c11_comparison_report()      # C1.1 SM4 (BRCD không hẹn)
│   └── create_c11_sm2_comparison_report()  # C1.1 SM2 (sửa chữa 72h)
├── Báo cáo so sánh C1.2
│   ├── create_c12_comparison_report()      # C1.2 SM1 (hỏng lại)
│   ├── create_sm1_c12_excluded_file()      # Tạo file SM1-C12 sau GT
│   └── create_c12_ti_le_bao_hong_comparison_report()  # C1.2 SM4 (tỷ lệ báo hỏng)
└── Hàm chính
    └── process_exclusion_reports()  # Wrapper chạy tất cả
```

---

## File Input / Output

### 📥 File đầu vào (bắt buộc)

| File | Thư mục | Mô tả |
|------|---------|-------|
| `ds_phieu_loai_tru.xlsx` | du_lieu_tham_chieu | Danh sách BAOHONG_ID loại trừ |
| `SM4-C11.xlsx` | downloads/baocao_hanoi | C1.1 BRCD không hẹn |
| `SM2-C11.xlsx` | downloads/baocao_hanoi | C1.1 sửa chữa 72h |
| `SM1-C12.xlsx` | downloads/baocao_hanoi | C1.2 hỏng lại |
| `SM2-C12.xlsx` | downloads/baocao_hanoi | C1.2 tổng phiếu báo hỏng |

### File tham chiếu (tùy chọn)

| File | Thư mục | Mô tả |
|------|---------|-------|
| `Tonghop_thuebao_NVKT_DB_C12.xlsx` | du_lieu_tham_chieu | Tổng số thuê bao theo NVKT |

### 📤 File đầu ra (lưu vào `downloads/kq_sau_giam_tru/`)

| File | Nội dung |
|------|----------|
| `So_sanh_C11_SM4.xlsx` | So sánh C1.1 SM4 trước/sau giảm trừ |
| `So_sanh_C11_SM2.xlsx` | So sánh C1.1 SM2 trước/sau giảm trừ |
| `So_sanh_C12_SM1.xlsx` | So sánh C1.2 SM1 trước/sau giảm trừ |
| `SM1-C12_sau_giam_tru.xlsx` | File SM1-C12 đã loại trừ (cấu trúc gốc) |
| `SM4-C12-ti-le-su-co-dv-brcd.xlsx` | So sánh tỷ lệ báo hỏng BRCĐ |
| **`Tong_hop_giam_tru.xlsx`** | **Tổng hợp tất cả chỉ tiêu** |

---

## Các hàm chính

### `load_exclusion_list(exclusion_file)`
Đọc danh sách BAOHONG_ID cần loại trừ từ file Excel.

**Tham số:**
- `exclusion_file`: Đường dẫn file (mặc định: `du_lieu_tham_chieu/ds_phieu_loai_tru.xlsx`)

**Trả về:** `set` - Tập hợp các BAOHONG_ID

---

### `extract_nvkt_name(ten_kv)`
Trích xuất tên NVKT từ cột TEN_KV.

**Ví dụ:**
- `"Sơn Lộc 1 - Nguyễn Thành Sơn"` → `"Nguyễn Thành Sơn"`
- `"VNM3-Khuất Anh Chiến( VXN)"` → `"Khuất Anh Chiến"`

---

### `create_c11_comparison_report(exclusion_ids, output_dir)`
Tạo báo cáo so sánh C1.1 SM4 (Tỷ lệ phiếu sửa chữa BRCD đúng quy định - không hẹn).

**Output file:** `So_sanh_C11_SM4.xlsx`
- Sheet `So_sanh_chi_tiet`: Chi tiết theo NVKT
- Sheet `Thong_ke_tong_hop`: Tổng hợp chung
- Sheet `DS_phieu_loai_tru`: Danh sách phiếu bị loại

---

### `create_c11_sm2_comparison_report(exclusion_ids, output_dir)`
Tạo báo cáo so sánh C1.1 SM2 (Tỷ lệ phiếu sửa chữa BRCD trong 72h).

**Tiêu chí đạt:** Thời gian xử lý ≤ 72 giờ

**Output file:** `So_sanh_C11_SM2.xlsx`

---

### `create_c12_comparison_report(exclusion_ids, output_dir)`
Tạo báo cáo so sánh C1.2 SM1 (Tỷ lệ thuê bao báo hỏng lặp lại).

**Công thức:** `Tỷ lệ HLL = Số phiếu HLL / Số phiếu báo hỏng × 100`

**Output file:** `So_sanh_C12_SM1.xlsx`

---

### `process_exclusion_reports()`
**Hàm chính** - Chạy toàn bộ workflow giảm trừ:
1. Đọc danh sách loại trừ
2. Tạo thư mục output
3. Tạo các báo cáo so sánh
4. Tạo báo cáo tổng hợp

---

## Cách sử dụng

### Chạy độc lập

```bash
python exclusion_process.py
```

> [!WARNING]
> Cần đảm bảo các file input đã tồn tại trước khi chạy

### Import trong code

```python
from exclusion_process import process_exclusion_reports

# Chạy toàn bộ workflow
process_exclusion_reports()
```

### Tích hợp trong baocaohanoi.py

```python
ENABLE_EXCLUSION = True  # Bật tính năng giảm trừ

if ENABLE_EXCLUSION:
    process_exclusion_reports()
```

---

## Module sử dụng

| Module | Import | Điều kiện |
|--------|--------|-----------|
| `baocaohanoi.py` | `from exclusion_process import process_exclusion_reports` | `ENABLE_EXCLUSION = True` |
| `kpi_calculator.py` | Đọc file output để tính KPI sau giảm trừ | - |

---

## Output Example

```
✅ Đã đọc 50 mã BAOHONG_ID cần loại trừ
✅ Đã tạo thư mục xuất kết quả: downloads/kq_sau_giam_tru

================================================================================
TẠO BÁO CÁO SO SÁNH C1.1 (SM4-C11) TRƯỚC/SAU GIẢM TRỪ
================================================================================
✅ Đã đọc file, tổng số dòng thô: 500
✅ Đã loại trừ 30 phiếu, còn lại 470 phiếu
✅ Đã tạo báo cáo so sánh C1.1 (SM4-C11)
   - Tổng phiếu thô: 500
   - Phiếu loại trừ: 30
   - Tổng phiếu sau GT: 470
   - Tỷ lệ thô: 92.5% -> Sau GT: 95.2%

================================================================================
✅ HOÀN THÀNH TẠO BÁO CÁO SO SÁNH GIẢM TRỪ
   Kết quả được lưu tại: downloads/kq_sau_giam_tru
================================================================================
```

---

## Lưu ý

> [!IMPORTANT]
> File `ds_phieu_loai_tru.xlsx` phải có cột `BAOHONG_ID` chứa mã phiếu cần loại trừ

> [!NOTE]
> Module này chỉ được gọi khi `ENABLE_EXCLUSION = True` trong `baocaohanoi.py`

> [!TIP]
> Kết quả sau giảm trừ được dùng bởi `kpi_calculator.py` để tính KPI SAU GIẢM TRỪ
