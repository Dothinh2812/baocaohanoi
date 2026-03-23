# 🔧 SỬA LỖI LOGIC GIẢM TRỪ - NGÀY 26/01/2026

## 📋 TÓM TẮT

**Vấn đề phát hiện:** Logic tính toán sau giảm trừ **SAI** - giảm cả tử số và mẫu số, thay vì chỉ giảm tử số.

**Nguyên tắc đúng:**
- **Mẫu số (Tổng phiếu):** GIỮ NGUYÊN
- **Tử số (Phiếu đạt/đạt yêu cầu):** CHỈ GIẢM các phiếu bị loại trừ

---

## ❌ VẤN ĐỀ TRƯỚC KHI SỬA

### Ví dụ với C1.1 SM4 (Sửa chữa báo hỏng):

**TRƯỚC giảm trừ:**
- Tổng phiếu báo hỏng: 1516
- Phiếu hoàn thành đúng hạn: 1286
- Tỷ lệ: 1286/1516 = 84.83%

**SAU giảm trừ (LOGIC CŨ - SAI):**
- Loại trừ 367 phiếu (trong đó 268 phiếu đạt)
- Tổng phiếu: 1516 - 367 = 1149 ❌ (MẪU SỐ BỊ GIẢM!)
- Phiếu đạt: 1286 - 268 = 1018
- Tỷ lệ: 1018/1149 = 88.60% ❌ **SAI HOÀN TOÀN!**

**SAU giảm trừ (LOGIC MỚI - ĐÚNG):**
- Tổng phiếu: 1516 ✅ (GIỮ NGUYÊN MẪU SỐ)
- Phiếu đạt: 1286 - 268 = 1018
- Tỷ lệ: 1018/1516 = 67.15% ✅ **ĐÚNG!**

---

## 🔧 CÁC THAY ĐỔI ĐÃ THỰC HIỆN

### 1. Tạo Hàm Mới: `calculate_statistics_keep_denominator()`

**Vị trí:** `exclusion_process.py:794-865`

**Chức năng:**
Tính toán thống kê SAU giảm trừ với logic ĐÚNG:
- Nhận `df_before_stats` (thống kê TRƯỚC giảm trừ)
- Nhận `df_raw` và `exclusion_ids`
- GIỮ NGUYÊN mẫu số từ `df_before_stats`
- CHỈ GIẢM tử số dựa trên số phiếu đạt bị loại trừ

**Code:**
```python
def calculate_statistics_keep_denominator(df_before_stats, df_raw, exclusion_ids,
                                          has_ten_doi=True, dat_column='DAT_TT_KO_HEN', dat_value=1):
    """
    Tính toán thống kê SAU giảm trừ với LOGIC ĐÚNG:
    - Mẫu số (Tổng phiếu): GIỮ NGUYÊN từ df_before_stats
    - Tử số (Phiếu đạt): Giảm đi số phiếu đạt bị loại trừ
    """
    # ...implementation...

    # QUAN TRỌNG:
    tong_phieu_sau = tong_phieu_truoc  # GIỮ NGUYÊN MẪU SỐ
    so_phieu_dat_sau = so_phieu_dat_truoc - excluded_dat_count  # CHỈ GIẢM TỬ SỐ
```

---

### 2. Sửa `create_c11_comparison_report()` - C1.1 SM4

**Vị trí:** `exclusion_process.py:1018`

**Thay đổi:**
```python
# CŨ (SAI):
df_stats_after = calculate_statistics(df_excluded, has_ten_doi)

# MỚI (ĐÚNG):
df_stats_after = calculate_statistics_keep_denominator(
    df_stats_before, df_raw, exclusion_ids, has_ten_doi
)
```

---

### 3. Sửa `create_c11_sm2_comparison_report()` - C1.1 SM2 (TP1)

**Vị trí:** `exclusion_process.py:1274`

**Thay đổi:**
```python
# CŨ (SAI):
df_stats_after = calculate_statistics(df_excluded, has_ten_doi, dat_column='PHIEU_DAT', dat_value=1)

# MỚI (ĐÚNG):
df_stats_after = calculate_statistics_keep_denominator(
    df_stats_before, df_raw, exclusion_ids, has_ten_doi,
    dat_column='PHIEU_DAT', dat_value=1
)
```

---

### 4. Kiểm Tra `create_sm1_c12_excluded_file()` - C1.2 SM1

**Vị trí:** `exclusion_process.py:1629-1858`

**Kết luận:** ✅ **ĐÃ ĐÚNG** - Hàm này đã sử dụng logic đúng từ đầu:
- Dòng 1726-1736: Đọc SM2-C12 KHÔNG giảm trừ (mẫu số giữ nguyên)
- Dòng 1696-1709: Số phiếu HLL tính từ `df_sm1_excluded` (đã giảm trừ - tử số giảm)
- Dòng 1783: Tỷ lệ = Số phiếu HLL (sau GT) / Số phiếu báo hỏng (trước GT - giữ nguyên)

---

### 5. Kiểm Tra `create_c12_ti_le_bao_hong_comparison_report()` - C1.2 SM4

**Vị trí:** `exclusion_process.py:1859-2140`

**Kết luận:** ✅ **ĐÃ ĐÚNG** - Hàm này đã sử dụng logic đúng:
- Dòng 1936: Đếm số phiếu báo hỏng từ `df_excluded` (tử số giảm)
- Dòng 1993: Merge với `df_ref_clean` (Tổng TB từ file tham chiếu - mẫu số không đổi)

---

### 6. Sửa `create_c14_comparison_report()` - C1.4 (Độ hài lòng KH)

**Vị trí:** `exclusion_process.py:2321-2349`

**Thay đổi:**
```python
# CŨ (SAI):
df_after = df_excluded.groupby(...).agg({
    'IS_KHL': 'sum',
    'BAOHONG_ID': 'size'  # ❌ MẪU SỐ BỊ GIẢM
}).reset_index()

# MỚI (ĐÚNG):
# CHỈ tính số phiếu KHL từ dữ liệu đã loại trừ
df_khl_after = df_excluded.groupby(...).agg({
    'IS_KHL': 'sum'  # Chỉ đếm số phiếu KHL sau loại trừ
}).reset_index()

# Merge với df_before để GIỮ NGUYÊN Tổng phiếu KS (mẫu số)
df_after = pd.merge(
    df_before[merge_cols + ['Tổng phiếu KS']],  # ✅ Lấy mẫu số từ TRƯỚC GT
    df_khl_after,  # Lấy số phiếu KHL từ SAU GT
    on=merge_cols, how='left'
)
```

---

### 7. Sửa `create_c15_comparison_report()` - C1.5 (Thiết lập dịch vụ)

**Vị trí:** `exclusion_process.py:2589-2617`

**Thay đổi:**
```python
# CŨ (SAI):
df_after = df_excluded.groupby(...).agg({
    'IS_DAT': 'sum',
    'HDTB_ID': 'size'  # ❌ MẪU SỐ BỊ GIẢM
}).reset_index()

# MỚI (ĐÚNG):
# CHỈ tính số phiếu đạt từ dữ liệu đã loại trừ
df_dat_after = df_excluded.groupby(...).agg({
    'IS_DAT': 'sum'  # Chỉ đếm số phiếu đạt sau loại trừ
}).reset_index()

# Merge với df_before để GIỮ NGUYÊN Tổng Hoàn công (mẫu số)
df_after = pd.merge(
    df_before[merge_cols + ['Tổng Hoàn công']],  # ✅ Lấy mẫu số từ TRƯỚC GT
    df_dat_after,  # Lấy số phiếu đạt từ SAU GT
    on=merge_cols, how='left'
)
```

---

## 📊 TỔNG KẾT CÁC FILE ĐÃ SỬA

| File | Hàm | Trạng thái | Ghi chú |
|------|-----|-----------|---------|
| `exclusion_process.py` | `calculate_statistics_keep_denominator()` | ✅ MỚI | Hàm tính toán mới với logic đúng |
| `exclusion_process.py` | `create_c11_comparison_report()` | ✅ ĐÃ SỬA | C1.1 SM4 - Sửa chữa báo hỏng |
| `exclusion_process.py` | `create_c11_sm2_comparison_report()` | ✅ ĐÃ SỬA | C1.1 SM2 - Sửa chữa chủ động |
| `exclusion_process.py` | `create_sm1_c12_excluded_file()` | ✅ ĐÃ ĐÚNG | C1.2 SM1 - Không cần sửa |
| `exclusion_process.py` | `create_c12_ti_le_bao_hong_comparison_report()` | ✅ ĐÃ ĐÚNG | C1.2 SM4 - Không cần sửa |
| `exclusion_process.py` | `create_c14_comparison_report()` | ✅ ĐÃ SỬA | C1.4 - Độ hài lòng KH |
| `exclusion_process.py` | `create_c15_comparison_report()` | ✅ ĐÃ SỬA | C1.5 - Thiết lập dịch vụ |

---

## 🎯 NGUYÊN TẮC TÍNH TOÁN SAU GIẢM TRỪ

### Tất cả các chỉ tiêu C1.x:

| Chỉ tiêu | Công thức | Mẫu số (Sau GT) | Tử số (Sau GT) |
|---------|----------|-----------------|----------------|
| **C1.1 TP1** | Tỷ lệ SCCD ≤72h | Tổng phiếu SCCD (GIỮ NGUYÊN) | Phiếu hoàn thành - Phiếu đạt bị loại trừ |
| **C1.1 TP2** | Tỷ lệ BH đúng hạn | Tổng phiếu BH (GIỮ NGUYÊN) | Phiếu hoàn thành - Phiếu đạt bị loại trừ |
| **C1.2 TP1** | Tỷ lệ HLL | Tổng phiếu BH (GIỮ NGUYÊN) | Số HLL - HLL bị loại trừ |
| **C1.2 TP2** | Tỷ lệ sự cố BRCĐ | Tổng TB (từ file tham chiếu - GIỮ NGUYÊN) | Số phiếu BH - Phiếu bị loại trừ |
| **C1.4** | Độ hài lòng | Tổng phiếu KS (GIỮ NGUYÊN) | (Tổng - KHL) với KHL giảm |
| **C1.5** | Thiết lập đúng hạn | Tổng Hoàn công (GIỮ NGUYÊN) | Phiếu đạt - Phiếu đạt bị loại trừ |

### Quy tắc chung:

```
TRƯỚC giảm trừ:
  Tỷ lệ = Tử số (trước) / Mẫu số (trước) × 100

SAU giảm trừ (ĐÚNG):
  Mẫu số (sau) = Mẫu số (trước)  ← GIỮ NGUYÊN
  Tử số (sau) = Tử số (trước) - Số phiếu đạt bị loại trừ  ← CHỈ GIẢM TỬ
  Tỷ lệ = Tử số (sau) / Mẫu số (sau) × 100
```

---

## 🧪 CÁCH KIỂM TRA KẾT QUẢ

### 1. Chạy lại quy trình giảm trừ:

```bash
cd /home/vtst/baocaohanoi
python baocaohanoi.py
```

### 2. Kiểm tra file So_sanh_*.xlsx:

**Ví dụ:** `downloads/kq_sau_giam_tru/So_sanh_C11_SM4.xlsx`

**Sheet:** `Thong_ke_theo_don_vi`

Kiểm tra:
```
Đơn vị: TTVT Sơn Tây
Tổng phiếu (Thô): 1516
Tổng phiếu (Sau GT): 1516  ← PHẢI BẰNG NHAU!
Phiếu đạt (Thô): 1286
Phiếu đạt (Sau GT): 1018   ← CHỈ TỬ SỐ GIẢM
Tỷ lệ % (Thô): 84.83%
Tỷ lệ % (Sau GT): 67.15%   ← TỶ LỆ GIẢM DO TỬ SỐ GIẢM
```

### 3. Kiểm tra báo cáo Word:

**File:** `downloads/reports/Bao_cao_KPI_NVKT_01_2026.docx`

**Bảng 1.2: Thống kê điểm BSC theo đơn vị**

Kiểm tra cột "Trước" và "Sau" cho mỗi chỉ tiêu:
- Điểm "Sau" phải thấp hơn hoặc bằng điểm "Trước"
- Nếu điểm "Sau" cao hơn điểm "Trước" → SAI LOGIC!

---

## 📝 GHI CHÚ QUAN TRỌNG

1. **File C1.2 đặc biệt:**
   - C1.2 TP1 (SM1): Logic đã đúng từ đầu - không cần sửa
   - C1.2 TP2 (SM4): Logic đã đúng từ đầu - không cần sửa
   - Cả 2 đều GIỮ NGUYÊN mẫu số (SM2-C12 hoặc Tổng TB từ file tham chiếu)

2. **Không áp dụng cho C1.3:**
   - C1.3 (Kênh TSL) không có giảm trừ

3. **Tác động lên điểm BSC:**
   - Sau khi sửa logic, điểm BSC sau giảm trừ sẽ **GIẢM ĐI ĐÁNG KỂ**
   - Điều này là **ĐÚNG** vì phản ánh đúng tác động của việc loại trừ các phiếu đạt

---

## ⚠️ CẢNH BÁO

**TRƯỚC KHI CHẠY LẠI:**
- Backup các file báo cáo hiện tại
- So sánh kết quả trước/sau khi sửa
- Xác nhận với người dùng về logic mới

**SAU KHI SỬA:**
- Tỷ lệ sau giảm trừ sẽ GIẢM (đúng như thực tế)
- Điểm BSC sau giảm trừ sẽ GIẢM
- Điều này phản ánh ĐÚNG tác động của việc loại trừ

---

**Ngày cập nhật:** 2026-01-26
**Người thực hiện:** Claude Sonnet 4.5
**Trạng thái:** ✅ HOÀN THÀNH
