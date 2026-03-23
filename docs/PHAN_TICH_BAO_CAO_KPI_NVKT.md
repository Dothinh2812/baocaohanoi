# 📄 PHÂN TÍCH CHI TIẾT: Bao_cao_KPI_NVKT_01_2026.docx

## Mục lục

1. [Trang bìa](#trang-bìa)
2. [Phần 1: Tổng quan](#phần-1-tổng-quan)
3. [Phần 2: Chi tiết theo từng tổ](#phần-2-chi-tiết-theo-từng-tổ)
4. [Phần 3: Dữ liệu giảm trừ](#phần-3-sử-dụng-dữ-liệu-giảm-trừ-nếu-có)
5. [Tóm tắt bản đồ nguồn dữ liệu](#tóm-tắt-bản-đồ-nguồn-dữ-liệu)
6. [Quy trình tạo báo cáo](#-quy-trình-tạo-báo-cáo)

---

## CẤU TRÚC BÁO CÁO VÀ NGUỒN DỮ LIỆU

---

## TRANG BÌA

### 1. Tiêu đề: "BÁO CÁO KẾT QUẢ BSC/KPI THÁNG 01/2026"

- **Nguồn:** Tạo tự động từ biến `report_month` trong code
- **File:** `report_generator.py:4201`

### 2. Ngày tạo báo cáo

- **Nguồn:** Tạo tự động từ `datetime.now()`
- **File:** `report_generator.py:4199`

---

## PHẦN 1: TỔNG QUAN

### 1.1. Biểu đồ So Sánh Điểm BSC Thực Tế 4 Tổ (TRƯỚC GIẢM TRỪ)

**Loại:** Biểu đồ cột (Bar Chart)

**Nguồn Dữ Liệu:**

- **File chính:** `downloads/baocao_hanoi/c1.1 report.xlsx` → Sheet `TH_C1.1`
- **File bổ sung:**
  - `downloads/baocao_hanoi/c1.2 report.xlsx` → Sheet `TH_C1.2`
  - `downloads/baocao_hanoi/c1.3 report.xlsx` → Sheet `TH_C1.3`
  - `downloads/baocao_hanoi/c1.4 report.xlsx` → Sheet `TH_C1.4`
  - `downloads/baocao_hanoi/c1.5_chitiet_report.xlsx` → Sheet `TH_TTVTST`

**Cách Tính:**

- Load từ hàm `load_c1x_reports()` (`report_generator.py:111-166`)
- Hiển thị trên biểu đồ: Cột `Chỉ tiêu BSC` từ mỗi file
- Tạo biểu đồ bằng `create_team_comparison_chart()` (`report_generator.py:4230`)

---

### 1.1.b. Biểu đồ BSC Sau Giảm Trừ

**Loại:** Biểu đồ cột (Bar Chart)

**Nguồn Dữ Liệu:**

- **File:** `downloads/kq_sau_giam_tru/Tong_hop_Diem_BSC_Don_Vi.xlsx`
  - Sheet: `Tong_hop_Don_vi`
  - Cột sử dụng:
    - `don_vi` (Tên đơn vị)
    - `Diem_C1.1 (Sau)` (Điểm C1.1 sau giảm trừ)
    - `Diem_C1.2 (Sau)` (Điểm C1.2 sau giảm trừ)
    - `Diem_C1.4 (Sau)` (Điểm C1.4 sau giảm trừ)
    - `Diem_C1.5 (Sau)` (Điểm C1.5 sau giảm trừ)

**Cách Tính:**

- Load từ hàm `load_bsc_unit_scores_from_comparison()` (`report_generator.py:350-381`)
- Tạo biểu đồ bằng `create_team_bsc_after_exclusion_chart()` (`report_generator.py:4245`)

---

### 1.2. Bảng Thống Kê Điểm BSC Theo Đơn Vị

**Loại:** Bảng Word Table

**Cấu trúc:**

| Đơn vị | C1.1 (Trước/Sau) | C1.2 (Trước/Sau) | C1.3 | C1.4 (Trước/Sau) | C1.5 (Trước/Sau) |
|--------|------------------|------------------|------|------------------|------------------|

**Nguồn Dữ Liệu:**

#### Cột C1.1, C1.2, C1.4, C1.5 (Trước & Sau):

- **File:** `downloads/kq_sau_giam_tru/Tong_hop_Diem_BSC_Don_Vi.xlsx`
- **Sheet:** `Tong_hop_Don_vi`
- **Các cột:**
  - `Diem_C1.1 (Trước)`, `Diem_C1.1 (Sau)`
  - `Diem_C1.2 (Trước)`, `Diem_C1.2 (Sau)`
  - `Diem_C1.4 (Trước)`, `Diem_C1.4 (Sau)`
  - `Diem_C1.5 (Trước)`, `Diem_C1.5 (Sau)`

#### Cột C1.3 (Không có giảm trừ):

- **File:** `downloads/baocao_hanoi/c1.3 report.xlsx`
- **Sheet:** `TH_C1.3`
- **Cột:** `Chỉ tiêu BSC`

**Code Tham Chiếu:** `report_generator.py:4360-4503`

---

### 1.3. Số Liệu Chi Tiết Các Chỉ Tiêu BSC (Bảng và Biểu Đồ)

Phần này được tạo bởi hàm `add_c1x_overview_table()`. Mỗi chỉ tiêu có 2 bảng: **Trước giảm trừ** và **Sau giảm trừ**.

---

#### 1.3.1. C1.1 - Chất Lượng Sửa Chữa BRCĐ

##### Bảng C1.1 TP1 (Sửa Chữa Chủ Động - SM2) - TRƯỚC GIẢM TRỪ

**Nguồn:**

- **File:** `downloads/baocao_hanoi/c1.1 report.xlsx`
- **Sheet:** `TH_C1.1`
- **Cột sử dụng:**
  - `Đơn vị` → Tên tổ
  - `Đơn vị SCCD` → SM2 (Sửa chữa chủ động)
  - `Tổng phiếu SCCD` → Tổng số phiếu
  - `SCCD hoàn thành ≤72h` → Số phiếu hoàn thành đúng hạn
  - `Tỷ lệ (%)` → Tỷ lệ hoàn thành
  - `Chỉ tiêu BSC` → Điểm BSC

**Code:** Hàm `load_c1x_reports()` → key `'c11'`

##### Bảng C1.1 TP1 - SAU GIẢM TRỪ

**Nguồn:**

- **File:** `downloads/kq_sau_giam_tru/So_sanh_C11_SM2.xlsx`
- **Sheet:** `Thong_ke_theo_don_vi`
- **Cột sử dụng:**
  - `don_vi` → Tên tổ
  - `tong_phieu_sau_giam_tru` → Tổng phiếu sau loại trừ
  - `phieu_hoan_thanh_sau_giam_tru` → Phiếu hoàn thành sau loại trừ
  - `ty_le_sau_giam_tru` → Tỷ lệ % sau loại trừ
  - `diem_bsc_sau` → Điểm BSC sau giảm trừ

**Code:** `report_generator.py:278-307` → key `'c11_sm2'`

---

##### Bảng C1.1 TP2 (Sửa Chữa Báo Hỏng - SM4) - TRƯỚC GIẢM TRỪ

**Nguồn:**

- **File:** `downloads/baocao_hanoi/c1.1 report.xlsx`
- **Sheet:** `TH_C1.1`
- **Cột sử dụng:**
  - `Đơn vị` → Tên tổ
  - `Đơn vị BH` → SM4 (Báo hỏng)
  - `Tổng phiếu BH` → Tổng số phiếu báo hỏng
  - `BH hoàn thành đúng quy định` → Phiếu hoàn thành đúng hạn
  - `Tỷ lệ (%)` → Tỷ lệ hoàn thành
  - `Chỉ tiêu BSC` → Điểm BSC

##### Bảng C1.1 TP2 - SAU GIẢM TRỪ

**Nguồn:**

- **File:** `downloads/kq_sau_giam_tru/So_sanh_C11_SM4.xlsx`
- **Sheet:** `Thong_ke_theo_don_vi`
- **Cột tương tự như TP1**

**Code:** `report_generator.py:292-299` → key `'c11_sm4'`

---

#### 1.3.2. C1.2 - Tỷ Lệ Thuê Bao Báo Hỏng

##### Bảng C1.2 TP1 (Hỏng Lặp Lại ≥2 lần/7 ngày - SM1) - TRƯỚC GIẢM TRỪ

**Nguồn:**

- **File:** `downloads/baocao_hanoi/c1.2 report.xlsx`
- **Sheet:** `TH_C1.2`
- **Cột sử dụng:**
  - `Đơn vị` → Tên tổ
  - `SM1` → Chỉ báo SM1
  - `Tổng TB quản lý` → Tổng thuê bao quản lý
  - `TB báo hỏng ≥2 lần/7 ngày` → Số thuê bao hỏng lặp
  - `Tỷ lệ (%)` → Tỷ lệ hỏng lặp
  - `Chỉ tiêu BSC` → Điểm BSC

##### Bảng C1.2 TP1 - SAU GIẢM TRỪ

**Nguồn:**

- **File:** `downloads/kq_sau_giam_tru/So_sanh_C12_SM1.xlsx`
- **Sheet:** `Thong_ke_theo_don_vi`

**Code:** `report_generator.py:310-317` → key `'c12_sm1'`

---

##### Bảng C1.2 TP2 (Tỷ Lệ Sự Cố BRCĐ - SM4) - TRƯỚC GIẢM TRỪ

**Nguồn:**

- **File:** `downloads/baocao_hanoi/c1.2 report.xlsx`
- **Sheet:** `TH_C1.2`
- **Cột sử dụng:**
  - `Đơn vị` → Tên tổ
  - `SM4` → Chỉ báo SM4
  - `Tổng TB quản lý` → Tổng thuê bao
  - `Phiếu báo hỏng BRCĐ` → Số phiếu báo hỏng
  - `Tỷ lệ (‰)` → Tỷ lệ phần nghìn
  - `Chỉ tiêu BSC` → Điểm BSC

##### Bảng C1.2 TP2 - SAU GIẢM TRỪ

**Nguồn:**

- **File:** `downloads/kq_sau_giam_tru/SM4-C12-ti-le-su-co-dv-brcd.xlsx`
- **Sheet:** `Thong_ke_theo_don_vi`

**Code:** `report_generator.py:320-327` → key `'c12_sm4'`

---

#### 1.3.3. C1.3 - Chất Lượng Sửa Chữa Kênh TSL (Leased Line)

**Nguồn:**

- **File:** `downloads/baocao_hanoi/c1.3 report.xlsx`
- **Sheet:** `TH_C1.3`
- **Cột sử dụng:**
  - `Đơn vị` → Tên tổ
  - `Tổng phiếu` → Tổng phiếu TSL
  - `Hoàn thành đúng quy định` → Phiếu hoàn thành đúng hạn
  - `Tỷ lệ (%)` → Tỷ lệ hoàn thành
  - `Chỉ tiêu BSC` → Điểm BSC

**Lưu ý:** C1.3 **KHÔNG CÓ GIẢM TRỪ** (chỉ có bảng trước)

**Code:** `report_generator.py:139-146`

---

#### 1.3.4. C1.4 - Độ Hài Lòng Khách Hàng

##### Bảng C1.4 - TRƯỚC GIẢM TRỪ

**Nguồn:**

- **File:** `downloads/baocao_hanoi/c1.4 report.xlsx`
- **Sheet:** `TH_C1.4`
- **Cột sử dụng:**
  - `Đơn vị` → Tên tổ
  - `Tổng khảo sát` → Tổng số khảo sát
  - `Khách hàng hài lòng` → Số khách hài lòng
  - `Tỷ lệ (%)` → Tỷ lệ hài lòng
  - `Chỉ tiêu BSC` → Điểm BSC

##### Bảng C1.4 - SAU GIẢM TRỪ

**Nguồn:**

- **File:** `downloads/kq_sau_giam_tru/So_sanh_C14.xlsx`
- **Sheet:** `Thong_ke_theo_don_vi`

**Code:** `report_generator.py:330-336` → key `'c14'`

---

#### 1.3.5. C1.5 - Thiết Lập Dịch Vụ BRCĐ Đạt Thời Gian

##### Bảng C1.5 - TRƯỚC GIẢM TRỪ

**Nguồn:**

- **File:** `downloads/baocao_hanoi/c1.5_chitiet_report.xlsx`
- **Sheet:** `TH_TTVTST`
- **Cột sử dụng:**
  - `Đơn vị` → Tên tổ
  - `Tổng phiếu` → Tổng phiếu lắp đặt
  - `Hoàn thành đúng hạn` → Phiếu hoàn thành đúng hạn
  - `Tỷ lệ (%)` → Tỷ lệ hoàn thành
  - `Chỉ tiêu BSC` → Điểm BSC

##### Bảng C1.5 - SAU GIẢM TRỪ

**Nguồn:**

- **File:** `downloads/kq_sau_giam_tru/So_sanh_C15.xlsx`
- **Sheet:** `Thong_ke_theo_don_vi`

**Code:** `report_generator.py:338-345` → key `'c15'`

---

### 1.4. Tổng Quan Suy Hao Cao (SHC)

**Nguồn:**

- **File:** `downloads/baocao_hanoi/I15_Suyhao_*.xlsx` (multiple files)
- **Hoặc:** `suy_hao_cts.xlsx` (fallback)

**Cột sử dụng:**

- Tên đơn vị
- Số thuê bao suy hao cao
- Tỷ lệ SHC/TB quản lý

**Code:** Hàm `add_shc_overview_section()` (`report_generator.py:4517`)

---

### 1.5. Tổng Hợp Số Liệu Sau Giảm Trừ

#### Bảng So Sánh Trước/Sau Giảm Trừ

**Nguồn:**

- **File:** `downloads/kq_sau_giam_tru/Tong_hop_giam_tru.xlsx`
- **Cột:**
  - `Chi_tieu` → Tên chỉ tiêu (C1.1 TP1, C1.1 TP2, C1.2 TP1, C1.2 TP2, C1.4, C1.5)
  - `Truoc_giam_tru` → Số liệu trước loại trừ
  - `Sau_giam_tru` → Số liệu sau loại trừ
  - `Chenh_lech` → Chênh lệch (+/-)
  - `Phan_tram_thay_doi` → % thay đổi

**Code:** `report_generator.py:266-273`

#### Biểu Đồ Tỷ Lệ Sau Giảm Trừ

**Nguồn:** Cùng file `Tong_hop_giam_tru.xlsx` như trên

**Code:** Hàm `create_exclusion_bar_chart()` (`report_generator.py:4536`)

---

## PHẦN 2: CHI TIẾT THEO TỪNG TỔ

Phần này lặp lại cho 4 tổ: **Phúc Thọ, Quảng Oai, Suối Hai, Sơn Tây**

---

### 2.1. Bảng Điểm KPI Tổng Hợp (Trước Giảm Trừ)

**Nguồn:**

- **File:** `downloads/KPI/KPI_NVKT_TomTat.xlsx`
- **Cột sử dụng:**
  - `don_vi` → Tên tổ (filter theo tổ)
  - `nvkt` → Tên NVKT
  - `c1_1_score` → Điểm C1.1
  - `c1_2_score` → Điểm C1.2
  - `c1_3_score` → Điểm C1.3
  - `c1_4_score` → Điểm C1.4
  - `c1_5_score` → Điểm C1.5
  - `tong_diem_bsc` → Tổng điểm BSC

**Code:** `report_generator.py:101-102`

---

### 2.2. Bảng KPI Sau Giảm Trừ

**Nguồn:**

- **File:** `downloads/kq_sau_giam_tru/Tong_hop_Diem_BSC_Don_Vi.xlsx`
- **Sheet:** `Chi_tiet_Ca_nhan`
- **Cột sử dụng:**
  - `don_vi` → Tên tổ
  - `nvkt` → Tên NVKT
  - `Diem_C1.1 (Sau)` → Điểm C1.1 sau giảm trừ
  - `Diem_C1.2 (Sau)` → Điểm C1.2 sau giảm trừ
  - `Diem_C1.4 (Sau)` → Điểm C1.4 sau giảm trừ
  - `Diem_C1.5 (Sau)` → Điểm C1.5 sau giảm trừ

**Code:** `report_generator.py:376-378`

---

### 2.3. Biểu Đồ So Sánh NVKT (Trước Giảm Trừ)

**Loại:** Biểu đồ cột theo NVKT

**Nguồn:** Cùng file `KPI_NVKT_TomTat.xlsx` như mục 2.1

**Code:** Hàm `create_nvkt_bar_chart()` (`report_generator.py:4588`)

---

### 2.4. Biểu Đồ NVKT Sau Giảm Trừ

**Nguồn:** Cùng file như mục 2.2

**Code:** Hàm `create_nvkt_bar_chart_after_exclusion()` (`report_generator.py:4598`)

---

### 2.5. Bảng Chi Tiết KPI Từng NVKT

**Nguồn:**

- **File:** `downloads/KPI/KPI_NVKT_ChiTiet.xlsx`
- **Cột sử dụng (cho mỗi NVKT):**
  - `nvkt` → Tên NVKT
  - **C1.1 TP1:**
    - `c1_1_tp1_tong_phieu` → Tổng phiếu SCCD
    - `c1_1_tp1_hoan_thanh` → Phiếu hoàn thành ≤72h
    - `c1_1_tp1_ty_le` → Tỷ lệ %
    - `c1_1_tp1_diem` → Điểm thành phần
  - **C1.1 TP2:**
    - `c1_1_tp2_tong_phieu` → Tổng phiếu BH
    - `c1_1_tp2_hoan_thanh` → Phiếu BH hoàn thành
    - `c1_1_tp2_ty_le` → Tỷ lệ %
    - `c1_1_tp2_diem` → Điểm thành phần
  - **C1.2 TP1:**
    - `c1_2_tp1_tong_tb` → Tổng TB quản lý
    - `c1_2_tp1_tb_hong_lap` → TB hỏng ≥2 lần
    - `c1_2_tp1_ty_le` → Tỷ lệ %
    - `c1_2_tp1_diem` → Điểm thành phần
  - **C1.2 TP2:**
    - `c1_2_tp2_tong_tb` → Tổng TB
    - `c1_2_tp2_phieu_bh` → Phiếu báo hỏng
    - `c1_2_tp2_ty_le` → Tỷ lệ ‰
    - `c1_2_tp2_diem` → Điểm thành phần
  - **C1.4:**
    - `c1_4_tong_khao_sat` → Tổng khảo sát
    - `c1_4_hai_long` → Khách hài lòng
    - `c1_4_ty_le` → Tỷ lệ %
    - `c1_4_diem` → Điểm
  - **C1.5:**
    - `c1_5_tong_phieu` → Tổng phiếu lắp
    - `c1_5_dung_han` → Phiếu đúng hạn
    - `c1_5_ty_le` → Tỷ lệ %
    - `c1_5_diem` → Điểm

**Code:** `report_generator.py:105-106`

---

### 2.6. Bảng Chi Tiết NVKT Sau Giảm Trừ

**Nguồn:**

- **File:** `downloads/kq_sau_giam_tru/So_sanh_C11_SM4.xlsx` → Sheet `So_sanh_chi_tiet`
- **File:** `downloads/kq_sau_giam_tru/So_sanh_C11_SM2.xlsx` → Sheet `So_sanh_chi_tiet`
- **File:** `downloads/kq_sau_giam_tru/So_sanh_C12_SM1.xlsx` → Sheet `So_sanh_chi_tiet`
- **File:** `downloads/kq_sau_giam_tru/SM4-C12-ti-le-su-co-dv-brcd.xlsx` → Sheet `So_sanh_chi_tiet`
- **File:** `downloads/kq_sau_giam_tru/So_sanh_C14.xlsx` → Sheet `So_sanh_chi_tiet`
- **File:** `downloads/kq_sau_giam_tru/So_sanh_C15.xlsx` → Sheet `So_sanh_chi_tiet`

**Cột (ví dụ cho C1.1 SM4):**

- `NVKT_DB` → Tên NVKT
- `Tong_phieu_truoc` → Tổng phiếu trước loại trừ
- `Tong_phieu_sau` → Tổng phiếu sau loại trừ
- `Ty_le_truoc` → Tỷ lệ % trước
- `Ty_le_sau` → Tỷ lệ % sau
- `Diem_BSC_truoc` → Điểm trước
- `Diem_BSC_sau` → Điểm sau

**Code:** `report_generator.py:194-273`

---

## PHẦN 3: SỬ DỤNG DỮ LIỆU GIẢM TRỪ (Nếu có)

### 3.1. Danh Sách Phiếu Bị Loại Trừ

**Nguồn:**

- **File Input:** `du_lieu_tham_chieu/ds_phieu_loai_tru.xlsx`
- **Cột:**
  - `BAOHONG_ID` → ID phiếu báo hỏng cần loại trừ
  - `LY_DO` → Lý do loại trừ
  - `NGAY_LOAI_TRU` → Ngày loại trừ

**File này được sử dụng để lọc dữ liệu ở bước xử lý giảm trừ** (`exclusion_process.py`)

---

## TÓM TẮT BẢN ĐỒ NGUỒN DỮ LIỆU

### 📂 NGUỒN DỮ LIỆU GỐC (Trước Giảm Trừ)

| File | Sheet | Thành Phần Báo Cáo |
|------|-------|-------------------|
| `downloads/baocao_hanoi/c1.1 report.xlsx` | `TH_C1.1` | Bảng C1.1 TP1 + TP2, Biểu đồ BSC trước GT |
| `downloads/baocao_hanoi/c1.2 report.xlsx` | `TH_C1.2` | Bảng C1.2 TP1 + TP2 |
| `downloads/baocao_hanoi/c1.3 report.xlsx` | `TH_C1.3` | Bảng C1.3 (không có GT) |
| `downloads/baocao_hanoi/c1.4 report.xlsx` | `TH_C1.4` | Bảng C1.4 trước GT |
| `downloads/baocao_hanoi/c1.5_chitiet_report.xlsx` | `TH_TTVTST` | Bảng C1.5 trước GT |
| `downloads/KPI/KPI_NVKT_TomTat.xlsx` | Sheet 1 | Bảng KPI tổng hợp từng tổ (trước GT) |
| `downloads/KPI/KPI_NVKT_ChiTiet.xlsx` | Sheet 1 | Bảng chi tiết NVKT (trước GT) |

---

### 📂 NGUỒN DỮ LIỆU SAU GIẢM TRỪ

| File | Sheet | Thành Phần Báo Cáo |
|------|-------|-------------------|
| `downloads/kq_sau_giam_tru/So_sanh_C11_SM4.xlsx` | `So_sanh_chi_tiet` | Bảng C1.1 TP2 chi tiết trước/sau |
| | `Thong_ke_tong_hop` | Tổng hợp C1.1 TP2 |
| | `Thong_ke_theo_don_vi` | C1.1 TP2 theo đơn vị |
| `downloads/kq_sau_giam_tru/So_sanh_C11_SM2.xlsx` | (3 sheets tương tự) | C1.1 TP1 |
| `downloads/kq_sau_giam_tru/So_sanh_C12_SM1.xlsx` | (3 sheets tương tự) | C1.2 TP1 |
| `downloads/kq_sau_giam_tru/SM4-C12-ti-le-su-co-dv-brcd.xlsx` | (3 sheets tương tự) | C1.2 TP2 |
| `downloads/kq_sau_giam_tru/So_sanh_C14.xlsx` | (3 sheets tương tự) | C1.4 |
| `downloads/kq_sau_giam_tru/So_sanh_C15.xlsx` | (3 sheets tương tự) | C1.5 |
| `downloads/kq_sau_giam_tru/Tong_hop_giam_tru.xlsx` | Sheet 1 | Bảng tổng hợp so sánh, Biểu đồ GT |
| `downloads/kq_sau_giam_tru/Tong_hop_Diem_BSC_Don_Vi.xlsx` | `Tong_hop_Don_vi` | Bảng 1.2 (điểm BSC theo đơn vị), Biểu đồ BSC sau GT |
| | `Chi_tiet_Ca_nhan` | Bảng KPI cá nhân sau GT |

---

### 📂 NGUỒN THAM CHIẾU

| File | Mục Đích |
|------|---------|
| `du_lieu_tham_chieu/ds_phieu_loai_tru.xlsx` | Danh sách BAOHONG_ID cần loại trừ |
| `du_lieu_tham_chieu/Tonghop_thuebao_NVKT_DB_C12.xlsx` | Tổng thuê bao cho C1.2 |
| `dsnv.xlsx` | Danh sách nhân viên NVKT |
| `danhba.db` | Database danh bạ (25MB) |

---

## 🔄 QUY TRÌNH TẠO BÁO CÁO

```
1. Download dữ liệu từ baocao.hanoi.vnpt.vn
   └─> downloads/baocao_hanoi/*.xlsx

2. Xử lý giảm trừ (exclusion_process.py)
   Input: downloads/baocao_hanoi/*.xlsx + ds_phieu_loai_tru.xlsx
   └─> downloads/kq_sau_giam_tru/*.xlsx

3. Tính điểm KPI (kpi_calculator.py)
   Input: downloads/baocao_hanoi/*.xlsx
   └─> downloads/KPI/KPI_NVKT_*.xlsx

4. Tạo báo cáo Word (report_generator.py)
   Input:
      - downloads/KPI/*.xlsx
      - downloads/baocao_hanoi/*.xlsx
      - downloads/kq_sau_giam_tru/*.xlsx
   └─> downloads/reports/Bao_cao_KPI_NVKT_01_2026.docx
```

---

## CÔNG THỨC TÍNH ĐIỂM BSC

### C1.1 - Chất Lượng Sửa Chữa BRCĐ

| Thành Phần | Trọng Số | Công Thức Tính Điểm |
|-----------|---------|---------------------|
| **TP1 (SM2 - Sửa chữa chủ động)** | 30% | • ≥99% = 5 điểm<br>• 96-99% = 1 + 4×(KQ-96%)/3%<br>• <96% = 1 điểm |
| **TP2 (SM4 - Sửa chữa báo hỏng)** | 70% | • ≥85% = 5 điểm<br>• 82-85% = 4 + (KQ-82%)/3%<br>• <79% = 1 điểm |

**Điểm C1.1 = (TP1 × 0.30) + (TP2 × 0.70)**

---

### C1.2 - Tỷ Lệ Thuê Bao Báo Hỏng

| Thành Phần | Trọng Số | Công Thức Tính Điểm |
|-----------|---------|---------------------|
| **TP1 (SM1 - Hỏng lặp lại)** | 50% | • ≤2.5% = 5 điểm<br>• 2.5-4% = 5 - 4×(KQ-2.5%)/1.5%<br>• ≥4% = 1 điểm |
| **TP2 (SM4 - Tỷ lệ sự cố BRCĐ)** | 50% | • ≤2‰ = 5 điểm<br>• 2-3‰ = 5 - 4×(KQ-2‰)/1‰<br>• ≥3‰ = 1 điểm |

**Điểm C1.2 = (TP1 × 0.50) + (TP2 × 0.50)**

---

### C1.4 - Độ Hài Lòng Khách Hàng

| Công Thức |
|----------|
| • ≥99.5% = 5 điểm |
| • 95-99.5% = 1 + 4×(KQ-95%)/4.5% |
| • <95% = 1 điểm |

---

### C1.5 - Thiết Lập Dịch Vụ BRCĐ Đạt Thời Gian

| Công Thức |
|----------|
| • ≥99.5% = 5 điểm |
| • 89.5-99.5% = 1 + 4×(KQ-89.5%)/10% |
| • <89.5% = 1 điểm |

---

### Điểm BSC Tổng Hợp

**Công thức:**

```
BSC = (C1.1 × 40%) + (C1.2 × 40%) + (C1.4 × 10%) + (C1.5 × 10%)
```

**Phân loại:**

- **Excellent (Xuất sắc):** ≥4.5 điểm
- **Good (Tốt):** 3.5 - 4.5 điểm
- **Average (Trung bình):** <3.5 điểm

---

## CÁC FILE PYTHON LIÊN QUAN

### 1. baocaohanoi.py

**Chức năng:** Entry point chính - điều phối toàn bộ quy trình

**Các bước:**
1. Đăng nhập vào `baocao.hanoi.vnpt.vn` với OTP
2. Download báo cáo C1.1, C1.2, C1.4, C1.5, I1.5
3. Gọi `exclusion_process.py` để xử lý giảm trừ
4. Gọi `kpi_calculator.py` để tính điểm
5. Gọi `report_generator.py` để tạo báo cáo Word

---

### 2. exclusion_process.py

**Chức năng:** Xử lý loại trừ phiếu báo hỏng

**Input:**
- Danh sách loại trừ: `du_lieu_tham_chieu/ds_phieu_loai_tru.xlsx`
- Báo cáo gốc: `downloads/baocao_hanoi/*.xlsx`

**Output:**
- `downloads/kq_sau_giam_tru/So_sanh_C11_SM4.xlsx`
- `downloads/kq_sau_giam_tru/So_sanh_C11_SM2.xlsx`
- `downloads/kq_sau_giam_tru/So_sanh_C12_SM1.xlsx`
- `downloads/kq_sau_giam_tru/SM4-C12-ti-le-su-co-dv-brcd.xlsx`
- `downloads/kq_sau_giam_tru/So_sanh_C14.xlsx`
- `downloads/kq_sau_giam_tru/So_sanh_C15.xlsx`
- `downloads/kq_sau_giam_tru/Tong_hop_giam_tru.xlsx`
- `downloads/kq_sau_giam_tru/Tong_hop_Diem_BSC_Don_Vi.xlsx`

**Mỗi file có 3 sheets:**
- `So_sanh_chi_tiet` - So sánh chi tiết từng NVKT
- `Thong_ke_tong_hop` - Tổng hợp toàn TTVT
- `Thong_ke_theo_don_vi` - Thống kê theo từng tổ

---

### 3. kpi_calculator.py

**Chức năng:** Tính điểm BSC/KPI cho từng NVKT và đơn vị

**Input:**
- `downloads/baocao_hanoi/*.xlsx` (các báo cáo C1.x)

**Output:**
- `downloads/KPI/KPI_NVKT_TomTat.xlsx` - Tổng hợp điểm theo NVKT
- `downloads/KPI/KPI_NVKT_ChiTiet.xlsx` - Chi tiết thành phần TP1, TP2

**Áp dụng công thức:**
- C1.1: TP1 (30%) + TP2 (70%)
- C1.2: TP1 (50%) + TP2 (50%)
- C1.4: 100%
- C1.5: 100%
- BSC: C1.1 (40%) + C1.2 (40%) + C1.4 (10%) + C1.5 (10%)

---

### 4. report_generator.py

**Chức năng:** Tạo báo cáo Word với bảng biểu và biểu đồ

**Input:**
- `downloads/KPI/*.xlsx`
- `downloads/baocao_hanoi/*.xlsx`
- `downloads/kq_sau_giam_tru/*.xlsx`

**Output:**
- `downloads/reports/Bao_cao_KPI_NVKT_01_2026.docx`
- `downloads/reports/individual_reports/[Team]/Bao_cao_KPI_[NVKT_Name]_01_2026.docx`

**Thư viện sử dụng:**
- `python-docx` - Tạo file Word
- `matplotlib` - Tạo biểu đồ
- `pandas` - Xử lý dữ liệu
- `openpyxl` - Đọc Excel

---

## CẤU TRÚC THƯ MỤC

```
/home/vtst/baocaohanoi/
│
├── baocaohanoi.py                 # Entry point chính
├── exclusion_process.py           # Xử lý giảm trừ
├── kpi_calculator.py              # Tính điểm KPI
├── report_generator.py            # Tạo báo cáo Word
├── team_config.py                 # Cấu hình 4 tổ
├── config.py                      # Cấu hình hệ thống
│
├── du_lieu_tham_chieu/
│   ├── ds_phieu_loai_tru.xlsx     # Danh sách loại trừ
│   ├── Tonghop_thuebao_NVKT_DB_C12.xlsx
│   └── LOAI_TRU_C1.1_TP1.xlsx
│
├── downloads/
│   ├── baocao_hanoi/              # Báo cáo gốc (trước GT)
│   │   ├── c1.1 report.xlsx
│   │   ├── c1.2 report.xlsx
│   │   ├── c1.3 report.xlsx
│   │   ├── c1.4 report.xlsx
│   │   └── c1.5_chitiet_report.xlsx
│   │
│   ├── kq_sau_giam_tru/           # Kết quả sau giảm trừ
│   │   ├── So_sanh_C11_SM4.xlsx
│   │   ├── So_sanh_C11_SM2.xlsx
│   │   ├── So_sanh_C12_SM1.xlsx
│   │   ├── SM4-C12-ti-le-su-co-dv-brcd.xlsx
│   │   ├── So_sanh_C14.xlsx
│   │   ├── So_sanh_C15.xlsx
│   │   ├── Tong_hop_giam_tru.xlsx
│   │   └── Tong_hop_Diem_BSC_Don_Vi.xlsx
│   │
│   ├── KPI/                       # Điểm KPI
│   │   ├── KPI_NVKT_TomTat.xlsx
│   │   └── KPI_NVKT_ChiTiet.xlsx
│   │
│   └── reports/                   # Báo cáo Word
│       ├── Bao_cao_KPI_NVKT_01_2026.docx
│       └── individual_reports/
│           ├── Phuc_Tho/
│           ├── Quang_Oai/
│           ├── Suoi_Hai/
│           └── Son_Tay/
│
├── danhba.db                      # Database danh bạ
├── baocao_hanoi.db                # Database báo cáo
└── dsnv.xlsx                      # Danh sách nhân viên
```

---

## GHI CHÚ QUAN TRỌNG

1. **C1.3 không có giảm trừ** - Chỉ tiêu TSL (Leased Line) không áp dụng loại trừ phiếu

2. **4 Tổ BRCD:**
   - Phúc Thọ (9 NVKT)
   - Quảng Oai (10 NVKT)
   - Suối Hai (8 NVKT)
   - Sơn Tây (10 NVKT)

3. **Sheet quan trọng trong mỗi file So_sanh_*.xlsx:**
   - `So_sanh_chi_tiet` - Chi tiết từng NVKT với cả trước & sau
   - `Thong_ke_tong_hop` - Tổng hợp toàn TTVT Sơn Tây
   - `Thong_ke_theo_don_vi` - Thống kê theo 4 tổ

4. **File Tong_hop_Diem_BSC_Don_Vi.xlsx** là file quan trọng nhất chứa:
   - Sheet `Tong_hop_Don_vi` - Điểm BSC của 4 tổ (trước & sau)
   - Sheet `Chi_tiet_Ca_nhan` - Điểm BSC của từng NVKT (trước & sau)

5. **Trọng số BSC:**
   - C1.1: 40%
   - C1.2: 40%
   - C1.4: 10%
   - C1.5: 10%

---

**Ngày tạo tài liệu:** 2026-01-26
**Phiên bản:** 1.0
**Tác giả:** Phân tích tự động từ codebase
