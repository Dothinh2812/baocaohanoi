# HƯỚNG DẪN SỬ DỤNG HỆ THỐNG TRACKING LỊCH SỬ SUY HAO

## Tổng quan

Hệ thống này giúp theo dõi lịch sử suy hao tín hiệu của các thuê bao theo thời gian, với các tính năng:

- **Tracking từng thuê bao**: Theo dõi từng mã thuê bao xuất hiện/biến mất mỗi ngày
- **Số ngày liên tục**: Tính số ngày một thuê bao bị suy hao liên tục
- **Báo cáo biến động**: So sánh TĂNG/GIẢM theo ngày/tuần/tháng
- **Danh sách chi tiết**: Xem cụ thể thuê bao nào tăng, thuê bao nào giảm
- **Lưu trữ vĩnh viễn**: Database SQLite lưu toàn bộ lịch sử

---

## Cấu trúc Files

```
/home/vtst/baocaohanoi/
├── suy_hao_history.db              # Database lịch sử (QUAN TRỌNG - backup file này!)
├── init_suy_hao_db.py              # Script khởi tạo database (chỉ chạy 1 lần)
├── import_initial_data.py          # Script import dữ liệu khởi đầu
├── c1_process.py                   # File chính (đã sửa đổi)
├── c1_process_enhanced.py          # Hàm xử lý I1.5 với tracking
├── suy_hao_reports.py              # Module tạo báo cáo tuần/tháng
└── downloads/baocao_hanoi/
    ├── I1.5 report.xlsx            # File input (có thêm sheet mới)
    ├── Bao_cao_tuan_XX_YYYY.xlsx  # Báo cáo tuần
    └── Bao_cao_thang_MM_YYYY.xlsx # Báo cáo tháng
```

---

## Cách sử dụng hàng ngày

### 1. Xử lý báo cáo I1.5 hàng ngày

Sau khi download file `I1.5 report.xlsx` từ hệ thống VNPT:

```bash
# Cách 1: Chạy trực tiếp
python3 c1_process.py

# Cách 2: Chỉ chạy process_I15_report
python3 c1_process_enhanced.py

# Cách 3: Import vào code khác
from c1_process import process_I15_report
process_I15_report()
```

**Kết quả:**
- File `I1.5 report.xlsx` được xử lý và thêm các sheet mới:
  - `Sheet1`: Dữ liệu gốc đầy đủ
  - `TH_SHC_I15`: Tổng hợp theo NVKT_DB (như cũ)
  - `TH_SHC_theo_to`: Tổng hợp theo tổ (như cũ)
  - `shc_theo_SA`: Thống kê theo SA (như cũ)
  - **`Bien_dong_tong_hop`** (MỚI): Tổng hợp số liệu tăng/giảm theo NVKT_DB
  - **`Tang_moi`** (MỚI): Danh sách thuê bao mới xuất hiện hôm nay
  - **`Giam_het`** (MỚI): Danh sách thuê bao hết suy hao hôm nay + số ngày đã suy hao
  - **`Van_con`** (MỚI): Danh sách thuê bao tiếp tục suy hao + số ngày liên tục
  - [Các sheet NVKT_DB]: Chi tiết từng nhân viên (như cũ)

- Dữ liệu được lưu vào database `suy_hao_history.db`

---

### 2. Tạo báo cáo tuần

```python
from suy_hao_reports import generate_weekly_report

# Tạo báo cáo tuần 44 năm 2025
generate_weekly_report(2025, 44)

# Hoặc chỉ định file output
generate_weekly_report(2025, 44, output_file="bao_cao_tuan_44.xlsx")
```

**Kết quả:** File Excel với các sheet:
- `Tong_hop`: Bảng so sánh số liệu tuần này vs tuần trước theo NVKT_DB
- `Tang_moi`: Danh sách thuê bao mới trong tuần này
- `Giam_het`: Danh sách thuê bao hết suy hao trong tuần này
- `Van_con`: Danh sách thuê bao tiếp tục từ tuần trước

---

### 3. Tạo báo cáo tháng

```python
from suy_hao_reports import generate_monthly_report

# Tạo báo cáo tháng 11 năm 2025
generate_monthly_report(2025, 11)

# Hoặc chỉ định file output
generate_monthly_report(2025, 11, output_file="bao_cao_thang_11.xlsx")
```

**Kết quả:** File Excel với các sheet:
- `Tong_hop`: So sánh tháng này vs tháng trước theo NVKT_DB
- `Xu_huong_theo_ngay`: Biểu đồ số lượng thuê bao suy hao từng ngày trong tháng
- `Tang_moi`: Danh sách thuê bao mới trong tháng này
- `Giam_het`: Danh sách thuê bao hết suy hao trong tháng này
- `Van_con`: Danh sách thuê bao tiếp tục từ tháng trước

---

### 4. Tạo báo cáo xu hướng tùy chỉnh

```python
from suy_hao_reports import generate_trend_report

# Báo cáo từ ngày 01/11 đến 10/11
generate_trend_report('2025-11-01', '2025-11-10')
```

**Kết quả:** File Excel với:
- `Xu_huong_theo_don_vi`: Bảng pivot số liệu theo ngày và đơn vị
- `Xu_huong_theo_NVKT`: Xu hướng theo từng NVKT_DB

---

## Hiểu về dữ liệu biến động

### Logic phân loại

**TĂNG MỚI:**
- Thuê bao xuất hiện trong báo cáo hôm nay
- KHÔNG có trong báo cáo hôm qua
- Nghĩa: Thuê bao mới bắt đầu bị suy hao

**GIẢM/HẾT:**
- Thuê bao có trong báo cáo hôm qua
- KHÔNG có trong báo cáo hôm nay
- Nghĩa: Tín hiệu đã được cải thiện, không còn suy hao

**VẪN CÒN:**
- Thuê bao có trong cả 2 báo cáo (hôm qua VÀ hôm nay)
- Có cột "Số ngày liên tục" = số ngày thuê bao này bị suy hao liên tiếp
- Nghĩa: Vấn đề chưa được giải quyết

### Ví dụ thực tế

Ngày 02/11: 296 thuê bao suy hao
Ngày 03/11: 302 thuê bao suy hao

Kết quả:
- **TĂNG MỚI**: 33 thuê bao (mới xuất hiện)
- **GIẢM/HẾT**: 27 thuê bao (đã được sửa)
- **VẪN CÒN**: 269 thuê bao (chưa sửa xong)
- **Biến động ròng**: +6 thuê bao (302 - 296)

---

## Cấu trúc Database

### Bảng `suy_hao_snapshots`
Lưu snapshot đầy đủ mỗi ngày (598 bản ghi = 296 + 302)

### Bảng `suy_hao_tracking`
Theo dõi trạng thái từng thuê bao (329 thuê bao unique)
- `ngay_xuat_hien_dau_tien`: Ngày đầu tiên phát hiện suy hao
- `ngay_thay_cuoi_cung`: Ngày cuối cùng thấy thuê bao này
- `so_ngay_lien_tuc`: Số ngày suy hao liên tiếp
- `trang_thai`: DANG_SUY_HAO hoặc DA_HET_SUY_HAO

### Bảng `suy_hao_daily_changes`
Chi tiết biến động từng ngày (329 bản ghi)
- Mỗi bản ghi = 1 thuê bao có biến động
- `loai_bien_dong`: TANG_MOI, GIAM_HET, VAN_CON

### Bảng `suy_hao_daily_summary`
Tổng hợp nhanh theo NVKT_DB (33 bản ghi)

---

## Backup và Restore

### Backup database

```bash
# Cách 1: Copy file
cp suy_hao_history.db suy_hao_history_backup_$(date +%Y%m%d).db

# Cách 2: Export SQL
sqlite3 suy_hao_history.db .dump > backup.sql
```

### Restore database

```bash
# Từ file backup
cp suy_hao_history_backup_20251103.db suy_hao_history.db

# Từ SQL dump
sqlite3 suy_hao_history_new.db < backup.sql
```

---

## Query dữ liệu trực tiếp

### Xem thuê bao suy hao lâu nhất

```python
import sqlite3
import pandas as pd

conn = sqlite3.connect('suy_hao_history.db')

# Top 10 thuê bao suy hao nhiều ngày nhất
df = pd.read_sql_query("""
    SELECT account_cts, doi_one, nvkt_db, so_ngay_lien_tuc,
           ngay_xuat_hien_dau_tien, ngay_thay_cuoi_cung
    FROM suy_hao_tracking
    WHERE trang_thai = 'DANG_SUY_HAO'
    ORDER BY so_ngay_lien_tuc DESC
    LIMIT 10
""", conn)

print(df)
conn.close()
```

### Xem xu hướng theo tổ

```python
conn = sqlite3.connect('suy_hao_history.db')

df = pd.read_sql_query("""
    SELECT ngay_bao_cao, doi_one, COUNT(*) as so_luong
    FROM suy_hao_snapshots
    GROUP BY ngay_bao_cao, doi_one
    ORDER BY ngay_bao_cao, doi_one
""", conn)

# Pivot để hiển thị dạng bảng
pivot = df.pivot(index='ngay_bao_cao', columns='doi_one', values='so_luong')
print(pivot)

conn.close()
```

### Thống kê theo NVKT_DB

```python
conn = sqlite3.connect('suy_hao_history.db')

df = pd.read_sql_query("""
    SELECT nvkt_db,
           COUNT(*) as tong_so_tb,
           AVG(so_ngay_lien_tuc) as trung_binh_ngay,
           MAX(so_ngay_lien_tuc) as max_ngay
    FROM suy_hao_tracking
    WHERE trang_thai = 'DANG_SUY_HAO'
    GROUP BY nvkt_db
    ORDER BY tong_so_tb DESC
""", conn)

print(df)
conn.close()
```

---

## Troubleshooting

### Lỗi: "Không tìm thấy database"

**Nguyên nhân**: File `suy_hao_history.db` không tồn tại

**Giải pháp**:
```bash
python3 init_suy_hao_db.py
```

### Lỗi: "Không có dữ liệu ngày hôm qua"

**Nguyên nhân**: Chưa có dữ liệu trong database

**Giải pháp**:
- Lần đầu chạy sẽ không có so sánh
- Từ ngày thứ 2 trở đi sẽ có đầy đủ thông tin biến động

### Sheet TĂNG/GIẢM/VẪN CÒN trống

**Nguyên nhân**: Lần đầu tiên xử lý hoặc chưa có dữ liệu ngày hôm qua

**Giải pháp**: Bình thường, hãy chạy lại vào ngày hôm sau

---

## Lưu ý quan trọng

1. **Backup database thường xuyên**: File `suy_hao_history.db` rất quan trọng!

2. **Không xóa database**: Mọi thống kê đều dựa trên lịch sử trong database

3. **Chạy hàng ngày**: Để có dữ liệu liên tục, nên xử lý báo cáo I1.5 mỗi ngày

4. **Thứ tự xử lý**:
   - Sáng: Download I1.5 report mới
   - Chạy: `python3 c1_process.py`
   - Xem: Các sheet TĂNG/GIẢM/VẪN CÒN
   - Cuối tuần/tháng: Chạy báo cáo tuần/tháng

5. **File I1.5 report.xlsx được ghi đè**: File gốc sẽ được thêm sheet mới, backup nếu cần

---

## Liên hệ / Hỗ trợ

Nếu có vấn đề, kiểm tra:
1. File log khi chạy script
2. Database có dữ liệu không: `ls -lh suy_hao_history.db`
3. Số bản ghi: Chạy script kiểm tra database ở trên

---

**Phiên bản**: 1.0
**Ngày tạo**: 04/11/2025
**Người tạo**: Claude Code
