# Hệ thống quản lý báo cáo Hà Nội

Hệ thống quản lý và truy vấn dữ liệu báo cáo từ các file Excel vào database SQLite.

## Tổng quan

Hệ thống này giúp:
- Tổ chức dữ liệu báo cáo từ nhiều file Excel rải rác vào một database tập trung
- Theo dõi dữ liệu theo ngày, tuần, tháng
- Trích xuất và tạo báo cáo theo nhiều tiêu chí khác nhau
- Phân tích xu hướng và biến động

## Cấu trúc Database

Database gồm các bảng chính:

### 1. Bảng dữ liệu hàng ngày
- **hoan_cong**: Dữ liệu thiết bị hoàn công (FIBER và MyTV)
- **ngung_psc**: Dữ liệu thiết bị ngừng PSC
- **thuc_tang**: Tổng hợp thực tăng theo tổ/NVKT

### 2. Bảng chất lượng dịch vụ
- **suy_hao_cao**: Dữ liệu suy hao cao (I1.5)
- **bao_cao_c1**: Báo cáo chỉ tiêu C1.x
- **bao_cao_kr**: Báo cáo KR6/KR7

### 3. Bảng tổng hợp
- **bao_cao_tuan_thang**: Báo cáo theo tuần/tháng
- **xu_huong_theo_ngay**: Xu hướng theo ngày
- **bien_dong_suy_hao**: Biến động suy hao cao

### 4. Bảng khác
- **nhan_vien**: Danh sách nhân viên
- **don_vi**: Danh sách đơn vị
- **import_log**: Log quá trình import

## Cài đặt

### Yêu cầu
```bash
pip install pandas openpyxl sqlite3
```

### Tạo database và import dữ liệu
```bash
python3 import_data.py
```

Script sẽ:
1. Tạo database `baocao_hanoi.db`
2. Tạo các bảng theo schema
3. Tự động import tất cả file Excel trong thư mục `downloads/baocao_hanoi/`
4. Ghi log quá trình import

## Sử dụng

### 1. Import dữ liệu

#### Import tất cả file
```bash
python3 import_data.py
```

#### Các loại file được hỗ trợ
- `hoan_cong_*.xlsx`: Báo cáo hoàn công
- `mytv_hoan_cong_*.xlsx`: Báo cáo hoàn công MyTV
- `ngung_psc_*.xlsx`: Báo cáo ngừng PSC
- `mytv_ngung_psc_*.xlsx`: Báo cáo ngừng PSC MyTV
- `thuc_tang_*.xlsx`: Báo cáo thực tăng
- `I1.5 report*.xlsx`: Báo cáo suy hao cao

### 2. Truy vấn và tạo báo cáo

#### Báo cáo theo ngày
```bash
python3 bao_cao_query.py --loai ngay --ngay 2025-11-20 --loai-dv FIBER
```

#### Báo cáo theo tuần
```bash
python3 bao_cao_query.py --loai tuan --tuan 47 --nam 2025 --loai-dv FIBER
```

#### Báo cáo theo tháng
```bash
python3 bao_cao_query.py --loai thang --thang 11 --nam 2025 --loai-dv FIBER
```

#### Báo cáo xu hướng
```bash
python3 bao_cao_query.py --loai xu-huong --tu-ngay 2025-11-01 --den-ngay 2025-11-20 --loai-dv FIBER
```

#### Báo cáo suy hao cao
```bash
python3 bao_cao_query.py --loai suy-hao --ngay 2025-11-20
```

#### Top NVKT
```bash
# Top 10 NVKT có hoàn công nhiều nhất
python3 bao_cao_query.py --loai top-nvkt --tu-ngay 2025-11-01 --den-ngay 2025-11-20 --loai-dv FIBER --top 10 --sap-xep hoan_cong

# Top 10 NVKT có thực tăng cao nhất
python3 bao_cao_query.py --loai top-nvkt --tu-ngay 2025-11-01 --den-ngay 2025-11-20 --loai-dv FIBER --top 10 --sap-xep thuc_tang
```

#### Thống kê tổng quan
```bash
python3 bao_cao_query.py --loai thong-ke
```

#### Export ra Excel
Thêm tham số `--export` để xuất kết quả ra file Excel:
```bash
python3 bao_cao_query.py --loai ngay --ngay 2025-11-20 --export bao_cao_ngay_20112025.xlsx
```

### 3. Truy vấn SQL trực tiếp

#### Kết nối database
```bash
sqlite3 baocao_hanoi.db
```

#### Ví dụ truy vấn

##### Tổng hợp hoàn công theo đơn vị
```sql
SELECT
    don_vi,
    COUNT(*) as so_luong,
    COUNT(DISTINCT nvkt) as so_nvkt
FROM hoan_cong
WHERE ngay_bao_cao = '2025-11-20'
GROUP BY don_vi
ORDER BY so_luong DESC;
```

##### Xu hướng hoàn công theo ngày
```sql
SELECT
    ngay_bao_cao,
    loai_dv,
    COUNT(*) as so_luong
FROM hoan_cong
WHERE ngay_bao_cao BETWEEN '2025-11-01' AND '2025-11-20'
GROUP BY ngay_bao_cao, loai_dv
ORDER BY ngay_bao_cao;
```

##### Top NVKT theo thực tăng
```sql
SELECT
    h.nvkt,
    h.don_vi,
    COUNT(h.id) as hoan_cong,
    COUNT(n.id) as ngung_psc,
    COUNT(h.id) - COUNT(n.id) as thuc_tang
FROM hoan_cong h
LEFT JOIN ngung_psc n ON h.ngay_bao_cao = n.ngay_bao_cao AND h.nvkt = n.nvkt
WHERE h.ngay_bao_cao BETWEEN '2025-11-01' AND '2025-11-20'
GROUP BY h.nvkt, h.don_vi
ORDER BY thuc_tang DESC
LIMIT 10;
```

##### Biến động suy hao cao theo NVKT
```sql
SELECT
    doi_one,
    nvkt_db_normalized,
    COUNT(*) as so_tb_suy_hao
FROM suy_hao_cao
WHERE ngay_bao_cao = '2025-11-20'
GROUP BY doi_one, nvkt_db_normalized
ORDER BY so_tb_suy_hao DESC;
```

## Cấu trúc thư mục

```
baocaohanoi/
├── database_schema.sql          # Schema database
├── import_data.py               # Script import dữ liệu
├── bao_cao_query.py             # Script truy vấn và tạo báo cáo
├── analyze_excel_structure.py   # Script phân tích cấu trúc Excel
├── baocao_hanoi.db             # Database SQLite (được tạo sau khi chạy import)
├── excel_structure_analysis.json # Kết quả phân tích cấu trúc Excel
├── README_DATABASE.md           # File này
└── downloads/
    └── baocao_hanoi/            # Thư mục chứa các file Excel
        ├── hoan_cong_*.xlsx
        ├── ngung_psc_*.xlsx
        ├── thuc_tang_*.xlsx
        ├── I1.5 report*.xlsx
        └── ...
```

## Views (Truy vấn nhanh)

Database có sẵn các view để truy vấn nhanh:

### v_tong_hop_ngay
Tổng hợp hoàn công, ngừng PSC, thực tăng theo ngày và đơn vị
```sql
SELECT * FROM v_tong_hop_ngay WHERE ngay_bao_cao = '2025-11-20';
```

### v_tong_hop_nvkt
Tổng hợp theo NVKT
```sql
SELECT * FROM v_tong_hop_nvkt WHERE ngay_bao_cao = '2025-11-20';
```

### v_xu_huong_suy_hao
Xu hướng suy hao cao theo đơn vị
```sql
SELECT * FROM v_xu_huong_suy_hao ORDER BY ngay_bao_cao DESC LIMIT 30;
```

## Lưu ý

1. **Format tên file**: Script tự động trích xuất ngày từ tên file theo format `ddmmyyyy`. Ví dụ: `hoan_cong_20112025.xlsx` → ngày báo cáo là 2025-11-20

2. **Import lại dữ liệu**: Khi import lại cùng một file, dữ liệu sẽ được cập nhật (không trùng lặp) nhờ constraint UNIQUE

3. **Log import**: Kiểm tra bảng `import_log` để xem lịch sử import:
```sql
SELECT * FROM import_log ORDER BY created_at DESC;
```

4. **Backup database**: Nên backup database định kỳ:
```bash
cp baocao_hanoi.db baocao_hanoi_backup_$(date +%Y%m%d).db
```

## Troubleshooting

### Lỗi "no such table"
Chạy lại script tạo schema:
```bash
sqlite3 baocao_hanoi.db < database_schema.sql
```

### Lỗi import file Excel
Kiểm tra:
- File có đúng format không?
- Sheet name có đúng không? (Data, thuc_tang_theo_to, etc.)
- Các cột có đúng tên không?

### Database bị lỗi
Xóa và tạo lại:
```bash
rm baocao_hanoi.db
python3 import_data.py
```

## Phát triển thêm

Để thêm loại báo cáo mới:
1. Thêm bảng trong `database_schema.sql`
2. Thêm hàm import trong `import_data.py`
3. Thêm hàm query trong `bao_cao_query.py`

## Liên hệ

Nếu có vấn đề hoặc đề xuất, vui lòng liên hệ.
