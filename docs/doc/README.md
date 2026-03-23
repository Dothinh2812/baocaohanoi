# Baocao Hanoi - Automation Script

Script tự động tải báo cáo từ https://baocao.hanoi.vnpt.vn/

## Tính năng

- Tự động đăng nhập với OTP
- Tự động tải các báo cáo PTTB và vật tư thu hồi
- Tự động sử dụng ngày hiện tại cho báo cáo
- Lưu file vào thư mục `downloads/baocao_hanoi/`

## Cấu trúc URL báo cáo

### URL Parameters

Các tham số trong URL báo cáo:

```
https://baocao.hanoi.vnpt.vn/report/report-info-data?
  id=264354                    # ID báo cáo
  &vdvvt_id=9                  # ID vùng/đơn vị
  &vdenngay=27%2F10%2F2025     # Ngày báo cáo (dd/mm/yyyy - encoded)
  &vdonvi_id=14324             # ID đơn vị
  &vloai=1                     # Loại báo cáo
```

### Tham số ngày (`vdenngay`)

- **Format**: `dd/mm/yyyy`
- **URL Encoding**: Dấu `/` được encode thành `%2F`
- **Ví dụ**:
  - Ngày: `27/10/2025`
  - Encoded: `27%2F10%2F2025`

## Cài đặt

### 1. Cài đặt thư viện Python

```bash
pip install -r requirements.txt
playwright install chromium
```

Hoặc cài đặt từng thư viện:

```bash
pip install playwright pandas openpyxl
playwright install chromium
```

### 2. Chuẩn bị file dsnv.xlsx

Đặt file `dsnv.xlsx` (danh sách nhân viên) cùng thư mục với script. File cần có các cột:
- `Họ tên` - Tên đầy đủ nhân viên
- `Đơn vị` - Đơn vị công tác

## Cách sử dụng

### 1. Chạy script tự động

```bash
python baocaohanoi.py
```

Script sẽ:
1. Đăng nhập tự động (đọc OTP từ file hoặc chờ nhập thủ công)
2. Tải báo cáo PTTB Ngưng PSC (ID: 264354)
3. Tải báo cáo PTTB Hoàn công (ID: 260054)
4. Giữ trình duyệt mở 10 giây để kiểm tra
5. Xử lý và chuẩn hóa báo cáo Ngưng PSC (thêm cột NVKT, Đơn vị và 2 sheet thống kê)
6. Xử lý và chuẩn hóa báo cáo Hoàn công (thêm cột NVKT, Đơn vị và 2 sheet thống kê)
7. Tạo báo cáo Thực tăng (kết hợp 2 báo cáo trên)

### 2. File OTP

Script đọc OTP từ: `G:\My Drive\App- baocao\OTP-handle\otp_logs.txt`

Yêu cầu:
- File chứa mã OTP 6 chữ số
- File phải được tạo trong vòng 120 giây gần đây

**Tự động xóa OTP sau khi sử dụng:**
- Script sẽ tự động xóa nội dung file OTP sau khi điền thành công
- Tránh lỗi sử dụng lại OTP cũ trong lần chạy sau
- Đảm bảo mỗi OTP chỉ được sử dụng một lần

### 3. Thư mục tải về

Tất cả báo cáo sẽ được lưu vào:
```
downloads/baocao_hanoi/
```

## Cấu trúc code

### Các hàm chính

1. **`read_otp_from_file()`** - Đọc mã OTP từ file (trả về tuple: otp_code, file_path)
2. **`clear_otp_file(file_path)`** - Xóa nội dung file OTP sau khi sử dụng
3. **`login_baocao_hanoi()`** - Đăng nhập vào hệ thống
4. **`download_report_pttb_ngung_psc(page)`** - Tải báo cáo PTTB Ngưng PSC
5. **`download_report_pttb_hoan_cong(page)`** - Tải báo cáo PTTB Hoàn công
6. **`download_report_vattu_thuhoi(page)`** - Tải báo cáo vật tư thu hồi
7. **`process_ngung_psc_report()`** - Xử lý và chuẩn hóa báo cáo Ngưng PSC
8. **`process_hoan_cong_report()`** - Xử lý và chuẩn hóa báo cáo Hoàn công
9. **`create_thuc_tang_report()`** - Tạo báo cáo Thực tăng từ 2 báo cáo trên
10. **`main()`** - Hàm chính điều khiển workflow

### Cập nhật ngày tự động

Code tự động sử dụng ngày hiện tại:

```python
from datetime import datetime
from urllib.parse import quote

# Lấy ngày hiện tại
current_date = datetime.now().strftime("%d/%m/%Y")  # Ví dụ: 27/10/2025

# Encode cho URL
encoded_date = quote(current_date, safe='')  # Kết quả: 27%2F10%2F2025

# Tạo URL
report_url = f"https://baocao.hanoi.vnpt.vn/report/report-info-data?id=264354&vdvvt_id=9&vdenngay={encoded_date}&vdonvi_id=14324&vloai=1"
```

## Thông tin đăng nhập

- **Username**: your_username
- **Password**: your_password
- **OTP**: Đọc từ file hoặc nhập thủ công

## Xử lý báo cáo tự động

### 1. Xử lý báo cáo Ngưng PSC

Hàm `process_ngung_psc_report()` tự động xử lý báo cáo sau khi tải về:

### Các bước xử lý:

1. **Đọc file** `ngung_psc_DDMMYYYY.xlsx` và `dsnv.xlsx`
2. **Chuẩn hóa tên nhân viên** từ cột "Nhóm địa bàn":
   - `Đồng Mô 4 - Đỗ Minh Thăng` → `Đỗ Minh Thăng`
   - `VNM3-Khuất Anh Chiến( VXN)` → `Khuất Anh Chiến`
   - Loại bỏ phần trước dấu `-` và phần trong ngoặc đơn `()`
3. **Ghi kết quả** vào cột `NVKT`
4. **Tra cứu đơn vị** từ file `dsnv.xlsx` (cột "Họ tên" khớp với `NVKT`)
5. **Thêm cột** `Đơn vị` vào báo cáo
6. **Tạo 2 sheet thống kê**:
   - Sheet `ngung-psc-theo-to`: Thống kê số lượng TB theo Tổ (Đơn vị)
   - Sheet `ngung-psc-theo-NVKT`: Thống kê số lượng TB theo NVKT và Tổ
7. **Lưu file** với 3 sheet (Data + 2 sheet thống kê)

### Ví dụ chuyển đổi:

| Nhóm địa bàn (gốc) | NVKT (chuẩn hóa) | Đơn vị (tra cứu) |
|---|---|---|
| Đồng Mô 4 - Đỗ Minh Thăng | Đỗ Minh Thăng | TTVT Sơn Tây |
| VNM3-Khuất Anh Chiến( VXN) | Khuất Anh Chiến | TTVT Hà Đông |
| Lê Văn A | Lê Văn A | TTVT Ba Đình |

### Cấu trúc file kết quả:

File Excel sau khi xử lý sẽ có 3 sheet:

#### Sheet 1: Data (Dữ liệu gốc)
Dữ liệu đầy đủ với 2 cột bổ sung:
- `NVKT` - Tên nhân viên kỹ thuật đã chuẩn hóa
- `Đơn vị` - Tổ/Đơn vị công tác

#### Sheet 2: ngung-psc-theo-to
Thống kê số lượng thuê bao theo Tổ:

| Đơn vị | Số lượng TB |
|--------|-------------|
| TTVT Sơn Tây | 45 |
| TTVT Hà Đông | 38 |
| TTVT Ba Đình | 25 |
| ... | ... |
| TỔNG CỘNG | 150 |

#### Sheet 3: ngung-psc-theo-NVKT
Thống kê số lượng thuê bao theo NVKT:

| Đơn vị | NVKT | Số lượng TB |
|--------|------|-------------|
| TTVT Sơn Tây | Đỗ Minh Thăng | 12 |
| TTVT Sơn Tây | Nguyễn Văn A | 8 |
| TTVT Hà Đông | Khuất Anh Chiến | 15 |
| ... | ... | ... |
| TỔNG CỘNG | | 150 |

### 2. Xử lý báo cáo Hoàn công

Hàm `process_hoan_cong_report()` tự động xử lý báo cáo sau khi tải về:

#### Các bước xử lý:

1. **Đọc file** `hoan_cong_DDMMYYYY.xlsx` và `dsnv.xlsx`
2. **Chuẩn hóa tên nhân viên** từ cột "Nhân viên KT":
   - `VNPT016763-Nguyễn Quảng Ba` → `Nguyễn Quảng Ba`
   - Loại bỏ phần trước dấu `-` và phần trong ngoặc đơn `()`
3. **Ghi kết quả** vào cột `NVKT`
4. **Tra cứu đơn vị** từ file `dsnv.xlsx` (cột "Họ tên" khớp với `NVKT`)
5. **Thêm cột** `Đơn vị` vào báo cáo
6. **Tạo 2 sheet thống kê**:
   - Sheet `hoan-cong-theo-to`: Thống kê số lượng TB theo Tổ (Đơn vị)
   - Sheet `hoan-cong-theo-NVKT`: Thống kê số lượng TB theo NVKT và Tổ
7. **Lưu file** với 3 sheet (Data + 2 sheet thống kê)

#### Cấu trúc file kết quả:

File Excel sau khi xử lý sẽ có 3 sheet:

**Sheet 1: Data** - Dữ liệu gốc với cột NVKT và Đơn vị

**Sheet 2: hoan-cong-theo-to** - Thống kê theo Tổ

**Sheet 3: hoan-cong-theo-NVKT** - Thống kê theo NVKT

### 3. Tạo báo cáo Thực tăng

Hàm `create_thuc_tang_report()` tự động tạo báo cáo sau khi xử lý 2 báo cáo trên:

#### Công thức tính:

```
Thực tăng = Hoàn công - Ngưng PSC
```

#### Cấu trúc file `thuc_tang_DDMMYYYY.xlsx`:

File Excel có 2 sheet:

**Sheet 1: thuc_tang_theo_to** - Thống kê theo Tổ

| Đơn vị | Hoàn công | Ngưng PSC | Thực tăng | Tỷ lệ (%) |
|--------|-----------|-----------|-----------|----------|
| TTVT Ba Đình | 20 | 15 | 5 | 33.33 |
| TTVT Thanh Xuân | 12 | 10 | 2 | 20.00 |
| TTVT Hà Đông | 35 | 38 | -3 | -7.89 |
| TTVT Sơn Tây | 28 | 45 | -17 | -37.78 |
| TỔNG CỘNG | 120 | 150 | -30 | -20.00 |

**Sheet 2: thuc_tang_theo_NVKT** - Thống kê theo NVKT

| Đơn vị | NVKT | Hoàn công | Ngưng PSC | Thực tăng | Tỷ lệ (%) |
|--------|------|-----------|-----------|-----------|----------|
| TTVT Ba Đình | Lê Thị D | 9 | 3 | 6 | 200.00 |
| TTVT Hoàn Kiếm | Hoàng Thị F | 7 | 3 | 4 | 133.33 |
| TTVT Hà Đông | Nguyễn Quảng Ba | 12 | 12 | 0 | 0.00 |
| TTVT Sơn Tây | Đỗ Minh Thăng | 10 | 12 | -2 | -16.67 |
| TỔNG CỘNG | | 120 | 150 | -30 | -20.00 |

#### Tính năng:

- **Merge dữ liệu**: Kết hợp 2 báo cáo theo Đơn vị và NVKT
- **Tính toán tự động**: Thực tăng = Hoàn công - Ngưng PSC
- **Sắp xếp**: Theo Thực tăng giảm dần (cao nhất lên trước)
- **Top 5**: Hiển thị Top 5 Tổ và NVKT có Thực tăng cao nhất

### File cần thiết:

- `dsnv.xlsx` - Danh sách nhân viên với các cột:
  - `Họ tên` - Tên đầy đủ nhân viên
  - `Đơn vị` - Đơn vị công tác

### Xử lý đặc biệt:

**Tra cứu thông minh và chuẩn hóa tên NVKT**

Script sử dụng 2 cấp độ tra cứu để xử lý trường hợp viết hoa/thường không khớp:

1. **Exact match**: Thử khớp chính xác tên trước
2. **Case-insensitive match**: Nếu không khớp, thử so sánh lowercase

**QUAN TRỌNG**: Khi tìm thấy qua lowercase matching, tên NVKT sẽ được thay thế bằng tên chuẩn từ file dsnv.xlsx

Ví dụ:
- Báo cáo 1: `VNPT016765-Bùi Văn Cường` → Chuẩn hóa: `Bùi Văn Cường`
- Báo cáo 2: `VNM3-Bùi văn Cường` → Chuẩn hóa: `Bùi văn Cường`
- File dsnv: `Bùi Văn Cường` (tên chuẩn)
- Kết quả: Cả 2 đều được thay thế thành `Bùi Văn Cường` ✅

**Lợi ích:**
- Tránh trùng lặp bản ghi do viết hoa/thường khác nhau
- Đảm bảo tính nhất quán trong thống kê
- Tên NVKT luôn theo chuẩn trong file dsnv.xlsx

## Lưu ý

- Script chạy với trình duyệt có giao diện (headless=False)
- Timeout mặc định: 60-500 giây tùy từng thao tác
- Nếu không có OTP, script sẽ chờ 10 giây để nhập thủ công
- Trình duyệt sẽ tự động đóng sau 10 giây khi tải xong báo cáo
- File `dsnv.xlsx` phải nằm cùng thư mục với script

## Các báo cáo được tải

### Báo cáo 1: PTTB Ngưng PSC (ID: 264354)
- Hàm: `download_report_pttb_ngung_psc(page)`
- Báo cáo chi tiết thuê bao Ngưng PSC tạm tính
- Tự động lấy ngày hiện tại
- Tên file: `ngung_psc_DDMMYYYY.xlsx` (VD: `ngung_psc_27102025.xlsx`)
- File sẽ được ghi đè nếu đã tồn tại

### Báo cáo 2: PTTB Hoàn công (ID: 260054)
- Hàm: `download_report_pttb_hoan_cong(page)`
- Báo cáo lũy kế tháng hoàn công
- Tự động lấy ngày hiện tại
- Tên file: `hoan_cong_DDMMYYYY.xlsx` (VD: `hoan_cong_27102025.xlsx`)
- File sẽ được ghi đè nếu đã tồn tại

### Báo cáo 3: Vật tư thu hồi (ID: 270922)
- Hàm: `download_report_vattu_thuhoi(page)`
- Đơn vị: TTVT Sơn Tây
- Ngày cố định: 24/09/2025
- Hiện đang bị comment trong hàm main()

## Xử lý lỗi

Script có xử lý lỗi và in thông tin chi tiết:
- Timeout khi tải trang
- Không tìm thấy element
- Lỗi khi tải file
- Không đọc được OTP

## Ví dụ output

```
=== Bắt đầu đăng nhập vào baocao.hanoi.vnpt.vn ===
Đang truy cập trang đăng nhập...
Đang điền username...
Đang điền password...
Đang click button Đăng nhập...
Đang đợi trường nhập OTP...
Đang đọc mã OTP từ file...
✅ Found OTP code in file: 123456
Đang điền OTP: 123456
Đang click button xác nhận OTP...
✅ Đã xóa nội dung file OTP để tránh sử dụng lại
✅ Đăng nhập thành công!

=== Bắt đầu tải báo cáo PTTB Ngưng PSC ===
Ngày báo cáo: 27/10/2025
Đang truy cập: https://baocao.hanoi.vnpt.vn/report/report-info-data?id=264354&vdvvt_id=9&vdenngay=27%2F10%2F2025&vdonvi_id=14324&vloai=1
Đang đợi dữ liệu load...
Đang tìm button 'Xuất Excel'...
Đã tìm thấy button 'Xuất Excel', đang click...
Đang tìm và click '2.Tất cả dữ liệu'...
Đang tải file...
✅ Đã tải file về: downloads/baocao_hanoi/ngung_psc_27102025.xlsx

=== Bắt đầu tải báo cáo PTTB Hoàn công ===
Ngày báo cáo: 27/10/2025
Đang truy cập: https://baocao.hanoi.vnpt.vn/report/report-info-data?id=260054&vdvvt_id=9&vdenngay=27%2F10%2F2025&vdonvi_id=14324&vloai=1&vloai_bc=luyke_thang_hoancong
Đang đợi dữ liệu load...
Đang tải file...
✅ Đã tải file về: downloads/baocao_hanoi/hoan_cong_27102025.xlsx

✅ Hoàn thành tải báo cáo!
Trình duyệt sẽ giữ mở trong 10 giây để bạn kiểm tra.

Đang đóng trình duyệt...

=== Bắt đầu xử lý báo cáo Ngưng PSC ===
Đang đọc file: downloads/baocao_hanoi/ngung_psc_27102025.xlsx
Đang đọc file: dsnv.xlsx
Đang chuẩn hóa tên nhân viên kỹ thuật...
✅ Đã chuẩn hóa 150 tên nhân viên
Đang tra cứu đơn vị từ danh sách nhân viên...
✅ Đã tra cứu được đơn vị cho 145/150 bản ghi

📊 Thống kê cơ bản:
   - Tổng số bản ghi: 150
   - Số bản ghi có đơn vị: 145
   - Số bản ghi chưa có đơn vị: 5

📊 Đang tạo thống kê theo Tổ...
✅ Đã tạo thống kê cho 8 tổ
📊 Đang tạo thống kê theo NVKT...
✅ Đã tạo thống kê cho 45 NVKT

💾 Đang lưu file với 3 sheet...
✅ Đã lưu file: downloads/baocao_hanoi/ngung_psc_27102025.xlsx
   - Sheet 'Data': Dữ liệu đầy đủ (150 dòng)
   - Sheet 'ngung-psc-theo-to': Thống kê theo Tổ (9 dòng)
   - Sheet 'ngung-psc-theo-NVKT': Thống kê theo NVKT (46 dòng)

📊 Top 5 Tổ có nhiều TB ngưng PSC nhất:
   1. TTVT Sơn Tây: 45 TB
   2. TTVT Hà Đông: 38 TB
   3. TTVT Ba Đình: 25 TB
   4. TTVT Hoàn Kiếm: 18 TB
   5. TTVT Thanh Xuân: 12 TB

📊 Top 5 NVKT có nhiều TB ngưng PSC nhất:
   1. Khuất Anh Chiến (TTVT Hà Đông): 15 TB
   2. Đỗ Minh Thăng (TTVT Sơn Tây): 12 TB
   3. Nguyễn Văn A (TTVT Ba Đình): 10 TB
   4. Trần Văn B (TTVT Sơn Tây): 9 TB
   5. Lê Thị C (TTVT Hoàn Kiếm): 8 TB

✅ Hoàn thành xử lý báo cáo Ngưng PSC!

=== Bắt đầu xử lý báo cáo Hoàn công ===
Đang đọc file: downloads/baocao_hanoi/hoan_cong_27102025.xlsx
Đang đọc file: dsnv.xlsx
✅ Tìm thấy cột: 'Nhân viên KT' trong hoan_cong
✅ Tìm thấy cột: 'Họ tên' và 'đơn vị' trong dsnv
Đang chuẩn hóa tên nhân viên kỹ thuật...
✅ Đã chuẩn hóa 120 tên nhân viên
Đang tra cứu đơn vị từ danh sách nhân viên...
✅ Đã tra cứu được đơn vị cho 115/120 bản ghi

📊 Thống kê cơ bản:
   - Tổng số bản ghi: 120
   - Số bản ghi có đơn vị: 115
   - Số bản ghi chưa có đơn vị: 5

📊 Đang tạo thống kê theo Tổ...
✅ Đã tạo thống kê cho 7 tổ
📊 Đang tạo thống kê theo NVKT...
✅ Đã tạo thống kê cho 38 NVKT

💾 Đang lưu file với 3 sheet...
✅ Đã lưu file: downloads/baocao_hanoi/hoan_cong_27102025.xlsx
   - Sheet 'Data': Dữ liệu đầy đủ (120 dòng)
   - Sheet 'hoan-cong-theo-to': Thống kê theo Tổ (8 dòng)
   - Sheet 'hoan-cong-theo-NVKT': Thống kê theo NVKT (39 dòng)

📊 Top 5 Tổ có nhiều TB hoàn công nhất:
   1. TTVT Hà Đông: 35 TB
   2. TTVT Sơn Tây: 28 TB
   3. TTVT Ba Đình: 20 TB
   4. TTVT Hoàn Kiếm: 15 TB
   5. TTVT Thanh Xuân: 10 TB

📊 Top 5 NVKT có nhiều TB hoàn công nhất:
   1. Nguyễn Quảng Ba (TTVT Hà Đông): 12 TB
   2. Trần Văn C (TTVT Sơn Tây): 10 TB
   3. Lê Thị D (TTVT Ba Đình): 9 TB
   4. Phạm Văn E (TTVT Hà Đông): 8 TB
   5. Hoàng Thị F (TTVT Hoàn Kiếm): 7 TB

✅ Hoàn thành xử lý báo cáo Hoàn công!

=== Bắt đầu tạo báo cáo Thực tăng ===
Đang đọc dữ liệu từ file Ngưng PSC...
Đang đọc dữ liệu từ file Hoàn công...

📊 Đang tạo báo cáo Thực tăng theo Tổ...
✅ Đã tạo thống kê cho 8 tổ
📊 Đang tạo báo cáo Thực tăng theo NVKT...
✅ Đã tạo thống kê cho 45 NVKT

💾 Đang lưu file báo cáo Thực tăng...
✅ Đã lưu file: downloads/baocao_hanoi/thuc_tang_27102025.xlsx
   - Sheet 'thuc_tang_theo_to': Thống kê theo Tổ (9 dòng)
   - Sheet 'thuc_tang_theo_NVKT': Thống kê theo NVKT (46 dòng)

📊 Tổng quan:
   - Tổng Hoàn công: 120 TB
   - Tổng Ngưng PSC: 150 TB
   - Thực tăng: -30 TB

📊 Top 5 Tổ có Thực tăng cao nhất:
   1. TTVT Ba Đình: 5 TB (HC: 20, NP: 15)
   2. TTVT Thanh Xuân: 2 TB (HC: 12, NP: 10)
   3. TTVT Hoàn Kiếm: -3 TB (HC: 15, NP: 18)
   4. TTVT Hà Đông: -3 TB (HC: 35, NP: 38)
   5. TTVT Sơn Tây: -17 TB (HC: 28, NP: 45)

📊 Top 5 NVKT có Thực tăng cao nhất:
   1. Lê Thị D (TTVT Ba Đình): 6 TB (HC: 9, NP: 3)
   2. Hoàng Thị F (TTVT Hoàn Kiếm): 4 TB (HC: 7, NP: 3)
   3. Nguyễn Quảng Ba (TTVT Hà Đông): 0 TB (HC: 12, NP: 12)
   4. Trần Văn C (TTVT Sơn Tây): -2 TB (HC: 10, NP: 12)
   5. Phạm Văn E (TTVT Hà Đông): -2 TB (HC: 8, NP: 10)

✅ Hoàn thành tạo báo cáo Thực tăng!

✅ Hoàn thành toàn bộ quá trình!
```

## Tùy chỉnh

### Thay đổi ngày cho báo cáo cụ thể

Nếu muốn dùng ngày khác thay vì ngày hiện tại:

```python
# Thay vì:
current_date = datetime.now().strftime("%d/%m/%Y")

# Dùng:
current_date = "21/10/2025"  # Ngày cụ thể
encoded_date = quote(current_date, safe='')
```

### Thêm báo cáo mới

1. Lấy URL báo cáo từ trình duyệt
2. Tạo function mới theo mẫu `download_report_pttb_ngung_psc` hoặc `download_report_pttb_hoan_cong`
3. Thêm vào hàm `main()`

### Thay đổi thư mục lưu file

```python
download_dir = os.path.join("downloads", "baocao_hanoi")
# Hoặc đường dẫn tuyệt đối:
download_dir = r"C:\Users\YourName\Reports"
```

## Bảo mật

⚠️ **Lưu ý**: File này chứa thông tin đăng nhập hardcoded. Không commit lên Git hoặc chia sẻ công khai.

Nên sử dụng:
- Biến môi trường
- File config riêng (thêm vào .gitignore)
- Keyring/credential manager
