# HƯỚNG DẪN CẤU HÌNH FILE .ENV

## Tổng quan

Hệ thống sử dụng file `.env` để quản lý các thông tin nhạy cảm và cấu hình như:
- Username/Password đăng nhập
- URLs của các báo cáo
- Đường dẫn file OTP
- Timeouts và cấu hình khác

## Cài đặt lần đầu

### 1. Cài đặt thư viện python-dotenv

```bash
pip install python-dotenv
```

Hoặc thêm vào `requirements.txt`:
```
python-dotenv
```

### 2. Tạo file .env từ template

```bash
cp .env.example .env
```

### 3. Chỉnh sửa file .env

Mở file `.env` và cập nhật các giá trị:

```env
# Thông tin đăng nhập (BẮT BUỘC)
BAOCAO_USERNAME=your_username_here
BAOCAO_PASSWORD=your_password_here

# Đường dẫn file OTP (BẮT BUỘC nếu dùng OTP tự động)
OTP_FILE_PATH=/path/to/your/otp_logs.txt

# Các giá trị khác có thể giữ nguyên mặc định
```

---

## Cấu trúc file .env

### Thông tin đăng nhập

```env
BAOCAO_USERNAME=your_username
BAOCAO_PASSWORD=your_password
BAOCAO_URL=https://baocao.hanoi.vnpt.vn/
```

### Cấu hình OTP

```env
OTP_FILE_PATH=/home/vtst/otp/otp_logs.txt
OTP_MAX_AGE_SECONDS=120
```

- `OTP_FILE_PATH`: Đường dẫn tuyệt đối đến file chứa mã OTP
- `OTP_MAX_AGE_SECONDS`: Thời gian tối đa (giây) file OTP được coi là hợp lệ

### Report IDs

Các ID này được lấy từ URLs của hệ thống báo cáo VNPT:

```env
# C1.1 Report
REPORT_C11_ID=522457
REPORT_C11_MENU_ID=522561

# C1.2 Report
REPORT_C12_ID=522459
REPORT_C12_MENU_ID=522562

# ... và các report khác
```

### Timeouts

```env
PAGE_LOAD_TIMEOUT=60000         # 60 giây
NETWORK_IDLE_TIMEOUT=500000     # 500 giây
DOWNLOAD_TIMEOUT=120000         # 120 giây
```

Đơn vị: milliseconds (1000ms = 1 giây)

### Browser Settings

```env
BROWSER_HEADLESS=True           # True = chạy ẩn, False = hiện browser
ACCEPT_DOWNLOADS=True           # Cho phép tải file
```

---

## Sử dụng trong code

### Import Config class

```python
from config import Config

# Sử dụng các giá trị
print(Config.BAOCAO_USERNAME)
print(Config.OTP_FILE_PATH)
print(Config.PAGE_LOAD_TIMEOUT)
```

### Lấy Report URLs

```python
from config import Config

# Lấy URL đầy đủ của báo cáo C1.1
url = Config.get_report_url('c11')
# Output: https://baocao.hanoi.vnpt.vn/report/report-info?id=522457&menu_id=522561

# Lấy URL báo cáo I1.5
url = Config.get_report_url('i15')

# Lấy URL data report với ngày
from urllib.parse import quote
encoded_date = quote('01/11/2025', safe='')
url = Config.get_report_data_url('kr6_nvkt', encoded_date)
```

### In cấu hình hiện tại

```python
from config import Config

Config.print_config()
```

Output:
```
================================================================================
CURRENT CONFIGURATION
================================================================================
BAOCAO_URL: https://baocao.hanoi.vnpt.vn/
BAOCAO_USERNAME: thinhdx.hni
BAOCAO_PASSWORD: ********
OTP_FILE_PATH: /home/vtst/otp/otp_logs.txt
OTP_MAX_AGE_SECONDS: 120
PAGE_LOAD_TIMEOUT: 60000ms
BROWSER_HEADLESS: True
================================================================================
```

---

## Bảo mật

### ⚠️ QUAN TRỌNG

1. **KHÔNG ĐƯỢC commit file .env vào Git!**
   - File `.env` đã được thêm vào `.gitignore`
   - Chỉ commit file `.env.example` (không chứa thông tin nhạy cảm)

2. **Backup file .env an toàn:**
   ```bash
   # Backup vào thư mục riêng tư
   cp .env ~/.baocao_env_backup

   # Hoặc encrypt trước khi backup
   gpg -c .env  # Tạo file .env.gpg
   ```

3. **Chia sẻ cấu hình:**
   - Không gửi file `.env` qua email/chat
   - Sử dụng các công cụ password manager (1Password, Bitwarden)
   - Hoặc gửi từng giá trị riêng lẻ qua kênh bảo mật

4. **Permissions trên server:**
   ```bash
   # Chỉ owner được đọc/ghi
   chmod 600 .env

   # Kiểm tra permissions
   ls -la .env
   # Output: -rw------- 1 user user 1234 Nov  5 10:00 .env
   ```

---

## Troubleshooting

### Lỗi: "No module named 'dotenv'"

**Giải pháp:**
```bash
pip install python-dotenv
```

### Lỗi: "File .env not found"

**Nguyên nhân:** Chưa tạo file .env hoặc file ở sai thư mục

**Giải pháp:**
```bash
# Kiểm tra file có tồn tại không
ls -la .env

# Tạo từ template nếu chưa có
cp .env.example .env
```

### Lỗi: Config không load giá trị từ .env

**Nguyên nhân:** File .env có syntax sai hoặc có khoảng trắng thừa

**Giải pháp:**
```env
# ❌ SAI - có khoảng trắng
BAOCAO_USERNAME = my_username

# ✅ ĐÚNG - không có khoảng trắng
BAOCAO_USERNAME=my_username

# ❌ SAI - dùng dấu ngoặc kép sai
BAOCAO_PASSWORD="my password"

# ✅ ĐÚNG - không cần ngoặc kép nếu không có khoảng trắng
BAOCAO_PASSWORD=my_password

# ✅ ĐÚNG - dùng ngoặc kép nếu có khoảng trắng
BAOCAO_PASSWORD="my complex password"
```

### Kiểm tra giá trị đang dùng

```python
from config import Config

# In ra tất cả config (ẩn password)
Config.print_config()

# Kiểm tra giá trị cụ thể
print(f"Username: {Config.BAOCAO_USERNAME}")
print(f"OTP Path: {Config.OTP_FILE_PATH}")
print(f"File exists: {os.path.exists(Config.OTP_FILE_PATH)}")
```

---

## Migration từ code cũ

### Trước đây (hardcode):

```python
# login.py (CŨ)
username_field.fill("thinhdx.hni")
password_field.fill("A#f4v5hp")
page.goto("https://baocao.hanoi.vnpt.vn/", timeout=60000)
```

### Bây giờ (sử dụng .env):

```python
# login.py (MỚI)
from config import Config

username_field.fill(Config.BAOCAO_USERNAME)
password_field.fill(Config.BAOCAO_PASSWORD)
page.goto(Config.BAOCAO_URL, timeout=Config.PAGE_LOAD_TIMEOUT)
```

---

## Best Practices

1. **Môi trường khác nhau, file .env khác nhau:**
   ```
   .env.development   # Môi trường dev
   .env.staging       # Môi trường staging
   .env.production    # Môi trường production
   ```

   Load theo môi trường:
   ```python
   import os
   from dotenv import load_dotenv

   env = os.getenv('APP_ENV', 'development')
   load_dotenv(f'.env.{env}')
   ```

2. **Validate config khi khởi động:**
   ```python
   # config.py
   class Config:
       @classmethod
       def validate(cls):
           if not cls.BAOCAO_USERNAME:
               raise ValueError("BAOCAO_USERNAME không được để trống")
           if not cls.BAOCAO_PASSWORD:
               raise ValueError("BAOCAO_PASSWORD không được để trống")
           # ... validate khác

   # Gọi khi app khởi động
   Config.validate()
   ```

3. **Log config khi chạy (ẩn sensitive info):**
   ```python
   import logging

   logging.info(f"Starting with config:")
   logging.info(f"  BAOCAO_URL: {Config.BAOCAO_URL}")
   logging.info(f"  USERNAME: {Config.BAOCAO_USERNAME}")
   logging.info(f"  PASSWORD: {'*' * 8}")
   ```

---

## Template cho .env.example

File `.env.example` nên chứa:
- Tất cả các keys cần thiết
- Giá trị mặc định hoặc placeholder
- Comments giải thích

```env
# Thông tin đăng nhập hệ thống báo cáo
BAOCAO_USERNAME=your_username_here
BAOCAO_PASSWORD=your_password_here
BAOCAO_URL=https://baocao.hanoi.vnpt.vn/

# Đường dẫn file OTP
# Đây là file chứa mã OTP 6 chữ số được tạo tự động
OTP_FILE_PATH=/path/to/your/otp_logs.txt
OTP_MAX_AGE_SECONDS=120

# ... các config khác với comments rõ ràng
```

---

## Checklist

Khi setup môi trường mới:

- [ ] Cài đặt `python-dotenv`
- [ ] Copy `.env.example` thành `.env`
- [ ] Cập nhật username/password trong `.env`
- [ ] Cập nhật đường dẫn OTP file
- [ ] Kiểm tra permissions của `.env` (600)
- [ ] Verify `.env` trong `.gitignore`
- [ ] Test config: `python config.py`
- [ ] Test login: `python login.py`

---

**Phiên bản:** 1.0
**Ngày tạo:** 05/11/2025
**Files liên quan:**
- [.env.example](.env.example): Template file
- [config.py](config.py): Config loader
- [login.py](login.py): Sử dụng Config
- [.gitignore](.gitignore): Bảo vệ .env
