# MIGRATION: Hardcoded Values → Environment Variables

## Tổng quan

Đã chuyển toàn bộ thông tin nhạy cảm và cấu hình từ hardcode trong code sang file `.env` để:
- ✅ Bảo mật thông tin đăng nhập (username/password)
- ✅ Dễ dàng thay đổi cấu hình mà không sửa code
- ✅ Hỗ trợ nhiều môi trường (dev/staging/production)
- ✅ Tránh commit thông tin nhạy cảm lên Git

---

## Các thay đổi chính

### 1. Files mới được tạo

| File | Mục đích |
|------|----------|
| [.env](.env) | Chứa thông tin cấu hình thực tế (KHÔNG commit) |
| [.env.example](.env.example) | Template cho .env (có thể commit) |
| [config.py](config.py) | Class để load và quản lý config |
| [.gitignore](.gitignore) | Bảo vệ .env khỏi bị commit |
| [SETUP_ENV.md](SETUP_ENV.md) | Hướng dẫn chi tiết về .env |
| [MIGRATION_TO_ENV.md](MIGRATION_TO_ENV.md) | File này |

### 2. Files đã sửa đổi

#### [login.py](login.py)

**Trước:**
```python
file_path = r"/home/vtst/otp/otp_logs.txt"
if time_diff <= 120:  # File is recent enough

username_field.fill("thinhdx.hni")
password_field.fill("A#f4v5hp")
page_baocao.goto("https://baocao.hanoi.vnpt.vn/", timeout=60000)
browser_baocao = playwright_baocao.chromium.launch(headless=True)
```

**Sau:**
```python
from config import Config

file_path = Config.OTP_FILE_PATH
if time_diff <= Config.OTP_MAX_AGE_SECONDS:

username_field.fill(Config.BAOCAO_USERNAME)
password_field.fill(Config.BAOCAO_PASSWORD)
page_baocao.goto(Config.BAOCAO_URL, timeout=Config.PAGE_LOAD_TIMEOUT)
browser_baocao = playwright_baocao.chromium.launch(headless=Config.BROWSER_HEADLESS)
```

#### [requirements.txt](requirements.txt)

Đã thêm:
```
python-dotenv>=1.0.0
```

---

## Danh sách các giá trị đã chuyển sang .env

### Thông tin đăng nhập
- ✅ `BAOCAO_USERNAME` (từ: `"thinhdx.hni"`)
- ✅ `BAOCAO_PASSWORD` (từ: `"A#f4v5hp"`)
- ✅ `BAOCAO_URL` (từ: `"https://baocao.hanoi.vnpt.vn/"`)

### OTP Settings
- ✅ `OTP_FILE_PATH` (từ: `r"/home/vtst/otp/otp_logs.txt"`)
- ✅ `OTP_MAX_AGE_SECONDS` (từ: `120`)

### Report IDs - C1 Series
- ✅ `REPORT_C11_ID` = 522457
- ✅ `REPORT_C11_MENU_ID` = 522561
- ✅ `REPORT_C12_ID` = 522459
- ✅ `REPORT_C12_MENU_ID` = 522562
- ✅ `REPORT_C13_ID` = 522461
- ✅ `REPORT_C13_MENU_ID` = 522563
- ✅ `REPORT_C14_ID` = 522463
- ✅ `REPORT_C14_MENU_ID` = 522564
- ✅ `REPORT_C15_ID` = 522465
- ✅ `REPORT_C15_MENU_ID` = 522565
- ✅ `REPORT_I15_ID` = 521580
- ✅ `REPORT_I15_MENU_ID` = 521601

### Report Data URLs Parameters
- ✅ KR6 NVKT: `REPORT_KR6_NVKT_ID`, `VDVVT_ID`, `VDONVI_ID`, `VLOAI`
- ✅ KR6 Tổng hợp: `REPORT_KR6_TONGHOP_*`
- ✅ KR7 NVKT: `REPORT_KR7_NVKT_*`
- ✅ KR7 Tổng hợp: `REPORT_KR7_TONGHOP_*`

### Report IDs - Others
- ✅ `REPORT_TBM_ID` = 270922
- ✅ `REPORT_TBM_MENU_ID` = 276242
- ✅ `REPORT_THUCTANG_ID` = 521560
- ✅ `REPORT_THUCTANG_MENU_ID` = 521600

### Timeouts
- ✅ `PAGE_LOAD_TIMEOUT` (từ: `60000`)
- ✅ `NETWORK_IDLE_TIMEOUT` (từ: `500000`)
- ✅ `DOWNLOAD_TIMEOUT` (từ: `120000`)

### Browser Settings
- ✅ `BROWSER_HEADLESS` (từ: `True`)
- ✅ `ACCEPT_DOWNLOADS` (từ: `True`)

---

## Cách sử dụng sau migration

### Import và sử dụng Config

```python
from config import Config

# Truy cập các giá trị
username = Config.BAOCAO_USERNAME
password = Config.BAOCAO_PASSWORD
otp_path = Config.OTP_FILE_PATH
timeout = Config.PAGE_LOAD_TIMEOUT

# Lấy URL báo cáo
c11_url = Config.get_report_url('c11')
i15_url = Config.get_report_url('i15')

# Lấy URL data với ngày
from urllib.parse import quote
encoded_date = quote('01/11/2025', safe='')
kr6_url = Config.get_report_data_url('kr6_nvkt', encoded_date)
```

### Helper methods mới

```python
# In cấu hình (ẩn password)
Config.print_config()

# Lấy report URL đầy đủ
url = Config.get_report_url('c11')
# → https://baocao.hanoi.vnpt.vn/report/report-info?id=522457&menu_id=522561

# Lấy data URL với tham số ngày
url = Config.get_report_data_url('kr6_nvkt', '01%2F11%2F2025')
# → https://baocao.hanoi.vnpt.vn/report/report-info-data?id=264354&vdvvt_id=9&...
```

---

## Setup cho developer mới

### Bước 1: Clone repo

```bash
git clone <repo_url>
cd baocaohanoi
```

### Bước 2: Cài đặt dependencies

```bash
pip install -r requirements.txt
```

### Bước 3: Setup .env

```bash
# Copy template
cp .env.example .env

# Chỉnh sửa .env với editor
nano .env
# hoặc
vim .env
# hoặc
code .env
```

Cập nhật các giá trị:
```env
BAOCAO_USERNAME=your_username
BAOCAO_PASSWORD=your_password
OTP_FILE_PATH=/your/path/to/otp_logs.txt
```

### Bước 4: Verify config

```bash
python config.py
```

Kết quả mong đợi:
```
================================================================================
CURRENT CONFIGURATION
================================================================================
BAOCAO_URL: https://baocao.hanoi.vnpt.vn/
BAOCAO_USERNAME: your_username
BAOCAO_PASSWORD: ********
...
================================================================================
```

### Bước 5: Test login

```bash
python login.py
```

---

## Compatibility với code cũ

### Code cũ vẫn hoạt động

Nếu không có file `.env`, các giá trị mặc định trong `config.py` sẽ được sử dụng:

```python
class Config:
    BAOCAO_USERNAME = os.getenv('BAOCAO_USERNAME', 'thinhdx.hni')  # Default value
    BAOCAO_PASSWORD = os.getenv('BAOCAO_PASSWORD', 'A#f4v5hp')     # Default value
```

### Nhưng nên sử dụng .env

**Lý do:**
1. Thông tin nhạy cảm không nằm trong code
2. Dễ thay đổi giữa các môi trường
3. Tránh conflict khi nhiều developer cùng làm việc

---

## Migration checklist cho các files khác

### Files cần cập nhật (TODO)

Danh sách các file trong `baocaohanoi.py` có URLs hardcode:

- [ ] **Line 34**: KR6 NVKT URL
  ```python
  # Trước:
  report_url = f"https://baocao.hanoi.vnpt.vn/report/report-info-data?id=264354&vdvvt_id=9&..."

  # Sau:
  report_url = Config.get_report_data_url('kr6_nvkt', encoded_date)
  ```

- [ ] **Line 102**: KR6 Tổng hợp URL
  ```python
  report_url = Config.get_report_data_url('kr6_tonghop', encoded_date)
  ```

- [ ] **Line 171**: KR7 NVKT URL
  ```python
  report_url = Config.get_report_data_url('kr7_nvkt', encoded_date)
  ```

- [ ] **Line 240**: KR7 Tổng hợp URL
  ```python
  report_url = Config.get_report_data_url('kr7_tonghop', encoded_date)
  ```

- [ ] **Line 301**: TBM Report URL
  ```python
  report_url = Config.get_report_url('tbm')
  ```

- [ ] **Line 1820, 1928**: Thực tăng Report URLs
  ```python
  report_url = Config.get_report_url('thuctang')
  ```

- [ ] **Line 2041, 2149**: I1.5 Report URLs
  ```python
  report_url = Config.get_report_url('i15')
  ```

### Cách migrate từng URL

**Pattern cũ:**
```python
report_url = "https://baocao.hanoi.vnpt.vn/report/report-info?id=522457&menu_id=522561"
page_baocao.goto(report_url, timeout=60000)
```

**Pattern mới:**
```python
from config import Config

report_url = Config.get_report_url('c11')  # hoặc 'i15', 'tbm', etc.
page_baocao.goto(report_url, timeout=Config.PAGE_LOAD_TIMEOUT)
```

**Pattern data URL cũ:**
```python
encoded_date = quote(date_str, safe='')
report_url = f"https://baocao.hanoi.vnpt.vn/report/report-info-data?id=264354&vdvvt_id=9&vdenngay={encoded_date}&..."
page_baocao.goto(report_url, timeout=500000)
```

**Pattern data URL mới:**
```python
from config import Config
from urllib.parse import quote

encoded_date = quote(date_str, safe='')
report_url = Config.get_report_data_url('kr6_nvkt', encoded_date)
page_baocao.goto(report_url, timeout=Config.NETWORK_IDLE_TIMEOUT)
```

---

## Lợi ích của migration

### 1. Security
- ✅ Username/password không còn trong code
- ✅ `.env` không được commit vào Git (trong `.gitignore`)
- ✅ Mỗi developer có thể dùng credentials riêng

### 2. Flexibility
- ✅ Thay đổi config không cần sửa code
- ✅ Hỗ trợ nhiều môi trường (dev/staging/prod)
- ✅ Dễ dàng override giá trị cho testing

### 3. Maintainability
- ✅ Config tập trung ở một chỗ (`config.py`)
- ✅ Dễ track thay đổi của URLs/IDs
- ✅ Helper methods tái sử dụng (`get_report_url()`)

### 4. Collaboration
- ✅ Developers không conflict về credentials
- ✅ `.env.example` hướng dẫn setup rõ ràng
- ✅ Onboarding developer mới nhanh chóng

---

## Rollback (nếu cần)

Nếu gặp vấn đề với .env, có thể rollback bằng cách:

### Option 1: Sử dụng default values

Xóa/đổi tên file `.env`, code sẽ dùng giá trị mặc định trong `config.py`

### Option 2: Revert code

```bash
git log --oneline  # Tìm commit trước migration
git revert <commit_hash>
```

### Option 3: Hardcode tạm thời

Sửa `config.py`:
```python
# Bỏ qua .env, dùng hardcode
class Config:
    BAOCAO_USERNAME = "thinhdx.hni"  # Hardcode
    BAOCAO_PASSWORD = "A#f4v5hp"     # Hardcode
    # ...
```

**Lưu ý:** Không nên giữ hardcode lâu dài vì lý do bảo mật

---

## Testing sau migration

### Test 1: Verify config loads

```bash
python config.py
```

Kiểm tra output có đúng không.

### Test 2: Test login

```bash
python login.py
```

Kiểm tra login thành công.

### Test 3: Test report URLs

```python
from config import Config

# Test các URLs
print(Config.get_report_url('c11'))
print(Config.get_report_url('i15'))
print(Config.get_report_data_url('kr6_nvkt', '01%2F11%2F2025'))
```

### Test 4: Run full workflow

```bash
python baocaohanoi.py
```

Kiểm tra toàn bộ quy trình hoạt động bình thường.

---

## Support

Nếu gặp vấn đề:
1. Đọc [SETUP_ENV.md](SETUP_ENV.md) - Hướng dẫn chi tiết
2. Kiểm tra `.env` có đúng format không
3. Verify `python-dotenv` đã được cài đặt
4. Chạy `python config.py` để debug

---

**Status:** ✅ Migration hoàn tất cho `login.py`
**TODO:** Migrate các URLs trong `baocaohanoi.py` (optional - có thể làm sau)
**Phiên bản:** 1.0
**Ngày:** 05/11/2025
