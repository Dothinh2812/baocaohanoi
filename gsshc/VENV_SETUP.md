# 🐍 Virtual Environment Setup

## ✅ Hoàn Thành!

Virtual environment đã được tạo và tất cả dependencies đã cài đặt thành công.

---

## 📁 Cấu Trúc

```
/home/vtst/s2/
├── venv/                          # Virtual environment
│   ├── bin/
│   │   ├── python3
│   │   ├── pip
│   │   ├── activate              # Script để activate
│   │   └── ...
│   └── lib/                       # Installed packages
├── gsmnv.py                       # Main application
├── requirements.txt               # Dependencies list
├── run.sh                         # Script để chạy ứng dụng
└── ...
```

---

## 🚀 Cách Chạy Ứng Dụng

### Option 1: Sử Dụng Script (Dễ Nhất)

```bash
cd /home/vtst/s2

# Chạy với port mặc định (8007)
./run.sh

# Hoặc chạy với port tùy chỉnh
./run.sh 8009
```

### Option 2: Manual Activation

```bash
cd /home/vtst/s2

# Activate virtual environment
source venv/bin/activate

# Chạy ứng dụng
python3 gsmnv.py

# Hoặc với port tùy chỉnh
PORT=8009 python3 gsmnv.py

# Deactivate khi xong
deactivate
```

### Option 3: Direct (Không Activate)

```bash
cd /home/vtst/s2

# Chạy trực tiếp
venv/bin/python3 gsmnv.py
```

---

## 📦 Installed Packages

### Google Cloud APIs
- `google-cloud-vision` - OCR processing
- `google-cloud-storage` - Store images
- `google-auth` - Authentication

### Web Framework
- `fastapi` - Web API framework
- `uvicorn[standard]` - ASGI server
- `pydantic` - Data validation

### AI/ML
- `openai` - OpenAI API integration
- `protobuf` - Data serialization

### Google APIs
- `gspread` - Google Sheets integration
- `oauth2client` - Google authentication

### Telegram
- `python-telegram-bot` - Telegram bot API

### Data Processing
- `pandas` - Data manipulation
- `numpy` - Numerical computing
- `openpyxl` - Excel support

### HTTP/Requests
- `requests` - HTTP library
- `httpx` - Async HTTP client

### Other
- `datetime` - Date/time handling
- `six` - Python 2/3 compatibility
- `pytz` - Timezone support
- `tqdm` - Progress bars

---

## 🔍 Kiểm Tra Installation

```bash
# Activate virtual environment
source venv/bin/activate

# Kiểm tra Python version
python3 --version

# Kiểm tra pip version
pip --version

# List all installed packages
pip list

# Kiểm tra specific package
pip show google-cloud-vision
```

---

## 📝 Cập Nhật Dependencies

Nếu cần thêm package mới:

```bash
# Activate virtual environment
source venv/bin/activate

# Cài đặt package mới
pip install package-name

# Cập nhật requirements.txt
pip freeze > requirements.txt
```

---

## 🧹 Dọn Dẹp

### Xóa Virtual Environment (Nếu Cần)

```bash
# Deactivate trước
deactivate

# Xóa folder venv
rm -rf venv/
```

### Tạo Lại Virtual Environment

```bash
# Tạo venv mới
python3 -m venv venv

# Cài đặt dependencies
source venv/bin/activate
pip install -r requirements.txt
```

---

## 🐛 Troubleshooting

### Lỗi: "command not found: activate"

**Giải pháp:**
```bash
# Đảm bảo bạn đang ở đúng directory
cd /home/vtst/s2

# Sử dụng đúng path
source venv/bin/activate
```

### Lỗi: "ModuleNotFoundError"

**Giải pháp:**
```bash
# Chắc chắn virtual environment được activate
source venv/bin/activate

# Kiểm tra package được cài đặt
pip list | grep package-name

# Cài đặt lại nếu cần
pip install -r requirements.txt
```

### Lỗi: "Permission denied" khi chạy run.sh

**Giải pháp:**
```bash
# Cấp quyền execute
chmod +x run.sh

# Chạy lại
./run.sh
```

---

## 📊 Thông Tin Hữu Ích

### Virtual Environment là gì?

Virtual environment là một isolated Python environment giúp:
- ✅ Tách biệt dependencies từ system Python
- ✅ Tránh xung đột phiên bản package
- ✅ Dễ dàng quản lý dependencies
- ✅ Tạo reproducible environments

### Tại sao nên dùng?

- ✅ Không ảnh hưởng đến system packages
- ✅ Dễ dàng chia sẻ project
- ✅ Dễ dàng xóa/cập nhật
- ✅ Best practice trong Python development

---

## 🎯 Quick Commands

```bash
# Navigate to project
cd /home/vtst/s2

# Activate venv
source venv/bin/activate

# Check Python
python3 --version

# Run app
./run.sh

# Or manual run
python3 gsmnv.py

# Install new package
pip install package-name

# Update requirements
pip freeze > requirements.txt

# Deactivate
deactivate
```

---

## ✨ Summary

| Task | Command |
|------|---------|
| Activate venv | `source venv/bin/activate` |
| Run app | `./run.sh` or `./run.sh 8009` |
| Install package | `pip install package-name` |
| Update requirements | `pip freeze > requirements.txt` |
| List packages | `pip list` |
| Deactivate venv | `deactivate` |

---

## 🚀 Bây Giờ

Chạy ứng dụng:

```bash
cd /home/vtst/s2
./run.sh
```

Hoặc chạy với port khác:

```bash
./run.sh 8007
```

---

**Status: ✅ READY TO USE**
