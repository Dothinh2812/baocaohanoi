# Hướng dẫn sử dụng hệ thống Login - VNPT Hà Nội Dashboard

## 📋 Tổng quan

Hệ thống login đã được tích hợp hoàn chỉnh vào Dashboard với các tính năng:
- ✅ Đăng nhập/Đăng xuất
- ✅ Đổi mật khẩu (bắt buộc lần đầu)
- ✅ Phân quyền Admin/User
- ✅ Quản lý users (chỉ admin)
- ✅ Reset password (chỉ admin)
- ✅ Ghi log đăng nhập
- ✅ Session timeout 12 giờ

## 🔐 Thông tin đăng nhập

### Admin
- **Username:** `thinhdx.hni`
- **Password mặc định:** `abc123`
- **Quyền:** Quản trị toàn bộ hệ thống

### Users thông thường
- **Username:** Theo file `username.xlsx` (ví dụ: `nhunv.hni`, `minhda.hni`, ...)
- **Password mặc định:** `abc123`
- **Quyền:** Xem dashboard

**Tổng số users:** 156 users

## 🚀 Khởi động hệ thống

### Development mode
```bash
cd /home/vtst/dash
source venv/bin/activate
python3 dashboard.py
```

### Production mode (khuyến nghị)
```bash
cd /home/vtst/dash
source venv/bin/activate
gunicorn -c gunicorn_config.py dashboard:app
```

Sau đó truy cập: `http://localhost:5009`

## 📖 Hướng dẫn sử dụng

### 1. Đăng nhập lần đầu

1. Truy cập trang login: `http://localhost:5009/login`
2. Nhập username (ví dụ: `thinhdx.hni`)
3. Nhập password mặc định: `abc123`
4. Click "Đăng nhập"
5. **Quan trọng:** Hệ thống sẽ tự động chuyển đến trang đổi mật khẩu
6. Nhập mật khẩu cũ: `abc123`
7. Nhập mật khẩu mới (tối thiểu 6 ký tự)
8. Xác nhận mật khẩu mới
9. Click "Đổi mật khẩu"

### 2. Đổi mật khẩu (sau khi đã login)

1. Click vào tên user ở góc trên bên phải
2. Chọn "Đổi mật khẩu"
3. Nhập mật khẩu hiện tại
4. Nhập mật khẩu mới
5. Xác nhận mật khẩu mới
6. Click "Đổi mật khẩu"

### 3. Đăng xuất

1. Click vào tên user ở góc trên bên phải
2. Chọn "Đăng xuất"

### 4. Quản lý Users (chỉ Admin)

#### Xem danh sách users
1. Đăng nhập với tài khoản admin
2. Click vào tên admin ở góc trên bên phải
3. Chọn "Quản lý Users"
4. Sẽ thấy danh sách tất cả 156 users

#### Tìm kiếm user
- Nhập tên hoặc username vào ô tìm kiếm
- Danh sách sẽ tự động lọc theo từ khóa

#### Reset mật khẩu user
1. Tìm user cần reset
2. Click nút "Reset PW"
3. Xác nhận
4. Mật khẩu của user sẽ được reset về `abc123`
5. User phải đổi mật khẩu khi đăng nhập lần tiếp theo

#### Vô hiệu hóa/Kích hoạt user
1. Tìm user cần thao tác
2. Click nút "Vô hiệu" để vô hiệu hóa user
3. Click nút "Kích hoạt" để kích hoạt lại user đã bị vô hiệu hóa
4. User bị vô hiệu hóa không thể đăng nhập

## 📊 Thống kê và Log

### Thống kê users (Admin)
Trang quản lý users hiển thị:
- Tổng số users
- Users đang hoạt động
- Số admins
- Users chưa đổi mật khẩu

### Log đăng nhập
File log: `/home/vtst/dash/logs/login_history.csv`

Ghi lại:
- Timestamp: Thời gian
- Username: Tên đăng nhập
- IP Address: Địa chỉ IP
- User Agent: Thông tin trình duyệt
- Action: Hành động (login/logout/change_password/reset_password/...)
- Status: Trạng thái (success/failed)

## 🔒 Bảo mật

### Mật khẩu
- Được hash bằng `pbkdf2:sha256` (rất an toàn)
- Không lưu plain text
- Yêu cầu tối thiểu 6 ký tự

### Session
- Timeout: **12 giờ** không hoạt động
- HttpOnly cookies (chống XSS)
- SameSite: Lax (chống CSRF)

### Phân quyền
- User thường: Chỉ xem dashboard
- Admin: Quản lý toàn bộ users

## 📂 Cấu trúc files

```
/home/vtst/dash/
├── dashboard.py              # Backend Flask (đã cập nhật)
├── auth.py                   # Module xác thực
├── username.xlsx             # Database users (Excel)
├── setup_users.py            # Script khởi tạo users
├── templates/
│   ├── login.html           # Trang đăng nhập
│   ├── change_password.html # Trang đổi mật khẩu
│   └── admin_users.html     # Trang quản lý users
├── dashboard.html            # Dashboard chính (đã cập nhật)
├── logs/
│   └── login_history.csv    # Log đăng nhập
└── flask_session/            # Session storage (tự động tạo)
```

## ❓ Troubleshooting

### Lỗi: "Module not found: flask_session"
```bash
source venv/bin/activate
pip install flask-session
```

### Lỗi: "Permission denied" khi ghi Excel
```bash
chmod 666 username.xlsx
```

### Lỗi: Session không hoạt động
```bash
# Xóa folder flask_session và restart
rm -rf flask_session/
python3 dashboard.py
```

### Quên mật khẩu admin
```bash
# Chạy lại script setup để reset tất cả về abc123
source venv/bin/activate
python3 setup_users.py
```

### Không thể đăng nhập
1. Kiểm tra username có trong file `username.xlsx`
2. Kiểm tra cột `is_active` = True
3. Kiểm tra log: `logs/login_history.csv`

## 🔧 Cấu hình nâng cao

### Thay đổi session timeout
Sửa file `dashboard.py`:
```python
app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(hours=12)  # Đổi số giờ ở đây
```

### Thay đổi password mặc định
Sửa file `setup_users.py`:
```python
default_password_hash = generate_password_hash('abc123', method='pbkdf2:sha256')
# Thay 'abc123' thành mật khẩu mong muốn
```

### Thêm admin mới
1. Mở file `username.xlsx`
2. Tìm user cần phân quyền
3. Đổi cột `role` từ `user` thành `admin`
4. Lưu file
5. Reload trang

## 📞 Hỗ trợ

Nếu gặp vấn đề, kiểm tra:
1. Log file: `logs/login_history.csv`
2. Gunicorn error log: `logs/gunicorn_error.log`
3. File Excel: `username.xlsx` (đảm bảo có đủ cột)

---

**Phiên bản:** 1.0
**Ngày cập nhật:** 2025-11-18
**Người phát triển:** VNPT Hà Nội IT Team
