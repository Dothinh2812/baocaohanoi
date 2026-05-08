#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Module xác thực và quản lý users cho Dashboard
Lưu trữ user data trong file username.xlsx
"""
import pandas as pd
import csv
from datetime import datetime
from functools import wraps
from flask import session, redirect, url_for, request, flash, jsonify
from werkzeug.security import generate_password_hash, check_password_hash
import os
from threading import Lock

# Lock để đảm bảo thread-safe khi đọc/ghi Excel
file_lock = Lock()

EXCEL_FILE = os.getenv('DASHV4_USER_FILE', 'username.xlsx')
LOGIN_LOG_FILE = os.getenv('DASHV4_LOGIN_LOG_FILE', 'logs/login_history.csv')

# Đảm bảo thư mục logs tồn tại
login_log_dir = os.path.dirname(os.path.abspath(LOGIN_LOG_FILE))
if login_log_dir:
    os.makedirs(login_log_dir, exist_ok=True)

# Khởi tạo file log nếu chưa có
if not os.path.exists(LOGIN_LOG_FILE):
    with open(LOGIN_LOG_FILE, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(['timestamp', 'username', 'ip_address', 'user_agent', 'action', 'status'])


def get_all_users():
    """Đọc tất cả users từ Excel file"""
    with file_lock:
        try:
            df = pd.read_excel(EXCEL_FILE)
            return df
        except Exception as e:
            print(f"Lỗi khi đọc file Excel: {e}")
            return None


def get_user_by_username(username):
    """Lấy thông tin user theo username"""
    df = get_all_users()
    if df is None:
        return None

    user = df[df['username'] == username]
    if user.empty:
        return None

    return user.iloc[0].to_dict()


def update_user(username, updates):
    """
    Cập nhật thông tin user trong Excel
    updates: dict chứa các cột cần update
    """
    with file_lock:
        try:
            df = pd.read_excel(EXCEL_FILE)

            # Tìm user
            user_idx = df[df['username'] == username].index
            if user_idx.empty:
                return False

            # Update các trường
            for key, value in updates.items():
                df.loc[user_idx, key] = value

            # Lưu lại
            df.to_excel(EXCEL_FILE, index=False)
            return True
        except Exception as e:
            print(f"Lỗi khi update user: {e}")
            return False


def verify_password(username, password):
    """Xác thực username và password"""
    user = get_user_by_username(username)
    if not user:
        return False, "Tên đăng nhập không tồn tại"

    if not user.get('is_active', True):
        return False, "Tài khoản đã bị vô hiệu hóa"

    if check_password_hash(user['password'], password):
        return True, user

    return False, "Mật khẩu không đúng"


def change_user_password(username, old_password, new_password):
    """Đổi mật khẩu user"""
    # Xác thực mật khẩu cũ
    success, result = verify_password(username, old_password)
    if not success:
        return False, result

    # Hash mật khẩu mới
    new_password_hash = generate_password_hash(new_password, method='pbkdf2:sha256')

    # Update vào Excel
    updates = {
        'password': new_password_hash,
        'is_first_login': False
    }

    if update_user(username, updates):
        return True, "Đổi mật khẩu thành công"
    else:
        return False, "Lỗi khi lưu mật khẩu mới"


def reset_user_password(username, new_password='abc123'):
    """Reset mật khẩu user về mặc định (chỉ admin)"""
    new_password_hash = generate_password_hash(new_password, method='pbkdf2:sha256')

    updates = {
        'password': new_password_hash,
        'is_first_login': True
    }

    if update_user(username, updates):
        return True, f"Đã reset mật khẩu về: {new_password}"
    else:
        return False, "Lỗi khi reset mật khẩu"


def toggle_user_active(username):
    """Kích hoạt/vô hiệu hóa user"""
    user = get_user_by_username(username)
    if not user:
        return False, "User không tồn tại"

    new_status = not user.get('is_active', True)
    updates = {'is_active': new_status}

    if update_user(username, updates):
        status_text = "kích hoạt" if new_status else "vô hiệu hóa"
        return True, f"Đã {status_text} user {username}"
    else:
        return False, "Lỗi khi cập nhật trạng thái"


def update_user_role(username, new_role):
    """Cập nhật role của user (admin/user)"""
    if new_role not in ['admin', 'user']:
        return False, "Role không hợp lệ"

    updates = {'role': new_role}

    if update_user(username, updates):
        return True, f"Đã đổi role của {username} thành {new_role}"
    else:
        return False, "Lỗi khi cập nhật role"


def log_login_activity(username, ip_address, user_agent, action, status):
    """Ghi log hoạt động đăng nhập"""
    try:
        with open(LOGIN_LOG_FILE, 'a', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            writer.writerow([timestamp, username, ip_address, user_agent, action, status])
    except Exception as e:
        print(f"Lỗi khi ghi log: {e}")


def login_required(f):
    """Decorator yêu cầu đăng nhập"""
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'username' not in session:
            if request.path.startswith('/api/'):
                return jsonify({'error': 'Vui lòng đăng nhập để tiếp tục'}), 401
            flash('Vui lòng đăng nhập để tiếp tục', 'warning')
            return redirect(url_for('auth.login'))
        return f(*args, **kwargs)
    return decorated_function


def admin_required(f):
    """Decorator yêu cầu quyền admin"""
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'username' not in session:
            if request.path.startswith('/api/'):
                return jsonify({'error': 'Vui lòng đăng nhập để tiếp tục'}), 401
            flash('Vui lòng đăng nhập để tiếp tục', 'warning')
            return redirect(url_for('auth.login'))

        user = get_user_by_username(session['username'])
        if not user or user.get('role') != 'admin':
            if request.path.startswith('/api/'):
                return jsonify({'error': 'Bạn không có quyền truy cập tài nguyên này'}), 403
            flash('Bạn không có quyền truy cập trang này', 'danger')
            return redirect(url_for('index'))

        return f(*args, **kwargs)
    return decorated_function


def get_login_stats():
    """Lấy thống kê đăng nhập (cho admin)"""
    try:
        df = pd.read_csv(LOGIN_LOG_FILE)
        return {
            'total_logins': len(df[df['action'] == 'login']),
            'failed_logins': len(df[(df['action'] == 'login') & (df['status'] == 'failed')]),
            'recent_activities': df.tail(20).to_dict('records')
        }
    except Exception as e:
        print(f"Lỗi khi đọc log: {e}")
        return None
