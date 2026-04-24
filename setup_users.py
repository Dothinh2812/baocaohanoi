#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script để cập nhật file username.xlsx với các cột mới cho hệ thống login
"""
import pandas as pd
from werkzeug.security import generate_password_hash

# Đọc file Excel hiện tại
excel_file = 'username.xlsx'
df = pd.read_excel(excel_file)

print(f"Đọc file {excel_file}...")
print(f"Số lượng users: {len(df)}")

# Thêm các cột mới nếu chưa có
if 'password' not in df.columns:
    # Hash mật khẩu mặc định 'abc123' cho tất cả users
    default_password_hash = generate_password_hash('abc123', method='pbkdf2:sha256')
    df['password'] = default_password_hash
    print("✓ Đã thêm cột 'password' với mật khẩu mặc định: abc123")

if 'role' not in df.columns:
    # Tất cả users mặc định là 'user', chỉ thinhdx.hni là 'admin'
    df['role'] = df['username'].apply(lambda x: 'admin' if x == 'thinhdx.hni' else 'user')
    print("✓ Đã thêm cột 'role' - admin: thinhdx.hni")

if 'is_first_login' not in df.columns:
    # Tất cả users cần đổi mật khẩu lần đầu
    df['is_first_login'] = True
    print("✓ Đã thêm cột 'is_first_login' = True")

if 'is_active' not in df.columns:
    # Tất cả users đều active
    df['is_active'] = True
    print("✓ Đã thêm cột 'is_active' = True")

# Lưu lại file Excel
df.to_excel(excel_file, index=False)
print(f"\n✅ Đã cập nhật file {excel_file} thành công!")
print(f"\nCấu trúc mới:")
print(df.head())
print(f"\nThông tin admin:")
admin = df[df['username'] == 'thinhdx.hni']
if not admin.empty:
    print(admin.to_string())
else:
    print("⚠️  CẢNH BÁO: Không tìm thấy user 'thinhdx.hni'")
