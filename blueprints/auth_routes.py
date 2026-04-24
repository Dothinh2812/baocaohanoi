from urllib.parse import urlencode

from flask import Blueprint, current_app, flash, redirect, render_template, request, session, url_for

from app_helpers import csrf_protect
from auth import (
    admin_required,
    change_user_password,
    get_all_users,
    get_user_by_username,
    log_login_activity,
    login_required,
    reset_user_password,
    toggle_user_active,
    verify_password,
)


auth_bp = Blueprint('auth', __name__)


@auth_bp.route('/login', methods=['GET', 'POST'])
@csrf_protect
def login():
    if 'username' in session:
        return redirect(url_for('index'))

    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '')

        if not username or not password:
            flash('Vui lòng nhập đầy đủ tên đăng nhập và mật khẩu', 'danger')
            return render_template('login.html')

        success, result = verify_password(username, password)
        if success:
            user = result
            session['username'] = username
            session['name'] = user.get('name', username)
            session['role'] = user.get('role', 'user')

            log_login_activity(
                username=username,
                ip_address=request.remote_addr,
                user_agent=request.headers.get('User-Agent', ''),
                action='login',
                status='success',
            )

            if user.get('is_first_login', False):
                flash('Vui lòng đổi mật khẩu mặc định để tiếp tục', 'warning')
                return redirect(url_for('auth.change_password'))

            flash(f'Xin chào {user.get("name", username)}!', 'success')
            return redirect(url_for('index'))

        log_login_activity(
            username=username,
            ip_address=request.remote_addr,
            user_agent=request.headers.get('User-Agent', ''),
            action='login',
            status='failed',
        )
        flash(result, 'danger')

    return render_template('login.html')


@auth_bp.route('/logout')
def logout():
    username = session.get('username', 'unknown')
    log_login_activity(
        username=username,
        ip_address=request.remote_addr,
        user_agent=request.headers.get('User-Agent', ''),
        action='logout',
        status='success',
    )
    session.clear()
    flash('Đã đăng xuất thành công', 'success')
    return redirect(url_for('auth.login'))


@auth_bp.route('/external/sh-portal')
@login_required
def redirect_sh_portal():
    username = session.get('username', '').strip()
    if not username:
        flash('Không xác định được tài khoản đăng nhập hiện tại', 'danger')
        return redirect(url_for('index'))

    target_url = current_app.config['SH_PORTAL_URL'].strip()
    username_param = current_app.config['SH_PORTAL_USERNAME_PARAM'].strip() or 'username'

    separator = '&' if '?' in target_url else '?'
    redirect_url = f"{target_url}{separator}{urlencode({username_param: username})}"
    return redirect(redirect_url)


@auth_bp.route('/change-password', methods=['GET', 'POST'])
@login_required
@csrf_protect
def change_password():
    username = session.get('username')
    user = get_user_by_username(username)

    if request.method == 'POST':
        old_password = request.form.get('old_password', '')
        new_password = request.form.get('new_password', '')
        confirm_password = request.form.get('confirm_password', '')

        if not old_password or not new_password or not confirm_password:
            flash('Vui lòng nhập đầy đủ thông tin', 'danger')
            return render_template('change_password.html', current_user=user, is_first_login=user.get('is_first_login', False))

        if new_password != confirm_password:
            flash('Mật khẩu mới và xác nhận không khớp', 'danger')
            return render_template('change_password.html', current_user=user, is_first_login=user.get('is_first_login', False))

        if len(new_password) < 6:
            flash('Mật khẩu phải có ít nhất 6 ký tự', 'danger')
            return render_template('change_password.html', current_user=user, is_first_login=user.get('is_first_login', False))

        success, message = change_user_password(username, old_password, new_password)
        if success:
            log_login_activity(
                username=username,
                ip_address=request.remote_addr,
                user_agent=request.headers.get('User-Agent', ''),
                action='change_password',
                status='success',
            )
            flash('Đổi mật khẩu thành công!', 'success')
            return redirect(url_for('index'))

        flash(message, 'danger')

    return render_template('change_password.html', current_user=user, is_first_login=user.get('is_first_login', False))


@auth_bp.route('/admin/users')
@admin_required
def admin_users():
    users_df = get_all_users()
    if users_df is None:
        flash('Lỗi khi đọc danh sách users', 'danger')
        return redirect(url_for('index'))

    stats = {
        'total': len(users_df),
        'active': len(users_df[users_df['is_active'] == True]),
        'admins': len(users_df[users_df['role'] == 'admin']),
        'first_login': len(users_df[users_df['is_first_login'] == True]),
    }
    return render_template('admin_users.html', users=users_df, stats=stats)


@auth_bp.route('/admin/reset-password/<username>', methods=['POST'])
@admin_required
@csrf_protect
def admin_reset_password(username):
    success, message = reset_user_password(username)
    if success:
        log_login_activity(
            username=f"admin:{session.get('username')}",
            ip_address=request.remote_addr,
            user_agent=request.headers.get('User-Agent', ''),
            action=f'reset_password:{username}',
            status='success',
        )
        flash(message, 'success')
    else:
        flash(message, 'danger')

    return redirect(url_for('auth.admin_users'))


@auth_bp.route('/admin/toggle-user/<username>', methods=['POST'])
@admin_required
@csrf_protect
def admin_toggle_user(username):
    if username == session.get('username'):
        flash('Không thể vô hiệu hóa tài khoản của chính mình', 'danger')
        return redirect(url_for('auth.admin_users'))

    success, message = toggle_user_active(username)
    if success:
        log_login_activity(
            username=f"admin:{session.get('username')}",
            ip_address=request.remote_addr,
            user_agent=request.headers.get('User-Agent', ''),
            action=f'toggle_user:{username}',
            status='success',
        )
        flash(message, 'success')
    else:
        flash(message, 'danger')

    return redirect(url_for('auth.admin_users'))
