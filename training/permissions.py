"""Permission service cho module đào tạo.

Admin role thỏa mãn mọi yêu cầu (spec §7.1). Quyền module lưu trong
training_user_roles, độc lập cột role của users.xlsx.
"""

from training.db import read_connection
from training.errors import ErrorCode, TrainingError

PERMISSION_SCOPE_DENIED = ErrorCode.PERMISSION_SCOPE_DENIED
ADMIN_ROLE = "admin"


def _user_roles(db_path, username):
    conn = read_connection(db_path)
    try:
        rows = conn.execute(
            "SELECT role_code FROM training_user_roles WHERE username=?",
            (username,),
        ).fetchall()
        return {row["role_code"] for row in rows}
    finally:
        conn.close()


def has_module_role(db_path, username, required_role):
    """True nếu user có required_role hoặc là admin."""
    roles = _user_roles(db_path, username)
    return required_role in roles or ADMIN_ROLE in roles


def module_roles(db_path, username):
    """Trả các role module đã cấp cho người dùng."""
    return _user_roles(db_path, username) if username else set()


def require_module_role(db_path, username, required_role):
    """Raise TrainingError(PERMISSION_SCOPE_DENIED) nếu không có quyền."""
    if not has_module_role(db_path, username, required_role):
        raise TrainingError(
            ErrorCode.PERMISSION_SCOPE_DENIED,
            f"Người dùng {username!r} không có quyền {required_role!r}.",
            status=403,
        )


def user_audience_codes(db_path, username):
    """Trả danh sách audience_code của user."""
    conn = read_connection(db_path)
    try:
        rows = conn.execute(
            "SELECT audience_code FROM training_user_audiences WHERE username=?",
            (username,),
        ).fetchall()
        return [row["audience_code"] for row in rows]
    finally:
        conn.close()
