"""Catalog service: danh mục lĩnh vực/chủ đề/đối tượng và RBAC helpers.

Quản lý cây phân loại và quyền module. Seeding mặc định cho MVP.
"""

from training import time_policy
from training.db import read_connection, write_connection
from repositories.training_repository import gen_id, write_audit

DEFAULT_DOMAINS = [
    ("quality", "Chất lượng dịch vụ", 1),
    ("subscriber_growth", "Phát triển thuê bao", 2),
    ("automatic_configuration", "Cấu hình tự động", 3),
    ("subscriber_marketing", "Tiếp thị thuê bao", 4),
    ("retention", "Gia hạn và duy trì thuê bao", 5),
    ("technical_operations", "Điều hành kỹ thuật", 6),
    ("materials", "Vật tư và tài sản", 7),
]

DEFAULT_AUDIENCES = [
    ("to_truong", "Tổ trưởng", None),
    ("nvkt", "Nhân viên kỹ thuật", None),
    ("b2a", "Nhân viên B2A", None),
]

MODULE_ROLES = ("learner", "editor", "exam_manager", "admin")


def seed_defaults(db_path, unit_code):
    """Seed danh mục mặc định nếu chưa có. Idempotent."""
    conn = write_connection(db_path)
    try:
        for code, name, sort_order in DEFAULT_DOMAINS:
            conn.execute(
                "INSERT INTO training_domains (code, name, sort_order) VALUES (?, ?, ?) "
                "ON CONFLICT(code) DO NOTHING",
                (code, name, sort_order),
            )
        for code, name, resp in DEFAULT_AUDIENCES:
            conn.execute(
                "INSERT INTO training_audiences (code, name, responsibility_text) "
                "VALUES (?, ?, ?) ON CONFLICT(code) DO NOTHING",
                (code, name, resp),
            )
        conn.commit()
    finally:
        conn.close()


def list_domains(db_path):
    conn = read_connection(db_path)
    try:
        return [dict(r) for r in conn.execute(
            "SELECT code, name, sort_order FROM training_domains ORDER BY sort_order"
        ).fetchall()]
    finally:
        conn.close()


def list_audiences(db_path):
    conn = read_connection(db_path)
    try:
        return [dict(r) for r in conn.execute(
            "SELECT code, name, responsibility_text FROM training_audiences ORDER BY code"
        ).fetchall()]
    finally:
        conn.close()


def list_user_roles(db_path, username):
    conn = read_connection(db_path)
    try:
        return [r["role_code"] for r in conn.execute(
            "SELECT role_code FROM training_user_roles WHERE username=?",
            (username,),
        ).fetchall()]
    finally:
        conn.close()


def grant_role(db_path, unit_code, actor, username, role_code):
    if role_code not in MODULE_ROLES:
        raise ValueError(f"role_code {role_code!r} không hợp lệ")
    conn = write_connection(db_path)
    try:
        conn.execute(
            "INSERT INTO training_user_roles (id, username, role_code, granted_at_ms) "
            "VALUES (?, ?, ?, ?) ON CONFLICT(username, role_code) DO NOTHING",
            (gen_id("role"), username, role_code, time_policy.utc_now_ms()),
        )
        write_audit(
            conn, actor=actor, unit_code=unit_code, action="grant_role",
            entity_type="user_role", entity_id=f"{username}:{role_code}",
        )
        conn.commit()
    finally:
        conn.close()


def revoke_role(db_path, unit_code, actor, username, role_code):
    conn = write_connection(db_path)
    try:
        conn.execute(
            "DELETE FROM training_user_roles WHERE username=? AND role_code=?",
            (username, role_code),
        )
        write_audit(
            conn, actor=actor, unit_code=unit_code, action="revoke_role",
            entity_type="user_role", entity_id=f"{username}:{role_code}",
        )
        conn.commit()
    finally:
        conn.close()


def set_user_audiences(db_path, unit_code, actor, username, audience_codes):
    conn = write_connection(db_path)
    try:
        conn.execute(
            "DELETE FROM training_user_audiences WHERE username=?", (username,)
        )
        for code in audience_codes:
            conn.execute(
                "INSERT INTO training_user_audiences (id, username, audience_code) "
                "VALUES (?, ?, ?)",
                (gen_id("aud"), username, code),
            )
        write_audit(
            conn, actor=actor, unit_code=unit_code, action="set_audiences",
            entity_type="user_audiences", entity_id=username,
            after={"audience_codes": list(audience_codes)},
        )
        conn.commit()
    finally:
        conn.close()
