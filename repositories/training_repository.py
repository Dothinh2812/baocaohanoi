"""Data access layer cho training.db.

Helper dùng chung: sinh ID, ghi audit log. Không giữ transaction khi gọi
AI/Excel/file I/O.
"""

import uuid

from training import time_policy


def gen_id(prefix="id"):
    """Sinh ID dạng <prefix>_<uuid_hex>."""
    return f"{prefix}_{uuid.uuid4().hex}"


def write_audit(
    conn,
    *,
    actor,
    unit_code,
    action,
    entity_type,
    entity_id,
    before=None,
    after=None,
    request_id=None,
    ip=None,
    user_agent=None,
):
    conn.execute(
        """
        INSERT INTO training_audit_log
            (id, actor, unit_code, action, entity_type, entity_id,
             before_json, after_json, request_id, ip, user_agent, created_at_ms)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            gen_id("aud"),
            actor,
            unit_code,
            action,
            entity_type,
            str(entity_id),
            _json(before),
            _json(after),
            request_id,
            ip,
            user_agent,
            time_policy.utc_now_ms(),
        ),
    )


def _json(value):
    import json

    if value is None:
        return None
    return json.dumps(value, ensure_ascii=False, default=str)
