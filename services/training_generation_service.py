"""Generation service: AI job queue với atomic claim, lease, retry.

Worker claim job bằng BEGIN IMMEDIATE. Lease hết hạn có thể được claim lại.
Completed job chỉ tạo drafts, không publish.
"""

import json

from training import constants, time_policy
from training.db import read_connection, write_connection
from training.errors import ErrorCode, TrainingError
from repositories.training_repository import gen_id, write_audit
from services.training_question_service import import_question_batch


def enqueue_job(
    db_path, *, unit_code, actor, source_document_version_ids,
    target_audience_codes, requested_count, idempotency_key=None,
    max_retries=3,
):
    now = time_policy.utc_now_ms()
    conn = write_connection(db_path)
    try:
        if idempotency_key:
            existing = conn.execute(
                "SELECT id FROM ai_generation_jobs WHERE idempotency_key=?",
                (idempotency_key,),
            ).fetchone()
            if existing:
                return existing["id"]
        job_id = gen_id("job")
        conn.execute(
            """INSERT INTO ai_generation_jobs
            (id, status, idempotency_key, request_payload_json,
             source_document_version_ids_json, target_audience_codes_json,
             requested_count, unit_code, created_by, created_at_ms, max_retries)
             VALUES (?, 'pending', ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                job_id, idempotency_key,
                json.dumps({"requested_count": requested_count}, ensure_ascii=False),
                json.dumps(source_document_version_ids, ensure_ascii=False),
                json.dumps(target_audience_codes, ensure_ascii=False),
                requested_count, unit_code, actor, now, max_retries,
            ),
        )
        write_audit(conn, actor=actor, unit_code=unit_code, action="enqueue_job",
                    entity_type="ai_generation_job", entity_id=job_id)
        conn.commit()
        return job_id
    finally:
        conn.close()


def claim_next_job(db_path, *, worker_id, lease_seconds):
    """Atomic claim via BEGIN IMMEDIATE. Trả job dict hoặc None."""
    now = time_policy.utc_now_ms()
    lease_expires = now + lease_seconds * 1000
    conn = write_connection(db_path)
    try:
        conn.execute("BEGIN IMMEDIATE")
        row = conn.execute(
            """SELECT * FROM ai_generation_jobs
            WHERE status = 'pending'
               OR (status = 'running' AND lease_expires_at_ms IS NOT NULL
                    AND lease_expires_at_ms <= ?)
            ORDER BY created_at_ms
            LIMIT 1""",
            (now,),
        ).fetchone()
        if not row:
            conn.execute("ROLLBACK")
            return None
        retry_count = row["retry_count"] + 1 if row["status"] == "running" else row["retry_count"]
        conn.execute(
            """UPDATE ai_generation_jobs
            SET status='running', claimed_by=?, lease_expires_at_ms=?, heartbeat_at_ms=?, retry_count=?
            WHERE id=?""",
            (worker_id, lease_expires, now, retry_count, row["id"]),
        )
        conn.commit()
        return dict(row) | {"status": "running", "claimed_by": worker_id, "retry_count": retry_count}
    except Exception:
        try:
            conn.execute("ROLLBACK")
        except Exception:
            pass
        raise
    finally:
        conn.close()


def heartbeat(db_path, *, job_id, worker_id, lease_seconds):
    now = time_policy.utc_now_ms()
    lease_expires = now + lease_seconds * 1000
    conn = write_connection(db_path)
    try:
        updated = conn.execute(
            """UPDATE ai_generation_jobs
            SET heartbeat_at_ms=?, lease_expires_at_ms=?
            WHERE id=? AND claimed_by=? AND status='running'
              AND lease_expires_at_ms IS NOT NULL AND lease_expires_at_ms > ?""",
            (now, lease_expires, job_id, worker_id, now),
        ).rowcount
        if updated != 1:
            raise TrainingError(ErrorCode.CONFLICT,
                                "Worker không còn sở hữu lease của job", status=409)
        conn.commit()
    finally:
        conn.close()


def complete_job(db_path, *, job_id, worker_id, batch, provider_metadata):
    """Lưu batch + import drafts. Completed job chỉ tạo drafts."""
    now = time_policy.utc_now_ms()
    conn = write_connection(db_path)
    try:
        conn.execute("BEGIN IMMEDIATE")
        job = conn.execute(
            "SELECT * FROM ai_generation_jobs WHERE id=?", (job_id,)
        ).fetchone()
        if not job:
            raise TrainingError(ErrorCode.NOT_FOUND, "Job không tồn tại", status=404)
        if (job["status"] != "running" or job["claimed_by"] != worker_id
                or job["lease_expires_at_ms"] is None
                or job["lease_expires_at_ms"] <= now):
            raise TrainingError(ErrorCode.CONFLICT,
                                "Worker không còn sở hữu lease của job", status=409)
        conn.execute(
            "INSERT OR REPLACE INTO ai_generation_batches "
            "(id, job_id, schema_version, provider, model, prompt_version, usage_json, "
            "raw_response_text, questions_payload_json, created_at_ms) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
            (
                gen_id("batch"), job_id,
                batch.get("schema_version", "1.0"),
                provider_metadata.get("provider", "unknown"),
                provider_metadata.get("model"),
                provider_metadata.get("prompt_version"),
                json.dumps(provider_metadata.get("usage"), ensure_ascii=False) if provider_metadata.get("usage") else None,
                provider_metadata.get("raw_response"),
                json.dumps(batch, ensure_ascii=False),
                now,
            ),
        )
        import_question_batch(
            db_path, unit_code=job["unit_code"], actor=job["created_by"],
            batch=batch, status="draft", conn=conn,
        )
        updated = conn.execute(
            """UPDATE ai_generation_jobs
            SET status='completed', completed_at_ms=?
            WHERE id=? AND status='running' AND claimed_by=?
              AND lease_expires_at_ms IS NOT NULL AND lease_expires_at_ms > ?""",
            (now, job_id, worker_id, now),
        ).rowcount
        if updated != 1:
            raise TrainingError(ErrorCode.CONFLICT,
                                "Worker không còn sở hữu lease của job", status=409)
        write_audit(conn, actor=worker_id, unit_code=job["unit_code"], action="complete_job",
                    entity_type="ai_generation_job", entity_id=job_id)
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def fail_job(db_path, *, job_id, worker_id, error_code, error_detail):
    now = time_policy.utc_now_ms()
    conn = write_connection(db_path)
    try:
        conn.execute("BEGIN IMMEDIATE")
        job = conn.execute(
            "SELECT retry_count, max_retries FROM ai_generation_jobs WHERE id=?",
            (job_id,),
        ).fetchone()
        if not job:
            raise TrainingError(ErrorCode.NOT_FOUND, "Job không tồn tại", status=404)
        now = time_policy.utc_now_ms()
        owned = conn.execute(
            """SELECT 1 FROM ai_generation_jobs
            WHERE id=? AND status='running' AND claimed_by=?
              AND lease_expires_at_ms IS NOT NULL AND lease_expires_at_ms > ?""",
            (job_id, worker_id, now),
        ).fetchone()
        if not owned:
            raise TrainingError(ErrorCode.CONFLICT,
                                "Worker không còn sở hữu lease của job", status=409)
        new_retry_count = job["retry_count"] + 1
        if new_retry_count >= job["max_retries"]:
            updated = conn.execute(
                """UPDATE ai_generation_jobs
                SET status='failed', error_code=?, error_detail=?, retry_count=?
                WHERE id=? AND status='running' AND claimed_by=?
                  AND lease_expires_at_ms IS NOT NULL AND lease_expires_at_ms > ?""",
                (error_code, error_detail, new_retry_count, job_id, worker_id, now),
            ).rowcount
        else:
            updated = conn.execute(
                """UPDATE ai_generation_jobs
                SET status='pending', error_code=?, error_detail=?,
                    claimed_by=NULL, lease_expires_at_ms=NULL, retry_count=?
                WHERE id=? AND status='running' AND claimed_by=?
                  AND lease_expires_at_ms IS NOT NULL AND lease_expires_at_ms > ?""",
                (error_code, error_detail, new_retry_count, job_id, worker_id, now),
            ).rowcount
        if updated != 1:
            raise TrainingError(ErrorCode.CONFLICT,
                                "Worker không còn sở hữu lease của job", status=409)
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def cancel_job(db_path, *, unit_code, actor, job_id):
    conn = write_connection(db_path)
    try:
        conn.execute(
            "UPDATE ai_generation_jobs SET status='cancelled' WHERE id=? AND status IN ('pending','running')",
            (job_id,),
        )
        write_audit(conn, actor=actor, unit_code=unit_code, action="cancel_job",
                    entity_type="ai_generation_job", entity_id=job_id)
        conn.commit()
    finally:
        conn.close()


def get_job(db_path, job_id):
    conn = read_connection(db_path)
    try:
        row = conn.execute(
            "SELECT * FROM ai_generation_jobs WHERE id=?", (job_id,)
        ).fetchone()
        return dict(row) if row else None
    finally:
        conn.close()


def list_jobs(db_path, *, status=None, actor=None, page=1, page_size=25):
    """Paginated list of generation jobs (non-sensitive columns only)."""
    page = max(1, page)
    page_size = max(1, min(page_size, 100))
    conn = read_connection(db_path)
    try:
        where = []
        params = []
        if status:
            where.append("status = ?")
            params.append(status)
        if actor:
            where.append("created_by = ?")
            params.append(actor)
        clause = ("WHERE " + " AND ".join(where)) if where else ""
        total = conn.execute(
            f"SELECT COUNT(*) AS c FROM ai_generation_jobs {clause}", params
        ).fetchone()["c"]
        offset = (page - 1) * page_size
        rows = conn.execute(
            f"""SELECT id, status, requested_count, created_by, created_at_ms,
                       claimed_by, retry_count, error_code, completed_at_ms
                FROM ai_generation_jobs {clause}
                ORDER BY created_at_ms DESC LIMIT ? OFFSET ?""",
            params + [page_size, offset],
        ).fetchall()
        return {"items": [dict(r) for r in rows], "page": page,
                "page_size": page_size, "total": total}
    finally:
        conn.close()


def get_job_detail(db_path, job_id):
    """Full job DTO including batch metadata (provider/model/usage)."""
    conn = read_connection(db_path)
    try:
        job = conn.execute(
            "SELECT * FROM ai_generation_jobs WHERE id=?", (job_id,)
        ).fetchone()
        if not job:
            return None
        result = dict(job)
        batch = conn.execute(
            "SELECT schema_version, provider, model, prompt_version, usage_json, "
            "created_at_ms FROM ai_generation_batches WHERE job_id=?",
            (job_id,),
        ).fetchone()
        if batch:
            result["batch_meta"] = dict(batch)
        else:
            result["batch_meta"] = None
        return result
    finally:
        conn.close()
