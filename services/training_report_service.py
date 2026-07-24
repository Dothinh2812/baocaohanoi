"""Finalize và report snapshot bất biến cho kỳ thi."""

import hashlib
import json

from training import constants, time_policy
from training.db import read_connection, write_connection
from repositories.training_repository import gen_id, write_audit
from services import training_attempt_service as attempts
from services import training_exam_service as exams

REPORT_SCHEMA_VERSION = "1.0"
REPORT_ALGORITHM_VERSION = "v1"


def _checksum(payload):
    raw = json.dumps(payload, ensure_ascii=False, sort_keys=True)
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


def _build_payload(db_path, exam_id):
    conn = read_connection(db_path)
    try:
        assignments = [dict(row) for row in conn.execute(
            "SELECT * FROM exam_assignments WHERE exam_event_id=? ORDER BY username", (exam_id,)
        ).fetchall()]
        individual = []
        for assignment in assignments:
            result = conn.execute(
                """SELECT r.raw_score, r.maximum_score, r.percent, r.passed, r.topic_breakdown_json
                   FROM exam_results r
                   JOIN exam_attempts a ON a.id=r.attempt_id
                   WHERE a.assignment_id=?""",
                (assignment["id"],),
            ).fetchone()
            individual.append({
                "assignment": assignment,
                "result": ({**dict(result), "topic_breakdown": json.loads(result["topic_breakdown_json"])}
                           if result else None),
            })
        completed = [row for row in individual if row["result"]]
        return {
            "schema_version": REPORT_SCHEMA_VERSION,
            "exam_id": exam_id,
            "summary": {
                "assigned": len(assignments),
                "completed": len(completed),
                "expired": sum(a["status"] == "expired" for a in assignments),
                "passed": sum(bool(row["result"]["passed"]) for row in completed),
                "failed": sum(not bool(row["result"]["passed"]) for row in completed),
            },
            "individual": individual,
        }
    finally:
        conn.close()


def finalize_exam(db_path, *, unit_code, actor, exam_id):
    """Chốt kỳ thi idempotently; không giữ transaction qua scoring/report build."""
    exam = exams.get_exam(db_path, exam_id)
    if not exam:
        raise ValueError("Kỳ thi không tồn tại")
    if exam["status"] == constants.ExamStatus.OPEN:
        exams.close_exam(db_path, unit_code=unit_code, actor=actor, exam_id=exam_id)

    conn = read_connection(db_path)
    try:
        active_attempt_ids = [row["id"] for row in conn.execute(
            """SELECT a.id FROM exam_attempts a
               JOIN exam_assignments x ON x.id=a.assignment_id
               WHERE x.exam_event_id=? AND a.status='active'""",
            (exam_id,),
        ).fetchall()]
    finally:
        conn.close()

    for attempt_id in active_attempt_ids:
        attempts.submit_attempt(db_path, unit_code=unit_code, actor=actor, attempt_id=attempt_id)

    conn = write_connection(db_path)
    try:
        conn.execute("BEGIN IMMEDIATE")
        conn.execute(
            """UPDATE exam_assignments SET status='expired'
               WHERE exam_event_id=? AND status='assigned'
                 AND id NOT IN (SELECT assignment_id FROM exam_attempts)""",
            (exam_id,),
        )
        conn.commit()
    finally:
        conn.close()

    payload = _build_payload(db_path, exam_id)
    conn = write_connection(db_path)
    try:
        conn.execute("BEGIN IMMEDIATE")
        existing = conn.execute(
            "SELECT * FROM exam_report_snapshots WHERE exam_event_id=? ORDER BY revision DESC LIMIT 1",
            (exam_id,),
        ).fetchone()
        if existing:
            conn.rollback()
            return {"revision": existing["revision"], "payload": json.loads(existing["payload_json"])}
        now = time_policy.utc_now_ms()
        conn.execute(
            """INSERT INTO exam_report_snapshots
               (id, exam_event_id, revision, schema_version, report_algorithm_version,
                payload_json, checksum, previous_revision, created_by, created_at_ms)
               VALUES (?, ?, 1, ?, ?, ?, ?, NULL, ?, ?)""",
            (gen_id("report"), exam_id, REPORT_SCHEMA_VERSION, REPORT_ALGORITHM_VERSION,
             json.dumps(payload, ensure_ascii=False), _checksum(payload), actor, now),
        )
        conn.execute(
            "UPDATE exam_events SET status='closed', finalized_at_ms=?, finalized_by=? WHERE id=?",
            (now, actor, exam_id),
        )
        write_audit(conn, actor=actor, unit_code=unit_code, action="finalize_exam",
                    entity_type="exam_event", entity_id=exam_id)
        conn.commit()
        return {"revision": 1, "payload": payload}
    finally:
        conn.close()


def get_report_snapshot(db_path, exam_id, revision=None):
    conn = read_connection(db_path)
    try:
        if revision is None:
            row = conn.execute(
                "SELECT revision, payload_json FROM exam_report_snapshots "
                "WHERE exam_event_id=? ORDER BY revision DESC LIMIT 1",
                (exam_id,),
            ).fetchone()
        else:
            row = conn.execute(
                "SELECT revision, payload_json FROM exam_report_snapshots "
                "WHERE exam_event_id=? AND revision=?",
                (exam_id, revision),
            ).fetchone()
        if not row:
            return None
        return {"revision": row["revision"], "payload": json.loads(row["payload_json"])}
    finally:
        conn.close()
