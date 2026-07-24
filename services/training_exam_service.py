"""Exam service: template, exam event, assignment, state transitions.

Template chỉ tham chiếu published question versions. State machine:
draft → ready → open → closed/cancelled. Không pause, không mở lại.
Open/close dùng compare-and-set + audit.
"""

import json
import sqlite3

from training import constants, time_policy
from training.db import read_connection, write_connection
from training.errors import ErrorCode, TrainingError
from repositories.training_repository import gen_id, write_audit

SHUFFLE_ALGORITHM_VERSION = "v1"
_AFTER_CLOSE_COMMIT_HOOK = None


def create_template(
    db_path, *, unit_code, actor, code, title, target_audience_code,
    question_version_ids, duration_seconds, pass_score_percent,
    shuffle_questions=False, shuffle_options=False,
):
    """Tạo fixed template từ published question versions."""
    if not question_version_ids:
        raise TrainingError(ErrorCode.TEMPLATE_QUESTION_INVALID, "Template cần ít nhất một câu hỏi")
    if len(set(question_version_ids)) != len(question_version_ids):
        raise TrainingError(ErrorCode.TEMPLATE_QUESTION_DUPLICATE, "Template không được chứa câu hỏi trùng")
    conn = write_connection(db_path)
    try:
        placeholders = ",".join("?" * len(question_version_ids))
        versions = conn.execute(
            f"SELECT id, publication_status FROM question_versions WHERE id IN ({placeholders})",
            question_version_ids,
        ).fetchall()
        if len(versions) != len(question_version_ids):
            raise TrainingError(
                ErrorCode.TEMPLATE_QUESTION_INVALID,
                "Question version không tồn tại",
                status=400,
            )
        unpublished = next((version for version in versions if version["publication_status"] != "published"), None)
        if unpublished:
            raise TrainingError(
                ErrorCode.TEMPLATE_QUESTION_INVALID,
                f"Question version {unpublished['id']} chưa được publish",
                status=400,
            )
        incompatible = conn.execute(
            f"""SELECT qv.id FROM question_versions qv
            WHERE qv.id IN ({placeholders})
              AND EXISTS (
                  SELECT 1 FROM question_audiences qa
                  WHERE qa.question_version_id = qv.id
              )
              AND NOT EXISTS (
                  SELECT 1 FROM question_audiences qa
                  WHERE qa.question_version_id = qv.id AND qa.audience_code = ?
              )""",
            [*question_version_ids, target_audience_code],
        ).fetchone()
        if incompatible:
            raise TrainingError(
                ErrorCode.TEMPLATE_QUESTION_INVALID,
                f"Question version {incompatible['id']} không phù hợp đối tượng",
                status=400,
            )
        template_id = gen_id("tpl")
        now = time_policy.utc_now_ms()
        conn.execute(
            """INSERT INTO exam_templates
            (id, code, title, target_audience_code, total_questions, duration_seconds,
             pass_score_percent, shuffle_questions, shuffle_options, locked, created_by, created_at_ms)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 0, ?, ?)""",
            (template_id, code, title, target_audience_code,
             0, duration_seconds, pass_score_percent,
             int(shuffle_questions), int(shuffle_options), actor, now),
        )
        for seq, qv_id in enumerate(question_version_ids, start=1):
            conn.execute(
                "INSERT INTO exam_template_items (id, template_id, sequence_number, question_version_id, points) "
                "VALUES (?, ?, ?, ?, 1.0)",
                (gen_id("ti"), template_id, seq, qv_id),
            )
        total_questions = conn.execute(
            "SELECT COUNT(*) FROM exam_template_items WHERE template_id=?", (template_id,)
        ).fetchone()[0]
        conn.execute(
            "UPDATE exam_templates SET total_questions=? WHERE id=?", (total_questions, template_id)
        )
        write_audit(conn, actor=actor, unit_code=unit_code, action="create_template",
                    entity_type="exam_template", entity_id=template_id)
        conn.commit()
        return {"id": template_id, "code": code, "title": title,
                "total_questions": total_questions, "locked": 0}
    finally:
        conn.close()


def get_template_items(db_path, template_id):
    conn = read_connection(db_path)
    try:
        rows = conn.execute(
            "SELECT * FROM exam_template_items WHERE template_id=? ORDER BY sequence_number",
            (template_id,),
        ).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def update_template(
    db_path, *, unit_code, actor, template_id, question_version_ids,
    shuffle_questions, shuffle_options,
):
    """Update mutable template presentation settings and its question list."""
    conn = write_connection(db_path)
    try:
        conn.execute("BEGIN IMMEDIATE")
        template = conn.execute(
            "SELECT locked, target_audience_code FROM exam_templates WHERE id=?", (template_id,)
        ).fetchone()
        if not template:
            raise TrainingError(ErrorCode.NOT_FOUND, "Template không tồn tại", status=404)
        used = conn.execute(
            "SELECT 1 FROM exam_events WHERE template_id=?", (template_id,)
        ).fetchone()
        if template["locked"] or used:
            raise TrainingError(
                ErrorCode.TEMPLATE_IMMUTABLE,
                "Template đã khóa hoặc đã được dùng cho kỳ thi", status=409,
            )
        if not question_version_ids:
            raise TrainingError(ErrorCode.TEMPLATE_QUESTION_INVALID, "Template cần ít nhất một câu hỏi")
        if len(set(question_version_ids)) != len(question_version_ids):
            raise TrainingError(ErrorCode.TEMPLATE_QUESTION_DUPLICATE, "Template không được chứa câu hỏi trùng")
        placeholders = ",".join("?" * len(question_version_ids))
        versions = conn.execute(
            f"SELECT id, publication_status FROM question_versions WHERE id IN ({placeholders})",
            question_version_ids,
        ).fetchall()
        if len(versions) != len(question_version_ids) or any(
            version["publication_status"] != "published" for version in versions
        ):
            raise TrainingError(ErrorCode.TEMPLATE_QUESTION_INVALID, "Question version không hợp lệ")
        incompatible = conn.execute(
            f"""SELECT qv.id FROM question_versions qv
            WHERE qv.id IN ({placeholders})
              AND EXISTS (
                  SELECT 1 FROM question_audiences qa
                  WHERE qa.question_version_id = qv.id
              )
              AND NOT EXISTS (
                  SELECT 1 FROM question_audiences qa
                  WHERE qa.question_version_id = qv.id AND qa.audience_code = ?
              )""",
            [*question_version_ids, template["target_audience_code"]],
        ).fetchone()
        if incompatible:
            raise TrainingError(
                ErrorCode.TEMPLATE_QUESTION_INVALID,
                f"Question version {incompatible['id']} không phù hợp đối tượng",
                status=400,
            )
        conn.execute("DELETE FROM exam_template_items WHERE template_id=?", (template_id,))
        for sequence_number, question_version_id in enumerate(question_version_ids, start=1):
            conn.execute(
                """INSERT INTO exam_template_items
                (id, template_id, sequence_number, question_version_id, points)
                VALUES (?, ?, ?, ?, 1.0)""",
                (gen_id("ti"), template_id, sequence_number, question_version_id),
            )
        conn.execute(
            """UPDATE exam_templates
            SET total_questions=?, shuffle_questions=?, shuffle_options=? WHERE id=?""",
            (len(question_version_ids), int(shuffle_questions), int(shuffle_options), template_id),
        )
        write_audit(conn, actor=actor, unit_code=unit_code, action="update_template",
                    entity_type="exam_template", entity_id=template_id)
        conn.commit()
    except TrainingError:
        if conn.in_transaction:
            conn.rollback()
        raise
    finally:
        conn.close()


def create_exam(
    db_path, *, unit_code, actor, code, title, template_id, target_audience_code,
    start_at_ms, end_at_ms, duration_seconds, pass_score_percent,
    description=None, reveal_answers_after_finalize=True,
):
    if end_at_ms <= start_at_ms:
        raise TrainingError(ErrorCode.VALIDATION_ERROR, "end_at phải sau start_at")
    conn = write_connection(db_path)
    try:
        template = conn.execute(
            "SELECT target_audience_code FROM exam_templates WHERE id=?", (template_id,)
        ).fetchone()
        if not template:
            raise TrainingError(ErrorCode.NOT_FOUND, "Template không tồn tại", status=404)
        if template["target_audience_code"] != target_audience_code:
            raise TrainingError(
                ErrorCode.AUDIENCE_MISMATCH,
                "Đối tượng kỳ thi phải trùng với đối tượng template",
            )
        exam_id = gen_id("exam")
        now = time_policy.utc_now_ms()
        conn.execute(
            """INSERT INTO exam_events
            (id, code, title, description, template_id, target_audience_code, status,
             start_at_ms, end_at_ms, duration_seconds, pass_score_percent,
             reveal_answers_after_finalize, created_by, created_at_ms)
            VALUES (?, ?, ?, ?, ?, ?, 'draft', ?, ?, ?, ?, ?, ?, ?)""",
            (exam_id, code, title, description, template_id, target_audience_code,
             start_at_ms, end_at_ms, duration_seconds, pass_score_percent,
             int(reveal_answers_after_finalize), actor, now),
        )
        write_audit(conn, actor=actor, unit_code=unit_code, action="create_exam",
                    entity_type="exam_event", entity_id=exam_id)
        conn.commit()
        return {"id": exam_id, "code": code, "status": constants.ExamStatus.DRAFT}
    finally:
        conn.close()


def _compare_and_set_exam_status(conn, exam_id, from_status, to_status):
    cursor = conn.execute(
        "UPDATE exam_events SET status=? WHERE id=? AND status=?",
        (to_status, exam_id, from_status),
    )
    return cursor.rowcount > 0


def ready_exam(db_path, *, unit_code, actor, exam_id):
    conn = write_connection(db_path)
    try:
        conn.execute("BEGIN IMMEDIATE")
        ok = _compare_and_set_exam_status(conn, exam_id, constants.ExamStatus.DRAFT, constants.ExamStatus.READY)
        if not ok:
            exam = conn.execute("SELECT status FROM exam_events WHERE id=?", (exam_id,)).fetchone()
            if not exam:
                raise TrainingError(ErrorCode.NOT_FOUND, "Kỳ thi không tồn tại", status=404)
            raise TrainingError(ErrorCode.CONFLICT,
                                f"Không thể ready: trạng thái hiện tại {exam['status']}", status=409)
        write_audit(conn, actor=actor, unit_code=unit_code, action="ready_exam",
                    entity_type="exam_event", entity_id=exam_id)
        conn.commit()
    except TrainingError:
        if conn.in_transaction:
            conn.rollback()
        raise
    finally:
        conn.close()


def open_exam(db_path, *, unit_code, actor, exam_id):
    conn = write_connection(db_path)
    try:
        conn.execute("BEGIN IMMEDIATE")
        exam = conn.execute(
            "SELECT status, start_at_ms, end_at_ms FROM exam_events WHERE id=?",
            (exam_id,),
        ).fetchone()
        if not exam:
            raise TrainingError(ErrorCode.NOT_FOUND, "Kỳ thi không tồn tại", status=404)
        now = time_policy.utc_now_ms()
        if exam["status"] == constants.ExamStatus.OPEN:
            return
        if exam["status"] != constants.ExamStatus.READY:
            raise TrainingError(ErrorCode.CONFLICT,
                                f"Kỳ thi không thể mở: trạng thái {exam['status']}", status=409)
        if now < exam["start_at_ms"] or now >= exam["end_at_ms"]:
            raise TrainingError(ErrorCode.EXAM_NOT_OPEN,
                                "Chưa đến hoặc đã hết thời gian thi", status=409)
        if not _compare_and_set_exam_status(conn, exam_id, exam["status"], constants.ExamStatus.OPEN):
            raise TrainingError(ErrorCode.CONFLICT, "Kỳ thi đã thay đổi trạng thái", status=409)
        write_audit(conn, actor=actor, unit_code=unit_code, action="open_exam",
                    entity_type="exam_event", entity_id=exam_id)
        conn.commit()
    except TrainingError:
        if conn.in_transaction:
            conn.rollback()
        raise
    finally:
        conn.close()


def close_exam(db_path, *, unit_code, actor, exam_id):
    conn = write_connection(db_path)
    try:
        conn.execute("BEGIN IMMEDIATE")
        now = time_policy.utc_now_ms()
        ok = conn.execute(
            "UPDATE exam_events SET status=?, closed_at_ms=? WHERE id=? AND status=?",
            (constants.ExamStatus.CLOSED, now, exam_id, constants.ExamStatus.OPEN),
        ).rowcount > 0
        if not ok:
            exam = conn.execute("SELECT status FROM exam_events WHERE id=?", (exam_id,)).fetchone()
            if not exam:
                raise TrainingError(ErrorCode.NOT_FOUND, "Kỳ thi không tồn tại", status=404)
            if exam["status"] == constants.ExamStatus.CLOSED:
                conn.rollback()
            else:
                raise TrainingError(ErrorCode.CONFLICT,
                                    f"Không thể đóng: trạng thái {exam['status']}", status=409)
        else:
            write_audit(conn, actor=actor, unit_code=unit_code, action="close_exam",
                        entity_type="exam_event", entity_id=exam_id)
            conn.commit()
    finally:
        conn.close()
    if _AFTER_CLOSE_COMMIT_HOOK:
        _AFTER_CLOSE_COMMIT_HOOK(exam_id)
    return _administratively_finish_active_attempts(
        db_path, unit_code=unit_code, actor=actor, exam_id=exam_id,
    )


def _administratively_finish_active_attempts(db_path, *, unit_code, actor, exam_id, ended_reason="exam_closed"):
    """Recover attempts after close; each attempt uses its own transaction."""
    from services import training_attempt_service as attempts

    conn = read_connection(db_path)
    try:
        attempt_rows = conn.execute(
            """SELECT a.id FROM exam_attempts a JOIN exam_assignments x ON x.id=a.assignment_id
               WHERE x.exam_event_id=? ORDER BY a.created_at_ms, a.id""", (exam_id,)
        ).fetchall()
    finally:
        conn.close()
    summary = {
        "processed_attempt_ids": [],
        "already_completed_ids": [],
        "failed_attempts": [],
    }
    for row in attempt_rows:
        attempt_id = row["id"]
        attempt = attempts.get_attempt(db_path, attempt_id)
        if attempt["status"] != constants.AttemptStatus.ACTIVE:
            summary["already_completed_ids"].append(attempt_id)
            continue
        try:
            outcome = attempts.administratively_submit_attempt(
                db_path, unit_code=unit_code, actor=actor, attempt_id=attempt_id,
                ended_reason=ended_reason, return_outcome=True,
            )
            if outcome["transitioned"]:
                summary["processed_attempt_ids"].append(attempt_id)
            else:
                summary["already_completed_ids"].append(attempt_id)
        except TrainingError as exc:
            summary["failed_attempts"].append({
                "attempt_id": attempt_id,
                "error_code": exc.code,
            })
    return summary


def cancel_exam(db_path, *, unit_code, actor, exam_id):
    conn = write_connection(db_path)
    try:
        conn.execute("BEGIN IMMEDIATE")
        exam = conn.execute("SELECT status FROM exam_events WHERE id=?", (exam_id,)).fetchone()
        if not exam:
            raise TrainingError(ErrorCode.NOT_FOUND, "Kỳ thi không tồn tại", status=404)
        if exam["status"] not in (constants.ExamStatus.DRAFT, constants.ExamStatus.READY):
            raise TrainingError(ErrorCode.CONFLICT,
                                f"Không thể hủy: trạng thái {exam['status']}", status=409)
        if not _compare_and_set_exam_status(
            conn, exam_id, exam["status"], constants.ExamStatus.CANCELLED,
        ):
            raise TrainingError(ErrorCode.CONFLICT, "Kỳ thi đã thay đổi trạng thái", status=409)
        write_audit(conn, actor=actor, unit_code=unit_code, action="cancel_exam",
                    entity_type="exam_event", entity_id=exam_id)
        conn.commit()
    except TrainingError:
        if conn.in_transaction:
            conn.rollback()
        raise
    finally:
        conn.close()


def create_assignments(db_path, *, unit_code, actor, exam_id, users, audience_code):
    """Snapshot users vào assignment."""
    now = time_policy.utc_now_ms()
    conn = write_connection(db_path)
    try:
        conn.execute("BEGIN IMMEDIATE")
        exam = conn.execute(
            "SELECT duration_seconds, target_audience_code, status, finalized_at_ms FROM exam_events WHERE id=?",
            (exam_id,),
        ).fetchone()
        if not exam:
            raise TrainingError(ErrorCode.NOT_FOUND, "Kỳ thi không tồn tại", status=404)
        if exam["status"] not in (constants.ExamStatus.DRAFT, constants.ExamStatus.READY) or exam["finalized_at_ms"] is not None:
            raise TrainingError(
                ErrorCode.CONFLICT,
                f"Không thể giao bài: trạng thái hiện tại {exam['status']}",
                status=409,
            )
        if exam["target_audience_code"] != audience_code:
            raise TrainingError(
                ErrorCode.AUDIENCE_MISMATCH,
                "Đối tượng giao bài phải trùng với đối tượng kỳ thi",
            )
        usernames = [user["username"] for user in users]
        if any(not isinstance(username, str) or not username.strip() for username in usernames):
            raise TrainingError(
                ErrorCode.VALIDATION_ERROR,
                "Username giao bài là bắt buộc",
            )
        if len(set(usernames)) != len(usernames):
            raise TrainingError(
                ErrorCode.ASSIGNMENT_ALREADY_EXISTS,
                "Không được giao trùng người dùng trong cùng yêu cầu",
                status=409,
            )
        if usernames:
            placeholders = ",".join("?" * len(usernames))
            existing_assignment = conn.execute(
                f"""SELECT username FROM exam_assignments
                WHERE exam_event_id=? AND audience_code=? AND username IN ({placeholders})
                LIMIT 1""",
                [exam_id, audience_code, *usernames],
            ).fetchone()
            if existing_assignment:
                raise TrainingError(
                    ErrorCode.ASSIGNMENT_ALREADY_EXISTS,
                    f"Người dùng {existing_assignment['username']} đã được giao bài",
                    status=409,
                )
            mismatched_user = conn.execute(
                f"""SELECT configured.username
                FROM (
                    SELECT DISTINCT username FROM training_user_audiences
                    WHERE username IN ({placeholders})
                ) configured
                WHERE NOT EXISTS (
                    SELECT 1 FROM training_user_audiences audience
                    WHERE audience.username=configured.username AND audience.audience_code=?
                )
                LIMIT 1""",
                [*usernames, audience_code],
            ).fetchone()
            if mismatched_user:
                raise TrainingError(
                    ErrorCode.AUDIENCE_MISMATCH,
                    f"Người dùng {mismatched_user['username']} không thuộc đối tượng được giao",
                )
        assignment_ids = []
        for user in users:
            assignment_id = gen_id("asg")
            conn.execute(
                """INSERT INTO exam_assignments
                (id, exam_event_id, username, display_name, team_code, team_name,
                 organization_code, organization_name, audience_code, status,
                 duration_seconds, assigned_at_ms)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'assigned', ?, ?)""",
                (assignment_id, exam_id, user["username"],
                 user.get("display_name", user["username"]),
                 user.get("team_code"), user.get("team_name"),
                 user.get("organization_code"), user.get("organization_name"),
                 audience_code, exam["duration_seconds"], now),
            )
            assignment_ids.append(assignment_id)
        write_audit(conn, actor=actor, unit_code=unit_code, action="create_assignments",
                    entity_type="exam_event", entity_id=exam_id,
                    after={"count": len(assignment_ids)})
        conn.commit()
        return assignment_ids
    except sqlite3.IntegrityError as exc:
        if conn.in_transaction:
            conn.rollback()
        if "UNIQUE constraint failed: exam_assignments.exam_event_id, exam_assignments.username, exam_assignments.audience_code" in str(exc):
            raise TrainingError(
                ErrorCode.ASSIGNMENT_ALREADY_EXISTS,
                "Người dùng đã được giao bài",
                status=409,
            ) from exc
        raise
    except TrainingError:
        if conn.in_transaction:
            conn.rollback()
        raise
    finally:
        conn.close()


def get_exam(db_path, exam_id):
    conn = read_connection(db_path)
    try:
        row = conn.execute("SELECT * FROM exam_events WHERE id=?", (exam_id,)).fetchone()
        return dict(row) if row else None
    finally:
        conn.close()


def get_assignment(db_path, assignment_id):
    conn = read_connection(db_path)
    try:
        row = conn.execute("SELECT * FROM exam_assignments WHERE id=?", (assignment_id,)).fetchone()
        return dict(row) if row else None
    finally:
        conn.close()


def get_assignments_for_exam(db_path, exam_id):
    conn = read_connection(db_path)
    try:
        rows = conn.execute(
            "SELECT * FROM exam_assignments WHERE exam_event_id=? ORDER BY username",
            (exam_id,),
        ).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def get_assignment_for_user(db_path, exam_id, username):
    conn = read_connection(db_path)
    try:
        rows = conn.execute(
            "SELECT * FROM exam_assignments WHERE exam_event_id=? AND username=? ORDER BY assigned_at_ms",
            (exam_id, username),
        ).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


_TEMPLATE_COLUMNS = (
    "id, code, title, target_audience_code, total_questions, duration_seconds, "
    "pass_score_percent, shuffle_questions, shuffle_options, locked, "
    "created_by, created_at_ms"
)


def _template_dto(row):
    return {
        "id": row["id"], "code": row["code"], "title": row["title"],
        "target_audience_code": row["target_audience_code"],
        "total_questions": row["total_questions"],
        "duration_seconds": row["duration_seconds"],
        "pass_score_percent": row["pass_score_percent"],
        "shuffle_questions": bool(row["shuffle_questions"]),
        "shuffle_options": bool(row["shuffle_options"]),
        "locked": bool(row["locked"]),
        "created_by": row["created_by"],
        "created_at_ms": row["created_at_ms"],
    }


def list_templates(db_path, *, page=1, page_size=25):
    """Return paginated template DTOs with explicit allowlist columns."""
    conn = read_connection(db_path)
    try:
        total = conn.execute(
            "SELECT COUNT(*) AS c FROM exam_templates"
        ).fetchone()["c"]
        offset = (page - 1) * page_size
        rows = conn.execute(
            f"""SELECT {_TEMPLATE_COLUMNS} FROM exam_templates
                ORDER BY created_at_ms DESC LIMIT ? OFFSET ?""",
            (page_size, offset),
        ).fetchall()
        return {
            "items": [_template_dto(row) for row in rows],
            "page": page, "page_size": page_size, "total": total,
        }
    finally:
        conn.close()


def get_template_detail(db_path, template_id):
    """Return template DTO + ordered items joined with question metadata."""
    conn = read_connection(db_path)
    try:
        row = conn.execute(
            f"SELECT {_TEMPLATE_COLUMNS} FROM exam_templates WHERE id=?",
            (template_id,),
        ).fetchone()
        if not row:
            raise TrainingError(ErrorCode.NOT_FOUND, "Mẫu đề không tồn tại.", status=404)
        item_rows = conn.execute(
            """SELECT ti.sequence_number, ti.question_version_id, qv.stem, qv.type,
                      qv.difficulty, ti.section_label, ti.points
               FROM exam_template_items ti
               JOIN question_versions qv ON qv.id = ti.question_version_id
               WHERE ti.template_id = ?
               ORDER BY ti.sequence_number""",
            (template_id,),
        ).fetchall()
        items = [{
            "sequence_number": r["sequence_number"],
            "question_version_id": r["question_version_id"],
            "stem": r["stem"], "type": r["type"], "difficulty": r["difficulty"],
            "section_label": r["section_label"], "points": r["points"],
        } for r in item_rows]
        detail = _template_dto(row)
        detail["items"] = items
        return detail
    finally:
        conn.close()


_EXAM_COLUMNS = (
    "id, code, title, template_id, target_audience_code, status, "
    "start_at_ms, end_at_ms, duration_seconds, pass_score_percent, "
    "reveal_answers_after_finalize, created_by, created_at_ms, finalized_at_ms"
)


def _exam_dto(row):
    return {
        "id": row["id"], "code": row["code"], "title": row["title"],
        "template_id": row["template_id"],
        "target_audience_code": row["target_audience_code"],
        "status": row["status"], "start_at_ms": row["start_at_ms"],
        "end_at_ms": row["end_at_ms"], "duration_seconds": row["duration_seconds"],
        "pass_score_percent": row["pass_score_percent"],
        "reveal_answers_after_finalize": bool(row["reveal_answers_after_finalize"]),
        "created_by": row["created_by"], "created_at_ms": row["created_at_ms"],
        "finalized_at_ms": row["finalized_at_ms"],
    }


def list_exams(db_path, *, page=1, page_size=25, status=None):
    """Return paginated exam DTOs; optional status filter."""
    conn = read_connection(db_path)
    try:
        where = []
        params = []
        if status:
            where.append("status = ?")
            params.append(status)
        clause = ("WHERE " + " AND ".join(where)) if where else ""
        total = conn.execute(
            f"SELECT COUNT(*) AS c FROM exam_events {clause}", params
        ).fetchone()["c"]
        offset = (page - 1) * page_size
        rows = conn.execute(
            f"""SELECT {_EXAM_COLUMNS} FROM exam_events {clause}
                ORDER BY created_at_ms DESC LIMIT ? OFFSET ?""",
            params + [page_size, offset],
        ).fetchall()
        return {
            "items": [_exam_dto(row) for row in rows],
            "page": page, "page_size": page_size, "total": total,
        }
    finally:
        conn.close()


def _assignment_summary(conn, exam_id):
    def _count(status_value):
        return conn.execute(
            "SELECT COUNT(*) AS c FROM exam_assignments "
            "WHERE exam_event_id=? AND status=?",
            (exam_id, status_value),
        ).fetchone()["c"]

    total = conn.execute(
        "SELECT COUNT(*) AS c FROM exam_assignments WHERE exam_event_id=?", (exam_id,)
    ).fetchone()["c"]
    in_progress = conn.execute(
        """SELECT COUNT(*) AS c FROM exam_assignments a
           JOIN exam_attempts att ON att.assignment_id = a.id
           WHERE a.exam_event_id=? AND att.status IN ('created', 'active')""",
        (exam_id,),
    ).fetchone()["c"]
    return {
        "total": total, "assigned": _count("assigned"),
        "completed": _count("completed"), "expired": _count("expired"),
        "cancelled": _count("cancelled"), "in_progress": in_progress,
    }


def get_exam_detail(db_path, exam_id):
    """Return exam DTO + nested template + assignment_summary."""
    conn = read_connection(db_path)
    try:
        row = conn.execute(
            f"SELECT {_EXAM_COLUMNS} FROM exam_events WHERE id=?", (exam_id,),
        ).fetchone()
        if not row:
            raise TrainingError(ErrorCode.NOT_FOUND, "Kỳ thi không tồn tại.", status=404)
        template_row = conn.execute(
            "SELECT id, code, title FROM exam_templates WHERE id=?", (row["template_id"],),
        ).fetchone()
        template = (
            {"id": template_row["id"], "code": template_row["code"], "title": template_row["title"]}
            if template_row else None
        )
        detail = _exam_dto(row)
        detail["template"] = template
        detail["assignment_summary"] = _assignment_summary(conn, exam_id)
        return detail
    finally:
        conn.close()


def get_exam_assignments_dto(db_path, exam_id):
    """Return stripped assignment DTOs (never exposes internal/technical columns)."""
    conn = read_connection(db_path)
    try:
        rows = conn.execute(
            """SELECT a.id, a.username, a.display_name, a.team_code, a.team_name,
                      a.organization_code, a.organization_name, a.audience_code,
                      a.status, a.assigned_at_ms, att.status AS attempt_status
               FROM exam_assignments a
               LEFT JOIN exam_attempts att ON att.assignment_id = a.id
               WHERE a.exam_event_id = ?
               ORDER BY a.username, a.assigned_at_ms""",
            (exam_id,),
        ).fetchall()
        return [{
            "id": r["id"], "username": r["username"], "display_name": r["display_name"],
            "team_code": r["team_code"], "team_name": r["team_name"],
            "organization_code": r["organization_code"],
            "organization_name": r["organization_name"],
            "audience_code": r["audience_code"], "status": r["status"],
            "assigned_at_ms": r["assigned_at_ms"],
            "attempt_status": r["attempt_status"],
        } for r in rows]
    finally:
        conn.close()


def get_my_assignments_dto(db_path, username):
    conn = read_connection(db_path)
    try:
        rows = conn.execute(
            """
            SELECT a.id, a.exam_event_id, a.username, a.display_name,
                   a.audience_code, a.status AS assignment_status,
                   a.duration_seconds, a.assigned_at_ms,
                   e.code AS exam_code, e.title AS exam_title,
                   e.status AS exam_status, e.start_at_ms, e.end_at_ms,
                   e.pass_score_percent, e.reveal_answers_after_finalize,
                   e.finalized_at_ms,
                   t.id AS attempt_id, t.status AS attempt_status,
                   t.deadline_at_ms
            FROM exam_assignments a
            JOIN exam_events e ON e.id = a.exam_event_id
            LEFT JOIN exam_attempts t ON t.assignment_id = a.id
            WHERE a.username = ?
            ORDER BY e.start_at_ms DESC, a.assigned_at_ms DESC
            """,
            (username,),
        ).fetchall()
    finally:
        conn.close()

    items = []
    for r in rows:
        items.append({
            "assignment_id": r["id"],
            "exam_id": r["exam_event_id"],
            "exam_code": r["exam_code"],
            "exam_title": r["exam_title"],
            "exam_status": r["exam_status"],
            "assignment_status": r["assignment_status"],
            "audience_code": r["audience_code"],
            "start_at_ms": r["start_at_ms"],
            "end_at_ms": r["end_at_ms"],
            "duration_seconds": r["duration_seconds"],
            "pass_score_percent": r["pass_score_percent"],
            "reveal_answers_after_finalize": bool(r["reveal_answers_after_finalize"]),
            "finalized": r["finalized_at_ms"] is not None,
            "assigned_at_ms": r["assigned_at_ms"],
            "attempt_id": r["attempt_id"],
            "attempt_status": r["attempt_status"],
            "deadline_at_ms": r["deadline_at_ms"],
        })
    return {"items": items}


def _get_all_users():
    """Indirection over auth.get_all_users for testability."""
    from auth import get_all_users
    return get_all_users()


def list_assignable_users(db_path, *, q=""):
    """Return stripped user DTOs ({username, display_name}); never exposes password.

    Phải tách biệt rõ: dữ liệu user từ auth (Excel) chỉ trả ra trường công khai.
    """
    users = _get_all_users()
    if users is None:
        return []
    try:
        records = users.to_dict("records")
    except AttributeError:
        records = list(users)
    needle = (q or "").lower()
    items = []
    for user in records:
        if not isinstance(user, dict):
            continue
        username = user.get("username")
        if not username:
            continue
        display_name = user.get("display_name") or user.get("name") or username
        is_active = user.get("is_active")
        if is_active in (0, "0", False, None):
            continue
        if needle and needle not in str(username).lower() \
                and needle not in str(display_name).lower():
            continue
        items.append({"username": username, "display_name": display_name})
    return items
