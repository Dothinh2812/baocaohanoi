"""Attempt service: start, autosave, submit, scoring.

Start dùng BEGIN IMMEDIATE; unique (assignment_id) là lớp bảo vệ race cuối cùng.
Submit dùng compare-and-set. Chấm chỉ từ snapshot.
"""

import hashlib
import json
import random
import secrets

from training import constants, time_policy
from training.db import read_connection, write_connection
from training.errors import ErrorCode, TrainingError
from repositories.training_repository import gen_id, write_audit
from services import training_exam_service as es
from services import training_scoring_service as scoring

SHUFFLE_ALGORITHM_VERSION = "shuffle_v1"


def _snapshot_checksum(items_data):
    raw = json.dumps(items_data, ensure_ascii=False, sort_keys=True)
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


def start_attempt(db_path, *, unit_code, actor, assignment_id):
    """Bắt đầu attempt: BEGIN IMMEDIATE, check exam/time/ownership, snapshot, commit.

    Idempotent: nếu đã có attempt active, trả lại.
    """
    conn = write_connection(db_path)
    try:
        conn.execute("BEGIN IMMEDIATE")
        assignment = conn.execute(
            "SELECT * FROM exam_assignments WHERE id=?", (assignment_id,)
        ).fetchone()
        if not assignment:
            conn.execute("ROLLBACK")
            raise TrainingError(ErrorCode.ASSIGNMENT_NOT_FOUND,
                                "Không tìm thấy lượt được giao", status=404)
        if assignment["username"] != actor:
            conn.execute("ROLLBACK")
            raise TrainingError(ErrorCode.PERMISSION_SCOPE_DENIED,
                                "Bạn không sở hữu lượt được giao này", status=403)

        existing = conn.execute(
            "SELECT * FROM exam_attempts WHERE assignment_id=?", (assignment_id,)
        ).fetchone()
        if existing:
            exam_window = conn.execute(
                "SELECT status, start_at_ms, end_at_ms FROM exam_events WHERE id=?",
                (assignment["exam_event_id"],),
            ).fetchone()
            now = time_policy.utc_now_ms()
            if (
                exam_window["status"] != constants.ExamStatus.OPEN
                or now < exam_window["start_at_ms"]
                or now >= exam_window["end_at_ms"]
            ):
                conn.execute("ROLLBACK")
                raise TrainingError(ErrorCode.EXAM_NOT_OPEN,
                                    "Kỳ thi chưa mở hoặc đã đóng", status=409)
            conn.execute("ROLLBACK")
            if existing["status"] == constants.AttemptStatus.CREATED:
                _activate_attempt(conn, existing, assignment)
                conn.commit()
                return _attempt_start_result(existing["id"], conn, assignment_id)
            return _attempt_start_result(existing["id"], conn, assignment_id)

        exam = conn.execute(
            """SELECT ee.*, et.shuffle_questions, et.shuffle_options
            FROM exam_events ee
            JOIN exam_templates et ON et.id=ee.template_id
            WHERE ee.id=?""", (assignment["exam_event_id"],)
        ).fetchone()
        now = time_policy.utc_now_ms()
        if exam["status"] != constants.ExamStatus.OPEN:
            conn.execute("ROLLBACK")
            raise TrainingError(ErrorCode.EXAM_NOT_OPEN, "Kỳ thi chưa mở hoặc đã đóng", status=409)
        if now < exam["start_at_ms"] or now >= exam["end_at_ms"]:
            conn.execute("ROLLBACK")
            raise TrainingError(ErrorCode.EXAM_NOT_OPEN, "Chưa đến hoặc đã hết thời gian thi", status=409)

        attempt_id = gen_id("att")
        random_seed = secrets.randbits(31)
        now = time_policy.utc_now_ms()
        deadline = time_policy.compute_attempt_deadline_ms(
            now, exam["end_at_ms"], assignment["duration_seconds"] * 1000,
        )
        conn.execute(
            """INSERT INTO exam_attempts
            (id, assignment_id, status, random_seed, shuffle_algorithm_version,
             snapshot_checksum, started_at_ms, deadline_at_ms, created_at_ms)
            VALUES (?, ?, 'active', ?, ?, '', ?, ?, ?)""",
            (attempt_id, assignment_id, random_seed, SHUFFLE_ALGORITHM_VERSION,
             now, deadline, now),
        )
        _build_snapshot(conn, attempt_id, exam, random_seed, SHUFFLE_ALGORITHM_VERSION)
        write_audit(conn, actor=actor, unit_code=unit_code, action="start_attempt",
                    entity_type="exam_attempt", entity_id=attempt_id)
        conn.commit()
        return _attempt_start_result(attempt_id, conn, assignment_id)
    except TrainingError:
        raise
    except Exception:
        try:
            conn.execute("ROLLBACK")
        except Exception:
            pass
        raise
    finally:
        conn.close()


def _activate_attempt(conn, attempt, assignment):
    now = time_policy.utc_now_ms()
    exam = conn.execute(
        "SELECT end_at_ms FROM exam_events WHERE id=?", (assignment["exam_event_id"],)
    ).fetchone()
    deadline = time_policy.compute_attempt_deadline_ms(
        attempt["started_at_ms"] or now, exam["end_at_ms"],
        assignment["duration_seconds"] * 1000,
    )
    conn.execute(
        "UPDATE exam_attempts SET status='active', started_at_ms=COALESCE(started_at_ms,?), deadline_at_ms=? WHERE id=?",
        (now, deadline, attempt["id"]),
    )


def _ordered_snapshot_inputs(conn, template_id, random_seed, algorithm_version,
                             shuffle_questions, shuffle_options):
    """Đọc template và áp dụng thứ tự snapshot ổn định theo attempt."""
    items = list(conn.execute(
        """SELECT ti.sequence_number, ti.points, ti.section_label, qv.*
        FROM exam_template_items ti
        JOIN question_versions qv ON ti.question_version_id = qv.id
        WHERE ti.template_id=? ORDER BY ti.sequence_number""",
        (template_id,),
    ).fetchall())
    if shuffle_questions:
        _shuffle_snapshot_values(items, _snapshot_rng(random_seed, algorithm_version, "questions"))

    ordered_items = []
    for item in items:
        options = [dict(row) for row in conn.execute(
            "SELECT option_code, option_text, display_order FROM question_options "
            "WHERE question_version_id=? ORDER BY display_order",
            (item["id"],),
        ).fetchall()]
        if shuffle_options:
            _shuffle_snapshot_values(
                options, _snapshot_rng(random_seed, algorithm_version, f"options:{item['id']}"),
            )
        ordered_items.append((item, options))
    return ordered_items


def _build_snapshot(conn, attempt_id, exam, random_seed, algorithm_version):
    """Đọc template items, tạo snapshot items + options."""
    ordered_items = _ordered_snapshot_inputs(
        conn, exam["template_id"], random_seed, algorithm_version,
        exam["shuffle_questions"], exam["shuffle_options"],
    )
    for sequence_number, (item, opts) in enumerate(ordered_items, start=1):
        item_id = gen_id("atti")
        topics = [r["topic_code"] for r in conn.execute(
            "SELECT topic_code FROM question_topics WHERE question_version_id=?",
            (item["id"],),
        ).fetchall()]
        evidence = [dict(r) for r in conn.execute(
            "SELECT document_version_id, block_id, extraction_revision, quoted_text, supports "
            "FROM question_sources WHERE question_version_id=?",
            (item["id"],),
        ).fetchall()]
        conn.execute(
            """INSERT INTO exam_attempt_items
            (id, attempt_id, sequence_number, question_version_id, type, stem, stimulus,
             language, correct_option_ids_json, explanation, distractor_rationales_json,
             difficulty, section_label, topic_codes_json, points, max_score, evidence_json)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (item_id, attempt_id, sequence_number, item["id"],
             item["type"], item["stem"], item["stimulus"], item["language"],
             item["correct_option_ids_json"], item["explanation"],
             item["distractor_rationales_json"], item["difficulty"],
             item["section_label"], json.dumps(topics, ensure_ascii=False),
             item["points"], item["max_score"],
             json.dumps(evidence, ensure_ascii=False)),
        )
        for idx, opt in enumerate(opts):
            conn.execute(
                "INSERT INTO exam_attempt_options (id, attempt_item_id, option_code, option_text, display_order) "
                "VALUES (?, ?, ?, ?, ?)",
                (gen_id("atto"), item_id, opt["option_code"], opt["option_text"], idx),
            )
    checksum = _snapshot_checksum_for_attempt(conn, attempt_id)
    conn.execute(
        "UPDATE exam_attempts SET snapshot_checksum=? WHERE id=?",
        (checksum, attempt_id),
    )


def _snapshot_rng(random_seed, algorithm_version, scope):
    """Tạo stream cục bộ ổn định cho từng phần snapshot."""
    material = f"{algorithm_version}:{random_seed}:{scope}".encode("utf-8")
    return random.Random(int.from_bytes(hashlib.sha256(material).digest(), "big"))


def _shuffle_snapshot_values(values, rng):
    """Xáo trộn ổn định và đảm bảo cờ shuffle tạo thứ tự mới khi có thể."""
    original = list(values)
    rng.shuffle(values)
    if len(values) > 1 and values == original:
        values.append(values.pop(0))


def _snapshot_checksum_for_attempt(conn, attempt_id):
    """Canonicalize toàn bộ dữ liệu snapshot đã lưu theo thứ tự presentation."""
    items = []
    for item in conn.execute(
        "SELECT * FROM exam_attempt_items WHERE attempt_id=? ORDER BY sequence_number", (attempt_id,)
    ).fetchall():
        data = {key: item[key] for key in (
            "sequence_number", "question_version_id", "type", "stem", "stimulus", "language",
            "correct_option_ids_json", "explanation", "distractor_rationales_json", "difficulty",
            "section_label", "topic_codes_json", "points", "max_score", "evidence_json",
        )}
        data["options"] = [{key: option[key] for key in (
            "option_code", "option_text", "display_order",
        )} for option in conn.execute(
            "SELECT * FROM exam_attempt_options WHERE attempt_item_id=? ORDER BY display_order",
            (item["id"],),
        ).fetchall()]
        items.append(data)
    return _snapshot_checksum(items)


def _verify_snapshot_checksum_with_conn(conn, attempt_id):
    stored = conn.execute(
        "SELECT snapshot_checksum FROM exam_attempts WHERE id=?", (attempt_id,)
    ).fetchone()
    return bool(stored) and stored["snapshot_checksum"] == _snapshot_checksum_for_attempt(conn, attempt_id)


def snapshot_order(db_path, attempt_id):
    """Đọc presentation order đã snapshot của attempt."""
    conn = read_connection(db_path)
    try:
        if not conn.execute("SELECT 1 FROM exam_attempts WHERE id=?", (attempt_id,)).fetchone():
            raise TrainingError(ErrorCode.NOT_FOUND, "Bài làm không tồn tại", status=404)
        items = conn.execute(
            "SELECT id, sequence_number, question_version_id FROM exam_attempt_items "
            "WHERE attempt_id=? ORDER BY sequence_number", (attempt_id,)
        ).fetchall()
        return [
            {
                "sequence_number": item["sequence_number"],
                "qv": item["question_version_id"],
                "options": [
                    (option["option_code"], option["display_order"])
                    for option in conn.execute(
                        "SELECT option_code, display_order FROM exam_attempt_options "
                        "WHERE attempt_item_id=? ORDER BY display_order", (item["id"],)
                    ).fetchall()
                ],
            }
            for item in items
        ]
    finally:
        conn.close()


def reproduce_snapshot_order(db_path, attempt_id):
    """Tái tạo presentation order chỉ từ snapshot bất biến của attempt."""
    return snapshot_order(db_path, attempt_id)


def verify_snapshot_checksum(db_path, attempt_id):
    """Kiểm tra checksum của toàn bộ snapshot item/option đã lưu."""
    conn = read_connection(db_path)
    try:
        attempt = conn.execute(
            "SELECT snapshot_checksum FROM exam_attempts WHERE id=?", (attempt_id,)
        ).fetchone()
        if not attempt:
            raise TrainingError(ErrorCode.NOT_FOUND, "Bài làm không tồn tại", status=404)
        return _verify_snapshot_checksum_with_conn(conn, attempt_id)
    finally:
        conn.close()


def _attempt_start_result(attempt_id, conn, assignment_id):
    row = conn.execute(
        "SELECT status, started_at_ms, deadline_at_ms FROM exam_attempts WHERE id=?",
        (attempt_id,),
    ).fetchone()
    return {
        "attempt_id": attempt_id,
        "status": row["status"],
        "started_at_ms": row["started_at_ms"],
        "deadline_at_ms": row["deadline_at_ms"],
    }


def get_attempt(db_path, attempt_id):
    conn = read_connection(db_path)
    try:
        row = conn.execute("SELECT * FROM exam_attempts WHERE id=?", (attempt_id,)).fetchone()
        return dict(row) if row else None
    finally:
        conn.close()


def response_count(db_path, attempt_id):
    """Đếm response hiện có của một attempt qua kết nối chỉ đọc."""
    conn = read_connection(db_path)
    try:
        return conn.execute(
            "SELECT COUNT(*) AS count FROM exam_responses WHERE attempt_id=?", (attempt_id,)
        ).fetchone()["count"]
    finally:
        conn.close()


def attempt_owner_username(db_path, attempt_id):
    conn = read_connection(db_path)
    try:
        row = conn.execute(
            """SELECT x.username FROM exam_attempts a
               JOIN exam_assignments x ON x.id=a.assignment_id WHERE a.id=?""",
            (attempt_id,),
        ).fetchone()
        return row["username"] if row else None
    finally:
        conn.close()


def get_attempt_learner_view(db_path, attempt_id):
    """Learner DTO: không có correct/explanation/evidence/scoring."""
    conn = read_connection(db_path)
    try:
        attempt = conn.execute("SELECT * FROM exam_attempts WHERE id=?", (attempt_id,)).fetchone()
        if not attempt:
            raise TrainingError(ErrorCode.NOT_FOUND, "Bài làm không tồn tại", status=404)
        items = []
        for item in conn.execute(
            "SELECT * FROM exam_attempt_items WHERE attempt_id=? ORDER BY sequence_number",
            (attempt_id,),
        ).fetchall():
            opts = [
                {"id": o["option_code"], "text": o["option_text"]}
                for o in conn.execute(
                    "SELECT option_code, option_text FROM exam_attempt_options "
                    "WHERE attempt_item_id=? ORDER BY display_order",
                    (item["id"],),
                ).fetchall()
            ]
            response = conn.execute(
                "SELECT selected_option_ids_json, client_revision FROM exam_responses "
                "WHERE attempt_item_id=?",
                (item["id"],),
            ).fetchone()
            items.append({
                "item_id": item["id"],
                "sequence": item["sequence_number"],
                "type": item["type"],
                "stem": item["stem"],
                "stimulus": item["stimulus"],
                "options": opts,
                "points": item["points"],
                "response": {
                    "selected_option_ids": json.loads(response["selected_option_ids_json"]) if response else [],
                    "client_revision": response["client_revision"] if response else 0,
                },
            })
        return {
            "attempt_id": attempt_id,
            "status": attempt["status"],
            "deadline_at_ms": attempt["deadline_at_ms"],
            "items": items,
        }
    finally:
        conn.close()


def save_response(db_path, *, attempt_id, attempt_item_id, selected_option_ids, client_revision):
    """Autosave compare-and-set: chỉ ghi khi revision > stored."""
    now = time_policy.utc_now_ms()
    conn = write_connection(db_path)
    try:
        conn.execute("BEGIN IMMEDIATE")
        if (
            not isinstance(client_revision, int)
            or isinstance(client_revision, bool)
            or client_revision <= 0
        ):
            raise TrainingError(ErrorCode.VALIDATION_ERROR,
                                "Phiên bản client phải là số nguyên dương")
        attempt = conn.execute(
            "SELECT status, deadline_at_ms FROM exam_attempts WHERE id=?",
            (attempt_id,),
        ).fetchone()
        if not attempt:
            raise TrainingError(ErrorCode.NOT_FOUND, "Bài làm không tồn tại", status=404)
        if attempt["status"] != constants.AttemptStatus.ACTIVE:
            raise TrainingError(ErrorCode.ATTEMPT_ALREADY_COMPLETED,
                                "Bài làm đã kết thúc", status=409)
        if now >= attempt["deadline_at_ms"]:
            raise TrainingError(ErrorCode.ATTEMPT_EXPIRED, "Đã hết thời gian làm bài", status=410)

        item = conn.execute(
            "SELECT type FROM exam_attempt_items WHERE id=? AND attempt_id=?",
            (attempt_item_id, attempt_id),
        ).fetchone()
        if not item:
            raise TrainingError(ErrorCode.ATTEMPT_ITEM_NOT_FOUND,
                                "Câu hỏi không thuộc bài làm", status=404)
        if (
            not isinstance(selected_option_ids, list)
            or any(not isinstance(option_id, str) for option_id in selected_option_ids)
            or len(set(selected_option_ids)) != len(selected_option_ids)
        ):
            raise TrainingError(ErrorCode.INVALID_OPTION_SELECTION,
                                "Lựa chọn đáp án không hợp lệ")
        option_ids = {
            row["option_code"] for row in conn.execute(
                "SELECT option_code FROM exam_attempt_options WHERE attempt_item_id=?",
                (attempt_item_id,),
            ).fetchall()
        }
        if any(option_id not in option_ids for option_id in selected_option_ids):
            raise TrainingError(ErrorCode.INVALID_OPTION_SELECTION,
                                "Lựa chọn đáp án không hợp lệ")
        if item["type"] == "single_choice" and len(selected_option_ids) > 1:
            raise TrainingError(ErrorCode.SINGLE_CHOICE_REQUIRES_ONE_OPTION,
                                "Câu hỏi một lựa chọn chỉ nhận tối đa một đáp án")

        existing = conn.execute(
            "SELECT client_revision FROM exam_responses WHERE attempt_id=? AND attempt_item_id=?",
            (attempt_id, attempt_item_id),
        ).fetchone()
        stored_rev = existing["client_revision"] if existing else 0
        if client_revision <= stored_rev:
            conn.rollback()
            return {"accepted": False, "stored_revision": stored_rev}

        selected_json = json.dumps(selected_option_ids, ensure_ascii=False)
        if existing:
            conn.execute(
                """UPDATE exam_responses
                SET selected_option_ids_json=?, client_revision=?, answered_at_ms=?
                WHERE attempt_id=? AND attempt_item_id=?""",
                (selected_json, client_revision, now, attempt_id, attempt_item_id),
            )
        else:
            conn.execute(
                """INSERT INTO exam_responses
                (id, attempt_id, attempt_item_id, selected_option_ids_json, client_revision, answered_at_ms)
                VALUES (?, ?, ?, ?, ?, ?)""",
                (gen_id("resp"), attempt_id, attempt_item_id, selected_json, client_revision, now),
            )
        conn.commit()
        return {"accepted": True, "stored_revision": client_revision}
    except TrainingError:
        if conn.in_transaction:
            conn.rollback()
        raise
    except Exception:
        try:
            conn.execute("ROLLBACK")
        except Exception:
            pass
        raise
    finally:
        conn.close()


def submit_attempt(db_path, *, unit_code, actor, attempt_id):
    """Submit: compare-and-set active→submitted/timed_out, score, create result.

    Idempotent: nếu đã có result, trả lại.
    """
    conn = write_connection(db_path)
    try:
        conn.execute("BEGIN IMMEDIATE")
        attempt = conn.execute(
            "SELECT * FROM exam_attempts WHERE id=?", (attempt_id,)
        ).fetchone()
        if not attempt:
            raise TrainingError(ErrorCode.NOT_FOUND, "Bài làm không tồn tại", status=404)
        attempt = dict(attempt)

        existing_result = conn.execute(
            "SELECT * FROM exam_results WHERE attempt_id=?", (attempt_id,)
        ).fetchone()
        if existing_result:
            if attempt["status"] == constants.AttemptStatus.ADMIN_SUBMITTED:
                raise TrainingError(ErrorCode.ATTEMPT_ALREADY_COMPLETED,
                                    "Kỳ thi đã đóng", status=409)
            conn.rollback()
            return _build_submit_response(attempt, existing_result, conn)

        if attempt["status"] in constants.ATTEMPT_TERMINAL_STATUSES:
            raise TrainingError(ErrorCode.ATTEMPT_ALREADY_COMPLETED,
                                "Bài làm đã kết thúc", status=409)

        now = time_policy.utc_now_ms()
        is_timeout = now >= attempt["deadline_at_ms"]
        exam_status = conn.execute(
            """SELECT ee.status FROM exam_events ee JOIN exam_assignments x
               ON x.exam_event_id=ee.id WHERE x.id=?""",
            (attempt["assignment_id"],),
        ).fetchone()["status"]
        if exam_status != constants.ExamStatus.OPEN and not is_timeout:
            raise TrainingError(ErrorCode.ATTEMPT_ALREADY_COMPLETED,
                                "Kỳ thi đã đóng", status=409)
        new_status = constants.AttemptStatus.TIMED_OUT if is_timeout else constants.AttemptStatus.SUBMITTED
        ended_reason = "timeout" if is_timeout else "submit"
        result = _complete_active_attempt(
            conn, unit_code=unit_code, actor=actor, attempt=attempt,
            status=new_status, ended_reason=ended_reason, action="submit_attempt",
        )
        conn.commit()
        return _build_submit_response(
            attempt | {"status": new_status, "submitted_at_ms": now}, result, conn
        )
    except TrainingError:
        if conn.in_transaction:
            conn.rollback()
        raise
    except Exception:
        try:
            conn.execute("ROLLBACK")
        except Exception:
            pass
        raise
    finally:
        conn.close()


def administratively_submit_attempt(db_path, *, unit_code, actor, attempt_id, ended_reason="exam_closed"):
    """Kết thúc attempt active theo thao tác quản trị, idempotently."""
    conn = write_connection(db_path)
    try:
        conn.execute("BEGIN IMMEDIATE")
        attempt = conn.execute("SELECT * FROM exam_attempts WHERE id=?", (attempt_id,)).fetchone()
        if not attempt:
            raise TrainingError(ErrorCode.NOT_FOUND, "Bài làm không tồn tại", status=404)
        attempt = dict(attempt)
        if attempt["status"] != constants.AttemptStatus.ACTIVE:
            result = conn.execute(
                "SELECT * FROM exam_results WHERE attempt_id=?", (attempt_id,)
            ).fetchone()
            conn.rollback()
            return dict(result) if result else None
        result = _complete_active_attempt(
            conn, unit_code=unit_code, actor=actor, attempt=attempt,
            status=constants.AttemptStatus.ADMIN_SUBMITTED, ended_reason=ended_reason,
            action="administratively_submit_attempt",
        )
        conn.commit()
        return dict(result)
    except TrainingError:
        if conn.in_transaction:
            conn.rollback()
        raise
    finally:
        conn.close()


def _complete_active_attempt(conn, *, unit_code, actor, attempt, status, ended_reason, action):
    """CAS terminal transition, immutable-snapshot scoring, result and audit in one transaction."""
    if not _verify_snapshot_checksum_with_conn(conn, attempt["id"]):
        raise TrainingError(ErrorCode.CONFLICT, "Snapshot bài làm không toàn vẹn", status=409)
    now = time_policy.utc_now_ms()
    cursor = conn.execute(
        "UPDATE exam_attempts SET status=?, submitted_at_ms=?, ended_reason=? WHERE id=? AND status='active'",
        (status, now, ended_reason, attempt["id"]),
    )
    if cursor.rowcount == 0:
        raise TrainingError(ErrorCode.ATTEMPT_ALREADY_COMPLETED,
                            "Bài làm không ở trạng thái active", status=409)
    assignment = conn.execute(
        "SELECT exam_event_id FROM exam_assignments WHERE id=?", (attempt["assignment_id"],)
    ).fetchone()
    exam = conn.execute(
        "SELECT pass_score_percent FROM exam_events WHERE id=?", (assignment["exam_event_id"],)
    ).fetchone()
    score_result = scoring.score_attempt_with_conn(conn, attempt["id"])
    result_id = scoring.store_result(
        conn, attempt["id"], score_result, pass_score_percent=exam["pass_score_percent"],
    )
    conn.execute("UPDATE exam_assignments SET status='completed' WHERE id=?", (attempt["assignment_id"],))
    write_audit(
        conn, actor=actor, unit_code=unit_code, action=action,
        entity_type="exam_attempt", entity_id=attempt["id"],
        after={"ended_reason": ended_reason} if action == "administratively_submit_attempt" else None,
    )
    return conn.execute("SELECT * FROM exam_results WHERE id=?", (result_id,)).fetchone()


def _build_submit_response(attempt, result, conn):
    return {
        "attempt_id": attempt["id"],
        "status": attempt["status"],
        "submitted_at_ms": attempt.get("submitted_at_ms"),
        "result": {
            "score": result["raw_score"],
            "maximum_score": result["maximum_score"],
            "percent": result["percent"],
            "passed": bool(result["passed"]),
        },
        "answers_released": False,
    }


def get_result(db_path, attempt_id):
    conn = read_connection(db_path)
    try:
        row = conn.execute("SELECT * FROM exam_results WHERE attempt_id=?", (attempt_id,)).fetchone()
        return dict(row) if row else None
    finally:
        conn.close()
