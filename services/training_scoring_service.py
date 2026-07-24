"""Scoring service: chấm chỉ từ immutable attempt snapshot.

Một đáp án: đúng nhận toàn bộ điểm, sai/không trả lời nhận 0.
Không trừ điểm âm trong MVP. Result hiệu lực = result gốc.
"""

import json

from training import constants, time_policy
from training.db import read_connection
from repositories.training_repository import gen_id


def score_attempt(db_path, attempt_id):
    """Đọc snapshot items + responses, tính điểm. Trả result dict."""
    conn = read_connection(db_path)
    try:
        return score_attempt_with_conn(conn, attempt_id)
    finally:
        conn.close()


def score_attempt_with_conn(conn, attempt_id):
    """Score sử dụng connection có sẵn (cho dùng trong transaction)."""
    items = [dict(r) for r in conn.execute(
        "SELECT * FROM exam_attempt_items WHERE attempt_id=? ORDER BY sequence_number",
        (attempt_id,),
    ).fetchall()]
    responses = {
        r["attempt_item_id"]: r for r in conn.execute(
            "SELECT * FROM exam_responses WHERE attempt_id=?", (attempt_id,),
        ).fetchall()
    }
    raw_score = 0.0
    maximum_score = 0.0
    topic_stats = {}

    for item in items:
        points = item["points"]
        maximum_score += points
        correct_ids = set(json.loads(item["correct_option_ids_json"]))
        response = responses.get(item["id"])
        selected = set(json.loads(response["selected_option_ids_json"])) if response else set()

        is_correct = selected == correct_ids and len(correct_ids) > 0
        earned = points if is_correct else 0.0
        raw_score += earned

        topics = json.loads(item.get("topic_codes_json") or "[]")
        if not topics:
            topics = ["uncategorized"]
        for tc in topics:
            stats = topic_stats.setdefault(
                tc,
                {"correct": 0, "total": 0, "points": 0.0, "max_points": 0.0},
            )
            stats["total"] += 1
            stats["max_points"] += points
            if is_correct:
                stats["correct"] += 1
                stats["points"] += earned

    percent = round((raw_score / maximum_score * 100), 2) if maximum_score > 0 else 0.0
    return {
        "raw_score": raw_score,
        "maximum_score": maximum_score,
        "percent": percent,
        "topic_breakdown": topic_stats,
    }


def compute_passed(score_result, pass_score_percent):
    return score_result["percent"] >= pass_score_percent


def store_result(conn, attempt_id, score_result, *, pass_score_percent, status="scored"):
    """Tạo result record trong transaction. Idempotent: gọi trong context đã kiểm tra."""
    now = time_policy.utc_now_ms()
    passed = compute_passed(score_result, pass_score_percent)
    result_id = gen_id("res")
    conn.execute(
        """INSERT OR IGNORE INTO exam_results
        (id, attempt_id, raw_score, maximum_score, percent, passed,
         topic_breakdown_json, scored_at_ms, status)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
        (result_id, attempt_id, score_result["raw_score"], score_result["maximum_score"],
         score_result["percent"], int(passed),
         json.dumps(score_result["topic_breakdown"], ensure_ascii=False),
         now, status),
    )
    for tc, stats in score_result["topic_breakdown"].items():
        conn.execute(
            "INSERT OR IGNORE INTO exam_result_topics "
            "(result_id, topic_code, correct_count, total_count, points, max_points) "
            "VALUES (?, ?, ?, ?, ?, ?)",
            (result_id, tc, stats["correct"], stats["total"],
             stats["points"], stats["max_points"]),
        )
    return result_id
