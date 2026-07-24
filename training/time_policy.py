"""Chính sách thời gian module đào tạo.

DB lưu UTC epoch milliseconds. Server là nguồn thời gian chuẩn.
Deadline = min(exam.end_at, attempt.started_at + assignment.duration_seconds).
"""

import time
from datetime import datetime, timezone


def utc_now_ms():
    """Thời gian hiện tại theo UTC, epoch milliseconds."""
    return int(time.time() * 1000)


def compute_attempt_deadline_ms(started_at_ms, exam_end_at_ms, duration_ms):
    """Deadline = min(exam_end, started + duration). Tất cả tham số theo ms."""
    return min(exam_end_at_ms, started_at_ms + duration_ms)


def is_past_deadline(deadline_at_ms, *, now_ms=None):
    if now_ms is None:
        now_ms = utc_now_ms()
    return now_ms >= deadline_at_ms


def ms_to_rfc3339(ms):
    """Epoch ms -> chuỗi RFC 3339 kết thúc Z (UTC)."""
    seconds = ms / 1000.0
    dt = datetime.fromtimestamp(seconds, tz=timezone.utc)
    return dt.strftime("%Y-%m-%dT%H:%M:%SZ")


def rfc3339_to_ms(text):
    """RFC 3339 (kết thúc Z) -> epoch ms."""
    text = text.rstrip("Z")
    dt = datetime.strptime(text, "%Y-%m-%dT%H:%M:%S").replace(tzinfo=timezone.utc)
    return int(dt.timestamp() * 1000)
