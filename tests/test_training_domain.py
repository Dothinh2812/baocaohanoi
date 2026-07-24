import time

import pytest

from training import constants, errors, time_policy


# --- time_policy ---

def test_utc_now_ms_is_integer_and_recent():
    now = time_policy.utc_now_ms()
    assert isinstance(now, int)
    assert now > 0
    assert abs(now - int(time.time() * 1000)) < 5000


def test_deadline_is_min_of_exam_end_and_start_plus_duration():
    started = 1_000_000
    duration = 25 * 60 * 1000
    exam_end = started + 60 * 60 * 1000
    deadline = time_policy.compute_attempt_deadline_ms(started, exam_end, duration)
    assert deadline == started + duration


def test_deadline_uses_exam_end_when_sooner_than_duration():
    started = 1_000_000
    duration = 60 * 60 * 1000
    exam_end = started + 25 * 60 * 1000
    deadline = time_policy.compute_attempt_deadline_ms(started, exam_end, duration)
    assert deadline == exam_end


def test_deadline_boundary_when_equal():
    started = 1_000_000
    duration = 25 * 60 * 1000
    exam_end = started + duration
    deadline = time_policy.compute_attempt_deadline_ms(started, exam_end, duration)
    assert deadline == exam_end


def test_is_expired_true_past_deadline():
    deadline = time_policy.utc_now_ms() - 1000
    assert time_policy.is_past_deadline(deadline) is True


def test_is_expired_false_before_deadline():
    deadline = time_policy.utc_now_ms() + 100_000
    assert time_policy.is_past_deadline(deadline) is False


def test_ms_to_rfc3339_z():
    ts = 1_721_731_200_000
    text = time_policy.ms_to_rfc3339(ts)
    assert text.endswith("Z")
    assert "2024" in text


# --- constants (enum stability) ---

def test_review_status_constants():
    assert constants.DocumentReviewStatus.DRAFT == "draft"
    assert constants.DocumentReviewStatus.ANALYZED == "analyzed"
    assert constants.DocumentReviewStatus.APPROVED == "approved"


def test_question_review_status_constants():
    assert constants.QuestionReviewStatus.DRAFT == "draft"
    assert constants.QuestionReviewStatus.NEEDS_REVIEW == "needs_review"
    assert constants.QuestionReviewStatus.APPROVED == "approved"
    assert constants.QuestionReviewStatus.REJECTED == "rejected"


def test_publication_status_constants():
    assert constants.PublicationStatus.UNPUBLISHED == "unpublished"
    assert constants.PublicationStatus.PUBLISHED == "published"
    assert constants.PublicationStatus.RETIRED == "retired_for_new_exams"


def test_exam_status_constants():
    assert constants.ExamStatus.DRAFT == "draft"
    assert constants.ExamStatus.READY == "ready"
    assert constants.ExamStatus.OPEN == "open"
    assert constants.ExamStatus.CLOSED == "closed"
    assert constants.ExamStatus.CANCELLED == "cancelled"


def test_assignment_status_constants():
    assert constants.AssignmentStatus.ASSIGNED == "assigned"
    assert constants.AssignmentStatus.COMPLETED == "completed"
    assert constants.AssignmentStatus.EXPIRED == "expired"
    assert constants.AssignmentStatus.CANCELLED == "cancelled"


def test_attempt_status_constants():
    assert constants.AttemptStatus.CREATED == "created"
    assert constants.AttemptStatus.ACTIVE == "active"
    assert constants.AttemptStatus.SUBMITTED == "submitted"
    assert constants.AttemptStatus.TIMED_OUT == "timed_out"
    assert constants.AttemptStatus.ADMIN_SUBMITTED == "administratively_submitted"
    assert constants.AttemptStatus.INVALIDATED == "invalidated"


def test_attempt_terminal_states():
    terminals = constants.ATTEMPT_TERMINAL_STATUSES
    assert "submitted" in terminals
    assert "timed_out" in terminals
    assert "administratively_submitted" in terminals
    assert "invalidated" in terminals
    assert "active" not in terminals
    assert "created" not in terminals


def test_result_status_constants():
    assert constants.ResultStatus.SCORED == "scored"
    assert constants.ResultStatus.INVALIDATED == "invalidated"


def test_job_status_constants():
    assert constants.JobStatus.PENDING == "pending"
    assert constants.JobStatus.RUNNING == "running"
    assert constants.JobStatus.COMPLETED == "completed"
    assert constants.JobStatus.FAILED == "failed"
    assert constants.JobStatus.CANCELLED == "cancelled"


# --- errors ---

def test_error_codes_are_stable_strings():
    assert errors.ErrorCode.ATTEMPT_ALREADY_ACTIVE == "ATTEMPT_ALREADY_ACTIVE"
    assert errors.ErrorCode.EXAM_NOT_OPEN == "EXAM_NOT_OPEN"
    assert errors.ErrorCode.VERSION_CONFLICT == "VERSION_CONFLICT"
    assert errors.ErrorCode.ASSIGNMENT_NOT_FOUND == "ASSIGNMENT_NOT_FOUND"
    assert errors.ErrorCode.DOCUMENT_HAS_BLOCKING_ISSUES == "DOCUMENT_HAS_BLOCKING_ISSUES"
    assert errors.ErrorCode.FINALIZATION_ALREADY_COMPLETED == "FINALIZATION_ALREADY_COMPLETED"
    assert errors.ErrorCode.PERMISSION_SCOPE_DENIED == "PERMISSION_SCOPE_DENIED"


def test_training_error_carries_code_and_details():
    err = errors.TrainingError(errors.ErrorCode.EXAM_NOT_OPEN, "Kỳ thi chưa mở", status=409)
    assert err.code == "EXAM_NOT_OPEN"
    assert err.status == 409
    envelope = err.to_envelope()
    assert envelope["error"]["code"] == "EXAM_NOT_OPEN"
    assert envelope["error"]["message"] == "Kỳ thi chưa mở"


def test_training_error_default_details_empty():
    err = errors.TrainingError(errors.ErrorCode.ATTEMPT_EXPIRED, "Hết giờ")
    assert err.details == {}
