"""Enum trạng thái module Đào tạo & sát hạch.

Giá trị chuỗi ổn định, dùng trong DB và API. Không đổi giá trị đã phát hành.
"""


class _Status:
    def __init__(self, value):
        self.value = value

    def __str__(self):
        return self.value

    def __eq__(self, other):
        if isinstance(other, _Status):
            return self.value == other.value
        return self.value == other

    def __hash__(self):
        return hash(self.value)

    def __repr__(self):
        return f"<Status {self.value!r}>"


class DocumentReviewStatus:
    DRAFT = "draft"
    ANALYZED = "analyzed"
    APPROVED = "approved"
    NEEDS_CONFIRMATION = "needs_confirmation"
    REJECTED = "rejected"


class QuestionReviewStatus:
    DRAFT = "draft"
    NEEDS_REVIEW = "needs_review"
    APPROVED = "approved"
    REJECTED = "rejected"


class PublicationStatus:
    UNPUBLISHED = "unpublished"
    PUBLISHED = "published"
    RETIRED = "retired_for_new_exams"


class ExamStatus:
    DRAFT = "draft"
    READY = "ready"
    OPEN = "open"
    CLOSED = "closed"
    CANCELLED = "cancelled"


class AssignmentStatus:
    ASSIGNED = "assigned"
    COMPLETED = "completed"
    EXPIRED = "expired"
    CANCELLED = "cancelled"


class AttemptStatus:
    CREATED = "created"
    ACTIVE = "active"
    SUBMITTED = "submitted"
    TIMED_OUT = "timed_out"
    ADMIN_SUBMITTED = "administratively_submitted"
    INVALIDATED = "invalidated"


class ResultStatus:
    SCORED = "scored"
    INVALIDATED = "invalidated"


class JobStatus:
    PENDING = "pending"
    RUNNING = "running"
    COMPLETED = "completed"
    FAILED = "failed"
    CANCELLED = "cancelled"


class IssueStatus:
    OPEN = "open"
    CONFIRMED = "confirmed"
    EXCLUDED = "excluded"
    RESOLVED = "resolved_by_new_version"


class IssueSeverity:
    LOW = "low"
    MEDIUM = "medium"
    HIGH = "high"


ATTEMPT_TERMINAL_STATUSES = frozenset({
    AttemptStatus.SUBMITTED,
    AttemptStatus.TIMED_OUT,
    AttemptStatus.ADMIN_SUBMITTED,
    AttemptStatus.INVALIDATED,
})

EXAM_TERMINAL_STATUSES = frozenset({
    ExamStatus.CLOSED,
    ExamStatus.CANCELLED,
})

QUESTION_TYPES_MVP = frozenset({"single_choice"})

QUESTION_TYPES_ALL = frozenset({
    "single_choice",
    "multiple_choice",
    "true_false",
    "scenario_single_choice",
    "scenario_multiple_choice",
})

MODULE_ROLES = frozenset({"learner", "editor", "exam_manager", "admin"})
