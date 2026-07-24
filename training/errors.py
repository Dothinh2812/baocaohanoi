"""Mã lỗi ổn định và exception module đào tạo."""


class ErrorCode:
    ASSIGNMENT_NOT_FOUND = "ASSIGNMENT_NOT_FOUND"
    ATTEMPT_ALREADY_ACTIVE = "ATTEMPT_ALREADY_ACTIVE"
    ATTEMPT_ALREADY_COMPLETED = "ATTEMPT_ALREADY_COMPLETED"
    ATTEMPT_EXPIRED = "ATTEMPT_EXPIRED"
    ATTEMPT_ITEM_NOT_FOUND = "ATTEMPT_ITEM_NOT_FOUND"
    EXAM_NOT_OPEN = "EXAM_NOT_OPEN"
    INVALID_OPTION_SELECTION = "INVALID_OPTION_SELECTION"
    QUESTION_SUPPLY_INSUFFICIENT = "QUESTION_SUPPLY_INSUFFICIENT"
    SINGLE_CHOICE_REQUIRES_ONE_OPTION = "SINGLE_CHOICE_REQUIRES_ONE_OPTION"
    TEMPLATE_QUESTION_INVALID = "TEMPLATE_QUESTION_INVALID"
    TEMPLATE_QUESTION_DUPLICATE = "TEMPLATE_QUESTION_DUPLICATE"
    VERSION_CONFLICT = "VERSION_CONFLICT"
    DOCUMENT_HAS_BLOCKING_ISSUES = "DOCUMENT_HAS_BLOCKING_ISSUES"
    FINALIZATION_ALREADY_COMPLETED = "FINALIZATION_ALREADY_COMPLETED"
    PERMISSION_SCOPE_DENIED = "PERMISSION_SCOPE_DENIED"
    NOT_FOUND = "NOT_FOUND"
    VALIDATION_ERROR = "VALIDATION_ERROR"
    CONFLICT = "CONFLICT"
    PROVIDER_ERROR = "PROVIDER_ERROR"


class TrainingError(Exception):
    def __init__(self, code, message, *, status=400, details=None):
        super().__init__(message)
        self.code = code
        self.message = message
        self.status = status
        self.details = details or {}

    def to_envelope(self):
        return {
            "error": {
                "code": self.code,
                "message": self.message,
                "details": self.details,
            }
        }


def training_error_response(err):
    """Trả (body_dict, http_status) cho một TrainingError."""
    return err.to_envelope(), err.status
