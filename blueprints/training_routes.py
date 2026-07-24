from functools import wraps

from flask import Blueprint, jsonify, render_template, request, session

import config
from app_helpers import add_no_cache_headers, csrf_protect
from auth import get_user_by_username
from training.errors import TrainingError
from training.permissions import has_module_role
from services import training_attempt_service as attempts
from services import training_exam_service as exams
from services import training_report_service as reports
from services import training_knowledge_service as knowledge
from services import training_question_service as questions

training_bp = Blueprint("training", __name__)


def _error_response(exc):
    return jsonify(exc.to_envelope()), exc.status


def _manager_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        username = session.get("username")
        user = get_user_by_username(username) if username else None
        if not (user and user.get("role") == "admin") and not has_module_role(
            config.TRAINING_DB_PATH, username, "exam_manager"
        ):
            return jsonify({"error": {"code": "PERMISSION_SCOPE_DENIED", "message": "Không có quyền quản lý kỳ thi.", "details": {}}}), 403
        return view(*args, **kwargs)
    return wrapped


def _attempt_owned_by_current_user(attempt_id):
    return attempts.attempt_owner_username(config.TRAINING_DB_PATH, attempt_id) == session.get("username")


@training_bp.route("/dao-tao-sat-hach")
def page_index():
    return render_template(
        "pages/training/index.html",
        current_user=get_user_by_username(session.get("username")),
        active_page="training",
    )


@training_bp.route("/api/training/knowledge", methods=["GET", "POST"])
@_manager_required
def knowledge_collection():
    if request.method == "GET":
        result = knowledge.list_documents(
            config.TRAINING_DB_PATH,
            status=request.args.get("status"),
            page=max(1, request.args.get("page", 1, type=int)),
            page_size=min(100, max(1, request.args.get("page_size", 25, type=int))),
        )
        return jsonify(result)

    token = request.form.get("csrf_token") or request.headers.get("X-CSRF-Token")
    from app_helpers import validate_csrf_token
    if not validate_csrf_token(token):
        return jsonify({"error": {"code": "CSRF_INVALID", "message": "CSRF token không hợp lệ", "details": {}}}), 400
    payload = request.get_json(silent=True) or {}
    content = payload.get("content_text", "")
    if not isinstance(content, str) or not content.strip():
        return jsonify({"error": {"code": "VALIDATION_ERROR", "message": "Nội dung văn bản là bắt buộc.", "details": {}}}), 400
    if len(content.encode("utf-8")) > config.TRAINING_MAX_TEXT_BYTES:
        return jsonify({"error": {"code": "VALIDATION_ERROR", "message": "Nội dung vượt giới hạn cấu hình.", "details": {}}}), 400
    try:
        result = knowledge.create_document(
            config.TRAINING_DB_PATH, unit_code=config.UNIT_CODE, actor=session["username"],
            document_code=payload["document_code"], title=payload["title"],
            content_text=content, classification=payload.get("classification"),
            audience_codes=payload.get("audience_codes"),
        )
        return jsonify(result), 201
    except (KeyError, ValueError) as exc:
        return jsonify({"error": {"code": "VALIDATION_ERROR", "message": str(exc), "details": {}}}), 400


@training_bp.route("/api/training/questions/import", methods=["POST"])
@csrf_protect
@_manager_required
def import_questions():
    payload = request.get_json(silent=True) or {}
    try:
        result = questions.import_question_batch(
            config.TRAINING_DB_PATH, unit_code=config.UNIT_CODE,
            actor=session["username"], batch=payload, status="draft",
        )
        return jsonify(result), 201
    except TrainingError as exc:
        return _error_response(exc)


@training_bp.route("/api/training/questions/<version_id>/approve", methods=["POST"])
@csrf_protect
@_manager_required
def approve_question(version_id):
    try:
        questions.add_review_action(config.TRAINING_DB_PATH, unit_code=config.UNIT_CODE,
                                    actor=session["username"], version_id=version_id, action="approve")
        return jsonify({"version_id": version_id, "review_status": "approved"})
    except TrainingError as exc:
        return _error_response(exc)


@training_bp.route("/api/training/questions/<version_id>/publish", methods=["POST"])
@csrf_protect
@_manager_required
def publish_question(version_id):
    try:
        questions.publish_question_version(config.TRAINING_DB_PATH, unit_code=config.UNIT_CODE,
                                           actor=session["username"], version_id=version_id)
        return jsonify({"version_id": version_id, "publication_status": "published"})
    except TrainingError as exc:
        return _error_response(exc)


@training_bp.route("/api/training/templates", methods=["POST"])
@csrf_protect
@_manager_required
def create_template():
    payload = request.get_json(silent=True) or {}
    try:
        result = exams.create_template(
            config.TRAINING_DB_PATH, unit_code=config.UNIT_CODE, actor=session["username"],
            code=payload["code"], title=payload["title"],
            target_audience_code=payload["target_audience_code"],
            question_version_ids=payload["question_version_ids"],
            duration_seconds=payload["duration_seconds"],
            pass_score_percent=payload.get("pass_score_percent", 80.0),
            shuffle_questions=bool(payload.get("shuffle_questions")),
            shuffle_options=bool(payload.get("shuffle_options")),
        )
        return jsonify(result), 201
    except (KeyError, ValueError) as exc:
        return jsonify({"error": {"code": "VALIDATION_ERROR", "message": str(exc), "details": {}}}), 400
    except TrainingError as exc:
        return _error_response(exc)


@training_bp.route("/api/training/exams", methods=["POST"])
@csrf_protect
@_manager_required
def create_exam():
    payload = request.get_json(silent=True) or {}
    try:
        result = exams.create_exam(
            config.TRAINING_DB_PATH, unit_code=config.UNIT_CODE, actor=session["username"],
            code=payload["code"], title=payload["title"], template_id=payload["template_id"],
            target_audience_code=payload["target_audience_code"], start_at_ms=payload["start_at_ms"],
            end_at_ms=payload["end_at_ms"], duration_seconds=payload["duration_seconds"],
            pass_score_percent=payload.get("pass_score_percent", 80.0),
            description=payload.get("description"),
        )
        return jsonify(result), 201
    except (KeyError, ValueError) as exc:
        return jsonify({"error": {"code": "VALIDATION_ERROR", "message": str(exc), "details": {}}}), 400
    except TrainingError as exc:
        return _error_response(exc)


@training_bp.route("/api/training/exams/<exam_id>/assignments", methods=["POST"])
@csrf_protect
@_manager_required
def create_assignments(exam_id):
    payload = request.get_json(silent=True) or {}
    try:
        assignment_ids = exams.create_assignments(
            config.TRAINING_DB_PATH, unit_code=config.UNIT_CODE, actor=session["username"],
            exam_id=exam_id, users=payload["users"], audience_code=payload["audience_code"],
        )
        return jsonify({"assignment_ids": assignment_ids}), 201
    except (KeyError, ValueError) as exc:
        return jsonify({"error": {"code": "VALIDATION_ERROR", "message": str(exc), "details": {}}}), 400
    except TrainingError as exc:
        return _error_response(exc)


@training_bp.route("/api/training/exams/<exam_id>/<action>", methods=["POST"])
@csrf_protect
@_manager_required
def transition_exam(exam_id, action):
    handlers = {"ready": exams.ready_exam, "open": exams.open_exam, "close": exams.close_exam}
    handler = handlers.get(action)
    if handler is None:
        return jsonify({"error": {"code": "NOT_FOUND", "message": "Thao tác không tồn tại.", "details": {}}}), 404
    try:
        handler(config.TRAINING_DB_PATH, unit_code=config.UNIT_CODE, actor=session["username"], exam_id=exam_id)
        return jsonify({"exam_id": exam_id, "action": action})
    except TrainingError as exc:
        return _error_response(exc)


@training_bp.route("/api/training/assignments/<assignment_id>/attempts", methods=["POST"])
@csrf_protect
def start_attempt(assignment_id):
    try:
        result = attempts.start_attempt(
            config.TRAINING_DB_PATH, unit_code=config.UNIT_CODE,
            actor=session["username"], assignment_id=assignment_id,
        )
        return jsonify(result)
    except TrainingError as exc:
        return _error_response(exc)


@training_bp.route("/api/training/attempts/<attempt_id>")
def get_attempt(attempt_id):
    try:
        if not _attempt_owned_by_current_user(attempt_id):
            raise TrainingError("PERMISSION_SCOPE_DENIED", "Bạn không sở hữu bài làm này.", status=403)
        view = attempts.get_attempt_learner_view(config.TRAINING_DB_PATH, attempt_id)
        response = jsonify(view)
        return add_no_cache_headers(response)
    except TrainingError as exc:
        return _error_response(exc)


@training_bp.route("/api/training/attempts/<attempt_id>/responses/<item_id>", methods=["PUT"])
@csrf_protect
def save_response(attempt_id, item_id):
    payload = request.get_json(silent=True) or {}
    try:
        if not _attempt_owned_by_current_user(attempt_id):
            raise TrainingError("PERMISSION_SCOPE_DENIED", "Bạn không sở hữu bài làm này.", status=403)
        result = attempts.save_response(
            config.TRAINING_DB_PATH, attempt_id=attempt_id, attempt_item_id=item_id,
            selected_option_ids=payload.get("selected_option_ids", []),
            client_revision=payload.get("client_revision", 0),
        )
        return jsonify(result)
    except TrainingError as exc:
        return _error_response(exc)


@training_bp.route("/api/training/attempts/<attempt_id>/submit", methods=["POST"])
@csrf_protect
def submit_attempt(attempt_id):
    try:
        if not _attempt_owned_by_current_user(attempt_id):
            raise TrainingError("PERMISSION_SCOPE_DENIED", "Bạn không sở hữu bài làm này.", status=403)
        return jsonify(attempts.submit_attempt(
            config.TRAINING_DB_PATH, unit_code=config.UNIT_CODE,
            actor=session["username"], attempt_id=attempt_id,
        ))
    except TrainingError as exc:
        return _error_response(exc)


@training_bp.route("/api/training/exams/<exam_id>/finalize", methods=["POST"])
@csrf_protect
@_manager_required
def finalize_exam(exam_id):
    try:
        return jsonify(reports.finalize_exam(
            config.TRAINING_DB_PATH, unit_code=config.UNIT_CODE,
            actor=session["username"], exam_id=exam_id,
        ))
    except TrainingError as exc:
        return _error_response(exc)


@training_bp.route("/api/training/exams/<exam_id>/report")
@_manager_required
def get_report(exam_id):
    try:
        report = reports.get_report_snapshot(config.TRAINING_DB_PATH, exam_id)
        if report is None:
            return jsonify({"error": {"code": "NOT_FOUND", "message": "Kỳ thi chưa được chốt.", "details": {}}}), 404
        return jsonify(report)
    except TrainingError as exc:
        return _error_response(exc)
