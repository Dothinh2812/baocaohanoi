from functools import wraps
import uuid

from flask import Blueprint, jsonify, render_template, request, send_file, session

import config
from app_helpers import add_no_cache_headers, csrf_protect
from auth import get_user_by_username
from training.errors import TrainingError
from training.permissions import has_module_role, module_roles
from services import training_attempt_service as attempts
from services import training_exam_service as exams
from services import training_report_service as reports
from services import training_knowledge_service as knowledge
from services import training_question_service as questions

training_bp = Blueprint("training", __name__)


def _error_response(exc):
    return jsonify(exc.to_envelope()), exc.status


def _module_role_required(role, message):
    def decorator(view):
        @wraps(view)
        def wrapped(*args, **kwargs):
            username = session.get("username")
            user = get_user_by_username(username) if username else None
            is_dashboard_admin = user and user.get("role") == "admin"
            roles = role if isinstance(role, tuple) else (role,)
            if not is_dashboard_admin and not any(
                has_module_role(config.TRAINING_DB_PATH, username, item) for item in roles
            ):
                return _error_response(TrainingError("PERMISSION_SCOPE_DENIED", message, status=403))
            return view(*args, **kwargs)
        return wrapped
    return decorator


_learner_required = _module_role_required("learner", "Không có quyền làm bài thi.")
_editor_required = _module_role_required("editor", "Không có quyền biên soạn nội dung.")
_exam_manager_required = _module_role_required("exam_manager", "Không có quyền quản lý kỳ thi.")
_question_reader_required = _module_role_required(
    ("editor", "exam_manager"), "Không có quyền xem ngân hàng câu hỏi."
)


def _attempt_owned_by_current_user(attempt_id):
    return attempts.attempt_owner_username(config.TRAINING_DB_PATH, attempt_id) == session.get("username")


@training_bp.route("/dao-tao-sat-hach")
def page_index():
    current_user = get_user_by_username(session.get("username"))
    roles = module_roles(config.TRAINING_DB_PATH, session.get("username"))
    if current_user and current_user.get("role") == "admin":
        roles.update({"learner", "editor", "exam_manager", "admin"})
    return render_template(
        "pages/training/index.html",
        current_user=current_user,
        module_roles=roles,
        active_page="training",
    )


@training_bp.route("/api/training/knowledge", methods=["GET", "POST"])
@_editor_required
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
@_editor_required
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


@training_bp.route("/api/training/questions/validate", methods=["POST"])
@csrf_protect
@_editor_required
def validate_questions():
    payload = request.get_json(silent=True)
    errors = questions.validate_question_batch(payload)
    if errors:
        return _error_response(TrainingError(
            "VALIDATION_ERROR", "Dữ liệu lô câu hỏi không hợp lệ.", status=400,
            details={"errors": errors},
        ))
    return jsonify({"valid": True, "errors": []})


@training_bp.route("/api/training/questions")
@_question_reader_required
def list_question_bank():
    try:
        return jsonify(questions.list_questions(
            config.TRAINING_DB_PATH,
            status=request.args.get("status"), audience=request.args.get("audience"),
            domain=request.args.get("domain"), topic=request.args.get("topic"),
            q=request.args.get("q"), page=max(1, request.args.get("page", 1, type=int)),
            page_size=min(100, max(1, request.args.get("page_size", 25, type=int))),
        ))
    except TrainingError as exc:
        return _error_response(exc)


@training_bp.route("/api/training/questions/<version_id>")
@_question_reader_required
def get_question_bank_detail(version_id):
    try:
        resp = jsonify(questions.get_question_management_detail(config.TRAINING_DB_PATH, version_id))
        resp.headers["Cache-Control"] = "no-store"
        return resp
    except TrainingError as exc:
        return _error_response(exc)


@training_bp.route("/api/training/questions/<version_id>/approve", methods=["POST"])
@csrf_protect
@_exam_manager_required
def approve_question(version_id):
    try:
        questions.add_review_action(config.TRAINING_DB_PATH, unit_code=config.UNIT_CODE,
                                    actor=session["username"], version_id=version_id, action="approve")
        return jsonify({"version_id": version_id, "review_status": "approved"})
    except TrainingError as exc:
        return _error_response(exc)


@training_bp.route("/api/training/questions/<version_id>/reject", methods=["POST"])
@csrf_protect
@_exam_manager_required
def reject_question(version_id):
    payload = request.get_json(silent=True) or {}
    comment = payload.get("comment")
    if comment is not None and not isinstance(comment, str):
        return _error_response(TrainingError(
            "VALIDATION_ERROR", "Nhận xét từ chối phải là chuỗi ký tự.", status=400,
        ))
    try:
        questions.add_review_action(
            config.TRAINING_DB_PATH, unit_code=config.UNIT_CODE,
            actor=session["username"], version_id=version_id, action="reject", comment=comment,
        )
        return jsonify({"version_id": version_id, "review_status": "rejected"})
    except TrainingError as exc:
        return _error_response(exc)


@training_bp.route("/api/training/questions/<version_id>/publish", methods=["POST"])
@csrf_protect
@_exam_manager_required
def publish_question(version_id):
    try:
        questions.publish_question_version(config.TRAINING_DB_PATH, unit_code=config.UNIT_CODE,
                                           actor=session["username"], version_id=version_id)
        return jsonify({"version_id": version_id, "publication_status": "published"})
    except TrainingError as exc:
        return _error_response(exc)


@training_bp.route("/api/training/templates", methods=["POST"])
@csrf_protect
@_exam_manager_required
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


@training_bp.route("/api/training/templates")
@_exam_manager_required
def list_templates():
    page = max(1, request.args.get("page", 1, type=int))
    page_size = min(100, max(1, request.args.get("page_size", 25, type=int)))
    try:
        return jsonify(exams.list_templates(
            config.TRAINING_DB_PATH, page=page, page_size=page_size,
        ))
    except TrainingError as exc:
        return _error_response(exc)


@training_bp.route("/api/training/templates/<template_id>")
@_exam_manager_required
def get_template_detail(template_id):
    try:
        return jsonify(exams.get_template_detail(config.TRAINING_DB_PATH, template_id))
    except TrainingError as exc:
        return _error_response(exc)


@training_bp.route("/api/training/exams", methods=["POST"])
@csrf_protect
@_exam_manager_required
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
            reveal_answers_after_finalize=bool(payload.get("reveal_answers_after_finalize", True)),
        )
        return jsonify(result), 201
    except (KeyError, ValueError) as exc:
        return jsonify({"error": {"code": "VALIDATION_ERROR", "message": str(exc), "details": {}}}), 400
    except TrainingError as exc:
        return _error_response(exc)


@training_bp.route("/api/training/exams/<exam_id>/assignments", methods=["POST"])
@csrf_protect
@_exam_manager_required
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


@training_bp.route("/api/training/my-assignments")
@_learner_required
@csrf_protect
def get_my_assignments():
    username = session.get("username")
    result = exams.get_my_assignments_dto(config.TRAINING_DB_PATH, username)
    return jsonify(result)


@training_bp.route("/api/training/exams")
@_exam_manager_required
def list_exams():
    page = max(1, request.args.get("page", 1, type=int))
    page_size = min(100, max(1, request.args.get("page_size", 25, type=int)))
    try:
        return jsonify(exams.list_exams(
            config.TRAINING_DB_PATH, page=page, page_size=page_size,
            status=request.args.get("status"),
        ))
    except TrainingError as exc:
        return _error_response(exc)


@training_bp.route("/api/training/exams/<exam_id>")
@_exam_manager_required
def get_exam_detail(exam_id):
    try:
        resp = jsonify(exams.get_exam_detail(config.TRAINING_DB_PATH, exam_id))
        add_no_cache_headers(resp)
        resp.headers["Cache-Control"] = "no-store"
        return resp
    except TrainingError as exc:
        return _error_response(exc)


@training_bp.route("/api/training/exams/<exam_id>/assignments")
@_exam_manager_required
def list_exam_assignments(exam_id):
    try:
        return jsonify({
            "items": exams.get_exam_assignments_dto(config.TRAINING_DB_PATH, exam_id),
        })
    except TrainingError as exc:
        return _error_response(exc)


@training_bp.route("/api/training/users")
@_exam_manager_required
def list_assignable_users():
    try:
        return jsonify({
            "items": exams.list_assignable_users(
                config.TRAINING_DB_PATH, q=request.args.get("q", ""),
            ),
        })
    except TrainingError as exc:
        return _error_response(exc)


@training_bp.route("/api/training/exams/<exam_id>/<action>", methods=["POST"])
@csrf_protect
@_exam_manager_required
def transition_exam(exam_id, action):
    handlers = {
        "ready": exams.ready_exam,
        "open": exams.open_exam,
        "close": exams.close_exam,
        "cancel": exams.cancel_exam,
    }
    handler = handlers.get(action)
    if handler is None:
        return jsonify({"error": {"code": "NOT_FOUND", "message": "Thao tác không tồn tại.", "details": {}}}), 404
    try:
        result = handler(
            config.TRAINING_DB_PATH, unit_code=config.UNIT_CODE,
            actor=session["username"], exam_id=exam_id,
        )
        response = {"exam_id": exam_id, "action": action}
        if action == "close":
            response["recovery_summary"] = result
        return jsonify(response)
    except TrainingError as exc:
        return _error_response(exc)


@training_bp.route("/api/training/assignments/<assignment_id>/attempts", methods=["POST"])
@csrf_protect
@_learner_required
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
@_learner_required
def get_attempt(attempt_id):
    try:
        if not _attempt_owned_by_current_user(attempt_id):
            raise TrainingError("PERMISSION_SCOPE_DENIED", "Bạn không sở hữu bài làm này.", status=403)
        view = attempts.get_attempt_learner_view(config.TRAINING_DB_PATH, attempt_id)
        response = jsonify(view)
        add_no_cache_headers(response)
        response.headers["Cache-Control"] = "no-store"
        return response
    except TrainingError as exc:
        return _error_response(exc)


@training_bp.route("/api/training/attempts/<attempt_id>/responses/<item_id>", methods=["PUT"])
@csrf_protect
@_learner_required
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
@_learner_required
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
@_exam_manager_required
def finalize_exam(exam_id):
    try:
        return jsonify(reports.finalize_exam(
            config.TRAINING_DB_PATH, unit_code=config.UNIT_CODE,
            actor=session["username"], exam_id=exam_id,
        ))
    except TrainingError as exc:
        return _error_response(exc)


@training_bp.route("/api/training/exams/<exam_id>/report")
@_exam_manager_required
def get_report(exam_id):
    try:
        report = reports.get_report_snapshot(config.TRAINING_DB_PATH, exam_id)
        if report is None:
            return jsonify({"error": {"code": "NOT_FOUND", "message": "Kỳ thi chưa được chốt.", "details": {}}}), 404
        return jsonify(report)
    except TrainingError as exc:
        return _error_response(exc)


@training_bp.route("/download/training/exams/<exam_id>/report.xlsx")
@_exam_manager_required
def download_report_excel(exam_id):
    report = reports.get_report_snapshot(config.TRAINING_DB_PATH, exam_id)
    if report is None:
        return jsonify({"error": {"code": "NOT_FOUND", "message": "Kỳ thi chưa được chốt.", "details": {}}}), 404
    from pathlib import Path

    filename = f"exam-report-{uuid.uuid4().hex}.xlsx"
    path = Path(config.TRAINING_EXPORT_DIR) / filename
    reports.export_report_excel(report["payload"], path)
    return send_file(path, as_attachment=True, download_name=f"bao-cao-{exam_id}.xlsx")
