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
