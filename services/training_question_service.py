"""Question service: validate Question Batch JSON, semantic checks, duplicate detection.

Validate Draft 2020-12; object chính additionalProperties: false.
Backend xác minh quote span, offset và content hash trước publish.
Exact normalized hash chặn publish; token/Jaccard similarity chỉ cảnh báo (sau MVP).
"""

import hashlib
import json
import os
import re
import unicodedata

from jsonschema import Draft202012Validator

from training import constants, time_policy
from training.db import read_connection, write_connection
from training.errors import ErrorCode, TrainingError
from repositories.training_repository import gen_id, write_audit

_SCHEMA_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "training", "schemas")

_question_batch_validator = None
_normalized_stem_hashes = {}


def _load_validator(name):
    schema_path = os.path.join(_SCHEMA_DIR, name)
    with open(schema_path, "r", encoding="utf-8") as fh:
        schema = json.load(fh)
    return Draft202012Validator(schema)


def _get_batch_validator():
    global _question_batch_validator
    if _question_batch_validator is None:
        _question_batch_validator = _load_validator("question_batch.schema.json")
    return _question_batch_validator


def _normalize_stem(text):
    """Normalize: lowercase, strip accents/diacritics, collapse whitespace."""
    if not text:
        return ""
    nfkd = unicodedata.normalize("NFKD", text)
    stripped = "".join(c for c in nfkd if not unicodedata.combining(c))
    lowered = stripped.lower()
    return re.sub(r"\s+", " ", lowered).strip()


def reset_duplicate_cache():
    global _normalized_stem_hashes
    _normalized_stem_hashes = {}


def record_normalized_stem(stem_text, version_id):
    h = hashlib.sha256(_normalize_stem(stem_text).encode("utf-8")).hexdigest()
    _normalized_stem_hashes[h] = version_id


def is_duplicate_normalized_stem(stem_text):
    h = hashlib.sha256(_normalize_stem(stem_text).encode("utf-8")).hexdigest()
    return h in _normalized_stem_hashes


def validate_question_batch(payload):
    """Trả list of error strings. Rỗng = hợp lệ."""
    errors = []
    validator = _get_batch_validator()
    schema_errors = sorted(validator.iter_errors(payload), key=lambda e: list(e.path))
    for err in schema_errors:
        path = ".".join(str(p) for p in err.path) or "(root)"
        errors.append(f"schema: {path}: {err.message}")

    questions = payload.get("questions", []) if isinstance(payload, dict) else []
    source_ids = set()
    if isinstance(payload, dict):
        batch = payload.get("batch", {})
        source_ids = set(batch.get("source_document_version_ids", []))

    for qi, q in enumerate(questions):
        prefix = f"question[{qi}]"
        options = q.get("options", [])
        option_ids = [o.get("id") for o in options if isinstance(o, dict)]
        seen = set()
        for oid in option_ids:
            if oid in seen:
                errors.append(f"{prefix}: duplicate option id {oid!r}")
            seen.add(oid)

        correct = q.get("correct_option_ids", [])
        if q.get("type") == "single_choice":
            if len(correct) != 1:
                errors.append(
                    f"{prefix}: single_choice requires exactly one correct_option_id, got {len(correct)}"
                )
        for cid in correct:
            if cid not in option_ids:
                errors.append(f"{prefix}: correct option {cid!r} does not exist in options")

        evidence = q.get("evidence", [])
        if not evidence:
            errors.append(f"{prefix}: at least one evidence required")
        for ei, ev in enumerate(evidence):
            ev_source = ev.get("document_version_id")
            if source_ids and ev_source not in source_ids:
                errors.append(
                    f"{prefix}: evidence[{ei}] document_version_id {ev_source!r} "
                    "not in batch source_document_version_ids"
                )

    return errors


def parse_question_batch_json(text):
    """Parse + validate. Trả (payload, errors)."""
    try:
        payload = json.loads(text)
    except (json.JSONDecodeError, TypeError) as exc:
        return None, [f"json parse error: {exc}"]
    errors = validate_question_batch(payload)
    if errors:
        return None, errors
    return payload, []


# --- DB-level question bank operations ---

def _stem_hash(stem):
    return hashlib.sha256(_normalize_stem(stem).encode("utf-8")).hexdigest()


def import_question_batch(db_path, *, unit_code, actor, batch, status="draft", conn=None):
    """Import validated batch → tạo question_items + versions + options + sources.

    Completed job chỉ tạo drafts, không publish.
    """
    errors = validate_question_batch(batch)
    if errors:
        raise TrainingError(
            ErrorCode.VALIDATION_ERROR, "; ".join(errors[:5]), status=400,
            details={"errors": errors},
        )
    now = time_policy.utc_now_ms()
    version_ids = []
    owns_connection = conn is None
    if owns_connection:
        conn = write_connection(db_path)
    try:
        for q in batch.get("questions", []):
            item_id = gen_id("qi")
            version_id = gen_id("qv")
            conn.execute(
                "INSERT INTO question_items (id, current_version_id, created_at_ms) VALUES (?, ?, ?)",
                (item_id, version_id, now),
            )
            classification = q.get("classification", {})
            conn.execute(
                """INSERT INTO question_versions
                (id, question_item_id, version_number, type, stem, stimulus, language,
                 correct_option_ids_json, explanation, distractor_rationales_json,
                 difficulty, cognitive_level, criticality, estimated_seconds,
                 max_score, normalized_stem_hash, review_status, publication_status,
                 created_by, created_at_ms)
                VALUES (?, ?, 1, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'unpublished', ?, ?)""",
                (
                    version_id, item_id, q["type"], q["stem"], q.get("stimulus"),
                    q.get("language", "vi"),
                    json.dumps(q["correct_option_ids"], ensure_ascii=False),
                    q.get("explanation"),
                    json.dumps(q.get("distractor_rationales") or {}, ensure_ascii=False),
                    q["difficulty"], q.get("cognitive_level"), q.get("criticality"),
                    q.get("estimated_seconds"), 1.0, _stem_hash(q["stem"]),
                    status, actor, now,
                ),
            )
            for idx, opt in enumerate(q.get("options", [])):
                conn.execute(
                    "INSERT INTO question_options (id, question_version_id, option_code, option_text, display_order) "
                    "VALUES (?, ?, ?, ?, ?)",
                    (gen_id("opt"), version_id, opt["id"], opt["text"], idx),
                )
            for ev in q.get("evidence", []):
                conn.execute(
                    "INSERT INTO question_sources "
                    "(id, question_version_id, document_version_id, block_id, extraction_revision, "
                    "quoted_text, quote_start, quote_end, supports) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
                    (gen_id("qsrc"), version_id, ev["document_version_id"], ev["block_id"],
                     ev.get("extraction_revision", 1), ev.get("quoted_text"),
                     ev.get("quote_start"), ev.get("quote_end"), ev.get("supports")),
                )
            for tc in classification.get("topic_codes", []):
                conn.execute(
                    "INSERT OR IGNORE INTO question_topics (question_version_id, topic_code) VALUES (?, ?)",
                    (version_id, tc),
                )
            for ac in classification.get("audience_codes", []):
                conn.execute(
                    "INSERT OR IGNORE INTO question_audiences (question_version_id, audience_code) VALUES (?, ?)",
                    (version_id, ac),
                )
            for ic in classification.get("indicator_codes", []):
                conn.execute(
                    "INSERT OR IGNORE INTO question_indicators (question_version_id, indicator_code) VALUES (?, ?)",
                    (version_id, ic),
                )
            version_ids.append(version_id)
            record_normalized_stem(q["stem"], version_id)
        write_audit(conn, actor=actor, unit_code=unit_code, action="import_questions",
                    entity_type="question_batch", entity_id=actor,
                    after={"count": len(version_ids), "status": status})
        if owns_connection:
            conn.commit()
    finally:
        if owns_connection:
            conn.close()
    return {"version_ids": version_ids}


def get_question_version(db_path, version_id):
    conn = read_connection(db_path)
    try:
        row = conn.execute(
            "SELECT * FROM question_versions WHERE id=?", (version_id,)
        ).fetchone()
        if not row:
            return None
        result = dict(row)
        result["correct_option_ids"] = json.loads(result.pop("correct_option_ids_json", "[]"))
        result["distractor_rationales"] = json.loads(result.pop("distractor_rationales_json", "{}"))
        opts = conn.execute(
            "SELECT option_code, option_text, display_order FROM question_options "
            "WHERE question_version_id=? ORDER BY display_order",
            (version_id,),
        ).fetchall()
        result["options"] = [dict(o) for o in opts]
        return result
    finally:
        conn.close()


def _question_status(review_status, publication_status):
    if publication_status == constants.PublicationStatus.PUBLISHED:
        return constants.PublicationStatus.PUBLISHED
    if review_status == constants.QuestionReviewStatus.APPROVED:
        return constants.QuestionReviewStatus.APPROVED
    if review_status == constants.QuestionReviewStatus.REJECTED:
        return constants.QuestionReviewStatus.REJECTED
    return constants.QuestionReviewStatus.DRAFT


def _question_classification(conn, version_id):
    domains = conn.execute(
        """SELECT DISTINCT kb.domain_code FROM question_sources qs
           JOIN knowledge_blocks kb ON kb.document_version_id=qs.document_version_id
             AND kb.extraction_revision=qs.extraction_revision AND kb.block_id=qs.block_id
           WHERE qs.question_version_id=? AND kb.domain_code IS NOT NULL
           ORDER BY kb.domain_code""",
        (version_id,),
    ).fetchall()
    mappings = {}
    for key, table, column in (
        ("topic_codes", "question_topics", "topic_code"),
        ("audience_codes", "question_audiences", "audience_code"),
        ("indicator_codes", "question_indicators", "indicator_code"),
    ):
        mappings[key] = [row[column] for row in conn.execute(
            f"SELECT {column} FROM {table} WHERE question_version_id=? ORDER BY {column}",
            (version_id,),
        ).fetchall()]
    return {"domain_codes": [row["domain_code"] for row in domains], **mappings}


def list_questions(
    db_path, *, status=None, audience=None, domain=None, topic=None,
    domain_code=None, page=1, page_size=25, q=None,
):
    """Return paginated, non-sensitive management list DTOs."""
    if domain is None:
        domain = domain_code
    if status not in (None, "draft", "approved", "published"):
        raise TrainingError(ErrorCode.VALIDATION_ERROR, "Trạng thái lọc không hợp lệ.", status=400)
    conn = read_connection(db_path)
    try:
        where = []
        params = []
        if status:
            if status == "draft":
                where.append("v.publication_status='unpublished' AND v.review_status IN ('draft', 'needs_review')")
            elif status == "approved":
                where.append("v.publication_status='unpublished' AND v.review_status='approved'")
            else:
                where.append("v.publication_status='published'")
        if q:
            where.append("v.stem LIKE ?")
            params.append(f"%{q}%")
        if audience:
            where.append(
                "EXISTS (SELECT 1 FROM question_audiences qa "
                "WHERE qa.question_version_id=v.id AND qa.audience_code=?)"
            )
            params.append(audience)
        if topic:
            where.append(
                "EXISTS (SELECT 1 FROM question_topics qt "
                "WHERE qt.question_version_id=v.id AND qt.topic_code=?)"
            )
            params.append(topic)
        if domain:
            where.append(
                """EXISTS (SELECT 1 FROM question_sources qs
                   JOIN knowledge_blocks kb ON kb.document_version_id=qs.document_version_id
                     AND kb.extraction_revision=qs.extraction_revision AND kb.block_id=qs.block_id
                   WHERE qs.question_version_id=v.id AND kb.domain_code=?)"""
            )
            params.append(domain)
        clause = ("WHERE " + " AND ".join(where)) if where else ""
        total = conn.execute(
            f"SELECT COUNT(*) AS c FROM question_versions v {clause}", params
        ).fetchone()["c"]
        offset = (page - 1) * page_size
        rows = conn.execute(
            f"""SELECT v.id, v.stem, v.type, v.difficulty, v.review_status,
                       v.publication_status, v.version_number, v.question_item_id
                FROM question_versions v {clause}
                ORDER BY v.created_at_ms DESC LIMIT ? OFFSET ?""",
            params + [page_size, offset],
        ).fetchall()
        items = []
        for row in rows:
            classification = _question_classification(conn, row["id"])
            items.append({
                "id": row["id"], "stem": row["stem"], "type": row["type"],
                "difficulty": row["difficulty"],
                "audience": classification["audience_codes"],
                "topic": classification["topic_codes"],
                "status": _question_status(row["review_status"], row["publication_status"]),
                "version": row["version_number"],
            })
        return {"items": items, "page": page,
                "page_size": page_size, "total": total}
    finally:
        conn.close()


def get_question_management_detail(db_path, version_id):
    """Return the management-only detail DTO for a question version."""
    conn = read_connection(db_path)
    try:
        row = conn.execute(
            """SELECT id, question_item_id, version_number, type, stem, stimulus, language,
                      correct_option_ids_json, explanation, distractor_rationales_json, difficulty,
                      cognitive_level, criticality, estimated_seconds, review_status,
                      publication_status, created_by, created_at_ms, approved_by, approved_at_ms
               FROM question_versions WHERE id=?""",
            (version_id,),
        ).fetchone()
        if not row:
            raise TrainingError(ErrorCode.NOT_FOUND, "Không tìm thấy câu hỏi", status=404)
        options = [
            {"id": option["option_code"], "text": option["option_text"], "order": option["display_order"]}
            for option in conn.execute(
                "SELECT option_code, option_text, display_order FROM question_options "
                "WHERE question_version_id=? ORDER BY display_order", (version_id,)
            ).fetchall()
        ]
        evidence = [dict(source) for source in conn.execute(
            """SELECT document_version_id, block_id, extraction_revision, quoted_text,
                      quote_start, quote_end, supports FROM question_sources
               WHERE question_version_id=? ORDER BY id""",
            (version_id,),
        ).fetchall()]
        reviews = [dict(review) for review in conn.execute(
            """SELECT action, reviewer, comment, created_at_ms FROM question_reviews
               WHERE question_version_id=? ORDER BY created_at_ms""",
            (version_id,),
        ).fetchall()]
        publications = [dict(event) for event in conn.execute(
            """SELECT action, actor, created_at_ms FROM training_audit_log
               WHERE entity_type='question_version' AND entity_id=? AND action='publish'
               ORDER BY created_at_ms""",
            (version_id,),
        ).fetchall()]
        return {
            "id": row["id"], "question_item_id": row["question_item_id"],
            "version": row["version_number"], "type": row["type"], "stem": row["stem"],
            "stimulus": row["stimulus"], "language": row["language"], "options": options,
            "correct_option_ids": json.loads(row["correct_option_ids_json"]),
            "explanation": row["explanation"],
            "distractor_rationales": json.loads(row["distractor_rationales_json"] or "{}"),
            "classification": _question_classification(conn, version_id),
            "difficulty": row["difficulty"], "cognitive_level": row["cognitive_level"],
            "criticality": row["criticality"], "estimated_seconds": row["estimated_seconds"],
            "evidence": evidence, "review_status": row["review_status"],
            "review_history": reviews,
            "publication": {
                "status": row["publication_status"], "approved_by": row["approved_by"],
                "approved_at_ms": row["approved_at_ms"],
            },
            "publication_history": publications, "created_by": row["created_by"],
            "created_at_ms": row["created_at_ms"],
        }
    finally:
        conn.close()


def add_review_action(db_path, *, unit_code, actor, version_id, action, comment=None):
    now = time_policy.utc_now_ms()
    conn = write_connection(db_path)
    try:
        ver = conn.execute(
            "SELECT review_status, publication_status FROM question_versions WHERE id=?",
            (version_id,),
        ).fetchone()
        if not ver:
            raise TrainingError(ErrorCode.NOT_FOUND, "Không tìm thấy câu hỏi", status=404)
        new_review = ver["review_status"]
        if action == "approve":
            new_review = constants.QuestionReviewStatus.APPROVED
        elif action == "reject":
            new_review = constants.QuestionReviewStatus.REJECTED
        elif action == "request_review":
            new_review = constants.QuestionReviewStatus.NEEDS_REVIEW
        conn.execute(
            "UPDATE question_versions SET review_status=? WHERE id=?",
            (new_review, version_id),
        )
        if new_review == constants.QuestionReviewStatus.APPROVED:
            conn.execute(
                "UPDATE question_versions SET approved_by=?, approved_at_ms=? WHERE id=?",
                (actor, now, version_id),
            )
        conn.execute(
            "INSERT INTO question_reviews (id, question_version_id, action, reviewer, comment, created_at_ms) "
            "VALUES (?, ?, ?, ?, ?, ?)",
            (gen_id("rev"), version_id, action, actor, comment, now),
        )
        write_audit(conn, actor=actor, unit_code=unit_code, action=f"review_{action}",
                    entity_type="question_version", entity_id=version_id)
        conn.commit()
    finally:
        conn.close()


def publish_question_version(db_path, *, unit_code, actor, version_id):
    """Publish: approved → published. Chặn duplicate, chặn thiếu evidence."""
    conn = write_connection(db_path)
    try:
        ver = conn.execute(
            "SELECT review_status, publication_status, normalized_stem_hash, stem FROM question_versions WHERE id=?",
            (version_id,),
        ).fetchone()
        if not ver:
            raise TrainingError(ErrorCode.NOT_FOUND, "Không tìm thấy câu hỏi", status=404)
        if ver["review_status"] != constants.QuestionReviewStatus.APPROVED:
            raise TrainingError(
                ErrorCode.CONFLICT,
                "Chỉ phát hành câu hỏi đã được duyệt (approved).",
                status=409,
            )
        existing = conn.execute(
            "SELECT id FROM question_versions WHERE normalized_stem_hash=? AND publication_status='published' AND id!=?",
            (ver["normalized_stem_hash"], version_id),
        ).fetchone()
        if existing:
            raise TrainingError(
                ErrorCode.CONFLICT,
                "Câu hỏi trùng (duplicate normalized stem) đã được phát hành.",
                status=409,
            )
        src_count = conn.execute(
            "SELECT COUNT(*) AS c FROM question_sources WHERE question_version_id=?",
            (version_id,),
        ).fetchone()["c"]
        if src_count == 0:
            raise TrainingError(
                ErrorCode.CONFLICT,
                "Không thể phát hành câu hỏi thiếu evidence.",
                status=409,
            )
        conn.execute(
            "UPDATE question_versions SET publication_status=? WHERE id=?",
            (constants.PublicationStatus.PUBLISHED, version_id),
        )
        conn.execute(
            "UPDATE question_items SET current_version_id=? WHERE current_version_id=?",
            (version_id, version_id),
        )
        write_audit(conn, actor=actor, unit_code=unit_code, action="publish",
                    entity_type="question_version", entity_id=version_id)
        conn.commit()
    finally:
        conn.close()


def update_question_draft(db_path, *, unit_code, actor, version_id, expected_version, updates):
    """PATCH draft với expected_version. Conflict → VERSION_CONFLICT."""
    conn = write_connection(db_path)
    try:
        ver = conn.execute(
            "SELECT version_number, review_status, publication_status FROM question_versions WHERE id=?",
            (version_id,),
        ).fetchone()
        if not ver:
            raise TrainingError(ErrorCode.NOT_FOUND, "Không tìm thấy câu hỏi", status=404)
        if ver["version_number"] != expected_version:
            raise TrainingError(ErrorCode.VERSION_CONFLICT,
                                "Version không khớp.", status=409)
        if ver["publication_status"] == constants.PublicationStatus.PUBLISHED:
            raise TrainingError(ErrorCode.CONFLICT,
                                "Câu hỏi đã phát hành không được sửa trực tiếp.",
                                status=409)
        allowed = {"stem", "stimulus", "explanation", "difficulty",
                   "cognitive_level", "criticality", "estimated_seconds"}
        set_clauses = []
        params = []
        for key, value in updates.items():
            if key in allowed:
                set_clauses.append(f"{key} = ?")
                params.append(value)
        if set_clauses:
            if "stem" in updates:
                set_clauses.append("normalized_stem_hash = ?")
                params.append(_stem_hash(updates["stem"]))
            params.extend([version_id])
            conn.execute(
                f"UPDATE question_versions SET {', '.join(set_clauses)} WHERE id=?",
                params,
            )
        write_audit(conn, actor=actor, unit_code=unit_code, action="update_draft",
                    entity_type="question_version", entity_id=version_id,
                    after={"fields": list(updates.keys())})
        conn.commit()
    finally:
        conn.close()
