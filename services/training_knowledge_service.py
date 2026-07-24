"""Knowledge service: paste-text ingestion, versioning, blocks, issues.

Không sửa tài liệu gốc qua kết quả AI. Thay nội dung đã duyệt tạo version mới.
Block ID ổn định trong (document_version_id, extraction_revision).
"""

import hashlib
import json

from training import constants, time_policy
from training.db import read_connection, write_connection
from repositories.training_repository import gen_id, write_audit

BLOCK_SEPARATOR = "\n\n"


def _sha256(text):
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def _split_blocks(content_text):
    """Chia text thành blocks theo đoạn (double newline). Trả list of (content, start, end)."""
    blocks = []
    pos = 0
    for chunk in content_text.split(BLOCK_SEPARATOR):
        chunk_stripped = chunk.strip()
        if chunk_stripped:
            start = content_text.find(chunk_stripped, pos)
            if start < 0:
                start = pos
            end = start + len(chunk_stripped)
            blocks.append((chunk_stripped, start, end))
            pos = end
        else:
            pos += len(BLOCK_SEPARATOR)
    if not blocks and content_text.strip():
        blocks.append((content_text.strip(), 0, len(content_text.strip())))
    return blocks


def create_document(
    db_path, *, unit_code, actor, document_code, title, content_text,
    classification=None, audience_codes=None, document_type="kpi_definition",
    issuer=None,
):
    """Tạo document + version 1 + blocks. Trả dict {document_id, version_id, version_number}."""
    classification = classification or {}
    audience_codes = audience_codes or []
    now = time_policy.utc_now_ms()
    doc_id = gen_id("doc")
    version_id = gen_id("docver")
    conn = write_connection(db_path)
    try:
        conn.execute(
            "INSERT INTO knowledge_documents "
            "(id, document_code, title, document_type, issuer, review_status, created_at_ms, updated_at_ms) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
            (doc_id, document_code, title, document_type, issuer,
             constants.DocumentReviewStatus.DRAFT, now, now),
        )
        conn.execute(
            "INSERT INTO knowledge_document_versions "
            "(id, document_id, version_number, content_text, content_sha256, review_status, created_at_ms, created_by) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
            (version_id, doc_id, 1, content_text, _sha256(content_text),
             constants.DocumentReviewStatus.DRAFT, now, actor),
        )
        _insert_version_mappings(conn, version_id, classification, audience_codes)
        _insert_blocks(conn, version_id, 1, content_text, classification, now)
        write_audit(conn, actor=actor, unit_code=unit_code, action="create_document",
                    entity_type="knowledge_document", entity_id=doc_id,
                    after={"document_code": document_code, "title": title})
        conn.commit()
    finally:
        conn.close()
    return {"document_id": doc_id, "version_id": version_id, "version_number": 1}


def _insert_version_mappings(conn, version_id, classification, audience_codes):
    domain_code = classification.get("domain_code")
    if domain_code:
        conn.execute(
            "UPDATE knowledge_document_versions SET review_status=review_status WHERE id=?",
            (version_id,),
        )
    for topic_code in classification.get("topic_codes", []):
        conn.execute(
            "INSERT OR IGNORE INTO knowledge_document_topics (document_version_id, topic_code) VALUES (?, ?)",
            (version_id, topic_code),
        )
    for audience_code in audience_codes:
        conn.execute(
            "INSERT OR IGNORE INTO knowledge_document_audiences (document_version_id, audience_code) VALUES (?, ?)",
            (version_id, audience_code),
        )


def _insert_blocks(conn, version_id, extraction_revision, content_text, classification, now_ms):
    domain_code = classification.get("domain_code")
    category_code = classification.get("category_code")
    doc_code = classification.get("document_code", "DOC")
    for index, (content, start, end) in enumerate(_split_blocks(content_text), start=1):
        block_id = f"{doc_code}-B{index:03d}"
        conn.execute(
            "INSERT INTO knowledge_blocks "
            "(id, document_version_id, extraction_revision, block_id, char_start, char_end, "
            "content, content_sha256, domain_code, category_code, created_at_ms) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
            (gen_id("blk"), version_id, extraction_revision, block_id,
             start, end, content, _sha256(content), domain_code, category_code, now_ms),
        )


def create_version(
    db_path, *, unit_code, actor, document_id, content_text,
    classification=None, audience_codes=None,
):
    """Tạo version mới cho document đã có."""
    classification = classification or {}
    audience_codes = audience_codes or []
    now = time_policy.utc_now_ms()
    conn = write_connection(db_path)
    try:
        row = conn.execute(
            "SELECT document_code FROM knowledge_documents WHERE id=?", (document_id,)
        ).fetchone()
        if not row:
            raise ValueError(f"document {document_id} không tồn tại")
        doc_code = row["document_code"]
        max_ver = conn.execute(
            "SELECT COALESCE(MAX(version_number), 0) AS m FROM knowledge_document_versions WHERE document_id=?",
            (document_id,),
        ).fetchone()["m"]
        version_number = max_ver + 1
        version_id = gen_id("docver")
        conn.execute(
            "INSERT INTO knowledge_document_versions "
            "(id, document_id, version_number, content_text, content_sha256, review_status, created_at_ms, created_by) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
            (version_id, document_id, version_number, content_text, _sha256(content_text),
             constants.DocumentReviewStatus.DRAFT, now, actor),
        )
        classification_with_code = dict(classification)
        classification_with_code["document_code"] = doc_code
        _insert_version_mappings(conn, version_id, classification_with_code, audience_codes)
        _insert_blocks(conn, version_id, 1, content_text, classification_with_code, now)
        conn.execute(
            "UPDATE knowledge_documents SET updated_at_ms=? WHERE id=?", (now, document_id)
        )
        write_audit(conn, actor=actor, unit_code=unit_code, action="create_version",
                    entity_type="knowledge_document_version", entity_id=version_id,
                    after={"document_id": document_id, "version_number": version_number})
        conn.commit()
    finally:
        conn.close()
    return {"document_id": document_id, "version_id": version_id, "version_number": version_number}


def get_version(db_path, version_id):
    conn = read_connection(db_path)
    try:
        row = conn.execute(
            "SELECT * FROM knowledge_document_versions WHERE id=?", (version_id,)
        ).fetchone()
        return dict(row) if row else None
    finally:
        conn.close()


def get_extraction_revision(db_path, version_id):
    conn = read_connection(db_path)
    try:
        row = conn.execute(
            "SELECT COALESCE(MAX(extraction_revision), 0) AS m "
            "FROM knowledge_blocks WHERE document_version_id=?",
            (version_id,),
        ).fetchone()
        return row["m"]
    finally:
        conn.close()


def reprocess_blocks(db_path, *, unit_code, actor, version_id, content_text):
    """Re-chunk tạo extraction revision mới. Không ghi đè revision cũ."""
    now = time_policy.utc_now_ms()
    conn = write_connection(db_path)
    try:
        doc_ver = conn.execute(
            "SELECT * FROM knowledge_document_versions WHERE id=?", (version_id,)
        ).fetchone()
        if not doc_ver:
            raise ValueError(f"version {version_id} không tồn tại")
        doc_row = conn.execute(
            "SELECT document_code FROM knowledge_documents WHERE id=?",
            (doc_ver["document_id"],),
        ).fetchone()
        doc_code = doc_row["document_code"] if doc_row else "DOC"
        current_rev = conn.execute(
            "SELECT COALESCE(MAX(extraction_revision), 0) AS m "
            "FROM knowledge_blocks WHERE document_version_id=?",
            (version_id,),
        ).fetchone()["m"]
        new_rev = current_rev + 1
        classification = {"document_code": doc_code,
                          "domain_code": None, "category_code": None}
        _insert_blocks(conn, version_id, new_rev, content_text, classification, now)
        write_audit(conn, actor=actor, unit_code=unit_code, action="reprocess_blocks",
                    entity_type="knowledge_blocks", entity_id=version_id,
                    after={"new_extraction_revision": new_rev})
        conn.commit()
        return new_rev
    finally:
        conn.close()


def list_blocks(db_path, version_id, extraction_revision=None):
    conn = read_connection(db_path)
    try:
        if extraction_revision is None:
            extraction_revision = get_extraction_revision(db_path, version_id)
        rows = conn.execute(
            "SELECT * FROM knowledge_blocks "
            "WHERE document_version_id=? AND extraction_revision=? ORDER BY block_id",
            (version_id, extraction_revision),
        ).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def create_issue(db_path, *, unit_code, actor, version_id, severity, issue_type,
                 description, block_id=None):
    now = time_policy.utc_now_ms()
    issue_id = gen_id("issue")
    conn = write_connection(db_path)
    try:
        conn.execute(
            "INSERT INTO knowledge_issues "
            "(id, document_version_id, block_id, severity, issue_type, description, status, created_at_ms) "
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
            (issue_id, version_id, block_id, severity, issue_type, description,
             constants.IssueStatus.OPEN, now),
        )
        write_audit(conn, actor=actor, unit_code=unit_code, action="create_issue",
                    entity_type="knowledge_issue", entity_id=issue_id)
        conn.commit()
    finally:
        conn.close()
    return issue_id


def list_issues(db_path, version_id):
    conn = read_connection(db_path)
    try:
        rows = conn.execute(
            "SELECT * FROM knowledge_issues WHERE document_version_id=? ORDER BY created_at_ms",
            (version_id,),
        ).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def resolve_issue(db_path, *, unit_code, actor, issue_id, status):
    now = time_policy.utc_now_ms()
    conn = write_connection(db_path)
    try:
        conn.execute(
            "UPDATE knowledge_issues SET status=?, resolved_by=?, resolved_at_ms=? WHERE id=?",
            (status, actor, now, issue_id),
        )
        write_audit(conn, actor=actor, unit_code=unit_code, action="resolve_issue",
                    entity_type="knowledge_issue", entity_id=issue_id, after={"status": status})
        conn.commit()
    finally:
        conn.close()


def has_blocking_issues(db_path, version_id):
    """True nếu có issue severity=high và status=open."""
    conn = read_connection(db_path)
    try:
        row = conn.execute(
            "SELECT COUNT(*) AS c FROM knowledge_issues "
            "WHERE document_version_id=? AND severity='high' AND status='open'",
            (version_id,),
        ).fetchone()
        return row["c"] > 0
    finally:
        conn.close()


def list_documents(db_path, *, domain_code=None, status=None, page=1, page_size=25):
    conn = read_connection(db_path)
    try:
        where = []
        params = []
        if domain_code:
            where.append("d.review_status = d.review_status")
        if status:
            where.append("d.review_status = ?")
            params.append(status)
        clause = ("WHERE " + " AND ".join(where)) if where else ""
        total = conn.execute(
            f"SELECT COUNT(*) AS c FROM knowledge_documents d {clause}", params
        ).fetchone()["c"]
        offset = (page - 1) * page_size
        rows = conn.execute(
            f"SELECT d.* FROM knowledge_documents d {clause} "
            "ORDER BY d.updated_at_ms DESC LIMIT ? OFFSET ?",
            params + [page_size, offset],
        ).fetchall()
        return {"items": [dict(r) for r in rows], "page": page,
                "page_size": page_size, "total": total}
    finally:
        conn.close()


def get_document(db_path, document_id):
    conn = read_connection(db_path)
    try:
        doc = conn.execute(
            "SELECT * FROM knowledge_documents WHERE id=?", (document_id,)
        ).fetchone()
        if not doc:
            return None
        versions = conn.execute(
            "SELECT id, version_number, review_status, created_at_ms, created_by "
            "FROM knowledge_document_versions WHERE document_id=? ORDER BY version_number",
            (document_id,),
        ).fetchall()
        return {"document": dict(doc), "versions": [dict(v) for v in versions]}
    finally:
        conn.close()
