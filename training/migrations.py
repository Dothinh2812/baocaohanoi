"""Schema migration runner cho training.db.

Có version, chạy trong transaction, dùng file lock per-instance để nhiều
Gunicorn worker không migrate đồng thời. Kiểm tra unit_code trước khi thay schema.
"""

import fcntl
import os
import sqlite3
import time

from training import time_policy
from training.db import write_connection

CURRENT_SCHEMA_VERSION = 8


class UnitCodeMismatchError(Exception):
    pass


def _lock_path(db_path):
    return os.path.abspath(db_path) + ".migrate.lock"


def _acquire_lock(db_path):
    lock_file = open(_lock_path(db_path), "w")
    fcntl.flock(lock_file.fileno(), fcntl.LOCK_EX)
    return lock_file


def _table_exists(conn, name):
    row = conn.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name=?", (name,)
    ).fetchone()
    return row is not None


def _applied_versions(conn):
    if not _table_exists(conn, "training_schema_migrations"):
        return set()
    return {
        row["version"]
        for row in conn.execute(
            "SELECT version FROM training_schema_migrations"
        ).fetchall()
    }


def _check_or_init_instance_metadata(conn, unit_code):
    """Kiểm tra unit_code trên DB đã có metadata; trả về True nếu đã có sẵn."""
    if not _table_exists(conn, "training_instance_metadata"):
        return False
    row = conn.execute(
        "SELECT unit_code FROM training_instance_metadata"
    ).fetchone()
    if row and row["unit_code"] != unit_code:
        raise UnitCodeMismatchError(
            f"training.db unit_code={row['unit_code']!r} != expected {unit_code!r}"
        )
    return row is not None


def _init_instance_metadata(conn, unit_code, schema_version):
    conn.execute(
        "INSERT INTO training_instance_metadata (unit_code, schema_version, created_at_ms) "
        "VALUES (?, ?, ?)",
        (unit_code, schema_version, time_policy.utc_now_ms()),
    )


def _record_migration(conn, version):
    conn.execute(
        "INSERT INTO training_schema_migrations (version, applied_at_ms) VALUES (?, ?)",
        (version, time_policy.utc_now_ms()),
    )


def _execute_script(conn, script):
    """Chạy từng statement để không khiến sqlite3 commit transaction hiện tại."""
    statement = ""
    for line in script.splitlines(keepends=True):
        statement += line
        if sqlite3.complete_statement(statement):
            if statement.strip():
                conn.execute(statement)
            statement = ""
    if statement.strip():
        conn.execute(statement)


# --- migration definitions ---

def migration_001(conn):
    """Foundation: instance metadata, schema migrations, audit log."""
    _execute_script(conn,
        """
        CREATE TABLE IF NOT EXISTS training_instance_metadata (
            unit_code TEXT NOT NULL PRIMARY KEY,
            schema_version INTEGER NOT NULL,
            created_at_ms INTEGER NOT NULL
        );

        CREATE TABLE IF NOT EXISTS training_schema_migrations (
            version INTEGER NOT NULL PRIMARY KEY,
            applied_at_ms INTEGER NOT NULL
        );

        CREATE TABLE IF NOT EXISTS training_audit_log (
            id TEXT NOT NULL PRIMARY KEY,
            actor TEXT NOT NULL,
            unit_code TEXT NOT NULL,
            action TEXT NOT NULL,
            entity_type TEXT NOT NULL,
            entity_id TEXT NOT NULL,
            before_json TEXT,
            after_json TEXT,
            request_id TEXT,
            ip TEXT,
            user_agent TEXT,
            created_at_ms INTEGER NOT NULL
        );
        CREATE INDEX IF NOT EXISTS idx_audit_entity
            ON training_audit_log (entity_type, entity_id, created_at_ms);
        """
    )


def migration_002(conn):
    """Catalog (domains/categories/indicators/topics/competencies/audiences/
    services/tags) + RBAC (user_roles/user_audiences)."""
    _execute_script(conn,
        """
        CREATE TABLE IF NOT EXISTS training_domains (
            code TEXT NOT NULL PRIMARY KEY,
            name TEXT NOT NULL,
            sort_order INTEGER NOT NULL DEFAULT 0
        );

        CREATE TABLE IF NOT EXISTS training_categories (
            code TEXT NOT NULL PRIMARY KEY,
            domain_code TEXT NOT NULL,
            name TEXT NOT NULL,
            FOREIGN KEY (domain_code) REFERENCES training_domains(code)
        );
        CREATE INDEX IF NOT EXISTS idx_categories_domain
            ON training_categories (domain_code);

        CREATE TABLE IF NOT EXISTS training_indicators (
            code TEXT NOT NULL PRIMARY KEY,
            name TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS training_topics (
            code TEXT NOT NULL PRIMARY KEY,
            name TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS training_competencies (
            code TEXT NOT NULL PRIMARY KEY,
            name TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS training_audiences (
            code TEXT NOT NULL PRIMARY KEY,
            name TEXT NOT NULL,
            responsibility_text TEXT
        );

        CREATE TABLE IF NOT EXISTS training_services (
            code TEXT NOT NULL PRIMARY KEY,
            name TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS training_tags (
            code TEXT NOT NULL PRIMARY KEY,
            name TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS training_user_roles (
            id TEXT NOT NULL PRIMARY KEY,
            username TEXT NOT NULL,
            role_code TEXT NOT NULL,
            granted_at_ms INTEGER NOT NULL,
            UNIQUE (username, role_code)
        );
        CREATE INDEX IF NOT EXISTS idx_user_roles_username
            ON training_user_roles (username);

        CREATE TABLE IF NOT EXISTS training_user_audiences (
            id TEXT NOT NULL PRIMARY KEY,
            username TEXT NOT NULL,
            audience_code TEXT NOT NULL,
            UNIQUE (username, audience_code)
        );
        """
    )


def migration_003(conn):
    """Kho tri thức: documents, versions, blocks, issues, rules, mappings."""
    _execute_script(conn,
        """
        CREATE TABLE IF NOT EXISTS knowledge_documents (
            id TEXT NOT NULL PRIMARY KEY,
            document_code TEXT NOT NULL,
            title TEXT NOT NULL,
            document_type TEXT NOT NULL DEFAULT 'kpi_definition',
            issuer TEXT,
            review_status TEXT NOT NULL DEFAULT 'draft',
            created_at_ms INTEGER NOT NULL,
            updated_at_ms INTEGER NOT NULL
        );
        CREATE INDEX IF NOT EXISTS idx_documents_code
            ON knowledge_documents (document_code);
        CREATE INDEX IF NOT EXISTS idx_documents_status
            ON knowledge_documents (review_status);

        CREATE TABLE IF NOT EXISTS knowledge_document_versions (
            id TEXT NOT NULL PRIMARY KEY,
            document_id TEXT NOT NULL,
            version_number INTEGER NOT NULL,
            content_text TEXT NOT NULL,
            content_sha256 TEXT NOT NULL,
            effective_from_ms INTEGER,
            effective_to_ms INTEGER,
            review_status TEXT NOT NULL DEFAULT 'draft',
            created_at_ms INTEGER NOT NULL,
            created_by TEXT NOT NULL,
            FOREIGN KEY (document_id) REFERENCES knowledge_documents(id),
            UNIQUE (document_id, version_number)
        );

        CREATE TABLE IF NOT EXISTS knowledge_document_topics (
            document_version_id TEXT NOT NULL,
            topic_code TEXT NOT NULL,
            PRIMARY KEY (document_version_id, topic_code)
        );

        CREATE TABLE IF NOT EXISTS knowledge_document_audiences (
            document_version_id TEXT NOT NULL,
            audience_code TEXT NOT NULL,
            PRIMARY KEY (document_version_id, audience_code)
        );

        CREATE TABLE IF NOT EXISTS knowledge_blocks (
            id TEXT NOT NULL PRIMARY KEY,
            document_version_id TEXT NOT NULL,
            extraction_revision INTEGER NOT NULL,
            block_id TEXT NOT NULL,
            heading_path_json TEXT,
            char_start INTEGER,
            char_end INTEGER,
            content TEXT NOT NULL,
            content_sha256 TEXT NOT NULL,
            domain_code TEXT,
            category_code TEXT,
            created_at_ms INTEGER NOT NULL,
            FOREIGN KEY (document_version_id) REFERENCES knowledge_document_versions(id),
            UNIQUE (document_version_id, extraction_revision, block_id)
        );
        CREATE INDEX IF NOT EXISTS idx_blocks_version_rev
            ON knowledge_blocks (document_version_id, extraction_revision);

        CREATE TABLE IF NOT EXISTS knowledge_block_topics (
            block_id TEXT NOT NULL,
            topic_code TEXT NOT NULL,
            PRIMARY KEY (block_id, topic_code)
        );

        CREATE TABLE IF NOT EXISTS knowledge_issues (
            id TEXT NOT NULL PRIMARY KEY,
            document_version_id TEXT NOT NULL,
            block_id TEXT,
            severity TEXT NOT NULL,
            issue_type TEXT NOT NULL,
            description TEXT NOT NULL,
            status TEXT NOT NULL DEFAULT 'open',
            resolved_by TEXT,
            resolved_at_ms INTEGER,
            created_at_ms INTEGER NOT NULL,
            FOREIGN KEY (document_version_id) REFERENCES knowledge_document_versions(id)
        );
        CREATE INDEX IF NOT EXISTS idx_issues_version_status
            ON knowledge_issues (document_version_id, status);

        CREATE TABLE IF NOT EXISTS knowledge_rules (
            id TEXT NOT NULL PRIMARY KEY,
            document_version_id TEXT NOT NULL,
            rule_code TEXT NOT NULL,
            rule_text TEXT NOT NULL,
            created_at_ms INTEGER NOT NULL,
            FOREIGN KEY (document_version_id) REFERENCES knowledge_document_versions(id)
        );
        CREATE INDEX IF NOT EXISTS idx_rules_code
            ON knowledge_rules (rule_code);
        """
    )


def migration_004(conn):
    """Kho câu hỏi: items, versions, options, sources, reviews, mappings."""
    _execute_script(conn,
        """
        CREATE TABLE IF NOT EXISTS question_items (
            id TEXT NOT NULL PRIMARY KEY,
            code TEXT,
            current_version_id TEXT,
            source_changed INTEGER NOT NULL DEFAULT 0,
            created_at_ms INTEGER NOT NULL
        );
        CREATE INDEX IF NOT EXISTS idx_items_code ON question_items (code);

        CREATE TABLE IF NOT EXISTS question_versions (
            id TEXT NOT NULL PRIMARY KEY,
            question_item_id TEXT NOT NULL,
            version_number INTEGER NOT NULL,
            type TEXT NOT NULL,
            stem TEXT NOT NULL,
            stimulus TEXT,
            language TEXT NOT NULL DEFAULT 'vi',
            correct_option_ids_json TEXT NOT NULL,
            explanation TEXT,
            distractor_rationales_json TEXT,
            difficulty TEXT NOT NULL,
            cognitive_level TEXT,
            criticality TEXT,
            estimated_seconds INTEGER,
            max_score REAL NOT NULL DEFAULT 1.0,
            scoring_policy_json TEXT,
            normalized_stem_hash TEXT,
            review_status TEXT NOT NULL DEFAULT 'draft',
            publication_status TEXT NOT NULL DEFAULT 'unpublished',
            created_by TEXT NOT NULL,
            created_at_ms INTEGER NOT NULL,
            approved_by TEXT,
            approved_at_ms INTEGER,
            FOREIGN KEY (question_item_id) REFERENCES question_items(id),
            UNIQUE (question_item_id, version_number)
        );
        CREATE INDEX IF NOT EXISTS idx_qversions_pub ON question_versions (publication_status);
        CREATE INDEX IF NOT EXISTS idx_qversions_review ON question_versions (review_status);

        CREATE TABLE IF NOT EXISTS question_options (
            id TEXT NOT NULL PRIMARY KEY,
            question_version_id TEXT NOT NULL,
            option_code TEXT NOT NULL,
            option_text TEXT NOT NULL,
            display_order INTEGER NOT NULL,
            FOREIGN KEY (question_version_id) REFERENCES question_versions(id),
            UNIQUE (question_version_id, option_code)
        );

        CREATE TABLE IF NOT EXISTS question_sources (
            id TEXT NOT NULL PRIMARY KEY,
            question_version_id TEXT NOT NULL,
            document_version_id TEXT NOT NULL,
            block_id TEXT NOT NULL,
            extraction_revision INTEGER NOT NULL,
            quoted_text TEXT,
            quote_start INTEGER,
            quote_end INTEGER,
            supports TEXT,
            FOREIGN KEY (question_version_id) REFERENCES question_versions(id)
        );
        CREATE INDEX IF NOT EXISTS idx_qsources_block ON question_sources (document_version_id, block_id);
        CREATE INDEX IF NOT EXISTS idx_qsources_version ON question_sources (question_version_id);

        CREATE TABLE IF NOT EXISTS question_topics (
            question_version_id TEXT NOT NULL,
            topic_code TEXT NOT NULL,
            PRIMARY KEY (question_version_id, topic_code)
        );
        CREATE TABLE IF NOT EXISTS question_indicators (
            question_version_id TEXT NOT NULL,
            indicator_code TEXT NOT NULL,
            PRIMARY KEY (question_version_id, indicator_code)
        );
        CREATE TABLE IF NOT EXISTS question_audiences (
            question_version_id TEXT NOT NULL,
            audience_code TEXT NOT NULL,
            PRIMARY KEY (question_version_id, audience_code)
        );
        CREATE TABLE IF NOT EXISTS question_competencies (
            question_version_id TEXT NOT NULL,
            competency_code TEXT NOT NULL,
            PRIMARY KEY (question_version_id, competency_code)
        );
        CREATE TABLE IF NOT EXISTS question_tags (
            question_version_id TEXT NOT NULL,
            tag_code TEXT NOT NULL,
            PRIMARY KEY (question_version_id, tag_code)
        );

        CREATE TABLE IF NOT EXISTS question_reviews (
            id TEXT NOT NULL PRIMARY KEY,
            question_version_id TEXT NOT NULL,
            action TEXT NOT NULL,
            reviewer TEXT NOT NULL,
            comment TEXT,
            created_at_ms INTEGER NOT NULL,
            FOREIGN KEY (question_version_id) REFERENCES question_versions(id)
        );
        CREATE INDEX IF NOT EXISTS idx_qreviews_version ON question_reviews (question_version_id);
        """
    )


def migration_005(conn):
    """AI generation jobs + batches."""
    _execute_script(conn,
        """
        CREATE TABLE IF NOT EXISTS ai_generation_jobs (
            id TEXT NOT NULL PRIMARY KEY,
            status TEXT NOT NULL DEFAULT 'pending',
            idempotency_key TEXT UNIQUE,
            request_payload_json TEXT NOT NULL,
            source_document_version_ids_json TEXT NOT NULL,
            target_audience_codes_json TEXT NOT NULL,
            requested_count INTEGER NOT NULL,
            created_by TEXT NOT NULL,
            created_at_ms INTEGER NOT NULL,
            claimed_by TEXT,
            lease_expires_at_ms INTEGER,
            heartbeat_at_ms INTEGER,
            retry_count INTEGER NOT NULL DEFAULT 0,
            max_retries INTEGER NOT NULL DEFAULT 3,
            error_code TEXT,
            error_detail TEXT,
            completed_at_ms INTEGER
        );
        CREATE INDEX IF NOT EXISTS idx_jobs_claim
            ON ai_generation_jobs (status, lease_expires_at_ms);

        CREATE TABLE IF NOT EXISTS ai_generation_batches (
            id TEXT NOT NULL PRIMARY KEY,
            job_id TEXT NOT NULL,
            schema_version TEXT NOT NULL,
            provider TEXT NOT NULL,
            model TEXT,
            prompt_version TEXT,
            usage_json TEXT,
            raw_response_text TEXT,
            questions_payload_json TEXT,
            created_at_ms INTEGER NOT NULL,
            FOREIGN KEY (job_id) REFERENCES ai_generation_jobs(id),
            UNIQUE (job_id)
        );
        """
    )


def migration_006(conn):
    """Đề và kỳ thi: templates, template_items, events, assignments."""
    _execute_script(conn,
        """
        CREATE TABLE IF NOT EXISTS exam_templates (
            id TEXT NOT NULL PRIMARY KEY,
            code TEXT NOT NULL,
            title TEXT NOT NULL,
            target_audience_code TEXT NOT NULL,
            total_questions INTEGER NOT NULL,
            duration_seconds INTEGER NOT NULL,
            pass_score_percent REAL NOT NULL DEFAULT 80.0,
            shuffle_questions INTEGER NOT NULL DEFAULT 0,
            shuffle_options INTEGER NOT NULL DEFAULT 0,
            locked INTEGER NOT NULL DEFAULT 0,
            created_by TEXT NOT NULL,
            created_at_ms INTEGER NOT NULL
        );
        CREATE INDEX IF NOT EXISTS idx_templates_code ON exam_templates (code);

        CREATE TABLE IF NOT EXISTS exam_template_items (
            id TEXT NOT NULL PRIMARY KEY,
            template_id TEXT NOT NULL,
            sequence_number INTEGER NOT NULL,
            question_version_id TEXT NOT NULL,
            section_label TEXT,
            points REAL NOT NULL DEFAULT 1.0,
            FOREIGN KEY (template_id) REFERENCES exam_templates(id),
            UNIQUE (template_id, sequence_number)
        );
        CREATE INDEX IF NOT EXISTS idx_titems_template_qv
            ON exam_template_items (template_id, question_version_id);

        CREATE TABLE IF NOT EXISTS exam_events (
            id TEXT NOT NULL PRIMARY KEY,
            code TEXT NOT NULL,
            title TEXT NOT NULL,
            description TEXT,
            template_id TEXT NOT NULL,
            target_audience_code TEXT NOT NULL,
            status TEXT NOT NULL DEFAULT 'draft',
            start_at_ms INTEGER NOT NULL,
            end_at_ms INTEGER NOT NULL,
            duration_seconds INTEGER NOT NULL,
            pass_score_percent REAL NOT NULL,
            reveal_answers_after_finalize INTEGER NOT NULL DEFAULT 1,
            created_by TEXT NOT NULL,
            created_at_ms INTEGER NOT NULL,
            finalized_at_ms INTEGER,
            finalized_by TEXT,
            FOREIGN KEY (template_id) REFERENCES exam_templates(id)
        );
        CREATE INDEX IF NOT EXISTS idx_events_status ON exam_events (status);
        CREATE INDEX IF NOT EXISTS idx_events_start ON exam_events (start_at_ms);

        CREATE TABLE IF NOT EXISTS exam_assignments (
            id TEXT NOT NULL PRIMARY KEY,
            exam_event_id TEXT NOT NULL,
            username TEXT NOT NULL,
            display_name TEXT NOT NULL,
            team_code TEXT,
            team_name TEXT,
            organization_code TEXT,
            organization_name TEXT,
            audience_code TEXT NOT NULL,
            status TEXT NOT NULL DEFAULT 'assigned',
            duration_seconds INTEGER NOT NULL,
            retake_of_assignment_id TEXT,
            assigned_at_ms INTEGER NOT NULL,
            user_source_updated_at_ms INTEGER,
            FOREIGN KEY (exam_event_id) REFERENCES exam_events(id),
            UNIQUE (exam_event_id, username, audience_code)
        );
        CREATE INDEX IF NOT EXISTS idx_assignments_event_user_status
            ON exam_assignments (exam_event_id, username, status);
        """
    )


def migration_007(conn):
    """Bài làm và kết quả: attempts, attempt_items/options, responses, results, reports."""
    _execute_script(conn,
        """
        CREATE TABLE IF NOT EXISTS exam_attempts (
            id TEXT NOT NULL PRIMARY KEY,
            assignment_id TEXT NOT NULL,
            status TEXT NOT NULL DEFAULT 'created',
            random_seed INTEGER NOT NULL,
            shuffle_algorithm_version TEXT NOT NULL,
            snapshot_checksum TEXT NOT NULL,
            started_at_ms INTEGER,
            deadline_at_ms INTEGER,
            submitted_at_ms INTEGER,
            ended_reason TEXT,
            created_at_ms INTEGER NOT NULL,
            FOREIGN KEY (assignment_id) REFERENCES exam_assignments(id),
            UNIQUE (assignment_id)
        );
        CREATE INDEX IF NOT EXISTS idx_attempts_assignment_status
            ON exam_attempts (assignment_id, status);

        CREATE TABLE IF NOT EXISTS exam_attempt_items (
            id TEXT NOT NULL PRIMARY KEY,
            attempt_id TEXT NOT NULL,
            sequence_number INTEGER NOT NULL,
            question_version_id TEXT NOT NULL,
            type TEXT NOT NULL,
            stem TEXT NOT NULL,
            stimulus TEXT,
            language TEXT NOT NULL,
            correct_option_ids_json TEXT NOT NULL,
            explanation TEXT,
            distractor_rationales_json TEXT,
            difficulty TEXT,
            section_label TEXT,
            topic_codes_json TEXT,
            competency_codes_json TEXT,
            points REAL NOT NULL,
            max_score REAL NOT NULL,
            scoring_policy_json TEXT,
            evidence_json TEXT,
            FOREIGN KEY (attempt_id) REFERENCES exam_attempts(id),
            UNIQUE (attempt_id, sequence_number)
        );

        CREATE TABLE IF NOT EXISTS exam_attempt_options (
            id TEXT NOT NULL PRIMARY KEY,
            attempt_item_id TEXT NOT NULL,
            option_code TEXT NOT NULL,
            option_text TEXT NOT NULL,
            display_order INTEGER NOT NULL,
            FOREIGN KEY (attempt_item_id) REFERENCES exam_attempt_items(id),
            UNIQUE (attempt_item_id, option_code)
        );

        CREATE TABLE IF NOT EXISTS exam_responses (
            id TEXT NOT NULL PRIMARY KEY,
            attempt_id TEXT NOT NULL,
            attempt_item_id TEXT NOT NULL,
            selected_option_ids_json TEXT NOT NULL DEFAULT '[]',
            client_revision INTEGER NOT NULL DEFAULT 0,
            answered_at_ms INTEGER,
            FOREIGN KEY (attempt_id) REFERENCES exam_attempts(id),
            FOREIGN KEY (attempt_item_id) REFERENCES exam_attempt_items(id),
            UNIQUE (attempt_id, attempt_item_id)
        );
        CREATE INDEX IF NOT EXISTS idx_responses_attempt ON exam_responses (attempt_id);

        CREATE TABLE IF NOT EXISTS exam_results (
            id TEXT NOT NULL PRIMARY KEY,
            attempt_id TEXT NOT NULL,
            raw_score REAL NOT NULL,
            maximum_score REAL NOT NULL,
            percent REAL NOT NULL,
            passed INTEGER NOT NULL,
            topic_breakdown_json TEXT NOT NULL,
            scored_at_ms INTEGER NOT NULL,
            status TEXT NOT NULL DEFAULT 'scored',
            FOREIGN KEY (attempt_id) REFERENCES exam_attempts(id),
            UNIQUE (attempt_id)
        );
        CREATE INDEX IF NOT EXISTS idx_results_attempt ON exam_results (attempt_id);

        CREATE TABLE IF NOT EXISTS exam_result_topics (
            result_id TEXT NOT NULL,
            topic_code TEXT NOT NULL,
            correct_count INTEGER NOT NULL,
            total_count INTEGER NOT NULL,
            points REAL NOT NULL,
            max_points REAL NOT NULL,
            PRIMARY KEY (result_id, topic_code)
        );

        CREATE TABLE IF NOT EXISTS exam_score_adjustments (
            id TEXT NOT NULL PRIMARY KEY,
            result_id TEXT NOT NULL,
            delta_score REAL NOT NULL,
            reason TEXT NOT NULL,
            adjusted_by TEXT NOT NULL,
            created_at_ms INTEGER NOT NULL,
            FOREIGN KEY (result_id) REFERENCES exam_results(id)
        );
        CREATE INDEX IF NOT EXISTS idx_adjustments_result ON exam_score_adjustments (result_id);

        CREATE TABLE IF NOT EXISTS exam_report_snapshots (
            id TEXT NOT NULL PRIMARY KEY,
            exam_event_id TEXT NOT NULL,
            revision INTEGER NOT NULL,
            schema_version TEXT NOT NULL,
            report_algorithm_version TEXT NOT NULL,
            payload_json TEXT NOT NULL,
            checksum TEXT NOT NULL,
            previous_revision INTEGER,
            created_by TEXT NOT NULL,
            created_at_ms INTEGER NOT NULL,
            FOREIGN KEY (exam_event_id) REFERENCES exam_events(id),
            UNIQUE (exam_event_id, revision)
        );
        """
    )


def migration_008(conn):
    """Khóa template item vào question version và loại trừ câu hỏi trùng."""
    # Keep foreign keys enabled; valid parent joins make the rebuild safe in the runner transaction.
    conn.execute("DROP TABLE IF EXISTS exam_template_items_new")
    conn.execute(
        """
        CREATE TABLE exam_template_items_new (
            id TEXT NOT NULL PRIMARY KEY,
            template_id TEXT NOT NULL,
            sequence_number INTEGER NOT NULL,
            question_version_id TEXT NOT NULL,
            section_label TEXT,
            points REAL NOT NULL DEFAULT 1.0,
            FOREIGN KEY (template_id) REFERENCES exam_templates(id),
            FOREIGN KEY (question_version_id) REFERENCES question_versions(id),
            UNIQUE (template_id, sequence_number),
            UNIQUE (template_id, question_version_id)
        )
        """
    )
    conn.execute(
        """INSERT INTO exam_template_items_new
        (id, template_id, sequence_number, question_version_id, section_label, points)
        SELECT ti.id, ti.template_id, ti.sequence_number, ti.question_version_id,
               ti.section_label, ti.points
        FROM exam_template_items ti
        JOIN exam_templates et ON et.id = ti.template_id
        JOIN question_versions qv ON qv.id = ti.question_version_id
        WHERE NOT EXISTS (
            SELECT 1
            FROM exam_template_items earlier
            JOIN exam_templates earlier_template ON earlier_template.id = earlier.template_id
            JOIN question_versions earlier_version ON earlier_version.id = earlier.question_version_id
            WHERE earlier.template_id = ti.template_id
              AND earlier.question_version_id = ti.question_version_id
              AND earlier.rowid < ti.rowid
        )"""
    )
    conn.execute("DROP TABLE exam_template_items")
    conn.execute("ALTER TABLE exam_template_items_new RENAME TO exam_template_items")
    conn.execute(
        "CREATE INDEX idx_titems_template_qv ON exam_template_items (template_id, question_version_id)"
    )


_MIGRATIONS = [
    (1, migration_001),
    (2, migration_002),
    (3, migration_003),
    (4, migration_004),
    (5, migration_005),
    (6, migration_006),
    (7, migration_007),
    (8, migration_008),
]


def run_migrations(db_path, unit_code, *, migrations=None):
    """Chạy migration pending với file lock + unit_code check."""
    lock_file = _acquire_lock(db_path)
    try:
        conn = write_connection(db_path)
        try:
            metadata_exists = _check_or_init_instance_metadata(conn, unit_code)
            applied = _applied_versions(conn)
            pending = [
                (version, fn)
                for version, fn in (migrations or _MIGRATIONS)
                if version not in applied
            ]
            if not pending:
                return applied
            conn.execute("BEGIN")
            for version, fn in pending:
                fn(conn)
                _record_migration(conn, version)
            max_version = max(v for v, _ in pending)
            if metadata_exists:
                conn.execute(
                    "UPDATE training_instance_metadata SET schema_version=?",
                    (max_version,),
                )
            else:
                _init_instance_metadata(conn, unit_code, max_version)
            conn.commit()
            return applied | {v for v, _ in pending}
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()
    finally:
        fcntl.flock(lock_file.fileno(), fcntl.LOCK_UN)
        lock_file.close()


def get_schema_version(db_path):
    conn = write_connection(db_path)
    try:
        if not _table_exists(conn, "training_instance_metadata"):
            return 0
        row = conn.execute(
            "SELECT schema_version FROM training_instance_metadata"
        ).fetchone()
        return row["schema_version"] if row else 0
    finally:
        conn.close()
