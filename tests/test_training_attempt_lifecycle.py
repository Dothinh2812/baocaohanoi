import json
import threading

import pytest

from training import db as training_db
from training import migrations, time_policy
from training.errors import ErrorCode, TrainingError
from services import training_exam_service as es
from services import training_attempt_service as att
from services import training_question_service as qs
from services import training_scoring_service as scoring
from services.training_catalog_service import seed_defaults


VALID_BATCH = {
    "schema_version": "1.0",
    "batch": {
        "title": "Test", "language": "vi",
        "source_document_version_ids": ["docver-001"],
        "target_audience_codes": ["nvkt"], "requested_count": 3,
    },
    "questions": [
        {
            "local_ref": f"Q{i}", "type": "single_choice",
            "stem": f"Câu {i}?",
            "options": [{"id": "A", "text": "sai"}, {"id": "B", "text": "đúng"}],
            "correct_option_ids": ["B" if i != 2 else "A"],
            "explanation": f"Giải thích {i}",
            "distractor_rationales": {"A": "sai"},
            "classification": {"domain_code": "quality", "topic_codes": [f"topic_{i}"]},
            "difficulty": "easy",
            "evidence": [
                {"document_version_id": "docver-001", "block_id": "DOC-B001",
                 "extraction_revision": 1, "quoted_text": "q", "supports": "correct_answer"}
            ],
        }
        for i in range(1, 4)
    ],
}


def _setup(monkeypatch, tmp_path):
    db_path = str(tmp_path / "training.db")
    monkeypatch.setattr(training_db, "TRAINING_DB_PATH", db_path)
    monkeypatch.setattr(training_db, "UNIT_CODE", "son_tay")
    migrations.run_migrations(db_path, "son_tay")
    seed_defaults(db_path, "son_tay")
    qs.reset_duplicate_cache()
    return db_path


def _make_open_exam_with_assignment(
    db_path, username="learner1", *, additional_users=(), shuffle_questions=False,
    shuffle_options=False,
):
    result = qs.import_question_batch(db_path, unit_code="son_tay", actor="ed",
                                       batch=VALID_BATCH, status="draft")
    for vid in result["version_ids"]:
        qs.add_review_action(db_path, unit_code="son_tay", actor="mgr", version_id=vid, action="approve")
        qs.publish_question_version(db_path, unit_code="son_tay", actor="mgr", version_id=vid)
    tpl = es.create_template(db_path, unit_code="son_tay", actor="mgr",
        code="T", title="T", target_audience_code="nvkt",
        question_version_ids=result["version_ids"], duration_seconds=600, pass_score_percent=80.0,
        shuffle_questions=shuffle_questions, shuffle_options=shuffle_options)
    now = time_policy.utc_now_ms()
    exam = es.create_exam(db_path, unit_code="son_tay", actor="mgr",
        code="E", title="E", template_id=tpl["id"], target_audience_code="nvkt",
        start_at_ms=now - 1000, end_at_ms=now + 3600_000,
        duration_seconds=600, pass_score_percent=80.0)
    assignment_ids = es.create_assignments(db_path, unit_code="son_tay", actor="mgr",
        exam_id=exam["id"],
        users=[{"username": username, "display_name": "Learner"}, *additional_users],
        audience_code="nvkt")
    es.ready_exam(db_path, unit_code="son_tay", actor="mgr", exam_id=exam["id"])
    es.open_exam(db_path, unit_code="son_tay", actor="mgr", exam_id=exam["id"])
    return exam["id"], assignment_ids[0]


# --- M3.3: start attempt + snapshot ---

def test_start_attempt_creates_active(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    _, assignment_id = _make_open_exam_with_assignment(db_path)
    result = att.start_attempt(db_path, unit_code="son_tay", actor="learner1",
                               assignment_id=assignment_id)
    assert result["status"] == "active"
    assert result["deadline_at_ms"] > result["started_at_ms"]
    assert result["attempt_id"]


def test_start_attempt_idempotent_returns_existing(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    _, assignment_id = _make_open_exam_with_assignment(db_path)
    r1 = att.start_attempt(db_path, unit_code="son_tay", actor="learner1", assignment_id=assignment_id)
    r2 = att.start_attempt(db_path, unit_code="son_tay", actor="learner1", assignment_id=assignment_id)
    assert r1["attempt_id"] == r2["attempt_id"]


def test_two_concurrent_starts_create_one_attempt(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    _, assignment_id = _make_open_exam_with_assignment(db_path)
    results = []
    barrier = threading.Barrier(2)

    def do_start():
        barrier.wait()
        try:
            r = att.start_attempt(db_path, unit_code="son_tay", actor="learner1", assignment_id=assignment_id)
            results.append(r["attempt_id"])
        except TrainingError:
            results.append(None)

    t1 = threading.Thread(target=do_start)
    t2 = threading.Thread(target=do_start)
    t1.start()
    t2.start()
    t1.join(timeout=15)
    t2.join(timeout=15)

    attempt_ids = [r for r in results if r is not None]
    assert len(set(attempt_ids)) == 1
    conn = training_db.read_connection(db_path)
    count = conn.execute(
        "SELECT COUNT(*) AS c FROM exam_attempts WHERE assignment_id=?", (assignment_id,)
    ).fetchone()["c"]
    conn.close()
    assert count == 1


def test_learner_view_has_no_scoring_data(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    _, assignment_id = _make_open_exam_with_assignment(db_path)
    result = att.start_attempt(db_path, unit_code="son_tay", actor="learner1", assignment_id=assignment_id)
    view = att.get_attempt_learner_view(db_path, result["attempt_id"])
    for item in view["items"]:
        assert "correct_option_ids" not in item
        assert "explanation" not in item
        assert "evidence" not in item
        assert "scoring_policy" not in item
        assert "max_score" not in item
        assert "options" in item
        assert "stem" in item
        assert "response" in item


def test_snapshot_has_correct_data(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    _, assignment_id = _make_open_exam_with_assignment(db_path)
    result = att.start_attempt(db_path, unit_code="son_tay", actor="learner1", assignment_id=assignment_id)
    conn = training_db.read_connection(db_path)
    items = conn.execute(
        "SELECT * FROM exam_attempt_items WHERE attempt_id=?", (result["attempt_id"],)
    ).fetchall()
    opts = conn.execute(
        "SELECT * FROM exam_attempt_options WHERE attempt_item_id IN "
        "(SELECT id FROM exam_attempt_items WHERE attempt_id=?)",
        (result["attempt_id"],),
    ).fetchall()
    conn.close()
    assert len(items) == 3
    assert len(opts) == 6


def _snapshot_order(db_path, attempt_id):
    conn = training_db.read_connection(db_path)
    try:
        items = conn.execute(
            "SELECT question_version_id FROM exam_attempt_items WHERE attempt_id=? ORDER BY sequence_number",
            (attempt_id,),
        ).fetchall()
        options = conn.execute(
            """SELECT o.option_code FROM exam_attempt_options o
            JOIN exam_attempt_items i ON i.id=o.attempt_item_id
            WHERE i.attempt_id=? ORDER BY i.sequence_number, o.display_order""",
            (attempt_id,),
        ).fetchall()
        checksum = conn.execute(
            "SELECT snapshot_checksum FROM exam_attempts WHERE id=?", (attempt_id,)
        ).fetchone()["snapshot_checksum"]
        return [row["question_version_id"] for row in items], [row["option_code"] for row in options], checksum
    finally:
        conn.close()


def _snapshot_presentation(db_path, attempt_id):
    conn = training_db.read_connection(db_path)
    try:
        items = []
        for item in conn.execute(
            "SELECT sequence_number, question_version_id FROM exam_attempt_items "
            "WHERE attempt_id=? ORDER BY sequence_number", (attempt_id,)
        ).fetchall():
            options = conn.execute(
                "SELECT option_code, display_order FROM exam_attempt_options "
                "WHERE attempt_item_id=(SELECT id FROM exam_attempt_items "
                "WHERE attempt_id=? AND sequence_number=?) ORDER BY display_order",
                (attempt_id, item["sequence_number"]),
            ).fetchall()
            items.append({
                "sequence_number": item["sequence_number"],
                "qv": item["question_version_id"],
                "options": [(option["option_code"], option["display_order"]) for option in options],
            })
        return items
    finally:
        conn.close()


def test_shuffle_snapshot_is_reproducible_and_includes_order_in_checksum(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    exam_id, assignment_id = _make_open_exam_with_assignment(
        db_path, additional_users=({"username": "learner2", "display_name": "Learner 2"},),
        shuffle_questions=True, shuffle_options=True,
    )
    second_assignment_id = es.get_assignment_for_user(db_path, exam_id, "learner2")[0]["id"]
    monkeypatch.setattr(att.secrets, "randbits", lambda _: 1)
    monkeypatch.setattr(att, "SHUFFLE_ALGORITHM_VERSION", "persisted-v1")

    first = att.start_attempt(db_path, unit_code="son_tay", actor="learner1", assignment_id=assignment_id)
    second = att.start_attempt(db_path, unit_code="son_tay", actor="learner2", assignment_id=second_assignment_id)
    monkeypatch.setattr(att, "SHUFFLE_ALGORITHM_VERSION", "current-v2")

    first_order = _snapshot_order(db_path, first["attempt_id"])
    second_order = _snapshot_order(db_path, second["attempt_id"])
    assert first_order == second_order
    snapshot = _snapshot_presentation(db_path, first["attempt_id"])
    assert att.snapshot_order(db_path, first["attempt_id"]) == snapshot
    assert att.read_snapshot_presentation(db_path, first["attempt_id"]) == snapshot
    template_items = es.get_template_items(db_path, es.get_exam(db_path, exam_id)["template_id"])
    assert first_order[0] != [item["question_version_id"] for item in template_items]
    assert first_order[1] != ["A", "B"] * 3


def test_shuffle_flags_false_preserve_template_and_option_order(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    exam_id, assignment_id = _make_open_exam_with_assignment(
        db_path, shuffle_questions=False, shuffle_options=False,
    )
    monkeypatch.setattr(att.secrets, "randbits", lambda _: 1)

    result = att.start_attempt(db_path, unit_code="son_tay", actor="learner1", assignment_id=assignment_id)

    item_order, option_order, _ = _snapshot_order(db_path, result["attempt_id"])
    template_items = es.get_template_items(db_path, es.get_exam(db_path, exam_id)["template_id"])
    assert item_order == [item["question_version_id"] for item in template_items]
    assert option_order == ["A", "B"] * 3


def test_question_shuffle_does_not_shuffle_options(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    exam_id, assignment_id = _make_open_exam_with_assignment(
        db_path, shuffle_questions=True, shuffle_options=False,
    )
    monkeypatch.setattr(att.secrets, "randbits", lambda _: 1)

    result = att.start_attempt(db_path, unit_code="son_tay", actor="learner1", assignment_id=assignment_id)

    item_order, option_order, _ = _snapshot_order(db_path, result["attempt_id"])
    template_items = es.get_template_items(db_path, es.get_exam(db_path, exam_id)["template_id"])
    assert item_order != [item["question_version_id"] for item in template_items]
    assert option_order == ["A", "B"] * 3


def test_option_shuffle_does_not_shuffle_questions(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    exam_id, assignment_id = _make_open_exam_with_assignment(
        db_path, shuffle_questions=False, shuffle_options=True,
    )
    monkeypatch.setattr(att.secrets, "randbits", lambda _: 1)

    result = att.start_attempt(db_path, unit_code="son_tay", actor="learner1", assignment_id=assignment_id)

    item_order, option_order, _ = _snapshot_order(db_path, result["attempt_id"])
    template_items = es.get_template_items(db_path, es.get_exam(db_path, exam_id)["template_id"])
    assert item_order == [item["question_version_id"] for item in template_items]
    assert option_order != ["A", "B"] * 3


def test_snapshot_checksum_includes_assigned_sequence_and_option_display_order(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    _, assignment_id = _make_open_exam_with_assignment(
        db_path, shuffle_questions=True, shuffle_options=True,
    )
    monkeypatch.setattr(att.secrets, "randbits", lambda _: 1)
    result = att.start_attempt(db_path, unit_code="son_tay", actor="learner1", assignment_id=assignment_id)
    assert att.verify_snapshot_checksum(db_path, result["attempt_id"]) is True
    conn = training_db.write_connection(db_path)
    try:
        conn.execute(
            "UPDATE exam_attempt_options SET display_order=99 WHERE attempt_item_id=("
            "SELECT id FROM exam_attempt_items WHERE attempt_id=? ORDER BY sequence_number LIMIT 1) "
            "AND display_order=0",
            (result["attempt_id"],),
        )
        conn.commit()
    finally:
        conn.close()
    assert att.verify_snapshot_checksum(db_path, result["attempt_id"]) is False


def test_persisted_snapshot_presentation_is_stable_after_template_mutation(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    exam_id, assignment_id = _make_open_exam_with_assignment(
        db_path, shuffle_questions=True, shuffle_options=True,
    )
    monkeypatch.setattr(att.secrets, "randbits", lambda _: 1)
    result = att.start_attempt(db_path, unit_code="son_tay", actor="learner1", assignment_id=assignment_id)
    expected = att.read_snapshot_presentation(db_path, result["attempt_id"])
    template_id = es.get_exam(db_path, exam_id)["template_id"]
    conn = training_db.write_connection(db_path)
    conn.execute("UPDATE exam_templates SET shuffle_questions=0, shuffle_options=0 WHERE id=?", (template_id,))
    conn.execute("DELETE FROM exam_template_items WHERE template_id=?", (template_id,))
    conn.commit()
    conn.close()

    assert att.read_snapshot_presentation(db_path, result["attempt_id"]) == expected


def test_snapshot_checksum_verification_detects_tampered_item_content(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    _, assignment_id = _make_open_exam_with_assignment(db_path)
    result = att.start_attempt(db_path, unit_code="son_tay", actor="learner1", assignment_id=assignment_id)
    assert att.verify_snapshot_checksum(db_path, result["attempt_id"]) is True
    conn = training_db.write_connection(db_path)
    conn.execute(
        "UPDATE exam_attempt_items SET stem='tampered' WHERE attempt_id=?", (result["attempt_id"],)
    )
    conn.commit()
    conn.close()
    assert att.verify_snapshot_checksum(db_path, result["attempt_id"]) is False


def test_snapshot_helpers_raise_not_found_for_missing_attempt(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    for helper in (att.snapshot_order, att.read_snapshot_presentation, att.verify_snapshot_checksum):
        with pytest.raises(TrainingError) as exc_info:
            helper(db_path, "missing")
        assert exc_info.value.code == ErrorCode.NOT_FOUND


@pytest.mark.parametrize("tamper", ["item", "option", "order"])
def test_submit_rejects_tampered_snapshot_with_one_audit_and_no_result(monkeypatch, tmp_path, tamper):
    db_path = _setup(monkeypatch, tmp_path)
    _, assignment_id = _make_open_exam_with_assignment(db_path)
    started = att.start_attempt(db_path, unit_code="son_tay", actor="learner1", assignment_id=assignment_id)
    conn = training_db.write_connection(db_path)
    try:
        if tamper == "item":
            conn.execute("UPDATE exam_attempt_items SET stem='tampered' WHERE attempt_id=?", (started["attempt_id"],))
        elif tamper == "option":
            conn.execute(
                "UPDATE exam_attempt_options SET option_text='tampered' WHERE attempt_item_id="
                "(SELECT id FROM exam_attempt_items WHERE attempt_id=? LIMIT 1)",
                (started["attempt_id"],),
            )
        else:
            conn.execute(
                "UPDATE exam_attempt_options SET display_order=99 WHERE attempt_item_id="
                "(SELECT id FROM exam_attempt_items WHERE attempt_id=? LIMIT 1) AND display_order=0",
                (started["attempt_id"],),
            )
        conn.commit()
    finally:
        conn.close()

    with pytest.raises(TrainingError) as exc:
        att.submit_attempt(db_path, unit_code="son_tay", actor="learner1", attempt_id=started["attempt_id"])
    assert exc.value.code == "ATTEMPT_SNAPSHOT_INVALID"
    conn = training_db.read_connection(db_path)
    try:
        audits = conn.execute(
            "SELECT COUNT(*) AS count FROM training_audit_log "
            "WHERE action='attempt_snapshot_invalid' AND entity_id=?", (started["attempt_id"],)
        ).fetchone()["count"]
        result = conn.execute("SELECT 1 FROM exam_results WHERE attempt_id=?", (started["attempt_id"],)).fetchone()
    finally:
        conn.close()
    assert audits == 1
    assert result is None


def test_direct_scoring_rejects_tampered_snapshot_with_audit(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    _, assignment_id = _make_open_exam_with_assignment(db_path)
    started = att.start_attempt(db_path, unit_code="son_tay", actor="learner1", assignment_id=assignment_id)
    conn = training_db.write_connection(db_path)
    try:
        conn.execute("UPDATE exam_attempt_items SET stem='tampered' WHERE attempt_id=?", (started["attempt_id"],))
        conn.commit()
    finally:
        conn.close()

    with pytest.raises(TrainingError) as exc:
        scoring.score_attempt(db_path, started["attempt_id"])
    assert exc.value.code == "ATTEMPT_SNAPSHOT_INVALID"
    conn = training_db.read_connection(db_path)
    try:
        audits = conn.execute(
            "SELECT COUNT(*) AS count FROM training_audit_log "
            "WHERE action='attempt_snapshot_invalid' AND entity_id=?", (started["attempt_id"],)
        ).fetchone()["count"]
    finally:
        conn.close()
    assert audits == 1


# --- M3.4: autosave ---

def test_autosave_stores_response(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    _, assignment_id = _make_open_exam_with_assignment(db_path)
    result = att.start_attempt(db_path, unit_code="son_tay", actor="learner1", assignment_id=assignment_id)
    view = att.get_attempt_learner_view(db_path, result["attempt_id"])
    item_id = view["items"][0]["item_id"]
    save_result = att.save_response(db_path, attempt_id=result["attempt_id"],
                                    attempt_item_id=item_id,
                                    selected_option_ids=["B"], client_revision=1)
    assert save_result["accepted"] is True
    assert save_result["stored_revision"] == 1


def test_old_revision_does_not_overwrite_new(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    _, assignment_id = _make_open_exam_with_assignment(db_path)
    result = att.start_attempt(db_path, unit_code="son_tay", actor="learner1", assignment_id=assignment_id)
    view = att.get_attempt_learner_view(db_path, result["attempt_id"])
    item_id = view["items"][0]["item_id"]
    att.save_response(db_path, attempt_id=result["attempt_id"], attempt_item_id=item_id,
                      selected_option_ids=["B"], client_revision=5)
    old = att.save_response(db_path, attempt_id=result["attempt_id"], attempt_item_id=item_id,
                            selected_option_ids=["A"], client_revision=2)
    assert old["accepted"] is False
    assert old["stored_revision"] == 5
    stored = att.get_attempt_learner_view(db_path, result["attempt_id"])["items"][0]["response"]
    assert stored["selected_option_ids"] == ["B"]
    assert stored["client_revision"] == 5


@pytest.mark.parametrize("client_revision", [0, True, "2"])
def test_autosave_rejects_non_positive_or_non_integer_revision(monkeypatch, tmp_path, client_revision):
    db_path = _setup(monkeypatch, tmp_path)
    _, assignment_id = _make_open_exam_with_assignment(db_path)
    result = att.start_attempt(db_path, unit_code="son_tay", actor="learner1", assignment_id=assignment_id)
    item_id = att.get_attempt_learner_view(db_path, result["attempt_id"])["items"][0]["item_id"]

    with pytest.raises(TrainingError) as exc_info:
        att.save_response(db_path, attempt_id=result["attempt_id"], attempt_item_id=item_id,
                          selected_option_ids=["B"], client_revision=client_revision)

    assert exc_info.value.code == ErrorCode.VALIDATION_ERROR


def test_autosave_rejects_item_from_another_attempt_without_response_mutation(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    exam_id, assignment_id = _make_open_exam_with_assignment(
        db_path, additional_users=({"username": "learner2", "display_name": "Learner 2"},),
    )
    other_assignment_id = es.get_assignment_for_user(db_path, exam_id, "learner2")[0]["id"]
    attempt = att.start_attempt(db_path, unit_code="son_tay", actor="learner1", assignment_id=assignment_id)
    other_attempt = att.start_attempt(
        db_path, unit_code="son_tay", actor="learner2", assignment_id=other_assignment_id,
    )
    own_item_id = att.get_attempt_learner_view(db_path, attempt["attempt_id"])["items"][0]["item_id"]
    foreign_item_id = att.get_attempt_learner_view(db_path, other_attempt["attempt_id"])["items"][0]["item_id"]
    att.save_response(db_path, attempt_id=attempt["attempt_id"], attempt_item_id=own_item_id,
                      selected_option_ids=["B"], client_revision=1)
    before_count = att.response_count(db_path, attempt["attempt_id"])

    with pytest.raises(TrainingError) as exc_info:
        att.save_response(db_path, attempt_id=attempt["attempt_id"], attempt_item_id=foreign_item_id,
                          selected_option_ids=["B"], client_revision=2)

    assert exc_info.value.code == ErrorCode.ATTEMPT_ITEM_NOT_FOUND
    assert att.response_count(db_path, attempt["attempt_id"]) == before_count
    assert att.get_attempt_learner_view(db_path, attempt["attempt_id"])["items"][0]["response"][
        "selected_option_ids"
    ] == ["B"]


def test_autosave_rejects_unknown_option_without_response_mutation(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    _, assignment_id = _make_open_exam_with_assignment(db_path)
    attempt = att.start_attempt(db_path, unit_code="son_tay", actor="learner1", assignment_id=assignment_id)
    item_id = att.get_attempt_learner_view(db_path, attempt["attempt_id"])["items"][0]["item_id"]
    att.save_response(db_path, attempt_id=attempt["attempt_id"], attempt_item_id=item_id,
                      selected_option_ids=["B"], client_revision=1)
    before_count = att.response_count(db_path, attempt["attempt_id"])

    with pytest.raises(TrainingError) as exc_info:
        att.save_response(db_path, attempt_id=attempt["attempt_id"], attempt_item_id=item_id,
                          selected_option_ids=["missing"], client_revision=2)

    assert exc_info.value.code == ErrorCode.INVALID_OPTION_SELECTION
    assert att.response_count(db_path, attempt["attempt_id"]) == before_count
    assert att.get_attempt_learner_view(db_path, attempt["attempt_id"])["items"][0]["response"][
        "selected_option_ids"
    ] == ["B"]


def test_autosave_rejects_multiple_options_for_single_choice_without_response_mutation(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    _, assignment_id = _make_open_exam_with_assignment(db_path)
    attempt = att.start_attempt(db_path, unit_code="son_tay", actor="learner1", assignment_id=assignment_id)
    item_id = att.get_attempt_learner_view(db_path, attempt["attempt_id"])["items"][0]["item_id"]
    att.save_response(db_path, attempt_id=attempt["attempt_id"], attempt_item_id=item_id,
                      selected_option_ids=["B"], client_revision=1)
    before_count = att.response_count(db_path, attempt["attempt_id"])

    with pytest.raises(TrainingError) as exc_info:
        att.save_response(db_path, attempt_id=attempt["attempt_id"], attempt_item_id=item_id,
                          selected_option_ids=["A", "B"], client_revision=2)

    assert exc_info.value.code == ErrorCode.SINGLE_CHOICE_REQUIRES_ONE_OPTION
    assert att.response_count(db_path, attempt["attempt_id"]) == before_count
    assert att.get_attempt_learner_view(db_path, attempt["attempt_id"])["items"][0]["response"][
        "selected_option_ids"
    ] == ["B"]


# --- M3.5: submit + scoring ---

def test_submit_creates_result(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    _, assignment_id = _make_open_exam_with_assignment(db_path)
    result = att.start_attempt(db_path, unit_code="son_tay", actor="learner1", assignment_id=assignment_id)
    view = att.get_attempt_learner_view(db_path, result["attempt_id"])
    for item in view["items"]:
        att.save_response(db_path, attempt_id=result["attempt_id"], attempt_item_id=item["item_id"],
                          selected_option_ids=["B"], client_revision=1)
    submit = att.submit_attempt(db_path, unit_code="son_tay", actor="learner1",
                                attempt_id=result["attempt_id"])
    assert submit["status"] == "submitted"
    assert submit["result"]["maximum_score"] == 3.0


def test_submit_is_idempotent(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    _, assignment_id = _make_open_exam_with_assignment(db_path)
    result = att.start_attempt(db_path, unit_code="son_tay", actor="learner1", assignment_id=assignment_id)
    view = att.get_attempt_learner_view(db_path, result["attempt_id"])
    for item in view["items"]:
        att.save_response(db_path, attempt_id=result["attempt_id"], attempt_item_id=item["item_id"],
                          selected_option_ids=["B"], client_revision=1)
    s1 = att.submit_attempt(db_path, unit_code="son_tay", actor="learner1", attempt_id=result["attempt_id"])
    s2 = att.submit_attempt(db_path, unit_code="son_tay", actor="learner1", attempt_id=result["attempt_id"])
    assert s1["result"]["score"] == s2["result"]["score"]


def test_scoring_from_snapshot(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    _, assignment_id = _make_open_exam_with_assignment(db_path)
    result = att.start_attempt(db_path, unit_code="son_tay", actor="learner1", assignment_id=assignment_id)
    view = att.get_attempt_learner_view(db_path, result["attempt_id"])
    # Q1,Q2 correct=B, Q3 correct=A. Answer all B → 2 correct.
    for item in view["items"]:
        att.save_response(db_path, attempt_id=result["attempt_id"], attempt_item_id=item["item_id"],
                          selected_option_ids=["B"], client_revision=1)
    submit = att.submit_attempt(db_path, unit_code="son_tay", actor="learner1", attempt_id=result["attempt_id"])
    assert submit["result"]["score"] == 2.0
    assert submit["result"]["percent"] == pytest.approx(66.67, abs=0.1)
    assert submit["result"]["passed"] is False


def test_submit_after_deadline_scores_saved_only(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    _, assignment_id = _make_open_exam_with_assignment(db_path)
    result = att.start_attempt(db_path, unit_code="son_tay", actor="learner1", assignment_id=assignment_id)
    view = att.get_attempt_learner_view(db_path, result["attempt_id"])
    for item in view["items"]:
        att.save_response(db_path, attempt_id=result["attempt_id"], attempt_item_id=item["item_id"],
                          selected_option_ids=["B"], client_revision=1)
    # simulate deadline passed
    conn = training_db.write_connection(db_path)
    conn.execute("UPDATE exam_attempts SET deadline_at_ms=? WHERE id=?",
                 (time_policy.utc_now_ms() - 1000, result["attempt_id"]))
    conn.commit()
    conn.close()
    submit = att.submit_attempt(db_path, unit_code="son_tay", actor="learner1", attempt_id=result["attempt_id"])
    assert submit["status"] == "timed_out"


def test_restart_does_not_lose_responses(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    _, assignment_id = _make_open_exam_with_assignment(db_path)
    result = att.start_attempt(db_path, unit_code="son_tay", actor="learner1", assignment_id=assignment_id)
    view = att.get_attempt_learner_view(db_path, result["attempt_id"])
    item_id = view["items"][0]["item_id"]
    att.save_response(db_path, attempt_id=result["attempt_id"], attempt_item_id=item_id,
                      selected_option_ids=["B"], client_revision=3)
    # simulate restart: reload view
    view2 = att.get_attempt_learner_view(db_path, result["attempt_id"])
    assert view2["items"][0]["response"]["client_revision"] == 3


def test_burst_submit_one_result(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    _, assignment_id = _make_open_exam_with_assignment(db_path)
    result = att.start_attempt(db_path, unit_code="son_tay", actor="learner1", assignment_id=assignment_id)
    view = att.get_attempt_learner_view(db_path, result["attempt_id"])
    for item in view["items"]:
        att.save_response(db_path, attempt_id=result["attempt_id"], attempt_item_id=item["item_id"],
                          selected_option_ids=["B"], client_revision=1)

    results = []
    barrier = threading.Barrier(5)

    def do_submit():
        barrier.wait()
        r = att.submit_attempt(db_path, unit_code="son_tay", actor="learner1", attempt_id=result["attempt_id"])
        results.append(r["result"]["score"])

    threads = [threading.Thread(target=do_submit) for _ in range(5)]
    for t in threads:
        t.start()
    for t in threads:
        t.join(timeout=15)

    conn = training_db.read_connection(db_path)
    count = conn.execute(
        "SELECT COUNT(*) AS c FROM exam_results WHERE attempt_id=?", (result["attempt_id"],)
    ).fetchone()["c"]
    conn.close()
    assert count == 1
    assert len(results) == 5


def test_close_administratively_submits_all_active_attempts_once(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    exam_id, first_assignment_id = _make_open_exam_with_assignment(
        db_path, additional_users=(
            {"username": "learner2", "display_name": "Learner 2"},
            {"username": "learner3", "display_name": "Learner 3"},
        ),
    )
    assignment_ids = [first_assignment_id, *[
        es.get_assignment_for_user(db_path, exam_id, username)[0]["id"]
        for username in ("learner2", "learner3")
    ]]
    attempt_ids = []
    for index, assignment_id in enumerate(assignment_ids, start=1):
        attempt = att.start_attempt(
            db_path, unit_code="son_tay", actor=f"learner{index}", assignment_id=assignment_id,
        )
        item_id = att.get_attempt_learner_view(db_path, attempt["attempt_id"])["items"][0]["item_id"]
        att.save_response(db_path, attempt_id=attempt["attempt_id"], attempt_item_id=item_id,
                          selected_option_ids=["B"], client_revision=1)
        attempt_ids.append(attempt["attempt_id"])

    es.close_exam(db_path, unit_code="son_tay", actor="mgr", exam_id=exam_id)
    es.close_exam(db_path, unit_code="son_tay", actor="mgr", exam_id=exam_id)

    conn = training_db.read_connection(db_path)
    try:
        attempts = conn.execute(
            "SELECT status, ended_reason FROM exam_attempts WHERE assignment_id IN (?, ?, ?) ORDER BY assignment_id",
            assignment_ids,
        ).fetchall()
        assert {(row["status"], row["ended_reason"]) for row in attempts} == {
            ("administratively_submitted", "exam_closed"),
        }
        assert conn.execute(
            "SELECT COUNT(*) AS c FROM exam_results WHERE attempt_id IN (?, ?, ?)", attempt_ids,
        ).fetchone()["c"] == 3
        assert conn.execute(
            "SELECT COUNT(*) AS c FROM training_audit_log WHERE action='close_exam' AND entity_id=?", (exam_id,),
        ).fetchone()["c"] == 1
        assert conn.execute(
            "SELECT COUNT(*) AS c FROM training_audit_log WHERE action='administratively_submit_attempt'"
        ).fetchone()["c"] == 3
    finally:
        conn.close()
    assert all(es.get_assignment(db_path, assignment_id)["status"] == "completed" for assignment_id in assignment_ids)


def test_close_racing_autosave_leaves_administrative_result(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    exam_id, assignment_id = _make_open_exam_with_assignment(db_path)
    started = att.start_attempt(db_path, unit_code="son_tay", actor="learner1", assignment_id=assignment_id)
    item_id = att.get_attempt_learner_view(db_path, started["attempt_id"])["items"][0]["item_id"]
    barrier = threading.Barrier(2)
    errors = []

    def autosave():
        barrier.wait()
        try:
            att.save_response(db_path, attempt_id=started["attempt_id"], attempt_item_id=item_id,
                              selected_option_ids=["B"], client_revision=1)
        except TrainingError as exc:
            errors.append(exc.code)

    def close():
        barrier.wait()
        es.close_exam(db_path, unit_code="son_tay", actor="mgr", exam_id=exam_id)

    threads = [threading.Thread(target=autosave), threading.Thread(target=close)]
    for thread in threads:
        thread.start()
    for thread in threads:
        thread.join(timeout=15)

    attempt = att.get_attempt(db_path, started["attempt_id"])
    assert attempt["status"] == "administratively_submitted"
    assert attempt["ended_reason"] == "exam_closed"
    assert att.get_result(db_path, started["attempt_id"]) is not None
    assert not errors or errors[0] in {"EXAM_NOT_OPEN", "ATTEMPT_ALREADY_COMPLETED"}


def test_closed_exam_blocks_start_autosave_and_learner_submit(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    exam_id, assignment_id = _make_open_exam_with_assignment(db_path)
    started = att.start_attempt(db_path, unit_code="son_tay", actor="learner1", assignment_id=assignment_id)
    item_id = att.get_attempt_learner_view(db_path, started["attempt_id"])["items"][0]["item_id"]
    es.close_exam(db_path, unit_code="son_tay", actor="mgr", exam_id=exam_id)

    with pytest.raises(TrainingError) as autosave_error:
        att.save_response(db_path, attempt_id=started["attempt_id"], attempt_item_id=item_id,
                          selected_option_ids=["B"], client_revision=1)
    with pytest.raises(TrainingError) as submit_error:
        att.submit_attempt(db_path, unit_code="son_tay", actor="learner1", attempt_id=started["attempt_id"])
    with pytest.raises(TrainingError) as start_error:
        att.start_attempt(db_path, unit_code="son_tay", actor="learner1", assignment_id=assignment_id)

    assert autosave_error.value.code == ErrorCode.ATTEMPT_ALREADY_COMPLETED
    assert submit_error.value.code == ErrorCode.ATTEMPT_ALREADY_COMPLETED
    assert start_error.value.code == ErrorCode.EXAM_NOT_OPEN


def test_learner_submit_rejects_every_administratively_submitted_attempt(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    _, assignment_id = _make_open_exam_with_assignment(db_path)
    started = att.start_attempt(db_path, unit_code="son_tay", actor="learner1", assignment_id=assignment_id)
    att.administratively_submit_attempt(
        db_path, unit_code="son_tay", actor="mgr", attempt_id=started["attempt_id"],
        ended_reason="recovery",
    )

    with pytest.raises(TrainingError) as exc_info:
        att.submit_attempt(db_path, unit_code="son_tay", actor="learner1", attempt_id=started["attempt_id"])

    assert exc_info.value.code == ErrorCode.ATTEMPT_ALREADY_COMPLETED


def test_autosave_rejects_response_at_exact_deadline(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    _, assignment_id = _make_open_exam_with_assignment(db_path)
    started = att.start_attempt(db_path, unit_code="son_tay", actor="learner1", assignment_id=assignment_id)
    item_id = att.get_attempt_learner_view(db_path, started["attempt_id"])["items"][0]["item_id"]
    deadline = 1_000_000
    conn = training_db.write_connection(db_path)
    try:
        conn.execute("UPDATE exam_attempts SET deadline_at_ms=? WHERE id=?", (deadline, started["attempt_id"]))
        conn.commit()
    finally:
        conn.close()
    monkeypatch.setattr(time_policy, "utc_now_ms", lambda: deadline)

    with pytest.raises(TrainingError) as exc_info:
        att.save_response(db_path, attempt_id=started["attempt_id"], attempt_item_id=item_id,
                          selected_option_ids=["B"], client_revision=1)

    assert exc_info.value.code == ErrorCode.ATTEMPT_EXPIRED


def test_close_retry_recovers_attempts_left_active_by_failed_close(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    exam_id, first_assignment_id = _make_open_exam_with_assignment(
        db_path, additional_users=({"username": "learner2", "display_name": "Learner 2"},),
    )
    second_assignment_id = es.get_assignment_for_user(db_path, exam_id, "learner2")[0]["id"]
    for actor, assignment_id in (("learner1", first_assignment_id), ("learner2", second_assignment_id)):
        att.start_attempt(db_path, unit_code="son_tay", actor=actor, assignment_id=assignment_id)

    original_submit = att.administratively_submit_attempt
    monkeypatch.setattr(att, "administratively_submit_attempt", lambda *_, **__: (_ for _ in ()).throw(RuntimeError("boom")))
    with pytest.raises(RuntimeError, match="boom"):
        es.close_exam(db_path, unit_code="son_tay", actor="mgr", exam_id=exam_id)
    monkeypatch.setattr(att, "administratively_submit_attempt", original_submit)

    es.close_exam(db_path, unit_code="son_tay", actor="mgr", exam_id=exam_id)

    conn = training_db.read_connection(db_path)
    try:
        assert conn.execute(
            "SELECT COUNT(*) AS c FROM exam_attempts WHERE status='active'"
        ).fetchone()["c"] == 0
        assert conn.execute(
            "SELECT COUNT(*) AS c FROM training_audit_log WHERE action='close_exam' AND entity_id=?", (exam_id,)
        ).fetchone()["c"] == 1
    finally:
        conn.close()


def test_administrative_submit_returns_existing_result_without_duplicate_audit(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    _, assignment_id = _make_open_exam_with_assignment(db_path)
    started = att.start_attempt(db_path, unit_code="son_tay", actor="learner1", assignment_id=assignment_id)

    first = att.administratively_submit_attempt(
        db_path, unit_code="son_tay", actor="mgr", attempt_id=started["attempt_id"],
    )
    second = att.administratively_submit_attempt(
        db_path, unit_code="son_tay", actor="mgr", attempt_id=started["attempt_id"],
    )

    conn = training_db.read_connection(db_path)
    try:
        audit_rows = conn.execute(
            "SELECT after_json FROM training_audit_log WHERE action='administratively_submit_attempt' "
            "AND entity_id=?", (started["attempt_id"],)
        ).fetchall()
    finally:
        conn.close()
    assert first == second
    assert first["attempt_id"] == started["attempt_id"]
    assert len(audit_rows) == 1
    assert json.loads(audit_rows[0]["after_json"])["ended_reason"] == "exam_closed"


@pytest.mark.parametrize(
    ("close_offset", "expected_status", "expected_reason"),
    [
        (-1, "administratively_submitted", "exam_closed"),
        (0, "timed_out", "timeout"),
        (1, "timed_out", "timeout"),
    ],
)
def test_close_classifies_active_attempt_by_deadline(monkeypatch, tmp_path, close_offset,
                                                     expected_status, expected_reason):
    db_path = _setup(monkeypatch, tmp_path)
    exam_id, assignment_id = _make_open_exam_with_assignment(db_path)
    started = att.start_attempt(db_path, unit_code="son_tay", actor="learner1", assignment_id=assignment_id)
    close_at = 2_000_000
    conn = training_db.write_connection(db_path)
    try:
        conn.execute("UPDATE exam_attempts SET deadline_at_ms=? WHERE id=?",
                     (close_at, started["attempt_id"]))
        conn.execute("UPDATE exam_events SET end_at_ms=? WHERE id=?", (close_at + 10_000, exam_id))
        conn.commit()
    finally:
        conn.close()
    monkeypatch.setattr(time_policy, "utc_now_ms", lambda: close_at + close_offset)

    es.close_exam(db_path, unit_code="son_tay", actor="mgr", exam_id=exam_id)

    attempt = att.get_attempt(db_path, started["attempt_id"])
    assert (attempt["status"], attempt["ended_reason"]) == (expected_status, expected_reason)
    assert att.get_result(db_path, started["attempt_id"]) is not None


def test_finalize_after_exam_end_times_out_active_attempt(monkeypatch, tmp_path):
    from services import training_report_service as reports

    db_path = _setup(monkeypatch, tmp_path)
    exam_id, assignment_id = _make_open_exam_with_assignment(db_path)
    started = att.start_attempt(db_path, unit_code="son_tay", actor="learner1", assignment_id=assignment_id)
    now = 2_000_000
    conn = training_db.write_connection(db_path)
    try:
        conn.execute("UPDATE exam_events SET end_at_ms=? WHERE id=?", (now, exam_id))
        conn.execute("UPDATE exam_attempts SET deadline_at_ms=? WHERE id=?", (now + 10_000, started["attempt_id"]))
        conn.commit()
    finally:
        conn.close()
    monkeypatch.setattr(time_policy, "utc_now_ms", lambda: now + 1)

    reports.finalize_exam(db_path, unit_code="son_tay", actor="mgr", exam_id=exam_id)

    attempt = att.get_attempt(db_path, started["attempt_id"])
    assert (attempt["status"], attempt["ended_reason"]) == ("timed_out", "timeout")


def test_close_mixed_expiry_assigns_one_terminal_result_per_attempt(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    exam_id, first_assignment_id = _make_open_exam_with_assignment(
        db_path, additional_users=({"username": "learner2", "display_name": "Learner 2"},),
    )
    second_assignment_id = es.get_assignment_for_user(db_path, exam_id, "learner2")[0]["id"]
    first = att.start_attempt(db_path, unit_code="son_tay", actor="learner1", assignment_id=first_assignment_id)
    second = att.start_attempt(db_path, unit_code="son_tay", actor="learner2", assignment_id=second_assignment_id)
    close_at = 2_000_000
    conn = training_db.write_connection(db_path)
    try:
        conn.execute("UPDATE exam_events SET end_at_ms=? WHERE id=?", (close_at + 10_000, exam_id))
        conn.execute("UPDATE exam_attempts SET deadline_at_ms=? WHERE id=?", (close_at - 1, first["attempt_id"]))
        conn.execute("UPDATE exam_attempts SET deadline_at_ms=? WHERE id=?", (close_at + 1, second["attempt_id"]))
        conn.commit()
    finally:
        conn.close()
    monkeypatch.setattr(time_policy, "utc_now_ms", lambda: close_at)

    es.close_exam(db_path, unit_code="son_tay", actor="mgr", exam_id=exam_id)
    es.close_exam(db_path, unit_code="son_tay", actor="mgr", exam_id=exam_id)

    assert (att.get_attempt(db_path, first["attempt_id"])["status"],
            att.get_attempt(db_path, first["attempt_id"])["ended_reason"]) == ("timed_out", "timeout")
    assert (att.get_attempt(db_path, second["attempt_id"])["status"],
            att.get_attempt(db_path, second["attempt_id"])["ended_reason"]) == (
                "administratively_submitted", "exam_closed")
    conn = training_db.read_connection(db_path)
    try:
        assert conn.execute("SELECT COUNT(*) AS count FROM exam_results WHERE attempt_id IN (?, ?)",
                            (first["attempt_id"], second["attempt_id"])).fetchone()["count"] == 2
    finally:
        conn.close()


def test_closed_exam_blocks_autosave_before_recovery_without_response_mutation(monkeypatch, tmp_path):
    db_path = _setup(monkeypatch, tmp_path)
    exam_id, assignment_id = _make_open_exam_with_assignment(db_path)
    started = att.start_attempt(db_path, unit_code="son_tay", actor="learner1", assignment_id=assignment_id)
    item_id = att.get_attempt_learner_view(db_path, started["attempt_id"])["items"][0]["item_id"]
    att.save_response(db_path, attempt_id=started["attempt_id"], attempt_item_id=item_id,
                      selected_option_ids=["A"], client_revision=1)
    closed = threading.Event()
    release = threading.Event()

    def pause_after_close(*_):
        closed.set()
        assert release.wait(timeout=10)

    monkeypatch.setattr(es, "_AFTER_CLOSE_COMMIT_HOOK", pause_after_close, raising=False)
    close_thread = threading.Thread(
        target=es.close_exam, kwargs={"db_path": db_path, "unit_code": "son_tay", "actor": "mgr", "exam_id": exam_id},
    )
    close_thread.start()
    assert closed.wait(timeout=10)

    with pytest.raises(TrainingError) as exc_info:
        att.save_response(db_path, attempt_id=started["attempt_id"], attempt_item_id=item_id,
                          selected_option_ids=["B"], client_revision=2)
    assert exc_info.value.code == ErrorCode.EXAM_NOT_OPEN
    response = att.get_attempt_learner_view(db_path, started["attempt_id"])["items"][0]["response"]
    assert response == {"selected_option_ids": ["A"], "client_revision": 1}

    release.set()
    close_thread.join(timeout=10)
    assert not close_thread.is_alive()
    assert att.get_attempt(db_path, started["attempt_id"])["status"] == "administratively_submitted"
