import copy
import pytest

from services import training_question_service as qs


VALID_SINGLE = {
    "schema_version": "1.0",
    "batch": {
        "title": "Câu hỏi C1.1 cho NVKT",
        "language": "vi",
        "source_document_version_ids": ["docver-001"],
        "target_audience_codes": ["nvkt"],
        "requested_count": 1,
    },
    "questions": [
        {
            "local_ref": "Q001",
            "type": "single_choice",
            "stem": "Thời gian sửa chữa được tính là bao nhiêu?",
            "stimulus": None,
            "options": [
                {"id": "A", "text": "9 giờ"},
                {"id": "B", "text": "1 giờ"},
                {"id": "C", "text": "30 phút"},
                {"id": "D", "text": "0 giờ"},
            ],
            "correct_option_ids": ["D"],
            "explanation": "Đáp án đúng là 0.",
            "distractor_rationales": {"A": "Sai", "B": "Sai", "C": "Sai"},
            "classification": {
                "domain_code": "quality",
                "indicator_codes": ["C1.1"],
                "topic_codes": ["repair_time_calculation"],
            },
            "difficulty": "medium",
            "cognitive_level": "apply",
            "criticality": "important",
            "estimated_seconds": 60,
            "evidence": [
                {
                    "document_version_id": "docver-001",
                    "block_id": "C1.1-B001",
                    "extraction_revision": 1,
                    "quoted_text": "thời gian tính bằng 0",
                    "supports": "correct_answer",
                }
            ],
        }
    ],
}


def test_valid_single_choice_passes():
    errors = qs.validate_question_batch(VALID_SINGLE)
    assert errors == []


def test_invalid_schema_missing_questions_fails():
    payload = copy.deepcopy(VALID_SINGLE)
    del payload["questions"]
    errors = qs.validate_question_batch(payload)
    assert len(errors) >= 1
    assert any("schema" in e.lower() or "questions" in e.lower() for e in errors)


def test_additional_properties_rejected():
    payload = copy.deepcopy(VALID_SINGLE)
    payload["batch"]["unexpected_field"] = "boom"
    errors = qs.validate_question_batch(payload)
    assert len(errors) >= 1


def test_duplicate_option_ids_fail():
    payload = copy.deepcopy(VALID_SINGLE)
    payload["questions"][0]["options"][1]["id"] = "A"
    errors = qs.validate_question_batch(payload)
    assert any("option" in e.lower() or "unique" in e.lower() for e in errors)


def test_single_choice_must_have_one_correct():
    payload = copy.deepcopy(VALID_SINGLE)
    payload["questions"][0]["correct_option_ids"] = ["A", "D"]
    errors = qs.validate_question_batch(payload)
    assert any("correct" in e.lower() or "single" in e.lower() for e in errors)


def test_correct_option_must_exist():
    payload = copy.deepcopy(VALID_SINGLE)
    payload["questions"][0]["correct_option_ids"] = ["Z"]
    errors = qs.validate_question_batch(payload)
    assert any("correct" in e.lower() or "exist" in e.lower() for e in errors)


def test_at_least_two_options_required():
    payload = copy.deepcopy(VALID_SINGLE)
    payload["questions"][0]["options"] = [{"id": "A", "text": "only"}]
    errors = qs.validate_question_batch(payload)
    assert any("option" in e.lower() for e in errors)


def test_evidence_required():
    payload = copy.deepcopy(VALID_SINGLE)
    payload["questions"][0]["evidence"] = []
    errors = qs.validate_question_batch(payload)
    assert any("evidence" in e.lower() for e in errors)


def test_evidence_block_must_match_batch_source():
    payload = copy.deepcopy(VALID_SINGLE)
    payload["questions"][0]["evidence"][0]["document_version_id"] = "docver-OTHER"
    errors = qs.validate_question_batch(payload)
    assert any("evidence" in e.lower() or "source" in e.lower() for e in errors)


def test_non_mvp_type_rejected():
    payload = copy.deepcopy(VALID_SINGLE)
    payload["questions"][0]["type"] = "multiple_choice"
    errors = qs.validate_question_batch(payload)
    assert any("type" in e.lower() or "single_choice" in e.lower() for e in errors)


def test_normalized_stem_hash_detects_duplicate():
    qs.reset_duplicate_cache()
    qs.record_normalized_stem("thời gian sửa chữa tính bằng 0 giờ", "qv1")
    is_dup = qs.is_duplicate_normalized_stem("THỜI GIAN sửa chữa tính bằng 0 giờ")
    assert is_dup is True


def test_normalized_stem_hash_no_false_positive():
    qs.reset_duplicate_cache()
    qs.record_normalized_stem("hoàn toàn khác", "qv1")
    assert qs.is_duplicate_normalized_stem("thời gian sửa chữa") is False
