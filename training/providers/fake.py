"""Fake AI provider — trả fixture JSON, không gọi mạng.

Dùng cho test và chế độ AI disabled. Trả batch hợp lệ schema.
"""

from training.providers.base import BaseProvider, GenerationResult


def _make_question(doc_ver_id, audience, index):
    return {
        "local_ref": f"FAKE-{index:03d}",
        "type": "single_choice",
        "stem": f"Câu hỏi mẫu {index} cho {audience}?",
        "stimulus": None,
        "options": [
            {"id": "A", "text": "Phương án A"},
            {"id": "B", "text": "Phương án B"},
            {"id": "C", "text": "Phương án C"},
            {"id": "D", "text": "Phương án D"},
        ],
        "correct_option_ids": ["B"],
        "explanation": "Đáp án B là đúng theo quy định mẫu.",
        "distractor_rationales": {
            "A": "Phương án A sai.",
            "C": "Phương án C sai.",
            "D": "Phương án D sai.",
        },
        "classification": {
            "domain_code": "quality",
            "topic_codes": ["sample_topic"],
            "audience_codes": [audience],
        },
        "difficulty": "medium",
        "cognitive_level": "remember",
        "criticality": "normal",
        "estimated_seconds": 45,
        "evidence": [
            {
                "document_version_id": doc_ver_id,
                "block_id": "DOC-B001",
                "extraction_revision": 1,
                "quoted_text": "Nội dung mẫu cho bằng chứng.",
                "supports": "correct_answer",
            }
        ],
    }


class FakeProvider(BaseProvider):
    name = "fake"

    def generate(self, *, source_document_version_ids, target_audience_codes,
                 requested_count, **kwargs):
        doc_ver_id = source_document_version_ids[0] if source_document_version_ids else "docver-fake"
        audience = target_audience_codes[0] if target_audience_codes else "nvkt"
        count = max(1, min(requested_count or 1, 10))
        questions = [_make_question(doc_ver_id, audience, i + 1) for i in range(count)]
        batch = {
            "schema_version": "1.0",
            "batch": {
                "title": f"Fake batch cho {audience}",
                "language": "vi",
                "source_document_version_ids": source_document_version_ids,
                "target_audience_codes": target_audience_codes,
                "requested_count": count,
            },
            "questions": questions,
        }
        return GenerationResult(
            batch=batch, provider="fake", model="fake-model",
            prompt_version="1.0", usage={"total_tokens": 0},
            raw_response=None,
        )


def generate(*, source_document_version_ids, target_audience_codes, requested_count, **kwargs):
    """Hàm tiện ích dùng FakeProvider."""
    return FakeProvider().generate(
        source_document_version_ids=source_document_version_ids,
        target_audience_codes=target_audience_codes,
        requested_count=requested_count,
        **kwargs,
    ).batch
