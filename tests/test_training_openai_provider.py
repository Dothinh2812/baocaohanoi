"""Tests cho OpenAI provider: lazy import, Structured Outputs, refusal, timeout, evidence validation.

Mock SDK transport — không gọi API thật.
"""
import json
import os
from unittest.mock import MagicMock, patch

import pytest

from training.errors import ErrorCode, TrainingError
from training.providers.openai_provider import OpenAIProvider, OpenAIProviderError


def _make_snapshot(doc_ver_id="docver-1", block_id="DOC-B001"):
    return {
        "document_versions": [{
            "document_version_id": doc_ver_id,
            "blocks": [{
                "block_id": block_id,
                "extraction_revision": 1,
                "char_start": 0,
                "char_end": 100,
                "content": "Nội dung block test " * 5,
                "content_sha256": "abc123",
                "domain_code": "quality",
                "category_code": None,
            }],
            "topic_codes": ["brcd_repair"],
            "audience_codes": ["nvkt"],
        }],
        "allowed_block_ids": {block_id},
        "allowed_document_version_ids": {doc_ver_id},
        "target_audience_codes": ["nvkt"],
    }


def _make_valid_payload(doc_ver_id="docver-1", block_id="DOC-B001"):
    return {
        "schema_version": "1.0",
        "batch": {
            "title": "Test batch",
            "language": "vi",
            "source_document_version_ids": [doc_ver_id],
            "target_audience_codes": ["nvkt"],
            "requested_count": 1,
        },
        "questions": [{
            "local_ref": "Q1",
            "type": "single_choice",
            "stem": "Câu hỏi test?",
            "stimulus": None,
            "options": [
                {"id": "A", "text": "Sai"},
                {"id": "B", "text": "Đúng"},
            ],
            "correct_option_ids": ["B"],
            "explanation": "Đúng.",
            "distractor_rationales": {"A": "Sai"},
            "classification": {
                "domain_code": "quality",
                "topic_codes": ["brcd_repair"],
                "audience_codes": ["nvkt"],
            },
            "difficulty": "easy",
            "cognitive_level": "remember",
            "criticality": "normal",
            "estimated_seconds": 30,
            "calculation": None,
            "evidence": [{
                "document_version_id": doc_ver_id,
                "block_id": block_id,
                "extraction_revision": 1,
                "quoted_text": "Nội dung trích dẫn",
                "quote_start": None,
                "quote_end": None,
                "supports": "correct_answer",
            }],
        }],
    }


class _MockUsage:
    def __init__(self):
        self.input_tokens = 100
        self.output_tokens = 50
        self.total_tokens = 150


class _MockResponse:
    def __init__(self, payload, *, refusal=None):
        self.output_text = json.dumps(payload) if payload else ""
        self.usage = _MockUsage()
        if refusal:
            refusal_item = MagicMock()
            refusal_item.type = "refusal"
            refusal_item.refusal = refusal
            self.output = [refusal_item]
        else:
            text_item = MagicMock()
            text_item.type = "output_text"
            self.output = [text_item]


class _MockOpenAIClient:
    def __init__(self, response=None, error=None):
        self._response = response
        self._error = error
        self.responses = MagicMock()
        if error:
            self.responses.create.side_effect = error
        else:
            self.responses.create.return_value = response


class TestOpenAIProviderInit:
    def test_disabled_raises(self, monkeypatch):
        monkeypatch.delenv("OPENAI_API_KEY", raising=False)
        monkeypatch.setenv("DASHV4_TRAINING_AI_ENABLED", "0")
        with pytest.raises(OpenAIProviderError) as exc_info:
            OpenAIProvider()
        assert "disabled" in exc_info.value.message.lower()

    def test_missing_key_raises(self, monkeypatch):
        monkeypatch.setenv("DASHV4_TRAINING_AI_ENABLED", "1")
        monkeypatch.delenv("OPENAI_API_KEY", raising=False)
        with pytest.raises(OpenAIProviderError) as exc_info:
            OpenAIProvider()
        assert "key" in exc_info.value.message.lower()

    def test_enabled_with_key_works(self, monkeypatch):
        monkeypatch.setenv("DASHV4_TRAINING_AI_ENABLED", "1")
        monkeypatch.setenv("OPENAI_API_KEY", "sk-test-fake")
        provider = OpenAIProvider()
        assert provider.name == "openai"

    def test_lazy_import_sdk_missing(self, monkeypatch):
        monkeypatch.setenv("DASHV4_TRAINING_AI_ENABLED", "1")
        monkeypatch.setenv("OPENAI_API_KEY", "sk-test")
        provider = OpenAIProvider(api_key="sk-test")
        with patch("builtins.__import__", side_effect=ImportError("no openai")):
            with pytest.raises(OpenAIProviderError, match="SDK"):
                provider._get_client()


class TestOpenAIProviderGenerate:
    def test_valid_structured_output(self, monkeypatch):
        monkeypatch.setenv("DASHV4_TRAINING_AI_ENABLED", "1")
        payload = _make_valid_payload()
        mock_resp = _MockResponse(payload)
        mock_client = _MockOpenAIClient(response=mock_resp)
        provider = OpenAIProvider(api_key="sk-test", model="gpt-4o-mini",
                                  client=mock_client)
        snapshot = _make_snapshot()
        result = provider.generate(
            source_document_version_ids=["docver-1"],
            target_audience_codes=["nvkt"],
            requested_count=1,
            snapshot=snapshot,
        )
        assert result.provider == "openai"
        assert result.batch["questions"][0]["stem"] == "Câu hỏi test?"

    def test_refusal_handled(self, monkeypatch):
        monkeypatch.setenv("DASHV4_TRAINING_AI_ENABLED", "1")
        mock_resp = _MockResponse(None, refusal="Content policy violation")
        mock_client = _MockOpenAIClient(response=mock_resp)
        provider = OpenAIProvider(api_key="sk-test", client=mock_client)
        snapshot = _make_snapshot()
        with pytest.raises(OpenAIProviderError, match="từ chối"):
            provider.generate(
                source_document_version_ids=["docver-1"],
                target_audience_codes=["nvkt"],
                requested_count=1,
                snapshot=snapshot,
            )

    def test_timeout_handled(self, monkeypatch):
        monkeypatch.setenv("DASHV4_TRAINING_AI_ENABLED", "1")
        mock_client = _MockOpenAIClient(
            error=TimeoutError("Request timed out after 120s"),
        )
        provider = OpenAIProvider(api_key="sk-test", client=mock_client)
        snapshot = _make_snapshot()
        with pytest.raises(OpenAIProviderError, match="timeout"):
            provider.generate(
                source_document_version_ids=["docver-1"],
                target_audience_codes=["nvkt"],
                requested_count=1,
                snapshot=snapshot,
            )

    def test_rate_limit_handled(self, monkeypatch):
        monkeypatch.setenv("DASHV4_TRAINING_AI_ENABLED", "1")
        mock_client = _MockOpenAIClient(
            error=Exception("rate_limit_exceeded (429)"),
        )
        provider = OpenAIProvider(api_key="sk-test", client=mock_client)
        snapshot = _make_snapshot()
        with pytest.raises(OpenAIProviderError, match="rate limit"):
            provider.generate(
                source_document_version_ids=["docver-1"],
                target_audience_codes=["nvkt"],
                requested_count=1,
                snapshot=snapshot,
            )

    def test_malformed_json_output(self, monkeypatch):
        monkeypatch.setenv("DASHV4_TRAINING_AI_ENABLED", "1")
        mock_resp = MagicMock()
        mock_resp.output_text = "not valid json {{{"
        mock_resp.usage = _MockUsage()
        text_item = MagicMock()
        text_item.type = "output_text"
        mock_resp.output = [text_item]
        mock_client = _MockOpenAIClient(response=mock_resp)
        provider = OpenAIProvider(api_key="sk-test", client=mock_client)
        snapshot = _make_snapshot()
        with pytest.raises(OpenAIProviderError, match="JSON"):
            provider.generate(
                source_document_version_ids=["docver-1"],
                target_audience_codes=["nvkt"],
                requested_count=1,
                snapshot=snapshot,
            )

    def test_empty_response(self, monkeypatch):
        monkeypatch.setenv("DASHV4_TRAINING_AI_ENABLED", "1")
        mock_resp = MagicMock()
        mock_resp.output_text = ""
        mock_resp.usage = None
        mock_resp.output = []
        mock_client = _MockOpenAIClient(response=mock_resp)
        provider = OpenAIProvider(api_key="sk-test", client=mock_client)
        snapshot = _make_snapshot()
        with pytest.raises(OpenAIProviderError, match="rỗng"):
            provider.generate(
                source_document_version_ids=["docver-1"],
                target_audience_codes=["nvkt"],
                requested_count=1,
                snapshot=snapshot,
            )

    def test_evidence_outside_selection_rejected(self, monkeypatch):
        monkeypatch.setenv("DASHV4_TRAINING_AI_ENABLED", "1")
        payload = _make_valid_payload()
        payload["questions"][0]["evidence"][0]["block_id"] = "OUTSIDE-B999"
        mock_resp = _MockResponse(payload)
        mock_client = _MockOpenAIClient(response=mock_resp)
        provider = OpenAIProvider(api_key="sk-test", client=mock_client)
        snapshot = _make_snapshot()
        with pytest.raises(OpenAIProviderError, match="không trong selection"):
            provider.generate(
                source_document_version_ids=["docver-1"],
                target_audience_codes=["nvkt"],
                requested_count=1,
                snapshot=snapshot,
            )

    def test_evidence_wrong_doc_version_rejected(self, monkeypatch):
        monkeypatch.setenv("DASHV4_TRAINING_AI_ENABLED", "1")
        payload = _make_valid_payload()
        payload["questions"][0]["evidence"][0]["document_version_id"] = "docver-FAKE"
        mock_resp = _MockResponse(payload)
        mock_client = _MockOpenAIClient(response=mock_resp)
        provider = OpenAIProvider(api_key="sk-test", client=mock_client)
        snapshot = _make_snapshot()
        with pytest.raises(OpenAIProviderError, match="không trong selection"):
            provider.generate(
                source_document_version_ids=["docver-1"],
                target_audience_codes=["nvkt"],
                requested_count=1,
                snapshot=snapshot,
            )

    def test_missing_snapshot_raises(self, monkeypatch):
        monkeypatch.setenv("DASHV4_TRAINING_AI_ENABLED", "1")
        provider = OpenAIProvider(api_key="sk-test")
        with pytest.raises(OpenAIProviderError, match="snapshot"):
            provider.generate(
                source_document_version_ids=["docver-1"],
                target_audience_codes=["nvkt"],
                requested_count=1,
            )

    def test_no_secret_in_result(self, monkeypatch):
        monkeypatch.setenv("DASHV4_TRAINING_AI_ENABLED", "1")
        payload = _make_valid_payload()
        mock_resp = _MockResponse(payload)
        mock_client = _MockOpenAIClient(response=mock_resp)
        provider = OpenAIProvider(api_key="sk-secret-key-do-not-leak",
                                  client=mock_client)
        snapshot = _make_snapshot()
        result = provider.generate(
            source_document_version_ids=["docver-1"],
            target_audience_codes=["nvkt"],
            requested_count=1,
            snapshot=snapshot,
        )
        result_str = json.dumps(result.__dict__, ensure_ascii=False, default=str)
        assert "sk-secret" not in result_str
        assert "sk-secret" not in json.dumps(result.batch, ensure_ascii=False)
        assert "sk-secret" not in json.dumps(result.raw_response or "", ensure_ascii=False)

    def test_structured_outputs_schema_used(self, monkeypatch):
        monkeypatch.setenv("DASHV4_TRAINING_AI_ENABLED", "1")
        payload = _make_valid_payload()
        mock_resp = _MockResponse(payload)
        mock_client = _MockOpenAIClient(response=mock_resp)
        provider = OpenAIProvider(api_key="sk-test", client=mock_client)
        snapshot = _make_snapshot()
        provider.generate(
            source_document_version_ids=["docver-1"],
            target_audience_codes=["nvkt"],
            requested_count=1,
            snapshot=snapshot,
        )
        call_kwargs = mock_client.responses.create.call_args.kwargs
        assert "text" in call_kwargs
        assert call_kwargs["text"]["format"]["type"] == "json_schema"
        assert call_kwargs["text"]["format"]["strict"] is True
        assert "schema" in call_kwargs["text"]["format"]


class TestSchemaConverter:
    def test_makes_all_required(self):
        from training.providers.schema_converter import get_openai_structured_schema
        schema = get_openai_structured_schema()

        def _check(node, path="root"):
            if isinstance(node, dict):
                if node.get("type") == "object" and "properties" in node:
                    props = set(node["properties"].keys())
                    required = set(node.get("required", []))
                    missing = props - required
                    assert not missing, f"{path}: missing required: {missing}"
                    assert node.get("additionalProperties") is False
                    for k, v in node["properties"].items():
                        _check(v, f"{path}.{k}")
                if node.get("type") == "array" and "items" in node:
                    _check(node["items"], f"{path}[]")

        _check(schema)

    def test_no_nullable_union(self):
        from training.providers.schema_converter import get_openai_structured_schema

        def _check(node, path="root"):
            if isinstance(node, dict):
                if isinstance(node.get("type"), list):
                    pytest.fail(f"{path}: nullable union type not allowed: {node['type']}")
                for k, v in node.get("properties", {}).items():
                    _check(v, f"{path}.{k}")
                if "items" in node:
                    _check(node["items"], f"{path}[]")

        _check(get_openai_structured_schema())
