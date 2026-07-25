"""OpenAI production adapter: Responses API + Structured Outputs.

Lazy import: app/CLI không-AI vẫn chạy khi thiếu SDK/key.
Provider nhận source blocks thật (không chỉ document_version_id).
Evidence phải neo theo document_version_id + block_id trong selection.
Không lưu API key trong DB/log/commit.

Xử lý: timeout, refusal, malformed output, provider error, retry.
Model/provider/key/timeout chỉ đọc từ config/env.
"""

import json
import os

from training.errors import ErrorCode, TrainingError
from training.providers.base import BaseProvider, GenerationResult
from training.providers.schema_converter import get_openai_structured_schema


class OpenAIProviderError(TrainingError):
    """Provider-specific error for OpenAI adapter."""


def _is_ai_enabled():
    return os.getenv("DASHV4_TRAINING_AI_ENABLED", "0").strip().lower() in {
        "1", "true", "yes", "on",
    }


def _get_api_key():
    return os.getenv("OPENAI_API_KEY")


def _get_model():
    return os.getenv("DASHV4_TRAINING_AI_MODEL", "gpt-4o-mini")


def _get_timeout():
    return int(os.getenv("DASHV4_TRAINING_GENERATION_TIMEOUT_SECONDS", "120"))


class OpenAIProvider(BaseProvider):
    name = "openai"

    def __init__(self, *, api_key=None, model=None, timeout=None, client=None):
        """Khởi tạo provider. Raise nếu AI disabled hoặc thiếu key.

        Args:
            api_key: override OPENAI_API_KEY (test only)
            model: override DASHV4_TRAINING_AI_MODEL (test only)
            timeout: override DASHV4_TRAINING_GENERATION_TIMEOUT_SECONDS (test only)
            client: inject SDK client (test mock only)
        """
        if not _is_ai_enabled():
            raise OpenAIProviderError(
                ErrorCode.PROVIDER_ERROR,
                "AI disabled qua DASHV4_TRAINING_AI_ENABLED=0",
                status=503,
            )
        self._api_key = api_key or _get_api_key()
        if not self._api_key:
            raise OpenAIProviderError(
                ErrorCode.PROVIDER_ERROR,
                "OPENAI_API_KEY chưa set trong environment",
                status=503,
            )
        self._model = model or _get_model()
        self._timeout = timeout or _get_timeout()
        self._client = client
        self._schema = get_openai_structured_schema()

    def _get_client(self):
        """Lazy import OpenAI SDK chỉ khi cần."""
        if self._client is not None:
            return self._client
        try:
            from openai import OpenAI
        except ImportError:
            raise OpenAIProviderError(
                ErrorCode.PROVIDER_ERROR,
                "openai SDK chưa cài đặt; chạy: pip install openai",
                status=503,
            )
        self._client = OpenAI(api_key=self._api_key)
        return self._client

    def _build_prompt(self, *, snapshot, requested_count):
        """Build instructions với source blocks thật.

        Provider nhận blocks text + block_id + offsets để AI neo evidence.
        Prompt yêu cầu evidence chỉ từ block IDs được phép.
        """
        block_summaries = []
        for doc_ver in snapshot["document_versions"]:
            for blk in doc_ver["blocks"]:
                block_summaries.append(
                    f"  - document_version_id={doc_ver['document_version_id']}, "
                    f"block_id={blk['block_id']}, "
                    f"extraction_revision={blk['extraction_revision']}, "
                    f"content={blk['content'][:500]}"
                )
        blocks_text = "\n".join(block_summaries) if block_summaries else "(không có blocks)"
        audiences = ", ".join(snapshot.get("target_audience_codes", []))
        source_ids = ", ".join(snapshot["allowed_document_version_ids"])

        instructions = (
            "Bạn là chuyên gia đào tạo VNPT. Sinh câu hỏi trắc nghiệm tiếng Việt "
            "dựa trên nội dung được cung cấp.\n\n"
            "NGUYÊN TẮC BẮT BUỘC:\n"
            "1. Mỗi câu hỏi phải có đúng 1 evidence neo về document_version_id + block_id "
            "từ danh sách blocks dưới đây. KHÔNG được bịa evidence hoặc dùng block ngoài danh sách.\n"
            "2. Stem: rõ ràng, ngắn gọn, đúng quy định VNPT.\n"
            "3. Options: 4 phương án A/B/C/D, chỉ 1 đúng, phân biệt rõ.\n"
            "4. explanation: giải thích tại sao đáp án đúng.\n"
            "5. quoted_text trong evidence: trích dẫn chính xác từ content của block.\n"
            f"6. classification.audience_codes: phải nằm trong [{audiences}].\n"
            f"7. batch.source_document_version_ids: phải đúng [{source_ids}].\n"
            f"8. Sinh đúng {requested_count} câu hỏi.\n\n"
            f"DANH SÁCH BLOCKS NGUỒN (chỉ được dùng các block này):\n{blocks_text}\n"
        )
        return instructions

    def _call_openai(self, *, instructions, requested_count):
        """Gọi OpenAI Responses API với Structured Outputs.

        Returns: (parsed_json, usage_dict, model_str)
        Raises: OpenAIProviderError cho mọi lỗi provider.
        """
        client = self._get_client()
        try:
            response = client.responses.create(
                model=self._model,
                instructions=instructions,
                input=[{
                    "role": "user",
                    "content": (
                        f"Sinh {requested_count} câu hỏi trắc nghiệm "
                        "theo JSON schema đã cấu hình."
                    ),
                }],
                text={
                    "format": {
                        "type": "json_schema",
                        "name": "question_batch",
                        "schema": self._schema,
                        "strict": True,
                    }
                },
                timeout=self._timeout,
            )
        except Exception as exc:
            raise self._classify_error(exc)

        for item in response.output:
            if item.type == "refusal":
                raise OpenAIProviderError(
                    ErrorCode.PROVIDER_ERROR,
                    f"OpenAI từ chối sinh nội dung: {item.refusal}",
                    status=502,
                )

        output_text = response.output_text
        if not output_text:
            raise OpenAIProviderError(
                ErrorCode.PROVIDER_ERROR,
                "OpenAI trả response rỗng (không có output_text)",
                status=502,
            )

        try:
            payload = json.loads(output_text)
        except json.JSONDecodeError as exc:
            raise OpenAIProviderError(
                ErrorCode.VALIDATION_ERROR,
                f"OpenAI output không phải JSON hợp lệ: {exc}",
                status=502,
            )

        usage = {}
        if hasattr(response, "usage") and response.usage:
            usage = {
                "prompt_tokens": getattr(response.usage, "input_tokens", 0),
                "completion_tokens": getattr(response.usage, "output_tokens", 0),
                "total_tokens": getattr(response.usage, "total_tokens", 0),
            }
        return payload, usage

    def _classify_error(self, exc):
        """Phân loại exception thành OpenAIProviderError với code phù hợp."""
        msg = str(exc).lower()
        if "timeout" in msg or "timed out" in msg:
            return OpenAIProviderError(
                ErrorCode.PROVIDER_ERROR,
                f"OpenAI timeout sau {self._timeout}s: {exc}",
                status=504,
            )
        if "refusal" in msg or "content_filter" in msg:
            return OpenAIProviderError(
                ErrorCode.PROVIDER_ERROR,
                f"OpenAI từ chối nội dung: {exc}",
                status=502,
            )
        if "rate_limit" in msg or "429" in msg:
            return OpenAIProviderError(
                ErrorCode.PROVIDER_ERROR,
                f"OpenAI rate limit: {exc}",
                status=429,
            )
        if "authentication" in msg or "api key" in msg or "401" in msg:
            return OpenAIProviderError(
                ErrorCode.PROVIDER_ERROR,
                f"OpenAI auth error: {exc}",
                status=503,
            )
        return OpenAIProviderError(
            ErrorCode.PROVIDER_ERROR,
            f"OpenAI error: {exc}",
            status=502,
        )

    def _validate_evidence_in_selection(self, payload, snapshot):
        """Xác minh evidence chỉ neo theo blocks trong selection.

        Reject evidence ngoài selection (block_id hoặc document_version_id).
        """
        allowed_doc_ids = snapshot["allowed_document_version_ids"]
        allowed_block_ids = snapshot["allowed_block_ids"]
        questions = payload.get("questions", [])
        for qi, q in enumerate(questions):
            evidence = q.get("evidence", [])
            if not evidence:
                raise OpenAIProviderError(
                    ErrorCode.VALIDATION_ERROR,
                    f"question[{qi}]: thiếu evidence",
                    status=502,
                )
            for ei, ev in enumerate(evidence):
                doc_id = ev.get("document_version_id")
                block_id = ev.get("block_id")
                if doc_id not in allowed_doc_ids:
                    raise OpenAIProviderError(
                        ErrorCode.VALIDATION_ERROR,
                        f"question[{qi}].evidence[{ei}]: document_version_id "
                        f"{doc_id!r} không trong selection",
                        status=502,
                    )
                if block_id not in allowed_block_ids:
                    raise OpenAIProviderError(
                        ErrorCode.VALIDATION_ERROR,
                        f"question[{qi}].evidence[{ei}]: block_id "
                        f"{block_id!r} không trong selection",
                        status=502,
                    )

    def generate(self, *, source_document_version_ids, target_audience_codes,
                 requested_count, snapshot=None, **kwargs):
        """Generate questions via OpenAI.

        Args:
            snapshot: generation snapshot từ resolve_generation_snapshot.
                      Bắt buộc để provider nhận blocks thật.
        """
        if snapshot is None:
            raise OpenAIProviderError(
                ErrorCode.VALIDATION_ERROR,
                "OpenAI provider yêu cầu snapshot (blocks thật)",
                status=400,
            )
        if not snapshot.get("document_versions"):
            raise OpenAIProviderError(
                ErrorCode.VALIDATION_ERROR,
                "Snapshot không có document versions / blocks",
                status=400,
            )

        instructions = self._build_prompt(
            snapshot=snapshot, requested_count=requested_count,
        )

        payload, usage = self._call_openai(
            instructions=instructions, requested_count=requested_count,
        )

        self._validate_evidence_in_selection(payload, snapshot)

        return GenerationResult(
            batch=payload,
            provider="openai",
            model=self._model,
            prompt_version="1.0",
            usage=usage,
            raw_response=json.dumps(payload, ensure_ascii=False),
        )
