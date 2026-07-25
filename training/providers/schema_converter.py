"""OpenAI Structured Outputs schema converter.

OpenAI Structured Outputs (strict mode) yêu cầu:
- Tất cả properties phải là required
- additionalProperties: false ở mọi cấp
- Không dùng union types (type: ["string", "null"])
-Không dùng $ref/$defs trong strict mode (phải inline)

Module này convert schema nội bộ (Draft 2020-12) sang OpenAI strict schema
bằng cách:
1. Inline tất cả nullable fields thành required string (default "")
2. Thêm default cho optional fields
3. Đảm bảo additionalProperties: false
"""

import copy
import json
import os

_SCHEMA_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "schemas")


def _load_internal_schema(name):
    path = os.path.join(_SCHEMA_DIR, name)
    with open(path, "r", encoding="utf-8") as fh:
        return json.load(fh)


def _make_strict(node):
    """Đệ quy convert JSON schema node sang OpenAI Structured Outputs strict."""
    if not isinstance(node, dict):
        return node
    result = dict(node)
    if result.get("type") == "object" or "properties" in result:
        result["additionalProperties"] = False
        props = result.get("properties", {})
        for key, sub in props.items():
            props[key] = _make_strict(sub)
        if "required" not in result:
            result["required"] = list(props.keys())
        else:
            existing = set(result["required"])
            for key in props:
                existing.add(key)
            result["required"] = list(existing)
    elif result.get("type") == "array":
        if "items" in result:
            result["items"] = _make_strict(result["items"])
        result.pop("minItems", None)
        result.pop("maxItems", None)
    elif isinstance(result.get("type"), list):
        types = result["type"]
        if "null" in types:
            non_null = [t for t in types if t != "null"]
            if non_null:
                result["type"] = non_null[0]
                if non_null[0] == "string":
                    result.setdefault("default", "")
                elif non_null[0] == "integer":
                    result.setdefault("default", 0)
                elif non_null[0] == "object":
                    result.setdefault("default", {})
    result.pop("$schema", None)
    result.pop("$id", None)
    result.pop("title", None)
    result.pop("minLength", None)
    result.pop("minimum", None)
    result.pop("maximum", None)
    return result


def get_openai_structured_schema():
    """Trả JSON schema cho OpenAI Structured Outputs (strict mode).

    Schema này là nguồn chuẩn duy nhất — convert từ question_batch.schema.json
    nội bộ sang định dạng OpenAI strict.
    """
    internal = _load_internal_schema("question_batch.schema.json")
    return _make_strict(copy.deepcopy(internal))
