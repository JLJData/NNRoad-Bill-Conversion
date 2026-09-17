"""Contracts that keep AI extraction independent from the code result."""
from __future__ import annotations

import copy
import re
from typing import Any

from bill_validation.contracts import decimal_text, nonempty, require, schema_fields


_SHA256_RE = re.compile(r"[0-9a-f]{64}")
_FORBIDDEN_REQUEST_KEYS = {
    "codecollection",
    "coderesult",
    "convertedfile",
    "differences",
    "expectedanswer",
    "expectedvalue",
    "resultsha256",
    "validationresult",
}


def _walk_keys(value: Any):
    if isinstance(value, dict):
        for key, child in value.items():
            yield str(key).replace("_", "").replace("-", "").casefold()
            yield from _walk_keys(child)
    elif isinstance(value, list):
        for child in value:
            yield from _walk_keys(child)


def _validate_documents(documents: Any) -> dict[str, dict]:
    require(isinstance(documents, list) and bool(documents), "Original documents are required")
    result = {}
    for document in documents:
        require(isinstance(document, dict), "Invalid original document")
        file_id = document.get("fileId")
        require(nonempty(file_id) and file_id not in result, "Document fileId must be unique")
        require(bool(_SHA256_RE.fullmatch(str(document.get("sha256", "")))), "Document SHA-256 is required")
        require(nonempty(document.get("mediaType")) and nonempty(document.get("sourceRef")),
                "Document mediaType and sourceRef are required")
        result[file_id] = document
    return result


def _model_schema(schema: dict) -> dict:
    """Return semantic extraction fields without code binding or expected values."""
    fields = schema_fields(schema)
    return {
        "schemaId": schema["schemaId"],
        "schemaVersion": schema["schemaVersion"],
        "decimalEncoding": schema["decimalEncoding"],
        "blankIsZero": schema["blankIsZero"],
        "fields": [
            {
                key: copy.deepcopy(field[key])
                for key in ("field", "labelZh", "type", "unit", "sourceKind", "requiredOutputEntry", "aliases")
                if key in field
            }
            for field in fields.values()
        ],
    }


def build_collection_request(*, run_id: str, documents: list[dict], schema: dict,
                             instructions: list[str] | None = None, period: str, currency: str) -> dict:
    """Build the only request shape accepted by providers.

    sourceRef is an opaque server-side reference. Provider adapters resolve it to the
    original document; it is not a converted workbook or code-derived intermediate.
    """
    request = {
        "requestVersion": 1,
        "runId": run_id,
        "documents": copy.deepcopy(documents),
        "schema": _model_schema(schema),
        "instructions": copy.deepcopy(instructions or []),
        "period": period,
        "currency": currency,
    }
    validate_collection_request(request)
    return request


def validate_collection_request(request: dict) -> None:
    require(isinstance(request, dict), "AI collection request must be an object")
    require(set(request) == {"requestVersion", "runId", "documents", "schema", "instructions", "period", "currency"},
            "AI collection request contains unsupported inputs")
    require(request.get("requestVersion") == 1 and nonempty(request.get("runId")), "Invalid request identity")
    require(not (_FORBIDDEN_REQUEST_KEYS & set(_walk_keys(request))), "AI request contains a code-result input")
    _validate_documents(request.get("documents"))
    fields = request.get("schema")
    require(isinstance(fields, dict) and nonempty(fields.get("schemaId")) and nonempty(fields.get("schemaVersion")),
            "AI request schema identity is required")
    require(isinstance(fields.get("fields"), list) and bool(fields["fields"]), "AI request fields are required")
    require(isinstance(request.get("instructions"), list)
            and all(nonempty(item) for item in request["instructions"]), "Instructions must be non-empty strings")
    require(nonempty(request.get("period")) and bool(re.fullmatch(r"[A-Z]{3}", str(request.get("currency", "")))),
            "Request period and currency are required")


def _validate_source(source: Any, documents: dict[str, dict]) -> None:
    require(isinstance(source, dict) and source.get("fileId") in documents, "AI source fileId is invalid")
    require(nonempty(source.get("location")), "AI source location is required")
    if "page" in source:
        require(type(source["page"]) is int and source["page"] >= 1, "AI source page is invalid")
    if "rawText" in source:
        require(isinstance(source["rawText"], str), "AI rawText must be a string")


def validate_ai_collection(collection: dict, schema: dict, documents: list[dict], *, run_id: str | None = None,
                           provider_id: str | None = None) -> None:
    """Validate structured extraction; this does not compare or authorize a bill."""
    fields = schema_fields(schema)
    document_map = _validate_documents(documents)
    require(isinstance(collection, dict) and collection.get("engine") == "ai", "Expected an AI collection")
    for key in ("schemaId", "schemaVersion"):
        require(collection.get(key) == schema[key], "AI collection " + key + " mismatch")
    if run_id is not None:
        require(collection.get("runId") == run_id, "AI collection runId mismatch")
    else:
        require(nonempty(collection.get("runId")), "AI collection runId is required")
    require(nonempty(collection.get("provider")), "AI provider is required")
    if provider_id is not None:
        require(collection.get("provider") == provider_id, "AI provider identity mismatch")
    require(nonempty(collection.get("model")), "AI model is required")
    require(nonempty(collection.get("period")) and bool(re.fullmatch(r"[A-Z]{3}", str(collection.get("currency", "")))),
            "AI period and currency are required")
    require(collection.get("automaticPassEnabled") is False, "AI collection cannot enable automatic PASS")
    records = collection.get("records")
    require(isinstance(records, list), "AI records must be a list")
    for record in records:
        require(isinstance(record, dict), "Invalid AI record")
        identity = record.get("sourceEntity")
        require(isinstance(identity, dict) and (nonempty(identity.get("employeeId")) or nonempty(identity.get("name"))),
                "AI source entity identity is required")
        values = record.get("fields")
        require(isinstance(values, list), "AI record fields must be a list")
        seen = set()
        for item in values:
            require(isinstance(item, dict) and item.get("field") in fields and item["field"] not in seen,
                    "Unknown or duplicate AI field")
            seen.add(item["field"])
            status = item.get("status")
            require(status in {"found", "missing", "unreadable", "not_applicable"}, "Unsupported AI field status")
            if status == "found":
                require(isinstance(item.get("value"), str), "AI decimal must be a string")
                decimal_text(item["value"])
            else:
                require("value" in item and item["value"] is None, "Unavailable AI value must be null")
            require(item.get("currency") == collection["currency"], "AI field currency mismatch")
            _validate_source(item.get("source"), document_map)
            if "confidence" in item:
                confidence = item["confidence"]
                require(isinstance(confidence, (int, float)) and not isinstance(confidence, bool)
                        and 0 <= confidence <= 1, "AI confidence must be between 0 and 1")
        require(seen == set(fields), "Every schema field must have an explicit AI result")
    unknown = collection.get("unknownItems")
    require(isinstance(unknown, list), "AI unknownItems must be a list")
    for item in unknown:
        require(isinstance(item, dict) and nonempty(item.get("label")), "Unknown item label is required")
        _validate_source(item.get("source"), document_map)
    issues = collection.get("issues")
    require(isinstance(issues, list) and all(isinstance(item, dict) and nonempty(item.get("code")) for item in issues),
            "AI issues must contain structured codes")
    ready = bool(records) and not issues and all(nonempty(record["sourceEntity"].get("name")) for record in records)
    require(collection.get("readiness") == ("ready_for_association" if ready else "review_required"),
            "AI readiness is inconsistent")
