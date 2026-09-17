"""Small, explicit contract for the first decimal-field collection reader."""
from __future__ import annotations

import re
from decimal import Decimal, InvalidOperation
from typing import Any


class ValidationError(ValueError):
    """Configuration, provenance or collection contract is invalid."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise ValidationError(message)


def nonempty(value: Any) -> bool:
    return isinstance(value, str) and bool(value.strip())


def decimal_text(value: Any) -> str:
    """No currency stripping, locale guessing, boolean coercion or rounding."""
    require(isinstance(value, (str, int, float, Decimal)) and not isinstance(value, bool),
            "Expected a finite decimal value")
    if isinstance(value, str):
        require(bool(re.fullmatch(r"[+-]?(?:\d+(?:\.\d*)?|\.\d+)", value.strip())),
                "Expected an unformatted decimal string")
    try:
        number = Decimal(str(value).strip())
    except (InvalidOperation, ValueError) as exc:
        raise ValidationError("Invalid decimal value") from exc
    require(number.is_finite(), "Expected a finite decimal value")
    return format(number, "f")


def schema_fields(schema: dict) -> dict[str, dict]:
    require(isinstance(schema, dict), "Schema must be an object")
    require(schema.get("formatVersion") == 0, "Unsupported schema formatVersion")
    require(nonempty(schema.get("schemaId")) and nonempty(schema.get("schemaVersion")),
            "Schema identity and version are required")
    require(schema.get("decimalEncoding") == "string" and schema.get("blankIsZero") is False,
            "Schema must preserve decimal strings and missing values")
    fields = schema.get("fields")
    require(isinstance(fields, list) and bool(fields), "Schema fields are required")
    result = {}
    for field in fields:
        require(isinstance(field, dict) and nonempty(field.get("field")), "Invalid field definition")
        key = field["field"]
        require(key not in result, "Duplicate schema field: " + key)
        require(field.get("type") == "decimal" and field.get("unit") == "currency",
                "Only decimal currency fields are supported: " + key)
        result[key] = field
    return result


def validate_collection(collection: dict, schema: dict) -> None:
    """Validate machine-readable results without declaring business agreement."""
    fields = schema_fields(schema)
    require(isinstance(collection, dict), "Collection must be an object")
    for key in ("schemaId", "schemaVersion"):
        require(collection.get(key) == schema[key], "Collection " + key + " mismatch")
    require(collection.get("engine") == "code", "Expected a code collection")
    require(collection.get("automaticPassEnabled") is False, "Reader cannot enable automatic PASS")
    digest = collection.get("resultSha256")
    require(isinstance(digest, str) and bool(re.fullmatch(r"[0-9a-f]{64}", digest)),
            "Result SHA-256 is required")
    require(nonempty(collection.get("period")) and bool(re.fullmatch(r"[A-Z]{3}", str(collection.get("currency", "")))),
            "Period and currency are required")
    records = collection.get("records")
    require(isinstance(records, list) and bool(records), "Collection records are required")
    entities = set()
    for record in records:
        require(isinstance(record, dict) and nonempty(record.get("entityKey")), "Entity key is required")
        require(record["entityKey"] not in entities, "Duplicate collection entity key")
        entities.add(record["entityKey"])
        values = record.get("fields")
        require(isinstance(values, list), "Record fields must be a list")
        seen = set()
        for item in values:
            require(isinstance(item, dict), "Invalid field result")
            key = item.get("field")
            require(isinstance(key, str) and key in fields and key not in seen,
                    "Unknown or duplicate collection field")
            seen.add(key)
            state = item.get("status")
            require(state in {"found", "missing", "unreadable"}, "Unsupported field status")
            if state == "found":
                require(isinstance(item.get("value"), str), "Found decimal must be a string")
                decimal_text(item["value"])
            else:
                require("value" in item and item["value"] is None, "Unavailable value must be null")
            require(item.get("currency") == collection["currency"], "Field currency mismatch")
            source = item.get("source")
            require(isinstance(source, dict) and source.get("fileSha256") == digest
                    and nonempty(source.get("sheet"))
                    and bool(re.fullmatch(r"[A-Z]{1,3}[1-9][0-9]*", str(source.get("cell", "")))),
                    "Invalid field provenance")
        require(seen == set(fields), "Every schema field must have an explicit result")
    issues = collection.get("issues")
    require(isinstance(issues, list) and all(isinstance(i, dict) and nonempty(i.get("code")) for i in issues),
            "Issues must contain structured error codes")
    ready = not issues and all(f["status"] == "found" for r in records for f in r["fields"])
    require(collection.get("readiness") == ("ready_for_comparison" if ready else "review_required"),
            "Collection readiness is inconsistent")
