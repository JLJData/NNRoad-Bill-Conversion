"""Deterministic Code-vs-AI comparison; no workflow approval side effects."""
from __future__ import annotations

from decimal import Decimal

from ai_collection.contracts import validate_ai_collection

from .contracts import decimal_text, require, schema_fields, validate_collection


def _name(value) -> str:
    return " ".join(str(value or "").split()).casefold()


def _field_map(record: dict) -> dict:
    return {item["field"]: item for item in record["fields"]}


def _evidence(item: dict | None):
    return item.get("source") if isinstance(item, dict) else None


def compare_collections(code: dict, ai: dict, schema: dict, documents: list[dict]) -> dict:
    """Associate only unique exact normalized names, then compare configured decimals.

    MATCH means the supplied records agree under diagnostic rules. It is never an
    approval: automaticPassEnabled and workflowDecision remain disabled.
    """
    fields = schema_fields(schema)
    validate_collection(code, schema)
    validate_ai_collection(ai, schema, documents)
    require(code["period"] == ai["period"], "Code and AI periods differ")
    require(code["currency"] == ai["currency"], "Code and AI currencies differ")

    result = {
        "schemaId": schema["schemaId"],
        "schemaVersion": schema["schemaVersion"],
        "period": code["period"],
        "currency": code["currency"],
        "associationMethod": "unique_exact_normalized_name_v1",
        "automaticPassEnabled": False,
        "workflowDecision": "NOT_EVALUATED",
        "differences": [],
        "matches": [],
        "summary": {},
    }

    def difference(kind: str, **details):
        result["differences"].append({"kind": kind, **details})

    code_names = {}
    for record in code["records"]:
        key = _name(record.get("observedName"))
        if not key:
            difference("CODE_ENTITY_UNAVAILABLE", entityKey=record["entityKey"])
            continue
        code_names.setdefault(key, []).append(record)
    ai_names = {}
    for index, record in enumerate(ai["records"]):
        key = _name(record["sourceEntity"].get("name"))
        if not key:
            difference("AI_ENTITY_UNAVAILABLE", aiRecordIndex=index,
                       sourceEntity=record["sourceEntity"])
            continue
        ai_names.setdefault(key, []).append((index, record))

    paired_code, paired_ai = set(), set()
    for key in sorted(set(code_names) | set(ai_names)):
        left, right = code_names.get(key, []), ai_names.get(key, [])
        if len(left) > 1 or len(right) > 1:
            difference("AMBIGUOUS_ENTITY", normalizedName=key,
                       codeEntityKeys=[record["entityKey"] for record in left],
                       aiRecordIndexes=[index for index, _ in right])
            continue
        if not left:
            index, record = right[0]
            difference("EXTRA_AI_ENTITY", aiRecordIndex=index, sourceEntity=record["sourceEntity"])
            continue
        if not right:
            difference("MISSING_AI_ENTITY", entityKey=left[0]["entityKey"],
                       observedName=left[0].get("observedName"))
            continue
        code_record = left[0]
        ai_index, ai_record = right[0]
        paired_code.add(code_record["entityKey"])
        paired_ai.add(ai_index)
        code_values, ai_values = _field_map(code_record), _field_map(ai_record)
        for field_name, field_schema in fields.items():
            code_item, ai_item = code_values[field_name], ai_values[field_name]
            common = {
                "entityKey": code_record["entityKey"],
                "observedName": code_record.get("observedName"),
                "field": field_name,
                "codeStatus": code_item["status"],
                "aiStatus": ai_item["status"],
                "codeValue": code_item.get("value"),
                "aiValue": ai_item.get("value"),
                "codeSource": _evidence(code_item),
                "aiSource": _evidence(ai_item),
            }
            if code_item["status"] != "found" or ai_item["status"] != "found":
                difference("FIELD_UNAVAILABLE", **common)
                continue
            comparison = field_schema.get("comparison") if isinstance(field_schema.get("comparison"), dict) else {}
            mode = comparison.get("mode", "compare_reported_fact")
            tolerance = Decimal(decimal_text(comparison.get("diagnosticTolerance", "0")))
            require(tolerance >= 0, "Diagnostic tolerance cannot be negative: " + field_name)
            code_number = Decimal(code_item["value"])
            ai_number = Decimal(ai_item["value"])
            delta = abs(code_number - ai_number)
            if delta > tolerance:
                difference("VALUE_MISMATCH", difference=decimal_text(delta),
                           diagnosticTolerance=decimal_text(tolerance), comparisonMode=mode, **common)
            elif mode != "compare_reported_fact":
                difference("FIELD_REQUIRES_REVIEW", difference=decimal_text(delta),
                           diagnosticTolerance=decimal_text(tolerance), comparisonMode=mode, **common)
            else:
                result["matches"].append({"entityKey": code_record["entityKey"], "field": field_name,
                                          "value": code_item["value"], "difference": decimal_text(delta)})

    for item in ai["unknownItems"]:
        difference("UNKNOWN_AI_ITEM", label=item["label"], aiSource=item["source"])
    for issue in code["issues"]:
        difference("CODE_COLLECTION_ISSUE", issue=issue)
    for issue in ai["issues"]:
        difference("AI_COLLECTION_ISSUE", issue=issue)

    kinds = {item["kind"] for item in result["differences"]}
    hard = {"VALUE_MISMATCH", "MISSING_AI_ENTITY", "EXTRA_AI_ENTITY"}
    if kinds & hard:
        status = "MISMATCH"
    elif kinds:
        status = "REVIEW_REQUIRED"
    else:
        status = "MATCH"
    result["comparisonStatus"] = status
    result["summary"] = {
        "codeEntityCount": len(code["records"]),
        "aiEntityCount": len(ai["records"]),
        "associatedEntityCount": len(paired_code),
        "matchedFieldCount": len(result["matches"]),
        "differenceCount": len(result["differences"]),
        "unknownItemCount": len(ai["unknownItems"]),
    }
    return result
