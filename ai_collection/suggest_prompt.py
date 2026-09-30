"""Suggest profile aiInstructions from CODE/AI diffs and a user request."""
from __future__ import annotations

import json
import os
from typing import Any, Callable

from bill_validation.contracts import nonempty, require

from .provider import AIProviderError, _default_http_post

_MAX_DIFFS = 80
_MAX_UNMATCHED = 40
_MAX_INSTRUCTION_CHARS = 2000
_MAX_INSTRUCTIONS = 20
_MAX_USER_REQUEST_CHARS = 4000
_MAX_EXISTING_INSTRUCTIONS = 50


def _trim_text(value: Any, limit: int) -> str:
    text = str(value or "").strip()
    if len(text) <= limit:
        return text
    return text[: max(0, limit - 1)] + "…"


def _compact_diffs(diffs: list[dict] | None) -> list[dict]:
    out: list[dict] = []
    for item in diffs or []:
        if not isinstance(item, dict):
            continue
        out.append({
            "type": _trim_text(item.get("type"), 40),
            "employeeName": _trim_text(item.get("employeeName"), 120),
            "fieldLabel": _trim_text(item.get("fieldLabel"), 200),
            "codeCell": _trim_text(item.get("codeCell"), 20),
            "codeValue": _trim_text(item.get("codeValue"), 120),
            "codeFormula": bool(item.get("codeFormula")),
            "aiCell": _trim_text(item.get("aiCell"), 20),
            "aiValue": _trim_text(item.get("aiValue"), 120),
            "aiFormula": bool(item.get("aiFormula")),
        })
        if len(out) >= _MAX_DIFFS:
            break
    return out


def _compact_unmatched(items: list[dict] | None) -> list[dict]:
    out: list[dict] = []
    for item in items or []:
        if not isinstance(item, dict):
            continue
        out.append({
            "name": _trim_text(item.get("name"), 120),
            "row": item.get("row"),
            "suggestedName": _trim_text(item.get("suggestedName"), 120) or None,
            "similarity": item.get("similarity"),
        })
        if len(out) >= _MAX_UNMATCHED:
            break
    return out


def _normalize_existing_instructions(raw: Any) -> list[str]:
    lines: list[str] = []
    if isinstance(raw, list):
        source = raw
    elif isinstance(raw, str):
        source = raw.splitlines()
    else:
        source = []
    seen: set[str] = set()
    for item in source:
        text = str(item or "").strip()
        if not text or text in seen:
            continue
        seen.add(text)
        lines.append(_trim_text(text, _MAX_INSTRUCTION_CHARS))
        if len(lines) >= _MAX_EXISTING_INSTRUCTIONS:
            break
    return lines


def _output_schema() -> dict:
    return {
        "type": "object",
        "additionalProperties": False,
        "required": ["explanation", "suggestedInstructions"],
        "properties": {
            "explanation": {
                "type": "string",
                "description": "Short rationale for the suggested prompts.",
            },
            "suggestedInstructions": {
                "type": "array",
                "maxItems": _MAX_INSTRUCTIONS,
                "items": {
                    "type": "string",
                    "maxLength": _MAX_INSTRUCTION_CHARS,
                },
            },
        },
    }


def build_suggest_prompt_payload(
    *,
    user_request: str,
    current_instructions: list[str] | None,
    comparison: dict,
    model: str,
    reasoning_effort: str = "medium",
) -> dict:
    """Build the OpenAI Responses API payload for prompt suggestion."""
    require(nonempty(user_request), "userRequest is required")
    require(len(user_request.strip()) <= _MAX_USER_REQUEST_CHARS, "userRequest is too long")
    require(nonempty(model), "model is required")
    require(
        reasoning_effort in {"none", "low", "medium", "high", "xhigh", "max"},
        "Unsupported OpenAI reasoning effort",
    )
    require(isinstance(comparison, dict), "comparison must be an object")

    existing = _normalize_existing_instructions(current_instructions)
    diffs = _compact_diffs(comparison.get("diffs") if isinstance(comparison.get("diffs"), list) else [])
    unmatched_code = _compact_unmatched(
        comparison.get("unmatchedCode") if isinstance(comparison.get("unmatchedCode"), list) else []
    )
    unmatched_ai = _compact_unmatched(
        comparison.get("unmatchedAi") if isinstance(comparison.get("unmatchedAi"), list) else []
    )

    context = {
        "task": "suggest_ai_instructions",
        "role": (
            "You help an office operator write short profileInstructions for an independent "
            "AI payroll bill filler. Those instructions are injected into a later AI run that "
            "reads only original supplier bills and fills the template last -L sheet."
        ),
        "rules": [
            "Follow the operator's request as the primary intent. Do not assume CODE is always correct or AI is always wrong unless the request says so.",
            "Use CODE/AI cell differences and unmatched employees only as evidence. Prefer concrete, reusable rules (field mapping, summing, skip, naming, arrears/backfill).",
            "Each suggestedInstructions item becomes one line in mapping.aiInstructions. Keep each line self-contained, actionable, and ideally English for model stability.",
            "Do not invent employees or amounts not supported by the evidence. Avoid generic advice like 'be careful'.",
            "Avoid duplicating currentInstructions unless you are refining them. Return 1-12 lines when possible; empty list only if nothing useful can be said.",
            "explanation should briefly tell the operator what pattern you noticed and why the lines help.",
        ],
        "userRequest": user_request.strip(),
        "currentInstructions": existing,
        "comparisonSummary": {
            "codeSheet": _trim_text(comparison.get("codeSheet"), 80),
            "aiSheet": _trim_text(comparison.get("aiSheet"), 80),
            "differenceCount": comparison.get("differenceCount"),
            "matchedEmployeeCount": comparison.get("matchedEmployeeCount"),
            "independentComparison": bool(comparison.get("independentComparison")),
            "diffTruncated": len(diffs) >= _MAX_DIFFS,
        },
        "diffs": diffs,
        "unmatchedCode": unmatched_code,
        "unmatchedAi": unmatched_ai,
    }
    return {
        "model": model,
        "reasoning": {"effort": reasoning_effort},
        "store": False,
        "input": [{
            "role": "user",
            "content": [{"type": "input_text", "text": json.dumps(context, ensure_ascii=False)}],
        }],
        "text": {
            "format": {
                "type": "json_schema",
                "name": "ai_instruction_suggestion",
                "strict": True,
                "schema": _output_schema(),
            }
        },
    }


def _parse_structured_response(response: dict) -> dict:
    if response.get("status") != "completed":
        raise AIProviderError("OpenAI response did not complete: " + str(response.get("status")))
    for output in response.get("output") or []:
        if output.get("type") != "message":
            continue
        for content in output.get("content") or []:
            if content.get("type") == "refusal":
                raise AIProviderError("OpenAI refused the prompt suggestion request")
            if content.get("type") == "output_text" and isinstance(content.get("text"), str):
                try:
                    return json.loads(content["text"])
                except json.JSONDecodeError as exc:
                    raise AIProviderError("OpenAI returned invalid structured JSON") from exc
    raise AIProviderError("OpenAI response contains no structured output")


def normalize_suggestion_result(raw: dict, *, model: str) -> dict:
    require(isinstance(raw, dict), "suggestion result must be an object")
    explanation = _trim_text(raw.get("explanation"), 4000)
    lines: list[str] = []
    seen: set[str] = set()
    for item in raw.get("suggestedInstructions") or []:
        text = str(item or "").strip()
        if not text or text in seen:
            continue
        if len(text) > _MAX_INSTRUCTION_CHARS:
            text = text[:_MAX_INSTRUCTION_CHARS]
        seen.add(text)
        lines.append(text)
        if len(lines) >= _MAX_INSTRUCTIONS:
            break
    return {
        "explanation": explanation,
        "suggestedInstructions": lines,
        "model": model,
    }


def suggest_ai_instructions(
    *,
    user_request: str,
    current_instructions: list[str] | None,
    comparison: dict,
    model: str | None = None,
    reasoning_effort: str | None = None,
    base_url: str | None = None,
    timeout_seconds: float | None = None,
    api_key: str | None = None,
    http_post: Callable[[str, dict[str, str], dict, float], dict] | None = None,
) -> dict:
    """Call OpenAI to suggest mapping.aiInstructions lines from a comparison package."""
    resolved_model = (model or os.environ.get("AI_VALIDATION_MODEL") or "gpt-5.6-luna").strip()
    resolved_effort = (
        reasoning_effort or os.environ.get("AI_VALIDATION_REASONING_EFFORT") or "medium"
    ).strip() or "medium"
    resolved_base = (base_url or os.environ.get("OPENAI_BASE_URL") or "https://api.openai.com/v1").rstrip("/")
    timeout = float(timeout_seconds if timeout_seconds is not None else 120)
    key = api_key or os.environ.get("OPENAI_API_KEY")
    if not nonempty(key):
        raise AIProviderError("OPENAI_API_KEY is not configured")

    payload = build_suggest_prompt_payload(
        user_request=user_request,
        current_instructions=current_instructions,
        comparison=comparison,
        model=resolved_model,
        reasoning_effort=resolved_effort,
    )
    poster = http_post or _default_http_post
    response = poster(
        resolved_base + "/responses",
        {"Authorization": "Bearer " + key, "Content-Type": "application/json"},
        payload,
        timeout,
    )
    parsed = _parse_structured_response(response)
    return normalize_suggestion_result(parsed, model=resolved_model)
