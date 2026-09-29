"""End-to-end AI comparison workbook orchestration.

This module is intentionally separate from the deterministic conversion runner.
It receives original supplier documents plus the current master template and
creates a new AI-only comparison workbook.

Optionally it may receive the CODE conversion workbook solely to copy person-name
identity cells onto the -L sheet before the model runs.  Amounts, fees and other
CODE values are never copied into the AI workbook.
"""
from __future__ import annotations

import hashlib
import json
import mimetypes
from pathlib import Path
from typing import Any

from bill_validation.contracts import nonempty, require

from .provider import AIProvider, OpenAIResponsesProvider
from .template_fill import (
    inspect_last_l_sheet,
    prepare_template_with_code_identities,
    write_dynamic_ai_template_copy,
)


_ALLOWED_SUFFIXES = {".xlsx", ".xlsm", ".pdf", ".csv"}


def list_ai_validation_profiles() -> list[dict[str, str]]:
    return [
        {"profileId": "taiwan-coral-sea", "engineId": "tw_payroll_calc", "status": "enabled", "mode": "template_driven"},
        {"profileId": "uk-payroll", "engineId": "uk_payroll_calc", "status": "enabled", "mode": "template_driven"},
        {"profileId": "uae-payroll", "engineId": "uae_payroll_calc", "status": "enabled", "mode": "template_driven"},
        {"profileId": "pakistan-payroll", "engineId": "pakistan_payroll_calc", "status": "enabled", "mode": "template_driven"},
        {"profileId": "india-payroll", "engineId": "india_payroll_calc", "status": "enabled", "mode": "template_driven"},
        {"profileId": "cyprus-payroll", "engineId": "cyprus_payroll_calc", "status": "enabled", "mode": "template_driven"},
    ]


def _documents(paths: list[Path], run_id: str) -> tuple[list[dict], dict[str, Path]]:
    require(bool(paths), "At least one original supplier document is required")
    documents: list[dict] = []
    resolved: dict[str, Path] = {}
    for index, raw_path in enumerate(paths, start=1):
        path = Path(raw_path).resolve()
        require(path.is_file(), "Original supplier document is missing")
        require(path.suffix.lower() in _ALLOWED_SUFFIXES, "Unsupported original supplier document type")
        file_id = "source-" + str(index)
        content = path.read_bytes()
        media_type = mimetypes.guess_type(path.name)[0]
        if path.suffix.lower() in {".xlsx", ".xlsm"}:
            media_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        elif path.suffix.lower() == ".csv":
            media_type = "text/csv"
        require(nonempty(media_type), "Original document media type is unknown")
        documents.append({
            "fileId": file_id,
            "sha256": hashlib.sha256(content).hexdigest(),
            "mediaType": media_type,
            "sourceRef": "ai-input://" + run_id + "/" + file_id,
        })
        resolved[file_id] = path
    return documents, resolved


def _instructions(source_hints: dict) -> list[str]:
    instructions = [str(item) for item in source_hints.get("instructions") or [] if nonempty(item)]
    # Only durable business rules belong here. Coordinates deliberately do not
    # enter the dynamic template-driven path.
    if source_hints.get("knownOutsideComparison"):
        instructions.append("Values supplied outside the bill and excluded from AI filling: "
                            + json.dumps(source_hints["knownOutsideComparison"], ensure_ascii=False,
                                         separators=(",", ":")))
    return instructions


def _column_mappings(source_hints: dict) -> dict[str, str]:
    """Read optional Office supplier -> current-template label hints."""
    raw = source_hints.get("columnRename")
    if raw is None:
        return {}
    require(isinstance(raw, dict), "AI columnRename hint must be an object")
    require(len(raw) <= 1000, "AI columnRename hint exceeds the entry limit")
    result: dict[str, str] = {}
    for source, target in raw.items():
        require(nonempty(source) and nonempty(target), "AI columnRename entries must be non-empty strings")
        result[str(source).strip()] = str(target).strip()
    return result


def run_ai_comparison_workbook(*, original_paths: list[str | Path], template_path: str | Path,
                               output_path: str | Path, run_id: str, period: str, currency: str,
                               profile_id: str = "taiwan-coral-sea", provider: AIProvider | None = None,
                               provider_options: dict[str, Any] | None = None,
                               source_hints: dict | None = None,
                               code_result_path: str | Path | None = None) -> dict:
    """Produce an AI-filled copy of the current template's last ``-L`` sheet.

    When ``code_result_path`` is provided, person names from the CODE result's
    last ``-L`` sheet are written into a prepared template first.  The model then
    only fills remaining blanks from the original supplier bills.

    The returned metadata deliberately has ``automaticPassEnabled=false``.  A
    caller may display or persist this artifact, but must not treat it as the
    formal conversion result.
    """
    require(nonempty(run_id), "AI validation run id is required")
    require(nonempty(period), "AI validation period is required")
    require(len(str(currency or "")) == 3 and str(currency).isalpha(), "AI validation currency is invalid")
    require(source_hints is None or isinstance(source_hints, dict), "AI source hints must be a JSON object")
    source_hints = source_hints or {}
    documents, resolved = _documents([Path(item) for item in original_paths], str(run_id))
    print(
        f"[ai] run={run_id} files={len(documents)} period={period} currency={currency} "
        f"profile={profile_id}",
        flush=True,
    )
    if provider is None:
        options = dict(provider_options or {})
        provider = OpenAIResponsesProvider(
            document_resolver=lambda document: resolved[document["fileId"]],
            **options,
        )
    work_template = Path(template_path).resolve()
    code_anchor = None
    if code_result_path is not None:
        prepared = Path(output_path).resolve().parent / ("ai-template-with-code-names-" + str(run_id) + ".xlsx")
        if prepared.exists():
            prepared.unlink()
        print("[ai] step1: anchoring CODE -L person names onto template …", flush=True)
        code_anchor = prepare_template_with_code_identities(
            template_path, code_result_path, prepared,
        )
        work_template = prepared
        print(
            f"[ai] step1 done employees={code_anchor.get('employeeCount')} "
            f"nameCells={code_anchor.get('writeCount')}",
            flush=True,
        )
    print("[ai] inspecting current last -L template …", flush=True)
    manifest = inspect_last_l_sheet(work_template)
    print(
        f"[ai] template sheet={manifest.get('sheetName')} "
        f"columns={len(manifest.get('columnContexts') or [])}",
        flush=True,
    )
    column_mappings = _column_mappings(source_hints)
    require(hasattr(provider, "plan_dynamic_template_fill"),
            "AI provider cannot perform dynamic template filling")
    print("[ai] step2: asking model for fill plan (this is the long wait) …", flush=True)
    plan = provider.plan_dynamic_template_fill(
        documents=documents,
        template_manifest=manifest,
        run_id=str(run_id),
        period=str(period),
        currency=str(currency).upper(),
        instructions=_instructions(source_hints),
        column_mappings=column_mappings,
        code_anchored_employees=(code_anchor or {}).get("employees"),
    )
    print(
        f"[ai] plan ready writes={len(plan.get('writes') or [])} "
        f"issues={len(plan.get('issues') or [])}; writing workbook …",
        flush=True,
    )
    artifact = write_dynamic_ai_template_copy(
        work_template, output_path, plan, documents, column_mappings=column_mappings,
    )
    print(f"[ai] workbook written sheet={artifact.get('sheetName')}", flush=True)
    return {
        "runId": str(run_id),
        "profileId": profile_id,
        "provider": provider.provider_id,
        "model": plan["model"],
        "schemaId": "current-template-last-L",
        "schemaVersion": manifest["templateSha256"][:16],
        "period": str(period),
        "currency": str(currency).upper(),
        "sheetName": artifact["sheetName"],
        "templateSha256": artifact["templateSha256"],
        "outputSha256": artifact["outputSha256"],
        "writeCount": artifact["writeCount"],
        "recordCount": None,
        "readiness": "review_required" if plan["issues"] else "ready_for_review",
        "issueCount": len(plan["issues"]),
        "planVersion": plan["planVersion"],
        "columnMappingHintCount": len(column_mappings),
        "issues": plan["issues"],
        "automaticPassEnabled": False,
        "formalResult": False,
        "codeIdentityAnchored": bool(code_anchor),
        "codeIdentityEmployeeCount": (code_anchor or {}).get("employeeCount") or 0,
    }
