"""Independent AI comparison by default; explicit legacy mode for regression.

Independent generation receives originals and a blank template. CODE is used
only for declared special-field ownership and the subsequent audited overlay.
The old anchored/reconciled path remains below the explicit mode dispatch.
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
    normalize_code_provenance_cells,
    prepare_template_with_code_identities,
    sync_code_owned_regions_from_code_result,
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


def _code_provenance_cells(source_hints: dict) -> list[dict[str, Any]]:
    return normalize_code_provenance_cells(source_hints.get("codeProvenanceCells"))


# Fixed-layout engines / AI profile ids whose employee block starts below metadata.
_PROFILE_DATA_START_ROW = {
    "india-payroll": 10,
    "india_payroll_calc": 10,
}


def _data_start_row(source_hints: dict, profile_id: str, *, sheet_name: str | None = None) -> int | None:
    """Optional first employee row for CODE identity anchoring / prefilled names.

    Prefer Office-merged ``targetL.dataStartRow`` from the convert mapping; fall
    back to engineId / known profile defaults / India-L sheet convention.
    ``None`` means header_row + 1.
    """
    target = source_hints.get("targetL")
    raw = None
    if isinstance(target, dict) and target.get("dataStartRow") is not None:
        raw = target.get("dataStartRow")
    elif source_hints.get("dataStartRow") is not None:
        raw = source_hints.get("dataStartRow")
    else:
        engine = str(source_hints.get("engineId") or source_hints.get("engine_id") or "").strip()
        if engine in _PROFILE_DATA_START_ROW:
            raw = _PROFILE_DATA_START_ROW[engine]
        elif profile_id in _PROFILE_DATA_START_ROW:
            raw = _PROFILE_DATA_START_ROW[profile_id]
        elif str(sheet_name or "").strip().casefold() == "india-l":
            raw = 10
    if raw is None:
        return None
    try:
        start = int(raw)
    except (TypeError, ValueError) as exc:
        raise ValueError("AI targetL.dataStartRow must be an integer") from exc
    require(start >= 1, "AI targetL.dataStartRow must be >= 1")
    return start


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
    mode = source_hints.get("comparisonMode", "independent")
    require(mode in {"independent", "legacy"}, "Unknown AI comparisonMode")
    if mode == "independent":
        from .independent_runner import run_independent
        return run_independent(
            documents=documents, resolved=resolved, template_path=template_path, output_path=output_path,
            run_id=str(run_id), period=str(period), currency=str(currency).upper(), profile_id=profile_id,
            provider=provider, source_hints=source_hints, code_result_path=code_result_path,
        )
    work_template = Path(template_path).resolve()
    # Peek sheet name so India-L can default dataStartRow=10 even when Office
    # sends a numeric DB profile id instead of india-payroll / engineId.
    peek_manifest = inspect_last_l_sheet(work_template)
    data_start_row = _data_start_row(
        source_hints, str(profile_id), sheet_name=str(peek_manifest.get("sheetName") or ""),
    )
    code_anchor = None
    if code_result_path is not None:
        prepared = Path(output_path).resolve().parent / ("ai-template-with-code-names-" + str(run_id) + ".xlsx")
        if prepared.exists():
            prepared.unlink()
        print(
            "[ai] step1: anchoring CODE -L person names onto template"
            + (f" (dataStartRow={data_start_row})" if data_start_row is not None else "")
            + " …",
            flush=True,
        )
        code_anchor = prepare_template_with_code_identities(
            template_path, code_result_path, prepared,
            data_start_row=data_start_row,
        )
        work_template = prepared
        print(
            f"[ai] step1 done employees={code_anchor.get('employeeCount')} "
            f"nameCells={code_anchor.get('writeCount')} "
            f"reused={code_anchor.get('reuseCount')} "
            f"clearedSample={code_anchor.get('clearedSampleCount')}",
            flush=True,
        )
    print("[ai] inspecting current last -L template …", flush=True)
    manifest = inspect_last_l_sheet(work_template)
    if data_start_row is not None:
        manifest["dataStartRow"] = data_start_row
    elif isinstance(code_anchor, dict) and code_anchor.get("dataStartRow") is not None:
        manifest["dataStartRow"] = code_anchor["dataStartRow"]
    print(
        f"[ai] template sheet={manifest.get('sheetName')} "
        f"columns={len(manifest.get('columnContexts') or [])}"
        + (
            f" dataStartRow={manifest.get('dataStartRow')}"
            if manifest.get("dataStartRow") is not None else ""
        ),
        flush=True,
    )
    column_mappings = _column_mappings(source_hints)
    provenance_cells = _code_provenance_cells(source_hints)
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
        code_provenance_cells=provenance_cells,
    )
    print(
        f"[ai] plan ready writes={len(plan.get('writes') or [])} "
        f"issues={len(plan.get('issues') or [])}; writing workbook …",
        flush=True,
    )
    artifact = write_dynamic_ai_template_copy(
        work_template, output_path, plan, documents,
        column_mappings=column_mappings,
        code_provenance_cells=provenance_cells,
    )
    final_plan = artifact.get("plan") or plan
    print(f"[ai] workbook written sheet={artifact.get('sheetName')}", flush=True)
    provenance_copy = {
        "copyCount": 0, "skippedCount": 0,
        "nonLCopyCount": 0, "formulaCopyCount": 0, "provenanceCopyCount": 0,
    }
    if code_result_path is not None:
        print(
            "[ai] step3: syncing CODE-owned sheets/formulas"
            + (f" + {len(provenance_cells)} provenance cell(s)" if provenance_cells else "")
            + " …",
            flush=True,
        )
        provenance_copy = sync_code_owned_regions_from_code_result(
            code_result_path, output_path,
            provenance_cells=provenance_cells,
            data_start_row=manifest.get("dataStartRow") or data_start_row,
        )
        print(
            f"[ai] step3 done copied={provenance_copy.get('copyCount')} "
            f"nonL={provenance_copy.get('nonLCopyCount')} "
            f"metadata={provenance_copy.get('metadataCopyCount')} "
            f"formulas={provenance_copy.get('formulaCopyCount')} "
            f"provenance={provenance_copy.get('provenanceCopyCount')} "
            f"skipped={provenance_copy.get('skippedCount')}",
            flush=True,
        )
        artifact["outputSha256"] = hashlib.sha256(Path(output_path).read_bytes()).hexdigest()
    return {
        "runId": str(run_id),
        "profileId": profile_id,
        "provider": provider.provider_id,
        "model": final_plan["model"],
        "schemaId": "current-template-last-L",
        "schemaVersion": manifest["templateSha256"][:16],
        "period": str(period),
        "currency": str(currency).upper(),
        "sheetName": artifact["sheetName"],
        "templateSha256": artifact["templateSha256"],
        "outputSha256": artifact["outputSha256"],
        "writeCount": artifact["writeCount"],
        "recordCount": None,
        "readiness": "review_required" if final_plan["issues"] else "ready_for_review",
        "issueCount": len(final_plan["issues"]),
        "planVersion": final_plan["planVersion"],
        "columnMappingHintCount": len(column_mappings),
        "codeProvenanceCellCount": len(provenance_cells),
        "codeProvenanceCopyCount": provenance_copy.get("provenanceCopyCount") or 0,
        "codeOwnedCopyCount": provenance_copy.get("copyCount") or 0,
        "codeOwnedMetadataCopyCount": provenance_copy.get("metadataCopyCount") or 0,
        "codeOwnedNonLCopyCount": provenance_copy.get("nonLCopyCount") or 0,
        "codeOwnedFormulaCopyCount": provenance_copy.get("formulaCopyCount") or 0,
        "issues": final_plan["issues"],
        "automaticPassEnabled": False,
        "formalResult": False,
        "codeIdentityAnchored": bool(code_anchor),
        "codeIdentityEmployeeCount": (code_anchor or {}).get("employeeCount") or 0,
    }
