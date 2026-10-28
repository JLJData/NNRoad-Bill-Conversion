"""Orchestration for the independent comparison experiment (never formal output)."""
from __future__ import annotations

import copy
import hashlib
import json
import os
import re
import tempfile
from pathlib import Path

from bill_validation.contracts import require
from .independent import (
    AUDIT_SHEET, PROVENANCE_SHEET, VERTICAL_LAYOUT, code_owned_columns, code_sheet_identity_hint,
    copy_independent_special_cells,
    embed_independent_audit, keep_ai_comparison_sheets, plan_sha256, prepare_independent_template,
    validate_independent_plan, validate_independent_template, write_independent_template_copy,
)
from .template_fill import list_l_sheets_in_workbook, normalize_code_provenance_cells, _normalized_semantic


def _multi_l_fill_mode(manifests: list, sheet_hints: dict, engine_id: str) -> str:
    """same_roster=同人多表（China-L / China-L (2)）；one_person_per_sheet=一人一表（UK-L (2)）。"""
    names = [str(item.get("sheetName") or "") for item in manifests]
    if str(engine_id or "").strip() == "china_hrone_payment_notice":
        return "same_roster"
    if names and all(n.strip().casefold().startswith("china-l") for n in names):
        return "same_roster"
    if any(str(item.get("layout") or "").strip() == VERTICAL_LAYOUT for item in manifests):
        return "one_person_per_sheet"
    keys = ["/".join(sorted({_normalized_semantic(person) for person in (sheet_hints.get(n) or "").split(" / ")
                             if _normalized_semantic(person)})) for n in names]
    keys = [k for k in keys if k]
    if len(keys) >= 2 and len(set(keys)) == 1:
        return "same_roster"
    if len(set(keys)) > 1:
        return "one_person_per_sheet"
    return "same_roster"


def _source_sheet_candidates(sheet_name: str, source_hints: dict, engine_id: str) -> list[str]:
    """Use explicit per-table bindings; a second target sheet does not imply another person."""
    copies = source_hints.get("lSheetCopies")
    if copies is None and engine_id:
        from convert_mapping import resolve_convert_mapping
        copies = resolve_convert_mapping(engine_id, {}).get("lSheetCopies")
    for item in copies or []:
        target, source = item.get("target") or {}, item.get("source") or {}
        if sheet_name in [target.get("sheet"), *(target.get("candidates") or [])]:
            return list(dict.fromkeys(n for n in [source.get("sheet"), *(source.get("candidates") or [])] if n))
    return {"china-l": ["S-Payslip"], "china-l (2)": ["S-Payroll Report"]}.get(sheet_name.casefold(), [])


def _multi_l_sheet_instructions(sheet_name: str, manifests: list, sheet_hints: dict, mode: str) -> list[str]:
    names = [str(item.get("sheetName") or "") for item in manifests]
    extra = [
        f"Fill only worksheet {sheet_name}. This CODE result has {len(manifests)} -L sheets: {', '.join(names)}.",
    ]
    hint = (sheet_hints.get(sheet_name) or "").strip()
    if mode == "one_person_per_sheet":
        extra.append(
            "Each -L sheet is a different person (or a different vertical employee pack). "
            "Do not copy another -L sheet's identity or amounts onto this sheet."
        )
        if hint:
            extra.append(f"This -L sheet corresponds to: {hint}")
        return extra
    extra.append(
        "These -L sheets share the same employee roster. They are different source tables "
        "for the same person(s) (e.g. payslip vs payroll report), not different people."
    )
    extra.append(
        "Fill this sheet from the original bill table that matches its columns; "
        "do not invent a second employee just because there is a second -L sheet."
    )
    if sheet_name.strip().casefold() == "china-l":
        extra.append("China-L is filled from S-Payslip (or the equivalent payslip sheet).")
    elif sheet_name.strip().casefold().startswith("china-l"):
        extra.append("China-L (2) is filled from S-Payroll Report (or the equivalent payroll report sheet).")
    if hint:
        extra.append(f"Employee identity on this sheet: {hint}")
    return extra


def run_independent(*, documents, resolved, template_path, output_path, run_id, period,
                    currency, profile_id, provider, source_hints, code_result_path):
    output = Path(output_path).resolve()
    require(not output.exists(), "Independent AI output must be a new file")
    # Durable audit survives the API's temporary upload-directory cleanup, including
    # failed attempts. Full plans are embedded in successful XLSX files as well.
    audit_dir = Path(os.environ.get("AI_VALIDATION_AUDIT_DIR") or
                     Path(__file__).resolve().parent.parent / "output" / "ai_audit")
    audit_dir.mkdir(parents=True, exist_ok=True)
    safe_run = re.sub(r"[^a-zA-Z0-9_.-]", "_", str(run_id))[:100]
    audit_file = audit_dir / (safe_run + "-" + hashlib.sha256(str(run_id).encode()).hexdigest()[:12] + ".json")
    require(not audit_file.exists(), "Independent audit runId already exists; use a new runId")
    audit = {"auditVersion": 1, "comparisonMode": "independent", "runId": run_id,
             "period": period, "currency": currency,
             "originalDocuments": [{k: v for k, v in d.items() if k != "sourceRef"} for d in documents],
             "attempts": [], "status": "preparing", "codeSpecialCopies": []}

    def persist():
        temporary = audit_file.with_suffix(".tmp")
        temporary.write_text(json.dumps(audit, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
        temporary.replace(audit_file)

    persist()
    output.parent.mkdir(parents=True, exist_ok=True)
    workspace = tempfile.TemporaryDirectory(prefix="ai-template-", dir=output.parent)
    prepared = Path(workspace.name) / "blank.xlsx"
    try:
        print("[ai] independent: preparing template without CODE identities or amounts", flush=True)
        code_l_names = list_l_sheets_in_workbook(code_result_path) if code_result_path else []
        multi_l = len(code_l_names) > 1
        if multi_l:
            print(f"[ai] CODE has {len(code_l_names)} -L sheets: {code_l_names}", flush=True)
        manifest = prepare_independent_template(
            template_path, prepared, layout=source_hints.get("targetL"),
            l_sheet_names=code_l_names if multi_l else None,
            engine_id=str(source_hints.get("engineId") or source_hints.get("engine_id") or "").strip() or None,
        )
        manifests = list(manifest.get("lSheetManifests") or [manifest])
        provenance = normalize_code_provenance_cells(source_hints.get("codeProvenanceCells"))
        require(not provenance or code_result_path is not None, "Special fields require the CODE result")
        for item in manifests:
            item["codeOwnedColumns"] = code_owned_columns(code_result_path, item, provenance)
            validate_independent_template(item)
        audit["template"] = copy.deepcopy(manifest)
        audit["lSheetNames"] = [item["sheetName"] for item in manifests]
        audit["originalTemplateSha256"] = hashlib.sha256(Path(template_path).read_bytes()).hexdigest()
        audit["status"] = "generating"
        persist()

        require(hasattr(provider, "plan_independent_template_fill"), "Provider does not support independent comparison")
        from .runner import _instructions
        base_instructions = _instructions(source_hints or {})
        plans = []
        artifacts = []
        sheet_hints = {
            item["sheetName"]: code_sheet_identity_hint(code_result_path, item["sheetName"])
            for item in manifests
        } if multi_l else {}
        engine_id = str(source_hints.get("engineId") or source_hints.get("engine_id") or "").strip()
        multi_mode = _multi_l_fill_mode(manifests, sheet_hints, engine_id) if multi_l else ""
        audit["multiLFillMode"] = multi_mode
        for index, sheet_manifest in enumerate(manifests):
            sheet_name = sheet_manifest["sheetName"]
            extra = []
            if multi_l:
                if multi_mode == "one_person_per_sheet":
                    sheet_manifest["codeSheetIdentityHint"] = sheet_hints.get(sheet_name) or ""
                    sheet_manifest["otherSheetIdentityHints"] = [
                        hint for name, hint in sheet_hints.items() if name != sheet_name and hint
                    ]
                else:
                    sheet_manifest["sourceSheetCandidates"] = _source_sheet_candidates(
                        sheet_name, source_hints, engine_id,
                    )
                extra.extend(_multi_l_sheet_instructions(
                    sheet_name, manifests, sheet_hints, multi_mode,
                ))
                if sheet_manifest.get("sourceSheetCandidates"):
                    extra.append("Every write on this sheet must cite its bound original source table: "
                                 + ", ".join(sheet_manifest["sourceSheetCandidates"]))
            print(f"[ai] independent: planning sheet {sheet_name} ({index + 1}/{len(manifests)})", flush=True)

            def validate(plan, _manifest=sheet_manifest):
                validate_independent_plan(plan, _manifest, documents, source_paths=resolved)

            def attempt(entry):
                audit["attempts"].append(copy.deepcopy(entry))
                persist()

            plan = provider.plan_independent_template_fill(
                documents=documents, template_manifest=sheet_manifest, template_path=prepared,
                run_id=run_id, period=period, currency=currency, validate=validate, on_attempt=attempt,
                instructions=base_instructions + extra,
            )
            validate(plan)
            artifact = write_independent_template_copy(
                prepared, output, plan, documents,
                source_paths=resolved, manifest=sheet_manifest,
                into_existing=index > 0,
            )
            plans.append(plan)
            artifacts.append(artifact)
        plan = plans[-1]
        artifact = artifacts[-1]
        audit["finalPlan"] = copy.deepcopy(plan)
        audit["finalPlans"] = copy.deepcopy(plans)
        audit["finalPlanSha256"] = plan_sha256({"plans": plans}) if multi_l else plan_sha256(plan)
        audit["rawAiWorkbookSha256"] = artifact["outputSha256"]
        special_records, special_issues = [], []
        l_names = {item["sheetName"] for item in manifests}
        for index, sheet_manifest in enumerate(manifests):
            sheet_name = sheet_manifest["sheetName"]
            scoped = [
                item for item in provenance
                if item.get("sheet") == sheet_name or (index == 0 and item.get("sheet") not in l_names)
            ]
            special = copy_independent_special_cells(code_result_path, output, scoped, sheet_manifest)
            special_records.extend(special["records"])
            special_issues.extend(special["issues"])
        audit["codeSpecialCopies"] = special_records
        audit["postprocessingIssues"] = special_issues
        audit["status"] = "completed"
        persist()
        embed_independent_audit(output, audit)
        keep_names = [item["sheetName"] for item in manifests]
        keep_ai_comparison_sheets(output, code_result_path, keep_l_sheets=keep_names if multi_l else None)
        issues = [issue for item in plans for issue in item.get("issues") or []] + special_issues
        write_count = sum(item.get("writeCount") or 0 for item in artifacts)
        return {
            "runId": run_id, "profileId": profile_id, "provider": provider.provider_id, "model": plan["model"],
            "schemaId": "independent-template-last-L" if not multi_l else "independent-template-multi-L",
            "schemaVersion": manifest["templateSha256"][:16],
            "comparisonMode": "independent", "planVersion": 4, "period": period, "currency": currency,
            "sheetName": keep_names[0], "lSheetNames": keep_names,
            "multiLFillMode": multi_mode,
            "templateSha256": manifest["templateSha256"],
            "outputSha256": hashlib.sha256(output.read_bytes()).hexdigest(), "writeCount": write_count,
            "firstPlanSha256": audit["attempts"][0]["planSha256"] if audit["attempts"] else None,
            "finalPlanSha256": audit["finalPlanSha256"], "attemptCount": len(audit["attempts"]),
            "auditSheet": AUDIT_SHEET, "provenanceSheet": PROVENANCE_SHEET,
            "codeIdentityAnchored": False, "columnMappingHintCount": 0,
            "codeProvenanceCellCount": len(provenance), "codeProvenanceCopyCount": len(special_records),
            "independentAccuracyIncludesCodeSpecials": False,
            "readiness": "review_required" if issues else "ready_for_review",
            "issueCount": len(issues), "issues": issues,
            "automaticPassEnabled": False, "formalResult": False,
        }
    except Exception as exc:
        audit["status"] = "failed"
        audit["error"] = str(exc)
        persist()
        # A failed overlay/audit must not leave a seemingly successful workbook.
        output.unlink(missing_ok=True)
        raise
    finally:
        workspace.cleanup()
