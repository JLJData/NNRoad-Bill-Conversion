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
    AUDIT_SHEET, PROVENANCE_SHEET, code_owned_columns, copy_independent_special_cells,
    embed_independent_audit, plan_sha256, prepare_independent_template,
    validate_independent_plan, validate_independent_template, write_independent_template_copy,
)
from .template_fill import normalize_code_provenance_cells


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
        manifest = prepare_independent_template(template_path, prepared, layout=source_hints.get("targetL"))
        provenance = normalize_code_provenance_cells(source_hints.get("codeProvenanceCells"))
        require(not provenance or code_result_path is not None, "Special fields require the CODE result")
        manifest["codeOwnedColumns"] = code_owned_columns(code_result_path, manifest, provenance)
        validate_independent_template(manifest)
        audit["template"] = copy.deepcopy(manifest)
        audit["originalTemplateSha256"] = hashlib.sha256(Path(template_path).read_bytes()).hexdigest()
        audit["status"] = "generating"
        persist()

        def validate(plan):
            validate_independent_plan(plan, manifest, documents, source_paths=resolved)

        def attempt(entry):
            audit["attempts"].append(copy.deepcopy(entry))
            persist()

        require(hasattr(provider, "plan_independent_template_fill"), "Provider does not support independent comparison")
        from .runner import _instructions
        plan = provider.plan_independent_template_fill(
            documents=documents, template_manifest=manifest, template_path=prepared,
            run_id=run_id, period=period, currency=currency, validate=validate, on_attempt=attempt,
            instructions=_instructions(source_hints or {}),
        )
        validate(plan)
        artifact = write_independent_template_copy(prepared, output, plan, documents,
                                                   source_paths=resolved, manifest=manifest)
        audit["finalPlan"] = copy.deepcopy(plan)
        audit["finalPlanSha256"] = plan_sha256(plan)
        audit["rawAiWorkbookSha256"] = artifact["outputSha256"]
        special = copy_independent_special_cells(code_result_path, output, provenance, manifest)
        audit["codeSpecialCopies"] = special["records"]
        audit["postprocessingIssues"] = special["issues"]
        audit["status"] = "completed"
        persist()
        embed_independent_audit(output, audit)
        issues = list(plan["issues"]) + special["issues"]
        return {
            "runId": run_id, "profileId": profile_id, "provider": provider.provider_id, "model": plan["model"],
            "schemaId": "independent-template-last-L", "schemaVersion": manifest["templateSha256"][:16],
            "comparisonMode": "independent", "planVersion": 4, "period": period, "currency": currency,
            "sheetName": artifact["sheetName"], "templateSha256": manifest["templateSha256"],
            "outputSha256": hashlib.sha256(output.read_bytes()).hexdigest(), "writeCount": artifact["writeCount"],
            "firstPlanSha256": audit["attempts"][0]["planSha256"] if audit["attempts"] else None,
            "finalPlanSha256": audit["finalPlanSha256"], "attemptCount": len(audit["attempts"]),
            "auditSheet": AUDIT_SHEET, "provenanceSheet": PROVENANCE_SHEET,
            "codeIdentityAnchored": False, "columnMappingHintCount": 0,
            "codeProvenanceCellCount": len(provenance), "codeProvenanceCopyCount": len(special["records"]),
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
