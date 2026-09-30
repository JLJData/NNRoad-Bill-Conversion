"""Inspect a changing workbook and safely fill its last ``-L`` worksheet.

The model may propose target cells, but this module owns all workbook writes.  It
copies the template, verifies its hash, rejects formulas/merged placeholders,
and resolves every written value from the validated AI collection instead of
accepting arbitrary values from a model-generated plan.
"""
from __future__ import annotations

import copy
import hashlib
import re
import shutil
from difflib import SequenceMatcher
from datetime import date
from decimal import Decimal, InvalidOperation
from io import BytesIO
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.utils.cell import coordinate_to_tuple, get_column_letter

from bill_validation.contracts import ValidationError, decimal_text, nonempty, require

from .contracts import validate_ai_collection


def _sha256(raw: bytes) -> str:
    return hashlib.sha256(raw).hexdigest()


def _last_l_sheet(sheetnames: list[str]) -> str:
    matches = [name for name in sheetnames if name.strip().casefold().endswith("-l")]
    require(bool(matches), "Template has no worksheet ending in -L")
    return matches[-1]


def _literal_label(value: Any, data_type: str | None = None) -> str:
    if value is None or data_type == "f" or isinstance(value, (int, float, date)):
        return ""
    return " ".join(str(value).split()).strip()


def _column_contexts(ws, header_row: int | None = None) -> tuple[int, list[dict]]:
    """Describe columns semantically without assuming field names or coordinates."""
    merged_labels: dict[tuple[int, int], tuple[str, str]] = {}
    for merged in ws.merged_cells.ranges:
        anchor = ws.cell(merged.min_row, merged.min_col)
        label = _literal_label(anchor.value, anchor.data_type)
        if not label:
            continue
        for row in range(merged.min_row, merged.max_row + 1):
            for column in range(merged.min_col, merged.max_col + 1):
                merged_labels[(row, column)] = (anchor.coordinate, label)

    row_scores: list[tuple[int, int]] = []
    scan_rows = min(max(1, ws.max_row), 200)
    for row in range(1, scan_rows + 1):
        labels = set()
        for column in range(1, ws.max_column + 1):
            cell = ws.cell(row, column)
            label = _literal_label(cell.value, cell.data_type)
            if label:
                labels.add((cell.coordinate, label))
            elif (row, column) in merged_labels:
                labels.add(merged_labels[(row, column)])
        row_scores.append((len(labels), row))
    if header_row is None:
        header_row = max(row_scores, key=lambda item: (item[0], item[1]))[1] if row_scores else 1

    contexts = []
    for column in range(1, ws.max_column + 1):
        labels = []
        seen = set()
        for row in range(1, header_row + 1):
            cell = ws.cell(row, column)
            label = _literal_label(cell.value, cell.data_type)
            coordinate = cell.coordinate
            if not label and (row, column) in merged_labels:
                coordinate, label = merged_labels[(row, column)]
            normalized = " ".join(label.casefold().split())
            if label and normalized not in seen:
                seen.add(normalized)
                labels.append({"cell": coordinate, "value": label})
        if not labels:
            continue
        path_parts = [item["value"] for item in labels]
        contexts.append({
            "column": column,
            "columnLetter": get_column_letter(column),
            "primaryLabel": labels[-1]["value"],
            # Parent -> child path so identical leaf labels under different
            # section headers stay distinguishable.
            "pathLabel": " / ".join(path_parts),
            "labels": labels,
        })
    return header_row, contexts


def _normalized_semantic(value: Any) -> str:
    text = " ".join(str(value or "").casefold().split())
    return re.sub(r"[^0-9a-z\u3400-\u9fff]+", "", text)


def _semantic_variants(value: Any) -> set[str]:
    """Keep the full label plus its Latin/numeric business-code portion.

    Bilingual template headers such as ``職災 GI`` must still match the exact
    supplier code ``GI``.  The joined Latin tokens preserve qualifiers such as
    ``2G``, ``EE``, ``ER`` and ``1T`` instead of matching any one token alone.
    """
    text = " ".join(str(value or "").casefold().split())
    result = {_normalized_semantic(text)}
    latin = "".join(re.findall(r"[0-9a-z]+", text))
    if latin:
        result.add(latin)
    return {item for item in result if item}


def _semantic_score(left: Any, right: Any) -> float:
    left_variants = _semantic_variants(left)
    right_variants = _semantic_variants(right)
    if not left_variants or not right_variants:
        return 0.0
    if left_variants & right_variants:
        return 1.0
    return max(SequenceMatcher(None, a, b).ratio()
               for a in left_variants for b in right_variants)


def _context_score(source_label: Any, context: dict) -> float:
    """Score a source label against a column using parent/child path + leaf."""
    primary = context.get("primaryLabel")
    path = context.get("pathLabel") or primary
    scores = [_semantic_score(source_label, primary), _semantic_score(source_label, path)]
    # Also compare slash/path-style source labels against the joined path.
    source_text = " ".join(str(source_label or "").replace("/", " ").split())
    if source_text:
        scores.append(_semantic_score(source_text, path))
        scores.append(_semantic_score(source_text, primary))
    return max(scores) if scores else 0.0


def _header_path_key(value: Any) -> str:
    """Normalize a parent/child header path for exact Excel-to-template matching."""
    return _normalized_semantic(str(value or "").replace("/", " "))


def _source_header_has_parent(label: Any) -> bool:
    text = " ".join(str(label or "").split())
    return " / " in text or (text.count("/") >= 1 and len(text.split("/")) >= 2)


def _exact_header_column(label: str, column_contexts: list[dict]) -> dict | None:
    """Lock a column when Excel source headers equal the template path.

    Parent+child pathLabel wins over leaf-only similarity.  A leaf match is used
    only when that primaryLabel is unique and the source label has no parent.
    """
    source_key = _header_path_key(label)
    if not source_key:
        return None
    path_hits = [
        item for item in column_contexts
        if _header_path_key(item.get("pathLabel") or item.get("primaryLabel")) == source_key
    ]
    if len(path_hits) == 1:
        return path_hits[0]
    if _source_header_has_parent(label):
        return None
    leaf_hits = [
        item for item in column_contexts
        if _header_path_key(item.get("primaryLabel")) == source_key
    ]
    return leaf_hits[0] if len(leaf_hits) == 1 else None


def _configured_target(source_label: str, column_mappings: dict[str, str] | None) -> str | None:
    if not column_mappings:
        return None
    source = _normalized_semantic(source_label)
    direct = [target for key, target in column_mappings.items()
              if _normalized_semantic(key) == source]
    if direct:
        return direct[0]
    bilingual_matches = {target for key, target in column_mappings.items()
                         if _semantic_score(key, source_label) == 1.0}
    if len(bilingual_matches) == 1:
        return next(iter(bilingual_matches))
    # Prefer full parent/child mapping keys before falling back to leaf-only.
    source_path = _normalized_semantic(str(source_label).replace("/", " "))
    path_matches = {target for key, target in column_mappings.items()
                    if _normalized_semantic(str(key).replace("/", " ")) == source_path}
    if len(path_matches) == 1:
        return next(iter(path_matches))
    source_child = _normalized_semantic(str(source_label).rsplit("/", 1)[-1])
    child_matches = {target for key, target in column_mappings.items()
                     if _normalized_semantic(str(key).rsplit("/", 1)[-1]) == source_child}
    return next(iter(child_matches)) if len(child_matches) == 1 else None


def _unambiguous_column(label: str, column_contexts: list[dict]) -> dict | None:
    scored = sorted(
        ((_context_score(label, item), item)
         for item in column_contexts if nonempty(item.get("primaryLabel"))),
        key=lambda pair: pair[0], reverse=True,
    )
    if not scored or scored[0][0] < 0.35:
        return None
    second_score = scored[1][0] if len(scored) > 1 else 0.0
    if scored[0][0] < second_score + 0.05:
        return None
    return scored[0][1]


def _is_identity_label(label: Any) -> bool:
    text = " ".join(str(label or "").casefold().split())
    if not text:
        return False
    if text in {"bu"} or "position" in text or "ft/pt" in text or text in {"ftpt"}:
        return True
    return (
        "cn name" in text or "en name" in text or text == "name" or text.endswith(" name")
        or "姓名" in text or "英文名" in text or "中文名" in text
        or "start date" in text or "end date" in text
        or ("employee" in text and "name" in text)
    )


def _is_person_name_label(label: Any) -> bool:
    """Person identity only — not company/client/BU/position/date fields."""
    text = " ".join(str(label or "").casefold().split())
    if not text:
        return False
    blocked = ("company", "client", "employer", "vendor", "bank", "file", "bu", "position")
    if any(token in text for token in blocked):
        return False
    if text in {"cn name", "en name", "name", "ee name", "employee name"}:
        return True
    if "cn name" in text or "en name" in text:
        return True
    if "姓名" in text or "中文名" in text or "英文名" in text:
        return True
    return ("employee" in text and "name" in text) or text.endswith(" name")


def _is_service_fee_label(label: Any) -> bool:
    """Service Fee / 服务费 — AI must leave these blank; CODE owns them."""
    text = " ".join(str(label or "").casefold().split())
    if not text:
        return False
    # Avoid matching unrelated "fee" labels such as Expense Fee / Operation Fee alone.
    if "service fee" in text or "服务费" in text:
        return True
    normalized = _normalized_semantic(text)
    return normalized in {"servicefee", "服务费"} or normalized.endswith("servicefee")


def _clean_cell_text(value: Any) -> str:
    if value is None or isinstance(value, bool):
        return ""
    if isinstance(value, (int, float, date)):
        return ""
    return " ".join(str(value).split()).strip()


def _parse_sheet_coordinate(location: Any) -> tuple[str, str] | None:
    text = str(location or "").strip()
    if "!" not in text:
        return None
    sheet_name, coordinate = text.rsplit("!", 1)
    sheet_name = sheet_name.strip().strip("'").replace("''", "'")
    coordinate = coordinate.strip().replace("$", "")
    if not sheet_name or not re.fullmatch(r"[A-Za-z]{1,3}[1-9][0-9]*", coordinate):
        return None
    return sheet_name, coordinate


def inspect_source_employee_layout(ws) -> dict[str, Any]:
    """Find person-name columns and which data rows look like real employees.

    Nameless numeric rows (common supplier subtotal/summary rows) are not employees.
    """
    header_row, contexts = _column_contexts(ws)
    name_contexts = [
        item for item in contexts
        if _is_person_name_label(item.get("primaryLabel")) or _is_person_name_label(item.get("pathLabel"))
    ]
    name_columns = [int(item["column"]) for item in name_contexts]
    # Prefer exact CN/EN name columns when present; avoid solitary vague "Name".
    precise = [
        item for item in name_contexts
        if _normalized_semantic(item.get("primaryLabel")) in {
            "cnname", "enname", "姓名", "中文名", "英文名", "employeename", "eename",
        }
        or "cn name" in " ".join(str(item.get("primaryLabel") or "").casefold().split())
        or "en name" in " ".join(str(item.get("primaryLabel") or "").casefold().split())
    ]
    if precise:
        name_contexts = precise
        name_columns = [int(item["column"]) for item in name_contexts]
    label_by_column = {
        int(item["column"]): str(item.get("primaryLabel") or item.get("pathLabel") or "")
        for item in name_contexts
    }
    rows: dict[int, dict[str, Any]] = {}
    max_row = min(max(1, ws.max_row or 1), header_row + 500)
    for row in range(header_row + 1, max_row + 1):
        names = []
        fields = []
        for column in name_columns:
            text = _clean_cell_text(ws.cell(row, column).value)
            if not text:
                continue
            names.append(text)
            fields.append({
                "column": column,
                "coordinate": get_column_letter(column) + str(row),
                "label": label_by_column.get(column) or "",
                "value": text,
            })
        rows[row] = {
            "hasName": bool(names),
            "names": names,
            "nameKey": " / ".join(_normalized_semantic(item) for item in names if _normalized_semantic(item)),
            "fields": fields,
        }
    return {
        "headerRow": header_row,
        "nameColumns": name_columns,
        "rows": rows,
    }


def list_named_source_employees(source_layouts: dict[str, dict[str, dict[str, Any]]]) -> list[dict[str, Any]]:
    """Flatten program-detected named employees across original XLSX sheets."""
    employees: list[dict[str, Any]] = []
    seen: set[str] = set()
    for file_id, sheets in source_layouts.items():
        if not isinstance(sheets, dict):
            continue
        for sheet_name, layout in sheets.items():
            if not isinstance(layout, dict) or not layout.get("nameColumns"):
                continue
            for source_row, row_info in sorted((layout.get("rows") or {}).items()):
                if not isinstance(row_info, dict) or not row_info.get("hasName"):
                    continue
                name_key = str(row_info.get("nameKey") or "")
                if not name_key or name_key in seen:
                    continue
                seen.add(name_key)
                employees.append({
                    "fileId": str(file_id),
                    "sheetName": str(sheet_name),
                    "sourceRow": int(source_row),
                    "names": list(row_info.get("names") or []),
                    "nameKey": name_key,
                    "displayName": " / ".join(row_info.get("names") or []) or name_key,
                    "fields": copy.deepcopy(row_info.get("fields") or []),
                })
    return employees


def _missing_employee_displays(plan: dict, source_layouts: dict[str, dict[str, dict[str, Any]]]) -> list[str]:
    covered_names: set[str] = set()
    for item in plan.get("writes") or []:
        if not isinstance(item, dict):
            continue
        source = item.get("source") if isinstance(item.get("source"), dict) else {}
        parsed = _parse_sheet_coordinate(source.get("location"))
        if not parsed:
            continue
        sheet_name, coordinate = parsed
        try:
            source_row, _column = coordinate_to_tuple(coordinate)
        except (TypeError, ValueError):
            continue
        layout = (source_layouts.get(str(source.get("fileId"))) or {}).get(sheet_name)
        if not layout:
            continue
        row_info = (layout.get("rows") or {}).get(source_row) or {}
        name_key = str(row_info.get("nameKey") or "")
        if name_key:
            covered_names.add(name_key)
    missing: list[str] = []
    for employee in list_named_source_employees(source_layouts):
        if employee["nameKey"] not in covered_names and employee["displayName"] not in missing:
            missing.append(employee["displayName"])
    return missing


def annotate_missing_source_employees(
    plan: dict,
    source_layouts: dict[str, dict[str, dict[str, Any]]],
) -> dict:
    """Replace MISSING_SOURCE_EMPLOYEES issues based on current writes vs source roster."""
    resolved = copy.deepcopy(plan)
    issues = [
        item for item in (resolved.get("issues") or [])
        if isinstance(item, dict) and item.get("code") != "MISSING_SOURCE_EMPLOYEES"
    ]
    missing = _missing_employee_displays(resolved, source_layouts)
    if missing:
        issues.append({
            "code": "MISSING_SOURCE_EMPLOYEES",
            "message": (
                "Source employees were not mapped to any kept -L row: "
                + ", ".join(missing[:20])
                + ("…" if len(missing) > 20 else "")
            ),
            "sourceLabel": "employee",
            "targetCell": None,
        })
    resolved["issues"] = issues
    return resolved


def seed_missing_employee_identity_writes(
    plan: dict,
    template_manifest: dict,
    source_layouts: dict[str, dict[str, dict[str, Any]]],
) -> dict:
    """Programmatically place CN/EN names for source employees the model omitted.

    Amount fields stay for the model; identity rows must not depend on the model
    noticing that the bill has multiple people.
    """
    resolved = copy.deepcopy(plan)
    writes = resolved.setdefault("writes", [])
    if not isinstance(writes, list):
        resolved["writes"] = []
        writes = resolved["writes"]
    issues = resolved.setdefault("issues", [])
    if not isinstance(issues, list):
        issues = []
        resolved["issues"] = issues

    contexts = [
        item for item in template_manifest.get("columnContexts") or []
        if isinstance(item, dict) and isinstance(item.get("column"), int)
    ]
    target_name_contexts = [
        item for item in contexts
        if _is_person_name_label(item.get("primaryLabel")) or _is_person_name_label(item.get("pathLabel"))
    ]
    if not target_name_contexts:
        return annotate_missing_source_employees(resolved, source_layouts)

    occupied_cells = {
        str(item.get("cell"))
        for item in template_manifest.get("nonemptyCells") or []
        if isinstance(item, dict) and nonempty(item.get("cell"))
    }
    formula_rows = {
        int(item["row"])
        for item in template_manifest.get("nonemptyCells") or []
        if isinstance(item, dict) and item.get("valueKind") == "formula" and isinstance(item.get("row"), int)
    }
    for item in writes:
        if isinstance(item, dict) and nonempty(item.get("targetCell")):
            occupied_cells.add(str(item["targetCell"]))

    header_row = int(template_manifest.get("headerRow") or 1)
    max_column = int(template_manifest.get("maxColumn") or 1)
    max_row = int(template_manifest.get("maxRow") or header_row) + 2000

    def _row_free(row: int) -> bool:
        if row in formula_rows:
            return False
        for ctx in target_name_contexts:
            cell = get_column_letter(int(ctx["column"])) + str(row)
            if cell in occupied_cells:
                return False
        return True

    covered_rows: dict[str, int] = {}
    for item in writes:
        if not isinstance(item, dict):
            continue
        source = item.get("source") if isinstance(item.get("source"), dict) else {}
        parsed = _parse_sheet_coordinate(source.get("location"))
        if not parsed:
            continue
        sheet_name, coordinate = parsed
        try:
            source_row, _column = coordinate_to_tuple(coordinate)
            target_row, _target_column = coordinate_to_tuple(str(item.get("targetCell")))
        except (TypeError, ValueError):
            continue
        layout = (source_layouts.get(str(source.get("fileId"))) or {}).get(sheet_name)
        if not layout:
            continue
        row_info = (layout.get("rows") or {}).get(source_row) or {}
        name_key = str(row_info.get("nameKey") or "")
        if name_key and name_key not in covered_rows:
            covered_rows[name_key] = target_row

    next_row = header_row + 1
    seeded = 0
    for employee in list_named_source_employees(source_layouts):
        if employee["nameKey"] in covered_rows:
            target_row = covered_rows[employee["nameKey"]]
        else:
            while next_row <= max_row and not _row_free(next_row):
                next_row += 1
            if next_row > max_row:
                issues.append({
                    "code": "EMPLOYEE_ROW_OVERFLOW",
                    "message": f"No free -L row left for employee {employee['displayName']}",
                    "sourceLabel": "employee",
                    "targetCell": None,
                })
                break
            target_row = next_row
            next_row += 1
        wrote_for_employee = False
        for field in employee.get("fields") or []:
            source_label = str(field.get("label") or "")
            value = str(field.get("value") or "")
            if not source_label or not value:
                continue
            context = _unambiguous_column(source_label, target_name_contexts)
            if context is None:
                scored = sorted(
                    ((_context_score(source_label, ctx), ctx) for ctx in target_name_contexts),
                    key=lambda pair: pair[0], reverse=True,
                )
                if scored and scored[0][0] >= 0.85:
                    context = scored[0][1]
            if context is None:
                continue
            target_cell = get_column_letter(int(context["column"])) + str(target_row)
            if target_cell in occupied_cells or int(context["column"]) > max_column:
                continue
            location = f"{employee['sheetName']}!{field['coordinate']}"
            writes.append({
                "targetCell": target_cell,
                "semanticLabel": context["primaryLabel"],
                "sourceLabel": source_label,
                "valueType": "text",
                "value": value,
                "confidence": 1.0,
                "source": {
                    "fileId": employee["fileId"],
                    "location": location,
                    "page": None,
                    "rawText": f"{source_label}: {value}",
                },
            })
            occupied_cells.add(target_cell)
            seeded += 1
            wrote_for_employee = True
        covered_rows[employee["nameKey"]] = target_row
        if wrote_for_employee:
            issues.append({
                "code": "EMPLOYEE_IDENTITY_SEEDED",
                "message": (
                    f"Seeded identity for {employee['displayName']} onto -L row {target_row} "
                    f"from {employee['sheetName']}!{employee['sourceRow']}"
                ),
                "sourceLabel": "employee",
                "targetCell": get_column_letter(int(target_name_contexts[0]["column"])) + str(target_row),
            })
    if seeded:
        resolved["issues"] = [
            item for item in issues if item.get("code") != "MISSING_SOURCE_EMPLOYEES"
        ]
    else:
        resolved["issues"] = issues
    return annotate_missing_source_employees(resolved, source_layouts)


def _name_token_set(names: Any) -> set[str]:
    tokens: set[str] = set()
    if not isinstance(names, (list, tuple)):
        return tokens
    for name in names:
        norm = _normalized_semantic(name)
        if norm:
            tokens.add(norm)
    return tokens


def _identity_data_start_row(template_manifest: dict | None, header_row: int) -> int:
    """First employee data row; skips period/metadata rows below the header."""
    raw = None if not isinstance(template_manifest, dict) else template_manifest.get("dataStartRow")
    if raw is None:
        return header_row + 1
    try:
        start = int(raw)
    except (TypeError, ValueError):
        return header_row + 1
    return max(start, header_row + 1)


def _prefilled_identity_rows(template_manifest: dict | None) -> dict[int, dict[str, Any]]:
    """Person names already present on -L data rows (e.g. CODE identity anchoring)."""
    if not isinstance(template_manifest, dict):
        return {}
    header_row = int(template_manifest.get("headerRow") or 1)
    data_start = _identity_data_start_row(template_manifest, header_row)
    name_columns: set[int] = set()
    for item in template_manifest.get("columnContexts") or []:
        if isinstance(item, dict) and isinstance(item.get("column"), int) and (
            _is_person_name_label(item.get("primaryLabel"))
            or _is_person_name_label(item.get("pathLabel"))
        ):
            name_columns.add(int(item["column"]))
    if not name_columns:
        return {}
    by_row: dict[int, list[str]] = {}
    for item in template_manifest.get("nonemptyCells") or []:
        if not isinstance(item, dict) or item.get("valueKind") == "formula":
            continue
        row = item.get("row")
        column = item.get("column")
        if not isinstance(row, int) or row < data_start:
            continue
        if column not in name_columns:
            continue
        text = _clean_cell_text(item.get("value"))
        if not text:
            continue
        by_row.setdefault(row, []).append(text)
    result: dict[int, dict[str, Any]] = {}
    for row, names in by_row.items():
        tokens = _name_token_set(names)
        if not tokens:
            continue
        result[row] = {
            "names": names,
            "tokens": tokens,
            "nameKey": " / ".join(
                _normalized_semantic(item) for item in names if _normalized_semantic(item)
            ),
            "displayName": " / ".join(names),
        }
    return result


def filter_inconsistent_employee_source_writes(
    plan: dict,
    *,
    source_layouts: dict[str, dict[str, dict[str, Any]]],
    template_manifest: dict | None = None,
) -> dict:
    """Drop summary-row facts and cross-employee contamination before -L writes.

    source_layouts: fileId -> sheetName -> inspect_source_employee_layout(...)
    """
    resolved = copy.deepcopy(plan)
    writes = resolved.get("writes")
    if not isinstance(writes, list):
        return resolved
    issues = resolved.setdefault("issues", [])
    if not isinstance(issues, list):
        issues = []
        resolved["issues"] = issues

    name_columns: set[int] = set()
    if isinstance(template_manifest, dict):
        for item in template_manifest.get("columnContexts") or []:
            if isinstance(item, dict) and isinstance(item.get("column"), int) and (
                _is_person_name_label(item.get("primaryLabel"))
                or _is_person_name_label(item.get("pathLabel"))
            ):
                name_columns.add(int(item["column"]))
    prefilled = _prefilled_identity_rows(template_manifest)

    def _layout_for(item: dict) -> tuple[str | None, int | None, dict | None]:
        source = item.get("source") if isinstance(item, dict) else None
        if not isinstance(source, dict):
            return None, None, None
        parsed = _parse_sheet_coordinate(source.get("location"))
        if not parsed:
            return None, None, None
        sheet_name, coordinate = parsed
        try:
            row, _column = coordinate_to_tuple(coordinate)
        except (TypeError, ValueError):
            return sheet_name, None, None
        layout = (source_layouts.get(str(source.get("fileId"))) or {}).get(sheet_name)
        return sheet_name, row, layout

    kept: list[dict] = []
    # Pass 1: drop facts taken from nameless source rows when the sheet has name columns.
    for item in writes:
        if not isinstance(item, dict):
            continue
        sheet_name, source_row, layout = _layout_for(item)
        if layout and source_row is not None and layout.get("nameColumns"):
            row_info = (layout.get("rows") or {}).get(source_row) or {}
            if not row_info.get("hasName"):
                issues.append({
                    "code": "SUMMARY_SOURCE_ROW_SKIPPED",
                    "message": (
                        f"Skipped source {item.get('source', {}).get('location')}: "
                        "row has no employee name (summary/total/padding row)"
                    ),
                    "sourceLabel": item.get("sourceLabel"),
                    "targetCell": item.get("targetCell"),
                    "sourceLocation": (item.get("source") or {}).get("location"),
                })
                continue
        kept.append(item)

    # Pass 1b: when -L rows already have CODE/template names, only accept source
    # facts from the same person (name-token overlap). Prevents 王品涵 amounts on 楊文凱's row.
    if prefilled:
        identity_kept: list[dict] = []
        for item in kept:
            try:
                target_row, _target_column = coordinate_to_tuple(str(item.get("targetCell")))
            except (TypeError, ValueError):
                identity_kept.append(item)
                continue
            pref = prefilled.get(target_row)
            if not pref:
                identity_kept.append(item)
                continue
            sheet_name, source_row, layout = _layout_for(item)
            if not layout or source_row is None:
                identity_kept.append(item)
                continue
            row_info = (layout.get("rows") or {}).get(source_row) or {}
            source_tokens = _name_token_set(row_info.get("names") or [])
            if not source_tokens:
                identity_kept.append(item)
                continue
            if not (source_tokens & pref["tokens"]):
                issues.append({
                    "code": "CODE_IDENTITY_MISMATCH_SKIPPED",
                    "message": (
                        f"Skipped {item.get('source', {}).get('location')}: target row "
                        f"{target_row} is anchored to {pref.get('displayName')!r}, but "
                        f"source row belongs to "
                        f"{' / '.join(row_info.get('names') or [])!r}"
                    ),
                    "sourceLabel": item.get("sourceLabel"),
                    "targetCell": item.get("targetCell"),
                    "sourceLocation": (item.get("source") or {}).get("location"),
                })
                continue
            identity_kept.append(item)
        kept = identity_kept

    # Pass 2: each target row may bind to at most one named source employee row.
    # Prefer CODE/template prefilled identity, then person-name writes, then majority vote.
    row_anchor: dict[int, tuple[str, str, int, str]] = {}

    def _best_source_for_tokens(
        tokens: set[str],
    ) -> tuple[str, str, int, str] | None:
        matches: list[tuple[int, str, str, int, str]] = []
        for file_id, sheets in source_layouts.items():
            if not isinstance(sheets, dict):
                continue
            for sheet_name, layout in sheets.items():
                if not isinstance(layout, dict):
                    continue
                for source_row, row_info in (layout.get("rows") or {}).items():
                    if not isinstance(row_info, dict) or not row_info.get("hasName"):
                        continue
                    source_tokens = _name_token_set(row_info.get("names") or [])
                    overlap = len(tokens & source_tokens)
                    if overlap <= 0:
                        continue
                    matches.append((
                        overlap,
                        str(file_id),
                        str(sheet_name),
                        int(source_row),
                        str(row_info.get("nameKey") or ""),
                    ))
        if not matches:
            return None
        matches.sort(key=lambda item: (-item[0], item[3]))
        best = matches[0]
        if len(matches) > 1 and matches[1][0] == best[0] and (
            matches[1][1], matches[1][2], matches[1][3]
        ) != (best[1], best[2], best[3]):
            return None
        return best[1], best[2], best[3], best[4]

    for target_row, pref in prefilled.items():
        matched = _best_source_for_tokens(pref["tokens"])
        if matched is not None:
            row_anchor[target_row] = matched

    for item in kept:
        try:
            target_row, target_column = coordinate_to_tuple(str(item.get("targetCell")))
        except (TypeError, ValueError):
            continue
        if target_row in row_anchor:
            continue
        if name_columns and target_column not in name_columns:
            continue
        if not name_columns and not _is_person_name_label(item.get("semanticLabel")) \
                and not _is_person_name_label(item.get("sourceLabel")):
            continue
        sheet_name, source_row, layout = _layout_for(item)
        if not layout or source_row is None or not sheet_name:
            continue
        row_info = (layout.get("rows") or {}).get(source_row) or {}
        if not row_info.get("hasName"):
            continue
        file_id = str((item.get("source") or {}).get("fileId") or "")
        name_key = str(row_info.get("nameKey") or "")
        row_anchor[target_row] = (file_id, sheet_name, source_row, name_key)

    # If no identity write anchored the row, use the majority named source row on that target.
    unanchored_votes: dict[int, dict[tuple[str, str, int, str], int]] = {}
    for item in kept:
        try:
            target_row, _target_column = coordinate_to_tuple(str(item.get("targetCell")))
        except (TypeError, ValueError):
            continue
        if target_row in row_anchor:
            continue
        sheet_name, source_row, layout = _layout_for(item)
        if not layout or source_row is None or not sheet_name:
            continue
        row_info = (layout.get("rows") or {}).get(source_row) or {}
        if not row_info.get("hasName"):
            continue
        file_id = str((item.get("source") or {}).get("fileId") or "")
        key = (file_id, sheet_name, source_row, str(row_info.get("nameKey") or ""))
        bucket = unanchored_votes.setdefault(target_row, {})
        bucket[key] = bucket.get(key, 0) + 1
    for target_row, bucket in unanchored_votes.items():
        winner = max(bucket.items(), key=lambda pair: (pair[1], -pair[0][2]))[0]
        row_anchor[target_row] = winner

    # Pass 3: drop writes that conflict with the anchored employee source row.
    consistent: list[dict] = []
    for item in kept:
        try:
            target_row, _target_column = coordinate_to_tuple(str(item.get("targetCell")))
        except (TypeError, ValueError):
            consistent.append(item)
            continue
        anchor = row_anchor.get(target_row)
        sheet_name, source_row, layout = _layout_for(item)
        if not anchor or not layout or source_row is None or not sheet_name:
            consistent.append(item)
            continue
        file_id = str((item.get("source") or {}).get("fileId") or "")
        anchor_file, anchor_sheet, anchor_row, _anchor_name = anchor
        if file_id == anchor_file and sheet_name == anchor_sheet and source_row != anchor_row:
            issues.append({
                "code": "SOURCE_ROW_MISMATCH_SKIPPED",
                "message": (
                    f"Skipped {item.get('source', {}).get('location')}: target row {target_row} "
                    f"is anchored to {anchor_sheet}!{anchor_row}, not mixed source rows"
                ),
                "sourceLabel": item.get("sourceLabel"),
                "targetCell": item.get("targetCell"),
                "sourceLocation": (item.get("source") or {}).get("location"),
            })
            continue
        consistent.append(item)

    # Pass 4: one employee identity -> one target row (keep the lowest row number).
    claimed_names: dict[str, int] = {}
    deduped: list[dict] = []
    blocked_rows: set[int] = set()
    for target_row, (_file_id, _sheet, _source_row, name_key) in sorted(row_anchor.items()):
        if not name_key:
            continue
        if name_key in claimed_names:
            blocked_rows.add(target_row)
            issues.append({
                "code": "DUPLICATE_EMPLOYEE_ROW_SKIPPED",
                "message": (
                    f"Skipped target row {target_row}: employee {name_key!r} already mapped "
                    f"to row {claimed_names[name_key]}"
                ),
                "targetCell": f"A{target_row}",
                "sourceLabel": "employee",
            })
        else:
            claimed_names[name_key] = target_row

    for item in consistent:
        try:
            target_row, _target_column = coordinate_to_tuple(str(item.get("targetCell")))
        except (TypeError, ValueError):
            deduped.append(item)
            continue
        if target_row in blocked_rows:
            continue
        deduped.append(item)

    # Pass 5: drop amount-only target rows that never received a person name.
    # Prefilled CODE/template identities count as named rows.
    rows_with_name: set[int] = set(prefilled)
    rows_with_writes: set[int] = set()
    for item in deduped:
        try:
            target_row, target_column = coordinate_to_tuple(str(item.get("targetCell")))
        except (TypeError, ValueError):
            continue
        rows_with_writes.add(target_row)
        if (name_columns and target_column in name_columns) or _is_person_name_label(
            item.get("semanticLabel")
        ) or _is_person_name_label(item.get("sourceLabel")):
            rows_with_name.add(target_row)
    # Only prune unnamed amount rows when the plan already identified at least one
    # named employee. Otherwise single-field / PDF probes would be wiped out.
    orphan_rows = (rows_with_writes - rows_with_name) if rows_with_name else set()
    final_writes: list[dict] = []
    for item in deduped:
        try:
            target_row, _target_column = coordinate_to_tuple(str(item.get("targetCell")))
        except (TypeError, ValueError):
            final_writes.append(item)
            continue
        if target_row in orphan_rows:
            issues.append({
                "code": "ORPHAN_EMPLOYEE_ROW_SKIPPED",
                "message": (
                    f"Skipped target row {target_row}: pay fields without an employee name"
                ),
                "sourceLabel": item.get("sourceLabel"),
                "targetCell": item.get("targetCell"),
            })
            continue
        final_writes.append(item)

    # De-dupe repeated orphan/duplicate issue messages per target row.
    compact_issues: list[dict] = []
    seen_issue_keys: set[tuple[Any, ...]] = set()
    for item in issues:
        if not isinstance(item, dict):
            continue
        key = (item.get("code"), item.get("targetCell"), item.get("sourceLocation"), item.get("message"))
        if key in seen_issue_keys:
            continue
        seen_issue_keys.add(key)
        compact_issues.append(item)
    resolved["writes"] = final_writes
    resolved["issues"] = compact_issues
    return annotate_missing_source_employees(resolved, source_layouts)


def extract_last_l_identity_rows(
    workbook_path: str | Path,
    *,
    data_start_row: int | None = None,
) -> dict[str, Any]:
    """Read person-name cells from a workbook's last ``-L`` sheet (typically CODE output)."""
    path = Path(workbook_path)
    require(path.is_file(), "Identity source workbook is missing")
    manifest = inspect_last_l_sheet(path)
    if data_start_row is not None:
        manifest["dataStartRow"] = int(data_start_row)
    name_contexts = [
        item for item in manifest.get("columnContexts") or []
        if isinstance(item, dict) and isinstance(item.get("column"), int) and (
            _is_person_name_label(item.get("primaryLabel"))
            or _is_person_name_label(item.get("pathLabel"))
        )
    ]
    require(bool(name_contexts), "CODE/result -L sheet has no person-name columns")
    wb = load_workbook(path, data_only=True, read_only=False, keep_links=False)
    try:
        ws = wb[manifest["sheetName"]]
        header_row = int(manifest["headerRow"])
        start_row = _identity_data_start_row(manifest, header_row)
        rows: list[dict[str, Any]] = []
        max_row = min(max(1, ws.max_row or 1), header_row + 500)
        for row in range(start_row, max_row + 1):
            fields = []
            names = []
            for context in name_contexts:
                column = int(context["column"])
                text = _clean_cell_text(ws.cell(row, column).value)
                if not text:
                    continue
                label = str(context.get("primaryLabel") or "")
                names.append(text)
                fields.append({
                    "column": column,
                    "columnLetter": get_column_letter(column),
                    "label": label,
                    "value": text,
                    "cell": get_column_letter(column) + str(row),
                })
            if not fields:
                continue
            rows.append({
                "row": row,
                "names": names,
                "displayName": " / ".join(names),
                "nameKey": " / ".join(
                    _normalized_semantic(item) for item in names if _normalized_semantic(item)
                ),
                "fields": fields,
            })
        return {
            "sheetName": manifest["sheetName"],
            "headerRow": header_row,
            "dataStartRow": start_row,
            "nameColumns": [int(item["column"]) for item in name_contexts],
            "employees": rows,
        }
    finally:
        wb.close()


def prepare_template_with_code_identities(
    template_path: str | Path,
    code_result_path: str | Path,
    output_path: str | Path,
    *,
    data_start_row: int | None = None,
) -> dict[str, Any]:
    """Step 1: copy the template and pre-fill -L person names from the CODE result.

    Only identity names come from CODE. Template-native formulas stay as-is; literal
    amounts are never copied. Blank and zero remain equivalent for later AI filling.
    """
    template = Path(template_path).resolve()
    code_path = Path(code_result_path).resolve()
    output = Path(output_path).resolve()
    require(template.is_file(), "Template is missing")
    require(code_path.is_file(), "CODE result is missing")
    require(template != output, "Prepared template must not overwrite the source template")
    require(not output.exists(), "Prepared template output already exists")
    identity = extract_last_l_identity_rows(code_path, data_start_row=data_start_row)
    employees = identity.get("employees") or []
    require(bool(employees), "CODE result -L sheet has no employee names to anchor")
    template_manifest = inspect_last_l_sheet(template)
    require(
        template_manifest["sheetName"] == identity["sheetName"],
        "CODE result and template last -L sheet names differ",
    )
    output.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(template, output)
    written = 0
    reused = 0
    cleared = 0
    anchored: list[dict[str, Any]] = []
    resolved_data_start = identity.get("dataStartRow")
    try:
        wb = load_workbook(output, data_only=False, read_only=False, keep_links=False)
        try:
            ws = wb[template_manifest["sheetName"]]
            header_row = int(template_manifest.get("headerRow") or identity.get("headerRow") or 1)
            start_row = _identity_data_start_row(
                {"dataStartRow": identity.get("dataStartRow") if data_start_row is None else data_start_row},
                header_row,
            )
            resolved_data_start = start_row
            # Drop master sample literals in the employee block so AI starts blank
            # (formulas stay). Otherwise occupied sample amounts make writes=0.
            max_row = min(max(1, ws.max_row or 1), start_row + 500)
            max_col = max(1, ws.max_column or 1)
            for row in range(start_row, max_row + 1):
                for col in range(1, max_col + 1):
                    cell = ws.cell(row, col)
                    if isinstance(cell, MergedCell) or cell.data_type == "f" or cell.value is None:
                        continue
                    cell.value = None
                    cleared += 1
            for employee in employees:
                row = int(employee["row"])
                row_fields = []
                for field in employee.get("fields") or []:
                    column = int(field["column"])
                    cell_ref = get_column_letter(column) + str(row)
                    cell = ws[cell_ref]
                    if isinstance(cell, MergedCell) or cell.data_type == "f":
                        continue
                    code_name = " ".join(str(field.get("value") or "").split()).strip()
                    if not code_name:
                        continue
                    existing = _clean_cell_text(cell.value)
                    if existing and _normalized_semantic(existing) == _normalized_semantic(code_name):
                        # Same name survived clear (unusual); count as reuse.
                        reused += 1
                    else:
                        cell.value = code_name
                        written += 1
                    row_fields.append({
                        "column": column,
                        "columnLetter": get_column_letter(column),
                        "label": field.get("label"),
                        "value": code_name,
                        "cell": cell_ref,
                    })
                if row_fields:
                    anchored.append({
                        "row": row,
                        "displayName": employee.get("displayName"),
                        "names": list(employee.get("names") or []),
                        "fields": row_fields,
                    })
            require(bool(anchored), "No CODE identity cells could be written into the template")
            for sheet in wb.worksheets:
                sheet.sheet_view.tabSelected = False
            wb.active = template_manifest["sheetIndex"]
            ws.sheet_view.tabSelected = True
            wb.save(output)
        finally:
            wb.close()
    except Exception:
        if output.exists():
            output.unlink()
        raise
    return {
        "sheetName": identity["sheetName"],
        "headerRow": identity.get("headerRow"),
        "dataStartRow": resolved_data_start,
        "employeeCount": len(anchored),
        "writeCount": written,
        "reuseCount": reused,
        "clearedSampleCount": cleared,
        "employees": anchored,
        "outputPath": str(output),
        "outputSha256": _sha256(output.read_bytes()),
    }


def _is_zero_like_write(item: dict) -> bool:
    if item.get("valueType") != "decimal":
        return False
    try:
        number = Decimal(decimal_text(item.get("value")))
    except (ValidationError, InvalidOperation, ValueError, TypeError):
        return False
    return number == 0


def resolve_dynamic_template_fill_targets(plan: dict, template_manifest: dict, *,
                                          column_mappings: dict[str, str] | None = None) -> dict:
    """Resolve target columns in code while retaining the model-selected row.

    A reviewed supplier/customer mapping wins.  Otherwise a uniquely matching
    current-template label is used.  Ambiguous columns stay untouched so the
    normal validator can accept or reject the model's candidate.
    """
    resolved = copy.deepcopy(plan)
    contexts = [item for item in template_manifest.get("columnContexts") or []
                if isinstance(item, dict) and isinstance(item.get("column"), int)]
    writes = resolved.get("writes")
    if not isinstance(writes, list):
        return resolved
    issues = resolved.setdefault("issues", [])
    if not isinstance(issues, list):
        issues = []
        resolved["issues"] = issues
    kept: list[dict] = []
    for item in writes:
        if not isinstance(item, dict) or not nonempty(item.get("sourceLabel")):
            if isinstance(item, dict):
                kept.append(item)
            continue
        if (
            _is_service_fee_label(item.get("sourceLabel"))
            or _is_service_fee_label(item.get("semanticLabel"))
        ):
            issues.append({
                "code": "SERVICE_FEE_SKIPPED",
                "message": (
                    f"Skipped Service Fee field {item.get('sourceLabel')!r}: "
                    "AI comparison intentionally leaves Service Fee blank"
                ),
                "sourceLabel": item.get("sourceLabel"),
                "targetCell": item.get("targetCell"),
            })
            continue
        if _is_zero_like_write(item):
            issues.append({
                "code": "ZERO_EQUIV_BLANK_SKIPPED",
                "message": (
                    f"Skipped zero value for {item.get('sourceLabel')!r}: "
                    "blank and 0 are equivalent; leave the cell empty"
                ),
                "sourceLabel": item.get("sourceLabel"),
                "targetCell": item.get("targetCell"),
            })
            continue
        try:
            row, column = coordinate_to_tuple(str(item.get("targetCell")))
        except (TypeError, ValueError):
            kept.append(item)
            continue
        target_occupied = next(
            (cell for cell in (template_manifest.get("nonemptyCells") or [])
             if isinstance(cell, dict) and cell.get("cell") == item.get("targetCell")),
            None,
        )
        if target_occupied and target_occupied.get("valueKind") == "formula":
            issues.append({
                "code": "FORMULA_CELL_SKIPPED",
                "message": (
                    f"Skipped {item.get('sourceLabel')!r} → {item.get('targetCell')}: "
                    "formula cells belong to the template and must not be AI-filled"
                ),
                "sourceLabel": item.get("sourceLabel"),
                "targetCell": item.get("targetCell"),
            })
            continue
        source_header = str(item["sourceLabel"])
        exact = _exact_header_column(source_header, contexts)
        configured = None if exact is not None else _configured_target(
            source_header, column_mappings,
        )
        if exact is not None:
            lookup = str(exact.get("pathLabel") or exact.get("primaryLabel") or source_header)
        else:
            lookup = configured or source_header
        if _is_service_fee_label(lookup):
            issues.append({
                "code": "SERVICE_FEE_SKIPPED",
                "message": (
                    f"Skipped Service Fee field {item.get('sourceLabel')!r}: "
                    "AI comparison intentionally leaves Service Fee blank"
                ),
                "sourceLabel": item.get("sourceLabel"),
                "targetCell": item.get("targetCell"),
            })
            continue
        context = exact or _unambiguous_column(lookup, contexts)
        if context is not None and _is_service_fee_label(context.get("primaryLabel")):
            issues.append({
                "code": "SERVICE_FEE_SKIPPED",
                "message": (
                    f"Skipped Service Fee field {item.get('sourceLabel')!r}: "
                    "AI comparison intentionally leaves Service Fee blank"
                ),
                "sourceLabel": item.get("sourceLabel"),
                "targetCell": item.get("targetCell"),
            })
            continue
        # Never retarget into a template-native formula cell.
        if context is not None:
            candidate = get_column_letter(int(context["column"])) + str(row)
            occupied = next(
                (cell for cell in (template_manifest.get("nonemptyCells") or [])
                 if isinstance(cell, dict) and cell.get("cell") == candidate),
                None,
            )
            if occupied and occupied.get("valueKind") == "formula":
                issues.append({
                    "code": "FORMULA_CELL_SKIPPED",
                    "message": (
                        f"Skipped {item.get('sourceLabel')!r} → {candidate}: "
                        "formula cells belong to the template and must not be AI-filled"
                    ),
                    "sourceLabel": item.get("sourceLabel"),
                    "targetCell": candidate,
                })
                continue
        if context is None and _is_identity_label(lookup):
            # Identity fields need an exact/high-confidence column; never keep a shifted guess.
            scored = sorted(
                ((_context_score(lookup, ctx), ctx) for ctx in contexts),
                key=lambda pair: pair[0], reverse=True,
            )
            if scored and scored[0][0] >= 0.85:
                second = scored[1][0] if len(scored) > 1 else 0.0
                if scored[0][0] >= second + 0.05:
                    context = scored[0][1]
        if context is None:
            # Drop clearly wrong model coordinates instead of writing shifted identity/pay cells.
            model_ctx = next((ctx for ctx in contexts if ctx.get("column") == column), None)
            model_score = _context_score(lookup, model_ctx) if model_ctx else 0.0
            if model_score < 0.35:
                issues.append({
                    "code": "UNRESOLVED_COLUMN_SKIPPED",
                    "message": (
                        f"Skipped source field {item.get('sourceLabel')!r}: "
                        "no unambiguous template column"
                    ),
                    "sourceLabel": item.get("sourceLabel"),
                    "targetCell": item.get("targetCell"),
                })
                continue
            kept.append(item)
            continue
        item["targetCell"] = get_column_letter(context["column"]) + str(row)
        item["semanticLabel"] = context["primaryLabel"]
        kept.append(item)
    resolved["writes"] = kept
    return resolved


def inspect_last_l_sheet(template_path: str | Path, *, max_nonempty_cells: int = 20000,
                         header_row: int | None = None) -> dict:
    """Return a bounded, read-only manifest for the last sheet ending in ``-L``."""
    path = Path(template_path)
    raw = path.read_bytes()
    wb = load_workbook(BytesIO(raw), data_only=False, read_only=False, keep_links=False)
    try:
        sheet_name = _last_l_sheet(list(wb.sheetnames))
        ws = wb[sheet_name]
        header_row, column_contexts = _column_contexts(ws, header_row=header_row)
        cells = []
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell, MergedCell) or cell.value is None:
                    continue
                require(len(cells) < max_nonempty_cells, "Template -L sheet exceeds inspection limit")
                value = cell.value
                cells.append({
                    "cell": cell.coordinate,
                    "row": cell.row,
                    "column": cell.column,
                    "value": str(value),
                    "valueKind": "formula" if cell.data_type == "f" else "literal",
                    "numberFormat": cell.number_format,
                })
        return {
            "manifestVersion": 1,
            "templateSha256": _sha256(raw),
            "sheetName": sheet_name,
            "sheetIndex": wb.sheetnames.index(sheet_name),
            "maxRow": ws.max_row,
            "maxColumn": ws.max_column,
            "headerRow": header_row,
            "columnContexts": column_contexts,
            "mergedRanges": [str(item) for item in ws.merged_cells.ranges],
            "nonemptyCells": cells,
        }
    finally:
        wb.close()


def validate_template_fill_plan(plan: dict, template_manifest: dict, collection: dict, schema: dict,
                                documents: list[dict]) -> None:
    """Validate model-proposed locations without trusting model-proposed values."""
    validate_ai_collection(collection, schema, documents)
    require(isinstance(plan, dict) and plan.get("planVersion") == 1, "Invalid template fill plan")
    require(plan.get("templateSha256") == template_manifest.get("templateSha256"),
            "Template fill plan is stale")
    require(plan.get("sheetName") == template_manifest.get("sheetName"),
            "Template fill plan targets the wrong worksheet")
    require(plan.get("automaticWriteEnabled") is False,
            "AI fill plan cannot authorize its own workbook write")
    writes = plan.get("writes")
    require(isinstance(writes, list), "Template fill writes must be a list")
    seen = set()
    records = collection["records"]
    schema_fields = {item["field"] for item in schema["fields"]}
    occupied = {item["cell"]: item for item in template_manifest.get("nonemptyCells") or []}
    max_row = template_manifest.get("maxRow")
    max_column = template_manifest.get("maxColumn")
    require(type(max_row) is int and type(max_column) is int and max_row >= 1 and max_column >= 1,
            "Invalid template worksheet bounds")
    for item in writes:
        require(isinstance(item, dict), "Invalid template fill write")
        cell = item.get("targetCell")
        try:
            row, column = coordinate_to_tuple(str(cell))
        except (TypeError, ValueError):
            raise ValidationError("Invalid target cell") from None
        require(1 <= row <= 1048576 and 1 <= column <= 16384, "Target cell exceeds Excel limits")
        require(column <= max_column and row <= max_row + 2000,
                "Target cell is outside the inspected -L table")
        require(cell not in seen, "Duplicate template fill target")
        seen.add(cell)
        existing = occupied.get(cell)
        if existing is not None:
            require(existing.get("valueKind") != "formula", "Cannot overwrite a formula cell")
            raise ValidationError("Cannot overwrite a non-empty template cell")
        record_index = item.get("recordIndex")
        require(type(record_index) is int and 0 <= record_index < len(records), "Invalid source record index")
        source_kind = item.get("sourceKind")
        require(source_kind in {"entity_name", "field"}, "Unsupported template fill source")
        allowed = {"targetCell", "recordIndex", "sourceKind", "field"}
        if source_kind == "field":
            require(item.get("field") in schema_fields, "Unknown template fill source field")
        else:
            require(item.get("field") is None, "Entity name write cannot reference a field")
        require(set(item) == allowed, "Template fill write contains unsupported data")
    issues = plan.get("issues")
    require(isinstance(issues, list)
            and all(isinstance(item, dict) and nonempty(item.get("code")) for item in issues),
            "Template fill issues must contain structured codes")


def _resolve_value(item: dict, collection: dict) -> str | int | float:
    record = collection["records"][item["recordIndex"]]
    if item["sourceKind"] == "entity_name":
        name = record["sourceEntity"].get("name")
        require(nonempty(name), "AI entity name is unavailable")
        return str(name)
    field = next(value for value in record["fields"] if value["field"] == item["field"])
    require(field["status"] == "found", "AI field is unavailable: " + item["field"])
    value = decimal_text(field["value"])
    number = Decimal(value)
    return int(number) if number == number.to_integral_value() else float(number)


def write_ai_template_copy(template_path: str | Path, output_path: str | Path, plan: dict, collection: dict,
                           schema: dict, documents: list[dict]) -> dict:
    """Write a new comparison workbook; never overwrite the template or an existing output."""
    source = Path(template_path).resolve()
    output = Path(output_path).resolve()
    require(source != output, "AI output must not overwrite the template")
    require(not output.exists(), "AI output already exists")
    manifest = inspect_last_l_sheet(source)
    validate_template_fill_plan(plan, manifest, collection, schema, documents)

    output.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, output)
    try:
        wb = load_workbook(output, data_only=False, read_only=False, keep_links=False)
        try:
            require(_last_l_sheet(list(wb.sheetnames)) == manifest["sheetName"],
                    "AI output worksheet order changed")
            ws = wb[manifest["sheetName"]]
            for item in plan["writes"]:
                cell = ws[item["targetCell"]]
                require(not isinstance(cell, MergedCell), "Cannot write to a merged placeholder")
                require(cell.data_type != "f", "Cannot overwrite a formula cell")
                cell.value = _resolve_value(item, collection)
            # The review UI opens the workbook's active sheet. Make the dynamically
            # selected last -L sheet visible immediately without removing other sheets.
            for sheet in wb.worksheets:
                sheet.sheet_view.tabSelected = False
            wb.active = manifest["sheetIndex"]
            ws.sheet_view.tabSelected = True
            wb.save(output)
        finally:
            wb.close()
    except Exception:
        if output.exists():
            output.unlink()
        raise
    return {
        "templateSha256": manifest["templateSha256"],
        "outputSha256": _sha256(output.read_bytes()),
        "sheetName": manifest["sheetName"],
        "writeCount": len(plan["writes"]),
        "issueCount": len(plan["issues"]),
        "automaticPassEnabled": False,
    }


def validate_dynamic_template_fill_plan(plan: dict, template_manifest: dict, documents: list[dict], *,
                                        column_mappings: dict[str, str] | None = None) -> None:
    """Validate template-driven values, locations and source evidence before any workbook write."""
    require(isinstance(plan, dict) and plan.get("planVersion") == 3, "Invalid dynamic template fill plan")
    require(plan.get("templateSha256") == template_manifest.get("templateSha256"),
            "Dynamic template fill plan is stale")
    require(plan.get("sheetName") == template_manifest.get("sheetName"),
            "Dynamic template fill plan targets the wrong worksheet")
    require(plan.get("automaticWriteEnabled") is False,
            "AI fill plan cannot authorize its own workbook write")
    require(nonempty(plan.get("model")), "Dynamic template fill model is required")
    writes = plan.get("writes")
    require(isinstance(writes, list), "Dynamic template fill writes must be a list")
    known_files = {item["fileId"] for item in documents}
    occupied = {item["cell"]: item for item in template_manifest.get("nonemptyCells") or []}
    max_row = template_manifest.get("maxRow")
    max_column = template_manifest.get("maxColumn")
    require(type(max_row) is int and type(max_column) is int and max_row >= 1 and max_column >= 1,
            "Invalid template worksheet bounds")
    expected = {"targetCell", "semanticLabel", "sourceLabel", "valueType", "value", "source", "confidence"}
    column_contexts = {
        item.get("column"): item for item in template_manifest.get("columnContexts") or []
        if isinstance(item, dict) and isinstance(item.get("column"), int)
    }
    context_list = list(column_contexts.values())
    issues = plan.get("issues")
    require(isinstance(issues, list), "Dynamic template fill issues must be a list")
    seen = set()
    kept: list[dict] = []
    for item in writes:
        require(isinstance(item, dict) and set(item) == expected,
                "Dynamic template fill write contains unsupported data")
        cell = item.get("targetCell")
        try:
            row, column = coordinate_to_tuple(str(cell))
        except (TypeError, ValueError):
            raise ValidationError("Invalid target cell") from None
        require(1 <= row <= 1048576 and 1 <= column <= 16384, "Target cell exceeds Excel limits")
        require(column <= max_column and row <= max_row + 2000,
                "Target cell is outside the inspected -L table")
        if cell in seen:
            issues.append({
                "code": "DUPLICATE_TARGET_SKIPPED",
                "message": f"Skipped duplicate write to {cell}",
                "targetCell": cell,
                "sourceLabel": item.get("sourceLabel"),
            })
            continue
        seen.add(cell)
        existing = occupied.get(cell)
        if existing is not None:
            if existing.get("valueKind") == "formula":
                issues.append({
                    "code": "FORMULA_CELL_SKIPPED",
                    "message": (
                        f"Skipped write to {cell}: formula cells belong to the template "
                        "and must not be AI-filled"
                    ),
                    "targetCell": cell,
                    "sourceLabel": item.get("sourceLabel"),
                })
                continue
            raise ValidationError("Cannot overwrite a non-empty template cell")
        context = column_contexts.get(column)
        require(context is not None and nonempty(context.get("primaryLabel")),
                "Dynamic target column has no unambiguous semantic label")
        semantic_label = item.get("semanticLabel")
        source_label = item.get("sourceLabel")
        require(nonempty(semantic_label) and nonempty(source_label),
                "Dynamic write semantic labels are required")
        path_label = context.get("pathLabel") or context.get("primaryLabel")
        if (
            _is_service_fee_label(source_label)
            or _is_service_fee_label(semantic_label)
            or _is_service_fee_label(path_label)
            or _is_service_fee_label(context.get("primaryLabel"))
        ):
            issues.append({
                "code": "SERVICE_FEE_SKIPPED",
                "message": (
                    f"Skipped Service Fee field {source_label!r}: "
                    "AI comparison intentionally leaves Service Fee blank"
                ),
                "targetCell": cell,
                "sourceLabel": source_label,
            })
            continue
        if _is_zero_like_write(item):
            issues.append({
                "code": "ZERO_EQUIV_BLANK_SKIPPED",
                "message": (
                    f"Skipped zero value for {source_label!r}: "
                    "blank and 0 are equivalent; leave the cell empty"
                ),
                "targetCell": cell,
                "sourceLabel": source_label,
            })
            continue
        require(
            _normalized_semantic(semantic_label) == _normalized_semantic(context.get("primaryLabel"))
            or _normalized_semantic(semantic_label) == _normalized_semantic(path_label),
            "Dynamic write semantic label does not match the target column",
        )
        value_type = item.get("valueType")
        value = item.get("value")
        require(value_type in {"text", "decimal", "date"} and nonempty(value),
                "Dynamic write value is invalid")
        if value_type == "decimal":
            try:
                decimal_text(value)
            except (ValidationError, InvalidOperation, ValueError):
                raise ValidationError("Dynamic decimal value is invalid") from None
        elif value_type == "date":
            try:
                date.fromisoformat(str(value))
            except (TypeError, ValueError):
                raise ValidationError("Dynamic date value must use YYYY-MM-DD") from None
        source = item.get("source")
        require(isinstance(source, dict)
                and set(source) == {"fileId", "location", "page", "rawText"}
                and source.get("fileId") in known_files
                and nonempty(source.get("location"))
                and nonempty(source.get("rawText")), "Dynamic write source evidence is invalid")
        require(_normalized_semantic(source_label) in _normalized_semantic(source.get("rawText")),
                "Dynamic source evidence must include the source field label")
        configured_target = _configured_target(str(source_label), column_mappings)
        comparison_label = configured_target or source_label
        selected_score = _context_score(comparison_label, context)
        scored_labels = [(_context_score(comparison_label, item), item) for item in context_list]
        best_score, best_context = max(
            scored_labels,
            key=lambda pair: (pair[0], pair[1].get("column") or 0),
            default=(0.0, {}),
        )
        best_label = (best_context or {}).get("pathLabel") or (best_context or {}).get("primaryLabel") or ""
        # Wrong-column guesses (e.g. section header "Pay Items") must not abort the
        # whole AI -L workbook. Drop the write and keep generating the rest.
        if best_score >= 0.35 and selected_score + 0.05 < best_score:
            issues.append({
                "code": "COLUMN_MISMATCH_SKIPPED",
                "message": (
                    f"Skipped source field {source_label!r}: it targeted "
                    f"{path_label!r}, but the closest template column is "
                    f"{best_label!r}"
                    + (f" (configured as {configured_target!r})" if configured_target else "")
                ),
                "targetCell": cell,
                "sourceLabel": source_label,
                "selectedLabel": path_label,
                "bestLabel": best_label,
            })
            continue
        confidence = item.get("confidence")
        require(type(confidence) in {int, float} and 0 <= confidence <= 1,
                "Dynamic write confidence is invalid")
        kept.append(item)
    plan["writes"] = kept
    require(all(isinstance(item, dict) and nonempty(item.get("code")) for item in issues),
            "Dynamic template fill issues must contain structured codes")

def _dynamic_value(item: dict) -> str | int | float | date:
    if item["valueType"] == "text":
        return str(item["value"])
    if item["valueType"] == "date":
        return date.fromisoformat(str(item["value"]))
    number = Decimal(decimal_text(item["value"]))
    return int(number) if number == number.to_integral_value() else float(number)


def normalize_code_provenance_cells(raw: Any) -> list[dict[str, Any]]:
    """Parse Office ``codeProvenanceCells`` (Excel 1-based sheet/row/col)."""
    if raw is None:
        return []
    require(isinstance(raw, list), "AI codeProvenanceCells must be a list")
    require(len(raw) <= 500, "AI codeProvenanceCells exceeds the entry limit")
    result: list[dict[str, Any]] = []
    seen: set[tuple[str, int, int]] = set()
    for item in raw:
        if not isinstance(item, dict):
            continue
        sheet = str(item.get("sheet") or "")
        try:
            row = int(item.get("row"))
            col = int(item.get("col"))
        except (TypeError, ValueError):
            continue
        if not sheet or row < 1 or col < 1:
            continue
        key = (sheet, row, col)
        if key in seen:
            continue
        seen.add(key)
        entry: dict[str, Any] = {"sheet": sheet, "row": row, "col": col}
        kind = " ".join(str(item.get("kind") or "").split()).strip()
        label = " ".join(str(item.get("label") or "").split()).strip()
        if kind:
            entry["kind"] = kind
        if label:
            entry["label"] = label
        result.append(entry)
    return result


def strip_code_provenance_writes(
    plan: dict,
    provenance_cells: list[dict[str, Any]],
    sheet_name: str,
) -> dict:
    """Drop model writes that target CODE special-source cells on the active -L sheet."""
    resolved = copy.deepcopy(plan)
    writes = resolved.get("writes")
    if not isinstance(writes, list) or not provenance_cells:
        return resolved
    reserved = {
        (str(item["sheet"]), int(item["row"]), int(item["col"]))
        for item in provenance_cells
        if isinstance(item, dict)
    }
    kept: list[dict] = []
    issues = resolved.setdefault("issues", [])
    if not isinstance(issues, list):
        issues = []
        resolved["issues"] = issues
    for item in writes:
        if not isinstance(item, dict):
            continue
        try:
            row, column = coordinate_to_tuple(str(item.get("targetCell")))
        except (TypeError, ValueError):
            kept.append(item)
            continue
        if (sheet_name, row, column) in reserved:
            issues.append({
                "code": "CODE_PROVENANCE_CELL_SKIPPED",
                "message": (
                    f"Skipped {item.get('targetCell')}: CODE special-source cell "
                    f"(kind={next((c.get('kind') for c in provenance_cells if c.get('sheet') == sheet_name and int(c['row']) == row and int(c['col']) == column), None)})"
                ),
                "targetCell": item.get("targetCell"),
                "sourceLabel": item.get("sourceLabel"),
            })
            continue
        kept.append(item)
    resolved["writes"] = kept
    return resolved


def copy_code_provenance_cells(
    code_result_path: str | Path,
    ai_output_path: str | Path,
    provenance_cells: list[dict[str, Any]],
) -> dict[str, Any]:
    """Copy CODE special-source cells onto the AI workbook (any sheet)."""
    code_path = Path(code_result_path).resolve()
    ai_path = Path(ai_output_path).resolve()
    require(code_path.is_file(), "CODE result is missing for provenance copy")
    require(ai_path.is_file(), "AI output is missing for provenance copy")
    cells = normalize_code_provenance_cells(provenance_cells)
    if not cells:
        return {"copyCount": 0, "skippedCount": 0}
    copied = 0
    skipped = 0
    code_wb = load_workbook(code_path, data_only=False, read_only=False, keep_links=False)
    try:
        ai_wb = load_workbook(ai_path, data_only=False, read_only=False, keep_links=False)
        try:
            for item in cells:
                sheet = item["sheet"]
                row = int(item["row"])
                col = int(item["col"])
                if sheet not in code_wb.sheetnames or sheet not in ai_wb.sheetnames:
                    skipped += 1
                    continue
                src = code_wb[sheet].cell(row, col)
                dst = ai_wb[sheet].cell(row, col)
                if isinstance(dst, MergedCell):
                    skipped += 1
                    continue
                dst.value = src.value
                copied += 1
            ai_wb.save(ai_path)
        finally:
            ai_wb.close()
    finally:
        code_wb.close()
    return {"copyCount": copied, "skippedCount": skipped}


def sync_code_owned_regions_from_code_result(
    code_result_path: str | Path,
    ai_output_path: str | Path,
    *,
    provenance_cells: list[dict[str, Any]] | None = None,
    data_start_row: int | None = None,
) -> dict[str, Any]:
    """Make CODE-owned regions match the formal result after AI fills ``-L``.

    1. Copy every non-last-``-L`` sheet cell from CODE (PN / regional EE etc.).
    2. On the last ``-L`` sheet, copy CODE metadata rows above the employee block
       (period dates etc.). Employee formulas are copied only at explicit special
       coordinates; ordinary AI inputs and template formulas stay untouched.
    3. Copy explicit ``codeProvenanceCells`` coordinates when provided.
    """
    code_path = Path(code_result_path).resolve()
    ai_path = Path(ai_output_path).resolve()
    require(code_path.is_file(), "CODE result is missing for CODE-owned sync")
    require(ai_path.is_file(), "AI output is missing for CODE-owned sync")
    cells = normalize_code_provenance_cells(provenance_cells)
    non_l_copied = 0
    formula_copied = 0
    metadata_copied = 0
    provenance_copied = 0
    skipped = 0
    code_wb = load_workbook(code_path, data_only=False, read_only=False, keep_links=False)
    try:
        ai_wb = load_workbook(ai_path, data_only=False, read_only=False, keep_links=False)
        try:
            last_l = _last_l_sheet(list(code_wb.sheetnames))
            # Copy explicit special coordinates first, so their counts are not
            # swallowed by the broader metadata/non-L synchronization below.
            for item in cells:
                sheet = item["sheet"]
                row = int(item["row"])
                col = int(item["col"])
                if sheet not in code_wb.sheetnames or sheet not in ai_wb.sheetnames:
                    skipped += 1
                    continue
                src = code_wb[sheet].cell(row, col)
                dst = ai_wb[sheet].cell(row, col)
                if isinstance(dst, MergedCell):
                    skipped += 1
                    continue
                if dst.value != src.value:
                    dst.value = src.value
                    provenance_copied += 1
            for sheet_name in code_wb.sheetnames:
                if sheet_name not in ai_wb.sheetnames:
                    skipped += 1
                    continue
                code_ws = code_wb[sheet_name]
                ai_ws = ai_wb[sheet_name]
                if sheet_name.strip().casefold() != last_l.strip().casefold():
                    for row in code_ws.iter_rows(
                        min_row=1, max_row=code_ws.max_row or 1,
                        max_col=code_ws.max_column or 1,
                    ):
                        for src in row:
                            if isinstance(src, MergedCell) or src.value is None:
                                continue
                            dst = ai_ws.cell(src.row, src.column)
                            if isinstance(dst, MergedCell):
                                skipped += 1
                                continue
                            if dst.value != src.value:
                                dst.value = src.value
                                non_l_copied += 1
                    continue
                # Prefer configured dataStartRow: period/date rows above it often
                # outscore the real header and break header-based band detection.
                if data_start_row is not None:
                    try:
                        employee_start = max(1, int(data_start_row))
                    except (TypeError, ValueError):
                        employee_start = None
                else:
                    employee_start = None
                if employee_start is None:
                    header_row, _contexts = _column_contexts(code_ws)
                    employee_start = header_row + 1
                for row in code_ws.iter_rows(
                    min_row=1, max_row=code_ws.max_row or 1,
                    max_col=code_ws.max_column or 1,
                ):
                    for src in row:
                        if isinstance(src, MergedCell) or src.value is None:
                            continue
                        dst = ai_ws.cell(src.row, src.column)
                        if isinstance(dst, MergedCell):
                            skipped += 1
                            continue
                        # Employee formulas outside the special list must not
                        # overwrite AI values and hide comparison differences.
                        if src.row >= employee_start:
                            continue
                        if src.data_type == "f":
                            if dst.value != src.value:
                                dst.value = src.value
                                formula_copied += 1
                            continue
                        # Headers + period/metadata literals above the employee block.
                        if src.row < employee_start and dst.value != src.value:
                            dst.value = src.value
                            metadata_copied += 1
            ai_wb.save(ai_path)
        finally:
            ai_wb.close()
    finally:
        code_wb.close()
    return {
        "copyCount": non_l_copied + formula_copied + metadata_copied + provenance_copied,
        "nonLCopyCount": non_l_copied,
        "formulaCopyCount": formula_copied,
        "metadataCopyCount": metadata_copied,
        "provenanceCopyCount": provenance_copied,
        "skippedCount": skipped,
    }


def write_dynamic_ai_template_copy(template_path: str | Path, output_path: str | Path, plan: dict,
                                   documents: list[dict], *,
                                   column_mappings: dict[str, str] | None = None,
                                   code_provenance_cells: list[dict[str, Any]] | None = None) -> dict:
    """Write an evidence-backed dynamic plan to a new copy of the current template."""
    source = Path(template_path).resolve()
    output = Path(output_path).resolve()
    require(source != output, "AI output must not overwrite the template")
    require(not output.exists(), "AI output already exists")
    manifest = inspect_last_l_sheet(source)
    plan = resolve_dynamic_template_fill_targets(
        plan, manifest, column_mappings=column_mappings,
    )
    plan = strip_code_provenance_writes(
        plan, normalize_code_provenance_cells(code_provenance_cells), manifest["sheetName"],
    )
    validate_dynamic_template_fill_plan(
        plan, manifest, documents, column_mappings=column_mappings,
    )
    output.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, output)
    try:
        wb = load_workbook(output, data_only=False, read_only=False, keep_links=False)
        try:
            require(_last_l_sheet(list(wb.sheetnames)) == manifest["sheetName"],
                    "AI output worksheet order changed")
            ws = wb[manifest["sheetName"]]
            for item in plan["writes"]:
                cell = ws[item["targetCell"]]
                require(not isinstance(cell, MergedCell), "Cannot write to a merged placeholder")
                require(cell.data_type != "f" and cell.value is None, "Cannot overwrite a template cell")
                cell.value = _dynamic_value(item)
            for sheet in wb.worksheets:
                sheet.sheet_view.tabSelected = False
            wb.active = manifest["sheetIndex"]
            ws.sheet_view.tabSelected = True
            wb.save(output)
        finally:
            wb.close()
    except Exception:
        if output.exists():
            output.unlink()
        raise
    return {
        "templateSha256": manifest["templateSha256"],
        "outputSha256": _sha256(output.read_bytes()),
        "sheetName": manifest["sheetName"],
        "writeCount": len(plan["writes"]),
        "issueCount": len(plan["issues"]),
        "automaticPassEnabled": False,
        "plan": plan,
    }
