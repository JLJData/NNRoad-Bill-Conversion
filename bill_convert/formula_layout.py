# -*- coding: utf-8 -*-
from __future__ import annotations

import re
from typing import Any

from openpyxl.worksheet.worksheet import Worksheet

from bill_convert.formula_copy import (
    copy_row_formulas,
    fix_ee_row_tw_refs,
    fix_tw_row_tw_ee_refs,
)
from bill_convert.mapping_spec import mapping_section
from bill_convert.person import norm_person_name


def _formula_templates_spec(mapping: dict[str, Any]) -> dict[str, Any]:
    raw = mapping.get("formulaTemplates")
    return raw if isinstance(raw, dict) else {}


def _formula_sheet_block(mapping: dict[str, Any], sheet_key: str) -> dict[str, Any]:
    block = _formula_templates_spec(mapping).get(sheet_key)
    return block if isinstance(block, dict) else {}


def _formula_detect_strategy(block: dict[str, Any]) -> str:
    if block.get("fixedLayout") is True:
        return "fixed"
    if block.get("autoDetectLayout") is False:
        return "fixed"
    raw = str(block.get("detectStrategy") or "alignTwL").strip().lower()
    if raw in ("fixed", "scan"):
        return raw
    return "alignTwL"


def _scan_first_row_with_pattern(ws: Worksheet, pattern: re.Pattern[str], *, max_row: int = 35) -> int | None:
    for row in range(1, max_row + 1):
        for col in range(1, (ws.max_column or 0) + 1):
            cell = ws.cell(row, col)
            if cell.data_type == "f" and isinstance(cell.value, str) and pattern.search(cell.value):
                return row
    return None


def _default_example_row(mapping: dict[str, Any], sheet_key: str, data_start_fallback: int) -> int:
    block = _formula_sheet_block(mapping, sheet_key)
    if block.get("defaultExampleRow") is not None:
        return int(block["defaultExampleRow"])
    return data_start_fallback


def tw_l_row_for_data_row(
    data_row: int,
    *,
    data_start: int,
    target_l_data_start: int,
) -> int:
    return target_l_data_start + (data_row - data_start)


def resolve_formula_rows_layout(
    wb,
    mapping: dict[str, Any],
    target_l_data_start: int,
    *,
    tw_sheet: str = "TW",
    tw_ee_sheet: str = "TW EE",
    fallback_tw_data_start: int = 9,
    fallback_tw_ee_data_start: int = 10,
    target_l_sheet: str = "TW-L",
) -> dict[str, int]:
    tw_block = _formula_sheet_block(mapping, "TW")
    ee_block = _formula_sheet_block(mapping, "TW EE")
    tw_strategy = _formula_detect_strategy(tw_block)
    ee_strategy = _formula_detect_strategy(ee_block)

    if tw_strategy == "fixed":
        tw_start = int(tw_block.get("dataStartRow") or fallback_tw_data_start)
    elif tw_strategy == "scan" and tw_sheet in wb.sheetnames:
        pat = re.compile(rf"'{re.escape(target_l_sheet)}'!\$?[A-Z]{{1,3}}\$?\d+", re.I)
        tw_start = _scan_first_row_with_pattern(wb[tw_sheet], pat) or target_l_data_start
    else:
        tw_start = target_l_data_start

    ee_offset = int(ee_block.get("dataStartOffset") or 1)
    if ee_strategy == "fixed":
        ee_start = int(ee_block.get("dataStartRow") or fallback_tw_ee_data_start)
    elif ee_strategy == "scan" and tw_ee_sheet in wb.sheetnames:
        pat = re.compile(r"TW!\$?[A-Z]{1,3}\$?\d+", re.I)
        ee_start = _scan_first_row_with_pattern(wb[tw_ee_sheet], pat) or (tw_start + ee_offset)
    else:
        ee_start = tw_start + ee_offset

    return {
        "tw_l_data_start": target_l_data_start,
        "tw_data_start": tw_start,
        "tw_ee_data_start": ee_start,
    }


def _employee_formula_styles(mapping: dict[str, Any]) -> list[dict[str, Any]]:
    raw = mapping.get("employeeFormulaStyles")
    if not isinstance(raw, list):
        return []
    return [x for x in raw if isinstance(x, dict)]


def _style_entry_matches_employee(entry: dict[str, Any], emp: dict[str, Any]) -> bool:
    cn = norm_person_name(emp.get("CN Name"))
    en = norm_person_name(emp.get("EN Name"))
    e_cn = norm_person_name(entry.get("cnName"))
    e_en = norm_person_name(entry.get("enName"))
    if e_cn and e_en:
        return e_cn == cn and e_en == en
    if e_cn:
        return e_cn == cn
    if e_en:
        return e_en == en
    return False


def _style_for_employee(mapping: dict[str, Any], emp: dict[str, Any]) -> dict[str, Any]:
    for entry in _employee_formula_styles(mapping):
        if _style_entry_matches_employee(entry, emp):
            return entry
    return {}


def _apply_default_formula_template_to_all(mapping: dict[str, Any]) -> bool:
    spec = _formula_templates_spec(mapping)
    if "applyDefaultToAllEmployees" in spec:
        return bool(spec["applyDefaultToAllEmployees"])
    return True


def apply_employee_formula_styles(
    wb,
    employees: list[dict[str, Any]],
    mapping: dict[str, Any],
    *,
    formula_rows: dict[str, int],
    tw_sheet: str = "TW",
    tw_ee_sheet: str = "TW EE",
    target_l_sheet: str = "TW-L",
) -> None:
    apply_all = _apply_default_formula_template_to_all(mapping)
    styles = _employee_formula_styles(mapping)
    if not apply_all and not styles:
        return
    tw_l_data_start_row = formula_rows["tw_l_data_start"]
    tw_data_start = formula_rows["tw_data_start"]
    tw_ee_data_start = formula_rows["tw_ee_data_start"]
    tw = wb[tw_sheet]
    ee = wb[tw_ee_sheet]
    tw_def = _default_example_row(mapping, "TW", tw_data_start)
    ee_def = _default_example_row(mapping, "TW EE", tw_ee_data_start)

    for i, emp in enumerate(employees):
        entry = _style_for_employee(mapping, emp)
        dst_tw = tw_data_start + i
        dst_ee = tw_ee_data_start + i

        has_override = bool(entry.get("twExampleRow") or entry.get("twEeExampleRow"))
        if not apply_all and not has_override:
            continue

        if has_override:
            src_tw = int(entry["twExampleRow"]) if entry.get("twExampleRow") else tw_def
            src_ee = int(entry["twEeExampleRow"]) if entry.get("twEeExampleRow") else ee_def
        else:
            src_tw = tw_def
            src_ee = ee_def

        src_tw_l = tw_l_row_for_data_row(
            src_tw, data_start=tw_data_start, target_l_data_start=tw_l_data_start_row
        )
        dst_tw_l = tw_l_data_start_row + i
        copy_row_formulas(
            tw, src_tw, dst_tw, src_tw_l, dst_tw_l, target_l_sheet=target_l_sheet
        )
        fix_tw_row_tw_ee_refs(tw, dst_tw, dst_ee)
        src_ee_l = tw_l_row_for_data_row(
            src_ee, data_start=tw_ee_data_start, target_l_data_start=tw_l_data_start_row
        )
        copy_row_formulas(
            ee, src_ee, dst_ee, src_ee_l, dst_tw_l, target_l_sheet=target_l_sheet
        )
        fix_ee_row_tw_refs(ee, dst_ee, dst_tw)
