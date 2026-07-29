# -*- coding: utf-8 -*-
"""Office 下发的 convert_mapping JSON 与引擎默认值合并。"""
from __future__ import annotations

import copy
from typing import Any


def _deep_merge(base: dict[str, Any], override: dict[str, Any]) -> dict[str, Any]:
    out = copy.deepcopy(base)
    for k, v in override.items():
        if k in out and isinstance(out[k], dict) and isinstance(v, dict):
            out[k] = _deep_merge(out[k], v)
        else:
            out[k] = copy.deepcopy(v)
    return out


ENGINE_DEFAULTS: dict[str, dict[str, Any]] = {
    "tw_payroll_calc": {
        "schemaVersion": 1,
        "sourceEmployeeSheet": {
            "sheet": "Payroll calculation",
            "candidates": ["Payroll calculation", "Payroll Calculation"],
            "headerScanMaxRow": 15,
            "headerMarkerKeys": ["BU", "CN Name", "EN Name"],
            "nameHeaders": ["CN Name", "EN Name"],
            "payrollDupMinCol": 14,
        },
        "sourceMetaSheet": {"sheet": "Summary", "optional": True},
        "targetL": {
            "sheet": "TW-L",
            "autoDetectLayout": True,
            "dataStartOffset": 2,
            "headerRow": 7,
            "dataStartRow": 9,
        },
        "columnRename": {
            "時薪 Hourly Rate": "Full Pay/Hourly Rate",
            "時數\nHours Worked": "Employment day/Hours Worked",
            "健保級距\nInsured Salary Grading - HI": "健保投保級距Insured Salary Grading - HI",
        },
        "formulaTemplates": {
            "applyDefaultToAllEmployees": True,
            "TW": {"detectStrategy": "alignTwL", "defaultExampleRow": 9},
            "TW EE": {"detectStrategy": "alignTwL", "dataStartOffset": 1, "defaultExampleRow": 10},
        },
        "employeeFormulaStyles": [],
        "skipSourceHeaders": [],
    },
    "china_payroll_calc": {
        "schemaVersion": 1,
        "sourceEmployeeSheet": {
            "sheet": "计算结果",
            "candidates": ["计算结果"],
            "headerRow": 1,
            "nameHeaders": ["姓名"],
        },
        "targetL": {"sheet": "China-L", "headerRow": 1, "dataStartRow": 2},
        "columnRename": {},
        "formulaTemplates": {
            "applyDefaultToAllEmployees": True,
            "China": {"defaultExampleRow": 9},
            "China EE": {"defaultExampleRow": 10, "dataStartOffset": 1},
        },
        "employeeFormulaStyles": [],
        "skipSourceHeaders": [],
    },
    "hk_vertical_l": {
        "schemaVersion": 1,
        "sourceEmployeeSheet": {
            "sheet": "Hong Kong-L",
            "candidates": ["Hong Kong-L"],
            "headerRow": 7,
            "dataStartRow": 8,
            "nameHeaders": ["Name of Employee", "EE Name", "Name"],
        },
        "targetL": {"sheet": "Hong Kong-L", "headerRow": 7, "dataStartRow": 8},
    },
}


def resolve_convert_mapping(engine_id: str, raw: dict[str, Any] | None) -> dict[str, Any]:
    # 旧引擎 id 兼容
    if engine_id == "china_hrone":
        engine_id = "china_payroll_calc"
    base = copy.deepcopy(ENGINE_DEFAULTS.get(engine_id, {"schemaVersion": 1}))
    if not raw:
        return base
    if not isinstance(raw, dict):
        return base
    override = copy.deepcopy(raw)
    for whole_key in ("columnRename", "skipSourceHeaders"):
        if whole_key in override:
            base[whole_key] = copy.deepcopy(override.pop(whole_key))
    return _deep_merge(base, override)


def find_sheet_name(sheetnames: list[str], spec: dict[str, Any] | None) -> str | None:
    if not spec:
        return None
    primary = str(spec.get("sheet") or "").strip()
    candidates = spec.get("candidates") or []
    names = list(sheetnames)
    lower_map = {n.lower(): n for n in names}

    def match_one(want: str) -> str | None:
        w = want.strip()
        if not w:
            return None
        if w in names:
            return w
        return lower_map.get(w.lower())

    if primary:
        hit = match_one(primary)
        if hit:
            return hit
    if isinstance(candidates, list):
        for c in candidates:
            hit = match_one(str(c))
            if hit:
                return hit
    return None
