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
        "columnRename": {},
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
    "hk_payroll_calc": {
        "schemaVersion": 1,
        "sourceEmployeeSheet": {
            "sheet": "Hong Kong-L",
            "candidates": ["Hong Kong-L"],
            "headerRow": 7,
            "dataStartRow": 8,
            "nameHeaders": ["Name of Employee", "EE Name", "Name"],
        },
        "targetL": {"sheet": "Hong Kong-L", "headerRow": 7, "dataStartRow": 8},
        "columnRename": {},
        "formulaTemplates": {
            "applyDefaultToAllEmployees": True,
            "Hong Kong": {"defaultExampleRow": 9},
            "Hong Kong EE": {"defaultExampleRow": 10, "dataStartOffset": 1},
        },
        "employeeFormulaStyles": [],
        "skipSourceHeaders": [],
    },
    "uk_payroll_calc": {
        # 中性默认；TopSource / EOR 列名别名见 PROFILE_MAPPING_OVERLAYS
        "schemaVersion": 1,
        "sourceEmployeeSheet": {
            "sheet": "UK-L",
            "candidates": ["UK-L", "UK-L (2)", "UK-L (3)"],
            "layout": "vertical_label_amount",
            "labelColumn": 1,
            "amountColumn": 2,
            "nameHeaders": ["Employee Name"],
            "nameLabel": "Employee Name",
        },
        "targetL": {
            "sheet": "UK-L",
            "candidates": ["UK-L"],
            "layout": "vertical_label_amount",
            "labelColumn": 1,
            "amountColumn": 2,
        },
        "columnRename": {},
        "formulaTemplates": {
            "applyDefaultToAllEmployees": False,
            "UK": {"defaultExampleRow": 9},
            "UK EE": {"defaultExampleRow": 9, "dataStartOffset": 0},
        },
        "employeeFormulaStyles": [],
        "skipSourceHeaders": [],
        "pnSheets": {"main": "UK", "ee": "UK EE", "l": "UK-L"},
    },
    # UAE 引擎默认保持中性；供应商差异见 PROFILE_MAPPING_OVERLAYS
    "uae_payroll_calc": {
        "schemaVersion": 1,
        "sourceEmployeeSheet": {
            "sheet": "UAE-L",
            "candidates": ["UAE-L"],
            "headerRow": 2,
            "dataStartRow": 3,
            "nameHeaders": ["Employee Name", "English Name"],
        },
        "targetL": {
            "sheet": "UAE-L",
            "candidates": ["UAE-L"],
            "headerRow": 2,
            "dataStartRow": 3,
        },
        "columnRename": {},
        "formulaTemplates": {
            "applyDefaultToAllEmployees": True,
            "UAE": {"defaultExampleRow": 9},
            "UAE EE": {"defaultExampleRow": 10, "dataStartOffset": 1},
        },
        "employeeFormulaStyles": [],
        "skipSourceHeaders": [],
        "pnSheets": {"main": "UAE", "ee": "UAE EE", "l": "UAE-L"},
        # 固定值写格：唯一约定。转换引擎与 Office「同步母版」都读这里，勿在 Java 再写死格子。
        "fixedValueWrites": [
            {
                "id": "uaeRecurringFee",
                "valueKey": "uaeRecurringFeeFixed",
                "sheet": "UAE",
                "columnLetter": "H",
                "dataStartRow": 9,
                "scope": "eachEmployee",
            }
        ],
    },
    # Pakistan 中性默认；Panda Work 见 PROFILE_MAPPING_OVERLAYS
    "pakistan_payroll_calc": {
        "schemaVersion": 1,
        "sourceEmployeeSheet": {
            "sheet": "Pakistan-L",
            "candidates": ["Pakistan-L"],
            "headerRow": 7,
            "dataStartRow": 8,
            "nameHeaders": ["Name of Employee", "Employee Name"],
        },
        "targetL": {
            "sheet": "Pakistan-L",
            "candidates": ["Pakistan-L"],
            "headerRow": 7,
            "dataStartRow": 8,
        },
        "columnRename": {},
        "formulaTemplates": {
            "applyDefaultToAllEmployees": True,
            "Pakistan": {"defaultExampleRow": 9},
            "Pakistan EE": {"defaultExampleRow": 10, "dataStartOffset": 1},
        },
        "employeeFormulaStyles": [],
        "skipSourceHeaders": [],
        "pnSheets": {"main": "Pakistan", "ee": "Pakistan EE", "l": "Pakistan-L"},
        "quarterSplitMonths": 3,
    },
    # Italy 中性默认；SafeGuard 见 PROFILE_MAPPING_OVERLAYS
    "italy_payroll_calc": {
        "schemaVersion": 1,
        "sourceEmployeeSheet": {
            "sheet": "Italy-L",
            "candidates": ["Italy-L"],
            "headerRow": 10,
            "dataStartRow": 11,
            "nameHeaders": ["Employee Name"],
        },
        "targetL": {
            "sheet": "Italy-L",
            "candidates": ["Italy-L"],
            "headerRow": 10,
            "dataStartRow": 11,
        },
        "columnRename": {},
        "formulaTemplates": {
            "applyDefaultToAllEmployees": True,
            "Italy": {"defaultExampleRow": 9},
            "Italy EE": {"defaultExampleRow": 10, "dataStartOffset": 1},
        },
        "employeeFormulaStyles": [],
        "skipSourceHeaders": [],
        "pnSheets": {"main": "Italy", "ee": "Italy EE", "l": "Italy-L"},
        "fixedValueWrites": [
            {
                "id": "italyFeeMin",
                "valueKey": "italyFeeMin",
                "sheet": "Italy-L",
                "headerNames": ["Fee Min", "SGWI Min"],
                "headerRow": 10,
                "dataStartRow": 11,
                "scope": "eachEmployee",
            }
        ],
    },
    # India 中性默认；Biz Solutions 见 PROFILE_MAPPING_OVERLAYS
    "india_payroll_calc": {
        "schemaVersion": 1,
        "sourceEmployeeSheet": {
            "sheet": "India-L",
            "candidates": ["India-L"],
            "headerRow": 4,
            "dataStartRow": 10,
            "nameHeaders": ["Employee Name", "Name"],
            "nameColumn": 2,
        },
        "targetL": {
            "sheet": "India-L",
            "candidates": ["India-L"],
            "headerRow": 4,
            "dataStartRow": 10,
        },
        "columnRename": {},
        "formulaTemplates": {
            "applyDefaultToAllEmployees": True,
            "India": {"defaultExampleRow": 9},
            "India EE": {"defaultExampleRow": 9, "dataStartOffset": 0},
        },
        "employeeFormulaStyles": [],
        "skipSourceHeaders": [],
        "pnSheets": {"main": "India", "ee": "India EE", "l": "India-L"},
    },
    # Cyprus 中性默认；A&T Technical 见 PROFILE_MAPPING_OVERLAYS
    "cyprus_payroll_calc": {
        "schemaVersion": 1,
        "sourceEmployeeSheet": {
            "sheet": "Cyprus-L",
            "candidates": ["Cyprus-L"],
            "headerRow": 7,
            "dataStartRow": 8,
            "nameHeaders": ["Name of Employee", "Employee Name"],
            "nameColumn": 2,
        },
        "targetL": {
            "sheet": "Cyprus-L",
            "candidates": ["Cyprus-L"],
            "headerRow": 7,
            "dataStartRow": 8,
        },
        "columnRename": {},
        "formulaTemplates": {
            "applyDefaultToAllEmployees": True,
            "Cyprus": {"defaultExampleRow": 9},
            "Cyprus EE": {"defaultExampleRow": 10, "dataStartOffset": 0},
        },
        "employeeFormulaStyles": [],
        "skipSourceHeaders": [],
        "pnSheets": {"main": "Cyprus", "ee": "Cyprus EE", "l": "Cyprus-L"},
        "fixedValueWrites": [
            {
                "id": "cyprusRecurringFee",
                "valueKey": "cyprusRecurringFee",
                "sheet": "Cyprus",
                "columnLetter": "I",
                "dataStartRow": 9,
                "scope": "eachEmployee",
            }
        ],
    },
}

# 列名对照不再内置默认：须在 Office「转换映射」中配置并保存。
# （历史别名曾用于 UK / Auxilium，已移除以免覆盖用户清空的映射。）

PROFILE_MAPPING_OVERLAYS: dict[str, dict[str, Any]] = {
    "topsource_uk": {
        "sourceEmployeeSheet": {
            "sheet": "UK-L",
            "candidates": ["UK-L", "UK-L (2)", "UK-L (3)"],
            "layout": "vertical_label_amount",
            "labelColumn": 1,
            "amountColumn": 2,
            "nameHeaders": ["Employee Name"],
            "nameLabel": "Employee Name",
        },
        "targetL": {
            "sheet": "UK-L",
            "candidates": ["UK-L"],
            "layout": "vertical_label_amount",
            "labelColumn": 1,
            "amountColumn": 2,
        },
        "columnRename": {},
    },
    "eor_uk": {
        "sourceEmployeeSheet": {
            "sheet": "UK-L",
            "candidates": ["UK-L"],
            "layout": "vertical_label_amount",
            "labelColumn": 1,
            "amountColumn": 2,
            "nameHeaders": ["Employee Name"],
            "nameLabel": "Employee Name",
        },
        "targetL": {
            "sheet": "UK-L",
            "candidates": ["UK-L"],
            "layout": "vertical_label_amount",
            "labelColumn": 1,
            "amountColumn": 2,
        },
        "columnRename": {},
    },
    "auxilium_uae": {
        "sourceEmployeeSheet": {
            "sheet": "UAE-L",
            "candidates": ["UAE-L"],
            "headerRow": 2,
            "dataStartRow": 3,
            "nameHeaders": ["Employee Name"],
        },
        "targetL": {
            "sheet": "UAE-L",
            "candidates": ["UAE-L"],
            "headerRow": 2,
            "dataStartRow": 3,
        },
        "columnRename": {},
    },
    "connect_uae": {
        "sourceEmployeeSheet": {
            "sheet": "UAE-L",
            "candidates": ["UAE-L"],
            "headerRow": 7,
            "dataStartRow": 8,
            "nameHeaders": ["English Name", "Employee Name"],
        },
        "targetL": {
            "sheet": "UAE-L",
            "candidates": ["UAE-L"],
            "headerRow": 7,
            "dataStartRow": 8,
        },
        "columnRename": {},
        "connectSalarySplit": {},
    },
    "panda_work_pk": {
        "sourceEmployeeSheet": {
            "sheet": "Pakistan-L",
            "candidates": ["Pakistan-L"],
            "headerRow": 7,
            "dataStartRow": 8,
            "nameHeaders": ["Name of Employee", "Employee Name"],
        },
        "targetL": {
            "sheet": "Pakistan-L",
            "candidates": ["Pakistan-L"],
            "headerRow": 7,
            "dataStartRow": 8,
        },
        "columnRename": {},
        "quarterSplitMonths": 3,
    },
    "safeguard_italy": {
        "sourceEmployeeSheet": {
            "sheet": "Italy-L",
            "candidates": ["Italy-L"],
            "headerRow": 10,
            "dataStartRow": 11,
            "nameHeaders": ["Employee Name"],
        },
        "targetL": {
            "sheet": "Italy-L",
            "candidates": ["Italy-L"],
            "headerRow": 10,
            "dataStartRow": 11,
        },
    },
    "biz_solutions_india": {
        "sourceEmployeeSheet": {
            "sheet": "India-L",
            "candidates": ["India-L"],
            "headerRow": 4,
            "dataStartRow": 10,
            "nameHeaders": ["Employee Name", "Name"],
            "nameColumn": 2,
        },
        "targetL": {
            "sheet": "India-L",
            "candidates": ["India-L"],
            "headerRow": 4,
            "dataStartRow": 10,
        },
        "columnRename": {},
    },
    "at_technical_cyprus": {
        "sourceEmployeeSheet": {
            "sheet": "Cyprus-L",
            "candidates": ["Cyprus-L"],
            "headerRow": 7,
            "dataStartRow": 8,
            "nameHeaders": ["Name of Employee", "Employee Name"],
            "nameColumn": 2,
        },
        "targetL": {
            "sheet": "Cyprus-L",
            "candidates": ["Cyprus-L"],
            "headerRow": 7,
            "dataStartRow": 8,
        },
        "columnRename": {},
    },
}


def resolve_convert_mapping(engine_id: str, raw: dict[str, Any] | None) -> dict[str, Any]:
    # 旧引擎 id 兼容
    if engine_id == "china_hrone":
        engine_id = "china_payroll_calc"
    if engine_id == "hk_vertical_l":
        engine_id = "hk_payroll_calc"
    base = copy.deepcopy(ENGINE_DEFAULTS.get(engine_id, {"schemaVersion": 1}))
    if not raw:
        merged = base
    elif not isinstance(raw, dict):
        merged = base
    else:
        override = copy.deepcopy(raw)
        for whole_key in (
            "columnRename",
            "skipSourceHeaders",
            "connectSalarySplit",
            "indiaSalarySplit",
            "indiaSalarySplits",
        ):
            if whole_key in override:
                base[whole_key] = copy.deepcopy(override.pop(whole_key))
        merged = _deep_merge(base, override)

    pid = ""
    if isinstance(merged, dict):
        pid = str(merged.get("pdfProfileId") or merged.get("_pdfProfileId") or "").strip()
    if pid and pid in PROFILE_MAPPING_OVERLAYS:
        overlay = copy.deepcopy(PROFILE_MAPPING_OVERLAYS[pid])
        # 布局键以 profile 为准；薪资拆分等保留用户配置
        for key in ("sourceEmployeeSheet", "targetL"):
            if key in overlay:
                merged[key] = copy.deepcopy(overlay[key])
        # columnRename 不从 profile overlay 注入：只认配置里保存的映射
        if "connectSalarySplit" in overlay and "connectSalarySplit" not in (raw or {}):
            merged["connectSalarySplit"] = copy.deepcopy(overlay["connectSalarySplit"])
        if "quarterSplitMonths" in overlay and "quarterSplitMonths" not in (raw or {}):
            merged["quarterSplitMonths"] = copy.deepcopy(overlay["quarterSplitMonths"])
        merged["pdfProfileId"] = pid
    # 写格约定始终以引擎默认为准，避免 Office 存过的 mapping 把格子钉死。
    engine_writes = (ENGINE_DEFAULTS.get(engine_id) or {}).get("fixedValueWrites")
    if isinstance(merged, dict):
        if engine_writes is not None:
            merged["fixedValueWrites"] = copy.deepcopy(engine_writes)
        else:
            merged.pop("fixedValueWrites", None)
    return merged


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
