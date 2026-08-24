# -*- coding: utf-8 -*-
"""India 映射：按员工维护薪资拆分 + PT/IIT（indiaSalarySplit）。

薪资：Basic / HRA / Telephone / LTA / Special / Wellness Stipend（合计应对齐 PDF CTC）
另两项：Professional tax / IIT（扣款项，不计入 CTC 校验）
Bonus 不在映射维护，写入时置 0。
"""
from __future__ import annotations

from typing import Any

from bill_convert.person import compact_person_name, score_person_name_match

# 计入 CTC 校验的薪资列
SALARY_KEYS = (
    "basic",
    "hra",
    "telephone",
    "lta",
    "special",
    "wellness",
)

# 扣款列（不计入 CTC）
DEDUCTION_KEYS = (
    "professionalTax",
    "iit",
)

SPLIT_KEYS = SALARY_KEYS + DEDUCTION_KEYS

# mapping 键 → India-L 员工 dict 字段名
SPLIT_TO_EMP_FIELD: dict[str, str] = {
    "basic": "Basic salary",
    "hra": "HRA",
    "telephone": "Telephone allowance",
    "lta": "LTA",
    "special": "Special allowance",
    "wellness": "Wellness Stipend",
    "professionalTax": "Professional tax",
    "iit": "IIT",
}

EMP_FIELD_TO_SPLIT: dict[str, str] = {v: k for k, v in SPLIT_TO_EMP_FIELD.items()}

# indiaSalarySplit provenance / 写入：India-L 表头名（按表头行动态解析列号）
INDIA_SPLIT_PROVENANCE_FIELDS: tuple[str, ...] = tuple(SPLIT_TO_EMP_FIELD.values()) + ("Bonus",)

# write_india_l 可写字段（表头名须与母版一致）
INDIA_L_KNOWN_DATA_FIELDS: tuple[str, ...] = (
    "Employee Name",
    "Business Tax",
    *INDIA_SPLIT_PROVENANCE_FIELDS,
    "Expense Claim",
    "Deduction",
)

_ALIAS: dict[str, str] = {
    "basic": "basic",
    "basic salary": "basic",
    "basicsalary": "basic",
    "hra": "hra",
    "telephone": "telephone",
    "telephone allowance": "telephone",
    "lta": "lta",
    "special": "special",
    "special allowance": "special",
    "wellness": "wellness",
    "wellness stipend": "wellness",
    "professional tax": "professionalTax",
    "professionaltax": "professionalTax",
    "pt": "professionalTax",
    "iit": "iit",
}

_RAW_KEYS = ("indiaSalarySplit", "indiaSalarySplits", "india_salary_split")


def _as_float(value: Any) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    text = str(value).strip().replace(",", "").replace("\xa0", "").replace("，", "")
    if not text:
        return None
    try:
        return float(text)
    except (TypeError, ValueError):
        return None


def _parse_entry(entry: Any) -> dict[str, float]:
    if not isinstance(entry, dict):
        return {}
    split: dict[str, float] = {}
    for k, v in entry.items():
        canon = _ALIAS.get(str(k or "").strip().lower().replace("_", " "))
        if not canon:
            continue
        num = _as_float(v)
        if num is not None:
            split[canon] = round(num, 2)
    return split


def parse_india_salary_splits(mapping: dict[str, Any] | None) -> dict[str, dict[str, float]]:
    """mapping.indiaSalarySplit: { employeeName: { basic, hra, ..., wellness, professionalTax, iit } }"""
    if not isinstance(mapping, dict):
        return {}
    raw = None
    for key in _RAW_KEYS:
        if key in mapping:
            raw = mapping.get(key)
            break
    out: dict[str, dict[str, float]] = {}
    if isinstance(raw, dict):
        for name, entry in raw.items():
            key = str(name or "").strip()
            if not key:
                continue
            split = _parse_entry(entry)
            if split:
                out[key] = split
    elif isinstance(raw, list):
        for item in raw:
            if not isinstance(item, dict):
                continue
            key = str(item.get("name") or item.get("employeeName") or item.get("Employee Name") or "").strip()
            if not key:
                continue
            split = _parse_entry(item)
            if split:
                out[key] = split
    return out


def lookup_salary_split(name: str, splits: dict[str, dict[str, float]]) -> dict[str, float] | None:
    if not name or not splits:
        return None
    if name in splits:
        return splits[name]
    compact = compact_person_name(name)
    for k, v in splits.items():
        if compact_person_name(k) == compact:
            return v
    best: dict[str, float] | None = None
    best_score = 0
    for k, v in splits.items():
        score = score_person_name_match(name, k)
        if score > best_score:
            best_score = score
            best = v
    return best if best_score >= 70 else None


def apply_salary_split_to_employee(
    emp: dict[str, Any],
    *,
    mapping: dict[str, Any] | None,
    ctc: float | None = None,
    warnings: list[str] | None = None,
    fallback_ctc_to_basic: bool = True,
) -> dict[str, Any]:
    """
    有映射则写入薪资六项 + PT/IIT；Bonus 置 0。
    无匹配时：若 fallback_ctc_to_basic，则 CTC 整笔进 Basic。
    """
    warn = warnings if warnings is not None else []
    name = str(emp.get("Employee Name") or "").strip()
    splits = parse_india_salary_splits(mapping)
    split = lookup_salary_split(name, splits)

    if split:
        emp["_india_salary_split_from_mapping"] = True
        for key, field in SPLIT_TO_EMP_FIELD.items():
            emp[field] = round(float(split.get(key) or 0.0), 2)
        emp["Bonus"] = 0.0

        salary_sum = round(sum(float(split.get(k) or 0.0) for k in SALARY_KEYS), 2)
        pdf_ctc = _as_float(ctc if ctc is not None else emp.get("_ctc"))
        if pdf_ctc is not None and abs(salary_sum - round(pdf_ctc, 2)) > 0.05:
            warn.append(
                f"{name}：映射薪资合计 {salary_sum} ≠ PDF CTC {round(pdf_ctc, 2)}"
            )
        return emp

    configured = ", ".join(splits.keys()) if splits else ""
    if splits:
        warn.append(
            f"{name or '员工'}：映射已有 indiaSalarySplit，但未匹配到该姓名"
            f"（已配置: {configured}）"
        )
        if not fallback_ctc_to_basic:
            return emp
    elif not fallback_ctc_to_basic:
        return emp
    else:
        raw_present = False
        if isinstance(mapping, dict):
            for key in _RAW_KEYS:
                if key in mapping and mapping.get(key) not in (None, "", {}, []):
                    raw_present = True
                    break
        if raw_present:
            warn.append(
                f"{name or '员工'}：indiaSalarySplit 有内容但无法解析出有效金额"
                f"（请确认键名为 basic/hra/telephone/lta/special/wellness/professionalTax/iit）"
            )
        else:
            warn.append(
                f"{name or '员工'}：映射未配置 indiaSalarySplit"
                f"（请在 Office「转换映射 → India · 薪资拆分」填写并点保存映射）"
            )

    total = _as_float(ctc if ctc is not None else emp.get("Basic salary") or emp.get("_ctc")) or 0.0
    emp["Basic salary"] = round(total, 2)
    for field in (
        "HRA",
        "Telephone allowance",
        "LTA",
        "Special allowance",
        "Bonus",
        "Wellness Stipend",
    ):
        emp[field] = 0.0
    if "Professional tax" not in emp:
        emp["Professional tax"] = 0.0
    if "IIT" not in emp:
        emp["IIT"] = 0.0
    if total:
        warn[-1] = warn[-1] + f"；已将 CTC {total} 全部计入 Basic salary"
    return emp


def resolve_field_columns_from_headers(
    header_map: dict[str, int],
    fields: tuple[str, ...] | list[str] | None = None,
) -> dict[str, int]:
    """表头名 → Excel 列号（1-based）；仅返回 header_map 中存在的字段。"""
    want = fields if fields is not None else INDIA_L_KNOWN_DATA_FIELDS
    return {name: header_map[name] for name in want if name in header_map}


def build_india_salary_split_cell_writes(
    employees: list[dict[str, Any]],
    *,
    sheet: str,
    data_start: int,
    field_cols: dict[str, int],
) -> list[dict[str, Any]]:
    """为映射 indiaSalarySplit 成功写入的员工行生成 cellProvenance。

    field_cols：由 resolve_field_columns_from_headers 从 India-L 表头解析，勿写死列号。
    """
    cells: list[dict[str, Any]] = []
    if not field_cols:
        return cells
    for idx, emp in enumerate(employees):
        if not emp.get("_india_salary_split_from_mapping"):
            continue
        row = data_start + idx
        name = str(emp.get("Employee Name") or "").strip()
        split_summary: dict[str, float] = {}
        for field in INDIA_SPLIT_PROVENANCE_FIELDS:
            val = emp.get(field)
            if val is None:
                continue
            if isinstance(val, (int, float)) and not isinstance(val, bool):
                split_key = EMP_FIELD_TO_SPLIT.get(field, field.lower())
                split_summary[split_key] = round(float(val), 2)
        for field, col in field_cols.items():
            if field not in INDIA_SPLIT_PROVENANCE_FIELDS:
                continue
            if field not in emp:
                continue
            val = emp.get(field)
            if val is None:
                continue
            split_key = EMP_FIELD_TO_SPLIT.get(field)
            detail: dict[str, Any] = {
                "employeeName": name,
                "field": field,
                "fieldLabel": field,
            }
            if split_key:
                detail["splitKey"] = split_key
            if split_summary:
                detail["splitSummary"] = dict(split_summary)
            cells.append(
                {
                    "kind": "indiaSalarySplit",
                    "sheet": sheet,
                    "row": row,
                    "col": col,
                    "sourceType": "mapping",
                    "source": "mapping.indiaSalarySplit",
                    "label": field,
                    "value": val,
                    "detail": detail,
                }
            )
    return cells
