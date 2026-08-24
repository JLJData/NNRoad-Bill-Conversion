# -*- coding: utf-8 -*-
"""Pakistan 映射：按员工 E.O.B.I / IT（如 Danfoss pakistanEmployeeFees）。"""
from __future__ import annotations

from typing import Any

from bill_convert.person import compact_person_name, score_person_name_match


def _as_float(value: Any) -> float | None:
    if value is None or value == "":
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def parse_pakistan_employee_fees(mapping: dict[str, Any] | None) -> dict[str, dict[str, float]]:
    """mapping.pakistanEmployeeFees: { employeeName: { eobi, it } }"""
    if not isinstance(mapping, dict):
        return {}
    raw = mapping.get("pakistanEmployeeFees")
    if not isinstance(raw, dict):
        return {}
    out: dict[str, dict[str, float]] = {}
    for name, entry in raw.items():
        key = str(name or "").strip()
        if not key or not isinstance(entry, dict):
            continue
        fee: dict[str, float] = {}
        eobi = _as_float(entry.get("eobi") if "eobi" in entry else entry.get("E.O.B.I"))
        it = _as_float(entry.get("it") if "it" in entry else entry.get("IT"))
        if eobi is not None:
            fee["eobi"] = eobi
        if it is not None:
            fee["it"] = it
        if fee:
            out[key] = fee
    return out


def lookup_employee_fee(name: str, fees: dict[str, dict[str, float]]) -> dict[str, float] | None:
    if not name or not fees:
        return None
    if name in fees:
        return fees[name]
    compact = compact_person_name(name)
    for k, v in fees.items():
        if compact_person_name(k) == compact:
            return v
    best: dict[str, float] | None = None
    best_score = 0
    for k, v in fees.items():
        score = score_person_name_match(name, k)
        if score > best_score:
            best_score = score
            best = v
    return best if best_score >= 70 else None


def make_pakistan_employee_fee_provenance(
    *,
    sheet: str,
    row: int,
    col: int,
    field: str,
    value: float,
    employee_name: str,
) -> dict[str, Any]:
    """单格 E.O.B.I / IT provenance（Excel 1-based）。"""
    return {
        "kind": "pakistanEmployeeFees",
        "sheet": sheet,
        "row": row,
        "col": col,
        "sourceType": "mapping",
        "source": "mapping.pakistanEmployeeFees",
        "label": field,
        "value": value,
        "detail": {
            "employeeName": employee_name,
            "field": field,
        },
    }


def apply_fees_to_employee_rows(
    employees: list[dict[str, Any]],
    fees: dict[str, dict[str, float]],
) -> list[str]:
    """把映射费用写入员工行字典的 E.O.B.I / IT；返回未命中姓名警告。"""
    if not fees:
        return []
    warnings: list[str] = []
    seen: set[str] = set()
    for emp in employees:
        name = str(emp.get("Name of Employee") or emp.get("Employee Name") or "").strip()
        if not name:
            continue
        fee = lookup_employee_fee(name, fees)
        if not fee:
            key = compact_person_name(name) or name
            if key not in seen:
                seen.add(key)
                warnings.append(f"{name}：映射 pakistanEmployeeFees 未匹配到 E.O.B.I / IT")
            continue
        if "eobi" in fee:
            emp["E.O.B.I"] = fee["eobi"]
        if "it" in fee:
            emp["IT"] = fee["it"]
    return warnings
