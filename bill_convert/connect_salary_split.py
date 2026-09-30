# -*- coding: utf-8 -*-
"""Connect UAE：mapping.connectSalarySplit → cellProvenance（蓝色 ⓘ）。"""
from __future__ import annotations

from typing import Any

# mapping 拆分键 → UAE-L 表头
SPLIT_TO_HEADER: dict[str, str] = {
    "basic": "Basic Salary",
    "housing": "Housing Allowance",
    "transport": "Transport",
}

HEADER_TO_SPLIT: dict[str, str] = {v: k for k, v in SPLIT_TO_HEADER.items()}


def build_connect_salary_split_cell_writes(
    employees: list[dict[str, Any]],
    *,
    sheet: str,
    data_start: int,
    header_map: dict[str, int],
    mapping: dict[str, Any] | None,
) -> list[dict[str, Any]]:
    """为配置了 connectSalarySplit 的员工写出 Basic/Housing/Transport 蓝标坐标。"""
    if not employees or not isinstance(mapping, dict):
        return []
    splits = mapping.get("connectSalarySplit")
    if not isinstance(splits, dict) or not splits:
        return []
    field_cols = {
        header: header_map[header]
        for header in SPLIT_TO_HEADER.values()
        if header in header_map
    }
    if not field_cols:
        return []

    # 复用 Connect PDF 侧的姓名匹配（含宽松包含）
    from pdf_ingest.profiles.connect_uae import _match_split

    cells: list[dict[str, Any]] = []
    for idx, emp in enumerate(employees):
        name = str(emp.get("English Name") or emp.get("Employee Name") or "").strip()
        split = _match_split(splits, name)
        if not split:
            continue
        row = data_start + idx
        split_summary: dict[str, float] = {}
        for key, header in SPLIT_TO_HEADER.items():
            val = emp.get(header)
            if isinstance(val, (int, float)) and not isinstance(val, bool):
                split_summary[key] = round(float(val), 2)
        for header, col in field_cols.items():
            if header not in emp or emp.get(header) is None:
                continue
            val = emp.get(header)
            split_key = HEADER_TO_SPLIT.get(header)
            detail: dict[str, Any] = {
                "employeeName": name,
                "field": header,
                "fieldLabel": header,
            }
            if split_key:
                detail["splitKey"] = split_key
            if split_summary:
                detail["splitSummary"] = dict(split_summary)
            cells.append(
                {
                    "kind": "connectSalarySplit",
                    "sheet": sheet,
                    "row": row,
                    "col": col,
                    "sourceType": "mapping",
                    "source": "mapping.connectSalarySplit",
                    "label": header,
                    "value": val,
                    "detail": detail,
                }
            )
    return cells
