# -*- coding: utf-8 -*-
from __future__ import annotations

from openpyxl.worksheet.worksheet import Worksheet

from xlsx_convert_utils import norm


def build_header_map(ws: Worksheet, header_row: int) -> dict[str, int]:
    mapping: dict[str, int] = {}
    for col in range(1, (ws.max_column or 0) + 1):
        key = norm(ws.cell(header_row, col).value)
        if key and key not in mapping:
            mapping[key] = col
    return mapping


def build_header_cols(ws: Worksheet, header_row: int) -> dict[str, list[int]]:
    mapping: dict[str, list[int]] = {}
    for col in range(1, (ws.max_column or 0) + 1):
        key = norm(ws.cell(header_row, col).value)
        if key:
            mapping.setdefault(key, []).append(col)
    return mapping


def resolve_target_col(cols: list[int], *, dup_min_col: int = 14) -> int | None:
    if not cols:
        return None
    if len(cols) == 1:
        return cols[0]
    payroll = [c for c in cols if c >= dup_min_col]
    return max(payroll) if payroll else max(cols)
