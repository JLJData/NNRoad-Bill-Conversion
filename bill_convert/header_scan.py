# -*- coding: utf-8 -*-
from __future__ import annotations

from openpyxl.worksheet.worksheet import Worksheet

from xlsx_convert_utils import norm


def find_header_row_by_markers(
    ws: Worksheet,
    *,
    marker_keys: frozenset[str],
    max_scan: int = 15,
    min_hits: int = 2,
    sheet_label: str = "",
) -> int:
    label = sheet_label or "工作表"
    for row in range(1, max_scan + 1):
        hits = 0
        for col in range(1, (ws.max_column or 0) + 1):
            if norm(ws.cell(row, col).value) in marker_keys:
                hits += 1
        if hits >= min_hits:
            return row
    raise ValueError(
        f"「{label}」前 {max_scan} 行未找到表头行（需含 {', '.join(sorted(marker_keys))} 中至少 {min_hits} 项）"
    )
