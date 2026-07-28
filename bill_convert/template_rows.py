# -*- coding: utf-8 -*-
from __future__ import annotations

from typing import Any

from openpyxl.worksheet.worksheet import Worksheet


def clear_row_values(ws: Worksheet, row: int) -> None:
    for col in range(1, (ws.max_column or 0) + 1):
        ws.cell(row, col).value = None


def count_data_slots(ws: Worksheet, data_start_row: int, *, marker_col: int = 2, max_slots: int = 9) -> int:
    n = 0
    for row in range(data_start_row, data_start_row + max_slots):
        if ws.cell(row, marker_col).value is not None:
            n += 1
        else:
            break
    return max(n, 1)


def list_formula_example_rows(
    ws: Worksheet,
    data_start_row: int,
    *,
    marker_col: int = 2,
    max_slots: int = 9,
) -> list[dict[str, Any]]:
    out: list[dict[str, Any]] = []
    for offset in range(max_slots):
        row = data_start_row + offset
        marker = ws.cell(row, marker_col).value
        has_formula = False
        for col in range(1, (ws.max_column or 0) + 1):
            cell = ws.cell(row, col)
            if cell.data_type == "f" and cell.value:
                has_formula = True
                break
        if offset > 0 and marker is None and not has_formula:
            break
        if marker is None and not has_formula and offset > 0:
            continue
        label = str(marker).strip() if marker is not None else ""
        if not label:
            label = f"第 {row} 行"
        else:
            label = f"第 {row} 行 · {label[:48]}"
        out.append({"row": row, "label": label})
    if not out:
        out.append({"row": data_start_row, "label": f"第 {data_start_row} 行（默认）"})
    return out
