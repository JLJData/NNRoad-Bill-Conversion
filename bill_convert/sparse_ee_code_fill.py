# -*- coding: utf-8 -*-
"""母版稀疏行未定义 EE Code 列单元格时，在转换副本写值前仅补 fill（不改正本母版 xlsx）。"""
from __future__ import annotations

from copy import copy

from openpyxl.cell.cell import Cell
from openpyxl.worksheet.worksheet import Worksheet


def _has_visible_fill(cell: Cell) -> bool:
    fill = cell.fill
    if fill is None:
        return False
    fill_type = getattr(fill, "fill_type", None) or getattr(fill, "patternType", None)
    if not fill_type or str(fill_type).lower() == "none":
        return False
    return True


def ensure_ee_code_fill_from_peer_row(
    ws: Worksheet,
    row: int,
    code_col: int,
    peer_col: int,
) -> None:
    """同行 peer 列（如 EE Name / 公式列）有底色而 Code 列无时，只复制 fill。"""
    if code_col == peer_col:
        return
    dst = ws.cell(row, code_col)
    if _has_visible_fill(dst):
        return
    src = ws.cell(row, peer_col)
    if _has_visible_fill(src):
        dst.fill = copy(src.fill)
