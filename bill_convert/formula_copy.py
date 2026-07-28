# -*- coding: utf-8 -*-
from __future__ import annotations

import re
from copy import copy

from openpyxl.worksheet.worksheet import Worksheet


def shift_row_formula(
    formula: str,
    from_row: int,
    to_row: int,
    *,
    target_l_from: int,
    target_l_to: int,
    target_l_sheet: str = "TW-L",
) -> str:
    if not formula or not isinstance(formula, str) or not formula.startswith("="):
        return formula

    placeholders: list[str] = []

    def stash_external(match: re.Match[str]) -> str:
        placeholders.append(match.group(0))
        return f"__EXT{len(placeholders) - 1}__"

    s = re.sub(r"'[^']+'!\$?[A-Z]{1,3}\$?\d+", stash_external, formula)
    s = re.sub(
        rf"(?<!\$)(?<![A-Z])([A-Z]{{1,3}}){from_row}(?!\d)",
        lambda m: f"{m.group(1)}{to_row}",
        s,
    )
    pat = re.compile(rf"'{re.escape(target_l_sheet)}'!([A-Z]{{1,3}})(\d+)")
    for idx, ref in enumerate(placeholders):
        ref = pat.sub(lambda m: f"'{target_l_sheet}'!{m.group(1)}{target_l_to}", ref)
        s = s.replace(f"__EXT{idx}__", ref)
    return s


def copy_row_formulas(
    ws: Worksheet,
    from_row: int,
    to_row: int,
    target_l_from: int,
    target_l_to: int,
    *,
    target_l_sheet: str = "TW-L",
) -> None:
    max_col = ws.max_column or 1
    for col in range(1, max_col + 1):
        src = ws.cell(from_row, col)
        dst = ws.cell(to_row, col)
        if src.has_style:
            dst.font = copy(src.font)
            dst.border = copy(src.border)
            dst.fill = copy(src.fill)
            dst.number_format = src.number_format
            dst.protection = copy(src.protection)
            dst.alignment = copy(src.alignment)
        if src.data_type == "f" and isinstance(src.value, str):
            dst.value = shift_row_formula(
                src.value,
                from_row,
                to_row,
                target_l_from=target_l_from,
                target_l_to=target_l_to,
                target_l_sheet=target_l_sheet,
            )
        elif src.value is not None and src.data_type != "f":
            dst.value = copy(src.value)


def fix_tw_row_tw_ee_refs(ws_tw: Worksheet, dst_row: int, ee_row: int) -> None:
    for col in range(1, (ws_tw.max_column or 0) + 1):
        cell = ws_tw.cell(dst_row, col)
        if cell.data_type == "f" and isinstance(cell.value, str):
            cell.value = re.sub(
                rf"'TW EE'!([A-Z]+){ee_row - 1}(?!\d)",
                lambda m: f"'TW EE'!{m.group(1)}{ee_row}",
                cell.value,
            )


def fix_ee_row_tw_refs(ws_ee: Worksheet, dst_row: int, tw_row: int) -> None:
    for col in range(1, (ws_ee.max_column or 0) + 1):
        cell = ws_ee.cell(dst_row, col)
        if cell.data_type == "f" and isinstance(cell.value, str):
            cell.value = re.sub(
                rf"TW!([A-Z]+){tw_row - 1}(?!\d)",
                lambda m: f"TW!{m.group(1)}{tw_row}",
                cell.value,
            )
