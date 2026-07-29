# -*- coding: utf-8 -*-
from __future__ import annotations

import re
from copy import copy
from typing import Any

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


def snapshot_row_cells(ws: Worksheet, row: int) -> list[dict[str, Any]]:
    """拷贝一行公式/值与样式，避免后续覆盖源行后丢失示例。"""
    out: list[dict[str, Any]] = []
    max_col = ws.max_column or 1
    for col in range(1, max_col + 1):
        src = ws.cell(row, col)
        item: dict[str, Any] = {
            "col": col,
            "value": src.value,
            "data_type": src.data_type,
        }
        if src.has_style:
            item["font"] = copy(src.font)
            item["border"] = copy(src.border)
            item["fill"] = copy(src.fill)
            item["number_format"] = src.number_format
            item["protection"] = copy(src.protection)
            item["alignment"] = copy(src.alignment)
        out.append(item)
    return out


def copy_row_formulas_from_snapshot(
    snapshot: list[dict[str, Any]],
    ws: Worksheet,
    from_row: int,
    to_row: int,
    target_l_from: int,
    target_l_to: int,
    *,
    target_l_sheet: str = "TW-L",
) -> None:
    for item in snapshot:
        col = int(item["col"])
        dst = ws.cell(to_row, col)
        if "font" in item:
            dst.font = item["font"]
            dst.border = item["border"]
            dst.fill = item["fill"]
            dst.number_format = item["number_format"]
            dst.protection = item["protection"]
            dst.alignment = item["alignment"]
        value = item.get("value")
        data_type = item.get("data_type")
        if data_type == "f" and isinstance(value, str):
            dst.value = shift_row_formula(
                value,
                from_row,
                to_row,
                target_l_from=target_l_from,
                target_l_to=target_l_to,
                target_l_sheet=target_l_sheet,
            )
        elif value is not None and data_type != "f":
            dst.value = copy(value)


def copy_row_formulas(
    ws: Worksheet,
    from_row: int,
    to_row: int,
    target_l_from: int,
    target_l_to: int,
    *,
    target_l_sheet: str = "TW-L",
) -> None:
    copy_row_formulas_from_snapshot(
        snapshot_row_cells(ws, from_row),
        ws,
        from_row,
        to_row,
        target_l_from,
        target_l_to,
        target_l_sheet=target_l_sheet,
    )


def fix_sheet_cross_refs(
    ws: Worksheet,
    dst_row: int,
    *,
    other_sheet: str,
    other_row: int,
    quoted: bool,
) -> None:
    """修正同行公式里对其它 sheet 上一行引用（复制扩展后常见 off-by-one）。"""
    prefix = f"'{other_sheet}'!" if quoted else f"{other_sheet}!"
    pat = re.compile(rf"{re.escape(prefix)}([A-Z]+){other_row - 1}(?!\d)")
    repl_prefix = prefix
    for col in range(1, (ws.max_column or 0) + 1):
        cell = ws.cell(dst_row, col)
        if cell.data_type == "f" and isinstance(cell.value, str):
            cell.value = pat.sub(lambda m: f"{repl_prefix}{m.group(1)}{other_row}", cell.value)


def fix_tw_row_tw_ee_refs(ws_tw: Worksheet, dst_row: int, ee_row: int) -> None:
    fix_sheet_cross_refs(ws_tw, dst_row, other_sheet="TW EE", other_row=ee_row, quoted=True)


def fix_ee_row_tw_refs(ws_ee: Worksheet, dst_row: int, tw_row: int) -> None:
    fix_sheet_cross_refs(ws_ee, dst_row, other_sheet="TW", other_row=tw_row, quoted=False)


def fix_china_row_china_ee_refs(ws_china: Worksheet, dst_row: int, ee_row: int) -> None:
    fix_sheet_cross_refs(ws_china, dst_row, other_sheet="China EE", other_row=ee_row, quoted=True)


def fix_ee_row_china_refs(ws_ee: Worksheet, dst_row: int, china_row: int) -> None:
    fix_sheet_cross_refs(ws_ee, dst_row, other_sheet="China", other_row=china_row, quoted=False)
