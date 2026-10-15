# -*- coding: utf-8 -*-
"""按源表列字母（如 CK）→ 目标 -L 表头字段；与 columnRename 并列，仅配置非空时生效。"""
from __future__ import annotations

import re
from typing import Any, Callable

from openpyxl.utils import column_index_from_string
from openpyxl.worksheet.worksheet import Worksheet

from xlsx_convert_utils import norm

_COL_LETTER_RE = re.compile(r"^\$?([A-Za-z]{1,3})(?:\$?\d+)?$")


def parse_column_cell_map(mapping: dict[str, Any] | None) -> list[dict[str, Any]]:
    if not isinstance(mapping, dict):
        return []
    raw = mapping.get("columnCellMap")
    if not isinstance(raw, list):
        return []
    out: list[dict[str, Any]] = []
    for item in raw:
        if not isinstance(item, dict):
            continue
        letter = _normalize_column_letter(
            item.get("sourceColumn") or item.get("sourceCell") or item.get("column")
        )
        if not letter:
            continue
        try:
            col_idx = int(column_index_from_string(letter))
        except (ValueError, TypeError):
            continue
        if col_idx < 1:
            continue
        target = str(item.get("target") or item.get("targetHeader") or "").strip()
        if not target:
            continue
        out.append(
            {
                "sourceColumn": letter.upper(),
                "sourceCol": col_idx,
                "target": target,
            }
        )
    return out


def _normalize_column_letter(raw: Any) -> str | None:
    text = str(raw or "").strip().upper().replace(" ", "")
    if not text:
        return None
    if "!" in text:
        text = text.rsplit("!", 1)[-1]
    m = _COL_LETTER_RE.match(text)
    if not m:
        return None
    return m.group(1).upper()


def inspect_headers_for_column_cell_map(mapping: dict[str, Any] | None) -> list[dict[str, Any]]:
    """供 mapping inspect 下拉展示（不参与 columnRename 命中检查）。"""
    headers: list[dict[str, Any]] = []
    for entry in parse_column_cell_map(mapping):
        letter = entry["sourceColumn"]
        key = f"$col:{letter}"
        headers.append(
            {
                "key": key,
                "label": f"列 {letter}",
                "child": letter,
                "parent": "",
                "col": entry["sourceCol"],
                "cellMap": True,
            }
        )
    return headers


def apply_column_cell_map_to_record(
    record: dict[str, Any],
    row: int,
    ws: Worksheet,
    mapping: dict[str, Any] | None,
    *,
    coerce_number: Callable[[Any], float | None] | None = None,
    is_formula_cell: Callable[[Any], bool] | None = None,
    overwrite: bool = False,
) -> None:
    """把源表指定列当前行的值写入 record[target]。"""
    entries = parse_column_cell_map(mapping)
    if not entries or row < 1:
        return

    def _is_formula(v: Any) -> bool:
        if is_formula_cell is not None:
            return bool(is_formula_cell(v))
        if isinstance(v, str) and v.startswith("="):
            return True
        return False

    for entry in entries:
        target = entry["target"]
        col = int(entry["sourceCol"])
        if not overwrite and target in record and record.get(target) not in (None, ""):
            continue
        val = ws.cell(row, col).value
        if _is_formula(val):
            continue
        if val is None or val == "":
            continue
        if coerce_number is not None:
            num = coerce_number(val)
            if num is not None:
                record[target] = num
                continue
        record[target] = val


def merge_column_cell_map_into_source_headers(
    source_headers: dict[str, int],
    mapping: dict[str, Any] | None,
) -> dict[str, int]:
    """可选：把列字母映射合成虚拟源表头 key（$col:CK），便于与 columnRename 联调。"""
    merged = dict(source_headers)
    for entry in parse_column_cell_map(mapping):
        merged[f"$col:{entry['sourceColumn']}"] = int(entry["sourceCol"])
    return merged
