# -*- coding: utf-8 -*-
from __future__ import annotations

from typing import Any

from openpyxl.worksheet.worksheet import Worksheet

from bill_convert.header_scan import find_header_row_by_markers
from bill_convert.mapping_spec import mapping_section
from convert_mapping import find_sheet_name
from xlsx_convert_utils import norm

DEFAULT_TARGET_L_HEADER_ROW = 7
DEFAULT_TARGET_L_MARKER_KEYS = frozenset({"BU", "CN Name", "EN Name"})


def _auto_detect_target_l(mapping: dict[str, Any]) -> bool:
    spec = mapping_section(mapping, "targetL")
    if "autoDetectLayout" in spec:
        return bool(spec.get("autoDetectLayout"))
    if "fixedLayout" in spec:
        return not bool(spec.get("fixedLayout"))
    return True


def target_l_marker_keys(mapping: dict[str, Any]) -> frozenset[str]:
    spec = mapping_section(mapping, "targetL")
    if isinstance(spec.get("headerMarkerKeys"), list):
        return frozenset(norm(str(x)) for x in spec["headerMarkerKeys"])
    src = mapping_section(mapping, "sourceEmployeeSheet")
    if isinstance(src.get("headerMarkerKeys"), list):
        return frozenset(norm(str(x)) for x in src["headerMarkerKeys"])
    return DEFAULT_TARGET_L_MARKER_KEYS


def find_target_l_header_row(
    ws: Worksheet,
    mapping: dict[str, Any],
    *,
    sheet_label: str | None = None,
    default_sheet_name: str = "TW-L",
) -> int:
    spec = mapping_section(mapping, "targetL")
    scan = int(spec.get("headerScanMaxRow") or 15)
    keys = target_l_marker_keys(mapping)
    label = sheet_label or str(spec.get("sheet") or default_sheet_name)
    return find_header_row_by_markers(
        ws, marker_keys=keys, max_scan=scan, sheet_label=label
    )


def resolve_target_l_layout(
    ws: Worksheet,
    mapping: dict[str, Any],
    *,
    sheet_label: str | None = None,
    default_sheet_name: str = "TW-L",
    fallback_header_row: int = DEFAULT_TARGET_L_HEADER_ROW,
) -> dict[str, int]:
    spec = mapping_section(mapping, "targetL")
    label = sheet_label or str(spec.get("sheet") or default_sheet_name)
    if _auto_detect_target_l(mapping):
        header_row = find_target_l_header_row(
            ws, mapping, sheet_label=label, default_sheet_name=default_sheet_name
        )
        if spec.get("dataStartRow") is not None:
            data_start = int(spec["dataStartRow"])
            if data_start <= header_row:
                data_start = header_row + int(spec.get("dataStartOffset") or 2)
        else:
            data_start = header_row + int(spec.get("dataStartOffset") or 2)
    else:
        header_row = int(spec.get("headerRow") or fallback_header_row)
        data_start = int(spec.get("dataStartRow") or (header_row + int(spec.get("dataStartOffset") or 2)))
    summary_row = int(spec["summaryRow"]) if spec.get("summaryRow") is not None else header_row + 1
    return {
        "header_row": header_row,
        "summary_row": summary_row,
        "data_start_row": data_start,
    }


def resolve_target_l_sheet_name(
    sheet_names: list[str],
    mapping: dict[str, Any],
    *,
    default_sheet_name: str = "TW-L",
) -> str:
    spec = mapping_section(mapping, "targetL")
    want = str(spec.get("sheet") or default_sheet_name)
    candidates = spec.get("candidates")
    find_spec: dict[str, Any] = {"sheet": want}
    if isinstance(candidates, list):
        find_spec["candidates"] = candidates
    else:
        find_spec["candidates"] = [want, default_sheet_name]
    name = find_sheet_name(sheet_names, find_spec)
    if not name:
        raise ValueError(f"母版中未找到 sheet「{want}」，现有: {sheet_names}")
    return name


def target_l_auto_detect_layout(mapping: dict[str, Any]) -> bool:
    return _auto_detect_target_l(mapping)
