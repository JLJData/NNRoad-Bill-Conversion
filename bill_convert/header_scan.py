# -*- coding: utf-8 -*-
from __future__ import annotations

from collections.abc import Callable
from typing import Any

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


def marker_keys_from_spec(spec: dict[str, Any] | None) -> frozenset[str]:
    if not isinstance(spec, dict):
        return frozenset()
    raw = spec.get("headerMarkerKeys")
    if not isinstance(raw, list):
        return frozenset()
    return frozenset(norm(str(x)) for x in raw if x is not None and str(x).strip())


def resolve_header_row(
    ws: Worksheet,
    *,
    spec: dict[str, Any] | None = None,
    marker_keys: frozenset[str] | None = None,
    max_scan: int | None = None,
    min_hits: int = 2,
    sheet_label: str = "",
    fixed_row: int | None = None,
    fallback: Callable[[], int] | None = None,
) -> int:
    """
    定位表头行：
    1) 配置了 ≥min_hits 个 headerMarkerKeys → 按标志列扫描
    2) 否则 fixed_row（来自 mapping.headerRow）
    3) 否则 fallback()（引擎原有启发式）
    """
    spec = spec if isinstance(spec, dict) else {}
    keys = marker_keys if marker_keys is not None else marker_keys_from_spec(spec)
    scan = max_scan if max_scan is not None else int(spec.get("headerScanMaxRow") or 20)
    if len(keys) >= min_hits:
        return find_header_row_by_markers(
            ws,
            marker_keys=keys,
            max_scan=scan,
            min_hits=min_hits,
            sheet_label=sheet_label or str(spec.get("sheet") or ""),
        )
    if fixed_row is not None:
        return int(fixed_row)
    if spec.get("headerRow") is not None:
        return int(spec["headerRow"])
    if fallback is not None:
        return fallback()
    raise ValueError(f"「{sheet_label or '工作表'}」无法定位表头行：未配置标志列且无固定行号")
