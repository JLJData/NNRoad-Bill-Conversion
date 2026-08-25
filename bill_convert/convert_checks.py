# -*- coding: utf-8 -*-
"""转换过程的低风险校验：缺列失败/回退警告、结果体检（不改计算结果）。"""
from __future__ import annotations

import re
from typing import Any

from bill_convert.headers import norm


def find_header_col(header_map: dict[str, int], candidates: list[str] | tuple[str, ...] | None) -> int | None:
    """按候选表头名解析列号（大小写/空白经 norm）。"""
    if not header_map or not candidates:
        return None
    for name in candidates:
        key = norm(name)
        if key and key in header_map:
            return int(header_map[key])
    lower_index = {norm(k).lower(): c for k, c in header_map.items() if norm(k)}
    for name in candidates:
        key = norm(name).lower()
        if key and key in lower_index:
            return int(lower_index[key])
    return None


def resolve_col_with_fallback(
    header_map: dict[str, int],
    *,
    field: str,
    header_names: list[str] | tuple[str, ...] | None = None,
    fallback: int | None = None,
    warnings: list[str] | None = None,
    strict: bool = False,
) -> int | None:
    """
    优先按表头找列；找不到时：
    - strict=True → 抛错
    - 有 fallback → 回退并警告（默认母版兼容）
    - 否则返回 None 并警告
    """
    names_list = list(header_names or [])
    col = find_header_col(header_map, names_list)
    if col is not None:
        return col
    names = " / ".join(str(x) for x in names_list if x)
    if strict:
        raise ValueError(f"未找到必需列「{field}」（表头候选: {names}）")
    if fallback is not None:
        if warnings is not None:
            warnings.append(
                f"未找到表头「{field}」({names})，已回退固定列 {fallback}；"
                f"若供应商调整了列位置请改映射/表头，勿依赖列号"
            )
        return int(fallback)
    if warnings is not None:
        warnings.append(f"未找到列「{field}」({names})，该字段跳过")
    return None


def parse_cell_ref(ref: Any, *, default_row: int, default_col: int) -> tuple[int, int]:
    """解析 A1 / C2 或 {row,col}；失败则返回默认。"""
    if isinstance(ref, dict):
        try:
            return int(ref.get("row") or default_row), int(ref.get("col") or default_col)
        except (TypeError, ValueError):
            return default_row, default_col
    text = str(ref or "").strip().upper()
    if not text:
        return default_row, default_col
    m = re.fullmatch(r"([A-Z]+)(\d+)", text)
    if not m:
        return default_row, default_col
    letters, row_s = m.group(1), m.group(2)
    col = 0
    for ch in letters:
        col = col * 26 + (ord(ch) - ord("A") + 1)
    return int(row_s), col


def check_column_rename_hits(
    rename: dict[str, Any] | None,
    source_headers: list[str] | set[str] | dict[str, Any],
    *,
    warnings: list[str] | None = None,
    strict_if_configured: bool = True,
) -> int:
    """
    columnRename 已配置时，检查是否至少命中 1 个源表头。
    返回命中数；strict 且 0 命中时抛错。
    """
    if not isinstance(rename, dict) or not rename:
        return 0
    if isinstance(source_headers, dict):
        keys = {norm(k) for k in source_headers.keys() if norm(k)}
    else:
        keys = {norm(k) for k in source_headers if norm(k)}
    if not keys:
        return 0
    keys_lower = {k.lower() for k in keys}
    hits = 0
    misses: list[str] = []
    for src in rename.keys():
        sk = norm(src)
        if not sk:
            continue
        if sk in keys or sk.lower() in keys_lower:
            hits += 1
        else:
            misses.append(str(src))
    if hits == 0 and rename:
        msg = (
            "columnRename 已配置但未命中任何源表头，请核对映射里的供应商列名是否与账单一致"
            + (f"（样例未命中: {', '.join(misses[:5])}）" if misses else "")
        )
        if strict_if_configured:
            raise ValueError(msg)
        if warnings is not None:
            warnings.append(msg)
    elif misses and warnings is not None:
        warnings.append(
            f"columnRename 部分未命中（{len(misses)} 项），例如: {', '.join(misses[:5])}"
        )
    return hits


def sanity_check_convert_result(result: dict[str, Any] | None) -> list[str]:
    """转换结果轻量体检：只产警告，不改文件/数值。"""
    out: list[str] = []
    if not isinstance(result, dict):
        return ["转换结果为空，请人工核对"]
    emp = result.get("employee_count")
    if emp is None and isinstance(result.get("employees"), list):
        emp = len(result["employees"])
    try:
        emp_n = int(emp) if emp is not None else None
    except (TypeError, ValueError):
        emp_n = None
    if emp_n == 0:
        out.append("转换结果员工数为 0，请核对源表姓名列/表头行是否变化")
    elif emp_n is not None and emp_n > 80:
        out.append(f"转换结果员工数偏多（{emp_n}），请确认是否误读表头或空行")
    warnings = result.get("warnings")
    if isinstance(warnings, list):
        fallback_n = sum(1 for w in warnings if isinstance(w, str) and "回退固定列" in w)
        if fallback_n >= 3:
            out.append(
                f"有 {fallback_n} 个字段靠固定列号回退写入，供应商账单列位置可能已变，建议按表头核对"
            )
    return out


def merge_warnings(*groups: list[str] | None) -> list[str]:
    seen: set[str] = set()
    out: list[str] = []
    for g in groups:
        if not g:
            continue
        for w in g:
            s = str(w).strip()
            if not s or s in seen:
                continue
            seen.add(s)
            out.append(s)
    return out
