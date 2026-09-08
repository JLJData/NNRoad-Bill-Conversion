# -*- coding: utf-8 -*-
"""PDF 正文：标签 → 金额 轻量抽取（供版式可配置配对）。"""
from __future__ import annotations

import re
from typing import Any


def norm_label(value: str | None) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).replace("\xa0", " ")).strip()


def norm_label_key(value: str | None) -> str:
    return norm_label(value).lower()


def parse_money(value: Any) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    text = str(value).strip().replace("\xa0", "").replace(" ", "")
    if not text:
        return None
    # 4.448,82 / 1.137,16
    if re.search(r",\d{2}$", text) and "." in text:
        text = text.replace(".", "").replace(",", ".")
    elif text.count(",") == 1 and "." not in text:
        text = text.replace(",", ".")
    else:
        text = text.replace(",", "")
    try:
        return float(text)
    except ValueError:
        return None


def extract_label_amounts(
    text: str,
    labels: list[str],
) -> list[dict[str, Any]]:
    """
    在文本中查找已知标签及其后金额。
    返回 [{label, value, start, end}, ...]，按出现位置排序。
    长标签优先，避免短标签吃掉长标签前缀。
    """
    if not text or not labels:
        return []
    uniq: list[str] = []
    seen: set[str] = set()
    for lab in labels:
        key = norm_label_key(lab)
        if not key or key in seen:
            continue
        seen.add(key)
        uniq.append(norm_label(lab))

    # 金额：先美式千分位 4,448.82，再欧式 4.448,82，避免 4,448.82 被吃成 4,44
    amount = (
        r"(?:"
        r"\d{1,3}(?:,\d{3})+(?:\.\d+)?"
        r"|\d{1,3}(?:\.\d{3})+,\d{2}"
        r"|\d+,\d{2}"
        r"|\d+\.\d+"
        r"|\d+"
        r")"
    )
    hits: list[dict[str, Any]] = []
    occupied: list[tuple[int, int]] = []
    for lab in sorted(uniq, key=len, reverse=True):
        pat = re.compile(re.escape(lab) + r"\s*" + amount, flags=re.I)
        for m in pat.finditer(text):
            span = (m.start(), m.end())
            if any(not (span[1] <= a or span[0] >= b) for a, b in occupied):
                continue
            # 金额在标签之后
            raw_amt = m.group(0)[len(lab) :].strip()
            # group 整段匹配；重新取尾部数字
            am = re.search(amount + r"\s*$", m.group(0), flags=re.I)
            raw_amt = am.group(0).strip() if am else raw_amt
            val = parse_money(raw_amt)
            if val is None:
                continue
            occupied.append(span)
            hits.append(
                {
                    "label": lab,
                    "value": val,
                    "start": m.start(),
                    "end": m.end(),
                }
            )
    hits.sort(key=lambda x: int(x["start"]))
    return hits


def build_label_rename_map(
    *,
    builtin: dict[str, str] | None = None,
    column_rename: dict[str, str] | None = None,
) -> dict[str, str]:
    """
    vendor_label_key(lower) → 目标字段名（Cyprus-L / 内部键）。
    columnRename（Office 保存）覆盖 builtin。
    """
    out: dict[str, str] = {}
    for src, dst in (builtin or {}).items():
        sk, dk = norm_label_key(src), norm_label(dst)
        if sk and dk:
            out[sk] = dk
    for src, dst in (column_rename or {}).items():
        sk, dk = norm_label_key(src), norm_label(dst)
        if sk and dk:
            out[sk] = dk
    return out


def apply_label_amounts(
    hits: list[dict[str, Any]],
    rename_map: dict[str, str],
) -> dict[str, float]:
    """把抽取结果按 rename 落到字段；同字段多次出现则累加。"""
    fields: dict[str, float] = {}
    for hit in hits:
        raw = norm_label(str(hit.get("label") or ""))
        key = norm_label_key(raw)
        target = rename_map.get(key) or raw
        if not target:
            continue
        val = hit.get("value")
        if not isinstance(val, (int, float)):
            continue
        fields[target] = float(fields.get(target, 0.0)) + float(val)
    return fields
