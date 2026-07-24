# -*- coding: utf-8 -*-
"""转换共用：文本规范化与单元格清洗。"""
from __future__ import annotations

import re
from typing import Any


def norm(text: Any) -> str:
    if text is None:
        return ""
    return str(text).replace("\uFEFF", "").strip()


def clean_value(value: Any) -> Any:
    """
    清洗源表单元格：空/#N/A → None；千分位数字 → int/float。
    排除 bool（True/False 是 int 子类，不能当金额写回）。
    """
    if value is None:
        return None
    s = norm(value)
    if s in ("", "#N/A", "#REF!", "#VALUE!", "-"):
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return value
    # 千分位逗号 / 中文逗号：4,149 → 4149（写成文本则公式引擎算不动）
    compact = s.replace(",", "").replace("，", "").replace(" ", "")
    try:
        if re.fullmatch(r"-?\d+(\.\d+)?", compact):
            return float(compact) if "." in compact else int(compact)
    except ValueError:
        pass
    return value
