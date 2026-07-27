# -*- coding: utf-8 -*-
"""转换共用：文本规范化与单元格清洗。"""
from __future__ import annotations

import re
from datetime import date, datetime, timedelta
from typing import Any


def norm(text: Any) -> str:
    if text is None:
        return ""
    return str(text).replace("\uFEFF", "").strip()


def is_excel_date_serial(n: float) -> bool:
    return 25000 <= n <= 65000 and abs(n - round(n)) < 1e-6


def excel_serial_to_datetime(serial: float) -> datetime:
    """Excel 1900 日期系统 → 仅日期部分的 datetime（供 openpyxl 写入）。"""
    days = int(round(float(serial)))
    base = datetime(1899, 12, 30)
    return base + timedelta(days=days)


def coerce_datetime_for_excel(value: Any) -> datetime | None:
    """
    把源表/中间值规范为 openpyxl 可写的 datetime（避免写字符串或裸序列进模板）。
    无法识别为日期则返回 None。
    """
    if value is None:
        return None
    if isinstance(value, datetime):
        return datetime(value.year, value.month, value.day)
    if isinstance(value, date):
        return datetime(value.year, value.month, value.day)
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if is_excel_date_serial(float(value)):
            return excel_serial_to_datetime(float(value))
        return None
    s = norm(value).lstrip("'").strip()
    if not s:
        return None
    for fmt in (
        "%Y-%m-%d",
        "%Y-%m-%d %H:%M:%S",
        "%Y/%m/%d",
        "%m/%d/%Y",
        "%d/%m/%Y",
        "%Y/%m/%d %H:%M:%S",
    ):
        try:
            d = datetime.strptime(s[:19], fmt).date()
            return datetime(d.year, d.month, d.day)
        except ValueError:
            continue
    if re.fullmatch(r"\d{4,5}(\.\d+)?", s):
        n = float(s)
        if is_excel_date_serial(n):
            return excel_serial_to_datetime(n)
    return None


def is_date_column_header(header: str) -> bool:
    """目标表头是否为日期列（写入时用 datetime，勿写文本/裸序列）。"""
    h = norm(header).lower()
    if not h:
        return False
    if "start date" in h or "end date" in h:
        return True
    if h in ("pay period", "payroll month", "period from", "period to"):
        return True
    if h.endswith(" date") or h.endswith("日期"):
        return True
    return False


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
