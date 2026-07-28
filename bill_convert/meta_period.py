# -*- coding: utf-8 -*-
from __future__ import annotations

import re
from datetime import date, datetime
from typing import Any

from convert_mapping import find_sheet_name
from xlsx_convert_utils import coerce_datetime_for_excel, norm


def parse_period(value: Any, payroll_month: Any = None) -> tuple[Any, Any]:
    """解析 Summary 中 Period 如 3/1-3/31"""
    if value is None:
        return None, None
    s = norm(value)
    m = re.match(r"(\d{1,2})/(\d{1,2})\s*-\s*(\d{1,2})/(\d{1,2})", s)
    if not m:
        return None, None

    year = None
    if payroll_month is not None:
        if isinstance(payroll_month, datetime):
            year = payroll_month.year
        elif isinstance(payroll_month, date):
            year = payroll_month.year
        else:
            ps = norm(payroll_month)
            for fmt in ("%Y-%m-%d", "%Y-%m-%d %H:%M:%S", "%Y/%m/%d"):
                try:
                    year = datetime.strptime(ps[:19], fmt).year
                    break
                except ValueError:
                    continue
    if year is None:
        year = datetime.now().year

    m1, d1, m2, d2 = (int(m.group(i)) for i in range(1, 5))
    return (
        datetime(year, m1, d1),
        datetime(year, m2, d2),
    )


def payroll_month_start(value: Any) -> datetime | None:
    dt = coerce_datetime_for_excel(value)
    if dt is None:
        return None
    return datetime(dt.year, dt.month, 1)


DEFAULT_SUMMARY_LABELS = {
    "Client": "company_name",
    "Payroll Month": "payroll_month",
    "Period": "period_raw",
    "Exchange rate": "exchange_rate",
}


def coerce_exchange_rate(value: Any) -> float | None:
    """Summary「Exchange rate」→ PN 汇率单元格数值。"""
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    s = norm(value).replace(",", "").strip()
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def read_summary_meta(
    wb,
    meta_spec: dict[str, Any] | None,
    *,
    default_sheet: str = "Summary",
    label_map: dict[str, str] | None = None,
    scan_max_row: int = 20,
) -> dict[str, Any]:
    spec = meta_spec if isinstance(meta_spec, dict) else {}
    optional = bool(spec.get("optional", True))
    sheet_name = find_sheet_name(list(wb.sheetnames), spec if spec else {"sheet": default_sheet})
    if not sheet_name:
        if optional:
            return {}
        want = spec.get("sheet") or default_sheet
        raise ValueError(f"未找到 Summary sheet「{want}」")
    ws = wb[sheet_name]
    labels = label_map or DEFAULT_SUMMARY_LABELS
    meta: dict[str, Any] = {}
    for row in range(1, scan_max_row + 1):
        label = norm(ws.cell(row, 1).value)
        val = ws.cell(row, 2).value
        for src_label, meta_key in labels.items():
            if label == norm(src_label):
                meta[meta_key] = val
    period_from, period_to = parse_period(meta.get("period_raw"), meta.get("payroll_month"))
    meta["period_from"] = period_from
    meta["period_to"] = period_to
    meta["payroll_month_start"] = payroll_month_start(meta.get("payroll_month"))
    fx = coerce_exchange_rate(meta.get("exchange_rate"))
    if fx is not None:
        meta["exchange_rate"] = fx
    elif "exchange_rate" in meta:
        meta["exchange_rate"] = None
    return meta
