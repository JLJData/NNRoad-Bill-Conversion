# -*- coding: utf-8 -*-
"""Pakistan Business Tax：仅用 PDF 金额（Danfoss mapping.pakistanBusinessTax）。"""
from __future__ import annotations

from decimal import Decimal, ROUND_HALF_UP
from typing import Any


BT_SINDH_HEADER = "BT Sindh USD"
BT_FEDERAL_HEADER = "BT Federal IT USD"


def _as_float(value: Any) -> float | None:
    if value is None or value == "":
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def parse_pakistan_business_tax_cfg(mapping: dict[str, Any] | None) -> dict[str, Any] | None:
    """返回规范化配置；未启用则 None。"""
    if not isinstance(mapping, dict):
        return None
    raw = mapping.get("pakistanBusinessTax")
    if not isinstance(raw, dict):
        return None
    mode = str(raw.get("mode") or "").strip()
    if mode not in ("invoiceDerived", "invoice_derived", "derived", "invoiceAmounts", "pdf"):
        if not bool(raw.get("enabled")):
            return None
    return {"mode": "invoiceAmounts"}


def _round2_half_up(value: float | Decimal) -> Decimal:
    """四舍五入保留两位小数。"""
    return Decimal(str(value)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def split_months_balanced(quarter_total: float, months: int = 3) -> list[float]:
    """
    前 months-1 月：总额/月数 四舍五入到 2 位；
    末月：总额 − 前几月之和（轧平，加总等于供应商季度金额）。
    """
    n = max(int(months or 1), 1)
    total = Decimal(str(quarter_total))
    if n == 1:
        return [float(_round2_half_up(total))]
    per = _round2_half_up(total / n)
    parts = [per] * (n - 1)
    last = total - sum(parts)
    parts.append(_round2_half_up(last))
    return [float(x) for x in parts]


def format_bt_coeff(value: float) -> str:
    text = f"{float(value):.10f}".rstrip("0").rstrip(".")
    return text or "0"


def business_tax_formula(sindh_monthly: float, federal_monthly: float) -> str:
    """Pakistan!F：={sindh}*PN!$B$33+{federal}*PN!$B$33"""
    a = format_bt_coeff(sindh_monthly)
    b = format_bt_coeff(federal_monthly)
    return f"={a}*PN!$B$33+{b}*PN!$B$33"


def make_pakistan_business_tax_provenance(
    *,
    sheet: str,
    row: int,
    col: int,
    formula: str,
    sindh_monthly: float,
    federal_monthly: float,
    employee_name: str | None = None,
) -> dict[str, Any]:
    """单格 Business Tax provenance（Excel 1-based）。"""
    detail: dict[str, Any] = {
        "sindhMonthly": sindh_monthly,
        "federalMonthly": federal_monthly,
    }
    if employee_name:
        detail["employeeName"] = employee_name
    return {
        "kind": "pakistanBusinessTax",
        "sheet": sheet,
        "row": row,
        "col": col,
        "sourceType": "mapping",
        "source": "mapping.pakistanBusinessTax",
        "label": "Business Tax",
        "value": formula,
        "detail": detail,
    }


def apply_derived_coeffs_to_rows(
    employees: list[dict[str, Any]],
    cfg: dict[str, Any] | None = None,
    *,
    split_months: int = 3,
) -> list[str]:
    """
    用 PDF 季度金额拆月写入 BT 列。
    同一人连续 split_months 行：前月 round-half-up 2 位，末月轧平。
    """
    del cfg
    warnings: list[str] = []
    months = max(int(split_months or 1), 1)
    i = 0
    while i < len(employees):
        emp0 = employees[i]
        name = str(emp0.get("Name of Employee") or emp0.get("Employee Name") or "").strip()
        # 同一人连续行（ingest 按人拆月）
        chunk = [emp0]
        j = i + 1
        while j < len(employees) and len(chunk) < months:
            emp_j = employees[j]
            name_j = str(emp_j.get("Name of Employee") or emp_j.get("Employee Name") or "").strip()
            if name_j != name:
                break
            chunk.append(emp_j)
            j += 1

        sindh_q = _as_float(
            emp0.get("_sindh_sales_tax_usd")
            if emp0.get("_sindh_sales_tax_usd") is not None
            else emp0.get("sindh_sales_tax_usd")
        )
        fed_q = _as_float(
            emp0.get("_federal_it_usd") if emp0.get("_federal_it_usd") is not None else emp0.get("federal_it_usd")
        )
        if sindh_q is None or fed_q is None:
            warnings.append(
                f"{name or '员工'}：PDF 未解析到 Sindh Sales Tax / Federal IT 金额，Business Tax 未写入"
            )
            i = j if j > i else i + 1
            continue

        n = len(chunk)
        sindh_parts = split_months_balanced(sindh_q, n)
        fed_parts = split_months_balanced(fed_q, n)
        for k, emp in enumerate(chunk):
            sindh_m = sindh_parts[k]
            fed_m = fed_parts[k]
            emp[BT_SINDH_HEADER] = sindh_m
            emp[BT_FEDERAL_HEADER] = fed_m
            emp["_bt_sindh_usd"] = sindh_m
            emp["_bt_federal_usd"] = fed_m
        i = j if j > i else i + 1
    return warnings
