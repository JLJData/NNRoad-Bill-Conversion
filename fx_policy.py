# -*- coding: utf-8 -*-
"""汇率策略：从 convert_mapping.fxPolicy / factStore / 供应商货币解析，供各引擎共用。

约定（勿写死 profile_code）：
- mode=vendor_bill：优先账单汇率，可 fallback api
- mode=api：网上汇率（可用 vendorCurrency）
- mode=fixed：固定值/母版公式，不调 API 覆盖
- mode=shared_fact：读 factStore / artifactBatch 中的共享汇率
- mode=none：不写汇率
"""
from __future__ import annotations

from typing import Any

from bill_convert.fact_store import get_batch_facts, get_fact_value
from fx_rate import fetch_usd_rates, get_usd_rate

# 与 TopSource 账单同源；同客户其它 UK 配置（如 EOR）按 pdfProfile 解析共享，不写死 profile_code
UK_VENDOR_BILL_FX_FACT = "uk.vendor_bill.fx_rate"


def fx_policy(mapping: dict[str, Any] | None) -> dict[str, Any]:
    if not isinstance(mapping, dict):
        return {}
    raw = mapping.get("fxPolicy")
    return dict(raw) if isinstance(raw, dict) else {}


def resolve_vendor_currency(mapping: dict[str, Any] | None, default: str | None = None) -> str | None:
    policy = fx_policy(mapping)
    if isinstance(mapping, dict):
        for key in ("vendorCurrency", "currency"):
            v = mapping.get(key)
            if v is not None and str(v).strip():
                return str(v).strip().upper()
    if policy.get("useVendorCurrency"):
        # 已在上面扫过 mapping；此处仅 defaultCurrency
        pass
    dc = policy.get("defaultCurrency") or default
    return str(dc).strip().upper() if dc else None


def _as_float(value: Any) -> float | None:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def read_shared_fx(mapping: dict[str, Any] | None, fact_key: str | None = None) -> float | None:
    """优先本批 artifactBatch，再 factStore。"""
    key = (fact_key or fx_policy(mapping).get("factKey") or UK_VENDOR_BILL_FX_FACT).strip()
    if not key:
        return None
    batch = get_batch_facts(mapping)
    if key in batch:
        fx = _as_float(batch.get(key))
        if fx is None and isinstance(batch.get(key), dict):
            fx = _as_float(batch[key].get("value"))
        if fx is not None and fx > 0:
            return fx
    fx = _as_float(get_fact_value(mapping, key))
    if fx is not None and fx > 0:
        return fx
    return None


def api_fx_for_currency(
    currency: str,
    *,
    adjustment: float = 1.0,
    invert: bool = False,
    rates: dict[str, float] | None = None,
) -> float:
    code = currency.strip().upper()
    table = rates if rates is not None else fetch_usd_rates()
    rate = get_usd_rate(code, table)
    if invert:
        if rate <= 0:
            raise RuntimeError(f"{code} 汇率无效")
        return round((1.0 / rate) * float(adjustment or 1.0), 10)
    return round(float(rate) * float(adjustment or 1.0), 10)


def build_fx_fact_update(fact_key: str, value: float, *, source: str | None = None) -> dict[str, Any]:
    entry: dict[str, Any] = {"value": float(value)}
    if source:
        entry["source"] = source
    return {fact_key: entry}


def _fmt_fx_num(value: float) -> str:
    text = f"{float(value):.10f}".rstrip("0").rstrip(".")
    return text or "0"


def fixed_fx_parts(mapping: dict[str, Any] | None) -> tuple[str | None, float | None]:
    """mode=fixed：返回 (Excel 公式, 乘积)。

    优先 mapping 上的 baseKey × adjustmentKey（默认 uaePnFxBase / uaePnFxAdjustment），
    缺省用 fxPolicy.defaultBase / defaultAdjustment。转换时应覆盖母版，不必先同步母版。
    旧字段 valueKey（uaePnFxRate）仅兼容为单值。
    """
    policy = fx_policy(mapping)
    data = mapping if isinstance(mapping, dict) else {}
    base_key = str(policy.get("baseKey") or "uaePnFxBase").strip()
    adj_key = str(policy.get("adjustmentKey") or "uaePnFxAdjustment").strip()
    has_base = base_key in data and data.get(base_key) not in (None, "")
    has_adj = adj_key in data and data.get(adj_key) not in (None, "")
    value_key = str(policy.get("valueKey") or "uaePnFxRate").strip()
    if not has_base and not has_adj:
        old = _as_float(data.get(value_key)) if value_key else None
        if old is not None and old > 0:
            return None, old
    base = _as_float(data.get(base_key)) if has_base else None
    adj = _as_float(data.get(adj_key)) if has_adj else None
    if base is None:
        base = _as_float(policy.get("defaultBase"))
    if adj is None:
        adj = _as_float(policy.get("defaultAdjustment"))
    if base is None:
        base = 3.6725
    if adj is None:
        adj = 0.97
    if base <= 0 or adj <= 0:
        return None, None
    formula = f"={_fmt_fx_num(base)}*{_fmt_fx_num(adj)}"
    return formula, float(base) * float(adj)


def fixed_fx_override(mapping: dict[str, Any] | None) -> float | None:
    """mode=fixed 时的覆盖值（乘积；无配置则 None）。"""
    _formula, product = fixed_fx_parts(mapping)
    return product
