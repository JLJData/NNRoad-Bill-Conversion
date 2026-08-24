# -*- coding: utf-8 -*-
"""汇率策略：从 convert_mapping.fxPolicy / factStore / 供应商货币解析，供各引擎共用。

约定（勿写死 profile_code）：
- mode=vendor_bill：优先账单汇率，可 fallback api
- mode=api：网上汇率（可用 vendorCurrency）
- mode=api_as_base：网上汇率作「基准」，与调整系数写成 PN 公式 =基准*系数
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


def _fx_keys(policy: dict[str, Any]) -> tuple[str, str, str]:
    base_key = str(policy.get("baseKey") or "uaePnFxBase").strip()
    adj_key = str(policy.get("adjustmentKey") or "uaePnFxAdjustment").strip()
    value_key = str(policy.get("valueKey") or "uaePnFxRate").strip()
    return base_key, adj_key, value_key


def read_fx_base_adjustment(
    mapping: dict[str, Any] | None,
) -> tuple[float | None, float | None, bool]:
    """读 mapping/policy 的基准与系数。

    返回 (base, adjustment, had_legacy_single_value)。
    base 可为 None（未配置且无 defaultBase），供 api_as_base 再填网上汇率。
    """
    policy = fx_policy(mapping)
    data = mapping if isinstance(mapping, dict) else {}
    base_key, adj_key, value_key = _fx_keys(policy)
    has_base = base_key in data and data.get(base_key) not in (None, "")
    has_adj = adj_key in data and data.get(adj_key) not in (None, "")
    if not has_base and not has_adj:
        old = _as_float(data.get(value_key)) if value_key else None
        if old is not None and old > 0:
            return None, None, True
    base = _as_float(data.get(base_key)) if has_base else None
    adj = _as_float(data.get(adj_key)) if has_adj else None
    if base is None:
        base = _as_float(policy.get("defaultBase"))
    if adj is None:
        adj = _as_float(policy.get("defaultAdjustment"))
    if adj is None:
        adj = 0.97
    return base, adj, False


def fixed_fx_parts(mapping: dict[str, Any] | None) -> tuple[str | None, float | None]:
    """mode=fixed：返回 (Excel 公式, 乘积)。

    优先 mapping 上的 baseKey × adjustmentKey（默认 uaePnFxBase / uaePnFxAdjustment），
    缺省用 fxPolicy.defaultBase / defaultAdjustment。转换时应覆盖母版，不必先同步母版。
    旧字段 valueKey（uaePnFxRate）仅兼容为单值。
    """
    policy = fx_policy(mapping)
    data = mapping if isinstance(mapping, dict) else {}
    base, adj, legacy = read_fx_base_adjustment(mapping)
    if legacy:
        _base_key, _adj_key, value_key = _fx_keys(policy)
        old = _as_float(data.get(value_key)) if value_key else None
        if old is not None and old > 0:
            return None, old
    if base is None:
        base = 3.6725  # UAE 历史缺省；India 等用 api_as_base / resolve_fx_write_parts
    if adj is None:
        adj = 0.97
    if base <= 0 or adj <= 0:
        return None, None
    formula = f"={_fmt_fx_num(base)}*{_fmt_fx_num(adj)}"
    return formula, float(base) * float(adj)


def resolve_fx_round_digits(mapping: dict[str, Any] | None) -> int | None:
    """mapping.roundDigitsKey：null=完整浮点；数字=ROUND 位数。键未出现时用 fxPolicy.roundDigits。"""
    policy = fx_policy(mapping)
    data = mapping if isinstance(mapping, dict) else {}
    round_key = str(policy.get("roundDigitsKey") or "").strip()
    if round_key and round_key in data:
        raw = data.get(round_key)
        if raw is None or raw == "":
            return None
        try:
            return int(raw)
        except (TypeError, ValueError):
            return None
    raw_def = policy.get("roundDigits")
    try:
        return int(raw_def) if raw_def is not None and raw_def != "" else None
    except (TypeError, ValueError):
        return None


def fx_display_number_format(round_digits: int | None) -> str:
    """PN 汇率格显示格式：与 ROUND 位数对齐；完整浮点用较多小数位。"""
    if round_digits is not None and round_digits >= 0:
        if round_digits == 0:
            return "0"
        return "0." + ("0" * int(round_digits))
    return "0.##########"


def resolve_fx_write_parts(
    mapping: dict[str, Any] | None,
    *,
    api_base: float | None = None,
) -> tuple[str | None, float | None, str]:
    """统一写出 PN 汇率公式。

    - 有 mapping/defaultBase 基准 → =基准*系数
    - 否则若传入 api_base（api_as_base）→ =网上基准*系数
    - roundDigitsKey（如 indiaPnFxRoundDigits）在 mapping 中：null=完整浮点，数字=ROUND(…,n)
      未出现该键时用 fxPolicy.roundDigits 引擎默认
    - 旧单值字段 → 无公式，仅乘积

    返回 (formula, product, source)。
    """
    policy = fx_policy(mapping)
    data = mapping if isinstance(mapping, dict) else {}
    base, adj, legacy = read_fx_base_adjustment(mapping)
    if legacy:
        _bk, _ak, value_key = _fx_keys(policy)
        old = _as_float(data.get(value_key)) if value_key else None
        if old is not None and old > 0:
            return None, old, "legacy_value"
    source = "mapping"
    if base is None:
        if api_base is not None and float(api_base) > 0:
            base = float(api_base)
            source = "api_as_base"
        else:
            base = _as_float(policy.get("defaultBase"))
            if base is not None:
                source = "defaultBase"
    if adj is None:
        adj = 0.97
    if base is None or base <= 0 or adj <= 0:
        return None, None, "none"

    round_digits = resolve_fx_round_digits(mapping)
    product = float(base) * float(adj)
    body = f"{_fmt_fx_num(base)}*{_fmt_fx_num(adj)}"
    if round_digits is not None and round_digits >= 0:
        product = round(product, round_digits)
        formula = f"=ROUND({body},{round_digits})"
    else:
        formula = f"={body}"
    return formula, product, source


def apply_fx_formula_to_cell_ex(
    cell,
    mapping: dict[str, Any] | None,
    *,
    api_base: float | None = None,
) -> tuple[float | None, str]:
    """写 PN 汇率公式并同步单元格显示位数；返回 (乘积, write_source)。"""
    formula, product, source = resolve_fx_write_parts(mapping, api_base=api_base)
    digits = resolve_fx_round_digits(mapping)
    cell.number_format = fx_display_number_format(digits)
    if formula:
        cell.value = formula
        return product, source
    if product is not None and product > 0:
        cell.value = float(product)
        return float(product), source
    return None, source


def apply_fx_formula_to_cell(cell, mapping: dict[str, Any] | None, *, api_base: float | None = None) -> float | None:
    """写 PN 汇率公式并同步单元格显示位数；返回乘积。"""
    product, _source = apply_fx_formula_to_cell_ex(cell, mapping, api_base=api_base)
    return product


def make_pn_fx_provenance(
    sheet: str,
    row: int,
    col: int,
    mapping: dict[str, Any] | None,
    value: float | None,
    *,
    write_source: str,
    fx_source: str | None = None,
) -> dict[str, Any] | None:
    """构建 PN/地区汇率格 provenance（Excel 1-based）。"""
    if value is None:
        return None
    policy = fx_policy(mapping)
    mode = str(policy.get("mode") or "").strip().lower()
    source_type = "mapping"
    source = fx_source or f"mapping.fxPolicy:{mode or 'fixed'}"
    fx_s = str(fx_source or "")
    if write_source in ("api_as_base", "api") or fx_s.startswith("api:"):
        source_type = "api"
    elif fx_s.startswith("source:") or fx_s.startswith("vendor:") or fx_s.startswith("summary:"):
        # 供应商账单/源表推算，不是在线 API
        source_type = "vendor"
        source = fx_s
    elif "shared_fact" in fx_s:
        source_type = "vendor"
        source = fx_s
    base_key, adj_key, value_key = _fx_keys(policy)
    round_key = str(policy.get("roundDigitsKey") or "").strip()
    detail: dict[str, Any] = {
        "mode": mode,
        "writeSource": write_source,
        # 供核对页「点来源 → 编映射」定位字段
        "editGroup": "fx",
        "baseKey": base_key,
        "adjustmentKey": adj_key,
        "valueKey": value_key,
    }
    if round_key:
        detail["roundDigitsKey"] = round_key
    if fx_source:
        detail["fxSource"] = fx_source
    return {
        "kind": "pnFxWrite",
        "sheet": str(sheet),
        "row": int(row),
        "col": int(col),
        "sourceType": source_type,
        "source": source,
        "label": "FX rate",
        "value": value,
        "detail": detail,
    }


def fixed_fx_override(mapping: dict[str, Any] | None) -> float | None:
    """mode=fixed 时的覆盖值（乘积；无配置则 None）。"""
    _formula, product = fixed_fx_parts(mapping)
    return product
