# -*- coding: utf-8 -*-
"""从 exchangerate-api 获取 USD 基准汇率。"""
from __future__ import annotations

import json
import urllib.error
import urllib.request

FX_API_URL = "https://api.exchangerate-api.com/v4/latest/USD"

# HK 母版原公式 =7.81039098*0.97，保留 0.97 调整系数
HK_FX_ADJUSTMENT = 0.97


def fetch_usd_rates(timeout: float = 15.0) -> dict[str, float]:
    try:
        with urllib.request.urlopen(FX_API_URL, timeout=timeout) as resp:
            data = json.loads(resp.read().decode("utf-8"))
    except urllib.error.URLError as exc:
        raise RuntimeError(f"获取汇率失败: {exc}") from exc

    rates = data.get("rates")
    if not isinstance(rates, dict) or not rates:
        raise RuntimeError("汇率 API 返回格式异常")
    return {str(k).upper(): float(v) for k, v in rates.items()}


def get_usd_rate(currency: str, rates: dict[str, float] | None = None) -> float:
    code = currency.strip().upper()
    table = rates if rates is not None else fetch_usd_rates()
    if code not in table:
        raise RuntimeError(f"汇率 API 未返回 {code}")
    return table[code]


def get_china_pn_fx_rate(rates: dict[str, float] | None = None) -> float:
    """China PN!B29：USD → CNY"""
    return get_usd_rate("CNY", rates)


def get_hk_pn_fx_rate(rates: dict[str, float] | None = None) -> float:
    """Hong Kong PN!B28：USD → HKD，再乘调整系数 0.97"""
    return round(get_usd_rate("HKD", rates) * HK_FX_ADJUSTMENT, 8)


def get_tw_pn_fx_rate(rates: dict[str, float] | None = None) -> float:
    """Taiwan PN!B31：USD → TWD"""
    return round(get_usd_rate("TWD", rates), 4)


def get_uk_gbp_per_usd(rates: dict[str, float] | None = None) -> float:
    """UK-L!D24：1 GBP = ? USD（母版用金额 GBP × D24 → USD；PN!B29=1/D24）。"""
    gbp_per_usd = get_usd_rate("GBP", rates)
    if gbp_per_usd <= 0:
        raise RuntimeError("GBP 汇率无效")
    return round(1.0 / gbp_per_usd, 6)


def get_uae_pn_fx_rate(rates: dict[str, float] | None = None) -> float:
    """UAE PN!B28：USD → AED（金额 AED / B28 → USD）。"""
    return round(get_usd_rate("AED", rates), 6)