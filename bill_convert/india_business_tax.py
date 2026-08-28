# -*- coding: utf-8 -*-
"""India Business Tax：PDF CGST+SGST；取整由 mapping 配置；多出来的未知金额行中止。"""
from __future__ import annotations

import re
from decimal import ROUND_HALF_UP, ROUND_UP, Decimal
from typing import Any

DEFAULT_MODE = "ROUND"
DEFAULT_DIGITS = 0

_MONEY_TAIL_RE = re.compile(
    r"(?P<amt>[0-9]{1,2}(?:,[0-9]{2,3})+\.\d{2}|[0-9]+\.\d{2})\s*$"
)
_SKIP_DESC_RE = re.compile(
    r"outsource\s+payroll|monthly\s+ctc|cgst|sgst|igst|\bgst\b|total|taxable|"
    r"round\s*off|invoice|gstin|hsn|sac|quantity|particular|description|"
    r"amount\s+in\s+words|rupees|bank|ifsc|account|\bpan\b|phone|email|"
    r"998224|page\s+\d|bill\s+to|place\s+of\s+supply",
    re.I,
)
_BT_FORMULA_RE = re.compile(
    r"^=\s*(?:(?P<fn>ROUND(?:UP)?)\s*\(\s*)?(?P<a>\d+(?:\.\d+)?)\s*\+\s*(?P<b>\d+(?:\.\d+)?)"
    r"(?:\s*,\s*(?P<d>\d+)\s*\))?\s*$",
    re.I,
)


def parse_india_business_tax_round_cfg(mapping: dict[str, Any] | None) -> dict[str, Any]:
    """mapping.indiaBusinessTaxRoundMode / indiaBusinessTaxRoundDigits。"""
    mode = DEFAULT_MODE
    digits: int | None = DEFAULT_DIGITS
    if isinstance(mapping, dict):
        raw_mode = str(mapping.get("indiaBusinessTaxRoundMode") or "").strip().upper()
        if raw_mode in {"ROUND", "ROUNDUP", "NONE", "RAW", "EXACT"}:
            mode = "NONE" if raw_mode in {"NONE", "RAW", "EXACT"} else raw_mode
        if "indiaBusinessTaxRoundDigits" in mapping:
            raw_d = mapping.get("indiaBusinessTaxRoundDigits")
            if raw_d is None or raw_d == "":
                mode = "NONE"
                digits = None
            else:
                try:
                    digits = int(raw_d)
                except (TypeError, ValueError):
                    digits = DEFAULT_DIGITS
                if digits < 0:
                    mode = "NONE"
                    digits = None
    if mode == "NONE":
        digits = None
    elif digits is None:
        digits = DEFAULT_DIGITS
    return {"mode": mode, "digits": digits}


def _money_text(value: float) -> str:
    text = format(Decimal(str(value)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP), "f")
    if "." in text:
        text = text.rstrip("0").rstrip(".")
    return text or "0"


def format_india_business_tax_formula(
    cgst: float,
    sgst: float,
    *,
    mode: str = DEFAULT_MODE,
    digits: int | None = DEFAULT_DIGITS,
) -> str:
    body = f"{_money_text(cgst)}+{_money_text(sgst)}"
    if mode == "NONE" or digits is None:
        return f"={body}"
    fn = "ROUNDUP" if str(mode).upper() == "ROUNDUP" else "ROUND"
    return f"={fn}({body},{int(digits)})"


def apply_india_business_tax_round(
    cgst: float,
    sgst: float,
    mapping: dict[str, Any] | None = None,
) -> dict[str, Any]:
    cfg = parse_india_business_tax_round_cfg(mapping)
    mode = str(cfg["mode"])
    digits = cfg["digits"]
    raw = Decimal(str(cgst)) + Decimal(str(sgst))
    formula = format_india_business_tax_formula(cgst, sgst, mode=mode, digits=digits)
    if mode == "NONE" or digits is None:
        value = float(raw.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))
    else:
        exp = Decimal("1") if int(digits) == 0 else Decimal("1e-%d" % int(digits))
        rounding = ROUND_UP if mode == "ROUNDUP" else ROUND_HALF_UP
        value = float(raw.quantize(exp, rounding=rounding))
    return {
        "value": value,
        "formula": formula,
        "mode": mode,
        "digits": digits,
        "cgst": float(cgst),
        "sgst": float(sgst),
    }


def parse_business_tax_formula(text: Any) -> dict[str, Any] | None:
    raw = str(text or "").replace(" ", "")
    m = _BT_FORMULA_RE.match(str(text or "").strip()) or _BT_FORMULA_RE.match(raw)
    if not m:
        return None
    cgst = float(m.group("a"))
    sgst = float(m.group("b"))
    fn = (m.group("fn") or "").upper()
    digits = m.group("d")
    mode = "NONE"
    ndigits: int | None = None
    if fn == "ROUNDUP":
        mode = "ROUNDUP"
        ndigits = int(digits) if digits is not None else 0
    elif fn == "ROUND":
        mode = "ROUND"
        ndigits = int(digits) if digits is not None else 0
    return {"cgst": cgst, "sgst": sgst, "mode": mode, "digits": ndigits, "formula": str(text).strip()}


def assert_no_unknown_invoice_amounts(
    text: str,
    *,
    ctc: float | None,
    cgst: float | None,
    sgst: float | None,
) -> None:
    """票面多出对不上 CTC/GST/合计的金额行 → 中止，避免 Expense/Deduction 被写成 0。"""
    known: list[float] = []
    for x in (
        ctc,
        cgst,
        sgst,
        (None if cgst is None or sgst is None else cgst + sgst),
        (
            None
            if ctc is None or cgst is None or sgst is None
            else ctc + cgst + sgst
        ),
    ):
        if x is None:
            continue
        known.append(round(float(x), 2))
        known.append(round(float(x), 0))

    unknown: list[str] = []
    for raw in (text or "").splitlines():
        line = raw.strip()
        if not line or len(line) > 160:
            continue
        m = _MONEY_TAIL_RE.search(line)
        if not m:
            continue
        amt = _parse_money(m.group("amt"))
        if amt is None or amt < 10:
            continue
        if any(abs(amt - k) < 0.051 for k in known):
            continue
        desc = line[: m.start()].strip(" :-|\t")
        if not desc or _SKIP_DESC_RE.search(desc):
            continue
        if re.fullmatch(r"[\d,%.\s]+", desc):
            continue
        unknown.append(f"{desc} = {amt}")
    if unknown:
        preview = "；".join(unknown[:5])
        raise ValueError(
            f"Biz Solutions PDF 出现未识别的费用行（{preview}），"
            f"版式可能已变更；已中止写出以免 Expense Claim / Deduction 被写成 0"
        )


def _parse_money(text: str) -> float | None:
    raw = (text or "").replace(",", "").replace(" ", "").replace("\xa0", "")
    try:
        return round(float(raw), 2)
    except ValueError:
        return None
