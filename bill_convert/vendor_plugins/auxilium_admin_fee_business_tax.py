# -*- coding: utf-8 -*-
"""
Auxilium 专属：Admin Fee 发票 PDF → UAE Business Tax。

公式：BusinessTax = currVat + (currVat - prevVat)
- currVat：factStore.latest（上传 Admin Fee PDF 识别写入）
- prevVat：factStore.total_vat（上期；新 PDF 识别到不同 VAT 时由旧 latest 晋升，或手工期初）

factStore 键：
- auxilium.admin_fee.latest_vat   最新 Admin Fee VAT（curr）
- auxilium.admin_fee.total_vat    上期 VAT（prev）

转换成功不回写 factStore；仅上传/识别 PDF 时由 Office 更新。
"""
from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path
from typing import Any

from bill_convert.fact_store import get_batch_facts, get_fact_value
from pdf_ingest.text_extract import extract_pdf_text

PLUGIN_ID = "auxilium_admin_fee_business_tax"
FACT_KEY_VAT = "auxilium.admin_fee.total_vat"  # committed / prev
FACT_KEY_LATEST_VAT = "auxilium.admin_fee.latest_vat"  # newest upload / curr

_VAT_RE = re.compile(
    r"Total\s+VAT\s*:?\s*(?:5%\s*)?([0-9]{1,3}(?:,[0-9]{3})*(?:\.[0-9]+)?|\d+(?:\.\d+)?)",
    re.IGNORECASE,
)
_INV_NO_RE = re.compile(r"Invoice\s+Number\s*\n?\s*(INV-\d+)", re.IGNORECASE)
_INV_DATE_RE = re.compile(
    r"Invoice\s+Date\s*\n?\s*(\d{1,2}\s+\w+\s+\d{4})",
    re.IGNORECASE,
)
_REF_RE = re.compile(r"Reference\s*\n?\s*([^\n]+)", re.IGNORECASE)


def _parse_money(text: str) -> float | None:
    s = (text or "").replace(",", "").strip()
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _parse_uk_date(text: str) -> str | None:
    raw = (text or "").strip()
    if not raw:
        return None
    for fmt in ("%d %b %Y", "%d %B %Y"):
        try:
            return datetime.strptime(raw, fmt).date().isoformat()
        except ValueError:
            continue
    return None


def looks_like_admin_fee_invoice(path: Path, text: str | None = None) -> bool:
    name = path.name.lower()
    if "admin" in name and "fee" in name:
        return True
    body = (text if text is not None else extract_pdf_text(path)).lower()
    if "total vat" in body and ("admin fee" in body or "management fee" in body):
        return True
    if "auxilium" in body and "total vat" in body and "tax invoice" in body:
        # 发票型 PDF，非 Payroll Draft
        return "payroll draft" not in body
    return False


def parse_admin_fee_invoice(path: Path) -> dict[str, Any]:
    text = extract_pdf_text(path)
    if not looks_like_admin_fee_invoice(path, text):
        raise ValueError(f"不是 Auxilium Admin Fee 发票: {path.name}")
    m_vat = _VAT_RE.search(text.replace("\xa0", " "))
    if not m_vat:
        raise ValueError(f"未找到 Total VAT: {path.name}")
    vat = _parse_money(m_vat.group(1))
    if vat is None:
        raise ValueError(f"Total VAT 无法解析: {path.name}")
    inv_no = None
    m_inv = _INV_NO_RE.search(text)
    if m_inv:
        inv_no = m_inv.group(1).strip()
    inv_date = None
    m_date = _INV_DATE_RE.search(text)
    if m_date:
        inv_date = _parse_uk_date(m_date.group(1))
    period_ref = None
    m_ref = _REF_RE.search(text)
    if m_ref:
        period_ref = m_ref.group(1).strip()
    return {
        FACT_KEY_VAT: vat,
        "auxilium.admin_fee.invoice_no": inv_no,
        "auxilium.admin_fee.invoice_date": inv_date,
        "auxilium.admin_fee.period_ref": period_ref,
        "auxilium.admin_fee.source_file": path.name,
    }


class AuxiliumAdminFeeBusinessTaxPlugin:
    plugin_id = PLUGIN_ID
    pdf_profile_ids = ("auxilium_uae",)

    def classify_path(self, path: Path) -> bool:
        p = Path(path)
        if p.suffix.lower() != ".pdf":
            return False
        try:
            return looks_like_admin_fee_invoice(p)
        except Exception:
            return False

    def parse_artifacts(self, paths: list[Path]) -> dict[str, Any]:
        if not paths:
            return {}
        parsed: list[dict[str, Any]] = []
        for p in paths:
            parsed.append(parse_admin_fee_invoice(Path(p)))
        # 多份时取发票日期最新；无日期则取列表最后一份
        def sort_key(item: dict[str, Any]) -> str:
            return str(item.get("auxilium.admin_fee.invoice_date") or "")

        parsed.sort(key=sort_key)
        chosen = parsed[-1]
        out = dict(chosen)
        if len(parsed) > 1:
            out["_warnings"] = [
                f"本批 {len(parsed)} 份 Admin Fee PDF，已采用最新 "
                f"{chosen.get('auxilium.admin_fee.source_file')} (VAT={chosen.get(FACT_KEY_VAT)})"
            ]
        return out

    def apply_to_workbook(
        self,
        wb,
        *,
        mapping: dict[str, Any],
        batch_facts: dict[str, Any],
        warnings: list[str],
        employee_count: int = 1,
    ) -> dict[str, Any] | None:
        facts = dict(batch_facts or get_batch_facts(mapping) or {})
        # curr：优先 latest（跨批最新发票）；勿把 factStore 已入账的 total_vat 当成 curr
        curr = facts.get(FACT_KEY_LATEST_VAT)
        if curr is None:
            curr = facts.get(FACT_KEY_VAT)  # 解析结果里同值
        if curr is None:
            curr = get_fact_value(mapping, FACT_KEY_LATEST_VAT)
        if curr is None:
            warnings.append(
                "Auxilium：无最新 Admin Fee Total VAT（请先上传 Admin Fee 发票，或与 Draft 同批上传），跳过 Business Tax"
            )
            return None
        try:
            curr_vat = float(curr)
        except (TypeError, ValueError):
            warnings.append(f"Auxilium：Total VAT 无效（{curr}），跳过 Business Tax")
            return None

        # prev：已入账上期（转换成功后写入）；勿与 latest 混用
        prev_raw = get_fact_value(mapping, FACT_KEY_VAT)
        prev_vat: float | None
        try:
            prev_vat = float(prev_raw) if prev_raw is not None and prev_raw != "" else None
        except (TypeError, ValueError):
            prev_vat = None

        if "UAE" not in wb.sheetnames:
            warnings.append("Auxilium：母版缺少 UAE sheet，无法写 Business Tax")
            return None

        uae = wb["UAE"]
        # 写在 Other Fee（A21）正上方，不覆盖结算区
        prev_row, curr_row = 19, 20
        uae.cell(prev_row, 1).value = "Admin VAT (prev)"
        uae.cell(prev_row, 2).value = prev_vat if prev_vat is not None else None
        uae.cell(curr_row, 1).value = "Admin VAT (curr/latest)"
        uae.cell(curr_row, 2).value = curr_vat
        # 清掉曾误写在 F7/F8 的值（母版 E7/E8 标签保留）
        uae.cell(7, 6).value = None
        uae.cell(8, 6).value = None

        # 每人一行都写同一公式（绝对引用）；汇总行只计第一人，避免 SUM 按人数翻倍
        data_start = 9
        n = max(int(employee_count or 1), 1)
        if prev_vat is None:
            warnings.append(
                "Auxilium：无已入账上期 Admin VAT，Business Tax 暂按最新 VAT 写入；"
                "请在映射中设置期初「上期 Total VAT」"
            )
            cell_val: Any = curr_vat
        else:
            cell_val = f"=$B${curr_row}+($B${curr_row}-$B${prev_row})"
        for i in range(n):
            uae.cell(data_start + i, 6).value = cell_val
        for row in range(data_start + n, data_start + 20):
            uae.cell(row, 6).value = None
        uae.cell(6, 6).value = f"=F{data_start}"

        # 不回写 factStore：上期/最新只在上传识别到 Admin Fee PDF 时由 Office 更新
        return None
