# -*- coding: utf-8 -*-
"""
Auxilium 专属：Admin Fee 发票 PDF → UAE Business Tax。

公式：BusinessTax = currVat + (currVat - prevVat)
- currVat：factStore.latest（较小发票号 / 本期）
- prevVat：factStore.total_vat（较大发票号 / 上期；或手工期初）

同批两份 Invoice INV-xxxxx.pdf（文件名不必含 admin fee）：
- 发票号数字大的 = 以前的 Admin Fee（prev）
- 发票号数字小的 = 本期 Admin Fee（curr）
发票号优先读 PDF 正文，其次文件名。

factStore 键：
- auxilium.admin_fee.latest_vat   本期 Admin Fee VAT（curr）
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
_INV_SEQ_RE = re.compile(r"INV-(\d+)", re.IGNORECASE)
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


def invoice_seq(inv_no: str | None, path: Path | None = None) -> int:
    """从发票号或文件名取 INV- 后的数字；读不到返回 -1。"""
    for raw in (inv_no or "", path.name if path is not None else ""):
        m = _INV_SEQ_RE.search(str(raw))
        if m:
            try:
                return int(m.group(1))
            except ValueError:
                continue
    return -1


def _read_pdf_text(path: Path) -> str:
    try:
        return extract_pdf_text(path) or ""
    except Exception:
        return ""


def looks_like_auxilium_invoice(path: Path, text: str | None = None) -> bool:
    """
    Auxilium 税票（Admin Fee 或普通 Invoice INV-xxxxx）。
    不依赖文件名含 admin fee：上传后常被改成 source_0.pdf。
    """
    p = Path(path)
    name = p.name.lower()
    body = (text if text is not None else _read_pdf_text(p)).lower().replace("\xa0", " ")
    if "payroll draft" in body:
        return False
    if "total vat" in body and ("admin fee" in body or "management fee" in body):
        return True
    if "auxilium" in body and "total vat" in body and "tax invoice" in body:
        return True
    if "tax invoice" in body and "total vat" in body and _INV_SEQ_RE.search(body):
        return True
    if "total vat" in body and _INV_SEQ_RE.search(body or name):
        return True
    # 正文抽不出字时：文件名带 INV-数字也视为候选（再在 parse 里抽 VAT）
    if _INV_SEQ_RE.search(p.name) and p.suffix.lower() == ".pdf":
        return True
    if "admin" in name and "fee" in name:
        return True
    return False


def looks_like_admin_fee_invoice(path: Path, text: str | None = None) -> bool:
    return looks_like_auxilium_invoice(path, text)


def parse_admin_fee_invoice(path: Path) -> dict[str, Any]:
    path = Path(path)
    text = _read_pdf_text(path)
    if not looks_like_auxilium_invoice(path, text):
        raise ValueError(f"不是 Auxilium 发票: {path.name}")
    body = text.replace("\xa0", " ")
    m_vat = _VAT_RE.search(body)
    if not m_vat:
        raise ValueError(f"未找到 Total VAT: {path.name}")
    vat = _parse_money(m_vat.group(1))
    if vat is None:
        raise ValueError(f"Total VAT 无法解析: {path.name}")
    inv_no = None
    m_inv = _INV_NO_RE.search(text)
    if m_inv:
        inv_no = m_inv.group(1).strip()
    if not inv_no:
        m_name = _INV_SEQ_RE.search(path.name)
        if m_name:
            inv_no = f"INV-{m_name.group(1)}"
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
        FACT_KEY_LATEST_VAT: vat,
        "auxilium.admin_fee.invoice_no": inv_no,
        "auxilium.admin_fee.invoice_date": inv_date,
        "auxilium.admin_fee.period_ref": period_ref,
        "auxilium.admin_fee.source_file": path.name,
        "auxilium.admin_fee.invoice_seq": invoice_seq(inv_no, path),
    }


def _facts_from_parsed(item: dict[str, Any], *, as_prev: bool = False) -> dict[str, Any]:
    vat = item.get(FACT_KEY_VAT)
    out = {
        FACT_KEY_VAT: vat,
        FACT_KEY_LATEST_VAT: vat,
        "auxilium.admin_fee.invoice_no": item.get("auxilium.admin_fee.invoice_no"),
        "auxilium.admin_fee.invoice_date": item.get("auxilium.admin_fee.invoice_date"),
        "auxilium.admin_fee.period_ref": item.get("auxilium.admin_fee.period_ref"),
        "auxilium.admin_fee.source_file": item.get("auxilium.admin_fee.source_file"),
        "auxilium.admin_fee.invoice_seq": item.get("auxilium.admin_fee.invoice_seq"),
    }
    if as_prev:
        out["auxilium.admin_fee.prev_invoice_no"] = item.get("auxilium.admin_fee.invoice_no")
        out["auxilium.admin_fee.prev_source_file"] = item.get("auxilium.admin_fee.source_file")
    return out


class AuxiliumAdminFeeBusinessTaxPlugin:
    plugin_id = PLUGIN_ID
    pdf_profile_ids = ("auxilium_uae",)

    def classify_path(self, path: Path) -> bool:
        p = Path(path)
        if p.suffix.lower() != ".pdf":
            return False
        try:
            return looks_like_auxilium_invoice(p)
        except Exception:
            return bool(_INV_SEQ_RE.search(p.name))

    def parse_artifacts(self, paths: list[Path]) -> dict[str, Any]:
        if not paths:
            return {}
        parsed: list[dict[str, Any]] = []
        for p in paths:
            parsed.append(parse_admin_fee_invoice(Path(p)))

        def sort_key(item: dict[str, Any]) -> tuple[int, str]:
            seq = item.get("auxilium.admin_fee.invoice_seq")
            try:
                n = int(seq)
            except (TypeError, ValueError):
                n = -1
            return (n, str(item.get("auxilium.admin_fee.invoice_date") or ""))

        parsed.sort(key=sort_key)
        if len(parsed) == 1:
            return _facts_from_parsed(parsed[0])

        # 同批多份：发票号数字大的 = 上期 Admin Fee；数字小的 = 本期
        curr = parsed[0]
        prev = parsed[-1]
        out = _facts_from_parsed(curr)
        out[FACT_KEY_LATEST_VAT] = curr.get(FACT_KEY_VAT)
        out[FACT_KEY_VAT] = prev.get(FACT_KEY_VAT)
        out["auxilium.admin_fee.prev_invoice_no"] = prev.get("auxilium.admin_fee.invoice_no")
        out["auxilium.admin_fee.prev_source_file"] = prev.get("auxilium.admin_fee.source_file")
        out["auxilium.admin_fee.prev_invoice_date"] = prev.get("auxilium.admin_fee.invoice_date")
        warnings = [
            "同批 Admin Fee 按发票号分流："
            f"上期(较大号) {prev.get('auxilium.admin_fee.invoice_no') or prev.get('auxilium.admin_fee.source_file')}"
            f" VAT={prev.get(FACT_KEY_VAT)}；"
            f"本期(较小号) {curr.get('auxilium.admin_fee.invoice_no') or curr.get('auxilium.admin_fee.source_file')}"
            f" VAT={curr.get(FACT_KEY_VAT)}"
        ]
        if len(parsed) > 2:
            warnings.append(f"本批 {len(parsed)} 份发票，仅采用号最小与号最大两份")
        out["_warnings"] = warnings
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
        batch_latest = facts.get(FACT_KEY_LATEST_VAT)
        batch_prev = facts.get(FACT_KEY_VAT)
        pair = bool(
            facts.get("auxilium.admin_fee.prev_invoice_no")
            or (
                batch_latest is not None
                and batch_prev is not None
                and str(batch_latest).strip() != ""
                and str(batch_prev).strip() != ""
                and str(batch_latest) != str(batch_prev)
            )
        )

        curr = batch_latest
        if curr is None:
            curr = get_fact_value(mapping, FACT_KEY_LATEST_VAT)
        if curr is None and not pair:
            curr = batch_prev
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

        if pair:
            prev_raw = batch_prev
        else:
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
