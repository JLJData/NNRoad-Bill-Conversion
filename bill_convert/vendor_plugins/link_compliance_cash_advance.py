# -*- coding: utf-8 -*-
"""
Link Compliance 专属：Tax Invoice PDF → PN Cash Advance（仅此一项）。

正文样例：
  3 Cash Advance
  - Paizal Nafis: IDR 40,165,612.50
  2,350.00

写入：
  PN 描述列「- Cash Advance for {name}」
  PN 金额列 IDR 数值（USD 列保留母版公式，如 =E18/$B$28）

fact 键：
  link_compliance.cash_advance.amount_idr
  link_compliance.cash_advance.employee_name
"""
from __future__ import annotations

import re
from pathlib import Path
from typing import Any

from bill_convert.fact_store import get_batch_facts, get_fact_value
from pdf_ingest.text_extract import extract_pdf_text

PLUGIN_ID = "link_compliance_cash_advance"
PROVENANCE_SOURCE = f"plugin:{PLUGIN_ID}"

FACT_KEY_AMOUNT = "link_compliance.cash_advance.amount_idr"
FACT_KEY_NAME = "link_compliance.cash_advance.employee_name"
FACT_KEY_SOURCE = "link_compliance.cash_advance.source_file"

PN_SHEET = "PN"
# 母版：Labor / Expense 之后、Service Fee 之前的空行（样例为第 18 行）
_DEFAULT_PN_ROW = 18
_DESC_COL = 1
_AMOUNT_COL = 5

_CASH_ADVANCE_BLOCK_RE = re.compile(
    r"Cash\s+Advance\s*"
    r"(?:\n|\r\n?|\s)+"
    r"-\s*([^:\n\r]+?)\s*:\s*IDR\s*([\d,]+(?:\.\d+)?)",
    re.IGNORECASE,
)
_CASH_ADVANCE_INLINE_RE = re.compile(
    r"Cash\s+Advance[^\n\r]{0,80}?"
    r"IDR\s*([\d,]+(?:\.\d+)?)",
    re.IGNORECASE,
)


def _parse_money(text: str) -> float | None:
    s = (text or "").replace(",", "").strip()
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _read_pdf_text(path: Path) -> str:
    try:
        return extract_pdf_text(path) or ""
    except Exception:
        return ""


def looks_like_link_compliance_invoice(path: Path, text: str | None = None) -> bool:
    """Link Compliance Tax Invoice（含 Cash Advance 旁路）。"""
    p = Path(path)
    if p.suffix.lower() != ".pdf":
        return False
    body = (text if text is not None else _read_pdf_text(p)).lower().replace("\xa0", " ")
    if not body.strip():
        name = p.name.lower()
        return "lcsg" in name or "link" in name and "compliance" in name
    if "link compliance" in body:
        return True
    if "cash advance" in body and ("tax invoice" in body or "nnroad" in body):
        return True
    return False


def parse_cash_advance_pdf(path: Path) -> dict[str, Any]:
    path = Path(path)
    text = _read_pdf_text(path)
    if not looks_like_link_compliance_invoice(path, text):
        raise ValueError(f"不是 Link Compliance Tax Invoice: {path.name}")
    body = text.replace("\xa0", " ")
    name: str | None = None
    amount: float | None = None
    m = _CASH_ADVANCE_BLOCK_RE.search(body)
    if m:
        name = (m.group(1) or "").strip() or None
        amount = _parse_money(m.group(2))
    if amount is None:
        m2 = _CASH_ADVANCE_INLINE_RE.search(body)
        if m2:
            amount = _parse_money(m2.group(1))
    if amount is None:
        raise ValueError(f"未找到 Cash Advance 金额: {path.name}")
    return {
        FACT_KEY_AMOUNT: amount,
        FACT_KEY_NAME: name,
        FACT_KEY_SOURCE: path.name,
    }


def _resolve_pn_cash_advance_row(ws, mapping: dict[str, Any] | None) -> int:
    """优先 mapping.cashAdvancePnRow；否则找已有 Cash Advance / 预留公式行；再否则默认 18。

    注意：母版 E15=SUM(E16:E18)，Cash Advance 必须落在 16–18 槽位内（样例为 18）。
    """
    if isinstance(mapping, dict):
        raw = mapping.get("cashAdvancePnRow")
        try:
            if raw is not None and int(raw) > 0:
                return int(raw)
        except (TypeError, ValueError):
            pass
    for row in range(1, (ws.max_row or 40) + 1):
        v = ws.cell(row, _DESC_COL).value
        if isinstance(v, str) and "cash advance" in v.lower():
            return row
    # 预留槽：F 列已有 =E{row}/$B$28，且描述/金额为空（模板第 18 行）
    for row in range(16, 19):
        fval = ws.cell(row, 6).value
        desc = ws.cell(row, _DESC_COL).value
        amt = ws.cell(row, _AMOUNT_COL).value
        if (
            isinstance(fval, str)
            and fval.startswith("=")
            and f"E{row}" in fval.replace("$", "")
            and (desc is None or str(desc).strip() == "")
            and (amt is None or str(amt).strip() == "")
        ):
            return row
    return _DEFAULT_PN_ROW


def apply_cash_advance_to_pn(
    wb,
    *,
    amount_idr: float,
    employee_name: str | None = None,
    mapping: dict[str, Any] | None = None,
) -> list[dict[str, Any]]:
    if PN_SHEET not in wb.sheetnames:
        return []
    ws = wb[PN_SHEET]
    row = _resolve_pn_cash_advance_row(ws, mapping)
    name = (employee_name or "").strip() or "Employee"
    desc = f"- Cash Advance for {name}"
    ws.cell(row, _DESC_COL).value = desc
    ws.cell(row, _AMOUNT_COL).value = float(amount_idr)
    return [
        {
            "sheet": PN_SHEET,
            "row": row,
            "col": _DESC_COL,
            "value": desc,
            "source": PROVENANCE_SOURCE,
            "field": "cashAdvanceDescription",
        },
        {
            "sheet": PN_SHEET,
            "row": row,
            "col": _AMOUNT_COL,
            "value": float(amount_idr),
            "source": PROVENANCE_SOURCE,
            "field": "cashAdvanceAmountIdr",
        },
    ]


class LinkComplianceCashAdvancePlugin:
    plugin_id = PLUGIN_ID
    pdf_profile_ids = ("link_compliance_id",)

    def classify_path(self, path: Path) -> bool:
        p = Path(path)
        if p.suffix.lower() != ".pdf":
            return False
        try:
            return looks_like_link_compliance_invoice(p)
        except Exception:
            return False

    def parse_artifacts(self, paths: list[Path]) -> dict[str, Any]:
        if not paths:
            return {}
        # 同批多份时取最后一份（通常最新）
        last = Path(paths[-1])
        parsed = parse_cash_advance_pdf(last)
        warnings: list[str] = []
        if len(paths) > 1:
            warnings.append(
                f"本批 {len(paths)} 份 Link Compliance PDF，Cash Advance 采用 {last.name}"
            )
        if warnings:
            parsed["_warnings"] = warnings
        return parsed

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
        amount = facts.get(FACT_KEY_AMOUNT)
        if amount is None:
            amount = get_fact_value(mapping, FACT_KEY_AMOUNT)
        if amount is None:
            warnings.append("Link Compliance：无 Cash Advance 金额（请附带 Tax Invoice PDF），跳过写入 PN")
            return None
        try:
            amount_f = float(amount)
        except (TypeError, ValueError):
            warnings.append(f"Link Compliance：Cash Advance 金额无效（{amount}），跳过")
            return None
        if amount_f <= 0:
            warnings.append("Link Compliance：Cash Advance 金额 ≤ 0，跳过")
            return None

        name = facts.get(FACT_KEY_NAME)
        if name is None:
            name = get_fact_value(mapping, FACT_KEY_NAME)
        # 回退：Indonesia-L 第一人姓名
        if not name and "Indonesia-L" in wb.sheetnames:
            try:
                name = wb["Indonesia-L"].cell(8, 2).value
            except Exception:
                name = None

        cell_writes = apply_cash_advance_to_pn(
            wb,
            amount_idr=amount_f,
            employee_name=str(name) if name else None,
            mapping=mapping,
        )
        return {"_cell_writes": cell_writes}
