# -*- coding: utf-8 -*-
"""
EOR Services Limited（UK）：
  - 文字发票 PDF → 填 UK PN 母版的 UK-L
  - EMPANPAY / Analysis of Payroll Totals Excel（B 标签 / D 金额）→ 同写 UK-L

PDF 样例（Invoice #6108）:
  Sarah Jane Walker-Monthly  3,634.70
    Monthly Salary- £3120, Empr Taxes £421.10, Empr Contributions £93.60
  Service Fee  200.00
  Invoice Total  £3,834.70

Excel 样例（EMPANPAY.xlsx）:
  Gross Pay / Employer NI / Employer Pension / PAYE Tax / Employee NI / Employee Pension …

映射兼容：
  - 引擎内置默认标签别名（零配置）
  - mappingJson.columnRename 覆盖（Office「列名对照」：供应商标签 → UK-L 标准标签）
  - sourceEmployeeSheet.labelColumn/amountColumn 控制竖表列（默认自动探测 A/B 或 B/D）

Service Fee / Recurring Fee(UK!H) 不是同一项 → 不自动写入 H 列（见 ERI 映射卡片）。
"""
from __future__ import annotations

import argparse
import re
import shutil
import sys
from dataclasses import asdict, dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any

from openpyxl import load_workbook

from fx_rate import fetch_usd_rates, get_uk_gbp_per_usd
from pdf_ingest.text_extract import extract_pdf_text
from pn_meta import PnMeta, apply_pn_meta
from region_templates import get_region_template

UK_L_SHEET = "UK-L"
UK_SHEET = "UK"

# UK-L 标签 → 单元格（金额在 B 列）
_LABEL_CELLS: dict[str, str] = {
    "Gross Salary": "B7",
    "Holiday Pay": "B8",
    "Car Allowance": "B9",
    "ER' NIC": "B10",
    "ER' Pension (Auto Enrolment)": "B11",
    "PAYE (Estimated)": "B14",
    "EE'NIC": "B15",
    "EE' Pension (Auto Enrolment)": "B16",
    "App Levy": "B22",
    "Payment Fees": "B26",
}

def _build_default_label_aliases() -> dict[str, str]:
    """小写源标签 → UK-L；含 convert_mapping 展示默认 + PDF/口语别名；可被 columnRename 覆盖。"""
    from convert_mapping import get_builtin_column_rename

    out: dict[str, str] = {}
    for src, dst in get_builtin_column_rename("eor_uk").items():
        out[str(src).strip().lower()] = str(dst).strip()
    # PDF / 口语等额外别名（不单独进 Office 默认表，避免噪声）
    out.update(
        {
            "gross salary": "Gross Salary",
            "monthly salary": "Gross Salary",
            "holiday pay": "Holiday Pay",
            "car allowance": "Car Allowance",
            "er' nic": "ER' NIC",
            "er nic": "ER' NIC",
            "empr taxes": "ER' NIC",
            "employer taxes": "ER' NIC",
            "er' pension (auto enrolment)": "ER' Pension (Auto Enrolment)",
            "er' pension": "ER' Pension (Auto Enrolment)",
            "empr contributions": "ER' Pension (Auto Enrolment)",
            "employer contributions": "ER' Pension (Auto Enrolment)",
            "paye (estimated)": "PAYE (Estimated)",
            "paye": "PAYE (Estimated)",
            "ee'nic": "EE'NIC",
            "ee' nic": "EE'NIC",
            "ee nic": "EE'NIC",
            "ee' pension (auto enrolment)": "EE' Pension (Auto Enrolment)",
            "ee' pension": "EE' Pension (Auto Enrolment)",
            "app levy": "App Levy",
            "payment fees": "Payment Fees",
        }
    )
    return out


_DEFAULT_LABEL_ALIASES: dict[str, str] = _build_default_label_aliases()

_SKIP_LABELS = {
    "description",
    "this period",
    "net pay",
    "standard earnings for ni",
    "employee ni rebate",
    "employer ni rebate",
    "employer class 1a",
    "loan repayments",
    "holiday fund accrued",
    "student/postgraduate loan",
    "details",
    "amount in gbp",
}


@dataclass
class EorUkParsed:
    supplier: str = "EOR Services Limited"
    invoice_no: str | None = None
    invoice_date: date | None = None
    employee_name: str | None = None
    gross_salary: float | None = None
    holiday_pay: float | None = None
    er_nic: float | None = None
    er_pension: float | None = None
    paye: float | None = None
    ee_nic: float | None = None
    ee_pension: float | None = None
    service_fee: float | None = None
    invoice_total: float | None = None
    currency: str = "GBP"
    raw_line_net: float | None = None  # 工资行 Net Amt
    source_kind: str = "pdf"  # pdf | excel_payroll_totals | excel_uk_l
    # 归一化后的 UK-L 标签 → 金额（Excel 路径主用）
    amounts: dict[str, float] = field(default_factory=dict)
    warnings: list[str] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        d = asdict(self)
        if self.invoice_date is not None:
            d["invoice_date"] = self.invoice_date.isoformat()
        return d


def _money(s: str) -> float | None:
    s = (s or "").strip()
    if not s:
        return None
    s = s.replace(",", "").replace("£", "").replace("￡", "").strip()
    try:
        return float(s)
    except ValueError:
        return None


def _uk_tax_year_label(d: date) -> str:
    """英国财年标签：4 月 6 日起算，如 2026-05-05 → 26-27。"""
    if d.month > 4 or (d.month == 4 and d.day >= 6):
        start = d.year
    else:
        start = d.year - 1
    return f"{str(start)[-2:]}-{str(start + 1)[-2:]}"


def parse_eor_uk_text(text: str) -> EorUkParsed:
    out = EorUkParsed()
    t = text.replace("\r", "\n")
    compact = re.sub(r"[ \t]+", " ", t)

    m = re.search(r"Invoice\s*No\s*[\n\r\s]*([0-9]+)", compact, re.I)
    if m:
        out.invoice_no = m.group(1).strip()

    m = re.search(
        r"Invoice\s*Date\s*[\n\r\s]*([0-9]{1,2}[/-][0-9]{1,2}[/-][0-9]{2,4})",
        compact,
        re.I,
    )
    if m:
        raw = m.group(1).replace("-", "/")
        for fmt in ("%d/%m/%Y", "%d/%m/%y", "%m/%d/%Y"):
            try:
                out.invoice_date = datetime.strptime(raw, fmt).date()
                break
            except ValueError:
                continue
        if out.invoice_date is None:
            out.warnings.append(f"无法解析发票日期: {raw}")

    # 员工名：描述行「Name-Monthly」
    m = re.search(
        r"([A-Za-z][A-Za-z .'-]{1,80}?)-Monthly",
        compact,
        re.I,
    )
    if m:
        out.employee_name = re.sub(r"\s+", " ", m.group(1)).strip(" -")

    # Salary / Empr Taxes / Contributions
    m = re.search(
        r"Monthly\s+Salary\s*[-–—]?\s*£?\s*([0-9,]+\.?[0-9]*)\s*,?\s*"
        r"Empr\s+Taxes?\s*£?\s*([0-9,]+\.?[0-9]*)\s*,?\s*"
        r"Empr\s+Contributions?\s*£?\s*([0-9,]+\.?[0-9]*)",
        compact,
        re.I,
    )
    if m:
        out.gross_salary = _money(m.group(1))
        out.er_nic = _money(m.group(2))
        out.er_pension = _money(m.group(3))
    else:
        out.warnings.append(
            "未匹配到「Monthly Salary / Empr Taxes / Empr Contributions」描述，版式可能已变更"
        )

    m = re.search(
        r"Service\s+Fee\s+([0-9,]+\.?[0-9]*)",
        compact,
        re.I,
    )
    if m:
        out.service_fee = _money(m.group(1))

    m = re.search(
        r"Invoice\s+Total\s*£?\s*([0-9,]+\.?[0-9]*)",
        compact,
        re.I,
    )
    if m:
        out.invoice_total = _money(m.group(1))

    # 工资行净额（描述后第一个金额块，作勾稽）
    m = re.search(
        r"-Monthly\s+([0-9,]+\.?[0-9]*)",
        compact,
        re.I,
    )
    if m:
        out.raw_line_net = _money(m.group(1))

    # 勾稽：Salary+NIC+Pension ≈ 工资行；+ServiceFee ≈ Total
    if (
        out.gross_salary is not None
        and out.er_nic is not None
        and out.er_pension is not None
        and out.raw_line_net is not None
    ):
        labor = round(out.gross_salary + out.er_nic + out.er_pension, 2)
        if abs(labor - out.raw_line_net) > 0.05:
            out.warnings.append(
                f"工资构成合计 {labor} 与行净额 {out.raw_line_net} 不一致"
            )
    if (
        out.raw_line_net is not None
        and out.service_fee is not None
        and out.invoice_total is not None
    ):
        expect = round(out.raw_line_net + out.service_fee, 2)
        if abs(expect - out.invoice_total) > 0.05:
            out.warnings.append(
                f"行净额+ServiceFee={expect} 与 Invoice Total {out.invoice_total} 不一致"
            )

    if not out.employee_name:
        out.warnings.append("未解析到员工姓名")
    if out.gross_salary is None:
        out.warnings.append("未解析到 Gross Salary")

    # 关键字段缺失 → 直接失败，避免写出偏数 UK-L
    if not out.employee_name:
        raise ValueError("EOR UK PDF 未解析到员工姓名，版式可能已变更")
    if out.gross_salary is None or out.er_nic is None or out.er_pension is None:
        raise ValueError(
            "EOR UK PDF 未解析到 Gross Salary / Empr Taxes / Empr Contributions，版式可能已变更"
        )
    if (
        out.raw_line_net is not None
        and out.gross_salary is not None
        and out.er_nic is not None
        and out.er_pension is not None
    ):
        labor = round(out.gross_salary + out.er_nic + out.er_pension, 2)
        if abs(labor - out.raw_line_net) > 0.05:
            raise ValueError(
                f"EOR UK PDF 工资构成合计 {labor} 与行净额 {out.raw_line_net} 不一致，已中止写出"
            )
    if (
        out.raw_line_net is not None
        and out.service_fee is not None
        and out.invoice_total is not None
    ):
        expect = round(out.raw_line_net + out.service_fee, 2)
        if abs(expect - out.invoice_total) > 0.05:
            raise ValueError(
                f"EOR UK PDF 行净额+ServiceFee={expect} 与 Invoice Total {out.invoice_total} 不一致，已中止写出"
            )

    return out


def parse_eor_uk_pdf(pdf_path: Path) -> EorUkParsed:
    text = extract_pdf_text(pdf_path)
    parsed = parse_eor_uk_text(text)
    # 版式指纹
    low = text.lower()
    if "eor services limited" not in low and "eorservices.co.uk" not in low:
        parsed.warnings.append(
            "正文未出现 EOR Services 关键字，可能不是本 profile 对应的发票"
        )
    return parsed


def _set_amount_by_label(ws, label: str, value: float | None) -> None:
    addr = _LABEL_CELLS.get(label)
    if not addr or value is None:
        return
    ws[addr] = float(value)


def _norm_label(v: Any) -> str:
    return str(v or "").strip()


def _resolve_eor_mapping(convert_mapping: dict[str, Any] | None) -> dict[str, Any]:
    from convert_mapping import resolve_convert_mapping

    raw = dict(convert_mapping) if isinstance(convert_mapping, dict) else {}
    raw.setdefault("pdfProfileId", "eor_uk")
    return resolve_convert_mapping("uk_payroll_calc", raw)


def _merge_label_aliases(column_rename: dict[str, Any] | None) -> dict[str, str]:
    """小写源标签 → UK-L 标准标签；columnRename 覆盖默认别名。"""
    out = dict(_DEFAULT_LABEL_ALIASES)
    if isinstance(column_rename, dict):
        for k, v in column_rename.items():
            src = _norm_label(k)
            dst = _norm_label(v)
            if src and dst:
                out[src.lower()] = dst
    return out


def _canonicalize_label(raw: str, aliases: dict[str, str]) -> str | None:
    key = _norm_label(raw)
    if not key:
        return None
    low = key.lower()
    if low in _SKIP_LABELS:
        return None
    if low in aliases:
        return aliases[low]
    # 已是 UK-L 标准标签
    for std in _LABEL_CELLS:
        if std.lower() == low:
            return std
    return None


def _cell_amount(v: Any) -> float | None:
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    return _money(str(v))


def _score_label_amount_cols(ws, label_col: int, amount_col: int) -> tuple[int, list[str]]:
    """按可识别标签数打分，供 A/B 与 B/D 自动探测。"""
    aliases = _merge_label_aliases(None)
    hit = 0
    labels: list[str] = []
    max_r = min(int(ws.max_row or 1), 80)
    for r in range(1, max_r + 1):
        raw = _norm_label(ws.cell(r, label_col).value)
        if not raw:
            continue
        canon = _canonicalize_label(raw, aliases)
        if canon is None:
            continue
        amt = _cell_amount(ws.cell(r, amount_col).value)
        if amt is None:
            continue
        hit += 1
        labels.append(raw)
    return hit, labels


def detect_uk_vertical_columns(
    ws,
    *,
    preferred_label: int | None = None,
    preferred_amount: int | None = None,
) -> tuple[int, int]:
    """
    探测竖表标签/金额列。
    优先 mapping 配置；再试 (1,2) UK-L / TopSource 与 (2,4) EMPANPAY。
    """
    candidates: list[tuple[int, int]] = []
    if preferred_label and preferred_amount:
        candidates.append((int(preferred_label), int(preferred_amount)))
    for pair in ((1, 2), (2, 4), (1, 3), (2, 3)):
        if pair not in candidates:
            candidates.append(pair)

    best = candidates[0]
    best_score = -1
    for pair in candidates:
        score, _ = _score_label_amount_cols(ws, pair[0], pair[1])
        if score > best_score:
            best_score = score
            best = pair
    return best


def _pick_payroll_sheet(wb, mapping: dict[str, Any] | None = None):
    from convert_mapping import find_sheet_name

    src_spec = (
        mapping.get("sourceEmployeeSheet")
        if isinstance(mapping, dict) and isinstance(mapping.get("sourceEmployeeSheet"), dict)
        else {}
    )
    name = find_sheet_name(list(wb.sheetnames), src_spec if src_spec else None)
    if name:
        return wb[name]
    for n in wb.sheetnames:
        low = str(n).strip().lower()
        if "payroll totals" in low or "analysis of payroll" in low:
            return wb[n]
    for n in wb.sheetnames:
        if str(n).strip().upper().startswith("UK-L"):
            return wb[n]
    return wb[wb.sheetnames[0]] if wb.sheetnames else None


def _looks_like_uk_l_workbook(path: Path) -> bool:
    wb = load_workbook(path, read_only=True, data_only=False)
    try:
        return any(str(n).strip().upper().startswith("UK-L") for n in wb.sheetnames)
    finally:
        wb.close()


def looks_like_eor_payroll_totals_excel(path: Path) -> bool:
    """内容探测：Analysis of Payroll Totals 或 Gross Pay + Employer NI。"""
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        names = [str(n) for n in wb.sheetnames]
        if any("payroll totals" in n.lower() for n in names):
            return True
        ws = wb[names[0]] if names else None
        if ws is None:
            return False
        blob = " ".join(
            _norm_label(ws.cell(r, c).value).lower()
            for r in range(1, min(20, int(ws.max_row or 1)) + 1)
            for c in range(1, min(6, int(ws.max_column or 1)) + 1)
        )
        return "gross pay" in blob and ("employer ni" in blob or "employer pension" in blob)
    except Exception:
        return False
    finally:
        wb.close()


def parse_eor_uk_payroll_excel(
    excel_path: Path,
    *,
    convert_mapping: dict[str, Any] | None = None,
) -> EorUkParsed:
    """EMPANPAY / Analysis of Payroll Totals 竖表 → EorUkParsed。"""
    mapping = _resolve_eor_mapping(convert_mapping)
    src_spec = (
        mapping.get("sourceEmployeeSheet")
        if isinstance(mapping.get("sourceEmployeeSheet"), dict)
        else {}
    )
    rename = mapping.get("columnRename") if isinstance(mapping.get("columnRename"), dict) else {}
    aliases = _merge_label_aliases(rename)

    out = EorUkParsed(source_kind="excel_payroll_totals", currency="GBP")
    wb = load_workbook(excel_path, data_only=True)
    try:
        ws = _pick_payroll_sheet(wb, mapping)
        if ws is None:
            raise ValueError("Excel 无可用工作表")

        label_col, amount_col = detect_uk_vertical_columns(
            ws,
            preferred_label=int(src_spec.get("labelColumn") or 0) or None,
            preferred_amount=int(src_spec.get("amountColumn") or 0) or None,
        )
        out.warnings.append(
            f"竖表列探测：标签列={label_col} 金额列={amount_col}"
            + (f"；columnRename={len(rename)} 条" if rename else "；使用内置默认别名")
        )

        amounts: dict[str, float] = {}
        unmapped: list[str] = []
        max_r = min(int(ws.max_row or 1), 120)
        for r in range(1, max_r + 1):
            raw = _norm_label(ws.cell(r, label_col).value)
            if not raw:
                continue
            low = raw.lower()
            if low in _SKIP_LABELS:
                continue
            amt = _cell_amount(ws.cell(r, amount_col).value)
            if amt is None:
                continue
            canon = _canonicalize_label(raw, aliases)
            if canon is None:
                unmapped.append(raw)
                continue
            # 同标准标签多源时后者覆盖（通常一行一项）
            amounts[canon] = float(amt)

        out.amounts = amounts
        out.gross_salary = amounts.get("Gross Salary")
        out.holiday_pay = amounts.get("Holiday Pay")
        out.er_nic = amounts.get("ER' NIC")
        out.er_pension = amounts.get("ER' Pension (Auto Enrolment)")
        out.paye = amounts.get("PAYE (Estimated)")
        out.ee_nic = amounts.get("EE'NIC")
        out.ee_pension = amounts.get("EE' Pension (Auto Enrolment)")

        if unmapped:
            out.warnings.append(
                "未映射标签（可在 Office「列名对照」补充）: " + ", ".join(unmapped[:8])
                + ("…" if len(unmapped) > 8 else "")
            )

        # EMPANPAY 无姓名；不写「未解析到员工姓名」以免 post_checks 升致命
        out.warnings.append(
            "Excel 竖表通常无员工姓名，UK-L 标题暂用 Employee；请在 PN 元数据补全"
        )

        if out.gross_salary is None:
            raise ValueError(
                "EOR UK Excel 未解析到 Gross Salary（Gross Pay），请检查列映射或版式"
            )
        if out.er_nic is None or out.er_pension is None:
            out.warnings.append(
                "未完整解析 ER NIC / ER Pension，请核对列名对照（Employer NI / Employer Pension）"
            )
    finally:
        wb.close()
    return out


def apply_to_workbook(
    wb,
    parsed: EorUkParsed,
    *,
    pn_meta: PnMeta | dict[str, Any] | None = None,
    registry_dir: Path | None = None,
    fill_fx: bool = True,
    convert_mapping: dict[str, Any] | None = None,
) -> tuple[list[str], float | None, PnMeta | None]:
    warnings = list(parsed.warnings)
    if UK_L_SHEET not in wb.sheetnames:
        raise ValueError(f"母版缺少 sheet「{UK_L_SHEET}」，现有: {wb.sheetnames}")

    ws = wb[UK_L_SHEET]
    name = parsed.employee_name or "Employee"
    fy = _uk_tax_year_label(parsed.invoice_date) if parsed.invoice_date else "YY-YY"
    ws["A3"] = f"{name}- Salary Calculation for FY {fy}"

    amounts = dict(parsed.amounts or {})
    if not amounts:
        if parsed.gross_salary is not None:
            amounts["Gross Salary"] = float(parsed.gross_salary)
        if parsed.holiday_pay is not None:
            amounts["Holiday Pay"] = float(parsed.holiday_pay)
        if parsed.er_nic is not None:
            amounts["ER' NIC"] = float(parsed.er_nic)
        if parsed.er_pension is not None:
            amounts["ER' Pension (Auto Enrolment)"] = float(parsed.er_pension)
        if parsed.paye is not None:
            amounts["PAYE (Estimated)"] = float(parsed.paye)
        if parsed.ee_nic is not None:
            amounts["EE'NIC"] = float(parsed.ee_nic)
        if parsed.ee_pension is not None:
            amounts["EE' Pension (Auto Enrolment)"] = float(parsed.ee_pension)

    is_excel = str(parsed.source_kind or "").startswith("excel")
    is_merged = str(parsed.source_kind or "") == "pdf_excel_merged"
    if not is_excel and not is_merged:
        # 纯 PDF：缺省项置 0
        for label in (
            "Holiday Pay",
            "Car Allowance",
            "PAYE (Estimated)",
            "EE'NIC",
            "EE' Pension (Auto Enrolment)",
            "App Levy",
        ):
            amounts.setdefault(label, 0.0)
        warnings.append(
            "EE 侧 PAYE / EE NIC / EE Pension 发票未提供，已置 0（需人工或其它来源）"
        )
    else:
        amounts.setdefault("Car Allowance", 0.0)
        amounts.setdefault("App Levy", 0.0)
        amounts.setdefault("Payment Fees", 0.0)

    for label, value in amounts.items():
        _set_amount_by_label(ws, label, value)

    if parsed.service_fee is not None:
        warnings.append(
            f"PDF Service Fee={parsed.service_fee}（与 PN Management Fee / Recurring Fee 不同项，未自动写入 UK!H）"
        )

    fx_rate = None
    if fill_fx:
        try:
            from fx_policy import UK_VENDOR_BILL_FX_FACT, api_fx_for_currency, read_shared_fx

            shared = read_shared_fx(convert_mapping, UK_VENDOR_BILL_FX_FACT)
            if shared is not None:
                fx_rate = float(shared)
            else:
                fx_rate = api_fx_for_currency("GBP", invert=True)
            ws["D24"] = fx_rate
        except Exception as exc:
            warnings.append(f"写入 UK-L!D24 汇率失败: {exc}")

    if UK_SHEET in wb.sheetnames:
        wb[UK_SHEET]["B9"] = name

    applied_pn = None
    if pn_meta is not None:
        applied_pn = apply_pn_meta(
            wb,
            pn_meta,
            registry_dir=registry_dir,
            reserve_invoice_number=True,
        )

    return warnings, fx_rate, applied_pn


def convert_pdf(
    pdf_path: Path,
    output_path: Path,
    *,
    template_path: Path | None = None,
    pn_meta: PnMeta | dict[str, Any] | None = None,
    registry_dir: Path | None = None,
    fill_fx: bool = True,
    convert_mapping: dict[str, Any] | None = None,
) -> dict[str, Any]:
    pdf_path = pdf_path.resolve()
    output_path = output_path.resolve()
    if not pdf_path.is_file():
        raise FileNotFoundError(f"PDF 不存在: {pdf_path}")

    tpl = (template_path or get_region_template("UK")).resolve()
    if not tpl.is_file():
        raise FileNotFoundError(f"UK 母版不存在: {tpl}")

    parsed = parse_eor_uk_pdf(pdf_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(tpl, output_path)

    wb = load_workbook(output_path)
    warnings, fx_rate, applied_pn = apply_to_workbook(
        wb,
        parsed,
        pn_meta=pn_meta,
        registry_dir=registry_dir or output_path.parent,
        fill_fx=fill_fx,
        convert_mapping=convert_mapping,
    )
    wb.save(output_path)
    wb.close()

    return {
        "ok": True,
        "profile_id": "eor_uk",
        "region": "UK",
        "source_kind": "pdf",
        "output": str(output_path),
        "employee_count": 1,
        "parsed": parsed.to_dict(),
        "warnings": warnings,
        "fx_rate": fx_rate,
        "pn_meta": applied_pn.to_dict() if applied_pn else None,
    }


def convert_excels(
    excel_paths: list[Path],
    output_path: Path,
    *,
    template_path: Path | None = None,
    pn_meta: PnMeta | dict[str, Any] | None = None,
    registry_dir: Path | None = None,
    fill_fx: bool = True,
    convert_mapping: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """EMPANPAY / 已成型 UK-L → 一份含 UK-L 的源表。"""
    paths = [Path(p).resolve() for p in excel_paths]
    if not paths:
        raise ValueError("未提供 Excel")
    for p in paths:
        if not p.is_file():
            raise FileNotFoundError(f"Excel 不存在: {p}")

    output_path = Path(output_path).resolve()
    uk_l_flags = [_looks_like_uk_l_workbook(p) for p in paths]
    if all(uk_l_flags):
        if len(paths) > 1:
            raise ValueError("多份已是 UK-L 的 Excel 无法自动合并，请只传一份")
        output_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(paths[0], output_path)
        return {
            "ok": True,
            "profile_id": "eor_uk",
            "region": "UK",
            "source_kind": "excel_uk_l",
            "output": str(output_path),
            "employee_count": None,
            "parsed": [],
            "warnings": ["源表已是 UK-L，已原样用作转换输入"],
            "fx_rate": None,
            "pn_meta": None,
        }

    if len(paths) > 1:
        raise ValueError(
            f"eor_uk Excel 暂仅支持单员工竖表（收到 {len(paths)} 份），请一次一份"
        )

    tpl = (template_path or get_region_template("UK")).resolve()
    if not tpl.is_file():
        raise FileNotFoundError(f"UK 母版不存在: {tpl}")

    parsed = parse_eor_uk_payroll_excel(paths[0], convert_mapping=convert_mapping)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(tpl, output_path)

    wb = load_workbook(output_path)
    warnings, fx_rate, applied_pn = apply_to_workbook(
        wb,
        parsed,
        pn_meta=pn_meta,
        registry_dir=registry_dir or output_path.parent,
        fill_fx=fill_fx,
        convert_mapping=convert_mapping,
    )
    wb.save(output_path)
    wb.close()

    return {
        "ok": True,
        "profile_id": "eor_uk",
        "region": "UK",
        "source_kind": "excel_payroll_totals",
        "output": str(output_path),
        "employee_count": 1,
        "parsed": parsed.to_dict(),
        "warnings": warnings,
        "fx_rate": fx_rate,
        "pn_meta": applied_pn.to_dict() if applied_pn else None,
    }


def _scalar_amounts_from_parsed(parsed: EorUkParsed) -> dict[str, float]:
    """从 PDF/Excel 标量字段拼出 UK-L 金额表（Excel.amounts 优先另合并）。"""
    amounts: dict[str, float] = {}
    if parsed.gross_salary is not None:
        amounts["Gross Salary"] = float(parsed.gross_salary)
    if parsed.holiday_pay is not None:
        amounts["Holiday Pay"] = float(parsed.holiday_pay)
    if parsed.er_nic is not None:
        amounts["ER' NIC"] = float(parsed.er_nic)
    if parsed.er_pension is not None:
        amounts["ER' Pension (Auto Enrolment)"] = float(parsed.er_pension)
    if parsed.paye is not None:
        amounts["PAYE (Estimated)"] = float(parsed.paye)
    if parsed.ee_nic is not None:
        amounts["EE'NIC"] = float(parsed.ee_nic)
    if parsed.ee_pension is not None:
        amounts["EE' Pension (Auto Enrolment)"] = float(parsed.ee_pension)
    return amounts


def _sync_scalars_from_amounts(parsed: EorUkParsed) -> None:
    am = parsed.amounts or {}
    parsed.gross_salary = am.get("Gross Salary", parsed.gross_salary)
    parsed.holiday_pay = am.get("Holiday Pay", parsed.holiday_pay)
    parsed.er_nic = am.get("ER' NIC", parsed.er_nic)
    parsed.er_pension = am.get("ER' Pension (Auto Enrolment)", parsed.er_pension)
    parsed.paye = am.get("PAYE (Estimated)", parsed.paye)
    parsed.ee_nic = am.get("EE'NIC", parsed.ee_nic)
    parsed.ee_pension = am.get("EE' Pension (Auto Enrolment)", parsed.ee_pension)


def merge_eor_uk_pdf_excel(pdf_parsed: EorUkParsed, excel_parsed: EorUkParsed) -> EorUkParsed:
    """
    混传合并：
    - 姓名 / 发票号 / 日期 / Service Fee / 勾稽字段 ← PDF
    - 工资竖表金额 ← Excel 优先；PDF 仅补 Excel 没有的项
    - 两边共有且差额 > 0.05 → warning（采用 Excel）
    """
    out = EorUkParsed(
        supplier=pdf_parsed.supplier or excel_parsed.supplier,
        invoice_no=pdf_parsed.invoice_no,
        invoice_date=pdf_parsed.invoice_date or excel_parsed.invoice_date,
        employee_name=pdf_parsed.employee_name or excel_parsed.employee_name,
        service_fee=pdf_parsed.service_fee,
        invoice_total=pdf_parsed.invoice_total,
        currency=pdf_parsed.currency or excel_parsed.currency or "GBP",
        raw_line_net=pdf_parsed.raw_line_net,
        source_kind="pdf_excel_merged",
    )
    out.warnings = list(pdf_parsed.warnings or []) + list(excel_parsed.warnings or [])
    # 合并场景不再需要「Excel 无姓名」类提示若 PDF 已有姓名
    if out.employee_name:
        out.warnings = [
            w
            for w in out.warnings
            if "Excel 竖表通常无员工姓名" not in str(w)
        ]

    pdf_am = _scalar_amounts_from_parsed(pdf_parsed)
    excel_am = dict(excel_parsed.amounts or {}) or _scalar_amounts_from_parsed(excel_parsed)
    merged: dict[str, float] = dict(pdf_am)
    for label, excel_val in excel_am.items():
        if label in merged and abs(float(merged[label]) - float(excel_val)) > 0.05:
            out.warnings.append(
                f"PDF/Excel「{label}」不一致：PDF={merged[label]} Excel={excel_val}，已采用 Excel"
            )
        merged[label] = float(excel_val)

    out.amounts = merged
    _sync_scalars_from_amounts(out)
    out.warnings.append(
        "已合并 PDF（姓名/发票/Service Fee）与 Excel（工资明细，含 EE 侧）；金额冲突以 Excel 为准"
    )

    if not out.employee_name:
        out.warnings.append(
            "合并后仍无员工姓名，UK-L 标题暂用 Employee；请在 PN 元数据补全"
        )
    if out.gross_salary is None:
        raise ValueError("EOR UK 合并后未得到 Gross Salary，请检查 PDF/Excel 或列名对照")
    return out


def _write_parsed_to_template(
    parsed: EorUkParsed,
    output_path: Path,
    *,
    template_path: Path | None = None,
    pn_meta: PnMeta | dict[str, Any] | None = None,
    registry_dir: Path | None = None,
    fill_fx: bool = True,
    convert_mapping: dict[str, Any] | None = None,
) -> dict[str, Any]:
    tpl = (template_path or get_region_template("UK")).resolve()
    if not tpl.is_file():
        raise FileNotFoundError(f"UK 母版不存在: {tpl}")
    output_path = Path(output_path).resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(tpl, output_path)
    wb = load_workbook(output_path)
    warnings, fx_rate, applied_pn = apply_to_workbook(
        wb,
        parsed,
        pn_meta=pn_meta,
        registry_dir=registry_dir or output_path.parent,
        fill_fx=fill_fx,
        convert_mapping=convert_mapping,
    )
    wb.save(output_path)
    wb.close()
    return {
        "ok": True,
        "profile_id": "eor_uk",
        "region": "UK",
        "source_kind": parsed.source_kind,
        "output": str(output_path),
        "employee_count": 1,
        "parsed": parsed.to_dict(),
        "warnings": warnings,
        "fx_rate": fx_rate,
        "pn_meta": applied_pn.to_dict() if applied_pn else None,
    }


def convert_sources(
    source_paths: list[Path],
    output_path: Path,
    *,
    template_path: Path | None = None,
    pn_meta: PnMeta | dict[str, Any] | None = None,
    registry_dir: Path | None = None,
    fill_fx: bool = True,
    convert_mapping: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """PDF / Excel：可单独转；混传则合并（PDF 元数据 + Excel 明细）。"""
    paths = [Path(p).resolve() for p in source_paths]
    pdfs = [p for p in paths if p.suffix.lower() == ".pdf"]
    excels = [p for p in paths if p.suffix.lower() in {".xlsx", ".xlsm", ".xls"}]

    if pdfs and excels:
        if len(pdfs) > 1:
            raise ValueError(
                f"eor_uk 混传时 PDF 仅支持 1 份（收到 {len(pdfs)} 份）"
            )
        if len(excels) > 1:
            raise ValueError(
                f"eor_uk 混传时 Excel 仅支持 1 份（收到 {len(excels)} 份）"
            )
        pdf_parsed = parse_eor_uk_pdf(pdfs[0])
        excel_parsed = parse_eor_uk_payroll_excel(
            excels[0], convert_mapping=convert_mapping
        )
        merged = merge_eor_uk_pdf_excel(pdf_parsed, excel_parsed)
        return _write_parsed_to_template(
            merged,
            output_path,
            template_path=template_path,
            pn_meta=pn_meta,
            registry_dir=registry_dir,
            fill_fx=fill_fx,
            convert_mapping=convert_mapping,
        )

    if pdfs:
        if len(pdfs) > 1:
            raise ValueError(
                f"eor_uk PDF 暂不支持批量（共 {len(pdfs)} 份），请一次只传 1 份"
            )
        return convert_pdf(
            pdfs[0],
            output_path,
            template_path=template_path,
            pn_meta=pn_meta,
            registry_dir=registry_dir,
            fill_fx=fill_fx,
            convert_mapping=convert_mapping,
        )

    if excels:
        return convert_excels(
            excels,
            output_path,
            template_path=template_path,
            pn_meta=pn_meta,
            registry_dir=registry_dir,
            fill_fx=fill_fx,
            convert_mapping=convert_mapping,
        )

    raise ValueError("未提供可用的 PDF 或 Excel 源")

def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="EOR UK 发票 PDF → UK-L PN 源表")
    parser.add_argument("pdf", type=Path, help="供应商 PDF 路径")
    parser.add_argument("-o", "--output", type=Path, help="输出 xlsx")
    parser.add_argument("-t", "--template", type=Path, help="UK PN 母版")
    args = parser.parse_args(argv)

    pdf = args.pdf.resolve()
    out = (
        args.output.resolve()
        if args.output
        else pdf.parent / f"UK_L_from_pdf_{pdf.stem}.xlsx"
    )
    try:
        result = convert_pdf(pdf, out, template_path=args.template)
    except Exception as exc:
        print(f"失败: {exc}", file=sys.stderr)
        return 1

    p = result["parsed"]
    print("完成")
    print(f"  输出: {result['output']}")
    print(f"  发票号: {p.get('invoice_no')}  日期: {p.get('invoice_date')}")
    print(f"  员工: {p.get('employee_name')}")
    print(
        f"  Gross={p.get('gross_salary')}  ER NIC={p.get('er_nic')}  "
        f"ER Pension={p.get('er_pension')}  ServiceFee={p.get('service_fee')}  "
        f"Total={p.get('invoice_total')}"
    )
    for w in result.get("warnings") or []:
        print(f"  ! {w}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
