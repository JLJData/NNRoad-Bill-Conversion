# -*- coding: utf-8 -*-
"""
UAE-L 横向源账单 → UAE PN（引擎 uae_payroll_calc）

用法:
  python -m profiles.uae_payroll_calc.convert <源.xlsx> [-o 输出.xlsx] [-t 母版.xlsx]

源账单: sheet「UAE-L」第 2 行表头、第 3 行起员工（可由 auxilium_uae ingest 产出）。
默认母版: templates/uae/template.xlsx
"""
from __future__ import annotations

import argparse
import re
import shutil
import sys
from copy import copy
from datetime import datetime
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.worksheet.worksheet import Worksheet

from bill_convert.formula_copy import shift_row_formula
from convert_mapping import find_sheet_name, resolve_convert_mapping
from fx_rate import fetch_usd_rates, get_uae_pn_fx_rate
from pn_meta import PnMeta, apply_pn_meta
from profiles.tw_payroll_calc.convert import match_ee_code
from region_templates import get_region_template
from xlsx_luckysheet_compat import apply_luckysheet_compat_uae
from xlsx_postprocess import postprocess_converted_xlsx

DEFAULT_TEMPLATE = get_region_template("UAE")

UAE_L_SHEET = "UAE-L"
UAE_SHEET = "UAE"
UAE_EE_SHEET = "UAE EE"
PN_SHEET = "PN"

UAE_L_HEADER_ROW = 2
UAE_L_DATA_START = 3
UAE_DATA_START = 9
UAE_EE_DATA_START = 10
MAX_EMPLOYEES = 20

_ACTIVE_MAPPING: dict[str, Any] | None = None


def _active_mapping() -> dict[str, Any]:
    return (
        _ACTIVE_MAPPING
        if isinstance(_ACTIVE_MAPPING, dict)
        else resolve_convert_mapping("uae_payroll_calc", None)
    )


def _norm(value: Any) -> str:
    if value is None:
        return ""
    return str(value).replace("\n", " ").strip()


def _as_float(value: Any) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "").replace("AED", "").replace("\xa0", "")
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def _header_map(ws: Worksheet, header_row: int) -> dict[str, int]:
    out: dict[str, int] = {}
    for col in range(1, (ws.max_column or 1) + 1):
        h = _norm(ws.cell(header_row, col).value)
        if h and h not in out:
            out[h] = col
    return out


def _cell_formula_text(value: Any) -> str | None:
    if isinstance(value, ArrayFormula):
        return value.text
    if isinstance(value, str) and value.startswith("="):
        return value
    return None


def _set_cell_value(cell, value: Any) -> None:
    """写入值；若目标是 ArrayFormula 则改为普通公式/值。"""
    cell.value = value


def parse_uae_l_employees(ws: Worksheet) -> list[dict[str, Any]]:
    headers = _header_map(ws, UAE_L_HEADER_ROW)
    if not headers:
        raise ValueError(f"「UAE-L」第 {UAE_L_HEADER_ROW} 行表头为空")
    employees: list[dict[str, Any]] = []
    for row in range(UAE_L_DATA_START, (ws.max_row or UAE_L_DATA_START) + 1):
        name = None
        if "Employee Name" in headers:
            name = ws.cell(row, headers["Employee Name"]).value
        emp_id = None
        if "Emp ID" in headers:
            emp_id = ws.cell(row, headers["Emp ID"]).value
        name_s = _norm(name)
        id_s = _norm(emp_id)
        if not name_s and not id_s:
            continue
        if name_s.upper().startswith("TOTAL") or id_s.upper().startswith("TOTAL"):
            continue
        row_data: dict[str, Any] = {}
        for h, col in headers.items():
            row_data[h] = ws.cell(row, col).value
        employees.append(row_data)
    return employees


def clear_uae_l_data(ws: Worksheet) -> None:
    max_row = max(ws.max_row or UAE_L_DATA_START, UAE_L_DATA_START + MAX_EMPLOYEES)
    max_col = min(ws.max_column or 1, 64)
    for row in range(UAE_L_DATA_START, max_row + 1):
        for col in range(1, max_col + 1):
            cell = ws.cell(row, col)
            # 保留 Gross 等表内公式列的结构：整行清空后由写入覆盖
            if type(cell).__name__ == "MergedCell":
                continue
            cell.value = None


def write_uae_l(ws: Worksheet, employees: list[dict[str, Any]]) -> None:
    headers = _header_map(ws, UAE_L_HEADER_ROW)
    if not headers:
        raise ValueError(f"「UAE-L」第 {UAE_L_HEADER_ROW} 行表头为空")
    clear_uae_l_data(ws)
    skip_formula_headers = {"EC - Gross Salary"}  # 保留母版行公式时跳过；清空后需重写值或公式
    for idx, emp in enumerate(employees):
        row = UAE_L_DATA_START + idx
        for h, col in headers.items():
            if h not in emp:
                continue
            if h in skip_formula_headers and emp.get(h) is None:
                continue
            val = emp[h]
            cell = ws.cell(row, col)
            if isinstance(val, datetime):
                cell.value = val
                cell.number_format = "yyyy/m/d"
            else:
                _set_cell_value(cell, val)
        # Gross 公式（无 Table 时用单元格加法兜底）
        gross_col = headers.get("EC - Gross Salary")
        if gross_col:
            parts = []
            for key in (
                "EC - Basic Salary",
                "EC - Housing Allowance",
                "EC - Transport Allowance",
                "EC - School Allowance",
                "EC - Other allowance",
                "EC - Mobile Allowance",
                "EC - Food Allowance",
            ):
                c = headers.get(key)
                if c:
                    parts.append(f"{ws.cell(row, c).coordinate}")
            if parts:
                ws.cell(row, gross_col).value = "=" + "+".join(parts)


def _copy_row_style_and_formula(
    ws: Worksheet,
    src_row: int,
    dest_row: int,
    max_col: int = 40,
    *,
    l_from: int | None = None,
    l_to: int | None = None,
) -> None:
    for c in range(1, max_col + 1):
        src = ws.cell(src_row, c)
        dest = ws.cell(dest_row, c)
        if type(dest).__name__ == "MergedCell" or type(src).__name__ == "MergedCell":
            continue
        if src.has_style:
            dest.font = copy(src.font)
            dest.border = copy(src.border)
            dest.fill = copy(src.fill)
            dest.number_format = src.number_format
            dest.protection = copy(src.protection)
            dest.alignment = copy(src.alignment)
        text = _cell_formula_text(src.value)
        if text:
            dest.value = shift_row_formula(
                text,
                src_row,
                dest_row,
                target_l_from=l_from if l_from is not None else -1,
                target_l_to=l_to if l_to is not None else -1,
                target_l_sheet=UAE_L_SHEET,
            )
        elif src.value is not None and not isinstance(src.value, ArrayFormula):
            dest.value = src.value


def _retarget_ee_refs(formula: str, ee_from: int, ee_to: int) -> str:
    if not isinstance(formula, str) or not formula.startswith("="):
        return formula
    return re.sub(
        rf"('UAE EE'!\$?[A-Z]{{1,3}})\$?{ee_from}(?!\d)",
        lambda m: f"{m.group(1)}{ee_to}",
        formula,
        flags=re.I,
    )


def expand_uae_employee_rows(wb, employee_count: int) -> None:
    n = max(int(employee_count), 1)
    if UAE_SHEET in wb.sheetnames:
        uae = wb[UAE_SHEET]
        for i in range(1, n):
            dest = UAE_DATA_START + i
            l_row = UAE_L_DATA_START + i
            ee_row = UAE_EE_DATA_START + i
            _copy_row_style_and_formula(
                uae,
                UAE_DATA_START,
                dest,
                max_col=40,
                l_from=UAE_L_DATA_START,
                l_to=l_row,
            )
            for c in range(1, 41):
                cell = uae.cell(dest, c)
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    cell.value = _retarget_ee_refs(cell.value, UAE_EE_DATA_START, ee_row)
            uae.cell(dest, 2).value = f"='{UAE_L_SHEET}'!B{l_row}"

    if UAE_EE_SHEET in wb.sheetnames:
        ee = wb[UAE_EE_SHEET]
        for i in range(1, n):
            dest = UAE_EE_DATA_START + i
            l_row = UAE_L_DATA_START + i
            _copy_row_style_and_formula(
                ee,
                UAE_EE_DATA_START,
                dest,
                max_col=40,
                l_from=UAE_L_DATA_START,
                l_to=l_row,
            )
            ee.cell(dest, 5).value = f"='{UAE_L_SHEET}'!B{l_row}"


# ---------- PN 多人：Labor×n + Expense×n（无 Office Rental）----------

_PN_LABOR_START = 16
_PN_AED_FMT = "#,##0.00"
_PN_USD_FMT = "$#,##0.00"


def _find_pn_row_by_label(ws: Worksheet, keyword: str, col: int = 1) -> int | None:
    key = keyword.lower()
    for row in range(1, (ws.max_row or 0) + 1):
        v = ws.cell(row, col).value
        if isinstance(v, str) and key in v.lower():
            return row
    return None


def _copy_pn_row_style(ws: Worksheet, src_row: int, dst_row: int) -> None:
    max_col = max(ws.max_column or 6, 6)
    for col in range(1, max_col + 1):
        src = ws.cell(src_row, col)
        dst = ws.cell(dst_row, col)
        if src.has_style:
            dst.font = copy(src.font)
            dst.border = copy(src.border)
            dst.fill = copy(src.fill)
            dst.number_format = src.number_format
            dst.protection = copy(src.protection)
            dst.alignment = copy(src.alignment)
    if ws.row_dimensions[src_row].height is not None:
        ws.row_dimensions[dst_row].height = ws.row_dimensions[src_row].height


def _ensure_merge_a_c(ws: Worksheet, row: int) -> None:
    rng = f"A{row}:C{row}"
    for m in list(ws.merged_cells.ranges):
        if m.min_row <= row <= m.max_row and m.min_col <= 1 and m.max_col >= 3:
            if m.min_row == row and m.max_row == row and m.min_col == 1 and m.max_col == 3:
                return
            try:
                ws.unmerge_cells(str(m))
            except ValueError:
                pass
    try:
        ws.merge_cells(rng)
    except ValueError:
        pass


def _collect_merges_from(ws: Worksheet, start_row: int) -> list[tuple[int, int, int, int]]:
    kept: list[tuple[int, int, int, int]] = []
    for m in list(ws.merged_cells.ranges):
        if m.min_row < start_row:
            continue
        kept.append((m.min_row, m.min_col, m.max_row, m.max_col))
        try:
            ws.unmerge_cells(str(m))
        except ValueError:
            pass
    return kept


def _restore_merges(
    ws: Worksheet,
    merges: list[tuple[int, int, int, int]],
    row_shift: int = 0,
) -> None:
    for min_r, min_c, max_r, max_c in merges:
        nr1, nr2 = min_r + row_shift, max_r + row_shift
        if nr1 < 1 or nr2 < nr1:
            continue
        try:
            ws.merge_cells(
                start_row=nr1,
                start_column=min_c,
                end_row=nr2,
                end_column=max_c,
            )
        except ValueError:
            pass


def _shift_row_heights(ws: Worksheet, start_row: int, amount: int) -> None:
    if amount == 0:
        return
    captured: dict[int, float] = {}
    max_r = max(ws.max_row or start_row, start_row)
    for r in list(ws.row_dimensions.keys()):
        if not isinstance(r, int) or r < start_row:
            continue
        h = ws.row_dimensions[r].height
        if h is not None:
            captured[r] = h
            max_r = max(max_r, r)
    for r in range(start_row, max_r + abs(amount) + 2):
        if r in ws.row_dimensions:
            ws.row_dimensions[r].height = None
    if amount > 0:
        for r in sorted(captured.keys(), reverse=True):
            ws.row_dimensions[r + amount].height = captured[r]
    else:
        for r in sorted(captured.keys()):
            new_r = r + amount
            if new_r >= start_row:
                ws.row_dimensions[new_r].height = captured[r]


def _pn_insert_rows(ws: Worksheet, idx: int, amount: int, fill_style_row: int | None = None) -> None:
    if amount <= 0:
        return
    merges = _collect_merges_from(ws, idx)
    _shift_row_heights(ws, idx, amount)
    ws.insert_rows(idx, amount)
    _restore_merges(ws, merges, row_shift=amount)
    if fill_style_row is not None:
        style_row = fill_style_row + amount if fill_style_row >= idx else fill_style_row
        for r in range(idx, idx + amount):
            _copy_pn_row_style(ws, style_row, r)
            _ensure_merge_a_c(ws, r)


def _count_pn_labor_slots(ws: Worksheet) -> int:
    n = 0
    for row in range(_PN_LABOR_START, _PN_LABOR_START + MAX_EMPLOYEES + 2):
        v = ws.cell(row, 1).value
        if isinstance(v, str) and "Labor cost" in v:
            n += 1
            continue
        break
    return max(n, 1)


def fit_uae_pn_employees(wb, employee_count: int) -> dict[str, int]:
    """
    PN 按人数扩行：Labor×n + Expense×n，重写 EOR/Service/Management/FX 引用。
    母版默认 1 人：Labor16 / Expense17 / Service20 / Mgmt21 / FX B28。
    """
    if PN_SHEET not in wb.sheetnames:
        return {}
    ws = wb[PN_SHEET]
    n = max(int(employee_count), 1)
    if n > MAX_EMPLOYEES:
        raise ValueError(f"员工数 {n} 超过模板上限 {MAX_EMPLOYEES}")

    old_n = _count_pn_labor_slots(ws)
    labor_start = _PN_LABOR_START
    expense_start = labor_start + old_n
    svc_row = _find_pn_row_by_label(ws, "Service Fee") or 20
    delta = n - old_n

    if delta > 0:
        _pn_insert_rows(ws, expense_start, delta, fill_style_row=labor_start)
        svc_row = _find_pn_row_by_label(ws, "Service Fee") or (svc_row + delta)
        expense_start = labor_start + n
        _pn_insert_rows(ws, svc_row, delta, fill_style_row=expense_start)
    elif delta < 0:
        expense_start = labor_start + old_n

        def _pn_delete_rows(wss: Worksheet, idx: int, amount: int) -> None:
            if amount <= 0:
                return
            merges = _collect_merges_from(wss, idx + amount)
            wss.delete_rows(idx, amount)
            _restore_merges(wss, merges, row_shift=-amount)

        _pn_delete_rows(ws, expense_start + n, -delta)
        _pn_delete_rows(ws, labor_start + n, -delta)

    expense_start = labor_start + n
    last_eor_detail = expense_start + n - 1
    eor_row = _find_pn_row_by_label(ws, "EOR/PEO Cost") or 15
    fx_row = _find_pn_row_by_label(ws, "FX rate") or 28
    svc = _find_pn_row_by_label(ws, "Service Fee")
    mgmt = _find_pn_row_by_label(ws, "Management Fee")
    if svc is None:
        svc = last_eor_detail + 3
    if mgmt is None:
        mgmt = svc + 1

    for i in range(n):
        uae_row = UAE_DATA_START + i
        labor_row = labor_start + i
        expense_row = expense_start + i
        if i > 0:
            _copy_pn_row_style(ws, labor_start, labor_row)
            _copy_pn_row_style(ws, expense_start, expense_row)
        _ensure_merge_a_c(ws, labor_row)
        _ensure_merge_a_c(ws, expense_row)

        ws.cell(labor_row, 1).value = (
            f'="- Labor cost for "&UAE!B{uae_row}&"  -  "&MONTH(UAE!B2)&"-"&YEAR(UAE!B2)'
        )
        e_labor = ws.cell(labor_row, 5)
        e_labor.value = f"=UAE!I{uae_row}+UAE!S{uae_row}"
        e_labor.number_format = _PN_AED_FMT
        f_labor = ws.cell(labor_row, 6)
        f_labor.value = f"=E{labor_row}/$B${fx_row}"
        f_labor.number_format = _PN_USD_FMT

        ws.cell(expense_row, 1).value = f'="- Expense claim for "&UAE!B{uae_row}'
        e_exp = ws.cell(expense_row, 5)
        e_exp.value = f"=UAE!AE{uae_row}"
        e_exp.number_format = _PN_AED_FMT
        f_exp = ws.cell(expense_row, 6)
        f_exp.value = f"=E{expense_row}/$B${fx_row}"
        f_exp.number_format = _PN_USD_FMT

    e_eor = ws.cell(eor_row, 5)
    e_eor.value = f"=SUM(E{labor_start}:E{last_eor_detail})"
    e_eor.number_format = _PN_AED_FMT
    f_eor = ws.cell(eor_row, 6)
    f_eor.value = f"=SUM(F{labor_start}:F{last_eor_detail})"
    f_eor.number_format = _PN_USD_FMT

    e_svc = ws.cell(svc, 5)
    e_svc.value = f"=SUM(E{mgmt}:E{mgmt + 1})"
    e_svc.number_format = _PN_AED_FMT
    f_svc = ws.cell(svc, 6)
    f_svc.value = f"=SUM(F{mgmt}:F{mgmt + 1})"
    f_svc.number_format = _PN_USD_FMT

    # Management：UAE!H6 为 Recurring Fee 合计（多人）
    ws.cell(mgmt, 1).value = '="- Management Fee - "&MONTH(UAE!B2)&"-"&YEAR(UAE!B2)'
    e_mgmt = ws.cell(mgmt, 5)
    e_mgmt.value = "=UAE!H6"
    e_mgmt.number_format = _PN_AED_FMT
    f_mgmt = ws.cell(mgmt, 6)
    f_mgmt.value = f"=E{mgmt}/$B${fx_row}"
    f_mgmt.number_format = _PN_USD_FMT
    _ensure_merge_a_c(ws, svc)
    _ensure_merge_a_c(ws, mgmt)

    # 扩行后重写合计区（EOR/PEO Service Cost、FX 旁合计、Sub/Tax/Total）
    total_row = _find_pn_row_by_label(ws, "EOR/PEO Service Cost") or (mgmt + 4)
    ws.cell(total_row, 5).value = f"=E{eor_row}+E{svc}"
    ws.cell(total_row, 5).number_format = _PN_AED_FMT
    ws.cell(total_row, 6).value = f"=F{eor_row}+F{svc}"
    ws.cell(total_row, 6).number_format = _PN_USD_FMT

    fx_row = _find_pn_row_by_label(ws, "FX rate") or fx_row
    # FX 行右侧「EOR/PEO Service Cost」USD
    ws.cell(fx_row, 6).value = f"=F{total_row}"
    ws.cell(fx_row, 6).number_format = _PN_USD_FMT
    # Sub Total / Tax / Total（相对 FX 行）
    sub_row = fx_row + 5
    tax_row = fx_row + 6
    grand_row = fx_row + 7
    # Outstanding..Bank = fx+1..fx+4，保持母版对 UAE!E22.. 的引用即可
    ws.cell(sub_row, 6).value = f"=SUM(F{fx_row}:F{fx_row + 4})"
    ws.cell(sub_row, 6).number_format = _PN_USD_FMT
    tax_cell = ws.cell(tax_row, 6)
    tax_cell.value = f"=UAE!F6/'PN'!B{fx_row}"
    tax_cell.number_format = _PN_USD_FMT
    ws.cell(grand_row, 6).value = f"=SUM(F{sub_row}:F{tax_row})"
    ws.cell(grand_row, 6).number_format = _PN_USD_FMT

    # Outstanding..Bank → UAE!E22..E25（扩行后行号仍指向地区表固定结算区）
    for i, uae_e in enumerate((22, 23, 24, 25), start=1):
        cell = ws.cell(fx_row + i, 6)
        cell.value = f"=UAE!E{uae_e}"
        cell.number_format = _PN_USD_FMT

    # 明细行 USD 除数若仍指向旧 B28，一并改到当前 FX 行
    for r in range(labor_start, last_eor_detail + 1):
        fcell = ws.cell(r, 6)
        if isinstance(fcell.value, str) and "$B$" in fcell.value:
            fcell.value = f"=E{r}/$B${fx_row}"
    for r in (mgmt,):
        fcell = ws.cell(r, 6)
        if isinstance(fcell.value, str):
            fcell.value = f"=E{r}/$B${fx_row}"

    return {
        "labor_start": labor_start,
        "expense_start": expense_start,
        "eor_row": eor_row,
        "svc_row": svc,
        "mgmt_row": mgmt,
        "fx_row": fx_row,
        "total_row": total_row,
        "employee_count": n,
    }


def set_recurring_fees(wb, employees: list[dict[str, Any]]) -> None:
    """UAE!H = Admin Fee × 1.5（样例 1312.5→1968）。"""
    if UAE_SHEET not in wb.sheetnames:
        return
    uae = wb[UAE_SHEET]
    for i, emp in enumerate(employees):
        admin = _as_float(emp.get("EC - Admin Fees"))
        if admin is None:
            admin = _as_float(emp.get("Admin Fees"))
        if admin is None:
            admin = _as_float(emp.get("_admin_fee"))
        if admin is None:
            continue
        fee = round(admin * 1.5, 6)
        uae.cell(UAE_DATA_START + i, 8).value = fee


def _resolve_pdf_profile_id(mapping: dict[str, Any] | None) -> str | None:
    if not isinstance(mapping, dict):
        return None
    for key in ("pdfProfileId", "_pdfProfileId"):
        val = mapping.get(key)
        if val is not None and str(val).strip():
            return str(val).strip()
    from bill_convert.fact_store import get_batch_facts

    batch = get_batch_facts(mapping)
    if any(str(k).startswith("auxilium.") for k in batch):
        return "auxilium_uae"
    return None


def _apply_vendor_plugins(wb, warnings: list[str], *, employee_count: int = 1) -> dict[str, Any]:
    """按 pdfProfile 加载供应商旁路插件（如 Auxilium Admin Fee → Business Tax）。"""
    from bill_convert.fact_store import get_batch_facts
    from bill_convert.vendor_plugins.runtime import apply_vendor_plugins

    mapping = _active_mapping()
    pdf_profile_id = _resolve_pdf_profile_id(mapping)
    if not pdf_profile_id:
        return {}
    return apply_vendor_plugins(
        wb,
        pdf_profile_id=pdf_profile_id,
        mapping=mapping,
        batch_facts=get_batch_facts(mapping),
        warnings=warnings,
        employee_count=employee_count,
    ) or {}


def set_period(wb, employees: list[dict[str, Any]]) -> None:
    if not employees or UAE_SHEET not in wb.sheetnames:
        return
    emp = employees[0]
    uae = wb[UAE_SHEET]
    for key, col in (("From", 2), ("To", 3)):
        val = emp.get(key)
        if isinstance(val, datetime):
            cell = uae.cell(2, col)
            cell.value = val
            cell.number_format = "yyyy/m/d"
        elif val is None:
            # 兜底：指向 UAE-L 首名员工账期
            uae.cell(2, col).value = f"='{UAE_L_SHEET}'!{'F' if col == 2 else 'G'}{UAE_L_DATA_START}"


def normalize_uae_other_fee_block(wb, *, fx_row: int = 28) -> None:
    """
    Other Fee 区（E22:E26）：
    - E23/E24 母版常空（值误在 F 列）→ 补 0
    - E25 Bank Charges 保留公式 =10+50/'PN'!B{fx}（多人扩行后 fx 行会下移，必须跟 PN 汇率格）
    """
    if UAE_SHEET not in wb.sheetnames:
        return
    uae = wb[UAE_SHEET]
    for row, default in ((22, 0), (23, 0), (24, 0), (26, 0)):
        cell = uae.cell(row, 5)
        val = cell.value
        if val is None or val == "":
            cell.value = default
        elif isinstance(val, str) and val.startswith("="):
            # 空结算项勿留坏公式
            cell.value = default
    # Bank Charges：固定规则 10+50/汇率；汇率格随 PN 扩行变化
    fx = max(int(fx_row or 28), 1)
    uae.cell(25, 5).value = f"=10+50/'PN'!B{fx}"
    uae.cell(21, 5).value = "=SUM(E22:E26)"
    for row in (23, 24):
        fcell = uae.cell(row, 6)
        if fcell.value == 0:
            fcell.value = None


def _pn_customer_id(pn_meta: PnMeta | dict[str, Any] | None) -> str | None:
    if pn_meta is None:
        return None
    if isinstance(pn_meta, PnMeta):
        cid = (pn_meta.customer_id or "").strip()
        return cid or None
    if isinstance(pn_meta, dict):
        cid = str(pn_meta.get("customer_id") or pn_meta.get("customerId") or "").strip()
        return cid or None
    return None


def apply_uae_ee_codes(
    wb,
    employees: list[dict[str, Any]],
    *,
    employee_directory: list[dict[str, Any]] | None = None,
    pn_meta: PnMeta | dict[str, Any] | None = None,
) -> list[str]:
    """
    UAE EE!B = Client Code（库客户编号，同 PN!B9）
    UAE EE!D = EE Code（按姓名匹配客户员工目录）
    """
    if UAE_EE_SHEET not in wb.sheetnames:
        return []
    ws = wb[UAE_EE_SHEET]
    client_code = _pn_customer_id(pn_meta)
    directory = list(employee_directory or [])
    warnings: list[str] = []

    for i, emp in enumerate(employees):
        row = UAE_EE_DATA_START + i
        # Client Code：有元数据则写入库值，否则公式指向 PN!B9
        if client_code:
            ws.cell(row, 2).value = client_code
        else:
            ws.cell(row, 2).value = "=PN!$B$9"
        # Client Name 始终跟 PN
        ws.cell(row, 3).value = "=PN!$B$8"

        excel_names = [_norm(emp.get("Employee Name"))]
        code, warn = match_ee_code([n for n in excel_names if n], directory)
        ws.cell(row, 4).value = code  # 匹配不到显式清空，避免母版残留
        if warn:
            warnings.append(f"UAE EE 第{i + 1}人：{warn}")
    return warnings


def apply_fx(wb, *, fill_fx: bool = True, fx_row: int | None = None) -> float | None:
    if not fill_fx or PN_SHEET not in wb.sheetnames:
        return None
    rates = fetch_usd_rates()
    fx = get_uae_pn_fx_rate(rates)
    row = fx_row or _find_pn_row_by_label(wb[PN_SHEET], "FX rate") or 28
    wb[PN_SHEET].cell(row, 2).value = fx
    return fx


def convert(
    source_path: Path,
    output_path: Path,
    template_path: Path,
    *,
    pn_meta: PnMeta | dict[str, Any] | None = None,
    employee_directory: list[dict[str, Any]] | None = None,
    registry_dir: Path | None = None,
    convert_mapping: dict[str, Any] | None = None,
    fill_fx: bool = True,
) -> dict[str, Any]:
    global _ACTIVE_MAPPING
    _ACTIVE_MAPPING = resolve_convert_mapping("uae_payroll_calc", convert_mapping)
    try:
        source_path = Path(source_path).resolve()
        output_path = Path(output_path).resolve()
        template_path = Path(template_path).resolve()
        if not source_path.is_file():
            raise FileNotFoundError(f"源文件不存在: {source_path}")
        if not template_path.is_file():
            raise FileNotFoundError(f"母版不存在: {template_path}")

        src_wb = load_workbook(source_path, data_only=False)
        try:
            l_name = find_sheet_name(list(src_wb.sheetnames), _active_mapping().get("sourceEmployeeSheet"))
            if not l_name or l_name not in src_wb.sheetnames:
                if UAE_L_SHEET in src_wb.sheetnames:
                    l_name = UAE_L_SHEET
                else:
                    raise ValueError(f"未找到 UAE-L，现有: {src_wb.sheetnames}")
            employees = parse_uae_l_employees(src_wb[l_name])
        finally:
            src_wb.close()

        if not employees:
            raise ValueError("UAE-L 未解析到员工行")

        output_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(template_path, output_path)
        wb = load_workbook(output_path)
        warnings: list[str] = []
        fact_store_updates: dict[str, Any] = {}
        fx = None
        applied_pn = None
        try:
            if UAE_L_SHEET not in wb.sheetnames:
                raise ValueError(f"母版缺少 {UAE_L_SHEET}")
            write_uae_l(wb[UAE_L_SHEET], employees)
            # 首行姓名写入 UAE B9（Table 公式可能失效时兜底）
            if UAE_SHEET in wb.sheetnames:
                name = _norm(employees[0].get("Employee Name"))
                if name:
                    wb[UAE_SHEET].cell(UAE_DATA_START, 2).value = name
            set_period(wb, employees)
            expand_uae_employee_rows(wb, len(employees))
            set_recurring_fees(wb, employees)
            fact_store_updates = _apply_vendor_plugins(wb, warnings, employee_count=len(employees))
            pn_layout = fit_uae_pn_employees(wb, len(employees))
            try:
                fx = apply_fx(wb, fill_fx=fill_fx, fx_row=pn_layout.get("fx_row"))
            except Exception as exc:
                warnings.append(f"写入 PN 汇率失败: {exc}")
            # 必须在 PN 扩行 + 写汇率之后，Bank Charges 才能指向正确的 B{fx}
            normalize_uae_other_fee_block(wb, fx_row=int(pn_layout.get("fx_row") or 28))

            if pn_meta is not None:
                applied_pn = apply_pn_meta(
                    wb,
                    pn_meta,
                    registry_dir=registry_dir or output_path.parent,
                    reserve_invoice_number=True,
                )
            warnings.extend(
                apply_uae_ee_codes(
                    wb,
                    employees,
                    employee_directory=employee_directory,
                    pn_meta=applied_pn or pn_meta,
                )
            )
            # Table1/ArrayFormula → 普通 A1，避免核对页 LuckySheet/HF 一直转圈
            apply_luckysheet_compat_uae(wb)
            wb.save(output_path)
        finally:
            wb.close()

        postprocess_converted_xlsx(output_path)
        return {
            "ok": True,
            "engine_id": "uae_payroll_calc",
            "region": "UAE",
            "output": str(output_path),
            "employee_count": len(employees),
            "fx_rate": fx,
            "warnings": warnings,
            "pn_meta": applied_pn.to_dict() if applied_pn else None,
            "fact_store_updates": fact_store_updates,
        }
    finally:
        _ACTIVE_MAPPING = None


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="UAE-L → UAE PN")
    parser.add_argument("source", type=Path)
    parser.add_argument("-o", "--output", type=Path)
    parser.add_argument("-t", "--template", type=Path, default=None)
    args = parser.parse_args(argv)
    src = args.source.resolve()
    out = args.output.resolve() if args.output else src.parent / f"PN_UAE_{src.stem}.xlsx"
    tpl = (args.template or DEFAULT_TEMPLATE).resolve()
    try:
        result = convert(src, out, tpl)
    except Exception as exc:
        print(f"失败: {exc}", file=sys.stderr)
        return 1
    print("完成", result.get("output"), "人数", result.get("employee_count"), "FX", result.get("fx_rate"))
    for w in result.get("warnings") or []:
        print(" !", w)
    return 0


if __name__ == "__main__":
    sys.exit(main())
