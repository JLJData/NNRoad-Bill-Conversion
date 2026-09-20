# -*- coding: utf-8 -*-
"""
Kyrgyzstan：Atlas Employment「Payroll cost calculation」Excel（或已是 Kyrgyzstan-L）→ Kyrgyzstan PN

用法:
  python -m profiles.kyrgyzstan_payroll_calc.convert <源.xlsx> [-o 输出.xlsx] [-t 母版.xlsx]

原则：
- PN / Kyrgyzstan / Kyrgyzstan EE 以母版公式为准；只写 Kyrgyzstan-L 数据与 PN 汇率。
- 税/净薪/总成本列保留母版公式（由 Base Salary 推导）。
- EE Code 必须来自员工库匹配，禁止沿用供应商账单工号。
- Atlas Invoice PDF（差旅垫付等）一期不自动写入；样例 PN 的 Expense 为 0。
"""
from __future__ import annotations

import argparse
import calendar
import re
import shutil
import sys
from copy import copy
from datetime import date, datetime
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.worksheet.worksheet import Worksheet

from bill_convert.convert_checks import (
    merge_warnings,
    parse_cell_ref,
    sanity_check_convert_result,
)
from bill_convert.formula_copy import shift_row_formula
from bill_convert.formula_layout import sort_employees_by_code
from bill_convert.headers import list_qualified_header_cells
from convert_mapping import find_sheet_name, resolve_convert_mapping
from fx_rate import get_usd_rate
from pn_meta import PnMeta, apply_pn_meta
from profiles.tw_payroll_calc.convert import match_ee_code
from region_templates import get_region_template
from xlsx_convert_utils import coerce_datetime_for_excel, norm
from xlsx_luckysheet_compat import apply_luckysheet_compat
from xlsx_postprocess import postprocess_converted_xlsx

DEFAULT_TEMPLATE = get_region_template("Kyrgyzstan")

KYRGYZSTAN_L_SHEET = "Kyrgyzstan-L"
KYRGYZSTAN_SHEET = "Kyrgyzstan"
KYRGYZSTAN_EE_SHEET = "Kyrgyzstan EE"
PN_SHEET = "PN"

KYRGYZSTAN_L_HEADER_ROW = 7
KYRGYZSTAN_L_DATA_START = 8
KYRGYZSTAN_DATA_START = 9
KYRGYZSTAN_EE_DATA_START = 10
MAX_EMPLOYEES = 20
_DATE_FMT = "yyyy/m/d"

COL_EE_CODE = 1
COL_NAME = 2

# 母版公式列 / EE Code：不从源表写入
_SKIP_WRITE_TARGETS = frozenset(
    {
        "No. of EE",
        "Gross Salary",
        "Personal Income Tax",
        "Pension fund",
        "Accumulative pension fund",
        "Net Salary",
        "Medical insurance fund",
        "Workers' Health Fund",
        "Workmen Compensation Insurance",
        "Other#2",
        "Total Cost",
    }
)

# 缺省写 0（样例 PN：Adj/Bonus/Other/Expense）；表头「Other 」经资格化后为 Other
_ZERO_FILL_KEYS = ("Salary Adjusment", "Bonus", "Other", "Expense Reimbursment")

_ACTIVE_MAPPING: dict[str, Any] | None = None

_MONTHS = {
    "jan": 1,
    "january": 1,
    "feb": 2,
    "february": 2,
    "mar": 3,
    "march": 3,
    "apr": 4,
    "april": 4,
    "may": 5,
    "jun": 6,
    "june": 6,
    "jul": 7,
    "july": 7,
    "aug": 8,
    "august": 8,
    "sep": 9,
    "sept": 9,
    "september": 9,
    "oct": 10,
    "october": 10,
    "nov": 11,
    "november": 11,
    "dec": 12,
    "december": 12,
}


def _active_mapping() -> dict[str, Any]:
    return (
        _ACTIVE_MAPPING
        if isinstance(_ACTIVE_MAPPING, dict)
        else resolve_convert_mapping("kyrgyzstan_payroll_calc", None)
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
    text = str(value).strip().replace(",", "").replace("\xa0", "")
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def _cell_formula_text(value: Any) -> str | None:
    if isinstance(value, ArrayFormula):
        return value.text
    if isinstance(value, str) and value.startswith("="):
        return value
    return None


def _set_cell_value(cell, value: Any) -> None:
    if isinstance(value, float):
        cell.value = round(value, 6) if abs(value) >= 1e-9 else 0
        fmt = str(cell.number_format or "")
        if "yy" in fmt.lower() or "h:" in fmt.lower() or "m/d" in fmt.lower():
            cell.number_format = "General"
    elif isinstance(value, int):
        cell.value = value
        fmt = str(cell.number_format or "")
        if "yy" in fmt.lower() or "h:" in fmt.lower() or "m/d" in fmt.lower():
            cell.number_format = "General"
    else:
        cell.value = value


def _kyrgyzstan_l_layout(*, target: bool = False) -> tuple[int, int]:
    mapping = _active_mapping()
    key = "targetL" if target else "sourceEmployeeSheet"
    spec = mapping.get(key) if isinstance(mapping.get(key), dict) else {}
    if not spec and target:
        spec = (
            mapping.get("sourceEmployeeSheet")
            if isinstance(mapping.get("sourceEmployeeSheet"), dict)
            else {}
        )
    header = int(spec.get("headerRow") or KYRGYZSTAN_L_HEADER_ROW)
    data_start = int(spec.get("dataStartRow") or KYRGYZSTAN_L_DATA_START)
    return header, data_start


def _meta_cell(key: str, default_row: int, default_col: int) -> tuple[int, int]:
    mapping = _active_mapping()
    meta = mapping.get("metaCells") if isinstance(mapping.get("metaCells"), dict) else {}
    return parse_cell_ref(meta.get(key), default_row=default_row, default_col=default_col)


def parse_pay_period_label(label: Any) -> tuple[int, int] | None:
    text = _norm(label)
    if not text:
        return None
    m = re.search(
        r"(jan|january|feb|february|mar|march|apr|april|may|jun|june|jul|july|"
        r"aug|august|sep|sept|september|oct|october|nov|november|dec|december)"
        r"[\s\-',]*(\d{2,4})",
        text,
        flags=re.I,
    )
    if not m:
        m2 = re.search(r"(\d{1,2})\s*/\s*(\d{4})", text)
        if m2:
            return int(m2.group(2)), int(m2.group(1))
        return None
    month = _MONTHS.get(m.group(1).lower())
    if not month:
        return None
    year = int(m.group(2))
    if year < 100:
        year += 2000
    return year, month


def period_bounds_from_label(label: Any) -> tuple[date, date] | None:
    parsed = parse_pay_period_label(label)
    if not parsed:
        return None
    year, month = parsed
    start = date(year, month, 1)
    end = date(year, month, calendar.monthrange(year, month)[1])
    return start, end


def _kyrgyzstan_l_formula_cols(ws: Worksheet, data_start: int) -> dict[int, str]:
    out: dict[int, str] = {}
    for col in range(1, (ws.max_column or 1) + 1):
        text = _cell_formula_text(ws.cell(data_start, col).value)
        if text:
            out[col] = text
    return out


def looks_like_kyrgyzstan_l(ws: Worksheet) -> bool:
    qualified = list_qualified_header_cells(ws, KYRGYZSTAN_L_HEADER_ROW)
    keys = {norm(str(h.get("key") or "")) for h in qualified}
    return "name of employee" in keys and "base salary" in keys


def looks_like_atlas_cost_sheet(ws: Worksheet) -> bool:
    blob = " ".join(_norm(ws.cell(r, 1).value).lower() for r in range(1, 6))
    if "atlas" in blob or "payrol" in blob or "payroll cost" in blob:
        return True
    for r in range(1, min((ws.max_row or 1), 40) + 1):
        if "gross salary" in _norm(ws.cell(r, 1).value).lower():
            return True
    return False


def _find_label_row(ws: Worksheet, *needles: str) -> int | None:
    wanted = [n.lower() for n in needles]
    for r in range(1, (ws.max_row or 0) + 1):
        label = _norm(ws.cell(r, 1).value).lower()
        if any(n in label for n in wanted):
            return r
    return None


def _eval_or_float(ws: Worksheet, row: int, col: int) -> float | None:
    """优先 data_only 缓存；公式格则尝试从关联格推算 GROSS。"""
    cell = ws.cell(row, col)
    direct = _as_float(cell.value)
    if direct is not None:
        return direct
    text = _cell_formula_text(cell.value)
    if not text:
        return None
    # E10 = 240*22 这类简单乘积
    m = re.fullmatch(r"=\s*([\d.]+)\s*\*\s*([\d.]+)\s*", text.replace(" ", ""))
    if m:
        return float(m.group(1)) * float(m.group(2))
    # D10 = I10*C5 / =E10 等：递归一层
    m2 = re.fullmatch(r"=\s*([A-Z]+)(\d+)\s*\*\s*([A-Z]+)(\d+)\s*", text, flags=re.I)
    if m2:
        a = _as_float(ws[f"{m2.group(1)}{m2.group(2)}"].value)
        b = _as_float(ws[f"{m2.group(3)}{m2.group(4)}"].value)
        if a is None:
            a = _eval_or_float(ws, int(m2.group(2)), _col_letter_to_index(m2.group(1)))
        if b is None:
            b = _eval_or_float(ws, int(m2.group(4)), _col_letter_to_index(m2.group(3)))
        if a is not None and b is not None:
            return a * b
    m3 = re.fullmatch(r"=\s*([A-Z]+)(\d+)\s*", text, flags=re.I)
    if m3:
        return _eval_or_float(ws, int(m3.group(2)), _col_letter_to_index(m3.group(1)))
    return None


def _col_letter_to_index(letter: str) -> int:
    n = 0
    for ch in letter.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def _read_fx_pair(ws: Worksheet) -> tuple[float | None, float | None]:
    usd_kgs = None
    eur_kgs = None
    for r in range(1, min((ws.max_row or 1), 12) + 1):
        label = _norm(ws.cell(r, 1).value).lower().replace(" ", "")
        val = _as_float(ws.cell(r, 3).value)
        if val is None or val <= 0:
            continue
        if "usd/kgs" in label or "usdkg" in label:
            usd_kgs = val
        elif "eur/kgs" in label or "eurkg" in label:
            eur_kgs = val
    return usd_kgs, eur_kgs


def _read_employee_name(ws: Worksheet) -> str:
    # A3 常见为人名；跳过标题行
    for r in (3, 2, 4):
        text = _norm(ws.cell(r, 1).value)
        low = text.lower()
        if not text:
            continue
        if any(
            x in low
            for x in (
                "atlas",
                "payrol",
                "payroll",
                "date:",
                "exchange",
                "cost calculation",
            )
        ):
            continue
        if len(text) >= 3:
            return text
    return ""


def _read_period_label(ws: Worksheet, source_path: Path | None = None) -> str | None:
    for r in range(1, min((ws.max_row or 1), 8) + 1):
        text = _norm(ws.cell(r, 1).value)
        if parse_pay_period_label(text):
            return text
    if source_path:
        if parse_pay_period_label(source_path.name):
            return source_path.name
    return None


def parse_atlas_cost_sheet(
    ws: Worksheet, warnings: list[str] | None = None, *, source_path: Path | None = None
) -> dict[str, Any] | None:
    name = _read_employee_name(ws)
    gross_row = _find_label_row(ws, "gross salary")
    if gross_row is None:
        if warnings is not None:
            warnings.append(f"sheet「{ws.title}」未找到 GROSS salary 行")
        return None

    base_kgs = _eval_or_float(ws, gross_row, 4)
    if base_kgs is None:
        # D = I*C5；尝试 E(USD)*C5
        usd = _eval_or_float(ws, gross_row, 5)
        usd_kgs, _ = _read_fx_pair(ws)
        if usd is not None and usd_kgs:
            base_kgs = usd * usd_kgs
    if base_kgs is None:
        if warnings is not None:
            warnings.append(f"sheet「{ws.title}」无法解析 GROSS salary (KGS)")
        return None

    bonus_row = gross_row
    bonus = _as_float(ws.cell(bonus_row, 6).value)  # Bonus KGS
    if bonus is None:
        bonus_usd = _as_float(ws.cell(bonus_row, 7).value)
        usd_kgs, _ = _read_fx_pair(ws)
        if bonus_usd and usd_kgs:
            bonus = bonus_usd * usd_kgs

    usd_kgs, eur_kgs = _read_fx_pair(ws)
    period_label = _read_period_label(ws, source_path)

    emp: dict[str, Any] = {
        "Name of Employee": name,
        "Employee Name": name,
        "Base Salary": float(base_kgs),
        "_fx_usd_kgs": usd_kgs,
        "_fx_eur_kgs": eur_kgs,
        "_period_label": period_label,
        "_source_sheet": ws.title,
    }
    for key in _ZERO_FILL_KEYS:
        emp[key] = 0
    if bonus is not None and abs(bonus) > 1e-9:
        emp["Bonus"] = float(bonus)

    # 一期不写差旅垫付（样例 PN Expense=0）；若 mapping 显式要求可再开
    mapping = _active_mapping()
    if mapping.get("writeAtlasExpense"):
        adv_row = _find_label_row(ws, "advance for business", "business travel", "expense")
        if adv_row:
            adv_usd = _as_float(ws.cell(adv_row, 9).value)
            if adv_usd is None:
                adv_kgs = _as_float(ws.cell(adv_row, 8).value)
            elif usd_kgs:
                adv_kgs = adv_usd * usd_kgs
            else:
                adv_kgs = None
            if adv_kgs is not None:
                emp["Expense Reimbursment"] = float(adv_kgs)

    return emp


def parse_kyrgyzstan_l_employees(
    ws: Worksheet, warnings: list[str] | None = None
) -> list[dict[str, Any]]:
    header_row, data_start = _kyrgyzstan_l_layout(target=True)
    headers = {
        str(h["key"]): int(h["col"])
        for h in list_qualified_header_cells(ws, header_row)
        if h.get("key") and h.get("col")
    }
    name_col = headers.get("Name of Employee") or COL_NAME
    base_col = headers.get("Base Salary") or 3
    employees: list[dict[str, Any]] = []
    for row in range(data_start, (ws.max_row or data_start) + 1):
        name = _norm(ws.cell(row, name_col).value)
        base = _as_float(ws.cell(row, base_col).value)
        if not name and base is None:
            continue
        if not name:
            continue
        emp: dict[str, Any] = {
            "Name of Employee": name,
            "Employee Name": name,
            "Base Salary": base if base is not None else 0,
        }
        for key in _ZERO_FILL_KEYS:
            col = headers.get(key)
            if col:
                val = _as_float(ws.cell(row, col).value)
                emp[key] = val if val is not None else 0
            else:
                emp[key] = 0
        employees.append(emp)
    if not employees and warnings is not None:
        warnings.append("Kyrgyzstan-L 未解析到员工行")
    return employees


def parse_source_workbook(source_path: Path, warnings: list[str] | None = None) -> list[dict[str, Any]]:
    wb = load_workbook(source_path, data_only=False)
    try:
        if KYRGYZSTAN_L_SHEET in wb.sheetnames and looks_like_kyrgyzstan_l(wb[KYRGYZSTAN_L_SHEET]):
            return parse_kyrgyzstan_l_employees(wb[KYRGYZSTAN_L_SHEET], warnings)

        mapping = _active_mapping()
        src_spec = (
            mapping.get("sourceEmployeeSheet")
            if isinstance(mapping.get("sourceEmployeeSheet"), dict)
            else {}
        )
        preferred = find_sheet_name(list(wb.sheetnames), src_spec)
        sheet_order = list(wb.sheetnames)
        if preferred and preferred in sheet_order:
            sheet_order = [preferred] + [s for s in sheet_order if s != preferred]

        employees: list[dict[str, Any]] = []
        for name in sheet_order:
            ws = wb[name]
            if not looks_like_atlas_cost_sheet(ws):
                continue
            emp = parse_atlas_cost_sheet(ws, warnings, source_path=source_path)
            if emp:
                employees.append(emp)
        if not employees and warnings is not None:
            warnings.append("未识别到 Atlas 成本测算表或 Kyrgyzstan-L")
        return employees
    finally:
        wb.close()


def write_kyrgyzstan_l_period(ws: Worksheet, employees: list[dict[str, Any]]) -> None:
    (from_r, from_c) = _meta_cell("periodFrom", 2, 3)
    (to_r, to_c) = _meta_cell("periodTo", 2, 5)
    label = None
    for emp in employees:
        label = emp.get("_period_label")
        if label:
            break
    bounds = period_bounds_from_label(label) if label else None
    if not bounds:
        return
    start, end = bounds
    ws.cell(from_r, from_c).value = coerce_datetime_for_excel(start)
    ws.cell(from_r, from_c).number_format = _DATE_FMT
    ws.cell(to_r, to_c).value = coerce_datetime_for_excel(end)
    ws.cell(to_r, to_c).number_format = _DATE_FMT


def write_kyrgyzstan_l(ws: Worksheet, employees: list[dict[str, Any]]) -> None:
    header_row, data_start = _kyrgyzstan_l_layout(target=True)
    write_kyrgyzstan_l_period(ws, employees)

    n = len(employees)
    formula_by_col = _kyrgyzstan_l_formula_cols(ws, data_start)
    max_col = max(ws.max_column or 22, 22)

    for i in range(1, n):
        _copy_row_style_and_formula(
            ws,
            data_start,
            data_start + i,
            max_col=max_col,
            l_from=data_start,
            l_to=data_start + i,
        )

    last_keep = data_start + max(n, 1) - 1
    max_row = max(ws.max_row or data_start, data_start + MAX_EMPLOYEES)
    for row in range(data_start, max_row + 1):
        for col in range(1, max_col + 1):
            cell = ws.cell(row, col)
            if type(cell).__name__ == "MergedCell":
                continue
            if row <= last_keep and col in formula_by_col:
                continue
            cell.value = None

    for i in range(n):
        row = data_start + i
        for col, formula in formula_by_col.items():
            cur = ws.cell(row, col).value
            if _cell_formula_text(cur):
                continue
            ws.cell(row, col).value = shift_row_formula(
                formula,
                data_start,
                row,
                target_l_from=data_start,
                target_l_to=row,
                target_l_sheet=KYRGYZSTAN_L_SHEET,
            )

    headers = {
        str(h["key"]): int(h["col"])
        for h in list_qualified_header_cells(ws, header_row)
        if h.get("key") and h.get("col")
    }
    headers_norm = {norm(k): (k, col) for k, col in headers.items()}
    name_col = COL_NAME
    if norm("Name of Employee") in headers_norm:
        name_col = headers_norm[norm("Name of Employee")][1]

    for idx, emp in enumerate(employees):
        row = data_start + idx
        name = _norm(emp.get("Employee Name") or emp.get("Name of Employee"))
        if name:
            ws.cell(row, name_col).value = name
        written: set[int] = set()
        for key, val in emp.items():
            if val is None or str(key).startswith("_"):
                continue
            if norm(key) in {norm(x) for x in _SKIP_WRITE_TARGETS} or norm(key) == norm(
                "Name of Employee"
            ):
                continue
            hit = headers_norm.get(norm(key))
            if not hit:
                continue
            _canon, col = hit
            if col in formula_by_col or col in written:
                continue
            _set_cell_value(ws.cell(row, col), val)
            written.add(col)


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
            dest.alignment = copy(src.alignment)
        text = _cell_formula_text(src.value)
        if text and l_from is not None and l_to is not None:
            dest.value = shift_row_formula(
                text,
                src_row,
                dest_row,
                target_l_from=l_from,
                target_l_to=l_to,
                target_l_sheet=KYRGYZSTAN_L_SHEET,
            )
        elif text:
            dest.value = shift_row_formula(text, src_row, dest_row)
        else:
            dest.value = None


def _retarget_l_refs(formula: str, l_from: int, l_to: int) -> str:
    if l_from == l_to:
        return formula
    return re.sub(
        rf"(Kyrgyzstan-L['\"]?!\$?[A-Za-z]+\$?){l_from}\b",
        rf"\g<1>{l_to}",
        formula,
        flags=re.I,
    )


def _retarget_ee_refs(formula: str, ee_from: int, ee_to: int) -> str:
    if ee_from == ee_to:
        return formula
    return re.sub(
        rf"(Kyrgyzstan EE['\"]?!\$?[A-Za-z]+\$?){ee_from}\b",
        rf"\g<1>{ee_to}",
        formula,
        flags=re.I,
    )


def _kyrgyzstan_ee_layout(ws: Worksheet) -> tuple[int, int, int]:
    """返回 (data_start_row, ee_code_col, ee_name_col)。列按表头；数据行用映射/默认行号。"""
    code_col, name_col = 4, 5
    for r in range(1, 12):
        for c in range(1, 16):
            h = re.sub(r"\s+", " ", _norm(ws.cell(r, c).value)).lower()
            if h == "ee code":
                code_col = c
            elif h == "ee name":
                name_col = c
    data_start = KYRGYZSTAN_EE_DATA_START
    ft = _active_mapping().get("formulaTemplates")
    block = (
        ft.get(KYRGYZSTAN_EE_SHEET)
        if isinstance(ft, dict) and isinstance(ft.get(KYRGYZSTAN_EE_SHEET), dict)
        else {}
    )
    if block.get("defaultExampleRow"):
        data_start = int(block["defaultExampleRow"])
    return data_start, code_col, name_col


def _kyrgyzstan_ee_match_names(
    wb,
    emp: dict[str, Any],
    *,
    l_header_row: int,
    l_data_start: int,
    index: int,
    ee_ws: Worksheet | None,
    ee_row: int,
    ee_name_col: int,
) -> list[str]:
    """匹配 EE Code 用的姓名：源表 + Kyrgyzstan-L + EE Name 列（与 Italy/UAE 一致）。"""
    out: list[str] = []
    seen: set[str] = set()

    def add(raw: Any) -> None:
        if _cell_formula_text(raw):
            return
        text = _norm(raw)
        if not text or text.startswith("="):
            return
        key = text.casefold()
        if key in seen:
            return
        seen.add(key)
        out.append(text)

    for key in ("Employee Name", "Name of Employee", "EE Name"):
        add(emp.get(key))
    if KYRGYZSTAN_L_SHEET in wb.sheetnames:
        l_ws = wb[KYRGYZSTAN_L_SHEET]
        want = {norm("Name of Employee"), norm("Employee Name")}
        for h in list_qualified_header_cells(l_ws, l_header_row):
            k = norm(str(h.get("key") or ""))
            col = h.get("col")
            if k in want and col:
                add(l_ws.cell(l_data_start + index, int(col)).value)
    if ee_ws is not None:
        add(ee_ws.cell(ee_row, ee_name_col).value)
    return out


def expand_kyrgyzstan_employee_rows(wb, employee_count: int) -> None:
    n = max(int(employee_count), 1)
    _, l_data_start = _kyrgyzstan_l_layout(target=True)
    if KYRGYZSTAN_SHEET in wb.sheetnames:
        main = wb[KYRGYZSTAN_SHEET]
        for i in range(1, n):
            dest = KYRGYZSTAN_DATA_START + i
            l_row = l_data_start + i
            ee_row = KYRGYZSTAN_EE_DATA_START + i
            _copy_row_style_and_formula(
                main,
                KYRGYZSTAN_DATA_START,
                dest,
                max_col=40,
                l_from=l_data_start,
                l_to=l_row,
            )
            for c in range(1, 41):
                cell = main.cell(dest, c)
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    cell.value = _retarget_l_refs(cell.value, l_data_start, l_row)
                    cell.value = _retarget_ee_refs(cell.value, KYRGYZSTAN_EE_DATA_START, ee_row)

    if KYRGYZSTAN_EE_SHEET in wb.sheetnames:
        ee = wb[KYRGYZSTAN_EE_SHEET]
        for i in range(1, n):
            dest = KYRGYZSTAN_EE_DATA_START + i
            l_row = l_data_start + i
            _copy_row_style_and_formula(
                ee,
                KYRGYZSTAN_EE_DATA_START,
                dest,
                max_col=40,
                l_from=l_data_start,
                l_to=l_row,
            )
            for c in range(1, 41):
                cell = ee.cell(dest, c)
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    cell.value = _retarget_l_refs(cell.value, l_data_start, l_row)


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


def apply_kyrgyzstan_ee_codes(
    wb,
    employees: list[dict[str, Any]],
    *,
    employee_directory: list[dict[str, Any]] | None = None,
    pn_meta: PnMeta | dict[str, Any] | None = None,
) -> list[str]:
    """员工库匹配 EE Code → 直接写 Kyrgyzstan EE 表 EE Code 列；同步写 Kyrgyzstan-L!A。"""
    warnings: list[str] = []
    directory = list(employee_directory or [])
    _, l_data_start = _kyrgyzstan_l_layout(target=True)
    client_code = _pn_customer_id(pn_meta)

    ee_ws = wb[KYRGYZSTAN_EE_SHEET] if KYRGYZSTAN_EE_SHEET in wb.sheetnames else None
    ee_data_start, ee_code_col, ee_name_col = KYRGYZSTAN_EE_DATA_START, 4, 5
    l_header_row, _ = _kyrgyzstan_l_layout(target=True)
    if ee_ws is not None:
        ee_data_start, ee_code_col, ee_name_col = _kyrgyzstan_ee_layout(ee_ws)

    for i, emp in enumerate(employees):
        emp.pop("No. of EE", None)
        emp.pop("_ee_code", None)

        row = ee_data_start + i
        excel_names = _kyrgyzstan_ee_match_names(
            wb,
            emp,
            l_header_row=l_header_row,
            l_data_start=l_data_start,
            index=i,
            ee_ws=ee_ws,
            ee_row=row,
            ee_name_col=ee_name_col,
        )
        code, warn = match_ee_code(excel_names, directory)
        if code:
            emp["No. of EE"] = code
            emp["_ee_code"] = code
            if KYRGYZSTAN_L_SHEET in wb.sheetnames:
                wb[KYRGYZSTAN_L_SHEET].cell(l_data_start + i, COL_EE_CODE).value = code
        elif KYRGYZSTAN_L_SHEET in wb.sheetnames:
            l_cell = wb[KYRGYZSTAN_L_SHEET].cell(l_data_start + i, COL_EE_CODE)
            if not _cell_formula_text(l_cell.value):
                l_cell.value = None
        if warn:
            warnings.append(f"Kyrgyzstan EE 第{i + 1}人：{warn}")

        if ee_ws is None:
            continue
        if client_code and not _cell_formula_text(ee_ws.cell(row, 2).value):
            ee_ws.cell(row, 2).value = client_code
        # EE Code：员工库工号直接覆盖母版公式（含旧 =L!A 引用）
        ee_ws.cell(row, ee_code_col).value = code
    return warnings


def _find_pn_row_by_label(ws: Worksheet, keyword: str, col: int = 1) -> int | None:
    key = keyword.lower()
    for row in range(1, (ws.max_row or 0) + 1):
        v = ws.cell(row, col).value
        if isinstance(v, str) and key in v.lower():
            return row
    return None


def fit_kyrgyzstan_pn_employees(wb, employee_count: int) -> dict[str, Any]:
    fx_row = _find_pn_row_by_label(wb[PN_SHEET], "FX rate") if PN_SHEET in wb.sheetnames else None
    return {"fx_row": fx_row or 28, "employee_count": employee_count}


def apply_fx(
    wb,
    employees: list[dict[str, Any]],
    *,
    fill_fx: bool = True,
    convert_mapping: dict | None = None,
) -> tuple[float | None, float | None, list[dict[str, Any]]]:
    """写 PN!B28 USD/KGS、PN!B29 EUR/KGS；优先供应商账单。"""
    if not fill_fx or PN_SHEET not in wb.sheetnames:
        return None, None, []
    from fx_policy import fx_policy, make_pn_fx_provenance

    mapping = convert_mapping if isinstance(convert_mapping, dict) else _active_mapping()
    policy = fx_policy(mapping)
    mode = str(policy.get("mode") or "vendor_bill").strip().lower()
    if mode == "none":
        return None, None, []

    usd_kgs = None
    eur_kgs = None
    for emp in employees:
        if usd_kgs is None:
            usd_kgs = _as_float(emp.get("_fx_usd_kgs"))
        if eur_kgs is None:
            eur_kgs = _as_float(emp.get("_fx_eur_kgs"))

    provenances: list[dict[str, Any]] = []
    pn = wb[PN_SHEET]
    usd_row = _find_pn_row_by_label(pn, "USD/KGS") or 28
    eur_row = _find_pn_row_by_label(pn, "EUR/KGS") or 29

    if mode in ("vendor_bill", "shared_fact") and usd_kgs and usd_kgs > 0:
        pn.cell(usd_row, 2).value = float(usd_kgs)
        provenances.append(
            make_pn_fx_provenance(
                PN_SHEET,
                usd_row,
                2,
                mapping,
                usd_kgs,
                write_source="source:vendor",
                fx_source="source:USD/KGS",
            )
        )
    elif mode != "none":
        try:
            usd_kgs = float(get_usd_rate("KGS"))
            pn.cell(usd_row, 2).value = usd_kgs
            provenances.append(
                make_pn_fx_provenance(
                    PN_SHEET,
                    usd_row,
                    2,
                    mapping,
                    usd_kgs,
                    write_source="api",
                    fx_source="api:KGS",
                )
            )
        except Exception:
            pass

    if mode in ("vendor_bill", "shared_fact") and eur_kgs and eur_kgs > 0:
        pn.cell(eur_row, 2).value = float(eur_kgs)
        provenances.append(
            make_pn_fx_provenance(
                PN_SHEET,
                eur_row,
                2,
                mapping,
                eur_kgs,
                write_source="source:vendor",
                fx_source="source:EUR/KGS",
            )
        )
    elif mode != "none" and eur_kgs is None:
        # 无供应商 EUR/KGS 时不硬猜；保留母版空值
        pass

    return usd_kgs, eur_kgs, [p for p in provenances if p]


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
    _ACTIVE_MAPPING = resolve_convert_mapping("kyrgyzstan_payroll_calc", convert_mapping)
    try:
        source_path = Path(source_path).resolve()
        output_path = Path(output_path).resolve()
        template_path = Path(template_path).resolve()
        if not source_path.is_file():
            raise FileNotFoundError(f"源文件不存在: {source_path}")
        if not template_path.is_file():
            raise FileNotFoundError(f"母版不存在: {template_path}")

        parse_warnings: list[str] = []
        employees = parse_source_workbook(source_path, parse_warnings)
        if not employees:
            raise ValueError("未解析到任何吉尔吉斯员工行")
        sort_employees_by_code(employees, employee_directory)

        output_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(template_path, output_path)

        wb = load_workbook(output_path)
        try:
            if KYRGYZSTAN_L_SHEET not in wb.sheetnames:
                raise ValueError(f"母版缺少 sheet「{KYRGYZSTAN_L_SHEET}」")

            write_kyrgyzstan_l(wb[KYRGYZSTAN_L_SHEET], employees)
            expand_kyrgyzstan_employee_rows(wb, len(employees))
            pn_layout = fit_kyrgyzstan_pn_employees(wb, len(employees))

            usd_kgs, eur_kgs, pn_fx_writes = apply_fx(
                wb, employees, fill_fx=fill_fx, convert_mapping=_ACTIVE_MAPPING
            )

            applied_pn = None
            if pn_meta is not None:
                applied_pn = apply_pn_meta(
                    wb,
                    pn_meta,
                    registry_dir=registry_dir or output_path.parent,
                    reserve_invoice_number=True,
                )

            ee_warnings = apply_kyrgyzstan_ee_codes(
                wb,
                employees,
                employee_directory=employee_directory,
                pn_meta=applied_pn or pn_meta,
            )

            apply_luckysheet_compat(wb, pn_sheet=PN_SHEET)
            wb.save(output_path)
        finally:
            wb.close()

        postprocess_converted_xlsx(output_path)
        warnings = merge_warnings(parse_warnings, ee_warnings)
        result = {
            "engine_id": "kyrgyzstan_payroll_calc",
            "region": "Kyrgyzstan",
            "employee_count": len(employees),
            "employees": [
                {
                    "name": e.get("Name of Employee") or e.get("Employee Name"),
                    "ee_code": e.get("_ee_code") or e.get("No. of EE"),
                    "base_salary": e.get("Base Salary"),
                }
                for e in employees
            ],
            "fx_usd_kgs": usd_kgs,
            "fx_eur_kgs": eur_kgs,
            "pn_fx_writes": pn_fx_writes,
            "pn_layout": pn_layout,
            "pn_meta": applied_pn.to_dict() if applied_pn else None,
            "warnings": warnings,
            "output": str(output_path),
        }
        extra = sanity_check_convert_result(result)
        if extra:
            result["warnings"] = merge_warnings(warnings, extra)
        return result
    finally:
        _ACTIVE_MAPPING = None


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Kyrgyzstan Atlas cost calc → PN")
    parser.add_argument("source", type=Path)
    parser.add_argument("-o", "--output", type=Path, default=None)
    parser.add_argument("-t", "--template", type=Path, default=DEFAULT_TEMPLATE)
    parser.add_argument("--no-fx", action="store_true")
    args = parser.parse_args(argv)
    out = args.output or Path(f"out_kyrgyzstan_{datetime.now():%Y%m%d_%H%M%S}.xlsx")
    result = convert(
        args.source,
        out,
        args.template,
        fill_fx=not args.no_fx,
    )
    print(result)
    for w in result.get("warnings") or []:
        print(f"[warn] {w}", file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
