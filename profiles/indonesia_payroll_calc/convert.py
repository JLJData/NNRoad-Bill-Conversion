# -*- coding: utf-8 -*-
"""
Indonesia：Link Compliance 工资明细 Excel（或已是 Indonesia-L）→ Indonesia PN

用法:
  python -m profiles.indonesia_payroll_calc.convert <源.xlsx> [-o 输出.xlsx] [-t 母版.xlsx]

原则：
- PN / Indonesia / Indonesia EE 以母版公式为准；只写 Indonesia-L 数据与必要元数据。
- EE Code 必须来自员工库匹配，禁止沿用供应商账单工号。
- Cash Advance 等 PDF 发票项一期不自动写入（人工 / 后续 pdf_ingest）。
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

from bill_convert.convert_checks import merge_warnings, parse_cell_ref, sanity_check_convert_result
from bill_convert.formula_copy import shift_row_formula
from bill_convert.formula_layout import sort_employees_by_code
from convert_mapping import find_sheet_name, resolve_convert_mapping
from fx_rate import get_indonesia_pn_fx_rate
from pn_meta import PnMeta, apply_pn_meta
from profiles.tw_payroll_calc.convert import match_ee_code
from region_templates import get_region_template
from xlsx_convert_utils import coerce_datetime_for_excel
from xlsx_luckysheet_compat import apply_luckysheet_compat
from xlsx_postprocess import postprocess_converted_xlsx

DEFAULT_TEMPLATE = get_region_template("Indonesia")

INDONESIA_L_SHEET = "Indonesia-L"
INDONESIA_SHEET = "Indonesia"
INDONESIA_EE_SHEET = "Indonesia EE"
PN_SHEET = "PN"

INDONESIA_L_HEADER_ROW = 7
INDONESIA_L_DATA_START = 8
INDONESIA_DATA_START = 9
INDONESIA_EE_DATA_START = 10
MAX_EMPLOYEES = 20
_DATE_FMT = "yyyy/m/d"

COL_EE_CODE = 1
COL_NAME = 2

# 目标 Indonesia-L 字段 → 默认列（表头匹配失败时兜底）
_TARGET_FIELD_COLS: dict[str, int] = {
    "name": 2,
    "base": 3,
    "ot": 4,
    "salary_adj": 5,
    "bonus": 6,
    "other": 7,
    "expense": 8,
    "pph21": 10,
    # jht_ee(K) 母版为公式，默认不写
    "jp_ee": 12,
    "bpjs_ee": 13,
    "jkk": 15,
    "jht_er": 16,
    "jp_er": 17,
    "jkm": 18,
    "bpjs_er": 19,
}

# 源 Link Compliance 字段 → 写入员工 dict 时使用的标准键（与 target 表头对齐）
_SOURCE_TO_TARGET_KEYS: dict[str, str] = {
    "base": "Base Salary",
    "ot": "OT",
    "salary_adj": "Salary Adjusment",
    "bonus": "Bonus",
    "other": "Other ",
    "expense": "Expense Reimbursment",
    "pph21": "INCOME TAX (PPH21)",
    "jp_ee": "JP 1%",
    "bpjs_ee": "BPJS KESEHATAN 1%",
    "jkk": "JKK 0.24%",
    "jht_er": "JHT 3.7%",
    "jp_er": "JP 2%",
    "jkm": "JKM  0.3%",
    "bpjs_er": "HEALTH INSURANCE (BPJS KESEHATAN) 4%",
}

# 缺省写 0 的薪资可选列（与样例 PN 一致：OT 留空，Adj/Bonus/Other/Expense 写 0）
_ZERO_FILL_KEYS = ("Salary Adjusment", "Bonus", "Other ", "Expense Reimbursment")

# 源表常缺、不必告警的字段
_OPTIONAL_SOURCE_FIELDS = frozenset({"ot", "salary_adj", "bonus", "other", "expense", "jht_ee"})

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
        else resolve_convert_mapping("indonesia_payroll_calc", None)
    )


def _norm(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).replace("\n", " ").strip()
    # 源表头常见 =UPPER("Employee Name")
    m = re.match(r'^=\s*UPPER\s*\(\s*"([^"]+)"\s*\)\s*$', text, flags=re.I)
    if m:
        return m.group(1).strip()
    m2 = re.match(r"^=\s*UPPER\s*\(\s*'([^']+)'\s*\)\s*$", text, flags=re.I)
    if m2:
        return m2.group(1).strip()
    return text


def _as_float(value: Any) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "").replace("\xa0", "").replace("IDR", "")
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
        if cell.number_format in (None, "General", ""):
            pass
        elif "yy" in str(cell.number_format).lower() or "h:" in str(cell.number_format).lower():
            cell.number_format = "General"
    elif isinstance(value, int):
        cell.value = value
        fmt = str(cell.number_format or "")
        if "yy" in fmt.lower() or "h:" in fmt.lower() or "m/d" in fmt.lower():
            cell.number_format = "General"
    else:
        cell.value = value


def _indonesia_l_layout(*, target: bool = False) -> tuple[int, int]:
    mapping = _active_mapping()
    key = "targetL" if target else "sourceEmployeeSheet"
    spec = mapping.get(key) if isinstance(mapping.get(key), dict) else {}
    if not spec and target:
        spec = (
            mapping.get("sourceEmployeeSheet")
            if isinstance(mapping.get("sourceEmployeeSheet"), dict)
            else {}
        )
    if target:
        header = int(spec.get("headerRow") or INDONESIA_L_HEADER_ROW)
        data_start = int(spec.get("dataStartRow") or INDONESIA_L_DATA_START)
    else:
        header = int(spec.get("headerRow") or 6)
        data_start = int(spec.get("dataStartRow") or 8)
    return header, data_start


def _field_header_names(field: str) -> list[str]:
    mapping = _active_mapping()
    custom = mapping.get("fieldHeaders") if isinstance(mapping.get("fieldHeaders"), dict) else {}
    raw = custom.get(field)
    out: list[str] = []
    seen: set[str] = set()

    def add(name: str) -> None:
        s = str(name).strip()
        if not s:
            return
        key = s.lower()
        if key in seen:
            return
        seen.add(key)
        out.append(s)

    if isinstance(raw, list):
        for x in raw:
            if x is not None:
                add(str(x))
    return out


def _header_map_single(ws: Worksheet, header_row: int) -> dict[str, int]:
    out: dict[str, int] = {}
    for col in range(1, (ws.max_column or 1) + 1):
        h = _norm(ws.cell(header_row, col).value)
        if h and h not in out:
            out[h] = col
    return out


def _header_map_dual(ws: Worksheet, header_row: int, sub_header_row: int | None) -> dict[str, int]:
    """双行表头：子行优先，空则用父行。"""
    out: dict[str, int] = {}
    max_col = ws.max_column or 1
    for col in range(1, max_col + 1):
        h = ""
        if sub_header_row:
            h = _norm(ws.cell(sub_header_row, col).value)
        if not h:
            h = _norm(ws.cell(header_row, col).value)
        if h and h not in out:
            out[h] = col
    return out


def _find_col(headers: dict[str, int], names: list[str]) -> int | None:
    lower = {k.lower(): v for k, v in headers.items()}
    for name in names:
        n = _norm(name)
        if not n:
            continue
        if n in headers:
            return headers[n]
        hit = lower.get(n.lower())
        if hit is not None:
            return hit
        # 宽松：去多余空格
        compact = re.sub(r"\s+", " ", n).lower()
        for k, col in headers.items():
            if re.sub(r"\s+", " ", k).lower() == compact:
                return col
    return None


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
        r"[\s\-']*(\d{2,4})",
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
    """样例 PN 使用整月起止（非 7–30 日截段）。"""
    parsed = parse_pay_period_label(label)
    if not parsed:
        return None
    year, month = parsed
    start = date(year, month, 1)
    end = date(year, month, calendar.monthrange(year, month)[1])
    return start, end


def _read_vendor_fx(ws: Worksheet) -> float | None:
    for row in range(1, min((ws.max_row or 1), 10) + 1):
        for col in range(1, min((ws.max_column or 1), 6) + 1):
            label = _norm(ws.cell(row, col).value).lower()
            if "exchange rate" in label or "1usd" in label.replace(" ", ""):
                for dc in (col + 1, col + 2, 3):
                    fx = _as_float(ws.cell(row, dc).value)
                    if fx is not None and fx > 0:
                        return fx
    return None


def _read_vendor_period_label(ws: Worksheet) -> str | None:
    for row in range(1, min((ws.max_row or 1), 8) + 1):
        for col in range(1, min((ws.max_column or 1), 6) + 1):
            text = _norm(ws.cell(row, col).value)
            if text.lower().startswith("period") or "september" in text.lower() or re.search(
                r"\b20\d{2}\b", text
            ):
                if "period" in text.lower() or parse_pay_period_label(text):
                    return text
    return None


def _indonesia_l_formula_cols(ws: Worksheet, data_start: int) -> dict[int, str]:
    out: dict[int, str] = {}
    for col in range(1, (ws.max_column or 1) + 1):
        text = _cell_formula_text(ws.cell(data_start, col).value)
        if text:
            out[col] = text
    return out


def looks_like_indonesia_l(ws: Worksheet) -> bool:
    headers = _header_map_single(ws, INDONESIA_L_HEADER_ROW)
    return bool(_find_col(headers, ["Name of Employee", "Base Salary"]))


def parse_link_compliance_employees(ws: Worksheet, warnings: list[str] | None = None) -> list[dict[str, Any]]:
    mapping = _active_mapping()
    src = (
        mapping.get("sourceEmployeeSheet")
        if isinstance(mapping.get("sourceEmployeeSheet"), dict)
        else {}
    )
    header_row = int(src.get("headerRow") or 6)
    sub_row = src.get("subHeaderRow")
    sub_header_row = int(sub_row) if sub_row is not None else header_row + 1
    data_start = int(src.get("dataStartRow") or 8)
    headers = _header_map_dual(ws, header_row, sub_header_row)

    name_col = _find_col(headers, _field_header_names("name") or ["EMPLOYEE NAME", "Employee Name"])
    if name_col is None:
        raise ValueError(f"未找到员工姓名列，表头={list(headers.keys())[:20]}")

    field_cols: dict[str, int] = {}
    for field in _SOURCE_TO_TARGET_KEYS:
        col = _find_col(headers, _field_header_names(field))
        if col is not None:
            field_cols[field] = col
        elif warnings is not None and field not in _OPTIONAL_SOURCE_FIELDS:
            warnings.append(f"源表未匹配字段「{field}」，将跳过")

    period_label = _read_vendor_period_label(ws)
    bounds = period_bounds_from_label(period_label) if period_label else None
    period_from = bounds[0] if bounds else None
    period_to = bounds[1] if bounds else None
    fx_rate = _read_vendor_fx(ws)

    employees: list[dict[str, Any]] = []
    max_row = max(ws.max_row or data_start, data_start)
    for row in range(data_start, max_row + 1):
        name = _norm(ws.cell(row, name_col).value)
        if not name or name.upper() == "TOTAL":
            continue
        # 跳过公式姓名
        if _cell_formula_text(ws.cell(row, name_col).value):
            continue
        emp: dict[str, Any] = {
            "Employee Name": name,
            "Name of Employee": name,
            "_period_from": period_from,
            "_period_to": period_to,
            "_period_label": period_label,
            "_fx_rate": fx_rate,
            "From": period_from,
            "To": period_to,
        }
        for field, col in field_cols.items():
            val = ws.cell(row, col).value
            if _cell_formula_text(val):
                # 尝试用 data_only 不可得；若是简单引用则仍可读缓存——此处跳过公式格
                # openpyxl 无 data_only=False 时公式单元格无数值缓存时 .value 仍是公式字符串
                continue
            num = _as_float(val)
            key = _SOURCE_TO_TARGET_KEYS[field]
            emp[key] = num if num is not None else val
        for zkey in _ZERO_FILL_KEYS:
            emp.setdefault(zkey, 0)
        employees.append(emp)
    return employees


def parse_indonesia_l_employees(ws: Worksheet, warnings: list[str] | None = None) -> list[dict[str, Any]]:
    header_row, data_start = _indonesia_l_layout(target=True)
    headers = _header_map_single(ws, header_row)
    name_col = _find_col(headers, _field_header_names("name") or ["Name of Employee"]) or COL_NAME
    (pf_r, pf_c) = _meta_cell("periodFrom", 2, 3)
    (pt_r, pt_c) = _meta_cell("periodTo", 2, 5)
    (fx_r, fx_c) = _meta_cell("fxRate", 4, 3)
    period_from = ws.cell(pf_r, pf_c).value
    period_to = ws.cell(pt_r, pt_c).value
    fx_rate = _as_float(ws.cell(fx_r, fx_c).value)

    employees: list[dict[str, Any]] = []
    max_row = max(ws.max_row or data_start, data_start)
    for row in range(data_start, max_row + 1):
        name = _norm(ws.cell(row, name_col).value)
        if not name:
            continue
        emp: dict[str, Any] = {
            "Employee Name": name,
            "Name of Employee": name,
            "_period_from": period_from,
            "_period_to": period_to,
            "_fx_rate": fx_rate,
            "From": period_from,
            "To": period_to,
        }
        # 注意：不读取 A 列供应商/旧 EE Code 作为权威工号
        for field, target_key in _SOURCE_TO_TARGET_KEYS.items():
            names = _field_header_names(field) or [target_key]
            col = _find_col(headers, names) or _TARGET_FIELD_COLS.get(field)
            if col is None:
                continue
            val = ws.cell(row, col).value
            if _cell_formula_text(val) or val is None or val == "":
                continue
            emp[target_key] = val
        for zkey in _ZERO_FILL_KEYS:
            emp.setdefault(zkey, 0)
        employees.append(emp)
    if not employees and warnings is not None:
        warnings.append("Indonesia-L 未解析到员工行")
    return employees


def _enrich_employees_from_data_only(
    source_path: Path,
    employees: list[dict[str, Any]],
    sheet_name: str,
) -> list[dict[str, Any]]:
    """用 Excel 缓存值补全公式列（如 BPJS 4%）。"""
    if not employees:
        return employees
    try:
        wb_val = load_workbook(source_path, data_only=True)
    except Exception:
        return employees
    try:
        if sheet_name not in wb_val.sheetnames:
            return employees
        ws = wb_val[sheet_name]
        mapping = _active_mapping()
        src_spec = (
            mapping.get("sourceEmployeeSheet")
            if isinstance(mapping.get("sourceEmployeeSheet"), dict)
            else {}
        )
        header_row = int(src_spec.get("headerRow") or 6)
        sub_row = src_spec.get("subHeaderRow")
        sub_header_row = int(sub_row) if sub_row is not None else header_row + 1
        data_start = int(src_spec.get("dataStartRow") or 8)
        headers = _header_map_dual(ws, header_row, sub_header_row)
        name_col = _find_col(headers, _field_header_names("name") or ["EMPLOYEE NAME", "Employee Name"])
        if name_col is None:
            return employees
        field_cols: dict[str, int] = {}
        for field in _SOURCE_TO_TARGET_KEYS:
            col = _find_col(headers, _field_header_names(field))
            if col is not None:
                field_cols[field] = col
        fx_rate = _read_vendor_fx(ws)
        idx = 0
        max_row = max(ws.max_row or data_start, data_start)
        for row in range(data_start, max_row + 1):
            name = _norm(ws.cell(row, name_col).value)
            if not name or name.upper() == "TOTAL":
                continue
            if idx >= len(employees):
                break
            emp = employees[idx]
            if fx_rate is not None:
                emp["_fx_rate"] = fx_rate
            for field, col in field_cols.items():
                key = _SOURCE_TO_TARGET_KEYS[field]
                existing = _as_float(emp.get(key))
                if existing is not None:
                    continue
                num = _as_float(ws.cell(row, col).value)
                if num is not None:
                    emp[key] = num
            idx += 1
        return employees
    finally:
        wb_val.close()


def parse_source_workbook(source_path: Path, warnings: list[str] | None = None) -> list[dict[str, Any]]:
    wb = load_workbook(source_path, data_only=False)
    try:
        if INDONESIA_L_SHEET in wb.sheetnames and looks_like_indonesia_l(wb[INDONESIA_L_SHEET]):
            return parse_indonesia_l_employees(wb[INDONESIA_L_SHEET], warnings)

        mapping = _active_mapping()
        src_spec = (
            mapping.get("sourceEmployeeSheet")
            if isinstance(mapping.get("sourceEmployeeSheet"), dict)
            else {}
        )
        sheet_name = find_sheet_name(list(wb.sheetnames), src_spec)
        if not sheet_name:
            sheet_name = wb.sheetnames[0] if wb.sheetnames else None
        if not sheet_name:
            raise ValueError("源工作簿无 sheet")
        employees = parse_link_compliance_employees(wb[sheet_name], warnings)
    finally:
        wb.close()

    return _enrich_employees_from_data_only(source_path, employees, sheet_name)


def write_indonesia_l_period(ws: Worksheet, employees: list[dict[str, Any]]) -> None:
    if not employees:
        return
    emp0 = employees[0]
    start = emp0.get("_period_from") or emp0.get("From")
    end = emp0.get("_period_to") or emp0.get("To")
    if start is None and end is None:
        bounds = period_bounds_from_label(emp0.get("_period_label"))
        if bounds:
            start, end = bounds
    (pf_r, pf_c) = _meta_cell("periodFrom", 2, 3)
    (pt_r, pt_c) = _meta_cell("periodTo", 2, 5)
    for (row, col), val in (((pf_r, pf_c), start), ((pt_r, pt_c), end)):
        if val is None:
            continue
        cell = ws.cell(row, col)
        dt = coerce_datetime_for_excel(val)
        if dt is not None:
            cell.value = dt
            cell.number_format = _DATE_FMT
        elif isinstance(val, str) and re.match(r"^\d{4}/\d{1,2}/\d{1,2}$", val.strip()):
            try:
                y, m, d = [int(x) for x in val.strip().split("/")]
                cell.value = datetime(y, m, d)
                cell.number_format = _DATE_FMT
            except ValueError:
                cell.value = val
        else:
            cell.value = val


def write_indonesia_l_fx(ws: Worksheet, employees: list[dict[str, Any]], fx: float | None) -> None:
    (fx_r, fx_c) = _meta_cell("fxRate", 4, 3)
    rate = fx
    if rate is None and employees:
        rate = _as_float(employees[0].get("_fx_rate"))
    if rate is None:
        return
    ws.cell(fx_r, fx_c).value = float(rate)


def write_indonesia_l(ws: Worksheet, employees: list[dict[str, Any]]) -> None:
    """只覆盖数据列，母版公式列保留/扩行复制。不写 EE Code（由员工库步骤写入）。"""
    header_row, data_start = _indonesia_l_layout(target=True)
    write_indonesia_l_period(ws, employees)

    n = len(employees)
    formula_by_col = _indonesia_l_formula_cols(ws, data_start)
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
                target_l_sheet=INDONESIA_L_SHEET,
            )

    headers = _header_map_single(ws, header_row)
    name_col = _find_col(headers, _field_header_names("name") or ["Name of Employee"]) or COL_NAME

    target_cols: dict[str, int] = {}
    for field, target_key in _SOURCE_TO_TARGET_KEYS.items():
        names = _field_header_names(field) or [target_key]
        col = _find_col(headers, names) or _TARGET_FIELD_COLS.get(field)
        if col is not None:
            target_cols[target_key] = col

    for idx, emp in enumerate(employees):
        row = data_start + idx
        name = _norm(emp.get("Employee Name") or emp.get("Name of Employee"))
        if name:
            ws.cell(row, name_col).value = name
        written: set[int] = set()
        for key, col in target_cols.items():
            if col in formula_by_col or col in written:
                continue
            if key not in emp:
                continue
            val = emp[key]
            if val is None:
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
                target_l_sheet=INDONESIA_L_SHEET,
            )
        elif text:
            dest.value = shift_row_formula(text, src_row, dest_row)
        else:
            dest.value = None


def _retarget_l_refs(formula: str, l_from: int, l_to: int) -> str:
    if l_from == l_to:
        return formula
    return re.sub(
        rf"(Indonesia-L['\"]?!\$?[A-Za-z]+\$?){l_from}\b",
        rf"\g<1>{l_to}",
        formula,
        flags=re.I,
    )


def _retarget_ee_refs(formula: str, ee_from: int, ee_to: int) -> str:
    if ee_from == ee_to:
        return formula
    return re.sub(
        rf"(Indonesia EE['\"]?!\$?[A-Za-z]+\$?){ee_from}\b",
        rf"\g<1>{ee_to}",
        formula,
        flags=re.I,
    )


def expand_indonesia_employee_rows(wb, employee_count: int) -> None:
    n = max(int(employee_count), 1)
    _, l_data_start = _indonesia_l_layout(target=True)
    if INDONESIA_SHEET in wb.sheetnames:
        main = wb[INDONESIA_SHEET]
        for i in range(1, n):
            dest = INDONESIA_DATA_START + i
            l_row = l_data_start + i
            ee_row = INDONESIA_EE_DATA_START + i
            _copy_row_style_and_formula(
                main,
                INDONESIA_DATA_START,
                dest,
                max_col=40,
                l_from=l_data_start,
                l_to=l_row,
            )
            for c in range(1, 41):
                cell = main.cell(dest, c)
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    cell.value = _retarget_l_refs(cell.value, l_data_start, l_row)
                    cell.value = _retarget_ee_refs(cell.value, INDONESIA_EE_DATA_START, ee_row)

    if INDONESIA_EE_SHEET in wb.sheetnames:
        ee = wb[INDONESIA_EE_SHEET]
        for i in range(1, n):
            dest = INDONESIA_EE_DATA_START + i
            l_row = l_data_start + i
            _copy_row_style_and_formula(
                ee,
                INDONESIA_EE_DATA_START,
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


def apply_indonesia_ee_codes(
    wb,
    employees: list[dict[str, Any]],
    *,
    employee_directory: list[dict[str, Any]] | None = None,
    pn_meta: PnMeta | dict[str, Any] | None = None,
) -> list[str]:
    """员工库匹配 EE Code → 仅写 Indonesia-L!A；Indonesia EE!D 为母版公式引用 L!A。

    禁止使用供应商账单中的工号。
    """
    warnings: list[str] = []
    directory = list(employee_directory or [])
    _, l_data_start = _indonesia_l_layout(target=True)
    client_code = _pn_customer_id(pn_meta)

    for i, emp in enumerate(employees):
        # 清除可能误带的供应商工号键
        emp.pop("No. of EE", None)
        emp.pop("_ee_code", None)

        name = _norm(emp.get("Employee Name") or emp.get("Name of Employee"))
        code, warn = match_ee_code([name] if name else [], directory)
        if code:
            emp["No. of EE"] = code
            emp["_ee_code"] = code
            if INDONESIA_L_SHEET in wb.sheetnames:
                wb[INDONESIA_L_SHEET].cell(l_data_start + i, COL_EE_CODE).value = code
        if warn:
            warnings.append(f"Indonesia EE 第{i + 1}人：{warn}")

        if INDONESIA_EE_SHEET not in wb.sheetnames:
            continue
        ee = wb[INDONESIA_EE_SHEET]
        row = INDONESIA_EE_DATA_START + i
        # Client Code：仅当母版该格不是公式时才写
        if client_code and not _cell_formula_text(ee.cell(row, 2).value):
            ee.cell(row, 2).value = client_code
        # EE Code 列若是公式（引用 L!A）则不覆盖
        if code and not _cell_formula_text(ee.cell(row, 4).value):
            ee.cell(row, 4).value = code
    return warnings


def _find_pn_row_by_label(ws: Worksheet, keyword: str, col: int = 1) -> int | None:
    key = keyword.lower()
    for row in range(1, (ws.max_row or 0) + 1):
        v = ws.cell(row, col).value
        if isinstance(v, str) and key in v.lower():
            return row
    return None


def fit_indonesia_pn_employees(wb, employee_count: int) -> dict[str, Any]:
    fx_row = _find_pn_row_by_label(wb[PN_SHEET], "FX rate") if PN_SHEET in wb.sheetnames else None
    return {"fx_row": fx_row or 28, "employee_count": employee_count}


def apply_fx(
    wb,
    employees: list[dict[str, Any]],
    *,
    fill_fx: bool = True,
    convert_mapping: dict | None = None,
) -> tuple[float | None, str | None]:
    """优先供应商账单汇率写 Indonesia-L!C4；否则 API IDR。"""
    if not fill_fx or INDONESIA_L_SHEET not in wb.sheetnames:
        return None, None
    from fx_policy import fx_policy

    mapping = convert_mapping if isinstance(convert_mapping, dict) else _active_mapping()
    policy = fx_policy(mapping)
    mode = str(policy.get("mode") or "vendor_bill").strip().lower()
    if mode == "none":
        return None, None

    vendor_fx = None
    if employees:
        vendor_fx = _as_float(employees[0].get("_fx_rate"))
    fx = None
    source = None
    if mode in ("vendor_bill", "shared_fact") and vendor_fx is not None and vendor_fx > 0:
        fx = float(vendor_fx)
        source = "source:vendor"
    elif mode == "none":
        return None, None
    else:
        try:
            fx = float(get_indonesia_pn_fx_rate())
            source = "api:IDR"
        except Exception as exc:
            if vendor_fx is not None and vendor_fx > 0:
                fx = float(vendor_fx)
                source = f"source:vendor(fallback after api fail: {exc})"
            else:
                raise

    if fx is not None:
        write_indonesia_l_fx(wb[INDONESIA_L_SHEET], employees, fx)
    return fx, source


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
    _ACTIVE_MAPPING = resolve_convert_mapping("indonesia_payroll_calc", convert_mapping)
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
            raise ValueError("未解析到任何印尼员工行")
        sort_employees_by_code(employees, employee_directory)

        output_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(template_path, output_path)
        wb = load_workbook(output_path)
        warnings: list[str] = list(parse_warnings)
        fx = None
        fx_source = None
        applied_pn = None
        try:
            if INDONESIA_L_SHEET not in wb.sheetnames:
                raise ValueError(f"母版缺少 {INDONESIA_L_SHEET}")
            write_indonesia_l(wb[INDONESIA_L_SHEET], employees)
            expand_indonesia_employee_rows(wb, len(employees))
            pn_layout = fit_indonesia_pn_employees(wb, len(employees))
            if len(employees) > 1:
                warnings.append(
                    f"Indonesia PN 多人明细行扩行暂定：已扩 Indonesia/Indonesia EE（{len(employees)} 人），请人工核对 PN"
                )
            try:
                fx, fx_source = apply_fx(
                    wb, employees, fill_fx=fill_fx, convert_mapping=_ACTIVE_MAPPING
                )
            except Exception as exc:
                warnings.append(f"写入汇率失败: {exc}")

            if pn_meta is not None:
                applied_pn = apply_pn_meta(
                    wb,
                    pn_meta,
                    registry_dir=registry_dir or output_path.parent,
                    reserve_invoice_number=True,
                )
            warnings.extend(
                apply_indonesia_ee_codes(
                    wb,
                    employees,
                    employee_directory=employee_directory,
                    pn_meta=applied_pn or pn_meta,
                )
            )

            apply_luckysheet_compat(wb, pn_sheet=PN_SHEET)
            wb.save(output_path)
        finally:
            wb.close()

        postprocess_converted_xlsx(output_path)
        warnings = merge_warnings(
            warnings,
            sanity_check_convert_result({"employee_count": len(employees), "warnings": warnings}),
        )
        return {
            "ok": True,
            "engine_id": "indonesia_payroll_calc",
            "region": "Indonesia",
            "output": str(output_path),
            "employee_count": len(employees),
            "employee_names": [
                _norm(e.get("Employee Name") or e.get("Name of Employee")) for e in employees
            ],
            "fx_rate": fx,
            "fx_source": fx_source,
            "warnings": warnings,
            "pn_meta": applied_pn.to_dict() if applied_pn else None,
            "pn_layout": pn_layout,
        }
    finally:
        _ACTIVE_MAPPING = None


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Link Compliance / Indonesia-L → Indonesia PN")
    parser.add_argument("source", type=Path)
    parser.add_argument("-o", "--output", type=Path, default=None)
    parser.add_argument("-t", "--template", type=Path, default=None)
    parser.add_argument("--no-fx", action="store_true")
    args = parser.parse_args(argv)
    source = args.source.resolve()
    output = (args.output or source.with_name(f"PN_Indonesia_{source.stem}.xlsx")).resolve()
    template = (args.template or DEFAULT_TEMPLATE).resolve()
    result = convert(source, output, template, fill_fx=not args.no_fx)
    print(result)
    return 0 if result.get("ok") else 1


if __name__ == "__main__":
    sys.exit(main())
