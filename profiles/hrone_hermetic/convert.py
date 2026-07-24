# -*- coding: utf-8 -*-
"""
HROne Co., Ltd. + Hermetic 账单 → PN 转换脚本

用法:
  python -m profiles.hrone_hermetic.convert <原始账单.xlsx> [-o 输出.xlsx] [-t 母版.xlsx]

默认母版: templates/china/template.xlsx
（含 PN / China / China EE / China-L 公式结构）
"""
from __future__ import annotations

import argparse
import re
import shutil
import sys
from copy import copy
from pathlib import Path
from typing import Any

from pn_meta import PnMeta, apply_pn_meta

from openpyxl import load_workbook
from openpyxl.cell.cell import Cell
from openpyxl.worksheet.worksheet import Worksheet

from fx_rate import fetch_usd_rates, get_china_pn_fx_rate
from region_templates import get_region_template
from xlsx_richtext_fix import migrate_inlinestr_richtext_to_shared_strings

DEFAULT_TEMPLATE = get_region_template("China")

CALC_SHEET_NAMES = ("计算结果",)
OTHER_FEE_NAMES = ("Other Fee",)
PAYMENT_NOTICE_NAMES = ("S-Payment Notice",)

CHINA_DATA_START_ROW = 9
CHINA_EE_DATA_START_ROW = 10
CHINA_L_DATA_START_ROW = 2
MAX_EMPLOYEES = 10


def norm(text: Any) -> str:
    if text is None:
        return ""
    return str(text).replace("\uFEFF", "").strip()


def find_sheet(wb, candidates: tuple[str, ...]) -> str | None:
    names = {n: n for n in wb.sheetnames}
    for c in candidates:
        if c in names:
            return c
    lower = {n.lower(): n for n in wb.sheetnames}
    for c in candidates:
        if c.lower() in lower:
            return lower[c.lower()]
    for n in wb.sheetnames:
        for c in candidates:
            if c in n:
                return n
    return None


def build_header_map(ws: Worksheet, header_row: int = 1) -> dict[str, int]:
    mapping: dict[str, int] = {}
    for col in range(1, (ws.max_column or 0) + 1):
        key = norm(ws.cell(header_row, col).value)
        if key and key not in mapping:
            mapping[key] = col
    return mapping


def clean_value(value: Any) -> Any:
    if value is None:
        return None
    s = norm(value)
    if s in ("", "#N/A", "#REF!", "#VALUE!", "-"):
        return None
    if isinstance(value, (int, float)):
        return value
    try:
        if re.fullmatch(r"-?\d+(\.\d+)?", s.replace(",", "")):
            compact = s.replace(",", "").replace("，", "").replace(" ", "")
            return float(compact) if "." in compact else int(compact)
    except ValueError:
        pass
    return value


def parse_expense_count(text: Any) -> int:
    if text is None:
        return 0
    m = re.search(r"(\d+)", str(text))
    return int(m.group(1)) if m else 0


def extract_by_label_col_a(ws: Worksheet, label: str, value_col: int) -> Any:
    target = norm(label)
    for row in range(1, (ws.max_row or 0) + 1):
        if norm(ws.cell(row, 1).value) == target:
            return clean_value(ws.cell(row, value_col).value)
    return None


def read_calc_employees(ws: Worksheet) -> list[dict[str, Any]]:
    headers = build_header_map(ws, 1)
    if "姓名" not in headers:
        raise ValueError("「计算结果」sheet 第 1 行须包含表头「姓名」")

    employees: list[dict[str, Any]] = []
    for row in range(2, (ws.max_row or 0) + 1):
        name = clean_value(ws.cell(row, headers["姓名"]).value)
        if name is None:
            continue
        record: dict[str, Any] = {}
        for hdr, col in headers.items():
            record[hdr] = clean_value(ws.cell(row, col).value)
        employees.append(record)

    if not employees:
        raise ValueError("「计算结果」中未找到有效员工行（姓名非空）")
    if len(employees) > MAX_EMPLOYEES:
        raise ValueError(f"员工数 {len(employees)} 超过模板上限 {MAX_EMPLOYEES}")
    return employees


def clear_china_l_data(ws: Worksheet, from_row: int = CHINA_L_DATA_START_ROW) -> None:
    max_row = max(ws.max_row or from_row, from_row + MAX_EMPLOYEES)
    max_col = ws.max_column or 1
    for row in range(from_row, max_row + 1):
        for col in range(1, max_col + 1):
            ws.cell(row, col).value = None


def write_china_l(ws: Worksheet, employees: list[dict[str, Any]]) -> None:
    target_headers = build_header_map(ws, 1)
    if not target_headers:
        raise ValueError("China-L 第 1 行表头为空")

    clear_china_l_data(ws)
    for idx, emp in enumerate(employees):
        row = CHINA_L_DATA_START_ROW + idx
        for hdr, val in emp.items():
            col = target_headers.get(hdr)
            if col is not None and val is not None:
                ws.cell(row, col).value = val


def shift_formula(formula: str, from_row: int, to_row: int, china_l_from: int, china_l_to: int) -> str:
    if not formula or not isinstance(formula, str) or not formula.startswith("="):
        return formula

    s = re.sub(
        r"'China-L'!([A-Z]{1,3})(\d+)",
        lambda m: f"'China-L'!{m.group(1)}{china_l_to}",
        formula,
    )
    s = re.sub(
        rf"(?<!\$)(?<![A-Z])([A-Z]{{1,3}}){from_row}(?!\d)",
        lambda m: f"{m.group(1)}{to_row}",
        s,
    )
    return s


def copy_row_formulas(ws: Worksheet, from_row: int, to_row: int, china_l_from: int, china_l_to: int) -> None:
    max_col = ws.max_column or 1
    for col in range(1, max_col + 1):
        src = ws.cell(from_row, col)
        dst = ws.cell(to_row, col)
        if src.data_type == "f" and isinstance(src.value, str):
            dst.value = shift_formula(src.value, from_row, to_row, china_l_from, china_l_to)
        elif src.value is not None and src.data_type != "f":
            dst.value = copy(src.value)


def copy_cell_style(src: Cell, dst: Cell) -> None:
    if src.has_style:
        dst.font = copy(src.font)
        dst.border = copy(src.border)
        dst.fill = copy(src.fill)
        dst.number_format = copy(src.number_format)
        dst.protection = copy(src.protection)
        dst.alignment = copy(src.alignment)


def expand_china_employee_rows(ws: Worksheet, employee_count: int) -> None:
    if employee_count <= 1:
        return
    template_row = CHINA_DATA_START_ROW
    for i in range(1, employee_count):
        target_row = template_row + i
        china_l_row = CHINA_L_DATA_START_ROW + i
        copy_row_formulas(ws, template_row, target_row, CHINA_L_DATA_START_ROW, china_l_row)


def expand_china_ee_rows(ws: Worksheet, employee_count: int) -> None:
    if employee_count <= 1:
        return
    template_row = CHINA_EE_DATA_START_ROW
    for i in range(1, employee_count):
        target_row = template_row + i
        china_l_row = CHINA_L_DATA_START_ROW + i
        china_row = CHINA_DATA_START_ROW + i
        copy_row_formulas(ws, template_row, target_row, CHINA_L_DATA_START_ROW, china_l_row)
        for col in range(1, (ws.max_column or 0) + 1):
            cell = ws.cell(target_row, col)
            if cell.data_type == "f" and isinstance(cell.value, str):
                cell.value = re.sub(
                    rf"China!([A-Z]+){CHINA_DATA_START_ROW}(?!\d)",
                    lambda m: f"China!{m.group(1)}{china_row}",
                    cell.value,
                )


def apply_china_specials(
    ws: Worksheet,
    employee_count: int,
    expense_count: int,
    other_amount: Any,
) -> None:
    for i in range(employee_count):
        row = CHINA_DATA_START_ROW + i
        ws.cell(row, 9).value = f"=40*PN!$B$29*{expense_count}"
        if other_amount is not None:
            ws.cell(row, 10).value = other_amount


def parse_source(source_path: Path) -> dict[str, Any]:
    wb = load_workbook(source_path, data_only=True, read_only=True)
    calc_name = find_sheet(wb, CALC_SHEET_NAMES)
    if not calc_name:
        wb.close()
        raise ValueError(f"未找到 sheet「计算结果」，现有: {wb.sheetnames}")

    employees = read_calc_employees(wb[calc_name])

    other_name = find_sheet(wb, OTHER_FEE_NAMES)
    payment_name = find_sheet(wb, PAYMENT_NOTICE_NAMES)

    other_amount = None
    expense_count = 0
    if other_name:
        other_ws = wb[other_name]
        other_amount = extract_by_label_col_a(other_ws, "其他", 4)
        expense_text = extract_by_label_col_a(other_ws, "报销服务费", 2)
        expense_count = parse_expense_count(expense_text)

    vendor_fx_rate = None
    if payment_name:
        vendor_fx_rate = clean_value(wb[payment_name]["C49"].value)

    wb.close()

    try:
        rates = fetch_usd_rates()
        fx_rate = get_china_pn_fx_rate(rates)
        fx_source = "api:CNY"
    except RuntimeError:
        fx_rate = vendor_fx_rate
        fx_source = "vendor:C49" if vendor_fx_rate is not None else "none"

    return {
        "employees": employees,
        "other_amount": other_amount,
        "expense_count": expense_count,
        "fx_rate": fx_rate,
        "fx_source": fx_source,
        "vendor_fx_rate": vendor_fx_rate,
    }


def convert(
    source_path: Path,
    output_path: Path,
    template_path: Path,
    *,
    pn_meta: PnMeta | dict[str, Any] | None = None,
    registry_dir: Path | None = None,
) -> dict[str, Any]:
    if not template_path.is_file():
        raise FileNotFoundError(f"母版不存在: {template_path}")
    if not source_path.is_file():
        raise FileNotFoundError(f"原始账单不存在: {source_path}")

    parsed = parse_source(source_path)
    employees = parsed["employees"]

    output_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(template_path, output_path)

    wb = load_workbook(output_path, rich_text=True)
    required = ("China-L", "China", "China EE", "PN")
    for name in required:
        if name not in wb.sheetnames:
            wb.close()
            raise ValueError(f"母版缺少 sheet「{name}」，现有: {wb.sheetnames}")

    applied_pn: PnMeta | None = None
    if pn_meta is not None:
        applied_pn = apply_pn_meta(
            wb,
            pn_meta,
            registry_dir=registry_dir or output_path.parent,
            reserve_invoice_number=True,
        )

    write_china_l(wb["China-L"], employees)

    if parsed["fx_rate"] is not None:
        wb["PN"]["B29"].value = parsed["fx_rate"]

    expand_china_employee_rows(wb["China"], len(employees))
    apply_china_specials(
        wb["China"],
        len(employees),
        parsed["expense_count"],
        parsed["other_amount"],
    )
    expand_china_ee_rows(wb["China EE"], len(employees))

    wb.save(output_path)
    wb.close()
    migrate_inlinestr_richtext_to_shared_strings(output_path)

    return {
        "employee_count": len(employees),
        "employee_names": [e.get("姓名") for e in employees],
        "fx_rate": parsed["fx_rate"],
        "fx_source": parsed.get("fx_source"),
        "vendor_fx_rate": parsed.get("vendor_fx_rate"),
        "other_amount": parsed["other_amount"],
        "expense_count": parsed["expense_count"],
        "output": str(output_path),
        "pn_meta": applied_pn.to_dict() if applied_pn else None,
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="China 供应商账单 → PN 自动转换")
    parser.add_argument("source", type=Path, help="原始账单 Excel 路径")
    parser.add_argument("-o", "--output", type=Path, help="输出 PN 路径（默认同目录 PN_auto_*.xlsx）")
    parser.add_argument("-t", "--template", type=Path, default=DEFAULT_TEMPLATE, help="PN 母版路径")
    args = parser.parse_args(argv)

    source = args.source.resolve()
    if args.output:
        output = args.output.resolve()
    else:
        output = source.parent / f"PN_auto_{source.stem}.xlsx"

    try:
        result = convert(source, output, args.template.resolve())
    except Exception as exc:
        print(f"转换失败: {exc}", file=sys.stderr)
        return 1

    print("转换完成")
    print(f"  输出: {result['output']}")
    print(f"  员工: {result['employee_count']} 人 → {result['employee_names']}")
    print(f"  汇率 PN!B29: {result['fx_rate']} ({result.get('fx_source')})")
    if result.get("vendor_fx_rate") is not None and result.get("fx_source") == "api:CNY":
        print(f"  供应商 C49 参考: {result['vendor_fx_rate']}")
    print(f"  Other Fee → China!J*: {result['other_amount']}")
    print(f"  报销笔数 → For Expense: {result['expense_count']}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
