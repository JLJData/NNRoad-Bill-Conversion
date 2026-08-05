# -*- coding: utf-8 -*-
"""
SafeGuard (SGWI) Italy Payroll Excel → Italy-L（profile: safeguard_italy）

源 sheet「Calculation」：
  姓名取供应商账单姓名列（通常 A 列，如 Matteo Cupi）
  Vacation Accruals → Italy-L「Vacation Leave」；「Vacation Accruals」置 0
  Fee Min 不写死：由后台 mapping.italyFeeMin 在引擎阶段写入

用法:
  python -m pdf_ingest.profiles.safeguard_italy <源.xlsx> [-o 输出.xlsx]
"""
from __future__ import annotations

import argparse
import re
import shutil
import sys
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from pn_meta import PnMeta
from region_templates import get_region_template

ITALY_L_SHEET = "Italy-L"
SRC_SHEET_CANDIDATES = ("Calculation", "calculation")

# 源表头 → Italy-L 表头（不含动态 Salary / Vacation 特殊处理）
# 目标名须与 Italy-L _norm 后表头一致（去首尾空白）
_STATIC_HEADER_MAP: dict[str, str] = {
    "PO Number": "PO Number",
    "SGWI %age Markup": "Fee %age Markup",
    "Fee %age Markup": "Fee %age Markup",
    "SGWI Minimum Currency": "Fee Minimum Currency",
    "Fee Minimum Currency": "Fee Minimum Currency",
    "Applied SGWI Minimum Currency": "Applied SGWI Minimum Currency",
    "Currency": "Currency",
    "Social Cost": "Social Cost",
    "Social contributions for accruals": "Social contributions for accruals",
    "13th and 14th Accrual": "13th and 14th Accrual",
    "Monthly Permitted Leave (Permessi Hours Ex-Fs)": "Monthly Permitted Leave (Permessi Hours Ex-Fs)",
    "Monthly Permitted Leave ROL Accrual": "Monthly Permitted Leave ROL Accrual",
    "Unemployment Fund (TFR) Accrual": "Unemployment Fund (TFR) Accrual",
    "Overheads": "Overheads",
    "Bilateral Trade Association": "Bilateral Trade Association",
    "Social Contributions Employer (INPS) April": "Social Contributions Employer (INPS) April",
    "Unemployment Fund (TFR) Accrual April": "Unemployment Fund (TFR) Accrual April",
}


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
    if not text or text == "-":
        return None
    try:
        return float(text)
    except ValueError:
        return None


def _money(value: Any) -> float | None:
    """金额统一两位小数，避免源表浮点尾巴把 PN USD 合计顶到差 0.01。"""
    num = _as_float(value)
    if num is None:
        return None
    return round(num, 2)


def _header_map(ws: Worksheet, header_row: int) -> dict[str, int]:
    out: dict[str, int] = {}
    for col in range(1, (ws.max_column or 1) + 1):
        h = _norm(ws.cell(header_row, col).value)
        if h and h not in out:
            out[h] = col
    return out


def _find_header_row(ws: Worksheet) -> int:
    for r in range(1, min(20, (ws.max_row or 1) + 1)):
        labels = {_norm(ws.cell(r, c).value).lower() for c in range(1, min(15, (ws.max_column or 1) + 1))}
        if "employee id" in labels or "employee name" in labels:
            return r
        if "po number" in labels and ("currency" in labels or "sgwi min" in labels):
            return r
    raise ValueError("未找到 SafeGuard 员工表头行（Employee ID / Employee Name）")


def _find_sheet(wb) -> Worksheet:
    for name in SRC_SHEET_CANDIDATES:
        if name in wb.sheetnames:
            return wb[name]
    # 兜底：第一张
    return wb[wb.sheetnames[0]]


def looks_like_safeguard_italy(path: Path) -> bool:
    try:
        wb = load_workbook(path, read_only=True, data_only=False)
    except Exception:
        return False
    try:
        ws = _find_sheet(wb)
        blob = []
        for r in range(1, min(12, (ws.max_row or 1) + 1)):
            for c in range(1, min(12, (ws.max_column or 1) + 1)):
                v = _norm(ws.cell(r, c).value).lower()
                if v:
                    blob.append(v)
        text = " ".join(blob)
        return ("italy" in text or "animal equality" in text) and (
            "sgwi" in text or "safeguard" in text or "pay period" in text
        )
    finally:
        wb.close()


def looks_like_italy_l_workbook(path: Path) -> bool:
    from profiles.italy_payroll_calc.convert import looks_like_italy_l_workbook as _fn

    return _fn(path)


def _read_meta(ws: Worksheet, header_row: int) -> dict[str, Any]:
    meta: dict[str, Any] = {}
    for r in range(1, header_row):
        label = _norm(ws.cell(r, 2).value).lower()
        val = ws.cell(r, 3).value
        if label == "customer":
            meta["_customer"] = _norm(val)
        elif label == "location":
            meta["_location"] = _norm(val)
        elif label == "pay period":
            meta["_pay_period"] = val
            meta["Pay Period"] = val
    # Invoice ID: 常见 R2 K 标签 / R3 K 号码
    if _norm(ws.cell(2, 11).value).lower().startswith("invoice"):
        meta["_invoice_id"] = ws.cell(3, 11).value
    return meta


def _salary_src_header(headers: dict[str, int]) -> str | None:
    for h in headers:
        if re.search(r"\bsalary\b", h, flags=re.I):
            return h
    return None


def _vac_accrual_header(headers: dict[str, int]) -> str | None:
    for h in headers:
        if re.search(r"vacation\s+accrual", h, flags=re.I):
            return h
    return None


def parse_safeguard_italy_excel(excel_path: Path) -> list[dict[str, Any]]:
    path = Path(excel_path).resolve()
    wb = load_workbook(path, data_only=True)
    try:
        ws = _find_sheet(wb)
        header_row = _find_header_row(ws)
        headers = _header_map(ws, header_row)
        meta = _read_meta(ws, header_row)

        # 姓名：优先 A 列非空；否则 Employee Name
        name_col_a = 1
        name_header_col = headers.get("Employee Name")
        id_col = headers.get("Employee ID")

        salary_h = _salary_src_header(headers)
        vac_h = _vac_accrual_header(headers)

        employees: list[dict[str, Any]] = []
        data_start = header_row + 1
        for row in range(data_start, (ws.max_row or data_start) + 1):
            name = _norm(ws.cell(row, name_col_a).value)
            if not name and name_header_col:
                name = _norm(ws.cell(row, name_header_col).value)
            if not name:
                # 有的版式姓名不在 A，但有 Employee ID
                if id_col and ws.cell(row, id_col).value not in (None, ""):
                    name = _norm(ws.cell(row, id_col).value)
                else:
                    continue
            low = name.lower()
            if "invoice total" in low or low.startswith("sgwi"):
                continue

            emp: dict[str, Any] = dict(meta)
            emp["Employee Name"] = name
            if id_col:
                emp["_employee_id"] = ws.cell(row, id_col).value

            for src_h, col in headers.items():
                src_key = _norm(src_h)
                tgt = _STATIC_HEADER_MAP.get(src_h) or _STATIC_HEADER_MAP.get(src_key)
                if not tgt:
                    continue
                tgt = _norm(tgt)
                val = ws.cell(row, col).value
                if val is None or val == "":
                    continue
                num = _money(val)
                if src_key in (
                    "po number",
                    "currency",
                    "sgwi minimum currency",
                    "fee minimum currency",
                    "applied sgwi minimum currency",
                ):
                    emp[tgt] = val
                elif num is not None:
                    emp[tgt] = num
                else:
                    emp[tgt] = val

            if salary_h:
                from profiles.italy_payroll_calc.convert import salary_header_for_period

                title = _norm(salary_header_for_period(meta.get("_pay_period")) or salary_h)
                emp[title] = _money(ws.cell(row, headers[salary_h]).value)

            # Vacation Accruals → Vacation Leave；Accruals 置 0
            if vac_h:
                vac_val = _money(ws.cell(row, headers[vac_h]).value) or 0.0
                emp["Vacation Leave"] = vac_val
                emp["Vacation Accruals"] = 0.0
            else:
                emp.setdefault("Vacation Accruals", 0.0)

            # Fee Min 留给 mapping；源 SGWI Min 仅作参考字段
            if "SGWI Min" in headers:
                emp["_source_fee_min"] = _money(ws.cell(row, headers["SGWI Min"]).value)
            if "Fee Min" in headers:
                emp["_source_fee_min"] = _money(ws.cell(row, headers["Fee Min"]).value)

            employees.append(emp)

        if not employees:
            raise ValueError(f"未解析到员工行: {path.name}")
        return employees
    finally:
        wb.close()


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
    """SafeGuard Calculation Excel → 含 Italy-L 的源表。"""
    del pn_meta, registry_dir, fill_fx
    from convert_mapping import resolve_convert_mapping
    from profiles.italy_payroll_calc.convert import set_fee_min, write_italy_l

    mapping_in = dict(convert_mapping) if isinstance(convert_mapping, dict) else {}
    mapping_in.setdefault("pdfProfileId", "safeguard_italy")
    mapping = resolve_convert_mapping("italy_payroll_calc", mapping_in)

    paths = [Path(p).resolve() for p in excel_paths]
    if not paths:
        raise ValueError("未提供 Excel")
    for p in paths:
        if not p.is_file():
            raise FileNotFoundError(f"Excel 不存在: {p}")

    output_path = Path(output_path).resolve()
    italy_l_flags = [looks_like_italy_l_workbook(p) for p in paths]
    if all(italy_l_flags):
        if len(paths) > 1:
            raise ValueError("多份已是 Italy-L 的 Excel 无法自动合并，请只传一份")
        output_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(paths[0], output_path)
        return {
            "ok": True,
            "profile_id": "safeguard_italy",
            "region": "Italy",
            "source_kind": "excel_italy_l",
            "output": str(output_path),
            "employee_count": None,
            "parsed": [],
            "warnings": ["源表已是 Italy-L，已原样用作转换输入"],
            "fx_rate": None,
            "pn_meta": None,
        }

    employees: list[dict[str, Any]] = []
    warnings: list[str] = []
    for p in paths:
        if looks_like_italy_l_workbook(p):
            raise ValueError("请不要混传已成型 Italy-L 与 SafeGuard 源账单")
        if not looks_like_safeguard_italy(p):
            warnings.append(f"文件可能不是 SafeGuard Italy 账单，仍尝试解析: {p.name}")
        employees.extend(parse_safeguard_italy_excel(p))

    tpl = (template_path or get_region_template("Italy")).resolve()
    if not tpl.is_file():
        raise FileNotFoundError(f"Italy 母版不存在: {tpl}")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(tpl, output_path)
    wb = load_workbook(output_path)
    try:
        if ITALY_L_SHEET not in wb.sheetnames:
            raise ValueError(f"母版缺少 {ITALY_L_SHEET}")
        # 临时激活 mapping，供 set_fee_min 读取
        import profiles.italy_payroll_calc.convert as italy_mod

        prev = italy_mod._ACTIVE_MAPPING
        italy_mod._ACTIVE_MAPPING = mapping
        try:
            write_italy_l(wb[ITALY_L_SHEET], employees)
            set_fee_min(wb, employees)
        finally:
            italy_mod._ACTIVE_MAPPING = prev
        wb.save(output_path)
    finally:
        wb.close()

    return {
        "ok": True,
        "profile_id": "safeguard_italy",
        "region": "Italy",
        "source_kind": "excel",
        "output": str(output_path),
        "employee_count": len(employees),
        "parsed": [
            {
                "employee_name": e.get("Employee Name"),
                "employee_id": e.get("_employee_id"),
                "pay_period": e.get("Pay Period") or e.get("_pay_period"),
                "invoice_id": e.get("_invoice_id"),
            }
            for e in employees
        ],
        "warnings": warnings,
        "fx_rate": None,
        "pn_meta": None,
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
    paths = [Path(p).resolve() for p in source_paths]
    if not paths:
        raise ValueError("未提供源文件")
    excels = [p for p in paths if p.suffix.lower() in (".xlsx", ".xlsm", ".xls")]
    other = [p for p in paths if p not in excels]
    if other:
        raise ValueError(
            "safeguard_italy 目前仅支持 Excel 源账单；"
            f"不支持: {[p.name for p in other]}"
        )
    return convert_excels(
        excels,
        output_path,
        template_path=template_path,
        pn_meta=pn_meta,
        registry_dir=registry_dir,
        fill_fx=fill_fx,
        convert_mapping=convert_mapping,
    )


def convert_pdf(*_args, **_kwargs):
    raise RuntimeError("safeguard_italy 主源为 Excel，不支持 PDF；请上传 SGWI Payroll xlsx")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="SafeGuard Italy Excel → Italy-L")
    parser.add_argument("source", type=Path)
    parser.add_argument("-o", "--output", type=Path, default=None)
    parser.add_argument("-t", "--template", type=Path, default=None)
    args = parser.parse_args(argv)
    source = args.source.resolve()
    output = (args.output or source.with_name(f"ItalyL_{source.stem}.xlsx")).resolve()
    result = convert_excels([source], output, template_path=args.template)
    print(result)
    return 0 if result.get("ok") else 1


if __name__ == "__main__":
    sys.exit(main())
