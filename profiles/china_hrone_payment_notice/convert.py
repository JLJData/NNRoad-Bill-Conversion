# -*- coding: utf-8 -*-
"""
HROne HK Payment Notice → China-L / China-L (2)
引擎 china_hrone_payment_notice

只按 mapping.lSheetCopies 填两张 L（默认 S-Payslip → China-L，
S-Payroll Report → China-L (2)）。不改 PN / China / China EE 公式；
汇率从供应商 S-Payment Notice!C51（可扫「汇率」标签）写入 PN 的 FX 格。
EE Code：按姓名（含拼音）匹配客户员工库，直接写入 China EE 的 EE Code 列（与 TW/UK 相同）。

用法:
  python -m profiles.china_hrone_payment_notice.convert <源.xlsx> [-o 输出.xlsx] [-t 母版.xlsx]
"""
from __future__ import annotations

import argparse
import re
import shutil
import sys
import tempfile
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.worksheet.worksheet import Worksheet

from bill_convert.l_sheet_copy import apply_l_sheet_copies
from convert_mapping import resolve_convert_mapping
from fx_policy import fx_policy, make_pn_fx_provenance
from pn_meta import PnMeta, apply_pn_meta
from profiles.tw_payroll_calc.convert import match_ee_code
from region_templates import get_region_template
from xlsx_convert_utils import clean_value, norm
from xlsx_luckysheet_compat import apply_luckysheet_compat
from xlsx_postprocess import postprocess_converted_xlsx
from xlsx_unlock import collect_unlock_passwords, unlock_xlsx

ENGINE_ID = "china_hrone_payment_notice"
DEFAULT_TEMPLATE = get_region_template("China")
PN_SHEET = "PN"
CHINA_SHEET = "China"
CHINA_EE_SHEET = "China EE"
CHINA_DATA_START_ROW = 9
CHINA_EE_DATA_START_ROW = 10
CHINA_CODE_COL = 3  # China!C
CHINA_NAME_COL = 2
PAYMENT_NOTICE_NAMES = ("S-Payment Notice", "Payment Notice", "付款通知")
_FX_LABEL_KEYS = (
    "汇率",
    "兑换率",
    "exchange rate",
    "fx rate",
    "usd/cny",
    "usd to cny",
    "美元汇率",
)
_SIMPLE_CELL_REF_RE = re.compile(r"^\$?([A-Za-z]+)\$?(\d+)$")
_CNY_USD_MIN, _CNY_USD_MAX = 4.0, 12.0

_ACTIVE_MAPPING: dict[str, Any] | None = None


def _active_mapping() -> dict[str, Any]:
    return (
        _ACTIVE_MAPPING
        if isinstance(_ACTIVE_MAPPING, dict)
        else resolve_convert_mapping(ENGINE_ID, None)
    )


def _find_sheet(wb, candidates: tuple[str, ...]) -> str | None:
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


def _payment_sheet_hints(mapping: dict[str, Any]) -> tuple[str, ...]:
    policy = fx_policy(mapping)
    hints = policy.get("sourceSheetHints") if isinstance(policy.get("sourceSheetHints"), list) else []
    names = [str(x).strip() for x in hints if str(x).strip()]
    return tuple(names) if names else PAYMENT_NOTICE_NAMES


def _payment_fx_cell(mapping: dict[str, Any]) -> str:
    policy = fx_policy(mapping)
    cell = str(policy.get("sourceCell") or "").strip().upper()
    return cell or "C51"


def _plausible_cny_per_usd(value: float) -> bool:
    return _CNY_USD_MIN <= float(value) <= _CNY_USD_MAX


def _coerce_fx_rate(value: Any) -> float | None:
    cleaned = clean_value(value)
    if isinstance(cleaned, (int, float)) and not isinstance(cleaned, bool):
        n = float(cleaned)
        return n if n > 0 else None
    text = norm(value)
    if not text:
        return None
    if text.startswith("="):
        body = text[1:].replace(" ", "").replace(",", "")
        try:
            n = float(body)
            return n if n > 0 else None
        except ValueError:
            pass
        if "*" in body and all(ch not in body for ch in "/()"):
            parts = body.split("*")
            if len(parts) == 2:
                try:
                    n = float(parts[0]) * float(parts[1])
                    return n if n > 0 else None
                except ValueError:
                    return None
        return None
    return None


def _fx_from_cell(ws: Worksheet, addr: str) -> float | None:
    try:
        raw = ws[addr].value
    except Exception:
        return None
    fx = _coerce_fx_rate(raw)
    if fx is not None:
        return fx
    text = norm(raw)
    if text.startswith("="):
        body = text[1:].replace("$", "").replace(" ", "")
        m = _SIMPLE_CELL_REF_RE.fullmatch(body)
        if m:
            try:
                return _coerce_fx_rate(ws[f"{m.group(1)}{m.group(2)}"].value)
            except Exception:
                return None
    return None


def _scan_fx_by_label(ws: Worksheet) -> tuple[float | None, str | None]:
    max_r = min(ws.max_row or 0, 80)
    max_c = min(ws.max_column or 0, 12)
    for row in range(1, max_r + 1):
        for col in range(1, max_c + 1):
            label = norm(ws.cell(row, col).value).lower()
            if not label or not any(k in label for k in _FX_LABEL_KEYS):
                continue
            for dc in (1, 2, 3):
                fx = _coerce_fx_rate(ws.cell(row, col + dc).value)
                if fx is not None and _plausible_cny_per_usd(fx):
                    return fx, f"label:{ws.cell(row, col).value}"
    return None, None


def _read_vendor_fx(
    src_wb,
    unlocked_path: Path,
    mapping: dict[str, Any],
) -> tuple[float | None, str | None]:
    """供应商账单汇率：先读映射格（默认 S-Payment Notice!C51），再扫「汇率」标签。"""
    hints = _payment_sheet_hints(mapping)
    cell = _payment_fx_cell(mapping)
    name = _find_sheet(src_wb, hints)
    fx = None
    source = None
    if name:
        fx = _fx_from_cell(src_wb[name], cell)
        if fx is not None:
            source = f"vendor:{name}!{cell}"
        if fx is None:
            fx, lab = _scan_fx_by_label(src_wb[name])
            if fx is not None:
                source = f"vendor:{name}!{lab}"

    if fx is None:
        try:
            wb = load_workbook(unlocked_path, data_only=False)
            try:
                name = _find_sheet(wb, hints)
                if name:
                    fx = _fx_from_cell(wb[name], cell)
                    if fx is not None:
                        source = f"vendor:{name}!{cell}(formula)"
                    if fx is None:
                        fx, lab = _scan_fx_by_label(wb[name])
                        if fx is not None:
                            source = f"vendor:{name}!{lab}"
            finally:
                wb.close()
        except Exception:
            pass

    if fx is not None and not _plausible_cny_per_usd(fx):
        return None, None
    return fx, source


def _find_pn_fx_row(ws: Worksheet) -> int | None:
    for col in (1, 2, 3):
        for row in range(1, (ws.max_row or 0) + 1):
            v = ws.cell(row, col).value
            if isinstance(v, str) and "fx rate" in v.lower():
                return row
    for col in (1, 2, 3):
        for row in range(1, (ws.max_row or 0) + 1):
            v = ws.cell(row, col).value
            if isinstance(v, str) and "汇率" in v:
                return row
    return None


def _apply_vendor_fx(dst_wb, src_wb, unlocked_path: Path, mapping: dict[str, Any]) -> dict[str, Any]:
    policy = fx_policy(mapping)
    mode = str(policy.get("mode") or "none").strip().lower()
    out: dict[str, Any] = {
        "fx_rate": None,
        "fx_source": "none",
        "fx_row": None,
        "pn_fx_write": None,
        "warnings": [],
    }
    if mode == "none":
        return out

    vendor_fx, vendor_src = _read_vendor_fx(src_wb, unlocked_path, mapping)
    if vendor_fx is None:
        cell = _payment_fx_cell(mapping)
        out["warnings"].append(
            f"供应商账单未读到汇率（S-Payment Notice!{cell} 或「汇率」标签），PN FX 格未改"
        )
        return out

    out["fx_rate"] = vendor_fx
    out["fx_source"] = vendor_src or "vendor_bill"
    if PN_SHEET not in dst_wb.sheetnames:
        out["warnings"].append("母版没有 PN 表，汇率已读到但未写入")
        return out

    fx_row = _find_pn_fx_row(dst_wb[PN_SHEET])
    if fx_row is None:
        out["warnings"].append("母版 PN 未找到 FX rate 行，汇率已读到但未写入")
        return out

    dst_wb[PN_SHEET].cell(fx_row, 2).value = vendor_fx
    out["fx_row"] = fx_row
    out["pn_fx_write"] = make_pn_fx_provenance(
        PN_SHEET,
        fx_row,
        2,
        mapping,
        float(vendor_fx),
        write_source="vendor",
        fx_source=out["fx_source"],
    )
    return out


def names_from_copies(copies: list[dict[str, Any]] | None) -> list[str]:
    names: list[str] = []
    for item in copies or []:
        for n in item.get("employeeNames") or []:
            if n and n not in names:
                names.append(n)
    return names


def _is_formula(value: Any) -> bool:
    return isinstance(value, str) and value.startswith("=")


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


def _header_text(value: Any) -> str:
    return re.sub(r"\s+", " ", norm(value)).strip().lower()


def _scan_china_ee_layout(ws: Worksheet) -> tuple[int, int, int, int]:
    """返回 (data_start_row, ee_code_col, ee_name_col, client_code_col)。按表头定位，兼容空 EE Code 格。"""
    code_col, name_col, client_col = 4, 5, 2
    header_row: int | None = None
    for r in range(1, 12):
        for c in range(1, 16):
            h = _header_text(ws.cell(r, c).value)
            if h == "ee code":
                code_col = c
                header_row = r if header_row is None else min(header_row, r)
            elif h == "ee name":
                name_col = c
                header_row = r if header_row is None else min(header_row, r)
            elif h == "client code":
                client_col = c
    start_from = (header_row or 3) + 2  # 跳过两行表头（EE Code 与 From/To 合并）
    for r in range(start_from, 22):
        a = norm(ws.cell(r, 1).value)
        if a and "eor" in a.lower():
            continue
        for col in (client_col, code_col, name_col):
            cell = ws.cell(r, col)
            if isinstance(cell, MergedCell):
                continue
            v = cell.value
            if v is not None and str(v).strip():
                return r, code_col, name_col, client_col
    return CHINA_EE_DATA_START_ROW, code_col, name_col, client_col


def _directory_row_by_code(code: str, directory: list[dict[str, Any]]) -> dict[str, Any] | None:
    want = norm(code)
    if not want:
        return None
    hits = [
        row
        for row in directory
        if isinstance(row, dict) and norm(row.get("employee_code") or row.get("employeeCode")) == want
    ]
    return hits[0] if hits else None


def _writable_cell(ws: Worksheet, row: int, col: int):
    cell = ws.cell(row, col)
    if not isinstance(cell, MergedCell):
        return cell
    for m in ws.merged_cells.ranges:
        if m.min_row <= row <= m.max_row and m.min_col <= col <= m.max_col:
            return ws.cell(m.min_row, m.min_col)
    return None


def _write_if_empty_or_value(cell, value: Any) -> None:
    """空格写入；已有公式保留（母版公式是真相）。"""
    if cell is None:
        return
    if value is None or value == "":
        if not _is_formula(cell.value):
            cell.value = None
        return
    if _is_formula(cell.value):
        return
    cell.value = value


def _apply_china_ee_from_directory(
    dst_wb,
    employee_names: list[str],
    employee_directory: list[dict[str, Any]] | None,
    *,
    pn_meta: PnMeta | dict[str, Any] | None = None,
) -> list[str]:
    """
    与 TW/UK/UAE 相同：按姓名匹配员工库 employee_code，写入地区 EE 表 EE Code 列。
    母版该格为空时写入；已有公式不覆盖。
    """
    warnings: list[str] = []
    directory = [r for r in (employee_directory or []) if isinstance(r, dict)]
    names = [str(n).strip() for n in employee_names if str(n).strip()]
    if not names:
        warnings.append("未从源表读到员工姓名，无法匹配 EE Code")
        return warnings

    client_code = _pn_customer_id(pn_meta)
    ee_ws = dst_wb[CHINA_EE_SHEET] if CHINA_EE_SHEET in dst_wb.sheetnames else None
    china_ws = dst_wb[CHINA_SHEET] if CHINA_SHEET in dst_wb.sheetnames else None
    data_start, code_col, _name_col, client_col = (
        _scan_china_ee_layout(ee_ws) if ee_ws is not None else (CHINA_EE_DATA_START_ROW, 4, 5, 2)
    )

    for i, name in enumerate(names):
        code, warn = match_ee_code([name], directory)
        if warn:
            warnings.append(f"China EE 第{i + 1}人：{warn}")
        if ee_ws is not None:
            ee_row = data_start + i
            if client_code:
                _write_if_empty_or_value(_writable_cell(ee_ws, ee_row, client_col), client_code)
            code_cell = _writable_cell(ee_ws, ee_row, code_col)
            if code_cell is not None:
                code_cell.value = code
        if china_ws is not None:
            china_row = CHINA_DATA_START_ROW + i
            _write_if_empty_or_value(china_ws.cell(china_row, CHINA_CODE_COL), code)
            hit = _directory_row_by_code(str(code or ""), directory)
            lib_name = ""
            if hit:
                lib_name = (
                    str(hit.get("employee_name") or hit.get("employeeName") or "").strip()
                    or str(hit.get("employee_name_en") or hit.get("employeeNameEn") or "").strip()
                )
            if lib_name:
                _write_if_empty_or_value(china_ws.cell(china_row, CHINA_NAME_COL), lib_name)
        if code:
            print(f"[china-ee-code] {name} → {code} (China EE row {data_start + i})")
    if ee_ws is None:
        warnings.append("母版没有 China EE 表，EE Code 未写入")
    elif not directory:
        warnings.append("未提供客户员工目录，无法匹配 EE Code")
    return warnings


def convert(
    source_path: Path,
    output_path: Path,
    template_path: Path,
    *,
    pn_meta: PnMeta | dict[str, Any] | None = None,
    registry_dir: Path | None = None,
    employee_directory: list[dict[str, Any]] | None = None,
    convert_mapping: dict[str, Any] | None = None,
) -> dict[str, Any]:
    global _ACTIVE_MAPPING
    _ACTIVE_MAPPING = resolve_convert_mapping(ENGINE_ID, convert_mapping)
    try:
        return _convert_impl(
            source_path,
            output_path,
            template_path,
            pn_meta=pn_meta,
            registry_dir=registry_dir,
            employee_directory=employee_directory,
        )
    finally:
        _ACTIVE_MAPPING = None


def _convert_impl(
    source_path: Path,
    output_path: Path,
    template_path: Path,
    *,
    pn_meta: PnMeta | dict[str, Any] | None = None,
    registry_dir: Path | None = None,
    employee_directory: list[dict[str, Any]] | None = None,
) -> dict[str, Any]:
    if not template_path.is_file():
        raise FileNotFoundError(f"母版不存在: {template_path}")
    if not source_path.is_file():
        raise FileNotFoundError(f"原始账单不存在: {source_path}")

    mapping = _active_mapping()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(template_path, output_path)

    with tempfile.TemporaryDirectory(prefix="cn_hrone_pn_") as tmp:
        unlocked = unlock_xlsx(
            source_path,
            Path(tmp),
            passwords=collect_unlock_passwords(source_path, mapping=mapping),
        )
        src_wb = load_workbook(unlocked, data_only=True)
        dst_wb = load_workbook(output_path, rich_text=True)
        try:
            copies = apply_l_sheet_copies(src_wb, dst_wb, mapping)
            applied_pn: PnMeta | None = None
            if pn_meta is not None:
                applied_pn = apply_pn_meta(
                    dst_wb,
                    pn_meta,
                    registry_dir=registry_dir or output_path.parent,
                    reserve_invoice_number=True,
                )
            fx_info = _apply_vendor_fx(dst_wb, src_wb, unlocked, mapping)
            ee_warnings = _apply_china_ee_from_directory(
                dst_wb,
                names_from_copies(copies),
                employee_directory,
                pn_meta=pn_meta,
            )
            if PN_SHEET in dst_wb.sheetnames:
                apply_luckysheet_compat(dst_wb, pn_sheet=PN_SHEET)
            dst_wb.save(output_path)
        finally:
            src_wb.close()
            dst_wb.close()

    postprocess_converted_xlsx(output_path)
    names: list[str] = names_from_copies(copies)
    warnings = list(fx_info.get("warnings") or [])
    warnings.extend(ee_warnings)
    return {
        "engine_id": ENGINE_ID,
        "employee_count": copies[0]["employeeCount"] if copies else 0,
        "employee_names": names,
        "l_sheet_copies": copies,
        "output": str(output_path),
        "pn_meta": applied_pn.to_dict() if applied_pn else None,
        "warnings": warnings,
        "fx_rate": fx_info.get("fx_rate"),
        "fx_source": fx_info.get("fx_source") or "none",
        "fx_row": fx_info.get("fx_row"),
        "pn_fx_write": fx_info.get("pn_fx_write"),
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="HROne Payment Notice → China-L / China-L (2)")
    parser.add_argument("source", type=Path, help="供应商 Excel")
    parser.add_argument("-o", "--output", type=Path, help="输出 PN 路径")
    parser.add_argument("-t", "--template", type=Path, default=DEFAULT_TEMPLATE, help="PN 母版")
    args = parser.parse_args(argv)
    source = args.source.resolve()
    output = args.output.resolve() if args.output else source.parent / f"PN_auto_{source.stem}.xlsx"
    try:
        result = convert(source, output, args.template.resolve())
    except Exception as exc:
        print(f"转换失败: {exc}", file=sys.stderr)
        return 1
    print("转换完成")
    print(f"  输出: {result['output']}")
    print(f"  员工: {result['employee_count']} 人 → {result['employee_names']}")
    for item in result.get("l_sheet_copies") or []:
        print(f"  {item.get('sourceSheet')} → {item.get('targetSheet')}")
    fx_row = result.get("fx_row")
    if result.get("fx_rate") is not None and fx_row:
        print(f"  汇率 PN!B{fx_row}: {result['fx_rate']} ({result.get('fx_source')})")
    for w in result.get("warnings") or []:
        print(f"  提示: {w}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
