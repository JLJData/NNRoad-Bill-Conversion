# -*- coding: utf-8 -*-
"""
HROne HK Payment Notice → China-L / China-L (2)
引擎 china_hrone_payment_notice

只按 mapping.lSheetCopies 填两张 L（默认 S-Payslip → China-L，
S-Payroll Report → China-L (2)）。不改 PN / China / China EE 公式；
汇率一律由 Office 注入 NNRoad 当月1号×0.97（mapping.nnroadExchangeRate），不读供应商账单。
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
from xlsx_convert_utils import norm
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

_ACTIVE_MAPPING: dict[str, Any] | None = None


def _active_mapping() -> dict[str, Any]:
    return (
        _ACTIVE_MAPPING
        if isinstance(_ACTIVE_MAPPING, dict)
        else resolve_convert_mapping(ENGINE_ID, None)
    )


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


def _as_float(value: Any) -> float | None:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        n = float(value)
        return n if n > 0 else None
    text = str(value).strip().replace(",", "")
    if not text:
        return None
    try:
        n = float(text)
        return n if n > 0 else None
    except ValueError:
        return None


def _read_nnroad_fx(mapping: dict[str, Any]) -> tuple[float | None, str | None, str | None]:
    """Office 注入的 NNRoad 当月1号×0.97。返回 (rate, source, error_detail)。

    本引擎不读供应商账单汇率；缺注入或未命中时 error_detail 非空。
    """
    block = mapping.get("nnroadExchangeRate") if isinstance(mapping, dict) else None
    if not isinstance(block, dict) or not block:
        return None, None, None
    status = str(block.get("status") or "").strip().upper()
    rate = _as_float(block.get("rate"))
    if rate is None:
        rate = _as_float(block.get("monthFirst097"))
    month = str(block.get("requestMonth") or "").strip()
    if status == "FOUND" and rate is not None:
        src = str(block.get("source") or "nnroad.exchangeRate.monthFirst097").strip()
        if month:
            src = f"{src}:{month}"
        return rate, src, None
    detail = str(block.get("message") or status or "not_found").strip()
    if month:
        detail = f"{detail}（{month}）"
    return None, None, detail or "not_found"


def _write_pn_fx(
    dst_wb,
    rate: float,
    mapping: dict[str, Any],
    fx_source: str,
    *,
    write_source: str,
) -> dict[str, Any]:
    out: dict[str, Any] = {
        "fx_rate": float(rate),
        "fx_source": fx_source,
        "fx_row": None,
        "pn_fx_write": None,
        "warnings": [],
    }
    if PN_SHEET not in dst_wb.sheetnames:
        out["warnings"].append("母版没有 PN 表，汇率已读到但未写入")
        return out
    fx_row = _find_pn_fx_row(dst_wb[PN_SHEET])
    if fx_row is None:
        out["warnings"].append("母版 PN 未找到 FX rate 行，汇率已读到但未写入")
        return out
    cell = dst_wb[PN_SHEET].cell(fx_row, 2)
    cell.value = float(rate)
    cell.number_format = "0.00"
    out["fx_row"] = fx_row
    out["pn_fx_write"] = make_pn_fx_provenance(
        PN_SHEET,
        fx_row,
        2,
        mapping,
        float(rate),
        write_source=write_source,
        fx_source=fx_source,
    )
    return out


def _apply_fx(dst_wb, mapping: dict[str, Any]) -> dict[str, Any]:
    policy = fx_policy(mapping)
    mode = str(policy.get("mode") or "none").strip().lower()
    empty: dict[str, Any] = {
        "fx_rate": None,
        "fx_source": "none",
        "fx_row": None,
        "pn_fx_write": None,
        "warnings": [],
    }
    if mode == "none":
        return empty

    nnroad_rate, nnroad_src, nnroad_err = _read_nnroad_fx(mapping)
    if nnroad_rate is not None:
        return _write_pn_fx(
            dst_wb,
            nnroad_rate,
            mapping,
            nnroad_src or "nnroad.exchangeRate.monthFirst097",
            write_source="api",
        )
    detail = nnroad_err or "未注入 nnroadExchangeRate"
    empty["warnings"].append(
        f"汇率应按当月1号×0.97 取自 NNRoad，但未取到（{detail}），PN FX 格未改（不用供应商账单汇率）"
    )
    return empty


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
        try:
            src_wb = load_workbook(unlocked, data_only=True)
        except Exception as exc:
            raise ValueError(f"源表解密后仍无法打开（{unlocked.name}）: {exc}") from exc
        try:
            dst_wb = load_workbook(output_path, rich_text=True)
        except Exception as exc:
            raise ValueError(f"母版不是有效 xlsx（{output_path.name}）: {exc}") from exc
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
            fx_info = _apply_fx(dst_wb, mapping)
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
