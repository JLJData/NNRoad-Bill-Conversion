# -*- coding: utf-8 -*-
"""
本机 Microsoft Excel COM 计算快照（模板发布三引擎验证 · 业务基准）

用法:
  python excel_com_snapshot.py <xlsx路径> [--sheet PN] [--max 300]

输出 JSON 到 stdout（UTF-8）:
  {
    "ok": true,
    "engine": "excel",
    "sheetName": "PN",
    "cells": [{"sheet","ref","row","col","formula","numberValue","textValue"}, ...],
    "truncated": false,
    "message": ""
  }

依赖: pip install pywin32
仅 Windows + 已安装桌面 Excel。
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any


def _col_to_a1(col_1based: int) -> str:
    n = col_1based
    s = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        s = chr(65 + rem) + s
    return s or "A"


def _to_a1(row_1based: int, col_1based: int) -> str:
    return f"{_col_to_a1(col_1based)}{row_1based}"


def _coerce_number(v: Any) -> float | None:
    if v is None:
        return None
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        f = float(v)
        return f if f == f and f not in (float("inf"), float("-inf")) else None
    s = str(v).strip().replace(",", "").replace("，", "").replace("%", "")
    if not s or s.startswith("="):
        return None
    try:
        f = float(s)
        return f if f == f else None
    except ValueError:
        return None


def snapshot_workbook(path: Path, sheet_filter: str | None, max_cells: int) -> dict[str, Any]:
    try:
        import win32com.client  # type: ignore
        import pythoncom  # type: ignore
    except ImportError as exc:
        return {
            "ok": False,
            "engine": "excel",
            "sheetName": sheet_filter or "",
            "cells": [],
            "truncated": False,
            "message": f"缺少 pywin32，请先 pip install pywin32（{exc}）",
        }

    if not path.is_file():
        return {
            "ok": False,
            "engine": "excel",
            "sheetName": sheet_filter or "",
            "cells": [],
            "truncated": False,
            "message": f"文件不存在: {path}",
        }

    pythoncom.CoInitialize()
    excel = None
    wb = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        excel.EnableEvents = False
        # 0 = xlCalculationAutomatic
        try:
            excel.Calculation = -4105
        except Exception:
            pass

        abs_path = str(path.resolve())
        wb = excel.Workbooks.Open(
            abs_path,
            UpdateLinks=0,
            ReadOnly=True,
            IgnoreReadOnlyRecommended=True,
        )
        try:
            excel.CalculateFullRebuild()
        except Exception:
            try:
                excel.CalculateFull()
            except Exception:
                try:
                    wb.Application.Calculate()
                except Exception:
                    pass

        want = (sheet_filter or "").strip().lower()
        cells_out: list[dict[str, Any]] = []
        used_sheet_name = ""
        truncated = False

        for i in range(1, int(wb.Worksheets.Count) + 1):
            if len(cells_out) >= max_cells:
                truncated = True
                break
            ws = wb.Worksheets(i)
            name = str(ws.Name or "")
            if want and name.lower() != want:
                continue
            if not used_sheet_name:
                used_sheet_name = name
            used = ws.UsedRange
            if used is None:
                continue
            rows = int(used.Rows.Count)
            cols = int(used.Columns.Count)
            base_row = int(used.Row)
            base_col = int(used.Column)
            for r_off in range(rows):
                if len(cells_out) >= max_cells:
                    truncated = True
                    break
                for c_off in range(cols):
                    if len(cells_out) >= max_cells:
                        truncated = True
                        break
                    cell = used.Cells(r_off + 1, c_off + 1)
                    try:
                        has_formula = bool(cell.HasFormula)
                    except Exception:
                        has_formula = False
                    if not has_formula:
                        continue
                    try:
                        formula = str(cell.Formula or "").strip()
                    except Exception:
                        formula = ""
                    if not formula:
                        continue
                    if not formula.startswith("="):
                        formula = "=" + formula
                    row1 = base_row + r_off
                    col1 = base_col + c_off
                    try:
                        raw_val = cell.Value
                    except Exception:
                        raw_val = None
                    try:
                        text_val = str(cell.Text or "").strip()
                    except Exception:
                        text_val = ""
                    num = _coerce_number(raw_val)
                    if num is None and text_val:
                        num = _coerce_number(text_val)
                    if num is None and (not text_val or text_val.startswith("=")):
                        continue
                    cells_out.append(
                        {
                            "sheet": name,
                            "ref": _to_a1(row1, col1),
                            "row": row1 - 1,
                            "col": col1 - 1,
                            "formula": formula,
                            "numberValue": num,
                            "textValue": text_val or ("" if num is None else str(num)),
                        }
                    )
            if want:
                break

        if want and not used_sheet_name:
            return {
                "ok": False,
                "engine": "excel",
                "sheetName": sheet_filter or "",
                "cells": [],
                "truncated": False,
                "message": f"未找到工作表: {sheet_filter}",
            }

        msg = ""
        if truncated:
            msg = f"公式格较多，仅抽取前 {max_cells} 个"
        return {
            "ok": True,
            "engine": "excel",
            "sheetName": used_sheet_name or (sheet_filter or ""),
            "cells": cells_out,
            "truncated": truncated,
            "message": msg,
            "cellCount": len(cells_out),
        }
    except Exception as exc:
        return {
            "ok": False,
            "engine": "excel",
            "sheetName": sheet_filter or "",
            "cells": [],
            "truncated": False,
            "message": f"Excel COM 失败: {exc}",
        }
    finally:
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if excel is not None:
                excel.Quit()
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Excel COM formula snapshot")
    parser.add_argument("path", help="xlsx/xlsm path")
    parser.add_argument("--sheet", default="PN", help="sheet name filter; empty=all")
    parser.add_argument("--max", type=int, default=300, help="max formula cells")
    args = parser.parse_args(argv)
    sheet = args.sheet
    if sheet is not None and sheet.strip() == "":
        sheet = None
    result = snapshot_workbook(Path(args.path), sheet, max(1, int(args.max)))
    sys.stdout.write(json.dumps(result, ensure_ascii=False))
    sys.stdout.flush()
    return 0 if result.get("ok") else 2


if __name__ == "__main__":
    raise SystemExit(main())
