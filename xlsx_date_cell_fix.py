# -*- coding: utf-8 -*-
"""
转换结果 xlsx：仅处理「模板/convert 已标日期格式」的单元格，把文本/带 ' 的值规范为 datetime。
不对 General 格式的数字做日期猜测（避免误伤金额/编号）。
"""
from __future__ import annotations

import re
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook

from xlsx_convert_utils import coerce_datetime_for_excel

_DATE_FMT_HINT = re.compile(r"y{1,4}|m/d|d/m|d-mmm|h:mm|年|月", re.I)


def _is_excel_date_format(fmt: str | None) -> bool:
    if not fmt or str(fmt).strip().lower() == "general":
        return False
    head = str(fmt).split(";")[0]
    if re.match(r"^[\s#0,]*\.?0*%?$", head.replace("$", "")) and not _DATE_FMT_HINT.search(head):
        return False
    return bool(_DATE_FMT_HINT.search(head))


def _parse_date_text(text: str) -> datetime | None:
    s = text.strip().lstrip("'").strip()
    if not s:
        return None
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%m/%d/%Y", "%d/%m/%Y", "%Y/%m/%d %H:%M:%S"):
        try:
            dt = datetime.strptime(s, fmt)
            return datetime(dt.year, dt.month, dt.day)
        except ValueError:
            continue
    return None


def normalize_template_date_cells(xlsx_path: Path | str) -> int:
    path = Path(xlsx_path)
    if not path.is_file():
        return 0
    wb = load_workbook(path)
    changed = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                fmt = cell.number_format or ""
                if not _is_excel_date_format(fmt):
                    continue
                val = cell.value
                if isinstance(val, datetime):
                    if val.hour or val.minute or val.second:
                        cell.value = datetime(val.year, val.month, val.day)
                        changed += 1
                    continue
                if isinstance(val, date) and not isinstance(val, datetime):
                    cell.value = datetime(val.year, val.month, val.day)
                    changed += 1
                    continue
                if isinstance(val, str):
                    dt = _parse_date_text(val)
                    if dt is not None:
                        cell.value = dt
                        changed += 1
                        continue
                if isinstance(val, (int, float)) and not isinstance(val, bool):
                    dt = coerce_datetime_for_excel(val)
                    if dt is not None:
                        cell.value = dt
                        changed += 1
    if changed:
        wb.save(path)
    return changed
