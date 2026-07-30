# -*- coding: utf-8 -*-
"""
公式兼容修复（与地区无关）：写入 xlsx 前规范化，供 HyperFormula / LuckySheet / Excel 共用。

- PN 母版误写 `="- "&+"Expense...` → 合法拼接
- A1 用 =MID(CELL("filename",A1),...) 取表名 → 静态表名（避免自引用）
- `=+'Sheet'!A1` 一元加号、EOMONTH(…)/TODAY() → HyperFormula 可算写法
"""
from __future__ import annotations

import calendar
import re
from datetime import date, timedelta

from openpyxl.worksheet.worksheet import Worksheet

_AMP_PLUS_RE = re.compile(r"&\s*\+")
_DASH_AMP_QUOTE_RE = re.compile(r'="-\s*"&"([^"]*)"')
_SHEET_TITLE_SELF_REF_RE = re.compile(
    r'CELL\s*\(\s*"filename"\s*,\s*\$?A\$?1\s*\)',
    re.IGNORECASE,
)
_UNARY_PLUS_RE = re.compile(r"^=\s*\+")
_EOMONTH_TODAY_RE = re.compile(
    r"EOMONTH\s*\(\s*TODAY\s*\(\s*\)\s*,\s*(-?\d+)\s*\)(?:\s*\+\s*(\d+))?",
    re.IGNORECASE,
)
_EOMONTH_HEAD_RE = re.compile(r"EOMONTH\s*\(", re.IGNORECASE)
_TODAY_RE = re.compile(r"TODAY\s*\(\s*\)", re.IGNORECASE)


def fix_pn_illegal_concat_formulas(ws: Worksheet) -> int:
    """母版误写 `="- "&+"Expense...`：Excel 能算，LuckySheet/部分引擎会 #VALUE!。"""
    n = 0
    for row in ws.iter_rows():
        for cell in row:
            v = cell.value
            if not (isinstance(v, str) and v.startswith("=") and _AMP_PLUS_RE.search(v)):
                continue
            nv = _AMP_PLUS_RE.sub("&", v)
            nv = _DASH_AMP_QUOTE_RE.sub(r'="- \1"', nv)
            if nv != v:
                cell.value = nv
                n += 1
    return n


def flatten_sheet_title_self_refs(wb) -> int:
    """
    母版 A1 用 =MID(CELL(\"filename\",A1),...) 取表名。
    Excel 允许这种自引用；LuckySheet 重算时会弹
    「公式不可引用其本身的单元格，会导致计算结果不准确」。
    转换时写成静态表名即可（缓存值本就是表名）。
    """
    n = 0
    for name in wb.sheetnames:
        cell = wb[name].cell(1, 1)
        v = cell.value
        if isinstance(v, str) and v.startswith("=") and _SHEET_TITLE_SELF_REF_RE.search(v):
            cell.value = name
            n += 1
    return n


def _eomonth(d: date, months: int) -> date:
    y = d.year + (d.month - 1 + months) // 12
    m = (d.month - 1 + months) % 12 + 1
    last = calendar.monthrange(y, m)[1]
    return date(y, m, last)


def _replace_eomonth_today(match: re.Match) -> str:
    months = int(match.group(1))
    plus = int(match.group(2) or 0)
    end = _eomonth(date.today(), months)
    if plus:
        end = end + timedelta(days=plus)
    return f"DATE({end.year},{end.month},{end.day})"


def _rewrite_eomonth_to_date(formula: str) -> str:
    """EOMONTH(date, n) → DATE(YEAR(date), MONTH(date)+n+1, 0)；HF 无 EOMONTH。"""
    if not _EOMONTH_HEAD_RE.search(formula):
        return formula
    f = formula
    for _ in range(20):
        m = _EOMONTH_HEAD_RE.search(f)
        if not m:
            break
        start = m.start()
        open_idx = m.end() - 1
        depth = 0
        end = -1
        in_str = False
        for i in range(open_idx, len(f)):
            ch = f[i]
            if ch == '"':
                in_str = not in_str
                continue
            if in_str:
                continue
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
                if depth == 0:
                    end = i
                    break
        if end < 0:
            break
        inner = f[open_idx + 1 : end]
        comma = -1
        d2 = 0
        s2 = False
        for i, ch in enumerate(inner):
            if ch == '"':
                s2 = not s2
                continue
            if s2:
                continue
            if ch == "(":
                d2 += 1
            elif ch == ")":
                d2 -= 1
            elif ch == "," and d2 == 0:
                comma = i
                break
        if comma < 0:
            break
        date_expr = inner[:comma].strip()
        months_expr = inner[comma + 1 :].strip()
        if not date_expr or not months_expr:
            break
        replacement = f"DATE(YEAR({date_expr}),MONTH({date_expr})+({months_expr})+1,0)"
        f = f[:start] + replacement + f[end + 1 :]
    return f


def normalize_formula_for_lucky(formula: str) -> str:
    """单条公式改写，供转换写盘前 / 扫描修复复用。"""
    if not (isinstance(formula, str) and formula.startswith("=")):
        return formula
    f = formula
    if _UNARY_PLUS_RE.match(f):
        f = _UNARY_PLUS_RE.sub("=", f, count=1)
    if _AMP_PLUS_RE.search(f):
        f = _AMP_PLUS_RE.sub("&", f)
        f = _DASH_AMP_QUOTE_RE.sub(r'="- \1"', f)
    if re.search(r"EOMONTH", f, re.I) and re.search(r"TODAY\s*\(", f, re.I):
        f = _EOMONTH_TODAY_RE.sub(_replace_eomonth_today, f)
    if _EOMONTH_HEAD_RE.search(f):
        f = _rewrite_eomonth_to_date(f)
    if _TODAY_RE.search(f) and not _EOMONTH_HEAD_RE.search(f):
        today = date.today()
        f = _TODAY_RE.sub(f"DATE({today.year},{today.month},{today.day})", f)
    return f


def fix_workbook_lucky_formulas(wb) -> int:
    """整本扫描：一元加号 / &+ / EOMONTH(TODAY()) / TODAY()。"""
    n = 0
    for name in wb.sheetnames:
        ws = wb[name]
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if not (isinstance(v, str) and v.startswith("=")):
                    continue
                nv = normalize_formula_for_lucky(v)
                if nv != v:
                    cell.value = nv
                    n += 1
    return n


def apply_luckysheet_compat(wb, *, pn_sheet: str | None = "PN") -> dict[str, int]:
    """转换写盘前统一兼容修复。"""
    stats = {
        "amp_plus": 0,
        "sheet_title": flatten_sheet_title_self_refs(wb),
        "formula_normalize": 0,
    }
    if pn_sheet and pn_sheet in wb.sheetnames:
        stats["amp_plus"] = fix_pn_illegal_concat_formulas(wb[pn_sheet])
    # 再扫一遍全表（含 China!BH 的 =+、China!B2 EOMONTH、PN 残留 &+）
    stats["formula_normalize"] = fix_workbook_lucky_formulas(wb)
    return stats
