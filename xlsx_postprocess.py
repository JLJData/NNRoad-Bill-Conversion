# -*- coding: utf-8 -*-
"""
转换结果写盘后的统一后处理：

1) 日期值/格式规范化（会通过 openpyxl 再次保存）
2) inlineStr 富文本 → sharedStrings（LuckyExcel 预览可读完整文本）
3) 主题填充落地 RGB + applyFill（必须在最后一次 openpyxl 保存之后）

此步骤不计算公式。Office 后端使用 LibreOffice 计算并保存 converted 公式缓存；
前端 HyperFormula 仅用于网页预览，不能替代正式落盘结果。
"""
from __future__ import annotations

from pathlib import Path

from xlsx_keep_images import require_pillow
from xlsx_richtext_fix import migrate_inlinestr_richtext_to_shared_strings
from xlsx_theme_fill_fix import materialize_theme_fills
from xlsx_date_cell_fix import normalize_template_date_cells
from xlsx_pn_font_fix import normalize_pn_billto_fonts


def postprocess_converted_xlsx(xlsx_path: Path | str) -> dict[str, int]:
    """
    按固定顺序后处理已保存的 xlsx。

    Returns:
        各步骤影响计数（便于日志/回归）。
    """
    path = Path(xlsx_path)
    require_pillow()
    date_cells = normalize_template_date_cells(path)
    migrate_inlinestr_richtext_to_shared_strings(path)
    theme_stats = materialize_theme_fills(path) or {}
    pn_fonts = normalize_pn_billto_fonts(path)
    return {
        "theme_fills": int(theme_stats.get("fills", 0) or 0),
        "theme_xfs": int(theme_stats.get("xfs", 0) or 0),
        "date_cells": int(date_cells or 0),
        "pn_fonts": int(pn_fonts or 0),
    }
