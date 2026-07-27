# -*- coding: utf-8 -*-
"""
转换结果写盘后的统一后处理：

1) inlineStr 富文本 → sharedStrings（LuckyExcel 预览可读完整文本）
2) 主题填充落地 RGB + applyFill（避免丢浅蓝底）

金额公式结果由前端 HyperFormula 重算，不再注入 PN 公式缓存 <v>。
"""
from __future__ import annotations

from pathlib import Path

from xlsx_richtext_fix import migrate_inlinestr_richtext_to_shared_strings
from xlsx_theme_fill_fix import materialize_theme_fills
from xlsx_date_cell_fix import normalize_template_date_cells


def postprocess_converted_xlsx(xlsx_path: Path | str) -> dict[str, int]:
    """
    按固定顺序后处理已保存的 xlsx。

    Returns:
        各步骤影响计数（便于日志/回归）。
    """
    path = Path(xlsx_path)
    migrate_inlinestr_richtext_to_shared_strings(path)
    theme_stats = materialize_theme_fills(path) or {}
    date_cells = normalize_template_date_cells(path)
    return {
        "theme_fills": int(theme_stats.get("fills", 0) or 0),
        "theme_xfs": int(theme_stats.get("xfs", 0) or 0),
        "date_cells": int(date_cells or 0),
    }
