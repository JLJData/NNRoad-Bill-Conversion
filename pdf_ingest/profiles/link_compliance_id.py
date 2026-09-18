# -*- coding: utf-8 -*-
"""
Link Compliance（Indonesia）Tax Invoice PDF

主转换请用工资明细 Excel（indonesia_payroll_calc）。
本 PDF 仅作旁路：抽取 Cash Advance → 写入 PN（见 vendor_plugins.link_compliance_cash_advance）。
"""
from __future__ import annotations

from pathlib import Path
from typing import Any

from bill_convert.vendor_plugins.link_compliance_cash_advance import (
    looks_like_link_compliance_invoice,
    parse_cash_advance_pdf,
)


def convert_pdf(
    pdf_path: Path,
    output_path: Path,
    *,
    template_path: Path | None = None,
    pn_meta: Any = None,
    registry_dir: Path | None = None,
    fill_fx: bool = True,
    convert_mapping: dict | None = None,
) -> dict[str, Any]:
    raise ValueError(
        "Link Compliance 主账单请上传工资明细 Excel；"
        "Tax Invoice PDF 仅用于抽取 Cash Advance（作旁路附件 / artifactBatch），"
        "勿单独走 pdf-to-source"
    )


def convert_pdfs(
    pdf_paths: list[Path],
    output_path: Path,
    **kwargs: Any,
) -> dict[str, Any]:
    return convert_pdf(pdf_paths[0] if pdf_paths else Path("."), output_path, **kwargs)
