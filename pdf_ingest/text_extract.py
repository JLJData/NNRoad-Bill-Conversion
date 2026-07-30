# -*- coding: utf-8 -*-
"""PDF 文本抽取公共工具。"""
from __future__ import annotations

from pathlib import Path


def extract_pdf_text(pdf_path: Path) -> str:
    try:
        from pypdf import PdfReader
    except ImportError as exc:
        raise RuntimeError("缺少依赖 pypdf，请执行: pip install pypdf") from exc

    reader = PdfReader(str(pdf_path))
    if reader.is_encrypted:
        try:
            reader.decrypt("")
        except Exception as exc:
            raise ValueError(f"PDF 已加密，无法读取: {pdf_path}") from exc
    parts: list[str] = []
    for page in reader.pages:
        parts.append(page.extract_text() or "")
    text = "\n".join(parts).strip()
    if not text:
        raise ValueError(
            f"未能从 PDF 抽出文字（可能是扫描件，需 OCR）: {pdf_path}"
        )
    return text
