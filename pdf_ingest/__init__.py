# -*- coding: utf-8 -*-
"""供应商 PDF → 源 Excel（各地区 -L / 计算结果）抽取入口。"""
from __future__ import annotations

from .registry import get_pdf_profile, list_pdf_profiles

__all__ = ["get_pdf_profile", "list_pdf_profiles"]
