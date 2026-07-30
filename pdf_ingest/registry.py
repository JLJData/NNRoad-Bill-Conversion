# -*- coding: utf-8 -*-
"""PDF 抽取 profile 注册：供应商版式 → 解析模块。"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Callable


@dataclass(frozen=True)
class PdfProfile:
    profile_id: str
    label: str
    region: str
    module: str
    description: str
    # 挂在哪些转换引擎下（Office 配置级联）
    engine_ids: tuple[str, ...]
    # 用于自动识别的关键字（全文小写匹配，命中任一即可）
    detect_keywords: tuple[str, ...]


PDF_PROFILES: dict[str, PdfProfile] = {
    "eor_uk": PdfProfile(
        profile_id="eor_uk",
        label="EOR Services",
        region="UK",
        module="pdf_ingest.profiles.eor_uk",
        description="EOR：PDF 按本版式解析；Excel 源则直接走 uk_payroll_calc",
        engine_ids=("uk_payroll_calc",),
        detect_keywords=(
            "eor services limited",
            "eorservices.co.uk",
            "zaka@eorservices.co.uk",
        ),
    ),
    "topsource_uk": PdfProfile(
        profile_id="topsource_uk",
        label="TopSource",
        region="UK",
        module="pdf_ingest.profiles.topsource_uk",
        description="TopSource：PDF 按本版式解析（多人一票多 PDF）；Excel 源则直接走 uk_payroll_calc",
        engine_ids=("uk_payroll_calc",),
        detect_keywords=(
            "topsource worldwide",
            "topsourceworldwide.com",
            "topsource registered",
        ),
    ),
}


def get_pdf_profile(profile_id: str) -> PdfProfile:
    p = PDF_PROFILES.get((profile_id or "").strip())
    if p is None:
        known = ", ".join(sorted(PDF_PROFILES))
        raise KeyError(f"未知 PDF profile「{profile_id}」，已知: {known}")
    return p


def list_pdf_profiles() -> list[PdfProfile]:
    return [PDF_PROFILES[k] for k in sorted(PDF_PROFILES)]


def detect_pdf_profile(text: str) -> PdfProfile | None:
    low = (text or "").lower()
    if not low.strip():
        return None
    for p in PDF_PROFILES.values():
        if any(k in low for k in p.detect_keywords):
            return p
    return None


def load_convert_fn(profile: PdfProfile) -> Callable:
    import importlib

    mod = importlib.import_module(profile.module)
    fn = getattr(mod, "convert_pdf", None)
    if not callable(fn):
        raise RuntimeError(f"模块 {profile.module} 缺少 convert_pdf()")
    return fn
