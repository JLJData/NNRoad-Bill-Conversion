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
    "auxilium_uae": PdfProfile(
        profile_id="auxilium_uae",
        label="Auxilium",
        region="UAE",
        module="pdf_ingest.profiles.auxilium_uae",
        description="Auxilium：Payroll Draft Excel → UAE-L；再经 uae_payroll_calc → PN",
        engine_ids=("uae_payroll_calc",),
        detect_keywords=(
            "payroll draft",
            "ax id",
            "nnroad (uae)",
        ),
    ),
    "connect_uae": PdfProfile(
        profile_id="connect_uae",
        label="Connect",
        region="UAE",
        module="pdf_ingest.profiles.connect_uae",
        description="Connect Resources：Tax Invoice PDF → UAE-L；再经 uae_payroll_calc → PN",
        engine_ids=("uae_payroll_calc",),
        detect_keywords=(
            "connect resources",
            "connect resources llc",
        ),
    ),
    "panda_work_pk": PdfProfile(
        profile_id="panda_work_pk",
        label="Panda Work Global",
        region="Pakistan",
        module="pdf_ingest.profiles.panda_work_pk",
        description="Panda Work Global：季度发票 PDF（一人一票）→ Pakistan-L；再经 pakistan_payroll_calc → PN",
        engine_ids=("pakistan_payroll_calc",),
        detect_keywords=(
            "panda work global",
            "pandaworkglobal.com",
            "panda work global (private) limited",
        ),
    ),
    "safeguard_italy": PdfProfile(
        profile_id="safeguard_italy",
        label="SafeGuard",
        region="Italy",
        module="pdf_ingest.profiles.safeguard_italy",
        description="SafeGuard (SGWI)：Payroll Excel → Italy-L；再经 italy_payroll_calc → PN",
        engine_ids=("italy_payroll_calc",),
        detect_keywords=(
            "sgwi",
            "safeguard",
            "sgwi vat code",
            "sgwi chargable",
        ),
    ),
    "biz_solutions_india": PdfProfile(
        profile_id="biz_solutions_india",
        label="Biz Solutions",
        region="India",
        module="pdf_ingest.profiles.biz_solutions_india",
        description="Biz Solutions：Tax Invoice PDF → India-L；薪资拆分+PT/IIT 见 mapping.indiaSalarySplit",
        engine_ids=("india_payroll_calc",),
        detect_keywords=(
            "biz solutions",
            "biz/26-",
            "biz/25-",
            "outsource payroll services",
            "payroll management services",
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
