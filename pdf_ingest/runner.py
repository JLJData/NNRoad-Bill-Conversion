# -*- coding: utf-8 -*-
"""供应商 PDF/Excel → 源 Excel 运行入口。"""
from __future__ import annotations

import importlib
import inspect
from pathlib import Path
from typing import Any

from pdf_ingest.registry import detect_pdf_profile, get_pdf_profile, load_convert_fn
from pdf_ingest.text_extract import extract_pdf_text
from pn_meta import PnMeta

_EXCEL_SUFFIXES = {".xlsx", ".xlsm", ".xls"}


def _resolve_profile(pdf_path: Path, profile_id: str | None):
    if profile_id:
        return get_pdf_profile(profile_id)
    text = extract_pdf_text(pdf_path)
    profile = detect_pdf_profile(text)
    if profile is None:
        raise ValueError(
            "无法自动识别 PDF 供应商版式，请显式传入 profile_id（如 eor_uk / topsource_uk）"
        )
    return profile


def _resolve_profile_for_sources(paths: list[Path], profile_id: str | None):
    if profile_id:
        return get_pdf_profile(profile_id)
    pdfs = [p for p in paths if p.suffix.lower() == ".pdf"]
    if pdfs:
        return _resolve_profile(pdfs[0], None)
    raise ValueError(
        "Excel 源须显式传入 profile_id（如 topsource_uk / auxilium_uae）；无法从正文自动识别"
    )


def _call_convert(fn, *args, **kwargs):
    sig = inspect.signature(fn)
    filtered = {k: v for k, v in kwargs.items() if k in sig.parameters}
    return fn(*args, **filtered)


def run_pdf_to_source(
    pdf_path: Path,
    output_path: Path,
    *,
    profile_id: str | None = None,
    template_path: Path | None = None,
    pn_meta: PnMeta | dict[str, Any] | None = None,
    registry_dir: Path | None = None,
    fill_fx: bool = True,
) -> dict[str, Any]:
    pdf_path = Path(pdf_path)
    output_path = Path(output_path)
    profile = _resolve_profile(pdf_path, profile_id)
    convert_fn = load_convert_fn(profile)
    return _call_convert(
        convert_fn,
        pdf_path,
        output_path,
        template_path=template_path,
        pn_meta=pn_meta,
        registry_dir=registry_dir or output_path.parent,
        fill_fx=fill_fx,
    )


def run_pdf_to_source_batch(
    pdf_paths: list[Path],
    output_path: Path,
    *,
    profile_id: str | None = None,
    template_path: Path | None = None,
    pn_meta: PnMeta | dict[str, Any] | None = None,
    registry_dir: Path | None = None,
    fill_fx: bool = True,
) -> dict[str, Any]:
    """多份 PDF → 一份源表（如 TopSource 一人一票 → 多张 UK-L）。"""
    paths = [Path(p) for p in pdf_paths]
    if not paths:
        raise ValueError("未提供 PDF")
    output_path = Path(output_path)
    profile = _resolve_profile(paths[0], profile_id)

    mod = importlib.import_module(profile.module)
    batch_fn = getattr(mod, "convert_pdfs", None)
    if callable(batch_fn):
        return _call_convert(
            batch_fn,
            paths,
            output_path,
            template_path=template_path,
            pn_meta=pn_meta,
            registry_dir=registry_dir or output_path.parent,
            fill_fx=fill_fx,
        )

    convert_fn = load_convert_fn(profile)
    result = _call_convert(
        convert_fn,
        paths[0],
        output_path,
        template_path=template_path,
        pn_meta=pn_meta,
        registry_dir=registry_dir or output_path.parent,
        fill_fx=fill_fx,
    )
    if len(paths) > 1:
        warnings = list(result.get("warnings") or [])
        warnings.append(
            f"当前 pdf_profile「{profile.profile_id}」不支持批量，仅处理了第 1 份，其余 {len(paths) - 1} 份已忽略"
        )
        result["warnings"] = warnings
    return result


def run_vendor_to_source_batch(
    source_paths: list[Path],
    output_path: Path,
    *,
    profile_id: str | None = None,
    template_path: Path | None = None,
    pn_meta: PnMeta | dict[str, Any] | None = None,
    registry_dir: Path | None = None,
    fill_fx: bool = True,
    convert_mapping: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """
    供应商 PDF 和/或 Excel → 一份源表。
    按扩展名分流；优先调用 profile 的 convert_sources / convert_excels / convert_pdfs。
    """
    paths = [Path(p) for p in source_paths]
    if not paths:
        raise ValueError("未提供源文件")
    output_path = Path(output_path)
    profile = _resolve_profile_for_sources(paths, profile_id)
    mod = importlib.import_module(profile.module)

    sources_fn = getattr(mod, "convert_sources", None)
    if callable(sources_fn):
        return _call_convert(
            sources_fn,
            paths,
            output_path,
            template_path=template_path,
            pn_meta=pn_meta,
            registry_dir=registry_dir or output_path.parent,
            fill_fx=fill_fx,
            convert_mapping=convert_mapping,
        )

    pdfs = [p for p in paths if p.suffix.lower() == ".pdf"]
    excels = [p for p in paths if p.suffix.lower() in _EXCEL_SUFFIXES]
    if pdfs and excels:
        raise ValueError("同一批次请不要混传 PDF 与 Excel")
    if excels:
        excel_fn = getattr(mod, "convert_excels", None)
        if not callable(excel_fn):
            raise ValueError(
                f"pdf_profile「{profile.profile_id}」暂不支持 Excel 源，请上传 PDF 或已成型的 UK-L Excel"
            )
        return _call_convert(
            excel_fn,
            excels,
            output_path,
            template_path=template_path,
            pn_meta=pn_meta,
            registry_dir=registry_dir or output_path.parent,
            fill_fx=fill_fx,
            convert_mapping=convert_mapping,
        )
    return run_pdf_to_source_batch(
        pdfs or paths,
        output_path,
        profile_id=profile.profile_id,
        template_path=template_path,
        pn_meta=pn_meta,
        registry_dir=registry_dir,
        fill_fx=fill_fx,
    )
