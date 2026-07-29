# -*- coding: utf-8 -*-
"""统一按 engine_id 调度转换（供 CLI / HTTP 调用）。"""
from __future__ import annotations

import importlib
import json
from datetime import date
from pathlib import Path
from typing import Any

from engines import get_engine
from pn_meta import PnMeta
from region_templates import get_region_template

BASE_DIR = Path(__file__).resolve().parent


def run_convert(
    *,
    engine_id: str,
    source_path: Path,
    output_path: Path,
    region: str | None = None,
    template_path: Path | None = None,
    pn_meta: PnMeta | dict[str, Any] | None = None,
    registry_dir: Path | None = None,
    employee_directory: list[dict[str, Any]] | None = None,
    convert_mapping: dict[str, Any] | None = None,
) -> dict[str, Any]:
    engine = get_engine(engine_id)
    module = importlib.import_module(engine.module)
    if not hasattr(module, "convert"):
        raise RuntimeError(f"引擎模块缺少 convert(): {engine.module}")

    if template_path is None:
        if not region:
            raise ValueError("未指定 template_path 时必须提供 region")
        template_path = get_region_template(region)

    source_path = Path(source_path).resolve()
    output_path = Path(output_path).resolve()
    template_path = Path(template_path).resolve()
    registry = Path(registry_dir).resolve() if registry_dir else (BASE_DIR)

    meta_obj: PnMeta | None = None
    if isinstance(pn_meta, PnMeta):
        meta_obj = pn_meta
    elif isinstance(pn_meta, dict) and pn_meta:
        meta_obj = PnMeta.from_dict(pn_meta)

    convert_kwargs: dict[str, Any] = {
        "pn_meta": meta_obj,
        "registry_dir": registry,
    }
    # 引擎签名支持时注入；china_payroll_calc / tw_payroll_calc 等各自消费，互不混用字段名
    import inspect

    sig = inspect.signature(module.convert)
    if "employee_directory" in sig.parameters:
        convert_kwargs["employee_directory"] = employee_directory
    # 即使 mapping 为空也传入，让引擎显式走「无映射」分支并打诊断
    if "convert_mapping" in sig.parameters:
        convert_kwargs["convert_mapping"] = convert_mapping

    result = module.convert(
        source_path,
        output_path,
        template_path,
        **convert_kwargs,
    )
    result = dict(result or {})
    result["engine_id"] = engine_id
    result["engine_label"] = engine.label
    result["region"] = region
    result["template"] = str(template_path)
    result["output"] = str(output_path)
    return result


def parse_pn_meta_payload(raw: str | None) -> dict[str, Any] | None:
    if not raw or not str(raw).strip():
        return None
    data = json.loads(raw)
    if not isinstance(data, dict):
        raise ValueError("pn_meta 必须是 JSON 对象")
    # 缺省账单日用今天，便于自动生成发票号
    if not data.get("invoice_date"):
        data["invoice_date"] = date.today().isoformat()
    return data


def parse_convert_mapping_payload(raw: str | None) -> dict[str, Any] | None:
    if not raw or not str(raw).strip():
        return None
    data = json.loads(raw)
    if not isinstance(data, dict):
        raise ValueError("convert_mapping 必须是 JSON 对象")
    return data


def parse_employee_directory_payload(raw: str | None) -> list[dict[str, Any]] | None:
    if not raw or not str(raw).strip():
        return None
    data = json.loads(raw)
    if not isinstance(data, list):
        raise ValueError("employee_directory 必须是 JSON 数组")
    out: list[dict[str, Any]] = []
    for item in data:
        if isinstance(item, dict):
            out.append(item)
    return out
