# -*- coding: utf-8 -*-
"""转换引擎注册表：按源账单格式划分，与输出地区（region）独立。"""
from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class ConvertEngine:
    engine_id: str
    label: str
    module: str
    description: str


ENGINES: dict[str, ConvertEngine] = {
    "china_hrone": ConvertEngine(
        engine_id="china_hrone",
        label="HROne China",
        module="profiles.china_hrone.convert",
        description="源账单 sheet「计算结果」→ China PN",
    ),
    "hk_vertical_l": ConvertEngine(
        engine_id="hk_vertical_l",
        label="HK Vertical-L",
        module="profiles.hk_vertical_l.convert",
        description="源账单 sheet「Hong Kong-L」→ Hong Kong PN",
    ),
    "tw_payroll_calc": ConvertEngine(
        engine_id="tw_payroll_calc",
        label="TW Payroll Calculation",
        module="profiles.tw_payroll_calc.convert",
        description="源账单 sheet「Payroll calculation」→ Taiwan PN",
    ),
}


def get_engine(engine_id: str) -> ConvertEngine:
    engine = ENGINES.get(engine_id)
    if engine is None:
        known = ", ".join(sorted(ENGINES))
        raise KeyError(f"未知转换引擎「{engine_id}」，已知: {known}")
    return engine


def list_engines() -> list[ConvertEngine]:
    return [ENGINES[k] for k in sorted(ENGINES)]
