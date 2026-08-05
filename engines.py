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
    "china_payroll_calc": ConvertEngine(
        engine_id="china_payroll_calc",
        label="China Payroll Calculation",
        module="profiles.china_payroll_calc.convert",
        description="源账单 sheet「计算结果」→ China PN",
    ),
    "hk_payroll_calc": ConvertEngine(
        engine_id="hk_payroll_calc",
        label="HK Payroll Calculation",
        module="profiles.hk_payroll_calc.convert",
        description="源账单 sheet「Hong Kong-L」→ Hong Kong PN",
    ),
    "tw_payroll_calc": ConvertEngine(
        engine_id="tw_payroll_calc",
        label="TW Payroll Calculation",
        module="profiles.tw_payroll_calc.convert",
        description="源账单 sheet「Payroll calculation」→ Taiwan PN",
    ),
    "uk_payroll_calc": ConvertEngine(
        engine_id="uk_payroll_calc",
        label="UK Payroll Calculation",
        module="profiles.uk_payroll_calc.convert",
        description="源账单 sheet「UK-L」（竖表）→ UK PN；PDF 版式可选 EOR / TopSource",
    ),
    "uae_payroll_calc": ConvertEngine(
        engine_id="uae_payroll_calc",
        label="UAE Payroll Calculation",
        module="profiles.uae_payroll_calc.convert",
        description="源账单 sheet「UAE-L」→ UAE PN；Excel 版式可选 Auxilium",
    ),
    "pakistan_payroll_calc": ConvertEngine(
        engine_id="pakistan_payroll_calc",
        label="Pakistan Payroll Calculation",
        module="profiles.pakistan_payroll_calc.convert",
        description="源账单 sheet「Pakistan-L」→ Pakistan PN；PDF 版式可选 Panda Work Global",
    ),
    "italy_payroll_calc": ConvertEngine(
        engine_id="italy_payroll_calc",
        label="Italy Payroll Calculation",
        module="profiles.italy_payroll_calc.convert",
        description="源账单 sheet「Italy-L」→ Italy PN；Excel 版式可选 SafeGuard",
    ),
    "india_payroll_calc": ConvertEngine(
        engine_id="india_payroll_calc",
        label="India Payroll Calculation",
        module="profiles.india_payroll_calc.convert",
        description="源账单 sheet「India-L」→ India PN；PDF 版式可选 Biz Solutions",
    ),
}

# 旧引擎 id 兼容
ENGINES["china_hrone"] = ENGINES["china_payroll_calc"]
ENGINES["hk_vertical_l"] = ENGINES["hk_payroll_calc"]

_ALIAS_ENGINE_IDS = frozenset({"china_hrone", "hk_vertical_l"})


def get_engine(engine_id: str) -> ConvertEngine:
    engine = ENGINES.get(engine_id)
    if engine is None:
        known = ", ".join(sorted(k for k in ENGINES if k not in _ALIAS_ENGINE_IDS))
        raise KeyError(f"未知转换引擎「{engine_id}」，已知: {known}")
    return engine


def list_engines() -> list[ConvertEngine]:
    # 列表不重复暴露兼容别名
    return [ENGINES[k] for k in sorted(ENGINES) if k not in _ALIAS_ENGINE_IDS]
