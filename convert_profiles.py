# -*- coding: utf-8 -*-
"""账单转换配置：供应商 + 客户 → 地区 + 转换引擎

正式路由配置已迁到 Office / Portal 表 portal_bill_convert_profile
（见 hrone-office-abp/sql/portal_bill_convert_profile.sql）。

本文件保留桌面 GUI / CLI 本地开发用的默认 PROFILES；线上执行应以
Office 下发的 engine_id + region + pn_meta 为准，勿再写死新客户映射。
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from engines import ConvertEngine, get_engine
from region_templates import get_region_template

BASE_DIR = Path(__file__).resolve().parent


@dataclass(frozen=True)
class ConvertProfile:
    profile_id: str
    supplier: str
    customer: str
    region: str  # 输出地区 → 决定 PN 母版
    engine: str  # 转换引擎 → 决定如何读源账单
    fx_cell: str
    fx_description: str
    output_prefix: str
    pn_customer_name: str  # PN!B8 默认客户名称
    pn_customer_id: str  # PN!B9 默认客户 ID
    pn_billing_address: str  # PN!B10 默认账单地址

    @property
    def convert_engine(self) -> ConvertEngine:
        return get_engine(self.engine)

    @property
    def module(self) -> str:
        """转换脚本模块（由 engine 解析）。"""
        return self.convert_engine.module

    @property
    def template(self) -> Path:
        """按地区复用的 PN 母版（templates/{region}/template.xlsx）。"""
        return get_region_template(self.region)


PROFILES: list[ConvertProfile] = [
    ConvertProfile(
        profile_id="hrone_hermetic",  # Office 种子 profile_code（示例客户对）
        supplier="HROne Co., Ltd.",
        customer="Hermetic",
        region="China",
        engine="china_hrone",
        fx_cell="PN!B29",
        fx_description="USD → CNY（exchangerate-api）",
        output_prefix="PN_Hermetic",
        pn_customer_name="Hermetic Solutions Group, Inc",
        pn_customer_id="CUS15253",
        pn_billing_address="Eight Neshaminy Interplex,Suite 221, Trevose. PA 19053",
    ),
    ConvertProfile(
        profile_id="topfdi_uecorp",  # Office 种子 profile_code（示例客户对）
        supplier="Top FDI",
        customer="UE Corp",
        region="Hong Kong",
        engine="hk_vertical_l",
        fx_cell="PN!B28",
        fx_description="USD → HKD × 0.97（exchangerate-api）",
        output_prefix="PN_UECorp",
        pn_customer_name="UE Corp",
        pn_customer_id="CUS1503",
        pn_billing_address="168 Georgetown Rd., Canonsburg, Pennsylvanis, USA",
    ),
    ConvertProfile(
        profile_id="peoplesearch_coralsea",  # Office 种子 profile_code（示例客户对）
        supplier="People Search",
        customer="Coral Sea",
        region="Taiwan",
        engine="tw_payroll_calc",
        fx_cell="PN!B31",
        fx_description="USD → TWD（exchangerate-api）",
        output_prefix="PN_CoralSea",
        pn_customer_name="Coral Sea",
        pn_customer_id="CUS1516",
        pn_billing_address="4205/41 Williams Esplanade, PALM COVE, QLD, Australia",
    ),
]


def list_suppliers() -> list[str]:
    return sorted({p.supplier for p in PROFILES})


def list_customers(supplier: str) -> list[str]:
    return sorted({p.customer for p in PROFILES if p.supplier == supplier})


def get_profile(supplier: str, customer: str) -> ConvertProfile | None:
    for p in PROFILES:
        if p.supplier == supplier and p.customer == customer:
            return p
    return None
