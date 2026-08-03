# -*- coding: utf-8 -*-
from __future__ import annotations

from pathlib import Path
from typing import Any, Protocol


class VendorPlugin(Protocol):
    """
    供应商专属旁路逻辑插件。
    - 只在匹配的 pdf_profile_id 下加载
    - classify/parse 在 vendor→源表阶段；apply 在地区引擎写完主表之后
    """

    plugin_id: str
    pdf_profile_ids: tuple[str, ...]

    def classify_path(self, path: Path) -> bool:
        """是否为本插件旁路文件（非主员工源）。"""
        ...

    def parse_artifacts(self, paths: list[Path]) -> dict[str, Any]:
        """
        解析旁路文件 → 本批事实（扁平 dict）。
        建议 key 带供应商前缀，如 auxilium.admin_fee.total_vat
        """
        ...

    def apply_to_workbook(
        self,
        wb,
        *,
        mapping: dict[str, Any],
        batch_facts: dict[str, Any],
        warnings: list[str],
        employee_count: int = 1,
    ) -> dict[str, Any] | None:
        """
        写入工作簿派生字段。
        返回需在转换成功后 commit 到 mapping.factStore 的更新；无则 None。
        """
        ...
