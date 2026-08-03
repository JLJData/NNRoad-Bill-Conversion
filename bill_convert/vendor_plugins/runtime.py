# -*- coding: utf-8 -*-
from __future__ import annotations

from pathlib import Path
from typing import Any

from bill_convert.vendor_plugins.registry import get_plugins_for_profile


def split_main_and_artifacts(
    paths: list[Path],
    *,
    pdf_profile_id: str | None,
) -> tuple[list[Path], list[Path], list[str]]:
    """按已注册插件拆分主源与旁路文件。"""
    plugins = get_plugins_for_profile(pdf_profile_id)
    if not plugins:
        return list(paths), [], []
    main: list[Path] = []
    artifacts: list[Path] = []
    warnings: list[str] = []
    for p in paths:
        path = Path(p)
        claimed = False
        for plugin in plugins:
            try:
                if plugin.classify_path(path):
                    artifacts.append(path)
                    claimed = True
                    break
            except Exception as exc:
                warnings.append(f"插件 {plugin.plugin_id} 分类失败 {path.name}: {exc}")
        if not claimed:
            main.append(path)
    return main, artifacts, warnings


def parse_artifact_facts(
    artifact_paths: list[Path],
    *,
    pdf_profile_id: str | None,
) -> tuple[dict[str, Any], list[str]]:
    plugins = get_plugins_for_profile(pdf_profile_id)
    facts: dict[str, Any] = {}
    warnings: list[str] = []
    if not artifact_paths or not plugins:
        return facts, warnings
    for plugin in plugins:
        mine = [p for p in artifact_paths if plugin.classify_path(Path(p))]
        if not mine:
            continue
        try:
            parsed = plugin.parse_artifacts(mine) or {}
        except Exception as exc:
            warnings.append(f"插件 {plugin.plugin_id} 解析失败: {exc}")
            continue
        for w in parsed.pop("_warnings", []) or []:
            warnings.append(str(w))
        facts.update(parsed)
    return facts, warnings


def apply_vendor_plugins(
    wb,
    *,
    pdf_profile_id: str | None,
    mapping: dict[str, Any],
    batch_facts: dict[str, Any] | None,
    warnings: list[str],
    employee_count: int = 1,
) -> dict[str, Any]:
    """对工作簿应用插件；返回需 commit 的 factStore 更新。"""
    plugins = get_plugins_for_profile(pdf_profile_id)
    commits: dict[str, Any] = {}
    facts = dict(batch_facts or {})
    for plugin in plugins:
        try:
            upd = plugin.apply_to_workbook(
                wb,
                mapping=mapping,
                batch_facts=facts,
                warnings=warnings,
                employee_count=employee_count,
            )
        except Exception as exc:
            warnings.append(f"插件 {plugin.plugin_id} 写入失败: {exc}")
            continue
        if isinstance(upd, dict):
            commits.update(upd)
    return commits
