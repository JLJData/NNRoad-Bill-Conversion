# -*- coding: utf-8 -*-
"""按地区复用的 PN 母版路径（China / Hong Kong / Taiwan / UK / UAE / Pakistan / Italy / India / Cyprus）。"""
from __future__ import annotations

from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
TEMPLATES_DIR = BASE_DIR / "templates"

# region 显示名 → 目录 slug
REGION_DIRS: dict[str, str] = {
    "China": "china",
    "Hong Kong": "hongkong",
    "Taiwan": "taiwan",
    "UK": "uk",
    "UAE": "uae",
    "Pakistan": "pakistan",
    "Italy": "italy",
    "India": "india",
    "Cyprus": "cyprus",
}

REGION_TEMPLATE_FILENAME = "template.xlsx"


def get_region_template(region: str) -> Path:
    """返回指定地区的默认 PN 母版路径。"""
    slug = REGION_DIRS.get(region)
    if slug is None:
        known = ", ".join(sorted(REGION_DIRS))
        raise KeyError(f"未知地区「{region}」，已知: {known}")
    return TEMPLATES_DIR / slug / REGION_TEMPLATE_FILENAME


def list_regions() -> list[str]:
    return sorted(REGION_DIRS)
