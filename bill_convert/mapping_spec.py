# -*- coding: utf-8 -*-
from __future__ import annotations

from typing import Any


def mapping_section(mapping: dict[str, Any] | None, key: str) -> dict[str, Any]:
    if not isinstance(mapping, dict):
        return {}
    block = mapping.get(key)
    return block if isinstance(block, dict) else {}
