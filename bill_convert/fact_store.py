# -*- coding: utf-8 -*-
"""转换配置 mappingJson.factStore：供应商插件跨期记忆。"""
from __future__ import annotations

from copy import deepcopy
from datetime import datetime, timezone
from typing import Any


FACT_STORE_KEY = "factStore"
ARTIFACT_BATCH_KEY = "artifactBatch"


def _as_store(mapping: dict[str, Any] | None) -> dict[str, Any]:
    if not isinstance(mapping, dict):
        return {}
    raw = mapping.get(FACT_STORE_KEY)
    return dict(raw) if isinstance(raw, dict) else {}


def get_fact_entry(mapping: dict[str, Any] | None, key: str) -> dict[str, Any] | None:
    store = _as_store(mapping)
    entry = store.get(key)
    return dict(entry) if isinstance(entry, dict) else None


def get_fact_value(mapping: dict[str, Any] | None, key: str, default: Any = None) -> Any:
    entry = get_fact_entry(mapping, key)
    if entry is None:
        return default
    return entry.get("value", default)


def set_fact_value(
    mapping: dict[str, Any],
    key: str,
    value: Any,
    *,
    meta: dict[str, Any] | None = None,
) -> dict[str, Any]:
    out = deepcopy(mapping) if isinstance(mapping, dict) else {}
    store = _as_store(out)
    entry: dict[str, Any] = {"value": value, "updatedAt": datetime.now(timezone.utc).isoformat()}
    if meta:
        entry.update({k: v for k, v in meta.items() if v is not None})
    store[key] = entry
    out[FACT_STORE_KEY] = store
    return out


def merge_fact_updates(mapping: dict[str, Any] | None, updates: dict[str, Any] | None) -> dict[str, Any]:
    """updates: {factKey: {value, ...meta}} 或 {factKey: scalar}"""
    out = deepcopy(mapping) if isinstance(mapping, dict) else {}
    if not updates:
        return out
    store = _as_store(out)
    now = datetime.now(timezone.utc).isoformat()
    for key, raw in updates.items():
        if not key:
            continue
        if isinstance(raw, dict) and "value" in raw:
            entry = dict(raw)
            entry.setdefault("updatedAt", now)
        else:
            entry = {"value": raw, "updatedAt": now}
        store[str(key)] = entry
    out[FACT_STORE_KEY] = store
    return out


def get_batch_facts(mapping: dict[str, Any] | None) -> dict[str, Any]:
    if not isinstance(mapping, dict):
        return {}
    raw = mapping.get(ARTIFACT_BATCH_KEY)
    return dict(raw) if isinstance(raw, dict) else {}


def set_batch_facts(mapping: dict[str, Any] | None, facts: dict[str, Any] | None) -> dict[str, Any]:
    out = deepcopy(mapping) if isinstance(mapping, dict) else {}
    if facts:
        out[ARTIFACT_BATCH_KEY] = dict(facts)
    else:
        out.pop(ARTIFACT_BATCH_KEY, None)
    return out
