# -*- coding: utf-8 -*-
from __future__ import annotations

from typing import Any

from bill_convert.vendor_plugins.auxilium_admin_fee_business_tax import AuxiliumAdminFeeBusinessTaxPlugin

_PLUGINS: list[Any] = [
    AuxiliumAdminFeeBusinessTaxPlugin(),
]


def list_plugins() -> list[Any]:
    return list(_PLUGINS)


def get_plugins_for_profile(pdf_profile_id: str | None) -> list[Any]:
    pid = (pdf_profile_id or "").strip()
    if not pid:
        return []
    return [p for p in _PLUGINS if pid in getattr(p, "pdf_profile_ids", ())]
