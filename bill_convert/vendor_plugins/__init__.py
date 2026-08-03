# -*- coding: utf-8 -*-
"""供应商旁路插件：按 pdfProfile 注册，互不影响。"""
from bill_convert.vendor_plugins.registry import get_plugins_for_profile, list_plugins

__all__ = ["get_plugins_for_profile", "list_plugins"]
