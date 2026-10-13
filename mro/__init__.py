# -*- coding: utf-8 -*-
"""
mro-sourcing-toolkit
====================
面向 MRO / 间接采购的轻量分析工具集：询价比价、供应商评分、库存补货测算。

核心引擎零第三方依赖（纯标准库），仅 Excel 读写需要 openpyxl。
"""
__version__ = "1.0.0"
__author__ = "Zhang Yujie"

from .models import Quote, InventoryItem, SupplierScore, ItemComparison  # noqa: F401
from .compare import compare_items, build_matrix, summary  # noqa: F401
from .scoring import score_suppliers, to_rows  # noqa: F401
from .inventory import analyze as analyze_inventory  # noqa: F401

__all__ = [
    "Quote", "InventoryItem", "SupplierScore", "ItemComparison",
    "compare_items", "build_matrix", "summary",
    "score_suppliers", "to_rows", "analyze_inventory",
]
