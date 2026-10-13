# -*- coding: utf-8 -*-
"""
数据读取：支持 Excel(.xlsx) 与 CSV，表头中英文混写也能识别。

不依赖 pandas：Excel 用 openpyxl 只读模式解析，CSV 用标准库 csv。
openpyxl 缺失时自动退化为「仅 CSV」，保证最小可用。
"""
from __future__ import annotations

import csv
import io
import json
import os
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

from .models import (
    FIELD_ALIASES, Quote, InventoryItem, _to_bool, _to_float, _to_str,
)


# 反向索引：别名 -> 标准字段名
_ALIAS_INDEX: Dict[str, str] = {}
for _field, _aliases in FIELD_ALIASES.items():
    for _a in _aliases:
        _ALIAS_INDEX[_a.strip().lower()] = _field


def normalize_header(h: Any) -> Optional[str]:
    """把表头映射成标准字段名，识别不了返回 None。"""
    if h is None:
        return None
    key = _to_str(h).lower().replace(" ", "").replace("（", "(").replace("）", ")")
    return _ALIAS_INDEX.get(key)


def _rows_from_csv(path_or_fp, encoding: str = "utf-8-sig") -> Tuple[List[str], List[List[Any]]]:
    if isinstance(path_or_fp, (str, os.PathLike)):
        with open(path_or_fp, "r", encoding=encoding, newline="") as f:
            return _read_csv_reader(f)
    if isinstance(path_or_fp, bytes):
        text = path_or_fp.decode(encoding, errors="replace")
        return _read_csv_reader(io.StringIO(text))
    return _read_csv_reader(path_or_fp)


def _read_csv_reader(fp) -> Tuple[List[str], List[List[Any]]]:
    reader = csv.reader(fp)
    rows = [r for r in reader if any(_to_str(c) for c in r)]
    if not rows:
        return [], []
    return rows[0], rows[1:]


def _rows_from_xlsx(path_or_fp) -> Tuple[List[str], List[List[Any]]]:
    try:
        from openpyxl import load_workbook
    except ImportError as exc:  # pragma: no cover
        raise RuntimeError(
            "读取 .xlsx 需要 openpyxl，请先执行： pip install openpyxl"
        ) from exc
    kwargs = {"data_only": True, "read_only": True}
    if isinstance(path_or_fp, (str, os.PathLike)):
        wb = load_workbook(path_or_fp, **kwargs)
    else:
        wb = load_workbook(io.BytesIO(path_or_fp.read()), **kwargs)
    ws = wb.worksheets[0]
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    rows = [r for r in rows if any(_to_str(c) for c in r)]
    wb.close()
    if not rows:
        return [], []
    return rows[0], rows[1:]


def read_table(source, sheet_hint: str = "") -> Tuple[List[str], List[List[Any]]]:
    """统一入口：根据后缀或内容判断 xlsx / csv。"""
    if isinstance(source, (str, os.PathLike)):
        p = str(source).lower()
        if p.endswith((".xlsx", ".xlsm")):
            return _rows_from_xlsx(source)
        return _rows_from_csv(source)
    # 类文件对象（Streamlit 上传）
    name = _to_str(getattr(source, "name", "")).lower()
    if name.endswith((".xlsx", ".xlsm")):
        return _rows_from_xlsx(source)
    return _rows_from_csv(source)


def _num_field(rec: Dict[str, Any], key: str, default: float) -> float:
    """
    取数值字段：空值用默认，填了 0 就必须是 0。

    这里不能用 `x or default` 的写法——0 是合法值（免税税率 0%、现货交期 0 天），
    被 `or` 吃掉会静默算错，是这类工具最容易埋雷的地方。
    """
    v = rec.get(key)
    if v is None or _to_str(v) == "":
        return default
    return _to_float(v, default)


def _dicts_from_table(source) -> List[Dict[str, Any]]:
    header, rows = read_table(source)
    mapped = [normalize_header(h) for h in header]
    out: List[Dict[str, Any]] = []
    for r in rows:
        rec: Dict[str, Any] = {}
        for i, col in enumerate(mapped):
            if not col:
                continue
            rec[col] = r[i] if i < len(r) else None
        if any(v not in (None, "") for v in rec.values()):
            out.append(rec)
    return out


# --------------------------------------------------------------------------
# 报价数据
# --------------------------------------------------------------------------
def load_quotes(source) -> List[Quote]:
    """读取报价明细，返回 Quote 列表。"""
    records = _dicts_from_table(source)
    quotes: List[Quote] = []
    for rec in records:
        q = Quote(
            item_code=_to_str(rec.get("item_code")),
            item_name=_to_str(rec.get("item_name")),
            category=_to_str(rec.get("category")),
            spec=_to_str(rec.get("spec")),
            unit=_to_str(rec.get("unit"), "EA"),
            qty=_to_float(rec.get("qty")),
            supplier=_to_str(rec.get("supplier")),
            unit_price=_num_field(rec, "unit_price", 0.0),
            tax_rate=_num_field(rec, "tax_rate", 0.13),
            tax_included=_to_bool(rec.get("tax_included"), True),
            lead_time=_num_field(rec, "lead_time", 0.0),
            moq=_num_field(rec, "moq", 0.0),
            freight=_num_field(rec, "freight", 0.0),
            payment_terms=_to_str(rec.get("payment_terms")),
            quality_score=_num_field(rec, "quality_score", 80.0),
            service_score=_num_field(rec, "service_score", 80.0),
            cert_ok=_to_bool(rec.get("cert_ok"), True),
            history_price=_num_field(rec, "history_price", 0.0),
            remark=_to_str(rec.get("remark")),
        )
        if q.supplier and q.unit_price >= 0:
            quotes.append(q)
    return quotes


# --------------------------------------------------------------------------
# 库存数据
# --------------------------------------------------------------------------
INV_ALIASES: Dict[str, tuple] = {
    "item_code":    ("物料编码", "编码", "item_code", "sku"),
    "item_name":    ("物料名称", "品名", "item_name", "name"),
    "category":     ("物料类别", "品类", "category"),
    "unit":         ("单位", "unit", "uom"),
    "unit_price":   ("单价", "采购单价", "unit_price", "price"),
    "annual_qty":   ("年耗用量", "年用量", "年需求量", "annual_qty", "annual_demand"),
    "order_cost":   ("单次订货成本", "订货成本", "order_cost"),
    "holding_rate": ("持有成本率", "年持有成本率", "holding_rate"),
    "lead_time":    ("采购提前期", "提前期", "交期", "lead_time"),
    "demand_std":   ("日需求标准差", "需求波动", "demand_std"),
    "cv":           ("变异系数", "cv"),
    "service_level":("服务水平", "service_level"),
    "current_stock":("当前库存", "现有库存", "current_stock", "stock"),
}

_INV_INDEX: Dict[str, str] = {}
for _f, _as in INV_ALIASES.items():
    for _a in _as:
        _INV_INDEX[_a.strip().lower()] = _f


def load_inventory(source) -> List[InventoryItem]:
    header, rows = read_table(source)
    mapped = [ _INV_INDEX.get(_to_str(h).strip().lower()) for h in header ]
    items: List[InventoryItem] = []
    for r in rows:
        rec = {c: (r[i] if i < len(r) else None) for i, c in enumerate(mapped) if c}
        if not rec:
            continue
        it = InventoryItem(
            item_code=_to_str(rec.get("item_code")),
            item_name=_to_str(rec.get("item_name")),
            category=_to_str(rec.get("category")),
            unit=_to_str(rec.get("unit"), "EA"),
            unit_price=_num_field(rec, "unit_price", 0.0),
            annual_qty=_num_field(rec, "annual_qty", 0.0),
            order_cost=_num_field(rec, "order_cost", 120.0),
            holding_rate=_num_field(rec, "holding_rate", 0.25),
            lead_time=_num_field(rec, "lead_time", 14.0),
            demand_std=_num_field(rec, "demand_std", 0.0),
            cv=_num_field(rec, "cv", 0.35),
            service_level=_num_field(rec, "service_level", 0.95),
            current_stock=_num_field(rec, "current_stock", 0.0),
        )
        if it.item_name or it.item_code:
            items.append(it)
    return items


def load_weights(path: str) -> Dict[str, float]:
    """从 JSON 读评分权重，文件不存在就用默认。"""
    from .models import DEFAULT_WEIGHTS
    if not path or not os.path.exists(path):
        return dict(DEFAULT_WEIGHTS)
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    w = dict(DEFAULT_WEIGHTS)
    for k, v in (data or {}).items():
        if k in w:
            w[k] = float(v)
    total = sum(w.values())
    if total <= 0:
        return dict(DEFAULT_WEIGHTS)
    return {k: v / total for k, v in w.items()}


def template_rows(kind: str = "quote") -> Tuple[List[str], List[List[Any]]]:
    """生成模板表头 + 一行示例，供 `mro init` 输出空白模板。"""
    if kind == "inventory":
        header = ["物料编码", "物料名称", "物料类别", "单位", "单价", "年耗用量",
                  "单次订货成本", "持有成本率", "采购提前期", "日需求标准差",
                  "变异系数", "服务水平", "当前库存"]
        sample = ["MRO-0001", "示例物料", "备品备件", "EA", 100, 1200,
                  120, 0.25, 14, "", 0.35, 0.95, 200]
        return header, [sample]
    header = ["物料编码", "物料名称", "物料类别", "规格型号", "单位", "需求数量",
              "供应商", "单价", "税率", "是否含税", "交期", "最小起订量", "运费",
              "付款条件", "质量评分", "服务评分", "资质合规", "历史单价", "备注"]
    sample = ["MRO-0001", "示例物料", "备品备件", "DN50", "EA", 100,
              "供应商A", 88.5, 0.13, "是", 7, 10, 0,
              "月结30天", 90, 85, "是", 92, ""]
    return header, [sample]
