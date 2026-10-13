# -*- coding: utf-8 -*-
"""
MRO 库存与补货分析：ABC 分类 + EOQ + 安全库存 + 再订货点。

MRO（维护、维修、运营物资）的典型特征是「种类多、单值低、需求波动大」，
用生产物料的套路管 MRO 只会把管理费吃掉。所以这里除了标准模型，
还额外输出「下单频次」和「管理成本」——C 类长尾物料降本的关键
不是压单价，而是减少下单次数（打包招标 / 双箱法 / 年度框架协议）。
"""
from __future__ import annotations

from typing import Dict, List, Sequence

from .models import InventoryItem


def classify_abc(items: Sequence[InventoryItem],
                 a_cut: float = 0.70,
                 b_cut: float = 0.90) -> List[InventoryItem]:
    """按年耗用金额降序做帕累托分类，并把结果写回对象的 _abc 属性。"""
    ranked = sorted(items, key=lambda x: x.annual_value, reverse=True)
    total = sum(x.annual_value for x in ranked)
    cum = 0.0
    for it in ranked:
        if total <= 0:
            it._abc = "C"
            continue
        cum += it.annual_value
        ratio = cum / total
        if ratio <= a_cut:
            it._abc = "A"
        elif ratio <= b_cut:
            it._abc = "B"
        else:
            it._abc = "C"
    return ranked


def analyze(items: Sequence[InventoryItem]) -> Dict:
    ranked = classify_abc(items)
    rows: List[Dict] = []
    for it in ranked:
        rows.append({
            "物料编码": it.item_code,
            "物料名称": it.item_name,
            "物料类别": it.category,
            "单位": it.unit,
            "单价": round(it.unit_price, 4),
            "年耗用量": round(it.annual_qty, 2),
            "年耗用金额": round(it.annual_value, 2),
            "ABC分类": it.abc_class,
            "经济订货批量EOQ": round(it.eoq, 2),
            "年订货次数": round(it.order_times, 2),
            "日均需求": round(it.daily_demand, 4),
            "安全库存": round(it.safety_stock, 2),
            "再订货点ROP": round(it.reorder_point, 2),
            "当前库存": round(it.current_stock, 2),
            "可支撑天数": round(it.stock_days, 1),
            "是否需补货": "是" if it.need_reorder else "否",
            "补货策略建议": it.policy(),
        })

    by_class: Dict[str, Dict[str, float]] = {}
    for it in ranked:
        b = by_class.setdefault(it.abc_class, {"数量": 0, "金额": 0.0, "年订货次数": 0.0})
        b["数量"] += 1
        b["金额"] += it.annual_value
        b["年订货次数"] += it.order_times

    # 管理成本视角：C 类物料种类多但金额小，合并下单能直接砍掉下单手续费
    total_orders = sum(it.order_times for it in ranked)
    order_cost_unit = ranked[0].order_cost if ranked else 120.0
    c_items = [it for it in ranked if it.abc_class == "C"]
    c_orders = sum(it.order_times for it in c_items)
    potential = c_orders * order_cost_unit * 0.6   # 打包后保守估计可省 60% 下单次数

    return {
        "rows": rows,
        "by_class": {k: {
            "物料数": v["数量"],
            "年金额": round(v["金额"], 2),
            "金额占比": round(v["金额"] / sum(x["金额"] for x in by_class.values()), 4)
            if sum(x["金额"] for x in by_class.values()) else 0.0,
            "年订货次数": round(v["年订货次数"], 1),
        } for k, v in sorted(by_class.items())},
        "全年订货次数": round(total_orders, 1),
        "年订货管理成本": round(total_orders * order_cost_unit, 2),
        "C类合并下单可节约管理费": round(potential, 2),
    }
