# -*- coding: utf-8 -*-
"""
询价比价引擎。

比价的三个专业口径，缺一个就容易被「低价陷阱」骗：
1. 不含税单价     —— 税率不同（13% / 3% / 免税）必须拉平，否则比的是税不是价
2. 到货等效单价   —— 把运费、最小起订量摊进去，避免「单价低但运费高」
3. 资质与交期前置 —— 合规不达标的低价报价直接排除，这是医药化工行业的硬约束
"""
from __future__ import annotations

from typing import Dict, List, Optional, Sequence

from .models import ItemComparison, Quote


def _group_by_item(quotes: Sequence[Quote]) -> Dict[str, List[Quote]]:
    groups: Dict[str, List[Quote]] = {}
    for q in quotes:
        key = q.item_code or q.item_name or f"__{id(q)}"
        groups.setdefault(key, []).append(q)
    return groups


def _valid(q: Quote) -> bool:
    """有效报价：有供应商、单价为正。"""
    return bool(q.supplier) and q.unit_price > 0


def compare_items(quotes: Sequence[Quote],
                  require_cert: bool = True) -> List[ItemComparison]:
    """
    按物料聚合报价并排序。

    require_cert=True 时，资质不全的报价不参与「最低价」评选，
    但会保留在明细里并标红——合规优先于价格。
    """
    results: List[ItemComparison] = []
    for key, group in _group_by_item(quotes).items():
        valid = [q for q in group if _valid(q)]
        if not valid:
            continue
        pool = [q for q in valid if q.cert_ok] if require_cert else list(valid)
        if not pool:                       # 全部资质不全 → 退回全量并报警
            pool = list(valid)
        ranked = sorted(pool, key=lambda q: (q.landed_unit_price, q.lead_time))

        # 影子最低价：忽略资质，仅用于风险提示（低于它说明有被合规卡掉的低价）
        shadow = min(valid, key=lambda q: (q.landed_unit_price, q.lead_time))

        base = valid[0]
        comp = ItemComparison(
            item_code=base.item_code,
            item_name=base.item_name,
            spec=base.spec,
            unit=base.unit,
            qty=max(q.qty for q in valid),
            quotes=list(valid),
            ranked=ranked,
            best=ranked[0] if ranked else None,
            second=ranked[1] if len(ranked) > 1 else None,
            worst=ranked[-1] if len(ranked) > 1 else ranked[0] if ranked else None,
            shadow_best=shadow,
        )
        results.append(comp)
    results.sort(key=lambda c: c.item_code)
    return results


def build_matrix(comparisons: Sequence[ItemComparison]) -> Dict[str, dict]:
    """
    生成「物料 × 供应商」比价矩阵（宽表），采购评审会上最常看的一张表。
    单元格为到货等效单价（不含税），并标注最优。
    """
    suppliers: List[str] = []
    for c in comparisons:
        for q in c.quotes:
            if q.supplier and q.supplier not in suppliers:
                suppliers.append(q.supplier)
    suppliers.sort()

    rows = []
    for c in comparisons:
        row = {
            "物料编码": c.item_code,
            "物料名称": c.item_name,
            "规格型号": c.spec,
            "单位": c.unit,
            "需求数量": c.qty,
            "报价家数": c.supplier_count,
        }
        for s in suppliers:
            qs = [q for q in c.quotes if q.supplier == s]
            row[s] = round(qs[0].landed_unit_price, 4) if qs else None
        best_sup = c.best.supplier if c.best else ""
        row["最优供应商"] = best_sup
        row["最优到货单价"] = round(c.best.landed_unit_price, 4) if c.best else None
        row["价差率"] = round(c.spread_rate, 4)
        row["风险提示"] = "；".join(c.risk_flags())
        row["建议"] = c.recommendation()
        rows.append(row)
    return {"suppliers": suppliers, "rows": rows}


def summary(comparisons: Sequence[ItemComparison]) -> Dict[str, float]:
    """汇总关键指标，用于报告首页与管理层汇报。"""
    total_best = sum(c.best_amount for c in comparisons)
    total_worst = sum((c.worst.landed_cost if c.worst else 0.0) for c in comparisons)
    total_history_saving = sum(c.saving_vs_history for c in comparisons)
    history_base = sum(
        (c.best.history_price * c.best.order_qty) if c.best and c.best.history_price > 0 else 0.0
        for c in comparisons
    )
    single = sum(1 for c in comparisons if c.single_source)
    risky = sum(1 for c in comparisons if c.risk_flags())
    return {
        "物料数": len(comparisons),
        "报价条目数": sum(len(c.quotes) for c in comparisons),
        "参与供应商数": len({q.supplier for c in comparisons for q in c.quotes if q.supplier}),
        "单一来源物料数": single,
        "存在风险提示的物料数": risky,
        "选定总金额(不含税)": round(total_best, 2),
        "若按最高价采购金额": round(total_worst, 2),
        "比价节约金额": round(total_worst - total_best, 2),
        "比价节约率": round((total_worst - total_best) / total_worst, 4) if total_worst else 0.0,
        "相对历史价节约": round(total_history_saving, 2),
        "历史价基准金额": round(history_base, 2),
        "历史降本率": round(total_history_saving / history_base, 4) if history_base else 0.0,
    }
