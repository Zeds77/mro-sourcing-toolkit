# -*- coding: utf-8 -*-
"""
供应商综合评分与份额建议。

价格只是采购决策的一个维度。这里用可配置的加权评分模型，
把「价格 / 交期 / 质量 / 服务 / 资质」五维拉到同一标尺上：
  - 价格分：  最低到货等效单价 / 本家单价 × 100（越低分越高）
  - 交期分：  在参与供应商中做 min-max 归一化，最快得 100
  - 质量分：  来料合格率等实测分（缺省 80）
  - 服务分：  响应、售后、开票配合度（缺省 80）
  - 资质分：  证照齐全 100，缺失 0（医药化工场景下建议一票否决）

份额建议遵循双源/三源策略：第一名 60%~70%，第二名 30%~40%，
避免独家供货的断供风险，同时保留议价筹码。
"""
from __future__ import annotations

from typing import Dict, List, Optional, Sequence

from .models import (
    DEFAULT_WEIGHTS, ItemComparison, Quote, SupplierScore, WEIGHT_LABELS,
)


def _valid_price(q: Quote) -> bool:
    return q.landed_unit_price > 0


def _minmax_score(value: float, lo: float, hi: float, reverse: bool = True) -> float:
    """把指标归一化到 0~100。reverse=True 表示数值越小越好。"""
    if hi <= lo:
        return 100.0
    x = (value - lo) / (hi - lo)
    x = min(max(x, 0.0), 1.0)
    return (1.0 - x) * 100.0 if reverse else x * 100.0


def score_suppliers(comparisons: Sequence[ItemComparison],
                    weights: Optional[Dict[str, float]] = None,
                    cert_veto: bool = True) -> List[SupplierScore]:
    w = dict(weights or DEFAULT_WEIGHTS)
    # 权重自动归一化：用户给的是比重还是百分比都能用，避免口径不一致导致分数失控
    _sum = sum(v for v in w.values() if v > 0)
    if _sum > 0 and abs(_sum - 1.0) > 1e-9:
        w = {k: (v / _sum if v > 0 else 0.0) for k, v in w.items()}

    # 收集每家供应商在各物料上的表现
    stats: Dict[str, Dict] = {}
    for c in comparisons:
        if not c.best:
            continue
        # 价格基准取「全部有效报价」的最低价（不受资质过滤影响），
        # 这样价格分不会超过 100，资质问题交给 cert 维度与 veto 处理。
        best_price = min((q.landed_unit_price for q in c.quotes if _valid_price(q)),
                         default=c.best.landed_unit_price)
        for q in c.quotes:
            s = q.supplier
            st = stats.setdefault(s, {
                "price_dev": [],      # 相对最优价的溢价率
                "lead_times": [],
                "quality": [],
                "service": [],
                "cert": [],
                "amount": 0.0,
                "quotes": 0,
                "wins": 0,
                "saving": [],
            })
            if best_price > 0:
                st["price_dev"].append((q.landed_unit_price - best_price) / best_price)
            st["lead_times"].append(q.lead_time)
            st["quality"].append(q.quality_score)
            st["service"].append(q.service_score)
            st["cert"].append(1.0 if q.cert_ok else 0.0)
            st["amount"] += q.landed_cost
            st["quotes"] += 1
            if c.best.supplier == s:
                st["wins"] += 1
            if q.history_price > 0:
                st["saving"].append(q.saving_rate or 0.0)

    if not stats:
        return []

    # 交期归一化需要全局极值
    all_lt = [lt for st in stats.values() for lt in st["lead_times"]]
    lt_lo, lt_hi = (min(all_lt), max(all_lt)) if all_lt else (0.0, 0.0)

    scores: List[SupplierScore] = []
    for s, st in stats.items():
        price_dev_avg = sum(st["price_dev"]) / len(st["price_dev"]) if st["price_dev"] else 0.0
        price_score = max(0.0, 100.0 - price_dev_avg * 100.0)
        lt_avg = sum(st["lead_times"]) / len(st["lead_times"]) if st["lead_times"] else 0.0
        lt_score = _minmax_score(lt_avg, lt_lo, lt_hi, reverse=True)
        q_score = sum(st["quality"]) / len(st["quality"]) if st["quality"] else 80.0
        s_score = sum(st["service"]) / len(st["service"]) if st["service"] else 80.0
        cert_score = (sum(st["cert"]) / len(st["cert"]) * 100.0) if st["cert"] else 100.0

        total = (
            w.get("price", 0) * price_score
            + w.get("lead_time", 0) * lt_score
            + w.get("quality", 0) * q_score
            + w.get("service", 0) * s_score
            + w.get("cert", 0) * cert_score
        )
        if cert_veto and cert_score < 100.0:
            total *= 0.85          # 资质不全显著扣分（但不直接归零，便于横向看差距）

        sc = SupplierScore(
            supplier=s,
            price_score=round(price_score, 2),
            lead_time_score=round(lt_score, 2),
            quality_score=round(q_score, 2),
            service_score=round(s_score, 2),
            cert_score=round(cert_score, 2),
            total_score=round(total, 2),
            quote_count=st["quotes"],
            win_count=st["wins"],
            cover_amount=round(st["amount"], 2),
            avg_lead_time=round(lt_avg, 1),
            avg_saving_rate=(round(sum(st["saving"]) / len(st["saving"]), 4)
                             if st["saving"] else None),
        )
        scores.append(sc)

    scores.sort(key=lambda x: x.total_score, reverse=True)
    _assign_share(scores)
    _assign_level(scores)
    return scores


def _assign_share(scores: List[SupplierScore]) -> None:
    """按排名给出建议份额：主供 + 辅供 + 备选。"""
    n = len(scores)
    if n == 0:
        return
    if n == 1:
        scores[0].suggested_share = 100.0
        return
    if n == 2:
        top, second = scores[0], scores[1]
        gap = top.total_score - second.total_score
        if gap >= 10:
            top.suggested_share, second.suggested_share = 70.0, 30.0
        elif gap >= 4:
            top.suggested_share, second.suggested_share = 60.0, 40.0
        else:
            top.suggested_share, second.suggested_share = 50.0, 50.0
        return
    scores[0].suggested_share = 60.0
    scores[1].suggested_share = 30.0
    scores[2].suggested_share = 10.0
    for s in scores[3:]:
        s.suggested_share = 0.0


def _assign_level(scores: List[SupplierScore]) -> None:
    for i, s in enumerate(scores):
        if s.total_score >= 85:
            s.level = "A级-战略供应商"
        elif s.total_score >= 75:
            s.level = "B级-合格供应商"
        elif s.total_score >= 60:
            s.level = "C级-观察/整改"
        else:
            s.level = "D级-建议淘汰"
        if i == 0:
            s.level += "（建议主供）"


def to_rows(scores: Sequence[SupplierScore]) -> List[Dict]:
    rows = []
    for i, s in enumerate(scores, 1):
        rows.append({
            "排名": i,
            "供应商": s.supplier,
            "综合得分": s.total_score,
            "价格分": s.price_score,
            "交期分": s.lead_time_score,
            "质量分": s.quality_score,
            "服务分": s.service_score,
            "资质分": s.cert_score,
            "报价条目数": s.quote_count,
            "低价命中数": s.win_count,
            "平均交期(天)": s.avg_lead_time,
            "覆盖金额(不含税)": s.cover_amount,
            "平均降本率": s.avg_saving_rate,
            "建议份额(%)": s.suggested_share,
            "等级": s.level,
        })
    return rows
