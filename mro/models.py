# -*- coding: utf-8 -*-
"""
数据模型与领域常量。

设计原则：
1. 零第三方依赖，只用 Python 标准库，便于在任何环境 clone 即跑；
2. 字段命名同时保留中英文，Excel 表头两种写法都能识别；
3. 所有金额计算显式区分「含税 / 不含税 / 到货总成本」，
   这是采购比价中最容易出错、也是最能体现专业度的地方。
"""
from __future__ import annotations

from dataclasses import dataclass, field, asdict
from typing import Any, Dict, Iterable, List, Optional

# --------------------------------------------------------------------------
# 字段别名：Excel / CSV 表头可能写中文也可能写英文，统一映射
# --------------------------------------------------------------------------
FIELD_ALIASES: Dict[str, tuple] = {
    "item_code":     ("物料编码", "物料编码/货号", "编码", "item_code", "material", "sku"),
    "item_name":     ("物料名称", "品名", "名称", "item_name", "material_name", "name"),
    "category":      ("物料类别", "品类", "分类", "category"),
    "spec":          ("规格型号", "规格", "型号", "spec", "model"),
    "unit":          ("单位", "unit", "uom"),
    "qty":           ("需求数量", "数量", "请购数量", "qty", "quantity"),
    "supplier":      ("供应商", "供应商名称", "报价单位", "supplier", "vendor"),
    "unit_price":    ("单价", "报价单价", "含税单价", "unit_price", "price"),
    "tax_rate":      ("税率", "tax_rate", "vat"),
    "tax_included":  ("是否含税", "含税", "tax_included"),
    "lead_time":     ("交期", "交货期", "交期(天)", "交期天数", "lead_time"),
    "moq":           ("最小起订量", "起订量", "moq"),
    "freight":       ("运费", "运费(元)", "freight", "shipping"),
    "payment_terms": ("付款条件", "付款方式", "账期", "payment_terms"),
    "quality_score": ("质量评分", "质量得分", "合格评分", "quality_score"),
    "service_score": ("服务评分", "服务得分", "响应评分", "service_score"),
    "cert_ok":       ("资质合规", "资质齐全", "证照齐全", "cert_ok", "certified"),
    "history_price": ("历史单价", "历史单价(不含税)", "目标价", "上期单价", "去年同期价",
                      "history_price", "last_price", "target_price"),
    "remark":        ("备注", "remark", "comment"),
}

# 置信水平 Z 值（安全库存计算用）
Z_SCORES: Dict[float, float] = {
    0.50: 0.000,
    0.80: 0.842,
    0.85: 1.036,
    0.90: 1.282,
    0.95: 1.645,
    0.97: 1.881,
    0.98: 2.054,
    0.99: 2.326,
    0.995: 2.576,
}

# 默认供应商评分权重（可在 config.json 中覆盖）
DEFAULT_WEIGHTS: Dict[str, float] = {
    "price": 0.35,     # 价格竞争力
    "lead_time": 0.20, # 交期
    "quality": 0.20,   # 质量 / 合格率
    "service": 0.15,   # 服务与响应
    "cert": 0.10,      # 资质与合规
}

WEIGHT_LABELS = {
    "price": "价格竞争力",
    "lead_time": "交期表现",
    "quality": "质量水平",
    "service": "服务响应",
    "cert": "资质合规",
}


def _to_float(v: Any, default: float = 0.0) -> float:
    if v is None or v == "":
        return default
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(",", "").replace("¥", "").replace("元", "")
    s = s.replace("%", "")
    try:
        return float(s)
    except ValueError:
        return default


def _to_int(v: Any, default: int = 0) -> int:
    return int(round(_to_float(v, default)))


def _to_bool(v: Any, default: bool = True) -> bool:
    """把 'Y'/'是'/'1'/'TRUE' 之类统一转成 bool。"""
    if v is None or v == "":
        return default
    if isinstance(v, bool):
        return v
    s = str(v).strip().upper()
    if s in ("Y", "YES", "TRUE", "1", "是", "有", "齐全", "√", "V", "OK"):
        return True
    if s in ("N", "NO", "FALSE", "0", "否", "无", "缺失", "×", "X"):
        return False
    return default


def _to_str(v: Any, default: str = "") -> str:
    if v is None:
        return default
    return str(v).strip()


@dataclass
class Quote:
    """一条供应商报价记录。"""
    item_code: str = ""
    item_name: str = ""
    category: str = ""
    spec: str = ""
    unit: str = ""
    qty: float = 0.0
    supplier: str = ""
    unit_price: float = 0.0      # 报价单价（是否含税看 tax_included）
    tax_rate: float = 0.13       # 默认 13%
    tax_included: bool = True
    lead_time: float = 0.0       # 天
    moq: float = 0.0             # 最小起订量
    freight: float = 0.0         # 运费（含税）
    payment_terms: str = ""
    quality_score: float = 80.0  # 缺省给基准分
    service_score: float = 80.0
    cert_ok: bool = True
    history_price: float = 0.0   # 历史价 / 目标价，用于算降本率
    remark: str = ""

    # ---------------- 派生指标 ----------------
    @property
    def net_unit_price(self) -> float:
        """不含税单价。"""
        if self.tax_included:
            return self.unit_price / (1 + self.tax_rate) if self.tax_rate > -1 else self.unit_price
        return self.unit_price

    @property
    def gross_unit_price(self) -> float:
        """含税单价。"""
        if self.tax_included:
            return self.unit_price
        return self.unit_price * (1 + self.tax_rate)

    @property
    def order_qty(self) -> float:
        """实际下单量：需求数量与起订量取大。"""
        return max(self.qty, self.moq, 0.0)

    @property
    def goods_amount(self) -> float:
        """货款金额（不含税）。"""
        return self.net_unit_price * self.order_qty

    @property
    def tax_amount(self) -> float:
        """税额。"""
        return self.gross_unit_price * self.order_qty - self.goods_amount

    @property
    def freight_net(self) -> float:
        """运费（折算不含税）。"""
        return self.freight / (1 + self.tax_rate) if self.tax_rate > -1 else self.freight

    @property
    def landed_cost(self) -> float:
        """到货总成本（不含税货款 + 不含税运费）。比价的核心口径。"""
        return self.goods_amount + self.freight_net

    @property
    def landed_unit_price(self) -> float:
        """到货等效单价（含运费、不含税）。跨供应商横向对比就用这个。"""
        if self.order_qty <= 0:
            return 0.0
        return self.landed_cost / self.order_qty

    @property
    def landed_unit_price_gross(self) -> float:
        """到货等效单价（含税口径，用于对账与预算）。"""
        return self.landed_unit_price * (1 + self.tax_rate)

    @property
    def saving_rate(self) -> Optional[float]:
        """相对历史价/目标价的降本率。无历史价时返回 None。"""
        if self.history_price <= 0:
            return None
        return (self.history_price - self.landed_unit_price) / self.history_price

    def to_row(self) -> Dict[str, Any]:
        d = asdict(self)
        d.update({
            "不含税单价": round(self.net_unit_price, 4),
            "含税单价": round(self.gross_unit_price, 4),
            "实际下单量": self.order_qty,
            "货款金额(不含税)": round(self.goods_amount, 2),
            "运费(不含税)": round(self.freight_net, 2),
            "到货总成本(不含税)": round(self.landed_cost, 2),
            "到货等效单价(不含税)": round(self.landed_unit_price, 4),
            "到货等效单价(含税)": round(self.landed_unit_price_gross, 4),
        })
        return d


@dataclass
class ItemComparison:
    """单个物料的多供应商比价结果。"""
    item_code: str
    item_name: str
    spec: str
    unit: str
    qty: float
    quotes: List[Quote] = field(default_factory=list)
    ranked: List[Quote] = field(default_factory=list)   # 按到货等效单价升序
    best: Optional[Quote] = None
    second: Optional[Quote] = None
    worst: Optional[Quote] = None
    shadow_best: Optional[Quote] = None   # 忽略资质的最低价，仅用于风险提示

    @property
    def supplier_count(self) -> int:
        return len({q.supplier for q in self.quotes if q.supplier})

    @property
    def spread_rate(self) -> float:
        """最高价相对最低价的溢价率，看议价空间。"""
        if not self.best or not self.worst or self.best.landed_unit_price <= 0:
            return 0.0
        return (self.worst.landed_unit_price - self.best.landed_unit_price) / self.best.landed_unit_price

    @property
    def gap_to_second(self) -> float:
        """次低价相对最低价的溢价率。<2% 说明两家基本打平，可分份额。"""
        if not self.best or not self.second or self.best.landed_unit_price <= 0:
            return 0.0
        return (self.second.landed_unit_price - self.best.landed_unit_price) / self.best.landed_unit_price

    @property
    def best_amount(self) -> float:
        return self.best.landed_cost if self.best else 0.0

    @property
    def saving_vs_worst(self) -> float:
        """选最低价相比选最高价节省的金额。"""
        if not self.best or not self.worst:
            return 0.0
        return self.worst.landed_cost - self.best.landed_cost

    @property
    def saving_vs_history(self) -> float:
        """相对历史价的节省金额（按最低价口径）。"""
        if not self.best or self.best.history_price <= 0:
            return 0.0
        return (self.best.history_price - self.best.landed_unit_price) * self.best.order_qty

    @property
    def single_source(self) -> bool:
        return self.supplier_count <= 1

    def risk_flags(self) -> List[str]:
        """风险提示：采购决策里最该被看见的东西。"""
        flags: List[str] = []
        if self.supplier_count == 0:
            flags.append("无有效报价")
            return flags
        if self.supplier_count == 1:
            flags.append("单一来源：建议开发备选供应商")
        if self.supplier_count == 2 and self.gap_to_second < 0.02:
            flags.append("两家报价接近：可按交期/服务分配份额")
        if self.spread_rate > 0.30:
            flags.append("价差>30%：需复核规格与品牌是否一致")

        # 被合规卡掉的低价（影子最低价）
        if self.shadow_best and self.best and self.shadow_best is not self.best:
            if not self.shadow_best.cert_ok:
                gap = (self.best.landed_unit_price - self.shadow_best.landed_unit_price) \
                      / self.shadow_best.landed_unit_price
                flags.append(
                    f"低价方【{self.shadow_best.supplier}】资质不全已被排除，"
                    f"合规优先多付 {gap:.1%}；可推动其补件后重新议价"
                )
        if self.best and not self.best.cert_ok:
            flags.append("最低价供应商资质不全：合规优先，需排除或补充材料")
        if self.best and self.second and self.best.lead_time > self.second.lead_time * 1.5:
            flags.append("最低价交期明显偏长：需与需求部门确认到货时间")
        if self.best and self.qty > 0 and self.best.moq > self.qty * 1.2:
            flags.append("起订量高于需求量：存在呆滞库存风险，可谈分批交付")

        # 付款条件风险（现金流视角）
        worst_pay = [q for q in self.quotes if q.payment_terms in ("预付全款", "款到发货")]
        if worst_pay and self.best and self.best.payment_terms not in ("预付全款", "款到发货"):
            flags.append("存在要求预付的报价：优先选择账期结算以保现金流")
        return flags

    def recommendation(self) -> str:
        if not self.best:
            return "无有效报价，需重新询价"
        if not self.best.cert_ok and self.second:
            return f"建议授予【{self.second.supplier}】：最低价方资质不全"
        if (self.shadow_best and self.shadow_best is not self.best
                and not self.shadow_best.cert_ok
                and self.shadow_best.landed_unit_price < self.best.landed_unit_price):
            return (f"建议授予【{self.best.supplier}】；"
                    f"同步要求【{self.shadow_best.supplier}】补齐资质后重新议价")
        if self.supplier_count == 1:
            return f"暂授【{self.best.supplier}】并同步开发第二货源"
        if self.second and self.gap_to_second < 0.02:
            return f"建议双源：【{self.best.supplier}】主供 / 【{self.second.supplier}】辅供"
        return f"建议授予【{self.best.supplier}】"


@dataclass
class SupplierScore:
    """供应商综合评分结果。"""
    supplier: str
    price_score: float = 0.0
    lead_time_score: float = 0.0
    quality_score: float = 0.0
    service_score: float = 0.0
    cert_score: float = 0.0
    total_score: float = 0.0
    quote_count: int = 0
    win_count: int = 0
    cover_amount: float = 0.0     # 覆盖金额（该供应商报价涉及的物料总额）
    avg_lead_time: float = 0.0
    avg_saving_rate: Optional[float] = None
    suggested_share: float = 0.0
    level: str = ""

    @property
    def dimensions(self) -> Dict[str, float]:
        return {
            "price": self.price_score,
            "lead_time": self.lead_time_score,
            "quality": self.quality_score,
            "service": self.service_score,
            "cert": self.cert_score,
        }


@dataclass
class InventoryItem:
    """库存 / 补货分析对象。"""
    item_code: str = ""
    item_name: str = ""
    category: str = ""
    unit: str = ""
    unit_price: float = 0.0
    annual_qty: float = 0.0          # 年耗用量
    order_cost: float = 120.0        # 每次下单的订货成本（询价+审批+收货+付款）
    holding_rate: float = 0.25       # 年持有成本率（资金+仓储+损耗+呆滞）
    lead_time: float = 14.0          # 采购提前期（天）
    demand_std: float = 0.0          # 日耗用标准差，缺省用 cv 推算
    cv: float = 0.35                 # 变异系数（MRO 需求波动大，默认偏高）
    service_level: float = 0.95
    current_stock: float = 0.0

    @property
    def annual_value(self) -> float:
        return self.unit_price * self.annual_qty

    @property
    def daily_demand(self) -> float:
        return self.annual_qty / 365.0 if self.annual_qty else 0.0

    @property
    def sigma_d(self) -> float:
        """日需求标准差：给了实测值就用实测，否则用 cv × 日均。"""
        if self.demand_std > 0:
            return self.demand_std
        return self.cv * self.daily_demand

    @property
    def eoq(self) -> float:
        """经济订货批量 sqrt(2DS/H)。"""
        h = self.unit_price * self.holding_rate
        if h <= 0 or self.annual_qty <= 0:
            return 0.0
        return (2 * self.annual_qty * self.order_cost / h) ** 0.5

    @property
    def safety_stock(self) -> float:
        """安全库存 SS = Z × σ_d × √LT。"""
        z = Z_SCORES.get(round(self.service_level, 3), 1.645)
        return z * self.sigma_d * (self.lead_time ** 0.5)

    @property
    def reorder_point(self) -> float:
        """再订货点 ROP = 日均需求 × 提前期 + 安全库存。"""
        return self.daily_demand * self.lead_time + self.safety_stock

    @property
    def order_times(self) -> float:
        """年订货次数。"""
        return self.annual_qty / self.eoq if self.eoq > 0 else 0.0

    @property
    def abc_class(self) -> str:
        return getattr(self, "_abc", "C")

    @property
    def stock_days(self) -> float:
        """现有库存可支撑天数。"""
        return self.current_stock / self.daily_demand if self.daily_demand > 0 else 0.0

    @property
    def need_reorder(self) -> bool:
        return self.current_stock <= self.reorder_point

    def policy(self) -> str:
        """按 ABC 分类给出补货策略——MRO 长尾物料降本的关键。"""
        c = self.abc_class
        if c == "A":
            return "A类：重点管控，逐单比价 + 签年度框架协议，设安全库存并月度复核"
        if c == "B":
            return "B类：常规ROP补货，季度比价，可合并同类项打包询价"
        return "C类：低值长尾，采用双箱法/定期补货，年度打包招标，减少下单频次降管理成本"
