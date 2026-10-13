# -*- coding: utf-8 -*-
"""
冒烟测试：不依赖 pytest，直接 `python tests/smoke_test.py` 即可运行。

覆盖点：
 1. 税率拉平：不同税率的报价必须换算到同一口径再比价
 2. MOQ 陷阱：起订量高于需求量时，实际下单量与总成本要跟着变
 3. 运费影响：单价低但运费高的一方，到货单价未必更低
 4. 合规优先：资质不全的低价报价不参与最低价评选
 5. ABC 分类：金额占比符合帕累托规律
 6. EOQ / 安全库存 / 再订货点：随参数单调变化
"""
from __future__ import annotations

import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from mro.models import Quote, InventoryItem, DEFAULT_WEIGHTS  # noqa: E402
from mro.compare import compare_items, build_matrix, summary  # noqa: E402
from mro.scoring import score_suppliers  # noqa: E402
from mro.inventory import analyze  # noqa: E402
from mro.loader import load_quotes, load_inventory  # noqa: E402


def q(**kw) -> Quote:
    base = dict(item_code="T1", item_name="测试物料", unit="EA", qty=100,
                supplier="S", unit_price=100.0, tax_rate=0.13, tax_included=True)
    base.update(kw)
    return Quote(**base)


class TestTaxNormalization(unittest.TestCase):
    """13% 专票 vs 3% 小规模：含税价相同时，不含税成本不同。"""

    def test_same_gross_price_different_tax(self):
        a = q(supplier="一般纳税人", unit_price=113.0, tax_rate=0.13, tax_included=True)
        b = q(supplier="小规模", unit_price=113.0, tax_rate=0.03, tax_included=True)
        self.assertAlmostEqual(a.net_unit_price, 100.0, places=4)
        self.assertAlmostEqual(b.net_unit_price, 109.7087, places=4)
        self.assertLess(a.landed_unit_price, b.landed_unit_price)

    def test_net_price_input(self):
        a = q(unit_price=100.0, tax_rate=0.13, tax_included=False)
        self.assertAlmostEqual(a.net_unit_price, 100.0, places=6)
        self.assertAlmostEqual(a.gross_unit_price, 113.0, places=6)


class TestLandingCost(unittest.TestCase):

    def test_moq_changes_order_qty(self):
        a = q(supplier="A", unit_price=10.0, qty=100, moq=0)
        b = q(supplier="B", unit_price=9.5, qty=100, moq=200)
        self.assertEqual(a.order_qty, 100)
        self.assertEqual(b.order_qty, 200)          # 起订量把下单量拉上去
        self.assertGreater(b.landed_cost, a.landed_cost)   # 单价低但总量翻倍

    def test_freight_flips_winner(self):
        a = q(supplier="A", unit_price=10.0, freight=0.0)
        b = q(supplier="B", unit_price=9.6, freight=600.0)
        self.assertLess(b.net_unit_price, a.net_unit_price)          # 单价 B 更低
        self.assertGreater(b.landed_unit_price, a.landed_unit_price)  # 到货价 A 更低

    def test_saving_rate(self):
        a = q(unit_price=100.0, tax_rate=0.0, history_price=120.0)
        self.assertAlmostEqual(a.landed_unit_price, 100.0, places=6)
        self.assertAlmostEqual(a.saving_rate, (120 - 100) / 120, places=4)


class TestCompare(unittest.TestCase):

    def setUp(self):
        self.quotes = [
            q(supplier="甲", unit_price=110.0, lead_time=5, history_price=100.0),
            q(supplier="乙", unit_price=100.0, tax_rate=0.03, tax_included=False,
              lead_time=15, history_price=100.0),
            q(supplier="丙", unit_price=90.0, lead_time=20, cert_ok=False, history_price=100.0),
        ]
        self.comps = compare_items(self.quotes, require_cert=True)

    def test_cert_first_excludes_low_price(self):
        c = self.comps[0]
        self.assertEqual(len(c.quotes), 3)
        self.assertEqual(c.best.supplier, "甲")     # 丙虽价低但资质不全
        self.assertEqual(c.shadow_best.supplier, "丙")
        self.assertTrue(any("资质不全" in f for f in c.risk_flags()))

    def test_ignore_cert(self):
        c = compare_items(self.quotes, require_cert=False)[0]
        self.assertEqual(c.best.supplier, "丙")

    def test_matrix_shape(self):
        m = build_matrix(self.comps)
        self.assertEqual(len(m["rows"]), 1)
        self.assertEqual(m["suppliers"], ["丙", "乙", "甲"])   # 按拼音/字典序

    def test_summary_saving(self):
        s = summary(self.comps)
        self.assertGreater(s["比价节约金额"], 0)
        self.assertGreater(s["历史降本率"], 0)


class TestScoring(unittest.TestCase):

    def test_price_score_capped_at_100(self):
        quotes = [
            q(supplier="甲", unit_price=113.0, lead_time=5),
            q(supplier="乙", unit_price=90.0, tax_rate=0.0, lead_time=20, cert_ok=False),
        ]
        comps = compare_items(quotes, require_cert=True)
        scores = {s.supplier: s for s in score_suppliers(comps, DEFAULT_WEIGHTS)}
        self.assertLessEqual(scores["乙"].price_score, 100.0)
        self.assertLess(scores["乙"].total_score, scores["甲"].total_score)

    def test_share_sums_to_100(self):
        quotes = [q(supplier=f"S{i}", unit_price=100 + i, lead_time=5 + i) for i in range(4)]
        comps = compare_items(quotes)
        scores = score_suppliers(comps, DEFAULT_WEIGHTS)
        self.assertAlmostEqual(sum(s.suggested_share for s in scores), 100.0, places=4)

    def test_weights_normalized(self):
        quotes = [q(supplier=f"S{i}", unit_price=100 + i) for i in range(3)]
        comps = compare_items(quotes)
        s1 = score_suppliers(comps, {"price": 1, "lead_time": 1, "quality": 0,
                                     "service": 0, "cert": 0})
        s2 = score_suppliers(comps, {"price": 100, "lead_time": 100, "quality": 0,
                                     "service": 0, "cert": 0})
        self.assertAlmostEqual(s1[0].total_score, s2[0].total_score, places=2)


class TestInventory(unittest.TestCase):

    def setUp(self):
        self.items = [
            InventoryItem(item_code=f"I{i}", unit_price=p, annual_qty=n,
                          order_cost=120, holding_rate=0.25, lead_time=14,
                          cv=0.35, service_level=0.95)
            for i, (p, n) in enumerate([(1000, 100), (50, 2000), (5, 20000), (1, 50000)])
        ]

    def test_abc_pareto(self):
        res = analyze(self.items)
        classes = {r["物料编码"]: r["ABC分类"] for r in res["rows"]}
        self.assertIn("A", classes.values())
        self.assertIn("C", classes.values())
        total = sum(r["年耗用金额"] for r in res["rows"])
        a_val = sum(r["年耗用金额"] for r in res["rows"] if r["ABC分类"] == "A")
        self.assertLessEqual(a_val / total, 0.95)

    def test_eoq_monotonic(self):
        a = InventoryItem(unit_price=100, annual_qty=1000, order_cost=120, holding_rate=0.25)
        b = InventoryItem(unit_price=100, annual_qty=4000, order_cost=120, holding_rate=0.25)
        self.assertAlmostEqual(b.eoq, a.eoq * 2, places=4)   # EOQ ∝ √D

    def test_safety_stock_grows_with_lead_time(self):
        a = InventoryItem(unit_price=10, annual_qty=3650, lead_time=7, cv=0.3)
        b = InventoryItem(unit_price=10, annual_qty=3650, lead_time=28, cv=0.3)
        self.assertAlmostEqual(b.safety_stock, a.safety_stock * 2, places=4)
        self.assertGreater(b.reorder_point, a.reorder_point)


class TestLoader(unittest.TestCase):

    def test_csv_roundtrip(self):
        import csv
        header = ["物料编码", "物料名称", "单位", "需求数量", "供应商", "单价", "税率",
                  "是否含税", "交期", "历史单价"]
        rows = [
            ["T1", "测试物料", "EA", 100, "甲", 113, 0.13, "是", 5, 100],
            ["T1", "测试物料", "EA", 100, "乙", 103, 0.03, "否", 15, 100],
        ]
        with tempfile.NamedTemporaryFile("w", suffix=".csv", delete=False,
                                        encoding="utf-8", newline="") as f:
            w = csv.writer(f); w.writerow(header); w.writerows(rows)
            path = f.name
        try:
            quotes = load_quotes(path)
            self.assertEqual(len(quotes), 2)
            self.assertEqual(quotes[0].supplier, "甲")
            self.assertTrue(quotes[0].tax_included)
            self.assertFalse(quotes[1].tax_included)
        finally:
            os.unlink(path)

    def test_zero_value_fields_are_kept(self):
        """回归测试：0 是合法值（免税税率、现货交期），不能被默认值覆盖。"""
        import csv
        header = ["物料编码", "物料名称", "单位", "需求数量", "供应商", "单价", "税率",
                  "是否含税", "交期", "最小起订量", "运费"]
        rows = [
            ["T1", "免税物料", "EA", 10, "甲", 100, 0, "是", 0, 0, 0],
        ]
        with tempfile.NamedTemporaryFile("w", suffix=".csv", delete=False,
                                        encoding="utf-8", newline="") as f:
            w = csv.writer(f); w.writerow(header); w.writerows(rows)
            path = f.name
        try:
            q1 = load_quotes(path)[0]
            self.assertEqual(q1.tax_rate, 0.0, "免税(0%)被错误改成了默认税率")
            self.assertEqual(q1.lead_time, 0.0, "现货(0天)被错误改成了默认交期")
            self.assertEqual(q1.moq, 0.0)
            self.assertEqual(q1.freight, 0.0)
            self.assertAlmostEqual(q1.net_unit_price, 100.0, places=6)
        finally:
            os.unlink(path)


if __name__ == "__main__":
    unittest.main(verbosity=2)
