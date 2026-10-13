# -*- coding: utf-8 -*-
"""
生成示例数据（sample_quotes.xlsx / sample_inventory.xlsx）。

场景设定：某医药化工企业的 MRO 与间接物料季度询价，
覆盖阀门管件、泵阀备件、电气、实验耗材、劳保、化工辅料等品类，
并刻意混入以下「真实采购里一定会遇到」的坑，用来验证比价口径是否正确：
  · 不同税率（13% / 3% 小规模 / 免税）
  · 含税价与不含税价混报
  · 运费另计 vs 包邮
  · 起订量高于需求量（MOQ 陷阱）
  · 资质不全的低价供应商（合规优先）
  · 交期长短差异
"""
from __future__ import annotations

import os
import sys

try:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit("需要 openpyxl： pip install openpyxl")

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(os.path.dirname(HERE), "data")
os.makedirs(DATA, exist_ok=True)

# ---------------------------------------------------------------- 报价数据
HEADER = ["物料编码", "物料名称", "物料类别", "规格型号", "单位", "需求数量",
          "供应商", "单价", "税率", "是否含税", "交期", "最小起订量", "运费",
          "付款条件", "质量评分", "服务评分", "资质合规", "历史单价", "备注"]

# (编码, 品名, 类别, 规格, 单位, 需求数量, 历史不含税单价)
# 注意：历史单价统一用「不含税」口径，与比价主口径一致，
#      否则含税/不含税混比会系统性虚增降本率。
ITEMS = [
    ("MRO-1001", "不锈钢球阀", "阀门管件", "DN50 304 法兰式", "只", 40, 237.2),
    ("MRO-1002", "离心泵机械密封", "泵阀备件", "IH50-32-160 碳化硅", "套", 12, 787.6),
    ("MRO-1003", "深沟球轴承", "传动备件", "6208-2RS", "个", 60, 28.8),
    ("MRO-1004", "洁净室丁腈手套", "劳保用品", "M 号 无粉 100只/盒", "盒", 300, 40.7),
    ("MRO-1005", "高效空气过滤器", "暖通耗材", "HEPA H13 610×610", "片", 24, 1044.2),
    ("MRO-1006", "电力电缆", "电气材料", "YJV-0.6/1kV 3×4", "米", 800, 10.44),
    ("MRO-1007", "无水乙醇", "实验试剂", "AR 500ml", "瓶", 120, 16.37),
    ("MRO-1008", "工业乙二醇", "化工辅料", "工业级 220kg/桶", "kg", 2200, 6.11),
    ("MRO-1009", "不锈钢压力表", "仪器仪表", "0-1.6MPa 径向", "只", 30, 85.0),
    ("MRO-1010", "抗磨液压油", "润滑油脂", "46# 200L/桶", "L", 600, 11.68),
    ("MRO-1011", "PP熔喷滤芯", "过滤耗材", "10寸 5μm", "支", 200, 7.61),
    ("MRO-1012", "外六角螺栓", "紧固件", "M12×50 304", "套", 1500, 2.08),
]

# 四家虚构供应商，定位各不相同（交期 / 税率 / 账期 / 资质 全面拉开差异）
SUPPLIERS = {
    # 综合型：交期快、资质全、服务好，价格中等
    "华创工业物资": dict(lt=7,  moq_rate=0.0,  freight=0.0,   tax=0.13, incl=True,
                     q=92, s=88, cert=True,  pay="月结30天"),
    # 化工品专精：小规模纳税人（3% 专票）、账期长，但交期慢、有起订量
    "瑞科供应链":   dict(lt=15, moq_rate=0.5,  freight=180.0, tax=0.03, incl=False,
                     q=88, s=90, cert=True,  pay="月结60天"),
    # 机电备件强：交期最快、质量最好，价格偏高
    "广联机电":     dict(lt=5,  moq_rate=0.0,  freight=0.0,   tax=0.13, incl=True,
                     q=95, s=82, cert=True,  pay="货到付款"),
    # 低值品价格屠夫：免税收据、要求预付，资质不全（合规风险）
    "通达五金":     dict(lt=20, moq_rate=1.0,  freight=260.0, tax=0.00, incl=True,
                     q=76, s=70, cert=False, pay="预付全款"),
}

# 品类价格因子矩阵：每家在不同品类各有强弱，制造真实的价格交叉
# 值 = 含税报价 / 历史价（越低越便宜）
PRICE_FACTOR = {
    #                 华创   瑞科   广联   通达
    "阀门管件":      {"华创工业物资": 0.93, "瑞科供应链": 1.05, "广联机电": 1.02, "通达五金": 0.95},
    "泵阀备件":      {"华创工业物资": 1.04, "瑞科供应链": 1.08, "广联机电": 0.90, "通达五金": 0.97},
    "传动备件":      {"华创工业物资": 0.92, "瑞科供应链": 1.10, "广联机电": 0.98, "通达五金": 0.88},
    "劳保用品":      {"华创工业物资": 1.00, "瑞科供应链": 1.02, "广联机电": 1.09, "通达五金": 0.84},
    "暖通耗材":      {"华创工业物资": 1.06, "瑞科供应链": 1.03, "广联机电": 0.88, "通达五金": 0.99},
    "电气材料":      {"华创工业物资": 1.02, "瑞科供应链": 1.06, "广联机电": 0.91, "通达五金": 0.96},
    "实验试剂":      {"华创工业物资": 1.08, "瑞科供应链": 0.89, "广联机电": 1.12, "通达五金": 0.92},
    "化工辅料":      {"华创工业物资": 1.05, "瑞科供应链": 0.86, "广联机电": 1.15, "通达五金": 0.90},
    "仪器仪表":      {"华创工业物资": 0.94, "瑞科供应链": 1.07, "广联机电": 1.00, "通达五金": 0.93},
    "润滑油脂":      {"华创工业物资": 1.03, "瑞科供应链": 0.88, "广联机电": 1.10, "通达五金": 0.94},
    "过滤耗材":      {"华创工业物资": 1.01, "瑞科供应链": 0.99, "广联机电": 1.05, "通达五金": 0.82},
    "紧固件":        {"华创工业物资": 1.07, "瑞科供应链": 1.01, "广联机电": 1.11, "通达五金": 0.80},
}


def build_rows() -> list:
    rows = []
    for idx, (code, name, cat, spec, unit, qty, hist) in enumerate(ITEMS):
        factors = PRICE_FACTOR.get(cat, {s: 1.0 for s in SUPPLIERS})
        for sup, cfg in SUPPLIERS.items():
            # 先确定该家在「含税口径」上的报价水平，再折算到它的开票口径：
            #   含税报价 → 直接写；不含税报价 → 除以 (1+税率)
            # 这样同一含税价下，3% 小规模纳税人的不含税成本反而更高，
            # 因为可抵扣进项少——这正是比价时必须拉平税率的现实原因。
            target_gross = hist * factors.get(sup, 1.0)
            price = target_gross if cfg["incl"] else target_gross / (1 + cfg["tax"])
            price = round(price, 2)

            moq = round(qty * cfg["moq_rate"] / 10) * 10
            amount = qty * target_gross
            rows.append([
                code, name, cat, spec, unit, qty,
                sup,
                price,
                cfg["tax"],
                "是" if cfg["incl"] else "否",
                cfg["lt"],
                moq,
                cfg["freight"] if amount < 20000 else 0.0,   # 大额订单免运费
                cfg["pay"],
                cfg["q"] + (idx % 3),          # 质量分略作波动
                cfg["s"] - (idx % 4),          # 服务分略作波动
                "是" if cfg["cert"] else "否",
                hist,
                "" if cfg["cert"] else "营业执照与危化品经营资质待补充",
            ])
    return rows


def write_quotes(path: str):
    wb = Workbook()
    ws = wb.active
    ws.title = "报价明细"
    ws.append(HEADER)
    for r in build_rows():
        ws.append(r)
    _beautify(ws, len(HEADER))
    wb.save(path)
    return path


# ---------------------------------------------------------------- 库存数据
INV_HEADER = ["物料编码", "物料名称", "物料类别", "单位", "单价", "年耗用量",
              "单次订货成本", "持有成本率", "采购提前期", "日需求标准差",
              "变异系数", "服务水平", "当前库存"]

INV_ROWS = [
    ["MRO-1001", "不锈钢球阀", "阀门管件", "只", 268.0, 160, 120, 0.25, 15, "", 0.35, 0.95, 20],
    ["MRO-1002", "离心泵机械密封", "泵阀备件", "套", 890.0, 48, 120, 0.25, 25, "", 0.50, 0.97, 6],
    ["MRO-1003", "深沟球轴承", "传动备件", "个", 32.5, 240, 120, 0.25, 10, "", 0.30, 0.95, 45],
    ["MRO-1004", "洁净室丁腈手套", "劳保用品", "盒", 46.0, 1200, 120, 0.25, 7, "", 0.20, 0.95, 180],
    ["MRO-1005", "高效空气过滤器", "暖通耗材", "片", 1180.0, 96, 120, 0.25, 30, "", 0.45, 0.98, 8],
    ["MRO-1006", "电力电缆", "电气材料", "米", 11.8, 3200, 120, 0.25, 12, "", 0.40, 0.95, 400],
    ["MRO-1007", "无水乙醇", "实验试剂", "瓶", 18.5, 480, 120, 0.25, 5, "", 0.25, 0.95, 60],
    ["MRO-1008", "工业乙二醇", "化工辅料", "kg", 6.9, 8800, 120, 0.25, 10, "", 0.30, 0.95, 900],
    ["MRO-1009", "不锈钢压力表", "仪器仪表", "只", 96.0, 120, 120, 0.25, 14, "", 0.35, 0.95, 15],
    ["MRO-1010", "抗磨液压油", "润滑油脂", "L", 13.2, 2400, 120, 0.25, 8, "", 0.25, 0.95, 300],
    ["MRO-1011", "PP熔喷滤芯", "过滤耗材", "支", 8.6, 800, 120, 0.25, 7, "", 0.30, 0.95, 60],
    ["MRO-1012", "外六角螺栓", "紧固件", "套", 2.35, 6000, 120, 0.25, 10, "", 0.35, 0.95, 500],
]


def write_inventory(path: str):
    wb = Workbook()
    ws = wb.active
    ws.title = "库存明细"
    ws.append(INV_HEADER)
    for r in INV_ROWS:
        ws.append(r)
    _beautify(ws, len(INV_HEADER))
    wb.save(path)
    return path


def _beautify(ws, ncol: int):
    fill = PatternFill("solid", fgColor="1F4E79")
    for j in range(1, ncol + 1):
        c = ws.cell(row=1, column=j)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = fill
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(j)].width = max(12, min(22, len(str(c.value)) * 2 + 4))
    ws.freeze_panes = "A2"


if __name__ == "__main__":
    p1 = write_quotes(os.path.join(DATA, "sample_quotes.xlsx"))
    p2 = write_inventory(os.path.join(DATA, "sample_inventory.xlsx"))
    print("生成完成：")
    print(" ", p1)
    print(" ", p2)
