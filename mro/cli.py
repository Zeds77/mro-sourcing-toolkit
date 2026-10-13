# -*- coding: utf-8 -*-
"""
命令行入口：python -m mro.cli <command>

命令：
  init       生成空白模板（quotes.xlsx / inventory.xlsx）
  run        跑完整分析，输出 Excel 报告
  compare    只做比价
  score      只做供应商评分
  inventory  只做库存补货测算
  demo       用内置示例数据一键演示
"""
from __future__ import annotations

import argparse
import datetime as _dt
import os
import sys
from typing import List, Optional

# Windows 控制台中文输出兜底
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

from .models import DEFAULT_WEIGHTS, WEIGHT_LABELS
from .loader import load_quotes, load_inventory, load_weights, template_rows
from .compare import compare_items, build_matrix, summary
from .scoring import score_suppliers, to_rows
from .inventory import analyze as analyze_inventory
from .report import write_report, build_context


def _print_table(rows: List[dict], max_col: int = 26):
    """控制台表格输出（不依赖任何第三方库）。"""
    if not rows:
        print("（无数据）")
        return
    cols: List[str] = []
    for r in rows:
        for k in r.keys():
            if k not in cols:
                cols.append(k)

    def fmt(v) -> str:
        if v is None:
            return ""
        if isinstance(v, float):
            return f"{v:,.4f}" if abs(v) < 1000 else f"{v:,.2f}"
        s = str(v)
        return s if len(s) <= max_col else s[: max_col - 1] + "…"

    widths = [max([len(str(c))] + [len(fmt(r.get(c))) for r in rows]) for c in cols]
    widths = [min(w, max_col) for w in widths]
    line = "  ".join(str(c).ljust(widths[i]) for i, c in enumerate(cols))
    print(line)
    print("-" * min(len(line), 160))
    for r in rows:
        print("  ".join(fmt(r.get(c)).ljust(widths[i]) for i, c in enumerate(cols)))


def _write_template(path: str, kind: str) -> str:
    header, sample = template_rows(kind)
    try:
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "报价明细" if kind == "quote" else "库存明细"
        ws.append(header)
        ws.append(sample)
        os.makedirs(os.path.dirname(os.path.abspath(path)) or ".", exist_ok=True)
        wb.save(path)
        return path
    except ImportError:
        import csv
        path = os.path.splitext(path)[0] + ".csv"
        with open(path, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f)
            w.writerow(header)
            w.writerow(sample)
        return path


def cmd_init(args) -> int:
    out_dir = args.out or "."
    os.makedirs(out_dir, exist_ok=True)
    p1 = _write_template(os.path.join(out_dir, "quotes_template.xlsx"), "quote")
    p2 = _write_template(os.path.join(out_dir, "inventory_template.xlsx"), "inventory")
    print(f"[OK] 已生成模板：\n  {p1}\n  {p2}")
    return 0


def cmd_run(args) -> int:
    quotes = load_quotes(args.quotes)
    if not quotes:
        print("[ERR] 未读到有效报价数据，请检查文件格式与表头。")
        return 2
    weights = load_weights(args.weights)
    comparisons = compare_items(quotes, require_cert=not args.ignore_cert)
    scores = score_suppliers(comparisons, weights, cert_veto=not args.ignore_cert)

    inv_result = None
    if args.inventory and os.path.exists(args.inventory):
        inv_result = analyze_inventory(load_inventory(args.inventory))

    ctx = build_context(
        comparisons, scores, weights, inv_result,
        generated_at=_dt.datetime.now().strftime("生成时间：%Y-%m-%d %H:%M"),
    )
    out = write_report(args.out, ctx)
    print(f"[OK] 报告已生成：{out}\n")
    print("=== 汇总 ===")
    _print_table([{"指标": k, "数值": v} for k, v in ctx["summary"].items()])
    print("\n=== 供应商评分 ===")
    _print_table(ctx["score_rows"])
    if args.top:
        print("\n=== 比价建议（前 %d 项）===" % args.top)
        _print_table(ctx["advice_rows"][: args.top])
    return 0


def cmd_compare(args) -> int:
    quotes = load_quotes(args.quotes)
    comparisons = compare_items(quotes, require_cert=not args.ignore_cert)
    print("=== 比价汇总 ===")
    _print_table([{"指标": k, "数值": v} for k, v in summary(comparisons).items()])
    print("\n=== 比价矩阵 ===")
    m = build_matrix(comparisons)
    _print_table(m["rows"], max_col=18)
    print("\n=== 决策建议 ===")
    rows = [{
        "物料编码": c.item_code, "物料名称": c.item_name,
        "最优供应商": c.best.supplier if c.best else "",
        "最优单价": round(c.best.landed_unit_price, 4) if c.best else None,
        "价差率": round(c.spread_rate, 4),
        "风险": "；".join(c.risk_flags()), "建议": c.recommendation(),
    } for c in comparisons]
    _print_table(rows, max_col=30)
    return 0


def cmd_score(args) -> int:
    quotes = load_quotes(args.quotes)
    weights = load_weights(args.weights)
    comparisons = compare_items(quotes, require_cert=not args.ignore_cert)
    scores = score_suppliers(comparisons, weights, cert_veto=not args.ignore_cert)
    print("=== 评分权重 ===")
    _print_table([{"维度": WEIGHT_LABELS.get(k, k), "权重": round(v, 3)} for k, v in weights.items()])
    print("\n=== 供应商评分 ===")
    _print_table(to_rows(scores), max_col=22)
    return 0


def cmd_inventory(args) -> int:
    items = load_inventory(args.inventory)
    res = analyze_inventory(items)
    print("=== ABC 分类汇总 ===")
    _print_table([{"ABC分类": k, **v} for k, v in res["by_class"].items()])
    print(f"\n全年订货次数：{res['全年订货次数']}")
    print(f"年订货管理成本：{res['年订货管理成本']:,.2f} 元")
    print(f"C类长尾合并下单预计可节约管理费：{res['C类合并下单可节约管理费']:,.2f} 元")
    print("\n=== 明细 ===")
    _print_table(res["rows"], max_col=20)
    return 0


def cmd_demo(args) -> int:
    here = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    data_dir = os.path.join(here, "data")
    q = os.path.join(data_dir, "sample_quotes.xlsx")
    inv = os.path.join(data_dir, "sample_inventory.xlsx")
    if not os.path.exists(q):
        print("[ERR] 缺少示例数据，请先运行： python -m mro.cli init")
        return 2
    args.quotes = q
    args.inventory = inv if os.path.exists(inv) else None
    args.weights = None
    args.ignore_cert = False
    args.top = 10
    args.out = args.out or os.path.join(here, "output", "demo_report.xlsx")
    return cmd_run(args)


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="mro",
        description="MRO 采购比价 / 供应商评分 / 库存补货分析工具",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "示例：\n"
            "  python -m mro.cli init\n"
            "  python -m mro.cli run -q quotes.xlsx -o output/report.xlsx\n"
            "  python -m mro.cli score -q quotes.xlsx\n"
            "  python -m mro.cli inventory -i inventory.xlsx\n"
            "  python -m mro.cli demo\n"
            "  streamlit run app.py\n"
        ),
    )
    sub = p.add_subparsers(dest="cmd")

    s = sub.add_parser("init", help="生成 Excel 模板")
    s.add_argument("-o", "--out", default=".", help="输出目录")
    s.set_defaults(func=cmd_init)

    s = sub.add_parser("run", help="完整分析并输出 Excel 报告")
    s.add_argument("-q", "--quotes", required=True, help="报价明细文件(.xlsx/.csv)")
    s.add_argument("-i", "--inventory", default="", help="库存明细文件（可选）")
    s.add_argument("-w", "--weights", default="", help="评分权重 JSON（可选）")
    s.add_argument("-o", "--out", default="output/report.xlsx", help="报告输出路径")
    s.add_argument("--ignore-cert", action="store_true", help="忽略资质校验（仅演练用）")
    s.add_argument("--top", type=int, default=10, help="控制台打印前 N 条建议")
    s.set_defaults(func=cmd_run)

    s = sub.add_parser("compare", help="仅执行比价")
    s.add_argument("-q", "--quotes", required=True)
    s.add_argument("--ignore-cert", action="store_true")
    s.set_defaults(func=cmd_compare)

    s = sub.add_parser("score", help="仅执行供应商评分")
    s.add_argument("-q", "--quotes", required=True)
    s.add_argument("-w", "--weights", default="")
    s.add_argument("--ignore-cert", action="store_true")
    s.set_defaults(func=cmd_score)

    s = sub.add_parser("inventory", help="库存与补货测算")
    s.add_argument("-i", "--inventory", required=True)
    s.set_defaults(func=cmd_inventory)

    s = sub.add_parser("demo", help="用内置示例数据演示")
    s.add_argument("-o", "--out", default="", help="报告输出路径")
    s.set_defaults(func=cmd_demo)

    return p


def main(argv: Optional[List[str]] = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    if not getattr(args, "cmd", None):
        parser.print_help()
        return 1
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main())
