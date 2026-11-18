#!/usr/bin/env python3
"""生成模拟卖家精灵导出，仅用于流程验证。默认输出到 /tmp，不写进 raw/。"""
from __future__ import annotations

import argparse
import random
from datetime import date
from pathlib import Path

from openpyxl import Workbook

MARKETS = {
    "男士无领衬衫": {"base_price": 2980, "base_sales": 52, "kw": "メンズ ノーカラーシャツ", "cv": 0.12},
    "女士打底衫": {"base_price": 2480, "base_sales": 61, "kw": "レディース インナー トップス", "cv": 0.18},
    "男士休闲裤": {"base_price": 2880, "base_sales": 44, "kw": "メンズ カジュアルパンツ", "cv": 0.24},
    "女士家居服": {"base_price": 2680, "base_sales": 38, "kw": "パジャマ レディース", "cv": 0.48},
}
BRANDS = ["ACTIBIZ", "nano・universe", "SHIPS", "Urban Research", "GAP", "UNIQLO", "무인양품", "MUJI"]


def save(ws_rows, path: Path, headers):
    wb = Workbook()
    ws = wb.active
    ws.append(headers)
    for r in ws_rows:
        ws.append(r)
    ws.freeze_panes = "A2"
    for col in ws.columns:
        width = max((len(str(c.value or "")) for c in col), default=10)
        ws.column_dimensions[col[0].column_letter].width = min(max(width + 2, 12), 46)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", type=Path, default=Path("/tmp/jp-apparel-demo"))
    args = ap.parse_args()
    raw = args.out / "raw"
    raw.mkdir(parents=True, exist_ok=True)
    rnd = random.Random(42)

    kw_rows, comp_rows, sales_rows = [], [], []
    asin_map = {}
    idx = 0
    for market, cfg in MARKETS.items():
        for i in range(14):
            idx += 1
            asin = f"B0{idx:08d}"
            asin_map.setdefault(market, []).append(asin)
            price = max(1500, int(cfg["base_price"] * rnd.uniform(0.82, 1.18)))
            sales = max(8, int(cfg["base_sales"] * rnd.uniform(0.35, 1.9)))
            comp_rows.append([
                asin,
                f"{market} 商品 {i+1} 長袖 カジュアル メンズ レディース",
                BRANDS[i % len(BRANDS)],
                f"卖家店铺 {i % 5 + 1}",
                price,
                sales,
                price * sales,
                max(3, int(rnd.gammavariate(2, 28))),
                round(rnd.uniform(3.7, 4.8), 1),
                market,
                market,
                "",
                f"202{rnd.randint(2,5)}-{rnd.randint(1,12):02d}",
                "",
            ])
            for back in range(24, 0, -1):
                y, m = 2026, 11 - back
                while m <= 0:
                    y -= 1
                    m += 12
                season = 1 + cfg["cv"] * 0.9 * (1 if m in (3, 4, 10, 11) else (-1 if m in (1, 2, 7, 8) else 0))
                sales_rows.append([f"{y}-{m:02d}", asin, max(1, int(sales * season * rnd.uniform(0.85, 1.15))), ""])

        kw_rows.append([cfg["kw"], market, int(18000 * rnd.uniform(0.6, 1.5)),
                        int(900 * rnd.uniform(0.5, 1.6)), rnd.choice(["低", "中", "高"])])

        for back in range(24, 0, -1):
            y, m = 2026, 11 - back
            while m <= 0:
                y -= 1
                m += 12
            season = 1 + cfg["cv"] * (1 if m in (3, 4, 10, 11) else (-1 if m in (1, 2, 7, 8) else 0))
            # 关键词趋势在下方统一按市场写文件
    for market, cfg in MARKETS.items():
        trend_rows = []
        for back in range(24, 0, -1):
            y, m = 2026, 11 - back
            while m <= 0:
                y -= 1
                m += 12
            season = 1 + cfg["cv"] * (1 if m in (3, 4, 10, 11) else (-1 if m in (1, 2, 7, 8) else 0))
            trend_rows.append([f"{y}-{m:02d}", cfg["kw"], int(18000 * season * rnd.uniform(0.9, 1.1)),
                               int(900 * season * rnd.uniform(0.9, 1.1))])
        save(trend_rows, raw / f"keyword_trend_{market}.xlsx", ["月份", "关键词", "月搜索量", "月购买量"])

    save(kw_rows, raw / "keyword_miner.xlsx", ["关键词", "市场", "月搜索量", "月购买量", "竞争度"])
    save(comp_rows, raw / "competitor_asins.xlsx",
         ["ASIN", "商品", "品牌", "卖家名称", "价格", "月销量", "月销售额", "评论数", "星级", "类目", "市场", "图片链接", "上架时间", "变体"])
    save(sales_rows, raw / "asin_sales_trend_all.xlsx", ["月份", "ASIN", "月销量", "月销售额"])

    print(f"模拟数据已生成：{raw}")
    print(f"  竞品 {len(comp_rows)} 行，ASIN趋势 {len(sales_rows)} 行，关键词趋势 {len(MARKETS)*24} 行")
    print("  注意：模拟数据只能验证流程，不能用于采购决策。")


if __name__ == "__main__":
    main()
