#!/usr/bin/env python3
"""生成空白采集模板 templates/product-research.xlsx。表头即列名，不要改列名。"""
from __future__ import annotations

import argparse
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

SHEETS = {
    "竞品ASIN": ["ASIN", "商品", "品牌", "卖家名称", "价格", "月销量", "月销售额", "评论数", "星级", "类目", "市场", "图片链接", "上架时间", "变体", "卖家数", "广告竞争", "退货率"],
    "关键词挖掘": ["关键词", "市场", "月搜索量", "月购买量", "竞争度"],
    "关键词趋势": ["月份", "关键词", "月搜索量", "月购买量"],
    "ASIN销量趋势": ["月份", "ASIN", "月销量", "月销售额"],
    "ASIN关键词反查": ["ASIN", "关键词", "月搜索量", "竞争度", "流量来源"],
}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", type=Path, default=Path(__file__).resolve().parent.parent / "templates" / "product-research.xlsx")
    args = ap.parse_args()
    wb = Workbook()
    wb.remove(wb.active)
    for name, headers in SHEETS.items():
        ws = wb.create_sheet(name)
        ws.append(headers)
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="2F5597")
        ws.freeze_panes = "A2"
        for i, h in enumerate(headers, 1):
            ws.column_dimensions[ws.cell(1, i).column_letter].width = max(14, len(h) * 2 + 4)
    args.out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(args.out)
    print(f"模板已生成：{args.out}")


if __name__ == "__main__":
    main()
