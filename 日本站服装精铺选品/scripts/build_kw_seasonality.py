#!/usr/bin/env python3
"""把新采的关键词逐月趋势整理进 raw/，并建立关键词->市场映射供 analyze.py 按市场算独立CV。
- 从 /tmp/kwt-<slug>/raw.xlsx 抽 JP-monthly-* sheet -> raw/keyword_trend_<slug>_monthly.xlsx
- 生成 raw/keyword_miner_fresh.xlsx：关键词 + 市场列（市场名严格对齐 competitor_asins.xlsx 的市场）
映射仅用语义明确对应的市场；不确定的市场不绑定（保持旧口径，报告标注）。
"""
import glob, os, openpyxl, sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RAW = os.path.join(ROOT, "raw")

# slug -> (关键词原文, [市场精确名])  市场名来自 report_data.json 的 market 字段
MAP = {
    "kutsushita": ("メンズ靴下", ["メンズカーフソックス", "メンズスポーツソックス"]),
    "belt":       ("メンズベルト", ["メンズベルト"]),
    "cap":        ("キャップメンズ", ["メンズキャップ"]),
    "pajama":     ("パジャマレディース", ["レディースパジャマ・ルームウェアセット"]),
    "skirt":      ("レディーススカート", ["レディーススカート"]),
    "denim":      ("デニムメンズ", ["レディースジーンズ"]),
    "inner":      ("メンズインナー", ["メンズインナーシャツ", "メンズボクサー"]),
    "chino":      ("チノパンメンズ", ["メンズロングパンツ"]),
    "cargo":      ("カーゴパンツメンズ", ["メンズロングパンツ"]),
    "muji_tshirt":("無地Tシャツ", ["メンズTシャツ・カットソー", "レディースTシャツ・カットソー"]),
    "tank":       ("タンクトップメンズ", ["メンズTシャツ・カットソー"]),
    "polo":       ("ポロシャツメンズ", ["メンズTシャツ・カットソー"]),
    "sweatpants": ("スウェットパンツ", ["レディースロングパンツ"]),
    "scarf":      ("レディースマフラー", []),   # 冬品，无对应春夏秋市场
    "tights":     ("レディースタイツ", []),      # 冬品
}

def extract_monthly(src):
    wb = openpyxl.load_workbook(src, read_only=True, data_only=True)
    sn = next((s for s in wb.sheetnames if "monthly" in s.lower() and not s.startswith("aba")), None)
    if not sn:
        return None
    return [list(r) for r in wb[sn].values]

def main():
    pairs = []  # (keyword, market)
    extracted = []
    for src in sorted(glob.glob("/tmp/kwt-*/raw.xlsx")):
        slug = os.path.basename(os.path.dirname(src)).replace("kwt-", "")
        if slug not in MAP:
            print("跳过未映射:", slug); continue
        _, markets = MAP[slug]
        rows = extract_monthly(src)
        if not rows:
            print("无monthly:", slug); continue
        # 用趋势文件里的真实关键词文本，保证 analyze.py 精确匹配
        hdr = [str(x) for x in rows[0]]
        ki = hdr.index("关键词")
        kw = str(rows[1][ki]) if len(rows) > 1 else slug
        # 写出 monthly 文件（保留原表头：站点/月份/关键词/ABA排名/月搜索量/...）
        out = os.path.join(RAW, f"keyword_trend_{slug}_monthly.xlsx")
        wb = openpyxl.Workbook(); ws = wb.active; ws.title = "monthly"
        for r in rows:
            ws.append(r)
        wb.save(out)
        extracted.append((slug, kw, len(rows) - 1))
        # 统一关键词文本为映射词（原文可能是长尾），保证趋势文件与映射一致
        for m in markets:
            pairs.append((kw, m))

    # keyword_miner_fresh.xlsx：让 analyze.py 建立 keyword->market
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Sheet1"
    ws.append(["关键词", "市场", "月搜索量", "月购买量", "竞争度", "来源文件"])
    for kw, m in pairs:
        ws.append([kw, m, None, None, None, "fresh-kw-trend"])
    out = os.path.join(RAW, "keyword_miner_fresh.xlsx")
    wb.save(out)

    print("抽取月度文件:")
    for slug, kw, n in extracted:
        print(f"  {slug:14s} {kw:14s} {n}月")
    print(f"映射对 {len(pairs)} 条 -> {out}")

if __name__ == "__main__":
    main()
