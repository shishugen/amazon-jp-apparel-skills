#!/usr/bin/env python3
"""店铺结构分析 + Store Research Score.

用法:
  python store-analysis.py <store-products.xlsx> [--scope brand|seller] [--label NAME] [--json OUT]

输入: SellerSprite 查竞品导出的 xlsx（66 列，见 references/sellersprite-fields.md）
输出: 店铺结构指标 + Store Research Score（缺维度 N/A、权重归一化）
"""
import argparse
import json
import statistics
import sys
from collections import Counter, defaultdict
from datetime import datetime, date

try:
    import openpyxl
except ImportError:
    sys.exit("需要 openpyxl: pip install -i https://pypi.tuna.tsinghua.edu.cn/simple openpyxl")


# Score 权重（与 SKILL.md §7 一致）
WEIGHTS = {
    "coverage": 15, "stable_30_100": 20, "stability": 15,
    "anti_hit": 10, "lifecycle": 10, "category_focus": 10,
    "low_review": 10, "price_band": 5, "new_success": 5,
}


def num(v, default=0.0):
    if v is None or v == "" or v == "--":
        return default
    try:
        return float(str(v).replace(",", ""))
    except ValueError:
        return default


def to_date(v):
    if not v:
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(str(v)[:19], fmt).date()
        except ValueError:
            continue
    return None


def read_rows(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    it = ws.iter_rows(values_only=True)
    header = [str(c).strip() if c is not None else "" for c in next(it)]
    rows = [dict(zip(header, r)) for r in it if any(x is not None for x in r)]
    wb.close()
    return rows


def aggregate_parents(rows):
    """按父ASIN聚合。取父行月销量；无父行取子体最大值。返回家族列表。"""
    groups = defaultdict(list)
    for r in rows:
        p = str(r.get("父ASIN") or r.get("ASIN") or "").strip()
        if p:
            groups[p].append(r)
    families = []
    for parent, members in groups.items():
        parent_rows = [m for m in members if str(m.get("ASIN") or "").strip() == parent]
        child_units = [num(m.get("月销量")) for m in members]
        if parent_rows:
            units = num(parent_rows[0].get("月销量"))
            source = "parent_row"
        else:
            units = max(child_units) if child_units else 0.0
            source = "max_child_estimated"
        base = parent_rows[0] if parent_rows else members[0]
        families.append({
            "parent_asin": parent,
            "units": units,
            "unit_source": source,
            "price": num(base.get("价格(円)")),
            "reviews": num(base.get("评分数")),
            "rating": num(base.get("评分")),
            "bsr": num(base.get("大类BSR")),
            "launch": to_date(base.get("上架时间")),
            "launch_days": num(base.get("上架天数")),
            "category": str(base.get("小类目") or base.get("大类目") or "").strip(),
            "seller_land": str(base.get("卖家所属地") or "").strip(),
            "variations": num(base.get("变体数")),
            "child_count": len(members),
            "motions": sum(child_units),
        })
    return families


def band(u):
    if u <= 0: return "none"
    if u < 10: return "very_low"
    if u < 30: return "low"
    if u < 50: return "stable_small"
    if u <= 100: return "core"
    if u < 200: return "mid_high"
    return "high"


def analyze(families, scope="brand", label="", price_min=2000, price_max=3000):
    total = len(families)
    active = [f for f in families if f["units"] > 0]
    coverage = len(active) / total if total else 0.0
    counts = Counter(band(f["units"]) for f in families)
    stable = counts["stable_small"] + counts["core"]

    ranked = sorted(families, key=lambda f: f["units"], reverse=True)
    total_units = sum(f["units"] for f in families)
    top3 = sum(f["units"] for f in ranked[:3]) / total_units if total_units else 0.0
    top10 = sum(f["units"] for f in ranked[:10]) / total_units if total_units else 0.0

    prices = [f["price"] for f in families if f["price"] > 0]
    in_band = sum(1 for p in prices if price_min <= p <= price_max) / len(prices) if prices else 0.0
    cats = Counter(f["category"] for f in families if f["category"])
    top3_cat = sum(c for _, c in cats.most_common(3)) / sum(cats.values()) if cats else 0.0

    today = date.today()
    ages = [(today - f["launch"]).days for f in families if f["launch"]]
    old12 = sum(1 for a in ages if a >= 365) / len(ages) if ages else 0.0
    old24 = sum(1 for a in ages if a >= 730) / len(ages) if ages else 0.0
    old_active = sum(1 for f in families if f["launch"] and (today - f["launch"]).days >= 365 and f["units"] >= 30)

    new12 = [f for f in families if f["launch"] and (today - f["launch"]).days <= 365]
    new_over30 = sum(1 for f in new12 if f["units"] >= 30)
    new_success = new_over30 / len(new12) if new12 else None

    low_review = sum(1 for f in families
                     if (f["reviews"] < 100 and 30 <= f["units"] <= 100)
                     or (f["reviews"] < 50 and f["units"] >= 30))

    lands = Counter(f["seller_land"] for f in families if f["seller_land"])

    dims = {}
    dims["coverage"] = 100.0 * min(1.0, coverage / 0.70)
    dims["stable_30_100"] = 100.0 * min(1.0, (stable / total) / 0.35) if total else 0.0
    dims["anti_hit"] = 100.0 * max(0.0, min(1.0, (0.75 - top3) / (0.75 - 0.35)))
    dims["lifecycle"] = 100.0 * min(1.0, old12 / 0.60)
    dims["category_focus"] = 100.0 * min(1.0, top3_cat / 0.70)
    dims["low_review"] = 100.0 * min(1.0, low_review / 8.0)
    dims["price_band"] = 100.0 * min(1.0, in_band / 0.60)
    dims["new_success"] = 100.0 * min(1.0, new_success / 0.40) if new_success is not None else None
    dims["stability"] = None  # 需逐月数据，见 sellersprite-history-export

    effective = {k: w for k, w in WEIGHTS.items() if dims.get(k) is not None}
    score_coverage = sum(effective.values())
    raw = sum(dims[k] * w for k, w in effective.items())
    score = raw / score_coverage if score_coverage else None

    if score is None or score_coverage < 70:
        grade = "C"
    elif score >= 80:
        grade = "A"
    elif score >= 70:
        grade = "B"
    else:
        grade = "C"

    return {
        "label": label, "scope": scope,
        "total_families": total, "active_families": len(active),
        "coverage": round(coverage, 4),
        "bands": dict(counts), "stable_30_100_count": stable,
        "top3_share_estimated": round(top3, 4), "top10_share_estimated": round(top10, 4),
        "price_median": round(statistics.median(prices), 2) if prices else None,
        "price_band_ratio": round(in_band, 4),
        "top3_category_share": round(top3_cat, 4),
        "old12_ratio": round(old12, 4), "old24_ratio": round(old24, 4),
        "old_active_ge30": old_active,
        "new12_count": len(new12), "new_success_snapshot": round(new_success, 4) if new_success is not None else None,
        "low_review_high_sales": low_review,
        "seller_lands": dict(lands),
        "dimensions": {k: (round(v, 2) if v is not None else "N/A") for k, v in dims.items()},
        "score_coverage": score_coverage,
        "score": round(score, 2) if score is not None else None,
        "grade": grade,
        "data_flags": ["Estimated: 销量/占比基于卖家精灵30天估算", "Snapshot: 新品成功率为当前快照口径",
                       "N/A: 季节稳定性缺逐月数据"] ,
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx")
    ap.add_argument("--scope", default="brand", choices=["brand", "seller"])
    ap.add_argument("--label", default="")
    ap.add_argument("--price-min", type=float, default=2000)
    ap.add_argument("--price-max", type=float, default=3000)
    ap.add_argument("--json", dest="json_out")
    args = ap.parse_args()

    rows = read_rows(args.xlsx)
    families = aggregate_parents(rows)
    result = analyze(families, args.scope, args.label or args.xlsx, args.price_min, args.price_max)
    text = json.dumps(result, ensure_ascii=False, indent=2)
    if args.json_out:
        with open(args.json_out, "w", encoding="utf-8") as f:
            f.write(text)
        print(f"写入 {args.json_out}")
    else:
        print(text)


if __name__ == "__main__":
    main()
