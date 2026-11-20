#!/usr/bin/env python3
"""亚马逊日本站服装精铺选品：读取卖家精灵导出 → 分析 → Excel/Markdown/JSON。

只依赖 openpyxl。所有金额 JPY，费率用小数。
用法见 SKILL.md。
"""
from __future__ import annotations

import argparse
import csv
import json
import math
import re
import statistics
import sys
import unicodedata
from collections import defaultdict
from datetime import date, datetime, timezone
from pathlib import Path

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit("缺少 openpyxl。请先安装：pip install openpyxl（本机用清华镜像）")


ROOT = Path(__file__).resolve().parent.parent
DEFAULT_PARAMS = ROOT / "templates" / "user-params.json"
MISSING = "缺失"

ALIASES = {
    "asin": ["asin", "商品asin", "父asin", "子asin", "父(子)体asin", "父体asin", "商品编号"],
    "title": ["商品", "商品名称", "标题", "title", "商品标题", "产品名称", "商品名"],
    "brand": ["品牌", "brand", "メーカー"],
    "gender": ["性别", "gender", "性別"],
    "price": ["价格", "售价", "平均单价", "buybox价格", "price", "日元价格", "当前价格"],
    "monthly_sales": ["月销量", "销量", "月销量(件)", "sales", "monthly sales", "estimated monthly sales", "预估月销量"],
    "monthly_revenue": ["月销售额", "销售额", "revenue", "月销售金额"],
    "reviews": ["评论数", "评论数量", "review count", "reviews", "rating count", "评价数"],
    "rating": ["评分", "星级", "rating", "star", "商品评分"],
    "bsr": ["bsr", "类目排名", "排名", "best sellers rank"],
    "category": ["类目", "所属类目", "分类", "category", "商品类目"],
    "market": ["市场", "分析市场", "品类", "细分市场", "group", "赛道", "关键词组"],
    "image": ["图片", "图片链接", "image", "image url", "商品图片", "主图"],
    "launch_date": ["上架时间", "上市时间", "上架日期", "launch date", "首次上架"],
    "seller_name": ["卖家名称", "卖家", "店铺", "店铺名", "seller", "seller name", "商家", "卖家店铺"],
    "seller_count": ["卖家数", "卖家数量", "seller count", "sellers"],
    "seller_country": ["卖家所属地", "seller country", "seller location", "卖家所在地"],
    "seller_info": ["卖家信息", "seller info", "商家信息"],
    "fulfillment": ["配送方式", "fulfillment", "配送", "buybox类型", "buybox类型(配送)"],
    "a_plus": ["a+页面", "a+", "a plus", "a+页面标记", "has a+"],
    "video": ["视频介绍", "商品视频", "有视频", "video", "视频"],
    "variation": ["变体", "颜色", "尺码", "子体", "variation", "规格"],
    "month": ["月份", "月", "month", "date", "日期"],
    "keyword": ["关键词", "keyword", "搜索词", "キーワード"],
    "search_volume": ["月搜索量", "搜索量", "search volume", "搜索指数"],
    "purchase_volume": ["月购买量", "购买量", "purchase volume"],
    "competition": ["竞争度", "竞争程度", "competition", "竞争激烈程度"],
    "ad_competition": ["广告竞争", "ad competition", "广告竞争度"],
    "return_rate": ["退货率", "return rate"],
}


def norm(value) -> str:
    if value is None:
        return ""
    return unicodedata.normalize("NFKC", str(value)).strip().lower()


def safe_str(value) -> str:
    if value is None:
        return ""
    return unicodedata.normalize("NFKC", str(value)).strip()


CHINA_SIGNALS = ["cn", "china", "中国", "中華人民共和国", "people's republic of china"]


def is_china_seller(item: dict) -> bool | None:
    """综合卖家所属地 + 卖家信息地址判断中国卖家；无信号返回 None(Unknown)，不猜测。"""
    vals = [safe_str(item.get("seller_country")), safe_str(item.get("seller_info"))]
    text = " ".join(vals).lower()
    if any(s in text for s in CHINA_SIGNALS):
        return True
    # 明确是日本（JP / 日本 / 日本国）则判否；其余信号不足 → Unknown
    jp = ["jp", "japan", "日本"]
    if any(s in text for s in jp) and not any(s in text for s in CHINA_SIGNALS):
        return False
    return None


def pass_seller_filters(item: dict, params: dict, gaps: list, market: str) -> bool:
    """按卖家所属地（含中国卖家判断）/ 配送方式过滤；默认 all 全部放行。

    卖家精灵 UI 里「是否为中国卖家」与「卖家所属地」是同一个筛选项（所属地输入框），
    因此这里统一用 seller_country 一个参数，不再拆成两个独立开关。
    """
    country = str(params.get("seller_country", "all") or "all").strip().upper()
    if country not in ("ALL", ""):
        if country in ("CN", "CHINA", "中国"):
            # 中国卖家：所属地字段常为空，用综合判断（所属地 + 卖家信息地址）
            if item["is_china_seller"] is not True:
                return False
        elif country in ("JP", "JAPAN", "日本"):
            if item["is_china_seller"] is not False:
                return False
        elif country == "UNKNOWN":
            if item["seller_country"] not in ("", MISSING, "N/A"):
                return False
        elif item["seller_country"].upper() != country:
            # 其他国家代码按所属地字段精确匹配
            return False

    fulfil = str(params.get("fulfillment", "all") or "all").strip().upper()
    if fulfil not in ("ALL", ""):
        # 卖家精灵 UI 为 AMZ / FBA / FBM 三个独立复选框，可多选（逗号分隔），精确匹配不做映射
        allowed = {x.strip() for x in fulfil.split(",") if x.strip()}
        if item["fulfillment"].upper() not in allowed:
            return False
    return True


def num(value):
    """把 Excel 里的数字/带单位字符串转成 float；失败返回 None。"""
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        v = float(value)
        return None if math.isnan(v) else v
    s = unicodedata.normalize("NFKC", str(value)).strip()
    if not s or s in {"-", "—", "N/A", "n/a", "null", "None", "无"}:
        return None
    pct = s.endswith("%")
    s = s.rstrip("%").replace("¥", "").replace("￥", "").replace(",", "")
    s = s.replace("件", "").replace("个", "").replace("日元", "").replace(" ", "")
    mult = 10000 if "万" in s else (1000 if "千" in s else 1)
    s = s.replace("万", "").replace("千", "")
    m = re.search(r"-?\d+(?:\.\d+)?", s)
    if not m:
        return None
    v = float(m.group()) * mult
    return v / 100 if pct else v


def norm_month(value):
    if value is None:
        return None
    if isinstance(value, (datetime, date)):
        return f"{value.year:04d}-{value.month:02d}"
    s = unicodedata.normalize("NFKC", str(value)).strip()
    m = re.search(r"(20\d{2})\s*[年/\-.]?\s*(\d{1,2})", s)
    if not m:
        return None
    y, mo = int(m.group(1)), int(m.group(2))
    return f"{y:04d}-{mo:02d}" if 1 <= mo <= 12 else None


def norm_asin(value):
    s = re.sub(r"[^A-Za-z0-9]", "", safe_str(value)).upper()
    return s if re.fullmatch(r"[A-Z0-9]{10}", s) else None


def asin_url(asin: str) -> str:
    return f"https://www.amazon.co.jp/dp/{asin}" if asin else ""


def find_header(rows):
    best_idx, best_score = 0, -1
    for idx, row in enumerate(rows[:12]):
        cells = {norm(c) for c in row if c is not None}
        score = sum(1 for aliases in ALIASES.values() if any(norm(a) in cells for a in aliases))
        if score > best_score:
            best_idx, best_score = idx, score
    return best_idx


def build_colmap(header):
    cells = [norm(c) for c in header]
    colmap = {}
    for key, aliases in ALIASES.items():
        for alias in aliases:
            a = norm(alias)
            if a in cells:
                colmap[key] = cells.index(a)
                break
    return colmap


def read_table(path: Path):
    if path.suffix.lower() == ".csv":
        with path.open("r", encoding="utf-8-sig", errors="replace", newline="") as fh:
            rows = [list(r) for r in csv.reader(fh)]
    else:
        wb = load_workbook(path, read_only=True, data_only=True)
        try:
            ws = wb[wb.sheetnames[0]]
            rows = [list(r) for r in ws.iter_rows(values_only=True)]
        finally:
            wb.close()
    if not rows:
        return [], [], 0
    hidx = find_header(rows)
    header = [safe_str(c) for c in rows[hidx]]
    colmap = build_colmap(header)
    out = []
    for r in rows[hidx + 1:]:
        rec = {k: (r[i] if i < len(r) else None) for k, i in colmap.items()}
        if not any(v not in (None, "") for v in rec.values()):
            continue
        rec["_source"] = path.name
        out.append(rec)
    return out, header, hidx + 1


def classify(path: Path, rows, header):
    keys = set(build_colmap(header))
    name = norm(path.stem)
    if "month" in keys and "asin" in keys and (keys & {"monthly_sales", "monthly_revenue"}):
        return "asin_sales"
    if "month" in keys and (keys & {"keyword", "search_volume"}):
        return "keyword_trend"
    if "asin" in keys and "keyword" in keys and not (keys & {"monthly_sales", "price"}):
        return "keyword_reverse"
    if "asin" in keys and "month" not in keys:
        return "competitor"
    if "keyword" in keys and "search_volume" in keys:
        return "keyword_miner"
    if "keyword" in keys and "month" not in keys:
        return "keyword_miner"
    if "asin" in keys:
        return "competitor"
    if "趋势" in name or "trend" in name:
        return "keyword_trend" if ("关键词" in name or "keyword" in name) else "asin_sales"
    if "反查" in name:
        return "keyword_reverse"
    if "竞品" in name or "competitor" in name or "选产品" in name:
        return "competitor"
    return "unknown"


def load_all(raw_dir: Path):
    data = {"competitor": [], "keyword_miner": [], "keyword_trend": [], "asin_sales": [], "keyword_reverse": [], "unknown": []}
    log = []
    for p in sorted(raw_dir.rglob("*")):
        if p.suffix.lower() not in {".xlsx", ".xlsm", ".csv"} or p.name.startswith("~$"):
            continue
        try:
            rows, header, hidx = read_table(p)
            kind = classify(p, rows, header)
            for r in rows:
                r["_type"] = kind
            data.setdefault(kind, []).extend(rows)
            log.append({
                "file": str(p.relative_to(raw_dir)) if raw_dir in p.parents else p.name,
                "type": kind,
                "rows": len(rows),
                "header_row": hidx,
                "modified": datetime.fromtimestamp(p.stat().st_mtime).astimezone().isoformat(timespec="seconds"),
            })
        except Exception as exc:  # 坏文件不中断整批
            log.append({"file": p.name, "type": "read_error", "rows": 0, "header_row": 0,
                        "modified": "", "error": str(exc)})
    return data, log


def load_params(path: Path | None) -> dict:
    base = json.loads(DEFAULT_PARAMS.read_text(encoding="utf-8")) if DEFAULT_PARAMS.exists() else {}
    if path and path.exists():
        user = json.loads(path.read_text(encoding="utf-8"))
        for k, v in user.items():
            if isinstance(v, dict) and isinstance(base.get(k), dict):
                base[k].update(v)
            else:
                base[k] = v
    return base


def median_or_none(values):
    vals = [v for v in values if isinstance(v, (int, float))]
    return float(statistics.median(vals)) if vals else None


def concentration(items, key_fn, total):
    """按 key 聚合销量，返回集中度 CR3/CR5 与 HHI。total<=0 返回 None。"""
    if not items or not total or total <= 0:
        return None
    grouped = defaultdict(float)
    for it in items:
        v = it.get("monthly_sales")
        if v is None:
            continue
        grouped[key_fn(it)] += v
    ordered = sorted(grouped.values(), reverse=True)
    if not ordered:
        return None
    return {
        "count": len(ordered),
        "cr3": sum(ordered[:3]) / total,
        "cr5": sum(ordered[:5]) / total,
        "hhi": sum((v / total) ** 2 for v in ordered) * 10000,
    }


def fmt(value, digits=1):
    if value is None or value == MISSING:
        return MISSING
    if isinstance(value, float):
        return f"{value:,.{digits}f}".rstrip("0").rstrip(".")
    if isinstance(value, int):
        return f"{value:,}"
    return safe_str(value) or MISSING


def compute_profit(price, costs):
    if price is None:
        return None
    c = costs or {}
    pieces = {
        "售价": price,
        "佣金": price * float(c.get("referral_rate", 0.10) or 0),
        "FBA配送费": float(c.get("fba_fee_jpy", 380) or 0),
        "采购成本": float(c.get("purchase_cost_jpy", 700) or 0),
        "头程物流": float(c.get("inbound_shipping_jpy", 180) or 0),
        "广告预留": price * float(c.get("ad_cost_rate", 0.10) or 0),
        "退货预留": price * float(c.get("return_reserve_rate", 0.05) or 0),
        "其他": float(c.get("other_cost_jpy", 0) or 0),
    }
    net = pieces["售价"] - sum(v for k, v in pieces.items() if k != "售价")
    pieces["单件净利"] = net
    pieces["净利率"] = net / price if price else None
    return pieces


SEASON_WINDOW_PRESETS = {
    "strict_year_round": (list(range(1, 13)), "全年1-12月"),
    "three_seasons_spring_summer_autumn": ([3, 4, 5, 6, 7, 8, 9, 10, 11], "春夏秋3-11月"),
}


def resolve_season_window(params):
    """按 seasonality 参数解析参与稳定性计算的月份窗口。
    三选一：strict_year_round / three_seasons_spring_summer_autumn / custom。
    只有 custom 读 season_window_months。返回 (月份集合, 标签)。"""
    pref = params.get("seasonality", "three_seasons_spring_summer_autumn")
    if pref in SEASON_WINDOW_PRESETS:
        months, label = SEASON_WINDOW_PRESETS[pref]
        return set(months), label
    if pref == "custom":
        raw = params.get("season_window_months") or []
        months = sorted({int(m) for m in raw if isinstance(m, (int, float)) and 1 <= int(m) <= 12})
        if not months:
            months = list(range(1, 13))
            label = "全年1-12月(custom空值回退)"
        else:
            label = "自定义" + ",".join(str(m) for m in months) + "月"
        return set(months), label
    # 兼容旧参数值
    if pref in ("stable", "flexible"):
        return set(range(1, 13)), "全年1-12月(旧参数兼容)"
    if pref == "strict":
        return set(range(1, 13)), "全年1-12月(旧参数兼容)"
    return set(range(1, 13)), "全年1-12月(未知参数回退)"


def score_stability(season, params):
    need = season.get("min_months_required") if season else None
    need = need if need is not None else int(params.get("min_months_for_seasonality", 6))
    if not season or season.get("months_available", 0) < need:
        return 8, "数据不足"
    cv = season.get("cv")
    if cv is None:
        return 8, "数据不足"
    th = params.get("thresholds", {})
    label = (season.get("window_label") or "目标窗口") + "稳定"
    if cv <= float(th.get("cv_excellent", 0.20)):
        return 25, label
    if cv <= float(th.get("cv_acceptable", 0.30)):
        return 20, "轻微波动"
    if cv <= float(th.get("cv_seasonal", 0.50)):
        return 11, "中等季节"
    return 4, "强季节"


def score_competition(ms, comps, params):
    score = 25.0
    th = params.get("thresholds", {})
    top10 = ms.get("top10_share")
    if top10 is not None:
        if top10 > float(th.get("top10_concentration_high", 0.55)):
            score -= 8
        elif top10 > float(th.get("top10_concentration_medium", 0.40)):
            score -= 4
    med_reviews = ms.get("median_reviews")
    if med_reviews is not None:
        if med_reviews > 200:
            score -= 7
        elif med_reviews > 100:
            score -= 4
        elif med_reviews > 50:
            score -= 2
    hhi = ms.get("brand_hhi")
    if hhi is not None and hhi > 2500:
        score -= 4
    seller_hhi = ms.get("seller_hhi")
    if seller_hhi is not None and seller_hhi > 2500:
        score -= 3
    if any(num(c.get("ad_competition")) and num(c.get("ad_competition")) > 2 for c in comps):
        score -= 3
    return max(0.0, min(25.0, round(score, 1)))


def score_profit(net, params):
    target = float(params.get("min_net_profit_jpy", 500) or 500)
    if net is None:
        return 0
    if net >= target * 1.5:
        return 20
    if net >= target:
        return 16
    if net >= target * 0.7:
        return 10
    if net > 0:
        return 5
    return 0


def score_differentiation(ms, params):
    low_ratio = ms.get("low_review_ratio")
    top10 = ms.get("top10_share")
    score = 8.0
    if low_ratio is not None:
        score += min(6.0, 6.0 * low_ratio)
    if top10 is not None:
        score += 4.0 * (1.0 - min(top10, 1.0))
    if ms.get("new_product_count") is not None and ms.get("new_product_count", 0) > 0:
        score += 2.0
    return max(0.0, min(20.0, round(score, 1)))


def score_risk(season, params, comps):
    score = 10.0
    cv = (season or {}).get("cv")
    if cv is not None:
        if cv > 0.5:
            score -= 4
        elif cv > 0.3:
            score -= 2
    need = (season or {}).get("min_months_required")
    need = need if need is not None else int(params.get("min_months_for_seasonality", 6))
    if (season or {}).get("months_available", 0) < need:
        score -= 3
    returns = [num(c.get("return_rate")) for c in comps]
    returns = [r for r in returns if r is not None]
    if returns and statistics.mean(returns) > 0.08:
        score -= 3
    low_variation = sum(1 for c in comps if safe_str(c.get("variation")))
    if comps and low_variation / len(comps) > 0.7:
        score -= 2
    return max(0.0, min(10.0, round(score, 1)))


def grade_of(total, params):
    th = params.get("thresholds", {})
    if total >= float(th.get("grade_a_min", 80)):
        return "A级 / 优先开发"
    if total >= float(th.get("grade_b_min", 70)):
        return "B级 / 可以测试"
    if total >= float(th.get("grade_c_min", 60)):
        return "C级 / 观察"
    return "淘汰"


def seasonality_for_market(market, asin_to_market, keyword_to_market, asin_trends, keyword_trends, params):
    months_back = int(params.get("history_months", 24) or 24)
    relevant_asin = {r for r, m in asin_to_market.items() if m == market}
    relevant_kw = set()
    for k, m in keyword_to_market.items():
        ms = m if isinstance(m, (list, tuple, set)) else [m]
        if market in ms:
            relevant_kw.add(k)

    asin_series = defaultdict(dict)
    for r in asin_trends:
        a = norm_asin(r.get("asin"))
        mo = norm_month(r.get("month"))
        if not a or not mo or (relevant_asin and a not in relevant_asin):
            continue
        v = num(r.get("monthly_sales"))
        if v is not None:
            asin_series[a][mo] = v

    kw_series = defaultdict(dict)
    for r in keyword_trends:
        k = safe_str(r.get("keyword"))
        mo = norm_month(r.get("month"))
        if not k or not mo or (relevant_kw and k not in relevant_kw):
            continue
        v = num(r.get("search_volume"))
        if v is not None:
            kw_series[k][mo] = v

    # 无专属数据（既无本市场ASIN销量，也无本市场关键词映射）→ 标数据不足
    # 不允许回退到"全部关键词求和"：那会把无关品曲线混进来，制造虚假稳定(CV≈0.2)
    if not asin_series and not relevant_kw:
        return None

    source = "竞品逐月销量"
    series = defaultdict(float)
    if asin_series:
        for a, bymonth in asin_series.items():
            for mo, v in bymonth.items():
                series[mo] += v
    elif kw_series:
        source = "关键词逐月搜索量"
        for k, bymonth in kw_series.items():
            for mo, v in bymonth.items():
                series[mo] += v
    else:
        return None

    all_months = sorted(series)[-months_back:]
    window, window_label = resolve_season_window(params)
    # 按窗口过滤：窗口外月份（如冬季12/1/2）不填0、不参与CV，仅留档披露
    months = [m for m in all_months if int(m.split("-")[1]) in window]
    excluded = [
        {"month": m, "value": round(series[m], 2)}
        for m in all_months if int(m.split("-")[1]) not in window
    ]
    # 最少月份阈值按窗口长度等比缩小：全年默认6 → 三季约4
    base_min = int(params.get("min_months_for_seasonality", 6))
    min_months = max(3, round(base_min * len(window) / 12))
    values = [series[m] for m in months if series[m] > 0]
    if not values:
        return {
            "source": source, "window_label": window_label,
            "window_months": sorted(window), "excluded_months": excluded,
            "months": [{"month": m, "value": round(series[m], 2)} for m in all_months],
            "months_available": 0, "min_months_required": min_months,
            "mean": None, "cv": None, "peak_trough": None, "stable_months": 0,
        }
    mean = statistics.mean(values)
    cv = statistics.pstdev(values) / mean if mean else None
    peak_trough = max(values) / min(values) if min(values) else None
    stable_count = sum(1 for v in values if mean and abs(v - mean) <= 0.2 * mean)
    return {
        "source": source,
        "window_label": window_label,
        "window_months": sorted(window),
        "excluded_months": excluded,
        "months": [{"month": m, "value": round(series[m], 2)} for m in all_months],
        "months_available": len(values),
        "min_months_required": min_months,
        "mean": round(mean, 2),
        "cv": round(cv, 4) if cv is not None else None,
        "peak_trough": round(peak_trough, 2) if peak_trough is not None else None,
        "stable_months": stable_count,
    }


def build_dashboard(comps: list, keywords: list, params: dict, ms: dict, new_def_months: int = 3) -> dict:
    """市场概览看板数据。缺字段的指标标 available=False，不编造。

    全部为「当前样本」口径，非日本站全市场。
    """
    from datetime import date
    today = date.today()

    def pct(n, d):
        return round(n / d * 100, 2) if d else None

    total_sales = sum(c["monthly_sales"] for c in comps if c["monthly_sales"] is not None)
    total_rev = sum(
        (c["monthly_revenue"] if c["monthly_revenue"] is not None
         else (c["price"] or 0) * (c["monthly_sales"] or 0))
        for c in comps
    )

    # 新品（上架 ≤ new_def_months 个月）
    def age_months(c):
        s = str(c.get("launch_date") or "")
        m = re.search(r"(\d{4})[-/](\d{1,2})", s)
        if not m:
            return None
        y, mo = int(m.group(1)), int(m.group(2))
        if not 1 <= mo <= 12:
            return None
        return (today.year - y) * 12 + (today.month - mo)

    new_items = [c for c in comps if (age_months(c) is not None and age_months(c) <= new_def_months)]
    new_sales = sum(c["monthly_sales"] for c in new_items if c["monthly_sales"] is not None)
    new_rev = sum(
        (c["monthly_revenue"] if c["monthly_revenue"] is not None
         else (c["price"] or 0) * (c["monthly_sales"] or 0))
        for c in new_items
    )

    # 卖家所属地分布（按销量）
    geo = {"中国": 0.0, "日本": 0.0, "其他/未知": 0.0}
    for c in comps:
        u = c["monthly_sales"] or 0
        if c.get("is_china_seller") is True:
            geo["中国"] += u
        elif c.get("is_china_seller") is False:
            geo["日本"] += u
        else:
            geo["其他/未知"] += u
    geo_share = {k: pct(v, total_sales) for k, v in geo.items()}

    # 配送方式分布（按销量）
    ful = {"AMZ": 0.0, "FBA": 0.0, "FBM": 0.0, "未知": 0.0}
    for c in comps:
        u = c["monthly_sales"] or 0
        f = str(c.get("fulfillment") or "").strip().upper()
        ful[f if f in ful else "未知"] += u
    ful_share = {k: pct(v, total_sales) for k, v in ful.items()}

    # 搜索购买比
    sv = sum((k.get("search_volume") or 0) for k in keywords)
    pv = sum((k.get("purchase_volume") or 0) for k in keywords)

    # 平均上架天数 / 平均评分数
    ages = [(today - datetime.strptime(re.search(r"(\d{4})[-/](\d{1,2})", str(c.get("launch_date"))).group(0).replace("/", "-"), "%Y-%m").date()).days
            for c in comps if c.get("launch_date") and re.search(r"(\d{4})[-/](\d{1,2})", str(c.get("launch_date")))]
    ratings = [c["rating"] for c in comps if c["rating"] is not None]

    # 头部商品集中度 CR-N 序列（商品/品牌/卖家），供"前N名"下拉联动
    def cr_series(key_fn):
        grouped = defaultdict(float)
        for c in comps:
            v = c.get("monthly_sales")
            if v is not None:
                grouped[key_fn(c)] += v
        ordered = sorted(grouped.values(), reverse=True)
        out = {}
        for n in (3, 5, 10, 20):
            out[f"cr{n}"] = round(sum(ordered[:n]) / total_sales, 4) if total_sales else None
        return out

    cr_product = cr_series(lambda c: c["asin"])
    cr_brand = cr_series(lambda c: c["brand"])
    cr_seller = cr_series(lambda c: c.get("seller") if c.get("seller") and c["seller"] != MISSING else c["asin"])

    # A+ / 视频销量占比（有 A+ 或有视频的商品销量之和）
    def has_flag(v):
        s = safe_str(v).upper()
        return s in ("Y", "YES", "TRUE", "1", "有", "A+", "有视频")
    ap_sales = sum((c["monthly_sales"] or 0) for c in comps
                   if has_flag(c.get("a_plus")) or has_flag(c.get("video")))
    ap_available = any(safe_str(c.get("a_plus")) or safe_str(c.get("video")) for c in comps)

    return {
        "scope": "当前样本（非日本站全市场）",
        "new_definition_months": new_def_months,
        "metrics": {
            "monthly_sales": {"value": round(total_sales, 2), "available": True, "label": "月销量"},
            "monthly_revenue": {"value": round(total_rev, 2), "available": True, "label": "月销售额(円)"},
            "product_cr": {"value": cr_product, "available": total_sales > 0, "label": "商品集中度", "fmt": "cr"},
            "brand_cr": {"value": cr_brand, "available": total_sales > 0, "label": "品牌集中度", "fmt": "cr"},
            "seller_cr": {"value": cr_seller, "available": total_sales > 0, "label": "卖家集中度", "fmt": "cr"},
            "new_sales": {"value": round(new_sales, 2), "available": True, "label": f"新品销量/占比(≤{new_def_months}月)"},
            "new_sales_share": {"value": pct(new_sales, total_sales), "available": True, "label": "新品销量占比", "fmt": "pct"},
            "new_revenue": {"value": round(new_rev, 2), "available": True, "label": "新品销售额(円)"},
            "geo_share": {"value": geo_share, "available": any(geo.values()), "label": "卖家所属地销量占比", "fmt": "bars"},
            "avg_price": {"value": ms.get("avg_price"), "available": ms.get("avg_price") is not None, "label": "平均价格(円)"},
            "fulfillment_share": {"value": ful_share, "available": any(ful.values()), "label": "配送方式销量占比", "fmt": "bars"},
            "a_plus_video": {"value": round(ap_sales / total_sales, 4) if total_sales else None, "available": ap_available, "label": "有A+有视频销量占比", "fmt": "pct", "need": "导出含 A+页面/视频介绍 列"},
            "return_rate": {"value": None, "available": False, "label": "退货率", "need": "导出含退货率列"},
            "search_purchase_ratio": {"value": round(pv / sv, 6) if sv else None, "available": sv > 0, "label": "搜索购买比", "fmt": "permille"},
            "last30_units": {"value": round(total_sales, 2), "available": True, "label": "近30天售出件数(样本)", "note": "等同样本月销量"},
            "product_features": {"value": None, "available": False, "label": "产品特征分析", "need": "导出含产品特征列"},
            "avg_listing_days": {"value": round(statistics.mean(ages)) if ages else None, "available": bool(ages), "label": "平均上架天数"},
            "avg_rating": {"value": round(statistics.mean(ratings), 2) if ratings else None, "available": bool(ratings), "label": "平均评分数"},
        },
    }


def analyze_market(market, competitors, asin_to_market, keyword_to_market, asin_trends, keyword_trends, keywords, params):
    gaps = []
    comps = []
    seen = set()
    for c in competitors:
        a = norm_asin(c.get("asin"))
        if not a or a in seen:
            continue
        seen.add(a)
        item = dict(c)
        item["asin"] = a
        item["title"] = safe_str(c.get("title")) or MISSING
        item["gender"] = safe_str(c.get("gender")) or MISSING
        item["brand"] = safe_str(c.get("brand")) or MISSING
        item["price"] = num(c.get("price"))
        item["monthly_sales"] = num(c.get("monthly_sales"))
        item["monthly_revenue"] = num(c.get("monthly_revenue"))
        item["reviews"] = num(c.get("reviews"))
        item["rating"] = num(c.get("rating"))
        item["bsr"] = num(c.get("bsr"))
        item["image"] = safe_str(c.get("image"))
        item["launch_date"] = norm_month(c.get("launch_date"))
        item["seller"] = safe_str(c.get("seller_name")) or MISSING
        item["seller_count"] = num(c.get("seller_count"))
        item["seller_country"] = safe_str(c.get("seller_country")) or MISSING
        item["seller_info"] = safe_str(c.get("seller_info")) or MISSING
        item["fulfillment"] = safe_str(c.get("fulfillment")) or MISSING
        item["is_china_seller"] = is_china_seller(item)
        item["a_plus"] = safe_str(c.get("a_plus"))
        item["video"] = safe_str(c.get("video"))
        item["ad_competition"] = num(c.get("ad_competition"))
        item["return_rate"] = num(c.get("return_rate"))
        item["url"] = asin_url(a)
        if item["monthly_sales"] is not None and item["monthly_sales"] > float(params.get("outlier_sales_max", 5000)):
            gaps.append(f"{market}：ASIN {a} 月销量 {fmt(item['monthly_sales'])} 异常，已排除")
            continue
        if not pass_seller_filters(item, params, gaps, market):
            continue
        comps.append(item)
    if not comps:
        return None

    comps.sort(key=lambda x: (x["monthly_sales"] if x["monthly_sales"] is not None else -1), reverse=True)
    total_sales = sum(c["monthly_sales"] for c in comps if c["monthly_sales"] is not None)
    total_revenue = sum(
        (c["monthly_revenue"] if c["monthly_revenue"] is not None
         else (c["price"] or 0) * (c["monthly_sales"] or 0))
        for c in comps
    )
    for c in comps:
        c["share"] = (c["monthly_sales"] / total_sales) if total_sales and c["monthly_sales"] is not None else None
        c["market"] = market

    top_n = min(10, len(comps))
    top10 = sum(c["monthly_sales"] or 0 for c in comps[:top_n]) / total_sales if total_sales else None
    top20_n = min(20, len(comps))
    top20 = sum(c["monthly_sales"] or 0 for c in comps[:top20_n]) / total_sales if total_sales else None

    product = concentration(comps, lambda c: c["asin"], total_sales)
    brand = concentration(comps, lambda c: c["brand"], total_sales)
    seller_comps = [c for c in comps if c.get("seller") and c["seller"] != MISSING]
    seller_total = sum(c["monthly_sales"] for c in seller_comps if c["monthly_sales"] is not None)
    seller = concentration(seller_comps, lambda c: c["seller"], seller_total)
    if not seller:
        gaps.append(f"{market}：竞品缺卖家名称，卖家集中度按缺失处理")
    prices = [c["price"] for c in comps if c["price"] is not None]

    low_max = float(params.get("low_review_max", 50))
    low_reviews = [c for c in comps if c["reviews"] is not None and c["reviews"] <= low_max]
    mid_sales = [c for c in comps if c["monthly_sales"] is not None
                 and float(params.get("monthly_sales_min_acceptable", 30)) <= c["monthly_sales"]
                 <= float(params.get("monthly_sales_max_reasonable", 100))]
    ms = {
        "market": market,
        "competitor_count": len(comps),
        "total_monthly_sales": round(total_sales, 2),
        "total_monthly_revenue": round(total_revenue, 2),
        "product_cr3": round(product["cr3"], 4) if product else None,
        "product_cr5": round(product["cr5"], 4) if product else None,
        "product_cr10": round(top10, 4) if top10 is not None else None,
        "product_hhi": round(product["hhi"], 1) if product else None,
        "top10_share": round(top10, 4) if top10 is not None else None,
        "top20_share": round(top20, 4) if top20 is not None else None,
        "brand_cr3": round(brand["cr3"], 4) if brand else None,
        "brand_cr5": round(brand["cr5"], 4) if brand else None,
        "brand_hhi": round(brand["hhi"], 1) if brand else None,
        "brand_count": brand["count"] if brand else 0,
        "seller_cr3": round(seller["cr3"], 4) if seller else None,
        "seller_cr5": round(seller["cr5"], 4) if seller else None,
        "seller_hhi": round(seller["hhi"], 1) if seller else None,
        "seller_count": seller["count"] if seller else 0,
        "seller_data_available": bool(seller),
        "low_review_count": len(low_reviews),
        "low_review_ratio": round(len(low_reviews) / len(comps), 4) if comps else None,
        "mid_sales_count": len(mid_sales),
        "avg_price": round(statistics.mean(prices), 2) if prices else None,
        "median_price": median_or_none(prices),
        "median_reviews": median_or_none([c["reviews"] for c in comps]),
        "median_rating": median_or_none([c["rating"] for c in comps]),
    }

    season = seasonality_for_market(market, asin_to_market, keyword_to_market, asin_trends, keyword_trends, params)
    if season is None:
        gaps.append(f"{market}：缺逐月趋势数据，季节性按数据不足处理")
    st_score, st_label = score_stability(season, params)
    if season:
        season["label"] = st_label

    in_range = [c["price"] for c in comps if c["price"] is not None
                and float(params.get("price_min", 2000)) <= c["price"] <= float(params.get("price_max", 3000))]
    suggested_price = median_or_none(in_range) if in_range else median_or_none([c["price"] for c in comps])
    if suggested_price is None:
        gaps.append(f"{market}：竞品缺价格，利润无法计算")
    profit = compute_profit(suggested_price, params.get("costs", {}))
    net = profit["单件净利"] if profit else None
    target_sales = int(params.get("monthly_sales_target", 50) or 50)
    if total_sales and total_sales < target_sales * 10:
        gaps.append(f"{market}：样本月销量合计 {fmt(total_sales)} 偏小，50件/月需谨慎验证")

    scores = {
        "stability": st_score,
        "competition": score_competition(ms, comps, params),
        "profit": score_profit(net, params),
        "differentiation": score_differentiation(ms, params),
        "risk": score_risk(season, params, comps),
    }
    total = round(sum(scores.values()), 1)

    category_counts = defaultdict(int)
    for c in comps:
        cat = safe_str(c.get("category")) or safe_str(c.get("market"))
        if cat and cat != MISSING:
            category_counts[cat] += 1
    main_category = max(category_counts, key=category_counts.get) if category_counts else market
    # 类目路径可能是一长串（ファッション:メンズ:…:ロングパンツ），取最后一级做展示名
    if ':' in str(main_category):
        main_category = str(main_category).split(':')[-1].strip() or market
    source_files = sorted({safe_str(c.get("_source")) for c in comps if c.get("_source")})

    dashboard = build_dashboard(comps, keywords, params, ms, new_def_months=3)

    candidate = {
        "market": market,
        # 市场名（小类目）本身已含性别+品类，直接用它做展示名，最简洁且不重复
        "name": f"{market} 精铺机会",
        "display_category": main_category,
        "category": main_category,
        "grade": grade_of(total, params),
        "total_score": total,
        "scores": scores,
        "seasonality": season or {"label": "数据不足", "months_available": 0, "cv": None, "months": []},
        "seasonality_label": st_label,
        "market_share": ms,
        "dashboard": dashboard,
        "suggested_price": suggested_price,
        "profit": profit,
        "net_profit": net,
        "monthly_net_profit": round(net * target_sales, 2) if net is not None else None,
        "target_monthly_sales": target_sales,
        "first_order_qty": int(params.get("first_order_qty", 100) or 100),
        "asins": [{
            "asin": c["asin"], "title": c["title"], "brand": c["brand"],
            "gender": c.get("gender", MISSING),
            "price": c["price"], "monthly_sales": c["monthly_sales"],
            "monthly_revenue": c["monthly_revenue"], "reviews": c["reviews"],
            "rating": c["rating"], "bsr": c["bsr"], "image": c["image"],
            "launch_date": c["launch_date"], "seller": c["seller"], "share": c["share"],
            "seller_country": c.get("seller_country", MISSING),
            "fulfillment": c.get("fulfillment", MISSING),
            "is_china_seller": c.get("is_china_seller"),
            "url": c["url"], "source": c.get("_source", MISSING),
        } for c in comps],
        "data_gaps": gaps,
        "source_files": source_files,
    }
    return candidate, comps


def build_report_data(datasets, log, params):
    competitors_raw = datasets.get("competitor", [])
    keywords_raw = datasets.get("keyword_miner", [])
    kw_trends = datasets.get("keyword_trend", [])
    asin_trends = datasets.get("asin_sales", [])

    keyword_to_market = {}
    for r in keywords_raw:
        k = safe_str(r.get("keyword"))
        m = safe_str(r.get("market"))
        if k and m:
            # 一个关键词可映射多个市场（如 靴下→カーフソックス+スポーツソックス）
            if k in keyword_to_market:
                prev = keyword_to_market[k]
                prev = prev if isinstance(prev, list) else [prev]
                if m not in prev:
                    prev.append(m)
                keyword_to_market[k] = prev
            else:
                keyword_to_market[k] = m
    asin_to_market = {}
    for r in competitors_raw:
        a = norm_asin(r.get("asin"))
        m = safe_str(r.get("market")) or safe_str(r.get("keyword")) or "默认市场"
        if a:
            asin_to_market[a] = m

    markets = set(asin_to_market.values())
    all_gaps = []
    candidates, all_comps = [], []
    for market in sorted(markets):
        comps = []
        for r in competitors_raw:
            m = safe_str(r.get("market")) or safe_str(r.get("keyword")) or "默认市场"
            if m == market:
                comps.append(r)
        res = analyze_market(market, comps, asin_to_market, keyword_to_market,
                             asin_trends, kw_trends, keywords_raw, params)
        if res:
            cand, comps_clean = res
            candidates.append(cand)
            all_comps.extend(comps_clean)
            all_gaps.extend(cand["data_gaps"])
    candidates.sort(key=lambda x: x["total_score"], reverse=True)

    kw_rows = []
    for r in keywords_raw:
        kw_rows.append({
            "keyword": safe_str(r.get("keyword")) or MISSING,
            "market": safe_str(r.get("market")) or MISSING,
            "search_volume": num(r.get("search_volume")),
            "purchase_volume": num(r.get("purchase_volume")),
            "competition": safe_str(r.get("competition")) or MISSING,
            "source": r.get("_source", MISSING),
        })

    if not candidates:
        all_gaps.append("没有可用的竞品 ASIN 数据：请导出 competitor_asins.xlsx 后重跑")
    if not kw_trends and not asin_trends:
        all_gaps.append("没有逐月趋势数据：无法判断全年稳定性")

    target = int(params.get("monthly_sales_target", 50) or 50)
    priority = [c for c in candidates if c["grade"].startswith("A")][:int(params.get("output", {}).get("priority_n", 5))]
    return {
        "generated_at": datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds"),
        "params": params,
        "summary": {
            "market_count": len(candidates),
            "candidate_count": len(candidates),
            "priority_count": len(priority),
            "avg_score": round(statistics.mean([c["total_score"] for c in candidates]), 1) if candidates else 0,
            "target_monthly_sales": target,
            "target_net_profit": params.get("min_net_profit_jpy"),
            "estimated_monthly_net": round(sum(c["monthly_net_profit"] or 0 for c in priority), 2),
        },
        "candidates": candidates,
        "competitors": all_comps,
        "keywords": kw_rows,
        "collection_log": log,
        "data_gaps": all_gaps,
    }


HEADERS = {
    "候选产品": ["市场", "产品", "等级", "总分", "稳定性", "竞争", "利润", "差异化", "风险", "均价JPY", "建议售价JPY", "单件净利JPY", "目标月销", "预计月净利JPY", "首批采购", "季节性", "月度CV", "商品CR3", "商品CR5", "商品HHI", "品牌CR3", "品牌CR5", "品牌HHI", "卖家CR3", "卖家CR5", "卖家HHI", "竞品数", "竞品ASIN", "数据来源"],
    "市场规模分析": ["市场", "竞品数", "样本月销量合计", "样本月销售额合计", "平均售价JPY", "中位售价JPY", "中位评论数", "品牌数", "卖家数", "月销30-100竞品数", "低评论竞品数"],
    "市场份额分析": ["市场", "商品CR3", "商品CR5", "商品CR10/TOP10", "商品HHI", "品牌CR3", "品牌CR5", "品牌HHI", "卖家CR3", "卖家CR5", "卖家HHI", "低评论占比", "样本说明"],
    "关键词数据": ["关键词", "市场", "月搜索量", "月购买量", "竞争度", "来源文件"],
    "竞品ASIN分析": ["市场", "ASIN", "商品标题", "性别", "品牌", "卖家", "卖家所属地", "中国卖家", "配送方式", "售价JPY", "月销量", "月销售额JPY", "评论数", "星级", "BSR", "样本份额", "上架月份", "商品链接", "来源文件"],
    "季节性分析": ["市场", "数据来源", "计算窗口", "有效月份数", "月均", "月度CV", "峰谷比", "稳定月数", "判定", "窗口外月份(排除)", "月度明细"],
    "逐月趋势明细": ["市场", "月份", "数值", "数据来源"],
    "价格与利润分析": ["市场", "样本均价JPY", "建议售价JPY", "佣金JPY", "FBA费JPY", "采购JPY", "头程JPY", "广告预留JPY", "退货预留JPY", "其他JPY", "单件净利JPY", "净利率", "目标月销", "预计月净利JPY"],
    "产品综合评分": ["市场", "目标季节稳定25", "竞争难度25", "利润空间20", "差异化20", "库存退货风险10", "合计100", "等级"],
    "数据采集记录": ["文件", "识别类型", "行数", "表头行", "修改时间", "错误"],
}


def write_workbook(data, path: Path):
    wb = Workbook()
    wb.remove(wb.active)

    def add(name, rows):
        ws = wb.create_sheet(name[:31])
        headers = HEADERS.get(name, [])
        if headers:
            ws.append(headers)
        for r in rows:
            ws.append(r)
        if headers:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for cell in ws[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="2F5597")
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for col in ws.columns:
            width = max((len(safe_str(c.value)) for c in col if c.value is not None), default=8)
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(width + 2, 10), 42)

    cand_rows = []
    for c in data["candidates"]:
        m = c["market_share"]
        cand_rows.append([
            c["market"], c["name"], c["grade"], c["total_score"],
            c["scores"]["stability"], c["scores"]["competition"], c["scores"]["profit"],
            c["scores"]["differentiation"], c["scores"]["risk"],
            m["avg_price"], c["suggested_price"], c["net_profit"], c["target_monthly_sales"],
            c["monthly_net_profit"], c["first_order_qty"], c["seasonality_label"],
            (c["seasonality"] or {}).get("cv"),
            m["product_cr3"], m["product_cr5"], m["product_hhi"],
            m["brand_cr3"], m["brand_cr5"], m["brand_hhi"],
            m["seller_cr3"], m["seller_cr5"], m["seller_hhi"],
            m["competitor_count"],
            ", ".join(a["asin"] for a in c["asins"][:10]),
            ", ".join(c["source_files"]),
        ])
    add("候选产品", cand_rows)

    add("市场规模分析", [[m["market"], m["competitor_count"], m["total_monthly_sales"], m["total_monthly_revenue"],
                           m["avg_price"], m["median_price"], m["median_reviews"], m["brand_count"],
                           m["seller_count"], m["mid_sales_count"], m["low_review_count"]]
                        for m in [c["market_share"] for c in data["candidates"]]])
    add("市场份额分析", [[m["market"], m["product_cr3"], m["product_cr5"], m["product_cr10"], m["product_hhi"],
                          m["brand_cr3"], m["brand_cr5"], m["brand_hhi"],
                          m["seller_cr3"], m["seller_cr5"], m["seller_hhi"], m["low_review_ratio"],
                          "样本估算；卖家集中度仅基于有卖家名称的样本"]
                        for m in [c["market_share"] for c in data["candidates"]]])
    add("关键词数据", [[k["keyword"], k["market"], k["search_volume"], k["purchase_volume"], k["competition"], k["source"]]
                     for k in data["keywords"]])
    def _cn_label(v):
        return "是" if v is True else ("否" if v is False else "Unknown")
    add("竞品ASIN分析", [[c["market"], c["asin"], c["title"], c.get("gender", MISSING), c["brand"], c.get("seller", MISSING),
                          c.get("seller_country", MISSING), _cn_label(c.get("is_china_seller")),
                          c.get("fulfillment", MISSING),
                          c["price"], c["monthly_sales"], c["monthly_revenue"], c["reviews"], c["rating"], c["bsr"],
                          c["share"], c["launch_date"], c["url"], c.get("_source", MISSING)] for c in data["competitors"]])
    add("季节性分析", [[c["market"], (c["seasonality"] or {}).get("source", MISSING),
                       (c["seasonality"] or {}).get("window_label", MISSING),
                       (c["seasonality"] or {}).get("months_available", 0), (c["seasonality"] or {}).get("mean"),
                       (c["seasonality"] or {}).get("cv"), (c["seasonality"] or {}).get("peak_trough"),
                       (c["seasonality"] or {}).get("stable_months"), c["seasonality_label"],
                       "; ".join(f"{x['month']}:{x['value']}" for x in (c["seasonality"] or {}).get("excluded_months", [])) or "无",
                       "; ".join(f"{x['month']}:{x['value']}" for x in (c["seasonality"] or {}).get("months", []))]
                      for c in data["candidates"]])
    monthly_rows = []
    for c in data["candidates"]:
        sea = c["seasonality"] or {}
        for x in sea.get("months", []):
            monthly_rows.append([c["market"], x["month"], x["value"], sea.get("source", MISSING)])
    add("逐月趋势明细", monthly_rows)
    add("价格与利润分析", [[c["market"], c["market_share"]["avg_price"], c["suggested_price"],
                          (c["profit"] or {}).get("佣金"), (c["profit"] or {}).get("FBA配送费"),
                          (c["profit"] or {}).get("采购成本"), (c["profit"] or {}).get("头程物流"),
                          (c["profit"] or {}).get("广告预留"), (c["profit"] or {}).get("退货预留"),
                          (c["profit"] or {}).get("其他"), c["net_profit"], (c["profit"] or {}).get("净利率"),
                          c["target_monthly_sales"], c["monthly_net_profit"]] for c in data["candidates"]])
    add("产品综合评分", [[c["market"], c["scores"]["stability"], c["scores"]["competition"], c["scores"]["profit"],
                         c["scores"]["differentiation"], c["scores"]["risk"], c["total_score"], c["grade"]]
                        for c in data["candidates"]])
    add("TOP20推荐产品", cand_rows[:20])
    add("TOP5优先开发产品", cand_rows[:5])
    add("数据采集记录", [[l.get("file", MISSING), l.get("type", MISSING), l.get("rows", 0), l.get("header_row", 0),
                         l.get("modified", MISSING), l.get("error", "")] for l in data["collection_log"]])

    ws = wb.create_sheet("数据缺口")
    ws.append(["缺口说明"])
    for g in data["data_gaps"]:
        ws.append([g])
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="C00000")
    ws.column_dimensions["A"].width = 90

    wb.save(path)


def write_markdown(data, path: Path):
    p = data["params"]
    s = data["summary"]
    lines = [
        "# 亚马逊日本站服装精铺选品报告",
        "",
        f"生成时间：{data['generated_at']}  ",
        f"目标：售价 ¥{fmt(p.get('price_min'))}–¥{fmt(p.get('price_max'))}，月销 {fmt(s['target_monthly_sales'])} 件，"
        f"单件净利 ≥ ¥{fmt(s['target_net_profit'])}，首批 {fmt(p.get('first_order_qty'))} 件/款。",
        "",
        "## 结论",
        "",
        f"- 候选市场/产品组合：**{s['candidate_count']}** 个",
        f"- 优先开发（A级）：**{s['priority_count']}** 个",
        f"- 平均评分：{s['avg_score']} / 100",
        f"- 优先组合预计月净利：¥{fmt(s['estimated_monthly_net'])}（按目标月销测算，未扣固定运营成本）",
        "",
        "## TOP 5 优先开发",
        "",
        "| # | 市场/产品 | 评分 | 等级 | 均价 | 建议售价 | 单件净利 | 目标月销 | 预计月净利 | CV | 商品CR10 | 品牌CR3 | 卖家CR3 |",
        "|---|---|---:|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|",
    ]
    for i, c in enumerate(data["candidates"][:5], 1):
        sea = c["seasonality"] or {}
        ms = c["market_share"]
        lines.append(
            f"| {i} | {c['name']}（{c['market']}） | {c['total_score']} | {c['grade']} | ¥{fmt(ms['avg_price'])} | ¥{fmt(c['suggested_price'])} | "
            f"¥{fmt(c['net_profit'])} | {c['target_monthly_sales']} | ¥{fmt(c['monthly_net_profit'])} | "
            f"{fmt(sea.get('cv'), 2)} | {fmt(ms['product_cr10'], 2)} | {fmt(ms['brand_cr3'], 2)} | {fmt(ms['seller_cr3'], 2)} |"
        )
    lines += ["", "## 逐款细节", ""]
    for i, c in enumerate(data["candidates"][:5], 1):
        ms = c["market_share"]
        lines += [
            f"### {i}. {c['name']}（{c['market']}） · {c['total_score']}分 · {c['grade']}",
            "",
            f"- 评分拆解：目标季节稳定 {c['scores']['stability']}/25，竞争 {c['scores']['competition']}/25，"
            f"利润 {c['scores']['profit']}/20，差异化 {c['scores']['differentiation']}/20，风险 {c['scores']['risk']}/10",
            f"- 季节性（窗口：{(c['seasonality'] or {}).get('window_label', MISSING)}）：{c['seasonality_label']}，"
            f"窗口内有效月份 {fmt((c['seasonality'] or {}).get('months_available'))}，"
            f"月度CV {fmt((c['seasonality'] or {}).get('cv'), 2)}，峰谷比 {fmt((c['seasonality'] or {}).get('peak_trough'), 2)}",
            f"- 样本市场：竞品 {ms['competitor_count']} 个，月销量合计 {fmt(ms['total_monthly_sales'])}",
            f"- 平均价格：¥{fmt(ms['avg_price'])}（中位 ¥{fmt(ms['median_price'])}）",
            f"- 集中度：商品 CR3 {fmt(ms['product_cr3'], 2)} / CR5 {fmt(ms['product_cr5'], 2)} / HHI {fmt(ms['product_hhi'], 0)}；"
            f"品牌 CR3 {fmt(ms['brand_cr3'], 2)} / CR5 {fmt(ms['brand_cr5'], 2)} / HHI {fmt(ms['brand_hhi'], 0)}；"
            f"卖家 CR3 {fmt(ms['seller_cr3'], 2)} / CR5 {fmt(ms['seller_cr5'], 2)} / HHI {fmt(ms['seller_hhi'], 0)}",
            f"- 利润：售价 ¥{fmt(c['suggested_price'])} − 佣金 ¥{fmt((c['profit'] or {}).get('佣金'))} − FBA ¥{fmt((c['profit'] or {}).get('FBA配送费'))} "
            f"− 采购 ¥{fmt((c['profit'] or {}).get('采购成本'))} − 头程 ¥{fmt((c['profit'] or {}).get('头程物流'))} "
            f"− 广告 ¥{fmt((c['profit'] or {}).get('广告预留'))} − 退货 ¥{fmt((c['profit'] or {}).get('退货预留'))} = **¥{fmt(c['net_profit'])}/件**",
            f"- 首批采购：{c['first_order_qty']} 件；月销目标 {c['target_monthly_sales']} 件；预计月净利 ¥{fmt(c['monthly_net_profit'])}",
            "- 数据来源：" + (", ".join(c["source_files"]) if c["source_files"] else MISSING),
            "- 逐月趋势（每月）：",
            "| 月份 | 数值 |",
            "|---|---:|",
        ]
        for x in (c["seasonality"] or {}).get("months", []):
            lines.append(f"| {x['month']} | {fmt(x['value'])} |")
        lines += [
            "",
            "| ASIN | 商品 | 品牌 | 售价 | 月销量 | 评论 | 样本份额 | 链接 |",
            "|---|---|---|---:|---:|---:|---:|---|",
        ]
        for a in c["asins"][:10]:
            title = a["title"].replace("|", "\\|")[:48]
            lines.append(
                f"| {a['asin']} | {title} | {a['brand']} | ¥{fmt(a['price'])} | {fmt(a['monthly_sales'])} | "
                f"{fmt(a['reviews'])} | {fmt(a['share'], 2)} | [打开]({a['url']}) |"
            )
        if c["data_gaps"]:
            lines.append("- 数据缺口：" + "；".join(c["data_gaps"]))
        lines.append("")

    _, season_label = resolve_season_window(p)
    lines += [
        "## 方法与边界",
        "",
        f"- 季节性计算窗口：**{season_label}**（参数 seasonality={p.get('seasonality')}）。"
        "窗口外月份（如冬季12/1/2月）不填0、不参与 CV 与评分，仅在「季节性分析」表留档披露。",
        "- 月度 CV = 总体标准差 / 均值，只基于窗口内有效月份。CV ≤20% 优秀，20–30% 可考虑，>30% 降分，>50% 强季节；"
        f"窗口内有效月份少于 {max(3, round(int(p.get('min_months_for_seasonality', 6)) * len(resolve_season_window(p)[0]) / 12))} 个月按“数据不足”。",
        "- 市场占比基于卖家精灵导出的**样本竞品**，销量为估算值。报告中的份额是样本份额，不能当作日本站全市场份额。",
        "- 评分：目标季节稳定25 / 竞争25 / 利润20 / 差异化20 / 库存退货风险10；A≥80，B≥70，C≥60，<60 淘汰。",
        "- 月销 50 件是目标而非硬淘汰线；月销 30 件但竞争小、利润稳、目标季节窗口内有需求的商品仍保留为潜力候选。",
        "- FBA 费、佣金率来自参数模板，**采购前必须按 Amazon 日本站最新费率核对**。",
        "- 父子变体销量不得重复累加；报告以父 ASIN 去重。",
        "",
        "## 数据缺口清单",
        "",
    ]
    lines += [f"- {g}" for g in data["data_gaps"]] or ["- 无"]
    lines += ["", "## 原始导出文件", ""]
    lines += [f"- {l.get('file', MISSING)} → {l.get('type', MISSING)}，{l.get('rows', 0)} 行" for l in data["collection_log"]]
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def check_only(datasets, log, params):
    counts = {k: len(v) for k, v in datasets.items() if k != "unknown"}
    print("=== 数据完整性检查 ===")
    for k, v in counts.items():
        print(f"  {k:16s} {v:5d} 行")
    required = {"competitor": "竞品 ASIN", "keyword_miner": "关键词挖掘", "keyword_trend": "关键词逐月趋势"}
    missing = [label for key, label in required.items() if counts.get(key, 0) == 0]
    if counts.get("asin_sales", 0) == 0:
        print("  ! asin_sales 为 0：季节性将退化到关键词趋势（可接受，但精度下降）")
    if missing:
        print("  缺失必需数据：" + "、".join(missing))
        return 1
    print("  必需数据齐全，可以运行分析。")
    return 0


def main():
    ap = argparse.ArgumentParser(description="日本站服装精铺选品分析")
    ap.add_argument("--raw", type=Path, default=ROOT / "raw")
    ap.add_argument("--params", type=Path)
    ap.add_argument("--out", type=Path, default=ROOT / "output")
    ap.add_argument("--check-only", action="store_true")
    args = ap.parse_args()

    params = load_params(args.params)
    if not args.raw.exists():
        sys.exit(f"raw 目录不存在：{args.raw}")
    datasets, log = load_all(args.raw)
    if args.check_only:
        return check_only(datasets, log, params)

    data = build_report_data(datasets, log, params)
    args.out.mkdir(parents=True, exist_ok=True)
    (args.out / "report_data.json").write_text(
        json.dumps(data, ensure_ascii=False, indent=2, default=str), encoding="utf-8"
    )
    write_workbook(data, args.out / "apparel-research.xlsx")
    write_markdown(data, args.out / "research-report.md")
    print(f"候选 {data['summary']['candidate_count']} 个，优先开发 {data['summary']['priority_count']} 个")
    print(f"输出：{args.out / 'apparel-research.xlsx'}")
    print(f"      {args.out / 'research-report.md'}")
    print(f"      {args.out / 'report_data.json'}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
