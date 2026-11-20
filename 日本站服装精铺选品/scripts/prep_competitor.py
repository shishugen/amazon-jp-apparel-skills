#!/usr/bin/env python3
"""把卖家精灵「选产品」按月快照导出整理成 analyze.py 可用的 raw/competitor_asins.xlsx。

核心目的：只保留**春夏秋能卖**的细分市场，而不是 11 月冬季快照里的一堆冬品。

做法：
1. 只读文件名里带月份的文件：`product_export_YYYY-MM_*.xlsx`（快照月份）
2. 只留「类目路径含 服＆ファッション小物」的行（服装本体，排除鞋/包/家纺/雨具）
3. 价格 2000–3000 日元、月销量 30–500
4. 硬冬品关键词剔除（裏起毛/ボア/フリース/ダウン/防寒…）
5. 按父 ASIN 去重（同一父体只留月销最高的子体）
6. **跨季门槛**：同一细分市场必须在 ≥2 个不同月份快照里都出现，且每月 ≥4 个竞品、合计 ≥8 个
7. 市场 = 小类目首行（清理换行）
"""
import glob
import json
import os
import re
import collections
from datetime import date
import openpyxl

BASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..')
RAW = os.path.join(BASE, 'raw')
SRC = os.path.join(BASE, 'raw-products')
PARAMS_PATH = os.path.join(BASE, 'output', 'params.json')

# 默认值；执行时由 output/params.json 覆盖（用户交互询问后写入）
PRICE_MIN, PRICE_MAX = 2000, 3000
SALES_MIN, SALES_MAX = 30, 500
NEW_PRODUCT_MONTHS = 0   # 0=不限；>0 只保留上架至今 ≤N 个月的新品
GENDER = 'all'           # all / female(女) / male(男) / female_leaning(偏女) / male_leaning(偏男) / neutral(中性) / unisex(中性,旧值兼容)
APPAREL_KEY = '服＆ファッション小物'
MIN_PER_MONTH = 4      # 每个快照月份至少几个竞品
MIN_TOTAL = 8          # 合计至少几个竞品


def load_params():
    """读 output/params.json 的可选过滤项，文件不存在则用默认。"""
    global PRICE_MIN, PRICE_MAX, NEW_PRODUCT_MONTHS, GENDER
    if not os.path.exists(PARAMS_PATH):
        print('未找到 output/params.json，用默认过滤：价格 2000–3000、新品不限')
        return
    try:
        p = json.load(open(PARAMS_PATH, encoding='utf-8'))
    except Exception as e:
        print(f'params.json 读取失败（{e}），用默认过滤')
        return
    PRICE_MIN = int(p.get('price_min', PRICE_MIN))
    PRICE_MAX = int(p.get('price_max', PRICE_MAX))
    NEW_PRODUCT_MONTHS = int(p.get('new_product_months', 0) or 0)
    GENDER = str(p.get('gender', 'all') or 'all').lower()
    print(f'params 过滤：价格 {PRICE_MIN}–{PRICE_MAX}、新品范围 '
          f'{"近"+str(NEW_PRODUCT_MONTHS)+"个月" if NEW_PRODUCT_MONTHS else "不限"}'
          f'、性别 {GENDER}')


def is_new_enough(launch, ref: date) -> bool:
    """上架时间距今 ≤ NEW_PRODUCT_MONTHS 个月；0=不限；无上架时间时不限（记缺口由 analyze 处理）。"""
    if not NEW_PRODUCT_MONTHS or not launch:
        return True
    ym = re.search(r'(\d{4})[-/](\d{1,2})', str(launch))
    if not ym:
        return True
    y, m = int(ym.group(1)), int(ym.group(2))
    if not (1 <= m <= 12):
        return True
    months = (ref.year - y) * 12 + (ref.month - m)
    return months <= NEW_PRODUCT_MONTHS

# 非服装品类：日本站把「服＆ファッション小物」当大类，里面混了配饰/袜/雨具。
# 用类目路径的【段】精确匹配（不能子串，否则 '服＆ファッション小物' 会被 'ファッション小物' 误伤）。
NON_APPAREL_SEGMENTS = [
    'ファッション小物',   # 帽子/太阳镜/腰带/围巾/手套/领带等配饰
    '靴下・レッグウェア', # 袜类
    '靴下', 'ソックス', 'レッグウェア',
    'レインウェア',       # 雨具（レインコート・ポンチョ）
    'シューズ', '靴', 'スニーカー', 'サンダル',
    'バッグ', 'かばん', '財布', 'リュック',
    'アクセサリー', 'ジュエリー', '腕時計', '時計',
    'スカーフ', 'マフラー', '手袋', 'ネクタイ',
]
# 段名里"含"这些词也算非服装（覆盖 スポーツソックス 这类复合段名）；仍按段判断，不做全路径子串
NON_APPAREL_SEG_CONTAINS = ['ソックス', '靴下', 'サングラス', '帽子', 'ベルト', 'レインコート', 'ポンチョ']

# 硬冬品：8 分以上的冬装信号，出现在标题里就剔除
WINTER_TERMS = [
    '裏起毛', '起毛', 'ボア', 'フリース', '中綿', 'キルティング', 'ムートン',
    'ダウンジャケット', 'ダウンコート', 'ダウンベスト',
    'マフラー', '手袋', 'ネックウォーマー', 'イヤーマフ',
    '防寒', 'あったか', '暖か', '極暖', 'ヒートテック', '保温',
    '厚手', '冬用', '真冬', 'スノー', 'ブーツ', 'ウィンター',
    'ニット帽子', 'セーター', 'タートルネック',
]

COLMAP = {
    'ASIN': 'ASIN',
    '商品标题': '商品标题',
    '品牌': '品牌',
    'BuyBox卖家': '卖家名称',
    '价格(円)': '价格',
    '月销量': '月销量',
    '月销售额(円)': '月销售额',
    '评分数': '评论数',
    '评分': '评分',
    '小类BSR': 'BSR',
    '上架时间': '上架时间',
    '小类目': '_小类目',
    '类目路径': '_类目路径',
    '商品主图': '图片链接',
    '父ASIN': '父ASIN',
    '卖家所属地': '卖家所属地',
    '配送方式': '配送方式',
    'BuyBox类型': 'BuyBox类型',
    'A+页面': 'A+页面',
    '视频介绍': '视频介绍',
    'SP广告': 'SP广告',
}

OUT_HEAD = ['市场', 'ASIN', '商品标题', '性别', '品牌', '卖家名称', '价格', '月销量', '月销售额',
            '评论数', '评分', 'BSR', '上架时间', '类目', '图片链接', '卖家所属地',
            '配送方式', 'BuyBox类型', 'A+页面', '视频介绍', 'SP广告',
            '快照月份', '来源文件']

# 性别标签：5 档，按类目路径(强信号)+标题关键词(弱信号)判定
# 女 / 偏女 / 中性 / 偏男 / 男
GENDER_LABELS = ['女', '偏女', '中性', '偏男', '男']
F_WORDS = ('レディース', '女性', 'ウィメンズ', 'ladies', 'women', 'womens')
M_WORDS = ('メンズ', '男性', 'men', 'mens')


def gender_label(path, title):
    """按标题性别词强弱给商品打性别标签。
    类目路径含明确性别 → 女/男（强信号，覆盖标题）；
    类目路径无性别 → 标题含女性词=偏女、含男性词=偏男、都没有=中性。
    """
    p = str(path or '').lower()
    t = str(title or '').lower()
    if 'レディース' in p:
        return '女'
    if 'メンズ' in p:
        return '男'
    has_f = any(w in t for w in F_WORDS)
    has_m = any(w in t for w in M_WORDS)
    if has_f and not has_m:
        return '偏女'
    if has_m and not has_f:
        return '偏男'
    return '中性'


def fnum(v):
    if v is None or v == '':
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def snapshot_month(name):
    m = re.search(r'(\d{4})-(\d{2})', name)
    return f'{m.group(1)}-{m.group(2)}' if m else None


def is_winter(title):
    t = str(title or '')
    return any(w in t for w in WINTER_TERMS)


def main():
    load_params()
    files = sorted(glob.glob(os.path.join(SRC, 'product_export_*.xlsx')))
    files = [f for f in files if snapshot_month(os.path.basename(f))]
    if not files:
        raise SystemExit('raw-products/ 里没有带月份的导出（形如 product_export_2026-05_xxx.xlsx）')

    rows = {}
    for f in files:
        month = snapshot_month(os.path.basename(f))
        wb = openpyxl.load_workbook(f, read_only=True)
        ws = wb.worksheets[0]
        it = ws.iter_rows(values_only=True)
        hdr = [str(h).strip() if h else '' for h in next(it)]
        idx = {h: i for i, h in enumerate(hdr) if h}
        n_in = n_app = n_win = n_nonapparel = n_gender = 0
        for r in it:
            if not r or not r[idx.get('ASIN', 0)]:
                continue
            n_in += 1
            path = str(r[idx.get('类目路径', -1)] or '')
            if APPAREL_KEY not in path or '傘' in path:
                continue
            # 性别过滤：先打 5 档标签，再按用户选择过滤（兼容旧值 female/male/unisex）
            label = None
            if GENDER != 'all':
                label = gender_label(path, r[idx.get('商品标题', -1)])
                want = {
                    'female': '女', 'male': '男',
                    'female_leaning': '偏女', 'male_leaning': '偏男',
                    'neutral': '中性', 'unisex': '中性',
                }.get(GENDER)
                if want and label != want:
                    n_gender += 1
                    continue
            # 非服装品类（配饰/袜/鞋/包/雨具等）：按路径段精确匹配剔除
            segs = [s.strip() for s in path.replace('\n', '').split(':')]
            if (any(seg in NON_APPAREL_SEGMENTS for seg in segs)
                    or any(w in seg for seg in segs for w in NON_APPAREL_SEG_CONTAINS)):
                n_nonapparel += 1
                continue
            n_app += 1
            price, sales = fnum(r[idx.get('价格(円)', -1)]), fnum(r[idx.get('月销量', -1)])
            if price is None or sales is None:
                continue
            if not (PRICE_MIN <= price <= PRICE_MAX) or not (SALES_MIN <= sales <= SALES_MAX):
                continue
            rec = {dst: (r[idx[src]] if src in idx else None) for src, dst in COLMAP.items()}
            # 性别标签：所有竞品都打，供报告/筛选使用（与 GENDER 过滤独立）
            rec['性别'] = gender_label(path, rec.get('商品标题'))
            # 新品范围：以该快照月份为基准算上架月龄，避免用今天年龄误杀历史快照
            if NEW_PRODUCT_MONTHS:
                ym = month.split('-')
                ref = date(int(ym[0]), int(ym[1]), 1)
                if not is_new_enough(rec.get('上架时间'), ref):
                    continue
            if is_winter(rec['商品标题']):
                n_win += 1
                continue
            asin = str(rec['ASIN']).strip()
            market = str(rec['_小类目'] or '').replace('\r', '').split('\n')[0].strip()
            if not asin or not market:
                continue
            if any(w in market for w in ('ニット帽子', 'セーター')):  # 小类目本身就是冬品
                n_win += 1
                continue
            rec['市场'] = market
            rec['类目'] = path.replace('\n', ' ')
            rec['快照月份'] = month
            rec['来源文件'] = os.path.basename(f)
            parent = str(rec.get('父ASIN') or '').strip()
            key = (month, parent if parent and parent != asin else asin)
            prev = rows.get(key)
            if prev and (fnum(prev['月销量']) or 0) >= sales:
                continue
            rows[key] = rec
        print(f'{month} {os.path.basename(f)}: 导出 {n_in} 行，服装 {n_app} 行，'
              f'剔非服装 {n_nonapparel} 条，剔冬品 {n_win} 条'
              + (f'，剔性别不符 {n_gender} 条' if GENDER != 'all' else ''))

    # 跨季门槛
    by_market = collections.defaultdict(lambda: collections.defaultdict(list))
    for rec in rows.values():
        by_market[rec['市场']][rec['快照月份']].append(rec)

    keep = {}
    for m, months in by_market.items():
        ok = [mm for mm, lst in months.items()
              if len(lst) >= MIN_PER_MONTH and sum(1 for x in lst if (fnum(x['月销量']) or 0) > 0) >= MIN_PER_MONTH]
        total = sum(len(months[mm]) for mm in ok)
        if len(ok) >= 2 and total >= MIN_TOTAL:
            keep[m] = {mm: months[mm] for mm in ok}

    # 同一商品可能同时出现在两个月，输出时按商品合并，取最近月份的数值，并记录出现过的月份
    merged = {}
    for m, months in keep.items():
        for mm in sorted(months):  # 月份升序，后写覆盖先写 → 保留最近月份数值
            for rec in months[mm]:
                parent = str(rec.get('父ASIN') or '').strip()
                key = (m, parent if parent and parent != rec['ASIN'] else rec['ASIN'])
                prev = merged.get(key)
                if prev:
                    prev['快照月份'] = f"{prev['快照月份']},{mm}"
                    for k in OUT_HEAD:
                        if k not in ('快照月份',):
                            prev[k] = rec.get(k)
                else:
                    merged[key] = dict(rec)

    final = list(merged.values())
    final.sort(key=lambda x: (x['市场'], -(fnum(x['月销量']) or 0)))

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'competitor_asins'
    ws.append(OUT_HEAD)
    for rec in final:
        ws.append([rec.get(h) for h in OUT_HEAD])
    out = os.path.join(RAW, 'competitor_asins.xlsx')
    wb.save(out)

    months_used = sorted({rec['快照月份'] for rec in final})
    print(f'\n去重后 {len(rows)} 个 ASIN，{len(by_market)} 个细分市场')
    print(f'跨季达标（≥2 个快照月、每月 ≥{MIN_PER_MONTH} 个、合计 ≥{MIN_TOTAL} 个）: {len(keep)} 个市场')
    print(f'输出 {len(final)} 行，快照月份 {months_used} -> {out}')
    for m, months in sorted(keep.items(), key=lambda kv: -sum(len(v) for v in kv[1].values())):
        tot = sum(fnum(x['月销量']) or 0 for lst in months.values() for x in lst)
        dist = ' / '.join(f'{mm}:{len(lst)}' for mm, lst in sorted(months.items()))
        print(f'  {m}: {dist}，样本月销合计 {int(tot)}')


if __name__ == '__main__':
    main()
