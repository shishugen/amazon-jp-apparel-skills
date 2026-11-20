#!/usr/bin/env python3
"""自动给 keyword_miner 补「市场」列，把逐月趋势关键词关联到竞品细分市场。

背景：analyze.py 的 seasonality_for_market() 靠 keyword_to_market（keyword_miner 的
「市场」列）把逐月趋势曲线挂到市场；没有市场列 → 该市场季节性恒为"数据不足"，
稳定性只拿 8/25 保底分。手工补映射容易漏、且大小写/空格/中点不一致会静默失效。

输入（都在 raw/）：
  - competitor_asins.xlsx    竞品表，取「市场」列（小类目名）作为候选市场
  - keyword_trend_*.xlsx     逐月趋势文件，取每个文件里的关键词文本
  - keyword_miner*.xlsx      关键词挖掘表（要补「市场」列的目标）

输出：
  - 原地修改 keyword_miner*.xlsx（备份到 raw-products/_archive_prev/）
  - raw/market_mapping_report.txt  映射明细 + 未能匹配的关键词（数据缺口，不编造）

匹配规则（保守，宁缺勿错）：
  1. 规范化：全角/半角中点统一、去空格、片假名/英文转小写、去常见后缀。
  2. 关键词 norm 后是市场 norm 的子串，或市场 norm 是关键词 norm 的子串 → 映射。
  3. 一个关键词命中多个市场 → 写多行（analyze 会把同关键词多市场读成 list）。
  4. 谁都没匹配上的趋势关键词 → 写进报告，留空，绝不猜。
"""
import sys, glob, os, re, shutil
from pathlib import Path

try:
    import openpyxl
except ImportError:
    sys.exit("缺少 openpyxl，用受管 Python 运行")

BASE = Path(__file__).resolve().parent.parent
RAW = BASE / 'raw'
ARCHIVE = BASE / 'raw-products' / '_archive_prev'


def norm(s):
    """规范化：小写、去空格/中点/标点、统一促音半角。"""
    s = str(s or '').strip().lower()
    s = s.replace('･', '').replace('・', '').replace('·', '').replace('．', '')
    s = re.sub(r'[\s\u3000]', '', s)
    s = re.sub(r'[（）()【】\[\]・､、,，.。\-_/]', '', s)
    return s


# 同义词：关键词常用词 -> 市场小类目里可能的说法（仅收录确定的同物异名，不猜）
SYNONYM = {
    'ボクサーパンツ': ['ボクサー'],
    'ショーツ': ['ショーツ'],
    'トップス': ['tシャツ', 'カットソー'],
    'インナー': ['インナー', 'キャミソール', 'タンクトップ'],
}


def tokens(s):
    """切成语义 token：按空格/中点切；再把性别前缀(レディース/メンズ)从粘连词里拆开。
    例：レディーストップス -> [レディース, トップス]，与市场 レディースTシャツ 可匹配。"""
    s = str(s or '').strip().lower()
    s = s.replace('･', ' ').replace('・', ' ').replace('·', ' ')
    out = []
    for t in re.split(r'[\s\u3000]+', s):
        if not t:
            continue
        m = re.match(r'^(レディース|メンズ)(.+)$', t)
        if m and len(m.group(2)) >= 2:
            out.extend([m.group(1), m.group(2)])
        else:
            out.append(t)
    return out


def match_score(kw, market):
    """关键词与市场匹配的保守打分：所有关键词 token 都能在市场里找到同义/原词 → 1，否则 0。
    处理词序不同（パジャマ メンズ vs メンズパジャマ），不处理语义跳跃。"""
    nk, nm = norm(kw), norm(market)
    if not nk or not nm:
        return 0
    # 先试整串包含（最快路径）
    if nk in nm or nm in nk:
        return 1
    kt, mt = tokens(kw), tokens(market)
    if not kt or not mt:
        return 0
    nm_tokens = [norm(t) for t in mt]
    # 每个关键词 token 都要能在市场 token 里找到自己或同义词
    for t in kt:
        nt = norm(t)
        if not nt:
            continue
        candidates = {nt} | {norm(x) for x in SYNONYM.get(nt, [])}
        # 也允许市场 token 包含该词（如 メンズボクサー 含 ボクサー）
        if not any(any(c in m or m in c for c in candidates) for m in nm_tokens):
            return 0
    return 1


def find_col(header, names):
    h = [norm(x) for x in header]
    for n in names:
        nn = norm(n)
        if nn in h:
            return h.index(nn)
    return None


def read_header(ws):
    return [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]


# 1. 市场清单
comp_file = RAW / 'competitor_asins.xlsx'
if not comp_file.exists():
    sys.exit('缺 raw/competitor_asins.xlsx，请先跑 prep_competitor.py')
wb = openpyxl.load_workbook(comp_file, read_only=True)
ws = wb[wb.sheetnames[0]]
hdr = read_header(ws)
mi = find_col(hdr, ['市场'])
markets = set()
if mi is not None:
    for r in ws.iter_rows(min_row=2, values_only=True):
        if r[mi]:
            markets.add(str(r[mi]).replace('\r', '').split('\n')[0].strip())
markets = sorted(markets)
print(f'市场清单 {len(markets)} 个: {markets}')

# 2. 趋势关键词（从每个 trend 文件的 monthly sheet 读第一行的关键词）
trend_kw = set()
for f in sorted(RAW.glob('keyword_trend_*.xlsx')):
    if '_monthly' in f.name:
        continue
    wb2 = openpyxl.load_workbook(f, read_only=True)
    ms = [w.title for w in wb2.worksheets if 'monthly' in w.title.lower()]
    if not ms:
        continue
    ws2 = wb2[ms[0]]
    h2 = read_header(ws2)
    ki = find_col(h2, ['关键词'])
    if ki is None:
        continue
    for r in ws2.iter_rows(min_row=2, values_only=True):
        if r[ki]:
            trend_kw.add(str(r[ki]).strip())
print(f'趋势关键词 {len(trend_kw)} 个: {sorted(trend_kw)}')

# 3. 建立 关键词 -> [市场] 映射（保守包含匹配）
mapping = {}
unmatched = []
for kw in sorted(trend_kw):
    nk = norm(kw)
    if not nk:
        continue
    hits = [m for m in markets if match_score(kw, m)]
    if hits:
        mapping[kw] = hits
    else:
        unmatched.append(kw)
print(f'可映射 {len(mapping)} 个，未匹配 {len(unmatched)} 个')

# 4. 写回 keyword_miner*.xlsx
targets = sorted(RAW.glob('keyword_miner*.xlsx'))
targets = [f for f in targets if '_archive' not in str(f)]
ARCHIVE.mkdir(parents=True, exist_ok=True)
for f in targets:
    shutil.copy2(f, ARCHIVE / f.name)
    wb3 = openpyxl.load_workbook(f)
    ws3 = wb3[wb3.sheetnames[0]]
    hdr3 = [c.value for c in ws3[1]]
    if '市场' not in [str(x) for x in hdr3]:
        mcol = ws3.max_column + 1
        ws3.cell(1, mcol, '市场')
    else:
        mcol = [str(x) for x in hdr3].index('市场') + 1
    kcol = [str(x) for x in hdr3].index('关键词') + 1
    # 清空旧市场列，重建
    for r in range(2, ws3.max_row + 1):
        ws3.cell(r, mcol, None)
    # 收集已有行 + 关键词位置
    existing = {}
    for r in range(2, ws3.max_row + 1):
        k = str(ws3.cell(r, kcol).value or '').strip()
        if k:
            existing.setdefault(k, []).append(r)
    # 写映射：每个关键词的每个市场一行（首行就地写，其余追加）
    for kw, ms in mapping.items():
        rows = existing.get(kw, [])
        if not rows:
            ws3.append([None] * max(kcol, mcol))
            rows = [ws3.max_row]
            # 修复(2026-11-20)：追加行必须写入关键词，否则该行 keyword 为空，
            # analyze 构建 keyword_to_market 时跳过，导致趋势无法关联到市场。
            ws3.cell(rows[0], kcol, kw)
        for idx, m in enumerate(ms):
            if idx < len(rows):
                ws3.cell(rows[idx], mcol, m)
            else:
                ws3.append([None] * max(kcol, mcol))
                r = ws3.max_row
                ws3.cell(r, kcol, kw)
                ws3.cell(r, mcol, m)
    wb3.save(f)
    print(f'{f.name}: 写入 {sum(len(v) for v in mapping.values())} 个市场标记')

# 5. 报告
report = RAW / 'market_mapping_report.txt'
with open(report, 'w', encoding='utf-8') as fh:
    fh.write('关键词 -> 市场 映射（自动，保守包含匹配）\n')
    fh.write('=' * 50 + '\n')
    for kw, ms in mapping.items():
        fh.write(f'{kw} -> {", ".join(ms)}\n')
    fh.write('\n未能匹配的趋势关键词（数据缺口，未编造）\n')
    fh.write('=' * 50 + '\n')
    for kw in unmatched:
        fh.write(f'{kw}\n')
print(f'映射报告: {report}')
