# 卖家精灵数据采集指南

## 站点与登录

- 用 `https://www.sellersprite.com`（国际站）。`sellersprite.cn` 在本机 SSL 握手失败。
- 日本站 `marketId=6`。
- 账号密码不写进脚本、日志或报告。用户手动在有头浏览器登录一次：

```bash
agent-browser open "https://www.sellersprite.com" --headed
```

登录态只在同一次浏览器会话内有效。采完再 `agent-browser close`。

本机如出现 `ERR_TUNNEL_CONNECTION_FAILED` / 502，读系统代理覆盖环境变量：

```bash
scutil --proxy | grep -E 'HTTPS?Port|HTTPS?Proxy'
export HTTP_PROXY=http://127.0.0.1:7897 HTTPS_PROXY=http://127.0.0.1:7897
export http_proxy=$HTTP_PROXY https_proxy=$HTTPS_PROXY
```

## 两类导出，机制不同

| 类型 | 例子 | 怎么拿文件 |
|---|---|---|
| 异步任务 | 关键词挖掘、选产品/竞品 ASIN | 点导出 → 打开 `https://www.sellersprite.com/v2/export-log` → 复制 `https://o.sellersprite.com/batch-exports/...xlsx` 直链 → `curl -L -o raw/xxx.xlsx <直链>`。直链无鉴权 |
| 即时下载 | 关键词逐月趋势、ASIN 逐月销量 | 点导出后浏览器直接下载。必须用 `agent-browser download @<ref> raw/xxx.xlsx`；`click` 或 JS `.click()` 不会保存到指定路径 |

`ref` 每次页面变化都会变，批量循环里必须重新 `snapshot` 取。

## 四张表的采集动作

### 1. 关键词挖掘 → `raw/keyword_miner.xlsx`

```bash
Q=$(python3 -c "import urllib.parse;print(urllib.parse.quote('メンズ ノーカラーシャツ'))")
agent-browser open "https://www.sellersprite.com/v3/keyword-miner/?q=${Q}&marketId=6&batch=0"
sleep 7
# 批量导出 → export-log 取直链 → curl
```

核心列：关键词、市场、月搜索量、月购买量、竞争度。

### 2. 关键词逐月趋势 → `raw/keyword_trend_<关键词>.xlsx`

```bash
# 在关键词挖掘结果里点行尾趋势图标
agent-browser eval "(()=>{const a=[...document.querySelectorAll('a')].find(x=>x.querySelector('.icon-google-trends-icon'));if(!a)return 0;a.click();return 1})()"
sleep 3
# 切「搜索趋势」tab（默认是 Google Trends）
agent-browser eval "(()=>{const t=[...document.querySelectorAll('[role=tab]')].find(x=>x.innerText.includes('搜索趋势'));if(t)t.click();return 1})()"
sleep 3
# 点「按月」
agent-browser eval "(()=>{const b=[...document.querySelectorAll('button')].find(x=>x.innerText.trim()==='按月'&&x.offsetParent);if(b)b.click();return 1})()"
sleep 2
REF=$(agent-browser snapshot | grep 'link.*导出' | grep -oE 'e[0-9]+' | head -1)
agent-browser download "@$REF" raw/keyword_trend_xxx.xlsx
```

输出约 59 行（表头 + 58 个月）。核心列：月份、关键词、月搜索量、月购买量。

### 3. 竞品 ASIN → `raw/competitor_asins.xlsx`

```bash
agent-browser open "https://www.sellersprite.com/v3/competitor-lookup"
sleep 5
# 从「选产品」按品类/价格带/月销区间筛，然后导出
```

核心列：ASIN、商品、品牌、卖家名称、价格、月销量、月销售额、评论数、星级、类目、市场、图片链接、上架时间、变体。要算卖家集中度就必须导出「卖家名称」列。

市场列写细分市场名（如「男士无领衬衫」）；没有这一列时脚本会用文件名或关键词兜底。

### 4. ASIN 逐月销量 → `raw/asin_sales_trend_<ASIN>.xlsx`

```bash
agent-browser open "https://www.sellersprite.com/v3/competitor-lookup"
# 输入 ASIN → 立即查询 → 点行内历史趋势图标
agent-browser eval "(()=>{const el=[...document.querySelectorAll('.icon-historical-trend')].find(e=>e.offsetParent);if(!el)return 0;el.click();return 1})()"
sleep 4
# 关键：切到「销量趋势」tab，不是「历史趋势」
REF=$(agent-browser snapshot | grep '导出产品' | grep -oE 'e[0-9]+' | head -1)
agent-browser download "@$REF" raw/asin_sales_trend_B0XXXXXXXX.xlsx
```

抽屉里的「历史趋势」tab 导出的是每日价格/BSR 变动，不是逐月销量。逐月销量只能用「销量趋势」tab 的「导出产品」。

## 脚本可识别的列名

`analyze.py` 用别名映射，不要求列名完全一致。常见别名：

| 规范字段 | 可接受列名 |
|---|---|
| ASIN | ASIN / 商品ASIN / 父ASIN / 父(子)体ASIN |
| 月销量 | 月销量 / 销量 / 预估月销量 / monthly sales |
| 价格 | 价格 / 售价 / 平均单价 / BuyBox价格 |
| 评论数 | 评论数 / 评论数量 / reviews |
| 月份 | 月份 / 月 / date / 日期 |
| 关键词 | 关键词 / keyword / 搜索词 |
| 市场 | 市场 / 分析市场 / 品类 / 细分市场 / 赛道 |

表头不在第 1 行也能识别（脚本会扫描前 12 行找表头）。缺列不会崩，会写进「数据缺口」。

## 采集清单

- [ ] `keyword_miner.xlsx`（必须）
- [ ] `keyword_trend_*.xlsx`（必须，至少覆盖候选关键词）
- [ ] `competitor_asins.xlsx`（必须，含真实 ASIN 和价格）
- [ ] `asin_sales_trend_*.xlsx`（建议，至少 TOP 竞品）
- [ ] `asin_keywords_*.xlsx`（建议）
- [ ] 每个文件记录采集时间（`templates/data-collection-log.md`）

## 数据边界

- 卖家精灵销量/销售额是估算值，不是后台真实订单。
- 竞品历史销量覆盖 1–27 个月不等，很多 ASIN 只有 2–3 个月，只能辅助佐证。全年判定以关键词逐月趋势为主。
- 有 ASIN 导出过万月销量属异常，脚本会标记 `outlier_sales_max` 并排除。
- 30 天快照不能证明全年。
