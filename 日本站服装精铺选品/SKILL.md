---
name: 日本站服装精铺选品
description: 亚马逊日本站服装精铺选品 Skill。用户提供卖家精灵账号并手动登录后，Skill 通过浏览器导出关键词、竞品 ASIN、逐月趋势数据，做市场占比、目标季节窗口（默认春夏秋3-11月，可配置全年）季节性、利润与 100 分评分，输出 Excel + Markdown + HTML（含真实 ASIN 与 Amazon.co.jp 链接）。当用户说"日本站服装选品""服装精铺""卖家精灵分析""春夏秋/三季选品""全年稳定的服装""月销50单也可以""验证全年可卖""导出历史趋势"时使用。已合并原 sss-amazon-jp-apparel-research 与 sellersprite-history-export。
agent_created: true
---

# 亚马逊日本站服装精铺选品专家

## 这个 Skill 解决什么

不是找爆款。找**目标季节窗口内有需求、淡旺季差异小、竞争低、月销 30–100 件也能赚钱**的服装精铺机会。

默认窗口为**春夏秋 3–11 月**（排除冬季 12/1/2 月），可由 `seasonality` 参数改为全年 —— 冬季不卖但经营季稳卖的产品允许进入候选。旧版"全年稳定"口径改用 `strict_year_round`。

输出三份报告：

| 文件 | 用途 |
|---|---|
| `output/index.html` | 主交付。离线可打开的交互选品页，按评分/销量/利润/CV 筛选排序，点产品看真实 ASIN 和 Amazon.co.jp 链接 |
| `output/apparel-research.xlsx` | 原始数据、市场份额、季节性、利润、评分、TOP20/TOP5 全部工作表 |
| `output/research-report.md` | 文字决策报告：为什么推荐、风险、采购建议、数据缺口 |

## 目录约定

```text
ssg-amazon-jp-apparel-niche-research/
├── SKILL.md
├── references/                # 分析口径，不要改脚本时凭感觉改这里
│   ├── methodology-redlines.md    # 数据真实性红线、样本边界、Parent/Child、中国卖家、服装风险、12项检查
│   ├── seasonality-analysis.md    # 目标季节窗口口径
│   └── sellersprite-data-guide.md # 采集动作与常见坑
├── scripts/
│   ├── analyze.py             # 读 raw/ → 分析 → xlsx + report_data.json
│   ├── render_report.py       # report_data.json → index.html
│   ├── make_sample_data.py    # 无账号时造模拟数据，验证流程
│   └── init_templates.py      # 生成空白采集模板 xlsx
├── templates/
│   ├── user-params.json       # 用户选品参数（每次运行复制一份改）
│   ├── raw-data-schema.md
│   ├── data-collection-log.md
│   └── product-research.xlsx  # 空白采集模板
├── raw/                       # 卖家精灵原始导出，只读，不要手工改
└── output/                    # 每次运行覆盖这里
```

## 执行流程

以下命令默认在 Skill 根目录执行。先进入目录并设好 Python：

```bash
cd /Users/shishugen/.workbuddy/skills/ssg-amazon-jp-apparel-niche-research
PY=/Users/shishugen/.workbuddy/binaries/python/envs/default/bin/python
```

若用的是工作区副本，把 `cd` 路径换成工作区里的 `ssg-amazon-jp-apparel-niche-research`。

### 0. 前置

1. 先加载 `agent-browser` Skill（浏览器操作必须用真实浏览器）。
2. 本机访问 ChatGPT / 海外站要走系统代理。默认环境变量里的代理可能返回 502，先读系统代理：

```bash
scutil --proxy | grep -E 'HTTPS?Port|HTTPS?Proxy'
# 若输出 7897：
export HTTP_PROXY=http://127.0.0.1:7897 HTTPS_PROXY=http://127.0.0.1:7897
export http_proxy=$HTTP_PROXY https_proxy=$HTTPS_PROXY
```

3. 站点用 `sellersprite.com`（国际站）。`sellersprite.cn` 在本机 SSL 握手失败，别用。
4. 账号密码不要问、不要存。让用户在有头浏览器里手动登录一次：

```bash
agent-browser open "https://www.sellersprite.com" --headed
# 用户完成登录/验证码后，后续步骤不要 close，登录态只在同一次会话有效
```

### 1. 收集用户参数

复制模板到 `output/params.json`，只问用户**没给的关键项**，其余用默认值：

```bash
cp templates/user-params.json output/params.json
```

| 参数 | 默认 | 说明 |
|---|---|---|
| `price_min` / `price_max` | 2000 / 3000 | 日元 |
| `monthly_sales_target` | 50 | 目标月销；30–100 都算可接受，不是硬淘汰线 |
| `min_net_profit_jpy` | 500 | 单件净利目标 |
| `first_order_qty` | 100 | 首批采购，单款 |
| `competition` | `low` | `low` / `medium` / `any` |
| `seasonality` | `three_seasons_spring_summer_autumn` | `three_seasons_spring_summer_autumn`(春夏秋3-11月，默认) / `strict_year_round`(全年) / `custom`(配 `season_window_months`)；旧值 stable/strict/flexible 回退全年 |
| `gender` | `all` | `all` / `men` / `women` / `unisex` |
| `categories` | `[]` | 空 = 全品类 |
| `return_risk` | `low` | `low` / `medium` / `any` |
| `history_months` | 24 | 12 或 24；24 更容易判季节性 |
| 成本项 | 见模板 | 采购、头程、FBA、佣金率、广告率、退货预留 |

**用户改过的值才问；模板已有默认值时不要反复确认。**

### 2. 采集数据（卖家精灵）

按 `references/sellersprite-data-guide.md` 执行。最小采集集：

| raw/ 文件名 | 内容 | 优先级 |
|---|---|---|
| `keyword_miner.xlsx` | 关键词、搜索量、购买量、竞争度 | 必须 |
| `keyword_trend_*.xlsx` | 单个关键词逐月搜索/购买趋势 | 必须 |
| `competitor_asins.xlsx` | 竞品 ASIN、价格、月销量、评论、品牌、类目、图片 | 必须 |
| `asin_sales_trend_*.xlsx` | 单 ASIN 逐月销量/销售额 | 建议 |
| `asin_keywords_*.xlsx` | ASIN 关键词反查 | 建议 |

关键坑（详见 reference）：

- 关键词挖掘 / 选产品是**异步导出**：点导出 → 打开 `https://www.sellersprite.com/v2/export-log` → 取 `o.sellersprite.com/batch-exports/...xlsx` 直链 → `curl -L -o raw/xxx.xlsx <直链>`。直链无鉴权。
- **历史趋势是即时下载**：必须 `agent-browser download @<ref> raw/xxx.xlsx`，用 `click` 或 JS `.click()` 不会落盘。ref 每次页面变化都会变，批量循环里重新取。
- 日本站 `marketId=6`。
- 竞品历史抽屉里“历史趋势”tab 导出的是每日价格/BSR，不是逐月销量；逐月销量要“销量趋势”tab 的“导出产品”。
- 整批采完再 `agent-browser close`。中间不要关。

### 3. 数据完整性检查

```bash
$PY scripts/analyze.py --raw raw --params output/params.json --check-only
```

缺数据会列出来。**缺什么补什么，不允许 AI 编造。** 检查通过后进入分析。

### 4. 分析与渲染

```bash
$PY scripts/analyze.py --raw raw --params output/params.json --out output
$PY scripts/render_report.py --data output/report_data.json --out output/index.html
```

`analyze.py` 会同时产出 `apparel-research.xlsx`、`research-report.md`、`report_data.json`。`render_report.py` 只读 JSON，改 HTML 样式时不用重跑分析。

`$PY` 用受管 Python：`/Users/shishugen/.workbuddy/binaries/python/envs/default/bin/python`（已装 openpyxl）。其他环境用 `python3` 并 `pip install openpyxl`。

### 5. 交付

用 `present_files` 打开这三个文件，顺序：`output/index.html` → `output/apparel-research.xlsx` → `output/research-report.md`。

口头只说结论：几个候选、几个优先开发、目标月销/净利、最大风险。不要复述整份报告。

## 分析口径（摘要，详见 references/）

出报告前**必须通读 `references/methodology-redlines.md` 并执行其中 12 项真实性检查**；数据真实性优先于凑齐候选数量。季节性窗口口径见 `references/seasonality-analysis.md`。

- **季节性**：先按 `seasonality` 确定窗口（默认春夏秋3–11月），月度 CV = 总体标准差 / 均值、峰谷比 = 最高月 / 最低月，**都只基于窗口内月份**。窗口外（冬季）不填0、不参与 CV 与评分，仅在「季节性分析」表的「窗口外月份」列留档披露。CV ≤ 20% 优秀，20–30% 可考虑，>30% 降分，>50% 强季节。窗口内有效月份少于 `max(3, round(6×窗口月数/12))`（三季=4）标“数据不足”。若销量几乎全靠窗口内单一月份（如8月水着），即便 CV 低也标季节性并降权。
- **逐月趋势**：HTML 详情、Markdown 报告和 Excel「逐月趋势明细」都逐月列出，不合并、不省略。
- **集中度**：商品（ASIN CR3/CR5/CR10 + HHI）、品牌（CR3/CR5 + HHI）、卖家（CR3/CR5 + HHI）三组；另给平均价格与中位价格。卖家集中度需要导出里有「卖家名称」列，缺则标缺失，不拿品牌冒充。
- **市场占比**：TOP10 / TOP20 销量占比、月销 30–100 件竞品数、低评论竞品数。卖家精灵是**估算值**，报告必须写“样本市场份额”，不能说成日本站全市场。
- **利润**：售价 − 佣金 − FBA − 采购 − 头程 − 广告 − 退货预留。费率全部来自 `params.json`，FBA 费率需用户按 Amazon 最新标准核对。
- **评分（100 分）**：目标季节稳定 25（窗口=全年时即原"全年稳定"）/ 竞争难度 25 / 利润 20 / 差异化 20 / 库存退货风险 10。A 级 80–100 优先开发，B 级 70–79 可测试，C 级 60–69 观察，<60 淘汰。
- **父子变体**：不能把子 ASIN 销量重复累加到父 ASIN；有变体数据时按父 ASIN 去重。

## 数据真实性红线

1. 每个 ASIN、销量、销售额、市场占比都必须来自 `raw/` 里的具体文件，报告记录来源文件名和采集时间。
2. 没采集到的字段写 `缺失`，不要填 0、不要估算、不要 AI 补造。
3. 模拟数据只用于流程验证，不能出现在交付报告里。`make_sample_data.py` 输出到 `/tmp`，不要写进 `raw/`。
4. 用户手动登录，账号密码不进入任何脚本、日志或报告。

## 无账号时怎么验证

```bash
$PY scripts/make_sample_data.py --out /tmp/jp-apparel-demo
$PY scripts/analyze.py --raw /tmp/jp-apparel-demo/raw \
  --params templates/user-params.json --out /tmp/jp-apparel-demo/output
$PY scripts/render_report.py --data /tmp/jp-apparel-demo/output/report_data.json \
  --out /tmp/jp-apparel-demo/output/index.html
```

看到 HTML 里有候选产品、ASIN 可点击跳转即算跑通。

## 常见故障

| 现象 | 处理 |
|---|---|
| `ERR_TUNNEL_CONNECTION_FAILED` / 502 | 用 `scutil --proxy` 读端口覆盖环境变量；别直接取消代理直连 |
| 点导出后 raw/ 没文件 | 即时下载型必须用 `agent-browser download @ref <path>` |
| snapshot 里 ref 失效 | 每轮操作后重新 `snapshot` 取 ref |
| 历史趋势只有 2–3 个月 | 只作佐证；全年判定以关键词逐月趋势为主 |
| 某 ASIN 月销百万级 | 异常值，标记并排除，写进数据缺口 |
| 用户说“月销 50 单太少” | 说明这是精铺不是爆款；30 单但竞争小、利润好仍保留为潜力候选 |
| `--check-only` 显示 `keyword_trend 0 行`，但文件明明下好了 | 趋势 xlsx 第一 sheet 是 `JP-weekly-*`，analyze.py 只读第一 sheet，判成 keyword_miner。用 openpyxl 提取 `JP-monthly-*` sheet 另存为只含 monthly 的单 sheet 文件（表头须含「月份/关键词/月搜索量」），再重跑 |
| 报告里价格/利润全「缺失」、市场全归「默认市场」 | 竞品导出列名带单位（如「价格(円)」）、且无「市场」列。预处理 competitor_asins.xlsx：价格(円)→价格、评分数→评论数、BuyBox卖家→卖家名称、月销售额(円)→月销售额，新增「市场」列（取「小类目」并清理换行） |
| 只有 1 个竞品的市场拿 A 级 89 分（竞争满分） | 小样本导致 CR/评论缺失→不扣分→虚高。预处理时只保留「样本≥8 且月销量>0 的样本≥5」的市场；否则报告的优先开发建议不可信 |
| 选产品页填了类目/价格/月销，点「开始筛选」结果却混入食品等无关类目 | 该表单提交不稳定。优先改用账号「我的导出」里的历史导出（直链 curl 下载），或走关键词挖掘主路径；不要在这张表单上反复耗时间 |
| 登录成功后弹「账号安全风险 / 绑定微信」swal2 弹窗 | 按 `agent-browser press Escape` 关闭；推广弹窗用 `document.getElementById('kolspriteTrialCloseBtn').click()`；不要用 `agent-browser close`（会丢登录态） |

## 数据预处理惯例（实测有效）

导出的原始 xlsx 常不能直接喂给 analyze.py，需在 `raw/` 里生成规范化的 `competitor_asins.xlsx`：
- 只保留样本量足够的市场（如 ≥8 竞品且 ≥5 个有效月销），避免小样本虚高分；
- 市场列 = 小类目（清理 CR/LF），这样每个细分市场独立评分；
- 关键词挖掘文件**不要**加「市场」列：留空可让 6 个逐月趋势作为「服装大盘季节性」共享给所有市场（报告需标注这是大盘参考，非该细分专属）。
