---
name: 日本站服装精铺选品
description: 亚马逊日本站服装精铺选品 Skill。卖家精灵数据可复用已采集文件或重新采集，做市场占比、目标季节窗口（默认春夏秋3-11月，可配置全年）季节性、利润与 100 分评分，输出 Excel + Markdown + HTML（含真实 ASIN 与 Amazon.co.jp 链接）。当用户说"日本站服装选品""服装精铺""卖家精灵分析""春夏秋/三季选品""全年稳定的服装""月销50单也可以""验证全年可卖""导出历史趋势""用已有数据重跑""重新采集"时使用。每次运行先问数据源：复用已采集数据 or 重新采集。已合并原 sss-amazon-jp-apparel-research 与 sellersprite-history-export。
agent_created: true
---

# 亚马逊日本站服装精铺选品专家

## 这个 Skill 解决什么

不是找爆款。找**目标季节窗口内有需求、淡旺季差异小、竞争低、月销 30–100 件也能赚钱**的服装精铺机会。

默认窗口为**春夏秋 3–11 月**（排除冬季 12/1/2 月），可由 `seasonality` 参数改为全年 —— 冬季不卖但经营季稳卖的产品允许进入候选。旧版"全年稳定"口径改用 `strict_year_round`。

输出三份报告：

| 文件 | 用途 |
|---|---|
| `output/index.html` | 主交付。离线可打开的交互选品页：顶部**市场概览看板**（18 张指标卡，随选中产品联动；含「头部商品 前3/5/10/20名」下拉控制商品/品牌/卖家集中度），下方候选产品表（评分/销量/利润/CV 筛选排序），点开看真实 ASIN 和 Amazon.co.jp 链接 |
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

### 0. 数据源选择（每次运行先问，必须交互）

**执行前第一问**：用 `AskUserQuestion` 询问数据源，不静默决定。

| 选项 | 含义 | 跳过什么 | 仍执行什么 |
|---|---|---|---|
| **复用已采集数据**（推荐） | 用 `raw/` 和 `raw-products/` 里现有文件重跑分析 | 跳过 §1 前置、§3 采集 | 参数询问、预处理、完整性检查、分析渲染、交付 |
| **重新采集** | 走卖家精灵重新导出全部原始数据 | 无 | 全部步骤 |

选「复用」前先检查数据是否存在并告知用户采集时间：

```bash
ls -la raw/competitor_asins.xlsx raw/keyword_miner*.xlsx raw/keyword_trend_*_monthly.xlsx 2>/dev/null
ls -la raw-products/product_export_*.xlsx 2>/dev/null
```

- 若 `raw/competitor_asins.xlsx` 或关键词文件缺失 → **不得静默复用**，提示用户改选「重新采集」或补齐文件。
- 复用档也要重跑预处理：`price_min`/`price_max`/`new_product_months` 变了会影响过滤结果（见 §4）。若本轮参数与上轮完全相同且用户只是要看结果，可跳过预处理直接用现有 `raw/competitor_asins.xlsx`。
- 用户本轮说「照上次」「用已有数据」「重跑」→ 直接按复用档，不必再问；说「重新采集」「重新获取」「更新数据」→ 重新采集档。
- 采集时间超过 **1 个月** 时提示用户：数据可能过期，是否仍复用。

### 1. 前置（仅「重新采集」档执行；复用档跳过，不要开浏览器）

1. 先加载 `agent-browser` Skill（浏览器操作必须用真实浏览器）。
2. 代理：**只在首次确认环境时读一次** `scutil --proxy | grep -E 'HTTPS?Port|HTTPS?Proxy'`。系统代理已全局配置，浏览器会自动用。
   **不要在 `open --headed` 的同一条命令里 `export HTTP_PROXY=...`**——这会触发 agent-browser daemon 用 `--proxy` 重启（端口变成 57970），Cookie 清空、用户白登录。2026-11-20 实测踩过两次。
3. 站点用 `sellersprite.com`（国际站）。`sellersprite.cn` 在本机 SSL 握手失败，别用。
4. 账号密码不要问、不要存。打开有头浏览器让用户手动登录：

```bash
agent-browser open "https://www.sellersprite.com" --headed
```

5. **用 `AskUserQuestion` 给登录选择框**，不要只发文字让用户回"已登录"：

```
问题："在卖家精灵窗口完成登录（含验证码）后，点下面的按钮继续"
选项：["已登录，继续采集", "登录失败/被踢回登录页", "取消本次采集"]
```

6. 用户选「已登录」后**立即自动验证**，不要信文字确认：

```bash
agent-browser eval "location.href"    # 必须含 /v2/welcome 或 /v2/me /v3/...，不能是 /user/login 或 about:blank
```

   - 若还在登录页 → 告诉用户"登录态丢失（可能浏览器被重启），请重新登录"，再弹一次选择框。
   - 若进 welcome → 按 Escape 关安全风险弹窗，推广弹窗用 `document.getElementById('kolspriteTrialCloseBtn').click()`。
   - **采完不要 `agent-browser close`**。保留浏览器和登录态给下一次对话复用，省掉重复登录。仅在用户明确说"关闭浏览器"时才 `close`。

**会话复用（下一次对话）**：开始采集前先探测现有 daemon 是否还活着且已登录：

```bash
agent-browser eval "location.href" 2>/dev/null
```

   - 返回 `/v2/welcome` / `/v2/me` / `/v3/...` → 登录态还在，跳过登录，直接进 §2 参数询问。
   - 返回登录页 / `about:blank` / 命令报错（daemon 已退出）→ 重新 `open --headed` 走上面的登录选择框。
   - 注意：daemon 可能被系统清理或用户手动关闭，复用失败是正常情况，不要报错退出，重新登录即可。

### 2. 收集用户参数（执行时必须交互询问）

复制模板到 `output/params.json`：

```bash
cp templates/user-params.json output/params.json
```

**执行前必须用 `AskUserQuestion` 主动询问下列可调项**（用户已明确给出的跳过，其余用默认值）。不要静默按默认值跑——用户要能选/输入：

| 序号 | 参数 | 默认值 | 让用户选/输入 | 写入 params 键 |
|---|---|---|---|---|
| Q1 | **低评论门槛** | ≤50 | ≤30 / ≤50 / ≤100 / 不限 | `low_review_max`（不限填 99999） |
| Q2 | **售价带** | 2000–3000 | 1000–2000 / 2000–3000 / 3000–5000 / 自定义（输入两个值） | `price_min` / `price_max` |
| Q3 | **新品范围** | 不限 | 只看近3月 / 近6月 / 近12月 / 不限 | `new_product_months`（0=不限） |
| Q4 | **列出数量** | 20 | 20 / 50 / 100 / 全部 | `output.top_n`（全部填 9999） |
| Q5 | **目标月销** | 50 | 30 / 50 / 100 / 200 / 自定义 | `monthly_sales_target` |
| Q6 | **单件净利目标** | ¥500 | ¥300 / ¥500 / ¥800 / ¥1000 | `min_net_profit_jpy` |
| Q7 | **性别**（5档细分） | 不限 | 女 / 偏女 / 中性 / 偏男 / 男 / 不限（单选，见下） | `gender`：`female`/`female_leaning`/`neutral`/`male_leaning`/`male`/`all` |
| Q8 | **季节窗口** | 春夏秋3-11月 | 春夏秋 / 全年 / 自定义月份 | `seasonality` / `season_window_months` |
| Q9 | **卖家所属地**（含是否中国卖家，同一个筛选项） | 不限 | 不限 / 中国 / 日本 / 其他（输入国家代码）/ 未知(字段空) | `seller_country` (`all`/`CN`/`JP`/国家代码/`unknown`) |
| Q10 | **配送方式**（多选） | 不限 | AMZ(亚马逊自营) / FBA / FBM 三个复选框可任意组合；都不选=不限 | `fulfillment`（`all` 或逗号分隔如 `AMZ,FBA`；**必须用 multiSelect 选择框**） |
| Q11 | **趋势历史月份**（新品期回溯窗口） | 24个月 | 12个月 / 24个月 / 36个月 / 自定义 | `history_months`（填数字，如 `12`） |

**Q11 说明**：控制季节性 CV / 峰谷比计算的回溯范围——从最新月份往前取 N 个月。12个月只看最近一年（对新品/新趋势敏感），24个月覆盖两年（默认，平衡），36个月看长周期。
注意：卖家精灵逐月趋势文件通常自带 24–59 个月，Q11 只是**截取参与评分的窗口**，不影响采集量；`min_months_for_seasonality` 会按 `max(3, round(6×窗口月数/12))` 自动校验样本是否够。

其余参数（`competition` / `categories` / `return_risk` / 成本项）用模板默认，用户主动提再改。

> **卖家精灵 UI 实测（用户截图 2026-11-19 + 导出核对）**：
> - 配送方式是 **AMZ / FBA / FBM 三个独立复选框**（不是我之前写的"只有 FBM/AMZ、FBA 映射 AMZ"）。**Q10 必须用 `AskUserQuestion` 的 `multiSelect: true` 选择框**，让用户一次勾选多个；选中几个就传逗号分隔值，如 `FBA,FBM`；一个都不选=不限（`all`）。不要做成单选框。
> - 「是否为中国卖家」与「卖家所属地」在卖家精灵里是**同一个筛选项**（就是「卖家所属地」输入框），所以本 Skill 也合并成 Q9 一个参数，不拆成两个。
> - 选中国/日本时：`is_china_seller` 综合 `卖家所属地` + `卖家信息`（地址含 CN/中国/China）判断，因为所属地字段常为空；信号不足标 `Unknown`，不猜测。选其他国家代码则按所属地字段精确匹配。

> **脚本已支持**：Q1/Q2/Q5/Q6/Q7/Q8/Q9/Q10/Q11 由 `analyze.py` 读取 params 生效；Q9/Q10 在竞品层过滤，并在 Excel「竞品ASIN分析」与 HTML 详情卡片显示卖家所属地/中国卖家/配送方式。
> Q2 售价带同时驱动 `prep_competitor.py`（读 `output/params.json`，未设置则用 2000–3000）。
> Q3 新品范围在 `prep_competitor.py` 按上架时间过滤（需导出含「上架时间」列，缺失则跳过并记数据缺口）。
> Q4 列出数量由 `render_report.py` 读取 `output.top_n`，控制"候选产品"主表显示行数；详情内竞品 ASIN 永远支持翻页/显示全部。
> Q11 `history_months` 由 `analyze.py` 的 `seasonality_for_market()` 读取（`months_back`），只影响参与 CV 计算的月份回溯，不改采集量。
> Q7 性别 5 档由 `prep_competitor.py` 的 `gender_label()` 判定并写入「性别」列；`analyze.py` 透传到 Excel/HTML。选偏/中性档位时采集关键词要去掉性别词才能采到对应品（见文末性别段）。

**每次执行都问，即使用户上一轮给过值**——除非用户本轮说"照上次"。问完把结果写回 `output/params.json` 再进入采集。

### 3. 采集数据（卖家精灵，仅「重新采集」档执行；复用档跳过）

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
- **历史趋势是即时下载**：必须 `agent-browser download @<ref> <绝对路径>`，用 `click` 或 JS `.click()` 不会落盘。ref 每次页面变化都会变，批量循环里重新取。
- **download 路径不要带 `.xlsx` 后缀**：agent-browser 会自动补一层，传 `xxx.xlsx` 会落成 `xxx.xlsx.xlsx`（2026-11-20 实测）。传 `/abs/path/keyword_trend_xxx` 即可，或下载后 `mv xxx.xlsx.xlsx xxx.xlsx`。
- 日本站 `marketId=6`。
- 竞品历史抽屉里“历史趋势”tab 导出的是每日价格/BSR，不是逐月销量；逐月销量要“销量趋势”tab 的“导出产品”。
- **整批采完不要 `agent-browser close`**，保留浏览器和登录态给下一次对话复用。只有用户明确要求关闭时才 `close`。

**采集后必须核对清单**（不核对就进预处理会静默漏数据）：

```bash
ls -la raw/competitor_asins.xlsx raw/keyword_miner*.xlsx 2>/dev/null
ls -la raw/keyword_trend_*_monthly.xlsx 2>/dev/null | wc -l   # 趋势文件数
ls raw/*.xlsx.xlsx 2>/dev/null && echo "有双后缀文件，需重命名"   # 应为 0 行
```

三者任一为空 → 回到对应采集步骤补齐，不要带着缺口往下跑。

### 4. 数据预处理（两档都执行：重新采集档必须；复用档见下）

把 `raw-products/` 的按月导出整理成 `raw/competitor_asins.xlsx`，并从关键词导出抽 monthly sheet：

```bash
$PY scripts/prep_competitor.py         # raw-products/product_export_YYYY-MM_*.xlsx → raw/competitor_asins.xlsx
$PY scripts/extract_monthly.py         # keyword_trend_*.xlsx → keyword_trend_*_monthly.xlsx（自动清多层冗余）
$PY scripts/patch_keyword_market.py    # 自动给 keyword_miner 补「市场」列，关联逐月趋势到细分市场
cat raw/market_mapping_report.txt      # 看映射明细 + 未匹配的趋势关键词（数据缺口）
```

- 脚本读 `output/params.json`：`price_min`/`price_max`/`new_product_months` 参数在这里生效，所以**参数变了必须重跑预处理**，否则竞品样本还是旧价格带。
- `patch_keyword_market.py` **必须在 analyze 之前跑**：没有「市场」列 → 该市场季节性恒为「数据不足」，稳定性只拿 8/25。脚本按关键词文本保守映射到竞品小类目（处理词序、空格、性别前缀、中点），映射不上的写进 `market_mapping_report.txt`，不编造。原 miner 备份到 `raw-products/_archive_prev/`。
- `extract_monthly.py` 已自动跳过 `*_monthly.xlsx` 并把残留的 `*_monthly_monthly*.xlsx` 移到 `raw-products/_archive_prev/`，不用再手工清理。
- 复用档若本轮参数与上一轮完全一致（尤其价格带、新品范围没变），可跳过本步，直接用现有 `raw/competitor_asins.xlsx`。只要有一个筛选参数不同就必须跑。
- 预处理规则详见文末「数据预处理惯例」：跨季快照门槛、剔非服装/冬品、按父 ASIN 去重。

### 5. 数据完整性检查

```bash
$PY scripts/analyze.py --raw raw --params output/params.json --check-only
```

缺数据会列出来。**缺什么补什么，不允许 AI 编造。** 检查通过后进入分析。

### 6. 分析与渲染

```bash
$PY scripts/analyze.py --raw raw --params output/params.json --out output
$PY scripts/render_report.py --data output/report_data.json --out output/index.html
```

`analyze.py` 会同时产出 `apparel-research.xlsx`、`research-report.md`、`report_data.json`。`render_report.py` 只读 JSON，改 HTML 样式时不用重跑分析。

`$PY` 用受管 Python：`/Users/shishugen/.workbuddy/binaries/python/envs/default/bin/python`（已装 openpyxl）。其他环境用 `python3` 并 `pip install openpyxl`。

### 7. 交付

**先复制到工作区再交付**（skill 的 `output/` 在 `~/.workbuddy/skills/` 下，用户在当前工作区看不到）：

```bash
cp output/index.html output/apparel-research.xlsx output/research-report.md <工作区>/output/
```

`<工作区>` = 会话的 Workspace Folder（见 user context）。然后 `present_files` 打开工作区副本，顺序：`index.html` → `apparel-research.xlsx` → `research-report.md`。

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
| 出来的品全是冬季货（パーカー/コート/裏起毛），和春夏秋目标不符 | 选产品默认「最近 30 天」= 当前月份。改「选择月份」为目标窗口两端月份（如 2026-05 + 2026-09），导出后在 `prep_competitor.py` 里走跨季门槛，详见下文「竞品必须用跨季快照」 |
| 选产品页 CSS 里给输入框 set value + 派发 input 事件，点「开始筛选」返回“暂无结果” | Element-UI 只认真实键盘事件。`eval` 里 `input.focus()`，再用 `agent-browser keyboard type` 逐字打，两框之间 `agent-browser press Tab`；JS 赋值不进入 Vue 状态 |
| 类目树勾了「レディース/メンズ」（2 级）后筛选结果为 0 | 该树只有勾到**叶子类目**才会命中商品；2 级节点返回 0 条。改用「包含关键词」筛标题（如 `レディース,メンズ`，多个英文逗号分隔），再在 `prep_competitor.py` 里按「类目路径」二次过滤服装 |
| 勾完类目点「确认选择」后页面标签没变化 | 复选框 click 与「确认选择」click 必须在**两次独立 eval** 之间隔 3 秒，同一次 eval 里连着点会读不到最新勾选状态 |
| 导出的竞品一半是同一父体的几十个子体，去重后只剩 200 个 | 选产品导出上限 2000 行且按月销降序截断，头部长变体 listing 会把名额吃满。对策：加「变体数 0–10」筛选 + 多导几批（不同月销区间/不同关键词）+ `prep_competitor.py` 按父 ASIN 去重 |
| 关键词趋势导出文件里有 5 个 sheet（weekly/monthly/aba） | 新版导出一次给全。`analyze.py` 只读第一 sheet（weekly），用 `scripts/extract_monthly.py` 抽出 `JP-monthly-*` 另存再跑 |
| 原始导出（product_export_*.xlsx）放 raw/ 会被 analyze.py 当成竞品混进来 | `load_all` 会 rglob 扫 raw/ 下所有 xlsx。原始大表统一放 `raw-products/`，`raw/` 只留规范化后的 `competitor_asins.xlsx` + 关键词文件 + `*_monthly.xlsx` |
| `agent-browser download @ref /path/xxx.xlsx` 报成功但找不到文件 | 目标路径若已含 `.xlsx`，agent-browser 会再补一层变成 `xxx.xlsx.xlsx`。传路径时**不要带后缀**，或下载后 `mv xxx.xlsx.xlsx xxx.xlsx`。批量下载后务必 `ls` 核对真实文件名 |
| 重新登录时浏览器被 relaunch、登录态丢失（代理从 7897 变成 57970） | 见 §1 第 2/6 步：登录阶段不要带代理 export；用选择框确认后必 `eval location.href` 验登录页。已固化到主流程 |
| 报告全淘汰、用户怀疑 skill 有 bug | 先逐项核 scores 再下结论。稳定8=无逐月趋势、4=CV>0.5；竞争被 CR10>0.55(-8)+中位评论>200(-7)+HHI>2500(-4/-3)叠加扣到个位数；利润16-20满分。全淘汰多为真实数据+口径组合的结果。三季口径**救不了夏peak冬谷品**（T恤三季内CV仍0.74-1.09）。真正有效的是补采缺失市场的逐月趋势并补 keyword_miner「市场」列映射（趋势词大小写/空格须精确匹配，如「メンズ tシャツ」小写t） |
| `extract_monthly.py` 反复生成 `*_monthly_monthly*.xlsx` 冗余文件 | 2026-11-20 已修：脚本现在跳过 `*_monthly.xlsx` 并在开头自动把残留多层文件移到 `raw-products/_archive_prev/`。若仍出现，说明用的是工作区旧副本，改回 skill 目录脚本 |

## 数据预处理惯例（实测有效）

### 关键：竞品必须用「跨季快照」，不能用最近 30 天

「最近 30 天」= 当前月份。11 月导出的样本全是 パーカー/コート/裏起毛/長袖パジャマ，直接跑出来就是一堆冬品，和目标季节窗口矛盾。**季节性参数只影响评分窗口，不会改变竞品样本本身。**

正确做法（2026-11-18 实测）：
1. 选产品页「选择月份」选 **目标窗口两端的月份**，例如春夏秋窗口就取 `2026-05`（春末夏初）+ `2026-09`（初秋）。
2. 各导出一次，文件名带上月份：`product_export_2026-05_gender.xlsx`、`product_export_2026-09_gender.xlsx`。
3. `scripts/prep_competitor.py` 会做**跨季门槛**：同一细分市场必须在两个月快照里各 ≥4 个竞品、合计 ≥8 个才入样；同时按标题剔除硬冬品（裏起毛/ボア/フリース/中綿/防寒/あったか/厚手/ダウン/マフラー/手袋/ブーツ/ニット帽子/セーター）和小类目级冬品。
4. 同一商品两个月都出现时，输出合并成一行（保留最近月份数值，`快照月份` 列记 `2026-05,2026-09`），避免同款被算两次销量。

其余惯例：

导出的原始 xlsx 常不能直接喂给 analyze.py，需在 `raw/` 里生成规范化的 `competitor_asins.xlsx`。
2026-11-18 已把整套流程写成脚本，直接跑即可：

```bash
$PY scripts/prep_competitor.py     # raw-products/product_export_YYYY-MM_*.xlsx → raw/competitor_asins.xlsx
$PY scripts/extract_monthly.py     # keyword_trend_*.xlsx → keyword_trend_*_monthly.xlsx（只留 monthly sheet）
```

`prep_competitor.py` 做的事：只留「类目路径含 服＆ファッション小物」（**并剔除非服装品类**）→ 价格 2000–3000、月销 30–500 → 标题级剔冬品 → **按父 ASIN 去重**（保留月销最高的子体）→ 市场 = 小类目首行 → 跨季门槛。

> **非服装剔除（重要）**：日本站把「服＆ファッション小物」当大类，里面混了配饰/袜/雨具/鞋包。2026-11-19 实测：24 个"市场"里 9 个不是服装（帽子109行、太阳镜48行、袜子73行、腰带26行、雨衣10行，合计占 46%）。
> 用 `NON_APPAREL_SEGMENTS`（路径段**精确匹配**）+ `NON_APPAREL_SEG_CONTAINS`（段名含 ソックス/帽子/サングラス/ベルト/ポンチョ 等）剔除。
> **不能用全路径子串**：`服＆ファッション小物` 会被 `ファッション小物` 误伤。剔除后市场从 24→15，全部为真服装。

- 只保留样本量足够的市场（≥8 竞品且 ≥5 个有效月销），避免小样本虚高分；
- 市场列 = 小类目（清理 CR/LF，取第一行，避免两个小类目名被拼成一串）；
- 关键词挖掘文件的「市场」列（2026-11-20 已脚本化）：`analyze.py` 的 `seasonality_for_market()` 靠 `keyword_to_market`（关键词挖掘行的「市场」列）把逐月趋势关联到细分市场；**没有市场列 → 该市场季节性恒为「数据不足」，稳定性只得 8/25 保底分**。当前 analyze 已禁止「全关键词求和」的全局回退，所以旧版"留空可让趋势作为大盘共享"的说法已失效。
  **现在直接跑 `scripts/patch_keyword_market.py` 自动补**，不要再手工 openpyxl。脚本按关键词文本映射到竞品小类目（处理词序不同如「パジャマ メンズ」vs「メンズパジャマ」、性别前缀粘连、空格、中点 `･`；内置少量确定同义词如 ボクサーパンツ↔ボクサー）。映射不上的趋势词写进 `raw/market_mapping_report.txt`，留空不猜。见 2026-11-20 男女不限实测（9市场映射成功，pants 类因竞品表无对应小类目留缺口）。

### 性别 5 档细分（gender，2026-11-20 改版）
`prep_competitor.py` 的 `gender_label(path, title)` 给**每条竞品**打标签，写入 competitor_asins.xlsx 的「性别」列，并在 Excel「竞品ASIN分析」与 HTML 详情卡片显示：

| 标签 | 判定规则 | params 值 |
|---|---|---|
| 女 | 类目路径含 レディース（强信号） | `female` |
| 男 | 类目路径含 メンズ（强信号） | `male` |
| 偏女 | 类目路径无性别词，但标题含 レディース/女性/ウィメンズ | `female_leaning` |
| 偏男 | 类目路径无性别词，但标题含 メンズ/男性 | `male_leaning` |
| 中性 | 类目路径和标题都无性别词 | `neutral`（`unisex` 兼容旧值） |

- `gender=all` 不过滤，但仍给每条打标签（供报告看性别构成）。指定档位则只保留该标签，日志打印「剔性别不符 N 条」。
- **关键**：当前采集若用 `keywords=レディース,メンズ` 筛标题，类目路径必带性别词 → 只会产出「女/男」，看不到偏/中性。要采到中性品，Q7 选「中性/偏女/偏男」时，采集关键词改用服装通用词（如 `Tシャツ,パンツ,ショーツ`）不要带性别词，性别交给预处理按标题判定。
- 选产品页用标题关键词筛选仍会混入配饰；非服装剔除在预处理按类目路径做，比标题关键词可靠。

### 选产品页关键词框/导出状态（2026-11-19 实测）
- 「请输入关键词，多个以英文逗号区分」有**两个**：第一个是「排除关键词」(URL 参数 `outOfKeywords`)，第二个才是「包含关键词」(URL `keywords`)。填错框会混入食品/雑貨。填完必须按 Enter 让值进入 URL 参数；判断依据直接看 URL 里的 `keywords=...`。
- 异步导出状态会**缓存**：JS 点 refresh 或重复点 ref 都拿不到最新状态。必须 `agent-browser open "https://www.sellersprite.com/v2/export-log"` **整页重新加载**，状态变「已完成」后行内才出现 `a.download-excel` 直链（`o.sellersprite.com/batch-exports/...xlsx`，curl 下载）。
- `agent-browser download <ref> <path>` 的相对路径会被解析到**当前工作区 cwd**，不是 skill 目录。必须传绝对路径。
- 备份文件（如 `_archive_competitor_male.xlsx`）绝不能留在 `raw/` 下——`analyze.load_all` 用 rglob 扫全部 xlsx，备份会被当数据读入。备份移到 `raw-products/` 或其子目录。
