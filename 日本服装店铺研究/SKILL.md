---
name: 日本服装店铺研究
description: >
  Amazon.co.jp 服装店铺研究 Skill（Store-Centric）。当用户要「找值得研究的服装店铺」「分析店铺经营结构/产品组合/爆款依赖度」「看哪些店铺精铺做得好、里面哪些产品值得跟进」时使用。
  通过用户授权的 SellerSprite 会话，按品牌/卖家反查店铺全量产品，做 Parent/Child 去重、销量结构、爆款依赖度、新品成功率、生命周期、价格带与品类集中度分析，
  计算 Store Research Score 并输出 A/B 级店铺 TOP10 产品。输出 HTML、Excel、Markdown。禁止虚构数据。
agent_created: true
---

# Amazon Japan Apparel Store Research Skill（V1.2 可执行版）

> **分工**：本 Skill 是 **店铺中心（store-centric）**——先找健康店铺，再找店里值得研究的产品。
> 若要 **市场/细分品类中心（niche-centric）**——直接找稳定需求、低竞争的细分市场和 ASIN，用 `sss-amazon-jp-apparel-research`。
> 两者数据可互通：本 Skill 输出的候选产品可直接喂给 niche 版做竞争/季节性/利润深挖。

## 0. 与 V1.2 规格文档的差异（已修正的三处硬伤）

| # | 规格文档 V1.2 | 本 Skill 修正 |
|---|---|---|
| 1 | 假设能拿到店铺 ID 与全量商品，未验证 | **Step 0 硬门槛**：先跑采集能力探测，确认字段与条数；拿不到就降级为「品牌集合」分析并明确标注，不假装是店铺 |
| 2 | 第19节稳定指数(100) 与第20节 Store Score(100) 并存、权重冲突 | **只保留一个 Store Research Score**；稳定指数不存在，其维度并入 Store Score |
| 3 | 用估算销量算占比并卡死百分比阈值 | 所有店铺级占比标注 `Estimated`；阈值放宽一档；不对 1% 差异下结论；报告统一声明估算口径 |

另补两条规格文档缺失的规则：
- **缺数据处理**：店铺评分缺维度标 `N/A`，不得按 0 或满分补；输出 `score_coverage`（有效权重和），<70 不得评 A/B。
- **样本边界**：「Top20 店铺」是**当前样本内** Top20，不是日本站全市场排名。

---

## 1. 核心判断

> **不是谁卖得最多，而是谁的店铺经营结构最值得研究。**

健康店铺的特征：多个产品持续出单、月销 30–100 的产品多、TOP3 依赖度低、低评论产品也能卖、生命周期长、新品持续成功、品类集中、价格带稳定。

**禁止**：仅通过高销量 ASIN 反推「优秀店铺」。必须先看店铺整体结构，再看单品。

---

## 2. 用户参数（未指定用默认）

```json
{
  "marketplace": "JP",
  "price_min_jpy": 2000,
  "price_max_jpy": 3000,
  "monthly_units_target": 50,
  "monthly_units_soft_min": 30,
  "monthly_units_soft_max": 100,
  "contribution_target_jpy": 500,
  "store_type": "精铺型",
  "categories": ["all_apparel"],
  "exclude": ["强季节", "圣诞", "万圣节", "情人节", "cosplay", "婚礼服", "强品牌", "高商标风险"],
  "top_stores": 20,
  "top_products_per_store": 10,
  "include_seasonality": true,
  "include_profit": true
}
```

说明：
- 月销量是软目标。20–30 件但竞争低、全年稳、评论少、利润好的产品仍可进候选池。
- 预算/交期未提供时：只输出按单款的采购建议，不假设预算。

---

## 3. Step 0 — 采集能力探测（硬门槛，必须先做）

实测确认（2026-11-18，账号 XPYY000005，sellersprite.com 国际站）：

- 站点：`sellersprite.cn` 本机 SSL 握手失败，**只用 `.com`**。
- 登录：必须 `agent-browser open <url> --headed` 由用户手动登录一次；`close` 后登录态失效，全程不要关浏览器。
- 账号可能弹「多人/多地登录」安全提醒——**不处理账号安全、不绕过**，只关闭弹窗继续；若要求绑定微信，交由用户处理。
- 反查入口：**查竞品 `/v3/competitor-lookup`**，可填「品牌 / 卖家 / 关键词 / ASIN」四选一，站点切日本站。
  - 无「店铺 ID → 全部商品」入口；`v3/product-store` 是个人产品库，`v2/store-tracking` 是上新监控，都不是全量采集。
  - **本 Skill 的「店铺」= 品牌或卖家的商品集合**，报告中必须写明是 `brand_scope` 还是 `seller_scope`，不得冒充 Amazon 官方 Store。

### 3.1 探测步骤

```bash
agent-browser open "https://www.sellersprite.com/v3/competitor-lookup" --headed
# 用户手动登录后：
agent-browser eval "(()=>{const e=[...document.querySelectorAll('span')].find(x=>x.innerText.trim()==='日本站');e.click()})()"
# 填一个服装品牌（默认 小倉屋 做探测）
agent-browser eval "(()=>{const s=Object.getOwnPropertyDescriptor(HTMLInputElement.prototype,'value').set;const i=[...document.querySelectorAll('input')].find(x=>(x.placeholder||'').includes('品牌 如'));s.call(i,'小倉屋');i.dispatchEvent(new Event('input',{bubbles:true}))})()"
agent-browser click "<立即查询 ref>"
sleep 8
agent-browser eval "document.body.innerText.match(/搜索结果数[:：]\s*(\d+)/)?.[1]"
```

### 3.2 导出（异步队列型）

点「导出」→ 弹窗「1-5分钟，去个人中心-我的导出」→ 等 90 秒左右 → `/v2/export-log` 取直链 → curl 下载。

```bash
# 取直链（无需鉴权）
agent-browser eval "(()=>{const r=[...document.querySelectorAll('tr')].find(x=>x.innerText.includes('Competitor-JP'));return [...r.querySelectorAll('a')].map(a=>a.href)})()"
curl -sL -o store-products.xlsx "<直链>"
```

### 3.3 已实测字段（66 列，直接对应用户规格第6节）

| 规格要求 | 实际列 | 状态 |
|---|---|---|
| Store / Seller | `品牌`、`BuyBox卖家`、`卖家信息`、`卖家首页`、`卖家所属地` | ✅ 比规格更全 |
| ASIN / Parent ASIN | `ASIN`、`父ASIN`、`SKU` | ✅ |
| Brand / Title / Category | `品牌`、`商品标题`、`类目路径`、`大类目`、`小类目` | ✅ |
| Price | `价格(円)`、`prime价格(円)`、`Coupon` | ✅ |
| Monthly Sales / Revenue | `月销量`、`月销售额(円)`、`子体销量`、`子体销售额(円)`、`销量环比增长率`、`销量同比增长率` | ✅ |
| Reviews / Rating | `评分数`、`月新增评分数`、`评分`、`留评率` | ✅ |
| BSR | `大类BSR`、`小类BSR`、`大类BSR增长率` | ✅ |
| Launch Date | `上架时间`、`上架天数` | ✅ |
| Variations | `变体数` | ✅ |
| 利润辅助 | `FBA(円)`、`毛利率`、`买家运费(円)`、`配送方式` | ⚠️ 毛利率为卖家精灵估算，非自算 |
| **逐月历史销量** | **无** | ❌ 需单独导出，见 §5.4 |

### 3.4 门槛判定

- **实测教训**：「小倉屋」返回 43 条商品，Parent 去重后只剩 **2 个家族**（服装颜色/尺码变体极多）。因此条数不等于样本量，一切结构分析基于**父体家族数**。
- 父体家族数 ≥ 10 且字段齐全 → 走完整流程。
- 父体家族数 < 10 → 标 `Small Sample`，不评 A/B，只做观察（小倉屋实测即属此类：TOP3 占 100%，典型爆款依赖）。
- 若账号无导出权限 → 停止，告知用户需要什么数据文件；**不得用模型知识补数据**。

---

## 4. 发现候选店铺（品牌/卖家集合）

1. **关键词反查**：用日文服装词（メンズ シャツ / レディース トップス / パーカー / ルームウェア / インナー 等）跑关键词挖掘，取高销量 ASIN 的品牌。
2. **ASIN 反查**：查竞品拿 `品牌` / `BuyBox卖家` / `卖家首页`。
3. **横向扩展**：对表现好的品牌/卖家，用「品牌」或「卖家」字段反查全部商品 → 这就是该「店铺」的全量集合。
   - 「品牌」反查已实测可用（43 条 / 66 列）。
   - 「卖家」反查输入框存在但**未实测返回完整性**；首次使用时先验证条数，不足则改用品牌口径或 Amazon 店铺页人工确认。
4. 每个候选集合落到 `input/raw/store-products/<scope-key>.xlsx`，**原始文件不得覆盖**。

采集日志写 `input/raw/collection-log.csv`：date / source / module / query / store / file / status / notes。

---

## 5. 数据结构化与清洗

### 5.1 Parent / Child 去重（强制）

- 市场规模、销量结构、店铺评分 → 按 `父ASIN` 聚合。
- 竞争分析 → 可保留 Child ASIN。
- **`父ASIN` 相同的一组，禁止把子体销量相加后再与父体销量相加。**
- 实测发现：导出中既有「父ASIN = 自身」的父行，也有「父ASIN = 另一 ASIN」的子行；聚合前先按 `父ASIN` 分组，取父行 `月销量` 为该家族销量。无父行时取子体销量最大值并标 `estimated_family_units`。

### 5.2 有效销售

`Monthly Sales > 0` 为有效 ASIN。
`Sales Coverage = 有销量有效Parent数 ÷ 全部有效Parent数`

### 5.3 销量分层

| 月销量 | 分层 |
|---:|---|
| 0 | 无销量 |
| 1–9 | 极低 |
| 10–29 | 低 |
| 30–49 | 稳定小销量 |
| 50–100 | 核心精铺区间 |
| 101–199 | 中高 |
| 200+ | 高 |

### 5.4 季节性数据（独立获取，非 30 天快照）

30 天快照**不能**证明全年稳定。需要时用 `sellersprite-history-export` 的「竞品逐月销量」流程：查竞品 → 行内历史趋势 → 「销量趋势」tab → 导出产品。
- 覆盖 1–27 个月不等，很多 ASIN 只有 2–3 个月 → <8 个月标 `Insufficient`，不参与季节稳定性评分。
- 缺失月份绝不填 0。
- 逐月数据拿不到时，Store Score 的「季节稳定性」维度标 N/A，权重不计入分母。

---

## 6. 店铺结构分析（每个候选集合必做）

1. 总 ASIN / 有效 Parent / Child 数
2. 有效销售数、覆盖率（≥70% 优秀 / 50–69% 良好 / 30–49% 一般 / <30% 较差）
3. 销量分层计数：30–100 数量与占比、50–100 数量、100+ 数量、200+ 数量
4. 爆款依赖度（**全部标注 Estimated**）：
   - `TOP3 Share = TOP3 家族销量合计 ÷ 集合全部家族销量`
   - `TOP10 Share` 同理
   - 判定放宽一档：≤35% 优秀 / 35–55% 一般 / >55% 高风险 / >75% 极高风险
5. 价格结构：最低/最高/平均/中位/主要区间占比；目标带 ¥2,000–3,000 的产品占比
6. 品类集中度：小类目占比，判断是否垂直（前 3 小类目合计 ≥60% 为聚焦）
7. 生命周期：上架天数均值/中位、>12 月占比、>24 月占比；老品（>12月）仍有 ≥30 销量的数量
8. 新品成功率：近 12 月上架 ASIN 中，当前销量 >0 / ≥30 / ≥50 / ≥100 的数量。
   **注意**：这是「当前快照口径」，不是严格的历史新品成功率，报告必须写明。有逐月数据时才算严格口径。
9. 低评论高销量：同时满足 `评分数<100 且月销30–100` 或 `评分数<50 且月销≥30` 的 ASIN 数量
10. 卖家类型：用 `卖家所属地` + `卖家信息` + `卖家首页` 综合判断 China / Japan / Other / Unknown。**不得仅凭品牌名判断**；拿不到就 Unknown。

---

## 7. Store Research Score（唯一 100 分制）

| 指标 | 权重 | 计分 |
|---|---:|---|
| 有效销售 ASIN 覆盖率 | 15 | ≥70% 满分，线性递减 |
| 30–100 销量 ASIN 占比 | 20 | 占比越高越高 |
| 销量稳定性 | 15 | 逐月 CV：≤20% 满分，20–30% 中，>30% 低；无逐月数据 → N/A |
| 爆款依赖度 | 10 | TOP3 ≤35% 满分，>75% 0 分 |
| 产品生命周期 | 10 | >12 月产品占比与老品持续销量 |
| 品类聚焦度 | 10 | 前 3 小类目占比 |
| 低评论销售能力 | 10 | 低评论高销量 ASIN 数量 |
| 价格带一致性 | 5 | 目标带 ¥2,000–3,000 占比 |
| 新品成功率 | 5 | ≥30 单新品占比；无上架时间 → N/A |

### 缺数据规则（强制）

- 缺失维度 = `N/A`，**不得按 0 或满分补**。
- `score_coverage = 有效维度权重之和`。
- 最终分 = `Σ(各维度得分 × 权重) / score_coverage × 100`，并标注 `normalized_from_X_points`。
- `score_coverage < 70` → 不得评 A/B，只能 C（观察）。

### 等级

- A：80–100 强烈建议研究
- B：70–79 值得研究
- C：60–69 持续观察
- D：<60 不作为主要研究对象

### 店铺研究结论模板

每个重点店铺输出 3–8 条数据依据，例如：
```
该店铺值得研究（Estimated，基于 30 天快照）：
1. 43 个有效父 ASIN 中 31 个有销量（覆盖率 72%）。
2. 15 个产品月销 30–100。
3. TOP3 占集合总销量 28%，非爆款依赖。
4. 近 12 月上架 12 款，其中 6 款当前月销 ≥30（快照口径）。
5. 7 个低评论产品仍稳定出单。
6. 核心价格带集中 ¥2,000–¥3,000。
7. 前 3 小类目占 74%，品类聚焦。
```
数据不足的条删掉，不补写。

---

## 8. 产品深挖（仅 A/B 级店铺）

对每个 A/B 集合输出 **TOP10 值得研究产品**：

| 字段 | 来源 |
|---|---|
| ASIN / 标题 / 图片 / 售价 | 导出列 |
| 月销量 / 销售额 | 导出列（父体口径） |
| 评分数 / 评分 / 留评率 / BSR | 导出列 |
| 上架时间 / 上架天数 | 导出列 |
| 变体数 / 尺码复杂度 | `变体数` + SKU 列 |
| 季节性 | §5.4，无数据标 N/A |
| 利润 | §9 |
| 竞争 / 机会 | 集合内份额 + 低评论可进入性 |
| 产品评分 | §8.1 |

链接只能用真实 ASIN：`https://www.amazon.co.jp/dp/{ASIN}`

### 8.1 产品评分（100 分）

| 维度 | 权重 |
|---|---:|
| 全年需求稳定性 | 25 |
| 市场竞争难度 | 25 |
| 利润空间 | 20 |
| 产品差异化空间 | 20 |
| 库存及退货风险 | 10 |

缺维度标 `N/A`，不得因缺数据得高分。等级 A 80–100 / B 70–79 / C 60–69 / D <60。

优先产品：① 月销 30–100；② 20–30 但低评论低竞争全年稳利润好；③ 100+ 但需判断是否爆款依赖。

### 8.2 主动识别不建议研究产品

品牌壁垒高、评论壁垒高、商标/外观专利风险、强季节、强趋势、爆款依赖、极端低价、高退货、尺码复杂、生命周期短。逐条给理由。

---

## 9. 利润模型

`净贡献 = 售价 − 采购 − 国内运输 − 国际头程 − Amazon佣金 − FBA配送 − 仓储 − 广告 − 退货预留 − 其他变动成本`

- 任何关键成本缺失：标 `Profit Data Incomplete`，不得声称确定净利润。
- 卖家精灵 `毛利率` 是估算，只作参考，不写进净贡献。
- 广告无实际数据时可给「假设场景」，必须标 `assumption / not_actual`。

---

## 10. 精铺组合与首批采购

- 组合表：产品 / 月销量 / 贡献；算总销量、总销售额、总贡献、库存资金、风险集中度。
- 首批采购量 = `预测日销量 × 供应链周期 + 安全库存`，不机械设 100 件。
- 新品默认覆盖 30–60 天需求；高风险偏 30 天。未提供预算/交期时只给模型，不假设。

---

## 11. 输出文件

```
output/
  index.html            # 主展示页
  apparel-research.xlsx # 12 个工作表
  research-report.md
input/raw/
  collection-log.csv
  store-products/<scope-key>.xlsx   # 原始导出，不覆盖
```

### HTML 必须包含

- Dashboard：店铺数 / A级 / B级 / 推荐产品 / 数据更新时间
- Store Ranking：样本内 Top20，Store Score、产品数、有销量比例、30–100 数量、TOP3 集中度、新品成功率
- Store Detail（点击店铺）：全部产品、销量/价格/评论分布、生命周期、新品表现、稳定产品、值得研究产品
- Product Detail：ASIN、标题、图片、售价、月销、评论、Rating、BSR、上架时间、季节性、利润、竞争、评分
- 筛选：店铺/ASIN 搜索、店铺评分、产品评分、月销量、价格、评论、季节性、卖家类型、推荐等级
- 点击 ASIN → `https://www.amazon.co.jp/dp/{ASIN}`（仅真实 ASIN）
- 单文件本地可打开，不依赖后端

### Excel 工作表

Executive Summary / Store Ranking / Store Analysis / Store Products / Candidate Products / ASIN Competitors / Market Share / Seasonality / Profit Model / Keyword Data / Seller Analysis / Data Quality & Sources

### Markdown 报告 18 节

执行摘要、用户参数、数据来源、数据质量、值得研究店铺、店铺销售结构、店铺经营模式、店铺评分、值得研究产品、ASIN竞品、季节性、利润、风险、精铺组合、首批采购、不建议研究店铺、不建议研究产品、数据限制。

---

## 12. 数据真实性（最高优先级）

绝对禁止虚构 ASIN / 店铺 / 销量 / 评论 / 价格 / 历史 / 市场份额 / 利润 / 卖家类型。
无法获取写 `Data unavailable`；缺失写 `N/A` / `Unknown`。

证据等级：
- A：多源交叉验证
- B：单一可靠数据源
- C：估算数据（卖家精灵销量、毛利率均属 C）
- D：无法验证

**市场份额只能是「当前样本份额」**：
`ASIN Share = ASIN月销 ÷ 样本总月销`；禁止写「占日本 Amazon 市场 X%」。

### 输出前真实性检查

```
CHECK-01 所有 ASIN 来自原始数据？
CHECK-02 链接由真实 ASIN 生成？
CHECK-03 无虚构销量/价格/评论？
CHECK-04 缺失数据未被填 0？
CHECK-05 父/子体未重复计算？
CHECK-06 店铺占比标注 Estimated？
CHECK-07 市场份额为样本口径？
CHECK-08 30天快照未被当作全年稳定？
CHECK-09 评分缺维度标 N/A 且 score_coverage 已算？
CHECK-10 卖家身份无过度推断？
CHECK-11 HTML/Excel/MD 核心数字一致？
```

任一关键检查失败 → 不得写「分析完成」，输出「数据质量存在问题」并列失败项。

---

## 13. 执行流程

```
读参数 → 建 collection log → 检查登录会话（用户手动）
→ Step 0 采集能力探测 → 关键词反查 → ASIN 反查品牌/卖家
→ 品牌/卖家反查全量商品 → 导出落盘 → Parent/Child 去重
→ 店铺结构分析 → 计算 Store Research Score（含缺数据处理）
→ 筛 A/B 级 → 逐店 TOP10 产品 → 竞品/季节性/利润/风险
→ 产品评分 → 精铺组合 → 采购建议
→ 生成 HTML / Excel / Markdown → 真实性检查 → 输出
```

禁止反过来仅通过高销量 ASIN 推断优秀店铺。
