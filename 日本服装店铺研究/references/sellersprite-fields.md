# SellerSprite 查竞品导出字段字典（日本站实测）

实测时间：2026-11-18
账号：XPYY000005
站点：sellersprite.com（国际站；`.cn` 本机 SSL 失败）
模块：`/v3/competitor-lookup` → 品牌/卖家/ASIN 反查 → 导出
样本：品牌「小倉屋」→ 44 行（1 表头 + 43 商品），66 列
导出文件：`Competitor-JP-Last-30-days-656844.xlsx`

## 字段分组与用途

| 组 | 列 | 在店铺研究中的用途 | 证据等级 |
|---|---|---|---|
| 标识 | ASIN / SKU / 父ASIN / 商品详情页链接 / 商品主图 | 去重、链接、展示；父ASIN 是家族聚合键 | B（直接导出） |
| 归属 | 品牌 / 品牌链接 / 卖家数 / BuyBox卖家 / BuyBox类型 / 卖家所属地 / 卖家信息 / 卖家首页 | 店铺集合定义、卖家类型识别（China/Japan/Other/Unknown） | B |
| 类目 | 类目路径 / 大类目 / 小类目 / 大类BSR / 大类BSR增长数 / 大类BSR增长率 / 小类BSR | 品类集中度、BSR 展示 | B |
| 销量 | 月销量 / 销量环比增长率 / 销量同比增长率 / 子体销量 / 子体销售额(円) / 月销售额(円) | 销量分层、覆盖率、TOP3/TOP10、组合；**均为估算值** | C |
| 变体 | 变体数 / SKU | Parent/Child 去重、尺码复杂度 | B |
| 价格 | 价格(円) / prime价格(円) / Coupon | 价格带结构 | B |
| 评论 | Q&A / 评分数 / 月新增评分数 / 评分 / 留评率 | 低评论高销量机会、评论壁垒 | B |
| 成本 | FBA(円) / 毛利率 / 买家运费(円) / 配送方式 / 商品重量 / 包装尺寸等 | 利润辅助；**毛利率是卖家精灵估算，不作净贡献** | C |
| 生命周期 | 上架时间 / 上架天数 | 新品识别、年龄结构 | B |
| 标签 | LQS / Best Seller标识 / Amazon's Choice / New Release标识 / A+页面 / 视频介绍 / SP广告 / 品牌故事 / 品牌广告 / 秒杀 / AC关键词 / 标签 | Listing 质量辅助 | B |

## 关键坑

1. **同一家族多行**：父 ASIN 相同的多行来自不同子体。聚合时按父ASIN 分组，取父行的 `月销量`；无父行时取子体最大值并标 `estimated_family_units`。
2. **无逐月历史**：此导出只有最近 30 天快照。季节性/严格新品成功率必须用 `sellersprite-history-export` 单独导出逐月数据。
3. **卖家字段稀疏**：并非每行都有 `卖家信息`/`卖家首页`（样本中部分为空）。空值标 Unknown，不得猜测。
4. **估算口径**：`月销量`、`毛利率`、`月销售额` 属估算，店铺级占比全部标 Estimated。
5. **导出异步**：点导出 → 1–5 分钟 → `/v2/export-log` 取直链 `https://o.sellersprite.com/batch-exports/...xlsx`，直链无鉴权可直接 curl。
6. **无店铺 ID 入口**：`v3/product-store` 是个人产品库，`v2/store-tracking` 是上新监控，都不是店铺全量采集。本 Skill 的「店铺」= 品牌或卖家的商品集合，报告需写明 `brand_scope` / `seller_scope`。

## 复现命令

```bash
agent-browser open "https://www.sellersprite.com/v3/competitor-lookup" --headed
# 用户手动登录；关闭"多人/多地登录"安全提醒弹窗（不绕过账号安全）
agent-browser eval "(()=>{const e=[...document.querySelectorAll('span')].find(x=>x.innerText.trim()==='日本站');e.click()})()"
agent-browser eval "(()=>{const s=Object.getOwnPropertyDescriptor(HTMLInputElement.prototype,'value').set;const i=[...document.querySelectorAll('input')].find(x=>(x.placeholder||'').includes('品牌 如'));s.call(i,'小倉屋');i.dispatchEvent(new Event('input',{bubbles:true}))})()"
agent-browser click "<立即查询 ref>"; sleep 8
agent-browser eval "(()=>{const b=[...document.querySelectorAll('button')].find(x=>x.innerText.trim()==='导出'&&x.offsetParent);b.click()})()"
sleep 90
agent-browser open "https://www.sellersprite.com/v2/export-log"
agent-browser eval "(()=>{const r=[...document.querySelectorAll('tr')].find(x=>x.innerText.includes('Competitor-JP'));return [...r.querySelectorAll('a')].map(a=>a.href)})()"
curl -sL -o store-products.xlsx "<直链>"
```
