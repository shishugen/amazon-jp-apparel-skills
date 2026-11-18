# 日本站 FBA 利润口径

## 公式

```text
单件净利 = 售价
         − 佣金（售价 × referral_rate）
         − FBA 配送费（fba_fee_jpy）
         − 采购成本（purchase_cost_jpy）
         − 头程物流（inbound_shipping_jpy）
         − 广告预留（售价 × ad_cost_rate）
         − 退货预留（售价 × return_reserve_rate）
         − 其他（other_cost_jpy）

净利率 = 单件净利 / 售价
预计月净利 = 单件净利 × 目标月销
```

全部金额 JPY，费率用小数（`0.10` = 10%）。

## 参数位置

`templates/user-params.json` → `costs`：

| 键 | 默认值 | 说明 |
|---|---:|---|
| `purchase_cost_jpy` | 700 | 1688/工厂采购成本，按实际报价改 |
| `inbound_shipping_jpy` | 180 | 头程分摊到单件 |
| `fba_fee_jpy` | 380 | 按尺寸/重量档位核对 |
| `referral_rate` | 0.10 | 服装类目佣金率按实际类目核对 |
| `ad_cost_rate` | 0.10 | 广告预留 |
| `return_reserve_rate` | 0.05 | 服装退货预留 |
| `other_cost_jpy` | 0 | 包装、贴标、仓储等 |

## 红线

- 这些默认值是**测算起点，不是 Amazon 官方报价**。采购前必须到 Seller Central 的费用页 / FBA Calculator 按实际 ASIN 尺寸重量、类目核对。
- 服装类目佣金率、低价商品规则、2026 年费率调整都可能变。**不要在报告里写死费率来自 Amazon，只写"来自参数模板，已按最新标准核对/待核对"**。
- 日本消费税、进口关税如适用，并入 `purchase_cost_jpy` 或 `other_cost_jpy`，不要漏算。
- 报告中的「预计月净利」未扣固定运营成本（店铺、软件、人工），不要当真实到手利润。

## 目标值

默认单件净利目标 ¥500。商品若月销 30 件、单件净利 ¥600，月净利 ¥18,000——精铺模式下也算可用候选，不能因为不是爆款就淘汰。

评分里的利润分：

| 单件净利 | 得分（20分） |
|---|---:|
| ≥ 目标 ×1.5 | 20 |
| ≥ 目标 | 16 |
| ≥ 目标 ×0.7 | 10 |
| > 0 | 5 |
| ≤ 0 | 0 |
