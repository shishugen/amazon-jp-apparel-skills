# 日本服装店铺研究（sss-amazon-jp-store-research）

Amazon.co.jp 服装店铺研究（Store-Centric）。先找经营结构健康的店铺集合，再找店内值得研究的产品。

> Skill 显示名为「日本服装店铺研究」；目录名保持英文 `sss-amazon-jp-store-research`，与调用路径/软链接解耦。

## 何时用

- 「找值得研究的服装店铺」「哪些店铺精铺做得好」
- 「分析店铺产品组合、爆款依赖度、新品成功率」
- 想从店铺视角而非单品视角做精铺选品

不想研究店铺、只想找稳定需求低竞争细分市场 → 用 `sss-amazon-jp-apparel-research`。

## 核心文件

- `SKILL.md`：完整流程、评分、输出规范。含相对 V1.2 规格文档的三处修正。
- `references/sellersprite-fields.md`：2026-11-18 实测的 66 列字段字典 + 采集命令 + 坑。
- `scripts/store-analysis.py`：Parent 去重、销量结构、TOP3/TOP10、新品成功率、Store Score（缺维度 N/A + 权重归一化）。已用实测数据跑通。
- `scripts/`、`templates/` 下其余文件软链接自 `sss-amazon-jp-apparel-research`，避免版本漂移。

## 快速验证脚本

```bash
python scripts/store-analysis.py <查竞品导出.xlsx> --scope brand --label 品牌名
```

## 实测结论（2026-11-18）

- `.com` 可用，`.cn` 本机 SSL 失败；必须 `--headed` 用户手动登录。
- 查竞品按品牌反查「小倉屋」→ 43 条商品 / 66 列，含卖家所属地、卖家首页、上架时间、父ASIN、变体数等。
- 43 个子体去重后仅 2 个父体家族 → 服装变体多，聚合必须按父ASIN。
- 导出异步，1–5 分钟后在 `/v2/export-log` 取无鉴权直链 curl 下载。
- 无逐月历史销量；季节性需 `sellersprite-history-export` 单独导出。
