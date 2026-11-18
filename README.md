# Amazon JP Apparel Skills

亚马逊日本站服装选品与店铺研究的 WorkBuddy Skill 集合（私有仓库）。

## 内容

| 文件夹 | SKILL.md 中文名 | 用途 |
|---|---|---|
| `日本站服装精铺选品/` | 日本站服装精铺选品 | 从卖家精灵导出关键词/竞品/逐月趋势，做市场占比、目标季节窗口（默认春夏秋3–11月）季节性、利润与 100 分评分，产出 HTML + Excel + Markdown |
| `日本服装店铺研究/` | 日本服装店铺研究 | 按品牌/卖家反查店铺全量产品，做 Parent/Child 去重、销量结构、爆款依赖度与店铺评分 |

## 目录结构

仅包含 skill 本体（SKILL.md、scripts、templates、references）。
`raw/`（采集原始数据）与 `output/`（运行报告）不入库，仅保留 `.gitkeep` 骨架，运行时自行生成。

## 本地使用

把对应文件夹复制到 `~/.workbuddy/skills/` 下即可（目录名保持英文亦可，SKILL.md 的 `name` 字段为中文显示名）：

```
~/.workbuddy/skills/<任意目录名>/SKILL.md
```

选品 skill 依赖 `openpyxl`（本机用清华镜像安装）：

```bash
/Users/shishugen/.workbuddy/binaries/python/envs/default/bin/pip install -i https://pypi.tuna.tsinghua.edu.cn/simple openpyxl
```

浏览器采集依赖 `agent-browser`，账号需用户手动在有头浏览器登录一次，不保存密码。

## 数据边界

所有销量/销售额均为卖家精灵估算值，报告中的份额是**样本份额**，不是日本站全市场份额。禁止虚构 ASIN 与数据。
