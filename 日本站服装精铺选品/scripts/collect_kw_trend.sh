#!/bin/bash
# 采集单个关键词的逐月搜索趋势到 raw/keyword_trend_<slug>_monthly.xlsx
# 用法: ./collect_kw_trend.sh "<日语关键词>" <slug>
set -u
export PATH="/Users/shishugen/.workbuddy/binaries/node/versions/22.22.2/bin:$PATH"
export HTTP_PROXY=http://127.0.0.1:7897 HTTPS_PROXY=http://127.0.0.1:7897
export http_proxy=$HTTP_PROXY https_proxy=$HTTPS_PROXY
KW="$1"; SLUG="$2"
ROOT="/Users/shishugen/.workbuddy/skills/ssg-amazon-jp-apparel-niche-research"
TMP="/tmp/kwt-${SLUG}"
mkdir -p "$TMP"

# 打开关键词挖掘页并等待
agent-browser open "https://www.sellersprite.com/v3/keyword-miner" >/dev/null 2>&1
sleep 5

# 填词（真实键盘）
agent-browser click "input[placeholder*='输入关键词']" >/dev/null 2>&1
agent-browser keyboard type "$KW" >/dev/null 2>&1
sleep 1
agent-browser click "button:has-text('立即查询')" >/dev/null 2>&1 || agent-browser eval "(()=>{const b=[...document.querySelectorAll('button')].find(e=>e.innerText.includes('立即查询'));if(b){b.click();return 1}return 0})()" >/dev/null 2>&1
sleep 7

# 点第一行趋势图标
agent-browser eval "(()=>{const el=[...document.querySelectorAll('.icon-google-trends-icon')].find(e=>e.offsetParent);if(!el)return 0;el.click();return 1})()" >/dev/null 2>&1
sleep 4
# 切搜索趋势 tab（默认可能是 Google Trends）
agent-browser eval "(()=>{const t=[...document.querySelectorAll('[role=tab]')].find(e=>e.offsetParent&&e.innerText.trim()==='搜索趋势');if(t)t.click();return 1})()" >/dev/null 2>&1
sleep 3
# 按月
agent-browser eval "(()=>{const b=[...document.querySelectorAll('span,button')].find(e=>e.offsetParent&&e.children.length===0&&e.innerText.trim()==='按月');if(b){b.click();return 1}return 0})()" >/dev/null 2>&1
sleep 2
# 找导出 ref（抽屉里 链接"导出"）
REF=$(agent-browser snapshot -i 2>&1 | grep 'link.*导出' | grep -oE 'e[0-9]+' | head -1)
if [ -z "$REF" ]; then
  echo "FAIL $SLUG: 无导出ref"; exit 1
fi
agent-browser download "@$REF" "$TMP/raw.xlsx" >/dev/null 2>&1
if [ ! -s "$TMP/raw.xlsx" ]; then echo "FAIL $SLUG: 下载失败"; exit 1; fi
echo "OK $SLUG: ref=$REF"
