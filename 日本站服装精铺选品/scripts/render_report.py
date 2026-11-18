#!/usr/bin/env python3
"""report_data.json → 单文件离线 HTML 选品报告。无外部 CDN，file:// 可直接打开。"""
from __future__ import annotations

import argparse
import json
from pathlib import Path


HTML = r"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>日本站服装精铺选品报告</title>
<style>
  :root{
    --bg:#f5f7fb; --panel:#ffffff; --ink:#1f2937; --muted:#6b7280;
    --line:#e5e7eb; --blue:#2563eb; --blue-soft:#eff6ff;
    --red:#dc2626; --amber:#d97706; --green:#059669; --violet:#7c3aed;
  }
  *{box-sizing:border-box}
  body{margin:0;background:var(--bg);color:var(--ink);font:14px/1.55 -apple-system,BlinkMacSystemFont,"PingFang SC","Hiragino Sans GB","Microsoft YaHei",sans-serif}
  header{background:linear-gradient(135deg,#1e3a8a,#2563eb);color:#fff;padding:28px 32px 24px}
  header h1{margin:0 0 6px;font-size:26px;letter-spacing:.3px}
  header .sub{opacity:.85;font-size:13px}
  .badges{margin-top:14px;display:flex;gap:8px;flex-wrap:wrap}
  .badge{background:rgba(255,255,255,.16);border:1px solid rgba(255,255,255,.28);border-radius:999px;padding:4px 12px;font-size:12px}
  .wrap{max-width:1360px;margin:0 auto;padding:22px 24px 60px}
  .stats{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:12px;margin-bottom:18px}
  .stat{background:var(--panel);border:1px solid var(--line);border-radius:12px;padding:16px 18px;box-shadow:0 1px 2px rgba(16,24,40,.04)}
  .stat .k{font-size:12px;color:var(--muted)}
  .stat .v{font-size:24px;font-weight:700;margin-top:4px}
  .stat .v small{font-size:12px;font-weight:500;color:var(--muted)}
  .card{background:var(--panel);border:1px solid var(--line);border-radius:12px;box-shadow:0 1px 2px rgba(16,24,40,.04);margin-bottom:18px;overflow:hidden}
  .card-h{padding:14px 18px;border-bottom:1px solid var(--line);display:flex;justify-content:space-between;align-items:center;gap:12px;flex-wrap:wrap}
  .card-h h2{margin:0;font-size:16px}
  .filters{display:flex;gap:10px;flex-wrap:wrap;align-items:center;padding:12px 18px;background:#fafbff;border-bottom:1px solid var(--line)}
  .filters label{font-size:12px;color:var(--muted);margin-right:4px}
  select,input[type=search]{border:1px solid var(--line);border-radius:8px;padding:6px 10px;background:#fff;color:var(--ink);font-size:13px;outline:none}
  select:focus,input:focus{border-color:var(--blue);box-shadow:0 0 0 3px rgba(37,99,235,.1)}
  table{width:100%;border-collapse:collapse;font-size:13px}
  th,td{padding:10px 12px;text-align:left;border-bottom:1px solid var(--line);white-space:nowrap}
  th{background:#f9fafb;color:var(--muted);font-weight:600;font-size:12px;cursor:pointer;user-select:none}
  th:hover{color:var(--blue)}
  tbody tr{cursor:pointer}
  tbody tr:hover{background:var(--blue-soft)}
  tbody tr.sel{background:#dbeafe;box-shadow:inset 3px 0 0 var(--blue)}
  .grade{display:inline-block;padding:2px 8px;border-radius:999px;font-size:11px;font-weight:600}
  .g-a{background:#dcfce7;color:#166534}.g-b{background:#dbeafe;color:#1e40af}.g-c{background:#fef3c7;color:#92400e}.g-x{background:#fee2e2;color:#991b1b}
  .score{font-weight:700;color:var(--blue)}
  .detail{display:grid;grid-template-columns:1.05fr 1fr;gap:18px;padding:18px}
  @media(max-width:900px){.detail{grid-template-columns:1fr}}
  .sec h3{margin:0 0 10px;font-size:14px;color:var(--muted);font-weight:600}
  .bar-row{display:grid;grid-template-columns:110px 1fr 60px;gap:8px;align-items:center;margin-bottom:7px;font-size:12px}
  .bar-bg{background:#eef2f7;border-radius:6px;height:9px;overflow:hidden}
  .bar-fg{height:100%;background:var(--blue);border-radius:6px}
  .asins{display:grid;grid-template-columns:repeat(auto-fill,minmax(250px,1fr));gap:10px}
  .asin{border:1px solid var(--line);border-radius:10px;padding:10px;display:flex;gap:10px;background:#fff}
  .asin img{width:54px;height:54px;object-fit:contain;border-radius:6px;background:#f9fafb;flex:0 0 auto}
  .asin .noimg{width:54px;height:54px;border-radius:6px;background:#eef2f7;display:flex;align-items:center;justify-content:center;color:var(--muted);font-size:11px;flex:0 0 auto}
  .asin .meta{min-width:0;flex:1}
  .asin .code{font-family:ui-monospace,SFMono-Regular,Menlo,monospace;font-weight:700;color:var(--blue);font-size:12px;text-decoration:none}
  .asin .title{font-size:12px;color:var(--muted);overflow:hidden;text-overflow:ellipsis;white-space:nowrap;margin:2px 0}
  .asin .nums{font-size:12px}
  .kv{display:grid;grid-template-columns:repeat(auto-fit,minmax(130px,1fr));gap:8px;margin-bottom:14px}
  .kv div{background:#f9fafb;border-radius:8px;padding:8px 10px}
  .kv .k{font-size:11px;color:var(--muted)}.kv .v{font-weight:600;font-size:14px}
  .note{font-size:12px;color:var(--muted);background:#fffbeb;border:1px solid #fde68a;border-radius:8px;padding:10px 12px;margin:0 18px 18px}
  .empty{padding:40px;text-align:center;color:var(--muted)}
  footer{text-align:center;color:var(--muted);font-size:12px;padding:20px}
  .spark{width:100%;height:120px;background:#f9fafb;border:1px solid var(--line);border-radius:8px}
  .monthly th,.monthly td{padding:4px 8px;font-size:12px;cursor:default}
  .monthly th{position:sticky;top:0;background:#f9fafb;z-index:1}
</style>
</head>
<body>
<header>
  <h1>日本站服装精铺选品报告</h1>
  <div class="sub" id="subtitle"></div>
  <div class="badges" id="badges"></div>
</header>
<div class="wrap">
  <div class="stats" id="stats"></div>

  <div class="card">
    <div class="card-h"><h2>候选产品</h2><span style="font-size:12px;color:var(--muted)">点行查看 ASIN 与趋势 · 点表头排序</span></div>
    <div class="filters">
      <div><label>等级</label><select id="fGrade"><option value="all">全部</option><option value="A">A 优先开发</option><option value="B">B 可测试</option><option value="C">C 观察</option></select></div>
      <div><label>最低分</label><select id="fScore"><option value="0">0</option><option value="60">60</option><option value="70">70</option><option value="80">80</option></select></div>
      <div><label>最大CV</label><select id="fCv"><option value="99">不限</option><option value="0.2">20%</option><option value="0.3">30%</option><option value="0.5">50%</option></select></div>
      <div><label>搜索</label><input id="fQ" type="search" placeholder="市场 / 产品 / ASIN"></div>
      <div><label>排序</label><select id="fSort"><option value="total_score">总分</option><option value="net_profit">单件净利</option><option value="monthly_net_profit">月净利</option><option value="cv">CV 低到高</option><option value="avg_price">均价</option><option value="product_cr10">商品CR10 低到高</option><option value="brand_cr3">品牌CR3 低到高</option><option value="seller_cr3">卖家CR3 低到高</option><option value="reviews">竞品评论少到多</option></select></div>
    </div>
    <div style="overflow:auto;max-height:520px">
      <table id="tbl"><thead><tr>
        <th data-k="rank">#</th><th data-k="name">产品 / 市场</th><th data-k="total_score">总分</th><th data-k="grade">等级</th>
        <th data-k="avg_price">均价</th><th data-k="suggested_price">建议售价</th><th data-k="net_profit">单件净利</th><th data-k="target_monthly_sales">目标月销</th>
        <th data-k="monthly_net_profit">预计月净利</th><th data-k="cv">CV</th><th data-k="product_cr10">商品CR10</th><th data-k="brand_cr3">品牌CR3</th><th data-k="seller_cr3">卖家CR3</th><th data-k="asin_count">ASIN数</th>
      </tr></thead><tbody></tbody></table>
    </div>
  </div>

  <div class="card">
    <div class="card-h"><h2 id="detailTitle">产品详情</h2><span id="detailSub" style="font-size:12px;color:var(--muted)"></span></div>
    <div class="detail" id="detail"></div>
    <div class="note" id="detailGaps"></div>
  </div>

  <div class="card">
    <div class="card-h"><h2>数据边界</h2></div>
    <div class="note" style="margin:14px 18px">
      卖家精灵销量/销售额为估算值；本报告展示的是<strong>样本市场份额</strong>，不等于日本站全市场份额。
      所有 ASIN 与链接来自本地导出文件；缺失字段标注“缺失”，未做补造。商品/品牌/卖家集中度均为样本口径，卖家集中度仅在导出含「卖家名称」列时有值。FBA 费与佣金率请在采购前按 Amazon 日本站最新标准核对。
    </div>
  </div>
</div>
<footer>本地离线报告 · 数据来自 raw/ 卖家精灵导出 · 生成于 <span id="gen"></span></footer>

<script id="payload" type="application/json">__PAYLOAD__</script>
<script>
const D = JSON.parse(document.getElementById('payload').textContent);
const $ = s => document.querySelector(s);
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const fmt = (v,d=1) => (v===null||v===undefined||v==='缺失') ? '缺失' : (typeof v === 'number' ? v.toLocaleString('zh-CN',{minimumFractionDigits:0,maximumFractionDigits:d}) : esc(v));
const gradeClass = g => g && g.startsWith('A') ? 'g-a' : (g && g.startsWith('B') ? 'g-b' : (g && g.startsWith('C') ? 'g-c' : 'g-x'));

let state = {grade:'all', minScore:0, maxCv:99, q:'', sort:'total_score', dir:-1, sel:0};

function rows(){
  let list = D.candidates.map((c,i)=>({...c, rank:i+1, cv:(c.seasonality||{}).cv,
    avg_price:c.market_share.avg_price, product_cr10:c.market_share.product_cr10,
    brand_cr3:c.market_share.brand_cr3, seller_cr3:c.market_share.seller_cr3,
    asin_count:c.asins.length}));
  const q = state.q.trim().toLowerCase();
  list = list.filter(c => {
    if(state.grade!=='all' && !(c.grade||'').startsWith(state.grade)) return false;
    if(c.total_score < state.minScore) return false;
    if(state.maxCv<99 && !(c.cv!==null && c.cv!==undefined && c.cv<=state.maxCv)) return false;
    if(q && !((c.name+' '+c.market+' '+c.asins.map(a=>a.asin).join(' ')).toLowerCase().includes(q))) return false;
    return true;
  });
  const k = state.sort;
  list.sort((a,b)=>{
    let av=a[k], bv=b[k];
    if(av===null||av===undefined) av=-1e9;
    if(bv===null||bv===undefined) bv=-1e9;
    if(typeof av==='string') return av.localeCompare(bv)*state.dir;
    return (av-bv)*state.dir;
  });
  return list;
}

function renderTable(){
  const list = rows();
  const tb = $('#tbl tbody');
  if(!list.length){ tb.innerHTML='<tr><td colspan="14" class="empty">没有符合条件的产品。放宽筛选试试。</td></tr>'; }
  else tb.innerHTML = list.map(c=>`
    <tr data-i="${c.rank-1}" class="${c.rank-1===state.sel?'sel':''}">
      <td>${c.rank}</td><td><strong>${esc(c.name)}</strong><br><span style="color:var(--muted);font-size:11px">${esc(c.market)}</span></td>
      <td class="score">${fmt(c.total_score)}</td><td><span class="grade ${gradeClass(c.grade)}">${esc(c.grade)}</span></td>
      <td>¥${fmt(c.avg_price)}</td><td>¥${fmt(c.suggested_price)}</td><td>¥${fmt(c.net_profit)}</td><td>${fmt(c.target_monthly_sales)}</td>
      <td>¥${fmt(c.monthly_net_profit)}</td><td>${fmt(c.cv,2)}</td><td>${fmt(c.product_cr10,2)}</td><td>${fmt(c.brand_cr3,2)}</td><td>${fmt(c.seller_cr3,2)}</td><td>${c.asin_count}</td>
    </tr>`).join('');
  tb.querySelectorAll('tr[data-i]').forEach(tr=>tr.onclick=()=>{state.sel=+tr.dataset.i;renderTable();renderDetail();});
  const cur = list.find(c=>c.rank-1===state.sel);
  if(cur) renderDetail(cur);
}

function spark(c){
  const months=(c.seasonality||{}).months||[];
  if(months.length<2) return '<div class="empty" style="padding:30px">逐月数据不足，无法画图</div>';
  const vals=months.map(m=>m.value), max=Math.max(...vals), min=Math.min(...vals), span=(max-min)||1;
  const W=520,H=110,pad=10;
  const pts=vals.map((v,i)=>[pad+i*(W-2*pad)/(vals.length-1), H-pad-((v-min)/span)*(H-2*pad)]);
  const line=pts.map(p=>p[0].toFixed(1)+','+p[1].toFixed(1)).join(' ');
  const area=`${pad},${H-pad} `+line+` ${W-pad},${H-pad}`;
  return `<svg class="spark" viewBox="0 0 ${W} ${H}" preserveAspectRatio="none">
    <polygon points="${area}" fill="rgba(37,99,235,.12)"/>
    <polyline points="${line}" fill="none" stroke="#2563eb" stroke-width="2.5" stroke-linejoin="round"/>
    ${pts.map(p=>`<circle cx="${p[0].toFixed(1)}" cy="${p[1].toFixed(1)}" r="3" fill="#2563eb"/>`).join('')}
  </svg><div style="font-size:11px;color:var(--muted);margin-top:4px">${esc(months[0].month)} → ${esc(months[months.length-1].month)} · 来源：${esc((c.seasonality||{}).source||'缺失')}</div>`;
}

function renderDetail(force){
  const c = force || D.candidates[state.sel];
  if(!c){ $('#detail').innerHTML='<div class="empty">暂无候选产品</div>'; return; }
  $('#detailTitle').textContent = c.name + ' · ' + c.total_score + ' 分';
  $('#detailSub').textContent = c.grade + ' · ' + c.market;
  const sea=c.seasonality||{}, ms=c.market_share;
  const bar=(label,v,d=0)=>`<div class="bar-row"><span>${label}</span><div class="bar-bg"><div class="bar-fg" style="width:${Math.min(100,Math.max(0,(v||0)*100))}%"></div></div><span>${fmt(v,d)}</span></div>`;
  $('#detail').innerHTML = `
    <div class="sec">
      <h3>关键指标</h3>
      <div class="kv">
        <div><div class="k">建议售价 / 单件净利</div><div class="v">¥${fmt(c.suggested_price)} / ¥${fmt(c.net_profit)}</div></div>
        <div><div class="k">目标月销 / 预计月净利</div><div class="v">${fmt(c.target_monthly_sales)} 件 / ¥${fmt(c.monthly_net_profit)}</div></div>
        <div><div class="k">季节性判定 / CV</div><div class="v">${esc(c.seasonality_label)} · ${fmt(sea.cv,2)}</div></div>
        <div><div class="k">首批采购</div><div class="v">${fmt(c.first_order_qty)} 件</div></div>
      </div>
      <h3>集中度（样本）</h3>
      <div class="kv">
        <div><div class="k">平均价格 / 中位价格</div><div class="v">¥${fmt(ms.avg_price)} / ¥${fmt(ms.median_price)}</div></div>
        <div><div class="k">商品集中度 CR3 / CR5</div><div class="v">${fmt(ms.product_cr3,2)} / ${fmt(ms.product_cr5,2)}</div></div>
        <div><div class="k">商品集中度 HHI</div><div class="v">${fmt(ms.product_hhi,0)}</div></div>
        <div><div class="k">品牌集中度 CR3 / CR5</div><div class="v">${fmt(ms.brand_cr3,2)} / ${fmt(ms.brand_cr5,2)}</div></div>
        <div><div class="k">品牌集中度 HHI</div><div class="v">${fmt(ms.brand_hhi,0)}</div></div>
        <div><div class="k">卖家集中度 CR3 / CR5</div><div class="v">${fmt(ms.seller_cr3,2)} / ${fmt(ms.seller_cr5,2)}</div></div>
        <div><div class="k">卖家集中度 HHI</div><div class="v">${fmt(ms.seller_hhi,0)}</div></div>
        <div><div class="k">品牌数 / 卖家数</div><div class="v">${fmt(ms.brand_count)} / ${fmt(ms.seller_count)}</div></div>
      </div>
      ${bar('商品 CR10（TOP10销量占比）', ms.product_cr10, 2)}
      ${bar('低评论竞品占比', ms.low_review_ratio, 2)}
      <div style="font-size:12px;color:var(--muted);margin-top:6px">
        竞品 ${fmt(ms.competitor_count)} 个 · 月销30–100件竞品 ${fmt(ms.mid_sales_count)} 个 · 卖家集中度仅基于有卖家名称的样本
      </div>
      <h3 style="margin-top:16px">逐月趋势（每月）</h3>
      ${spark(c)}
      <div style="max-height:180px;overflow:auto;margin-top:8px">
        <table class="monthly"><thead><tr><th>月份</th><th>数值</th></tr></thead><tbody>
        ${(sea.months||[]).map(m=>`<tr><td>${esc(m.month)}</td><td>${fmt(m.value)}</td></tr>`).join('')}
        </tbody></table>
      </div>
    </div>
    <div class="sec">
      <h3>竞品 ASIN（${c.asins.length} 个，点 ASIN 打开 Amazon.co.jp）</h3>
      <div class="asins">${c.asins.slice(0,12).map(a=>`
        <div class="asin">
          ${a.image?`<img src="${esc(a.image)}" alt="" onerror="this.style.display='none'">`:'<div class="noimg">无图</div>'}
          <div class="meta">
            <a class="code" href="${esc(a.url)}" target="_blank" rel="noopener">${esc(a.asin)}</a>
            <div class="title" title="${esc(a.title)}">${esc(a.title)}</div>
            <div class="nums">¥${fmt(a.price)} · 月销 ${fmt(a.monthly_sales)} · 评论 ${fmt(a.reviews)} · 份额 ${fmt(a.share,2)}</div>
            <div style="font-size:11px;color:var(--muted)">${esc(a.brand)} · 来源 ${esc(a.source)}</div>
          </div>
        </div>`).join('')}</div>
      ${c.asins.length>12?`<div style="font-size:12px;color:var(--muted);margin-top:8px">仅显示前 12 个，完整列表见 Excel。</div>`:''}
    </div>`;
  $('#detailGaps').textContent = c.data_gaps && c.data_gaps.length ? '数据缺口：'+c.data_gaps.join('；') : '数据缺口：无';
}

function init(){
  const p=D.params||{}, s=D.summary||{};
  $('#subtitle').textContent = `目标 ¥${fmt(p.price_min)}–¥${fmt(p.price_max)} · 月销 ${fmt(s.target_monthly_sales)} 件 · 单件净利 ≥ ¥${fmt(s.target_net_profit)} · 首批 ${fmt(p.first_order_qty)} 件/款`;
  $('#gen').textContent = D.generated_at||'';
  $('#badges').innerHTML = [
    `市场 ${s.market_count||0}`, `候选 ${s.candidate_count||0}`, `优先开发 ${s.priority_count||0}`,
    `平均评分 ${s.avg_score||0}`, `历史 ${p.history_months||0} 个月`
  ].map(x=>`<span class="badge">${esc(x)}</span>`).join('');
  const stats=[
    ['候选产品', s.candidate_count||0, '个'],['优先开发', s.priority_count||0, '个'],
    ['目标月销', s.target_monthly_sales||0, '件/款'],['目标单件净利', '¥'+fmt(s.target_net_profit), ''],
    ['优先组合月净利', '¥'+fmt(s.estimated_monthly_net), '测算'],['平均评分', s.avg_score||0, '/100']
  ];
  $('#stats').innerHTML=stats.map(x=>`<div class="stat"><div class="k">${esc(x[0])}</div><div class="v">${esc(x[1])} <small>${esc(x[2])}</small></div></div>`).join('');

  $('#fGrade').onchange=e=>{state.grade=e.target.value;state.sel=0;renderTable();};
  $('#fScore').onchange=e=>{state.minScore=+e.target.value;state.sel=0;renderTable();};
  $('#fCv').onchange=e=>{state.maxCv=+e.target.value;state.sel=0;renderTable();};
  $('#fQ').oninput=e=>{state.q=e.target.value;state.sel=0;renderTable();};
  $('#fSort').onchange=e=>{state.sort=e.target.value;state.dir=['cv','reviews','product_cr10','brand_cr3','seller_cr3'].includes(e.target.value)?1:-1;renderTable();};
  document.querySelectorAll('th[data-k]').forEach(th=>th.onclick=()=>{
    const k=th.dataset.k;
    if(k==='rank'){state.sort='total_score';state.dir=-1;}
    else if(['cv','product_cr10','brand_cr3','seller_cr3','reviews'].includes(k)){state.sort=k;state.dir=state.dir===1?-1:1;}
    else {state.sort=k;state.dir=state.sort===k?-state.dir:-1;}
    $('#fSort').value=['total_score','net_profit','monthly_net_profit','cv','avg_price','product_cr10','brand_cr3','seller_cr3','reviews'].includes(state.sort)?state.sort:'total_score';
    renderTable();
  });
  renderTable();
}
init();
</script>
</body>
</html>
"""


def main():
    ap = argparse.ArgumentParser(description="生成离线交互 HTML 报告")
    ap.add_argument("--data", type=Path, required=True)
    ap.add_argument("--out", type=Path, required=True)
    args = ap.parse_args()

    data = json.loads(args.data.read_text(encoding="utf-8"))
    payload = json.dumps(data, ensure_ascii=False, separators=(",", ":")).replace("<", "\\u003c")
    html = HTML.replace("__PAYLOAD__", payload)

    out = args.out
    if out.is_dir() or out.suffix.lower() != ".html":
        out = out / "index.html"
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(html, encoding="utf-8")
    print(f"HTML 已生成：{out}")


if __name__ == "__main__":
    main()
