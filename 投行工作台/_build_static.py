"""从 workbench.db 生成【脱敏静态演示版】单文件 HTML。
- 素材库文件名按 id 精确脱敏 + 正则兜底，剔除无关/高危条目
- 台账数据本身已是"某XX"脱敏示例，直接采用
- 初稿模块只放章节大纲，不引用具体金额
"""
import sqlite3, os, re, json, html, datetime

DB = 'C:/Users/公司银行部（超本3）/WorkBuddy/2026-09-09-19-51-07/投行工作台Flask版/workbench.db'
OUT_DIR = 'C:/Users/公司银行部（超本3）/WorkBuddy/2026-09-09-19-51-07/投行工作台Flask版/dist'
os.makedirs(OUT_DIR, exist_ok=True)

# ---------- 素材库脱敏映射 ----------
# 明确剔除：与业务无关 / 工程产物 / 高危
DROP_IDS = {786, 270, 271, 272}   # 广财大课程须知、app.py.txt、requirements.txt、通报格式.jpg

RENAME = {
    294: '投行部五年规划材料.docx',
    293: '投行部年度工作总结及计划.docx',
    287: '投行部年度述职报告.docx',
    292: '投行部半年度工作计划.docx',
    291: '投行部年度工作总结及计划（历史）.docx',
    290: '投行部半年度工作总结.docx',
    286: '投行业务经营思路.pptx',
    289: '投行部年度工作总结及计划（历史）.docx',
    288: '投行业务发展思路汇报材料.docx',
    285: '投行业务发展思路汇报.pptx',
    785: '投放客户存款推动表.xlsx',
    781: '产品情况汇总表.xlsx',
    782: '投行业务台账.xlsx',
    783: '投放客户存款推动表.xlsx',
    784: '结算渠道建设方案（提纲）.doc',
    279: '投行汇报材料.docx',
    280: '营收及非息汇报材料.docx',
    274: '半年度工作会发言材料.pdf',
    273: '季度非息分析及举措.docx',
    284: '跨境投放与外币存款专题汇报.pptx',
    278: '投行部半年度工作总结及计划.docx',
    276: '月度经营分析材料.docx',
    275: '月度经营分析材料.docx',
    282: '板块会汇报材料.docx',
    277: '板块会汇报材料.docx',
    283: '投行部工作汇报材料.pptx',
    281: '投行部战略研讨材料.docx',
    768: '上级沟通事项清单（投行反馈）.xlsx',
    767: '季度经营分析会发言稿.docx',
    688: '工作事项跟踪表.xlsx',
    421: '月度经营分析会材料.pdf',
    264: '投行汇报材料.docx',
    265: '板块专项汇报材料.pptx',
    266: '月度经营分析材料.pptx',
    268: '中长期贷款及中收举措讲稿.docx',
    269: '月度经营分析材料（投行）.pptx',
    262: '交流汇报材料.pptx',
    263: '信息科技工作汇报.pptx',
    261: '投贷联动营收分析.docx',
    258: '经营机构管理实施方案.pdf',
    259: '科技金融工作进展汇报.pdf',
    260: '科技金融年度信贷投放方案.pdf',
    267: '投行部工作汇报材料.pptx',
}

# 正则兜底：万一有未覆盖的新文件，扫掉机构名/日期/版本号/敏感词
FALLBACK = [
    (r'广州分行|深圳分行|总行|分行|支行', ''),
    (r'行领导', ''),
    (r'考核', '评估'),
    (r'专项行动', '工作方案'),
    (r'新台账', '台账'),
    (r'\d{6,8}[-_]?', ''),
    (r'^\d{3,4}[-_]?', ''),
    (r'\((\d+)\)|（(\d+)）', ''),
    (r'[Vv]\d+', ''),
    (r'\d+\.\d+(?=\.)', ''),
    (r'\(\d{4}\)|（\d{4}）', ''),
]

def sanitize(mid, name):
    if mid in RENAME:
        return RENAME[mid]
    s = name
    for pat, rep in FALLBACK:
        s = re.sub(pat, rep, s)
    s = re.sub(r'\s{2,}', ' ', s).strip(' -_（）()')
    return s or '业务材料'

con = sqlite3.connect(DB); cur = con.cursor()

# 素材库
mats = []
seen = set()
for mid, name, cat in cur.execute('SELECT id,name,cat FROM materials ORDER BY cat, mtime DESC'):
    if mid in DROP_IDS:
        continue
    n = sanitize(mid, name)
    key = (cat, n)
    if key in seen:
        continue
    seen.add(key)
    mats.append({'cat': cat, 'name': n})

# 台账
def rows(t, cols):
    cur2 = con.cursor()
    return [dict(zip(cols, r)) for r in cur2.execute(f'SELECT {",".join(cols)} FROM {t}')]

deals = rows('deals', ['name','client','sector','stage','amount','owner','teammate','status','note'])
exps  = rows('expansion_clients', ['name','sector','region','tag','source','status','owner','note'])
govs  = rows('gov_projects', ['name','region','sector','invest','demand','stage','source','owner'])
con.close()

def money(v):
    try:
        return f'{float(v):,.0f}'
    except Exception:
        return '-'

def esc(s):
    return html.escape(str(s or ''))

# ---------- 渲染 ----------
deal_total = sum(float(d['amount'] or 0) for d in deals)
gov_demand = sum(float(g['demand'] or 0) for g in govs)

def table(headers, body_rows):
    th = ''.join(f'<th>{esc(h)}</th>' for h in headers)
    tb = ''
    for r in body_rows:
        tb += '<tr>' + ''.join(f'<td>{c}</td>' for c in r) + '</tr>'
    return f'<table><tr>{th}</tr>{tb}</table>'

kpis = [
    ('在库商机', f'{len(deals)} <small>个 / {money(deal_total)} 万元</small>'),
    ('科技金融拓户', f'{len(exps)} <small>家 / 已接触 {sum(1 for e in exps if e["status"]=="已接触")} 家</small>'),
    ('综合融资储备', f'{len(govs)} <small>个 / 需求 {money(gov_demand)} 万元</small>'),
    ('素材库（脱敏）', f'{len(mats)} <small>份</small>'),
]
kpi_html = ''.join(
    f'<div class="kpi"><div class="t">{esc(t)}</div><div class="v">{v}</div></div>'
    for t, v in kpis)

deal_tbl = table(
    ['商机 / 客户', '行业', '阶段', '金额(万元)', '负责人', '状态'],
    [[f'<b>{esc(d["name"])}</b><div class="mini">{esc(d["client"])}</div>',
      esc(d['sector']), f'<span class="pill p1">{esc(d["stage"])}</span>',
      money(d['amount']), esc(d['owner']),
      f'<span class="pill p2">{esc(d["status"])}</span>'] for d in deals])

exp_tbl = table(
    ['企业名称', '行业', '区域', '标签', '来源', '负责人', '状态'],
    [[esc(e['name']), esc(e['sector']), esc(e['region']),
      f'<span class="pill p5">{esc(e["tag"])}</span>', esc(e['source']),
      esc(e['owner']), f'<span class="pill p2">{esc(e["status"])}</span>'] for e in exps])

gov_tbl = table(
    ['项目名称', '区域', '行业', '总投资(万元)', '融资需求(万元)', '阶段', '来源', '负责人'],
    [[esc(g['name']), esc(g['region']), esc(g['sector']),
      money(g['invest']), money(g['demand']),
      f'<span class="pill p3">{esc(g["stage"])}</span>', esc(g['source']), esc(g['owner'])] for g in govs])

# 素材库按分类
mat_html = ''
for cat in ['投行条线汇报', '综合汇报与上会', '年度总结及计划']:
    items = [m for m in mats if m['cat'] == cat]
    if not items:
        continue
    lis = ''.join(f'<li>{esc(m["name"])}</li>' for m in items)
    mat_html += f'<div class="card"><h2>{esc(cat)} <span class="mini">（{len(items)} 份）</span></h2><ul class="mlist">{lis}</ul></div>'

# 初稿大纲（脱敏，不含金额）
outline = [
    '一、效益情况分析',
    '二、目标及举措',
    '　　聚焦重点客户，做大中长期资产底盘，夯实项目储备',
    '二、境内外协同突围，全力拓展跨境资产增量蓝海',
    '三、精耕可转债结构化业务，加快收益回补节奏',
    '四、把握资本市场双重机遇，以投资突破撬动中收增长',
    '　（一）充分发挥AIC股性资金优势，寻求投资商机',
    '　　　1、稳定现金流　2、上市公司　3、股权投资',
    '　（二）把握港股资本市场历史机遇',
]
outline_html = '<pre class="draft">' + '\n'.join(esc(o) for o in outline) + '</pre>'

TABS = [
    ('home', '📊 首页看板', f'<div class="kpis">{kpi_html}</div>'
        '<div class="card"><h2>工作台说明</h2>'
        '<p style="line-height:1.9">本页面为投行部数智化工作台的<b>静态演示版</b>，用于向同事快速展示能力框架。</p>'
        '<ul class="mlist" style="margin-top:10px">'
        '<li>商机管理：商机台账、阶段漏斗、负责人与共享成员</li>'
        '<li>科技金融：S70 客群拓户名单、接触状态跟踪</li>'
        '<li>综合融资：发改委/推介清单项目储备与融资需求</li>'
        '<li>跨表匹配：名单与台账的 VLOOKUP 式模糊匹配</li>'
        '<li>素材库：汇报材料归档与按日扫描更新</li>'
        '<li>初稿生成：参照原稿框架生成汇报初稿</li>'
        '</ul></div>'),
    ('deals', '💡 商机管理', f'<div class="card"><h2>商机台账</h2>{deal_tbl}</div>'),
    ('tech', '🔬 科技金融', f'<div class="card"><h2>S70 拓户名单</h2>{exp_tbl}</div>'),
    ('gov', '🏛 综合融资', f'<div class="card"><h2>项目储备</h2>{gov_tbl}</div>'),
    ('matcher', '🔗 跨表匹配',
        '<div class="card"><h2>跨表匹配（能力说明）</h2>'
        '<p style="line-height:1.9">完整版支持上传两份名单（如科技局名单 vs 行内存量客户），'
        '自动做<strong>全角括号归一化</strong>与<strong>模糊匹配</strong>（相似度阈值 0.89），'
        '输出三类结果：已开户 / 疑似命中待确认 / 未命中可新拓。</p>'
        '<div class="note">静态演示版不含服务端计算，此页仅展示能力说明。</div></div>'),
    ('library', '📚 素材库',
        '<div class="note">以下为<strong>脱敏后</strong>的材料清单：已移除机构名称、'
        '日期批次号、版本号及内部敏感字样，且不提供任何文件下载或路径信息。</div>' + mat_html),
    ('draft', '📝 初稿生成',
        '<div class="card"><h2>汇报材料框架（脱敏示例）</h2>'
        '<p class="mini" style="margin-bottom:10px">完整版会读取原稿全文，保留结构并按段落标黄提示待更新数据，'
        '可导出 Word。此处仅展示框架。</p>' + outline_html + '</div>'),
]

tab_btns = ''.join(
    f'<a class="tab{" on" if i==0 else ""}" data-t="{tid}" href="javascript:void(0)">{label}</a>'
    for i, (tid, label, _) in enumerate(TABS))
tab_panes = ''.join(
    f'<div class="pane{" on" if i==0 else ""}" id="p-{tid}">{body}</div>'
    for i, (tid, _, body) in enumerate(TABS))

HTML = f'''<!DOCTYPE html>
<html lang="zh"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>投行部 · 数智化工作台（演示版）</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0;font-family:'Microsoft YaHei',sans-serif}}
body{{background:#f4f6fa;color:#1c2333;font-size:13px}}
.top{{background:#12224a;color:#fff;padding:14px 22px}}
.top h1{{font-size:17px}}.top .sub{{color:#8ea1cc;font-size:11px;margin-top:3px}}
.tabs{{background:#12224a;padding:0 12px;display:flex;overflow-x:auto}}
.tab{{color:#c9d4ee;padding:11px 15px;font-size:13px;white-space:nowrap;text-decoration:none}}
.tab.on{{background:#2456c8;color:#fff;font-weight:600;border-radius:8px 8px 0 0}}
.pane{{display:none;padding:18px 22px}}.pane.on{{display:block}}
.card{{background:#fff;border:1px solid #e5e9f0;border-radius:12px;padding:16px 18px;margin-bottom:14px}}
h2{{font-size:15px;margin-bottom:12px}}
table{{width:100%;border-collapse:collapse;font-size:12.5px}}
th{{background:#f7f9fc;color:#5b6478;text-align:left;padding:9px 10px;border-bottom:1px solid #e5e9f0;white-space:nowrap}}
td{{padding:9px 10px;border-bottom:1px solid #eef1f6;vertical-align:top}}
tr:hover td{{background:#f8faff}}
.kpis{{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:14px}}
.kpi{{background:#fff;border:1px solid #e5e9f0;border-radius:12px;padding:14px 16px}}
.kpi .t{{color:#5b6478;font-size:12px}}.kpi .v{{font-size:22px;font-weight:700;margin-top:6px}}
.kpi .v small{{font-size:12px;font-weight:400;color:#5b6478}}
.pill{{display:inline-block;padding:2px 9px;border-radius:99px;font-size:11px;font-weight:600}}
.p1{{background:#eaf0fd;color:#2456c8}}.p2{{background:#e8f6ee;color:#1a7f4d}}
.p3{{background:#fdf4e3;color:#b97a12}}.p5{{background:#f0ecfb;color:#6d4fc2}}
.mini{{font-size:11.5px;color:#5b6478}}
.note{{background:#fff8e6;border:1px solid #f0d9a8;color:#8a6d1f;border-radius:8px;padding:10px 14px;margin-bottom:12px;font-size:12px;line-height:1.7}}
.mlist{{margin:6px 0 0 18px;line-height:2;font-size:12.5px}}
pre.draft{{white-space:pre-wrap;font-size:13px;line-height:2;background:#fbfcfe;border:1px solid #e5e9f0;border-radius:8px;padding:16px}}
.foot{{padding:14px 22px;color:#8b95ab;font-size:11.5px;text-align:center}}
@media(max-width:900px){{.kpis{{grid-template-columns:repeat(2,1fr)}}}}
</style></head><body>
<div class="top"><h1>投行部 · 数智化工作台</h1>
<div class="sub">演示版 · 数据为示例/脱敏内容 · 编辑与协同请使用内网完整版</div></div>
<div class="tabs">{tab_btns}</div>
{tab_panes}
<div class="foot">静态演示版 · 生成于 {datetime.date.today().isoformat()} · 不含任何真实客户信息与内部文件</div>
<script>
document.querySelectorAll('.tab').forEach(function(a){{
  a.onclick=function(){{
    document.querySelectorAll('.tab').forEach(function(x){{x.classList.remove('on')}});
    document.querySelectorAll('.pane').forEach(function(x){{x.classList.remove('on')}});
    a.classList.add('on');
    document.getElementById('p-'+a.dataset.t).classList.add('on');
  }};
}});
</script></body></html>'''

out = os.path.join(OUT_DIR, 'index.html')
with open(out, 'w', encoding='utf-8') as f:
    f.write(HTML)

print('生成:', out, os.path.getsize(out), 'bytes')
print('素材库: 原始 47 -> 脱敏去重后', len(mats), '份（剔除', len(DROP_IDS), '条无关/高危）')
for cat in ['投行条线汇报', '综合汇报与上会', '年度总结及计划']:
    print('   ', cat, sum(1 for m in mats if m['cat'] == cat))
print('商机', len(deals), '拓户', len(exps), '融资项目', len(govs))
