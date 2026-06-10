"""
广告命名工具生成器（Google Sheets 版）
用法：python generate_tool.py
读取 Google Sheet → 生成 ad-naming-tool.html
维护枚举值：直接编辑 Google Sheet，刷新页面即可生效
"""

import json, os, sys, re
from pathlib import Path

# fix windows console encoding
sys.stdout.reconfigure(encoding='utf-8')

try:
    import pandas as pd
except ImportError:
    print("需要安装依赖：pip install pandas openpyxl")
    exit(1)

# ===== 配置 =====
SHEET_ID = "1zvP0nuD1n6d0G5GD_rTpsUL785k5xoOt3odfQ9yoTMk"
SHEET_NAME = "Sheet1"
SHEET_CSV_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/gviz/tq?tqx=out:csv&sheet={SHEET_NAME}"

EXCEL_PATH = Path(r"D:/serein/工作/媒介建设/ad campaign namebuilder.xlsx")
OUTPUT_HTML = Path(__file__).parent / "ad-naming-tool.html"

FIELD_TYPES = {
    "Market":           "dropdown",
    "Product Type":     "dropdown",
    "Product Code":     "dropdown",
    "Media Funnel":     "tag",
    "Ad Product":       "tag",
    "Bidding Goal":     "tag",
    "URL":              "tag",
    "Custom":           "custom",
}

CUSTOM_HELP = {
    "Search Ad": ["dsa","dsa-brand","dsa-product","dsa-generic","brand","product","generic","competitor"],
    "[PMAX] Pmax Ad": ["brand-feed","brand-asset","nonbrand-feed","nonbrand-asset","brand-all","nonbrand-all"],
}

PAGE_TITLE = "Ad Campaign 命名生成器 · Campaign Name Builder"
PAGE_SUBTITLE = "按照规范快速生成广告系列命名，支持复制 / 导出 · 数据来源：Google Sheets（在线维护）"


def js_key(name):
    return name.lower().replace(" ", "-")

def esc_js(s):
    return s.replace("\\","\\\\").replace("'","\\'").replace('"','\\"').replace("\n","\\n")

def esc_html(s):
    return s.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;").replace('"',"&quot;")


def load_data_from_excel():
    """从本地 Excel 读取数据，作为离线备用/初始嵌入"""
    df = pd.read_excel(EXCEL_PATH, sheet_name=0, header=0)
    data = {}
    for col in df.columns:
        vals = df[col].dropna().astype(str).str.strip().tolist()
        data[str(col)] = vals
    return data


def load_data_from_csv_text(csv_text):
    """解析 CSV 文本为字段字典"""
    import csv, io
    reader = csv.reader(io.StringIO(csv_text))
    headers = next(reader)
    data = {h: [] for h in headers}
    for row in reader:
        for i, h in enumerate(headers):
            if i < len(row) and row[i].strip():
                val = row[i].strip()
                if val not in data[h]:
                    data[h].append(val)
    return data


# ===== HTML TEMPLATES (plain strings, NO f-string nesting) =====

CSS = r"""
:root {
  --primary: #4F46E5;
  --primary-light: #EEF2FF;
  --primary-hover: #4338CA;
  --success: #10B981;
  --danger: #EF4444;
  --danger-light: #FEF2F2;
  --border: #E5E7EB;
  --text: #111827;
  --muted: #6B7280;
  --bg: #F3F4F6;
  --card: #FFFFFF;
  --shadow: 0 1px 3px rgba(0,0,0,0.08);
}
* { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family:-apple-system,BlinkMacSystemFont,'Segoe UI','PingFang SC','Microsoft YaHei',sans-serif;background:var(--bg);color:var(--text);min-height:100vh }

.header { background:var(--primary);padding:18px 28px;display:flex;align-items:center;gap:14px }
.header-icon { font-size:26px;line-height:1 }
.header-title { font-size:18px;font-weight:700;color:#fff }
.header-sub { font-size:12px;color:rgba(255,255,255,0.7);margin-top:2px }

.layout { display:grid;grid-template-columns:1fr 400px;gap:20px;max-width:1200px;margin:24px auto;padding:0 20px 40px;align-items:start }
@media(max-width:960px){ .layout{grid-template-columns:1fr} }

.card { background:var(--card);border:1px solid var(--border);border-radius:12px;padding:22px;box-shadow:var(--shadow) }
.card + .card { margin-top:16px }
.card-header { display:flex;align-items:center;gap:10px;margin-bottom:20px;padding-bottom:14px;border-bottom:1px solid var(--border) }
.card-icon { width:32px;height:32px;border-radius:8px;background:var(--primary-light);display:flex;align-items:center;justify-content:center;font-size:16px;flex-shrink:0 }
.card-title { font-size:15px;font-weight:700 }
.card-desc { font-size:12px;color:var(--muted);margin-top:1px }

.fields-grid { display:grid;grid-template-columns:1fr 1fr;gap:16px }
@media(max-width:600px){ .fields-grid{grid-template-columns:1fr} }
.field-full { grid-column:1/-1 }

.field { display:flex;flex-direction:column;gap:6px }
.field-label { font-size:12px;font-weight:600;color:var(--muted);text-transform:uppercase;letter-spacing:.05em;display:flex;align-items:center;gap:5px }
.field-label .req { color:var(--danger);font-size:14px;line-height:1 }
.field-label .optional { font-size:10px;font-weight:500;text-transform:none;color:var(--muted);background:#F3F4F6;padding:1px 5px;border-radius:4px }

.tag-wrap { display:flex;flex-wrap:wrap;gap:7px;padding:10px;border:1.5px solid var(--border);border-radius:10px;min-height:44px;cursor:default;transition:border-color .15s }
.tag { display:inline-flex;align-items:center;padding:4px 10px;border:1.5px solid var(--border);border-radius:6px;font-size:12px;font-weight:500;cursor:pointer;transition:all .12s;color:var(--muted);background:var(--bg);user-select:none;white-space:nowrap }
.tag:hover { border-color:var(--primary);color:var(--primary);background:var(--primary-light) }
.tag.active { border-color:var(--primary);background:var(--primary);color:#fff;font-weight:600 }

.select-wrap { position:relative }
.select-search { width:100%;padding:9px 36px 9px 12px;border:1.5px solid var(--border);border-radius:10px;font-size:14px;color:var(--text);background:var(--card);transition:border-color .15s,box-shadow .15s;cursor:pointer }
.select-search::placeholder { color:#9CA3AF;font-size:13px }
.select-search:focus { outline:none;border-color:var(--primary);box-shadow:0 0 0 3px rgba(79,70,229,.1) }
.select-search.has-value { color:var(--primary);font-weight:600 }
.select-arrow { position:absolute;right:10px;top:50%;transform:translateY(-50%);pointer-events:none;color:var(--muted);font-size:12px }
.dropdown { position:absolute;top:calc(100%+4px);left:0;right:0;background:var(--card);border:1.5px solid var(--border);border-radius:10px;box-shadow:0 8px 24px rgba(0,0,0,.08);max-height:240px;overflow-y:auto;z-index:100;display:none }
.dropdown.open { display:block }
.dropdown-item { padding:9px 14px;font-size:13px;cursor:pointer;transition:background .1s;color:var(--text) }
.dropdown-item:hover { background:var(--primary-light);color:var(--primary) }
.dropdown-item.selected { background:var(--primary-light);color:var(--primary);font-weight:600 }
.dropdown-item.no-match { color:var(--muted);font-style:italic;pointer-events:none }

.custom-input-wrap { position:relative }
.custom-input { width:100%;padding:9px 12px;border:1.5px solid var(--border);border-radius:10px;font-size:14px;color:var(--text);transition:border-color .15s,box-shadow .15s }
.custom-input:focus { outline:none;border-color:var(--primary);box-shadow:0 0 0 3px rgba(79,70,229,.1) }
.custom-input.error { border-color:var(--danger);background:var(--danger-light) }
.custom-input.ok { border-color:var(--success) }
.custom-error-msg { font-size:11px;color:var(--danger);margin-top:4px;display:none }
.custom-error-msg.show { display:block }
.custom-rules { font-size:11px;color:var(--muted);margin-top:4px;line-height:1.5 }
.custom-rules .rule { display:flex;align-items:center;gap:4px }
.custom-rules .rule.pass { color:var(--success) }
.custom-rules .rule.fail { color:var(--danger) }

.custom-help { margin-top:12px;border:1px solid var(--border);border-radius:10px;overflow:hidden }
.custom-help-header { display:flex;align-items:center;justify-content:space-between;padding:9px 14px;background:#F3F4F6;cursor:pointer;user-select:none;font-size:12px;font-weight:600;color:var(--muted);transition:background .12s }
.custom-help-header:hover { background:#E5E7EB }
.custom-help-toggle { font-size:11px;transition:transform .2s }
.custom-help-toggle.open { transform:rotate(180deg) }
.custom-help-body { display:none;padding:12px 14px 14px }
.custom-help-body.show { display:block }
.help-group { margin-bottom:12px }
.help-group:last-child { margin-bottom:0 }
.help-group-title { font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:var(--primary);margin-bottom:7px }
.help-tags { display:flex;flex-wrap:wrap;gap:6px }
.help-tag { padding:4px 10px;border:1px solid var(--border);border-radius:6px;font-size:11px;font-weight:500;cursor:pointer;transition:all .12s;color:var(--text);background:white }
.help-tag:hover { border-color:var(--primary);color:var(--primary);background:var(--primary-light) }

.result-sticky { position:sticky;top:20px }
.result-dark { background:#18181B;border-radius:12px;padding:20px 20px 16px;margin-bottom:14px }
.result-dark-label { font-size:11px;text-transform:uppercase;letter-spacing:.08em;color:#71717A;margin-bottom:10px;font-weight:600 }
.result-name-display { font-size:15px;font-weight:700;color:#F4F4F5;word-break:break-all;line-height:1.6;min-height:28px }
.seg-part { color:#A78BFA }
.seg-sep { color:#3F3F46 }
.seg-empty { color:#52525B;font-style:italic;font-size:13px;font-weight:400 }
.seg-custom { color:#34D399 }
.result-len { font-size:11px;color:#52525B;margin-top:8px;text-align:right }

.btn-primary { width:100%;padding:11px;border-radius:10px;font-size:14px;font-weight:700;cursor:pointer;border:none;background:var(--primary);color:white;transition:background .15s;display:flex;align-items:center;justify-content:center;gap:8px }
.btn-primary:hover { background:var(--primary-hover) }
.btn-primary.copied { background:var(--success) }
.btn-row { display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-top:10px }
.btn-secondary { padding:9px 10px;border-radius:10px;font-size:13px;font-weight:500;cursor:pointer;border:1.5px solid var(--border);background:var(--card);color:var(--text);transition:all .15s;display:flex;align-items:center;justify-content:center;gap:5px }
.btn-secondary:hover { border-color:var(--primary);color:var(--primary);background:var(--primary-light) }
.btn-download { width:100%;padding:8px;border-radius:10px;font-size:12px;font-weight:500;cursor:pointer;border:1.5px dashed var(--border);background:transparent;color:var(--muted);transition:all .15s;display:flex;align-items:center;justify-content:center;gap:5px;margin-top:8px }
.btn-download:hover { border-color:var(--primary);color:var(--primary) }

.breakdown-list { display:flex;flex-direction:column;gap:7px }
.breakdown-row { display:flex;align-items:center;gap:10px;padding:8px 10px;border-radius:8px;background:var(--bg) }
.breakdown-num { width:20px;height:20px;border-radius:5px;background:var(--primary-light);color:var(--primary);font-size:10px;font-weight:700;display:flex;align-items:center;justify-content:center;flex-shrink:0 }
.breakdown-field { font-size:12px;color:var(--muted);flex:1 }
.breakdown-val { font-size:13px;font-weight:600;color:var(--text) }
.breakdown-val.custom-val { color:var(--success) }
.empty-hint { font-size:13px;color:var(--muted);text-align:center;padding:14px 0 }

.history-scroller { max-height:200px;overflow-y:auto }
.history-row { display:flex;align-items:center;gap:8px;padding:8px 10px;border-radius:8px;cursor:pointer;transition:background .1s }
.history-row:hover { background:var(--primary-light) }
.history-name { font-size:12px;font-weight:500;flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap }
.history-time { font-size:10px;color:var(--muted);flex-shrink:0 }
.history-copy-btn { width:26px;height:26px;border-radius:6px;border:1px solid var(--border);background:var(--card);font-size:12px;cursor:pointer;display:flex;align-items:center;justify-content:center;flex-shrink:0;transition:all .1s }
.history-copy-btn:hover { background:var(--primary-light);border-color:var(--primary) }
::-webkit-scrollbar{width:4px;height:4px} ::-webkit-scrollbar-thumb{background:#D1D5DB;border-radius:2px}

/* Data source status bar */
.datasource-bar { display:flex;align-items:center;gap:8px;padding:8px 14px;border-radius:8px;font-size:11px;margin-bottom:16px }
.datasource-bar.ok { background:#ECFDF5;border:1px solid #A7F3D0;color:#065F46 }
.datasource-bar.error { background:#FEF2F2;border:1px solid:#FECACA;color:#991B1B }
.datasource-bar.loading { background:#FFFBEB;border:1px solid #FDE68A;color:#92400E }
.datasource-bar .ds-dot { width:8px;height:8px;border-radius:50%;flex-shrink:0 }
.datasource-bar.ok .ds-dot { background:#10B981 }
.datasource-bar.error .ds-dot { background:#EF4444 }
.datasource-bar.loading .ds-dot { background:#F59E0B;animation:pulse-dot 1s infinite }
.datasource-bar .ds-msg { flex:1 }
.datasource-bar .ds-refresh { cursor:pointer;font-size:11px;font-weight:600;white-space:nowrap;opacity:.7;transition:opacity .15s }
.datasource-bar .ds-refresh:hover { opacity:1 }
@keyframes pulse-dot { 0%,100%{opacity:1} 50%{opacity:.3} }

/* Loading skeleton */
.skeleton { background:linear-gradient(90deg,#f0f0f0 25%,#e0e0e0 50%,#f0f0f0 75%);background-size:200%100%;animation:skeleton-shimmer 1.5s infinite;border-radius:8px }
@keyframes skeleton-shimmer { 0%{background-position:-200% 0} 100%{background-position:200% 0} }
"""


JS_TEMPLATE = """
// ===== CONFIG =====
const SHEET_URL = '{SHEET_CSV_URL}';
const FALLBACK_DATA = {FALLBACK_JSON};

const FIELD_TYPES = {FIELD_TYPES_JSON};
let FIELDS_ORDER = [];
let FIELD_LABELS = {{}};
let OPTIONS = {{}};
let DATA_READY = false;

const sel = {{}};
let customValid = true;
const copyHistory = [];


// ===== DATA LOADING =====
async function loadData() {{
  showDatasourceStatus('loading', '正在从 Google Sheet 拉取最新数据...');

  try {{
    const resp = await fetch(SHEET_URL);
    if (!resp.ok) throw new Error('HTTP ' + resp.status);
    const csvText = await resp.text();
    const parsed = parseCSV(csvText);

    // Build field structures from parsed data
    FIELDS_ORDER = [];
    FIELD_LABELS = {{}};
    OPTIONS = {{}};

    for (const [colName, vals] of Object.entries(parsed)) {{
      const key = colName.toLowerCase().replace(/ /g, '-');
      FIELDS_ORDER.push(key);
      FIELD_LABELS[key] = colName;
      OPTIONS[key] = vals;
    }}

    renderAllFields();
    DATA_READY = true;
    const totalVals = Object.values(OPTIONS).reduce((a,b)=>a+b.length,0);
    showDatasourceStatus('ok', '数据已同步自 Google Sheet (' + Object.keys(OPTIONS).length + ' 字段, ' + totalVals + ' 枚举值) · ' + new Date().toLocaleTimeString('zh-CN',{{hour:'2-digit',minute:'2-digit'}}));
  }} catch(err) {{
    console.warn('Sheet fetch failed, using fallback:', err);
    loadFallback();
    showDatasourceStatus('error', '无法连接 Google Sheet，已使用内置备用数据（' + new Date().toLocaleString('zh-CN') + '）');
  }}
}}

function parseCSV(csvText) {{
  const lines = csvText.split('\\n').filter(l=>l.trim());
  if (lines.length < 2) throw new Error('Empty CSV');
  const headers = lines[0].split(',').map(h => h.trim().replace(/^"|"$/g, ''));
  const data = {{}};
  headers.forEach(h => data[h] = []);

  for (let i = 1; i < lines.length; i++) {{
    const cells = splitCSVLine(lines[i]);
    for (let j = 0; j < headers.length; j++) {{
      const val = (cells[j] || '').trim();
      if (val && !data[headers[j]].includes(val)) {{
        data[headers[j]].push(val);
      }}
    }}
  }}
  return data;
}}

function splitCSVLine(line) {{
  const result = [];
  let cur = '';
  let inQuotes = false;
  for (let i = 0; i < line.length; i++) {{
    const ch = line[i];
    if (ch === '"') {{ inQuotes = !inQuotes; }}
    else if (ch === ',' && !inQuotes) {{ result.push(cur); cur = ''; }}
    else {{ cur += ch; }}
  }}
  result.push(cur);
  return result;
}}

function loadFallback() {{
  FIELDS_ORDER = [];
  FIELD_LABELS = {{}};
  OPTIONS = {{}};
  for (const [colName, vals] of Object.entries(FALLBACK_DATA)) {{
    const key = colName.toLowerCase().replace(/ /g, '-');
    FIELDS_ORDER.push(key);
    FIELD_LABELS[key] = colName;
    OPTIONS[key] = vals;
  }}
  renderAllFields();
  DATA_READY = true;
}}

function showDatasourceStatus(state, msg) {{
  const el = document.getElementById('datasource-status');
  el.className = 'datasource-bar ' + state;
  el.innerHTML = '<span class="ds-dot"></span><span class="ds-msg">' + msg + '</span>'
    + '<span class="ds-refresh" onclick="loadData()">[REFRESH]</span>';
}}


// ===== RENDER FIELDS DYNAMICALLY =====
function renderAllFields() {{
  const container = document.getElementById('fields-container');
  let html = '';
  let dropdownCount = 0;

  for (const f of FIELDS_ORDER) {{
    const ft = FIELD_TYPES[f] || 'tag';
    const label = FIELD_LABELS[f];
    const opts = OPTIONS[f] || [];

    if (ft === 'dropdown') {{
      dropdownCount++;
      const isFullWidth = (f === 'product-code' || opts.length > 30);
      const phMap = {{'market':'搜索或选择市场...','product-type':'搜索产品类型...','product-code':'搜索产品型号（如 argus 4、rlc-410）...'}};
      const ph = phMap[f] || '搜索' + label + '...';
      html += '<div class="field' + (isFullWidth ? ' field-full' : '') + '">'
        + '<div class="field-label"><span class="req">*</span> ' + escHtml(label) + '</div>'
        + '<div class="select-wrap" id="wrap-' + f + '">'
        + '  <input class="select-search" id="search-' + f + '" placeholder="' + escHtml(ph) + '" autocomplete="off"'
        + '    onfocus="openDropdown(\\'' + f + '\\')" oninput="filterDropdown(\\'' + f + \\',this.value)" onblur="blurDropdown(\\'' + f + \\')">'
        + '  <span class="select-arrow">&#9662;</span>'
        + '  <div class="dropdown" id="drop-' + f + '"></div>'
        + '</div></div>';
    }} else if (ft === 'tag') {{
      const tags = opts.map(v =>
        '<div class="tag" data-val="' + escAttr(v) + '" onclick="selectTag(\\'' + f + \\',\\'' + escJs(v) + \\')">' + escHtml(v) + '</div>'
      ).join('');
      html += '<div class="field"><div class="field-label"><span class="req">*</span> ' + escHtml(label) + '</div>'
        + '<div class="tag-wrap" id="tags-' + f + '">' + tags + '</div></div>';
    }} else if (ft === 'custom') {{
      html += makeCustomHTML();
    }}
  }}

  container.innerHTML = '<div class="fields-grid">' + html + '</div>';
}}


function makeCustomHTML() {{
  const groups =
    '<div class="help-group"><div class="help-group-title">Search Ad</div><div class="help-tags">' +
    ['dsa','dsa-brand','dsa-product','dsa-generic','brand','product','generic','competitor'].map(v=>
      '<span class="help-tag" onclick="pickCustom(\\''+v+'\\')">'+v+'</span>').join('') +
    '</div></div>' +
    '<div class="help-group"><div class="help-group-title">Pmax Ad</div><div class="help-tags">' +
    ['brand-feed','brand-asset','nonbrand-feed','nonbrand-asset','brand-all','nonbrand-all'].map(v=>
      '<span class="help-tag" onclick="pickCustom(\\''+v+'\\')">'+v+'</span>').join('') +
    '</div></div>';

  return '<div class="field field-full">'
    + '<div class="field-label">Custom<span class="optional">可选</span></div>'
    + '<div class="custom-input-wrap">'
    + '  <input type="text" class="custom-input" id="custom-input" placeholder="自定义内容，如 rm、aon..."'
    + '    oninput="validateCustom(this)" maxlength="50">'
    + '</div>'
    + '<div class="custom-error-msg" id="custom-error"></div>'
    + '<div class="custom-rules" id="custom-rules-hint">'
    + '  <div class="rule" id="rule-lower">&#9675; 仅英文小写（不含中文及大写字母）</div>'
    + '  <div class="rule" id="rule-no-underscore">&#9675; 不能包含下划线 <code>_</code>，用 <code>-</code> 或空格连接词语</div>'
    + '</div>'
    + '<div class="custom-help">'
    + '  <div class="custom-help-header" onclick="toggleHelp()">'
    + '    &#128161; 常用 Custom 值参考'
    + '    <span class="custom-help-toggle" id="help-toggle">&#9662;</span>'
    + '  </div>'
    + '  <div class="custom-help-body" id="help-body">' + groups + '</div>'
    + '</div></div>';
}}


function escHtml(s) {{ return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }}
function escAttr(s) {{ return s.replace(/&/g,'&amp;').replace(/"/g,'&quot;').replace(/'/g,"&#39;"); }}
function escJs(s) {{ return s.replace(/\\\\/g,'\\\\\\\\').replace(/'/g,"\\\\'").replace(/"/g,'\\\\\\"').replace(/\\n/g,'\\\\n'); }}


// ===== DROPDOWN ENGINE =====
function buildDropdown(field) {{
  const opts = OPTIONS[field] || [];
  const drop = document.getElementById('drop-' + field);
  if (!drop) return;
  drop.innerHTML = opts.map(v =>
    '<div class="dropdown-item' + (sel[field]===v ? ' selected' : '') + '" data-val="' + escAttr(v) + '"'
      + ' onmousedown="pickDropdown(\\'' + field.replace(/'/g,"\\\\'") + '\\',\\'' + escJs(v) + \\',event)">' + escHtml(v) + '</div>'
  ).join('');
}}

function openDropdown(field) {{
  buildDropdown(field);
  document.getElementById('drop-' + field).classList.add('open');
  document.getElementById('search-' + field).select();
}}

function filterDropdown(field, q) {{
  const opts = OPTIONS[field] || [];
  const drop = document.getElementById('drop-' + field);
  if (!drop) return;
  drop.classList.add('open');
  const filtered = q.trim() === '' ? opts : opts.filter(v => v.toLowerCase().includes(q.toLowerCase()));
  if (filtered.length === 0) {{
    drop.innerHTML = '<div class="dropdown-item no-match">无匹配结果</div>';
  }} else {{
    drop.innerHTML = filtered.map(v =>
      '<div class="dropdown-item' + (sel[field]===v ? ' selected' : '') + '" data-val="' + escAttr(v) + '"'
        + ' onmousedown="pickDropdown(\\'' + field.replace(/'/g,"\\\\'") + '\\',\\'' + escJs(v) + \\',event)">' + escHtml(v) + '</div>'
    ).join('');
  }}
}}

function pickDropdown(field, val, e) {{
  e.preventDefault();
  sel[field] = val;
  const inp = document.getElementById('search-' + field);
  inp.value = val;
  inp.classList.add('has-value');
  document.getElementById('drop-' + field).classList.remove('open');
  updatePreview();
}}

function blurDropdown(field) {{
  setTimeout(() => {{
    const drop = document.getElementById('drop-' + field);
    if (drop) drop.classList.remove('open');
    const inp = document.getElementById('search-' + field);
    if (sel[field]) {{ inp.value = sel[field]; inp.classList.add('has-value'); }}
    else {{ inp.value = ''; inp.classList.remove('has-value'); }}
  }}, 150);
}}


// ===== TAG ENGINE =====
function selectTag(field, val) {{
  if (sel[field] === val) delete sel[field];
  else sel[field] = val;
  const container = document.getElementById('tags-' + field);
  if (container) {{
    container.querySelectorAll('.tag').forEach(t => t.classList.toggle('active', t.dataset.val === sel[field]));
  }}
  updatePreview();
}}


// ===== CUSTOM VALIDATION =====
function validateCustom(inp) {{
  const v = inp.value;
  const lowerOk = v === '' || /^[a-z0-9 \\-+]+$/.test(v);
  const noUnderscore = !v.includes('_');
  const r1 = document.getElementById('rule-lower'), r2 = document.getElementById('rule-no-underscore');
  const errEl = document.getElementById('custom-error');
  r1.className = 'rule ' + (v==='' ? '' : lowerOk ? 'pass' : 'fail');
  r1.textContent = (v===''||lowerOk ? '\u2713' : '\u2717') + ' 仅英文小写（不含中文及大写字母）';
  r2.className = 'rule ' + (v==='' ? '' : noUnderscore ? 'pass' : 'fail');
  r2.textContent = (v===''||noUnderscore ? '\u2713' : '\u2717') + ' 不能包含下划线 _，用 - 或空格连接词语';
  if (v !== '' && (!lowerOk || !noUnderscore)) {{
    inp.classList.add('error'); inp.classList.remove('ok'); customValid = false;
    const msgs=[];
    if(!lowerOk) msgs.push('仅允许英文小写、数字、连字符和空格');
    if(!noUnderscore) msgs.push('不能包含下划线 _');
    errEl.textContent=msgs.join('\uff1b'); errEl.classList.add('show');
  }} else {{
    inp.classList.remove('error');
    if(v!=='') inp.classList.add('ok'); else inp.classList.remove('ok');
    customValid=true; errEl.classList.remove('show');
  }}
  updatePreview();
}}


// ===== PREVIEW =====
function toStr(s){{ return s.trim(); }}

function buildName() {{
  const parts=[];
  for(const f of FIELDS_ORDER){{
    if(f==='custom'){{
      const v=document.getElementById('custom-input');
      const cv = v ? v.value.trim() : '';
      if(cv && customValid) parts.push({{field:f,val:toStr(cv),isCustom:true}});
    }}else{{
      if(sel[f]) parts.push({{field:f,val:toStr(sel[f])}});
    }}
  }}
  return parts;
}}

function updatePreview() {{
  const parts=buildName(), outEl=document.getElementById('preview-output'), lenEl=document.getElementById('preview-len');
  if(parts.length===0){{
    outEl.innerHTML='<span class="seg-empty">填写左侧字段以生成命名…</span>';
    lenEl.textContent=''; renderBreakdown([]); return;
  }}
  outEl.innerHTML=parts.map((p,i)=>(i>0?'<span class="seg-sep">_</span>':'')+'<span class="'+(p.isCustom?'seg-custom':'seg-part')+'">'+escHtml(p.val)+'</span>').join('');
  const fullName=parts.map(p=>p.val).join('_');
  lenEl.textContent=fullName.length+' 字符'; window._currentName=fullName;
  renderBreakdown(parts);
}}

function renderBreakdown(parts){{
  const el=document.getElementById('breakdown-list');
  if(parts.length===0){{ el.innerHTML='<div class="empty-hint">填写字段后自动解析…</div>'; return; }}
  el.innerHTML=parts.map((p,i)=>'<div class="breakdown-row"><div class="breakdown-num">'+(i+1)+'</div><div class="breakdown-field">'+escHtml(FIELD_LABELS[p.field])+'</div><div class="breakdown-val'+(p.isCustom?' custom-val':'')+'">'+escHtml(p.val)+'</div></div>').join('');
}}


// ===== COPY & EXPORT =====
function copyMain() {{
  const n=window._currentName;if(!n)return;doClipboard(n);
  const b=document.getElementById('copy-main-btn');b.textContent='\u2705 已复制！';b.classList.add('copied');
  setTimeout(()=>{{b.innerHTML='[OK] Copied';b.classList.remove('copied')}},2000);addHistory(n);
}}
function doClipboard(s){{ navigator.clipboard.writeText(s).catch(()=>{{const t=document.createElement('textarea');t.value=s;document.body.appendChild(t);t.select();document.execCommand('copy');document.body.removeChild(t)}})}}
function addHistory(n){{ copyHistory.unshift({{name:n,time:new Date().toLocaleTimeString('zh-CN',{{hour:'2-digit',minute:'2-digit'}})}});renderHistory(); }}
function renderHistory(){{
  const el=document.getElementById('history-scroller');
  if(copyHistory.length===0){{el.innerHTML='<div class="empty-hint">暂无记录…</div>';return}}
  el.innerHTML=copyHistory.slice(0,20).map(h=>'<div class="history-row" onclick="doClipboard(\\''+h.name.replace(/'/g,"\\\\'")+'\\')"><div class="history-name" title="'+escHtml(h.name)+'">'+escHtml(h.name)+'</div><div class="history-time">'+h.time+'</div><div class="history-copy-btn" onclick="event.stopPropagation();doClipboard(\\''+h.name.replace(/'/g,"\\\\'")+'\\')">COPY</div></div>').join('');
}}
function clearHistory(){{copyHistory.length=0;renderHistory()}}
function exportTXT(){{
  if(!copyHistory.length){{alert('暂无历史记录');return}}
  dlBlob('广告系列命名导出\\n导出时间：'+new Date().toLocaleString('zh-CN')+'\\n'+String.fromCharCode(8210).repeat(60)+'\\n\\n'+copyHistory.map((h,i)=>(i+1)+'. '+h.name).join('\\n'),'campaign-names-'+fmt()+'.txt','text/plain;charset=utf-8');
}}
function exportCSV(){{
  if(!copyHistory.length){{alert('暂无历史记录');return}}
  dlBlob('\\ufeff'+[['序号','命名','复制时间'],...copyHistory.map((h,i)=>[i+1,h.name,h.time])].map(r=>r.map(c=>'"'+c+'"').join(',')).join('\\n'),'campaign-names-'+fmt()+'.csv','text/csv;charset=utf-8');
}}
function downloadTool(){{dlBlob('<!DOCTYPE html>\\n'+document.documentElement.outerHTML,'ad-campaign-namebuilder.html','text/html;charset=utf-8')}}
function dlBlob(content,filename,type){{const u=URL.createObjectURL(new Blob([content],{{type}}));const a=document.createElement('a');a.href=u;a.download=filename;document.body.appendChild(a);a.click();document.body.removeChild(a);URL.revokeObjectURL(u)}}
function fmt(){{const d=new Date();return d.getFullYear()+String(d.getMonth()+1).padStart(2,'0')+String(d.getDate()).padStart(2,'0')}}

function toggleHelp(){{document.getElementById('help-body').classList.toggle('show');document.getElementById('help-toggle').classList.toggle('open')}}
function pickCustom(val){{const i=document.getElementById('custom-input');if(i){{i.value=val;validateCustom(i)}}}}

function resetAll() {{
  Object.keys(OPTIONS).forEach(f=>{{delete sel[f];const i=document.getElementById('search-'+f);if(i){{i.value='';i.classList.remove('has-value')}}}});
  document.querySelectorAll('.tag-wrap[id^="tags-"]').forEach(c=>{{const f=c.id.replace('tags-','');delete sel[f];c.querySelectorAll('.tag').forEach(t=>t.classList.remove('active'))}});
  const ci=document.getElementById('custom-input');if(ci){{ci.value='';ci.classList.remove('error','ok')}}
  const ce=document.getElementById('custom-error');if(ce)ce.classList.remove('show');
  const r1=document.getElementById('rule-lower');if(r1){{r1.className='rule';r1.textContent='\u25cb 仅英文小写（不含中文及大写字母）'}}
  const r2=document.getElementById('rule-no-underscore');if(r2){{r2.className='rule';r2.textContent='\u25cb 不能包含下划线 _，用 - 或空格连接词语'}}
  customValid=true;
  const hb=document.getElementById('help-body');if(hb)hb.classList.remove('show');
  const ht=document.getElementById('help-toggle');if(ht)ht.classList.remove('open');
  updatePreview();
}}

document.addEventListener('click',e=>{{if(!e.target.closest('.select-wrap'))document.querySelectorAll('.dropdown').forEach(d=>d.classList.remove('open'))}});


// ===== INIT =====
loadData();
"""


def generate_html(fallback_data):
    """Generate HTML with Google Sheet as primary data source, Excel data as fallback"""

    # Prepare fallback JSON (embedded in HTML for offline use)
    fallback_json = json.dumps(fallback_data, ensure_ascii=False)

    # Prepare field types JSON
    ft_dict = {}
    for col_name, ft in FIELD_TYPES.items():
        ft_dict[js_key(col_name)] = ft
    ft_json = json.dumps(ft_dict, ensure_ascii=False)

    js = JS_TEMPLATE.format(
        SHEET_CSV_URL=SHEET_CSV_URL,
        FALLBACK_JSON=fallback_json,
        FIELD_TYPES_JSON=ft_json,
    )

    # Clean surrogate characters
    clean_js = re.sub(r'[\ud800-\udfff]', '[EMOJI]', js)

    html = (
        "<!DOCTYPE html>\n<html lang=\"zh-CN\">\n<head>\n"
        "  <meta charset=\"UTF-8\">\n"
        "  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">\n"
        f"  <title>{PAGE_TITLE}</title>\n"
        f"  <style>{CSS}\n  </style>\n"
        "</head>\n<body>\n\n"
        "<div class=\"header\">\n"
        "  <div class=\"header-icon\">TAG</div>\n"
        "  <div>\n"
        f"    <div class=\"header-title\">{PAGE_TITLE}</div>\n"
        f"    <div class=\"header-sub\">{PAGE_SUBTITLE}</div>\n"
        "  </div>\n"
        "</div>\n\n"

        "<div class=\"layout\">\n\n"
        "  <div>\n"
        "    <div class=\"card\">\n"
        "      <div class=\"card-header\">\n"
        "        <div class=\"card-icon\">&#128203;</div>\n"
        "        <div>\n"
        "          <div class=\"card-title\">填写字段信息</div>\n"
        "          <div class=\"card-desc\">每个字段选择枚举值，或在 Custom 列自定义填写</div>\n"
        "        </div>\n"
        "      </div>\n\n"

        # Data source status bar
        '      <div id="datasource-status" class="datasource-bar loading">\n'
        '        <span class="ds-dot"></span>\n'
        '        <span class="ds-msg">正在加载数据...</span>\n'
        '        <span class="ds-refresh" onclick="loadData()">[REFRESH]</span>\n'
        "      </div>\n\n"

        # Dynamic fields container (rendered by JS after data loads)
        f"      <div id=\"fields-container\">\n"
        '        <div style="padding:40px 0;text-align:center">\n'
        '          <div class="skeleton" style="height:44px;width:100%;margin-bottom:16px"></div>\n'
        '          <div class="skeleton" style="height:44px;width:48%;display:inline-block;margin-right:2%"></div>\n'
        '          <div class="skeleton" style="height:44px;width:48%;display:inline-block"></div>\n'
        '          <div class="skeleton" style="height:44px;width:100%;margin-top:16px"></div>\n'
        "        </div>\n"
        "      </div>\n\n"

        "    </div>\n\n"

        '    <div style="display:flex;justify-content:flex-end;margin-top:12px">\n'
        '      <button class="btn-secondary" onclick="resetAll()" style="padding:8px 18px;border-radius:8px">\n'
        '        &#x1F504; 重置所有字段\n'
        "      </button>\n"
        "    </div>\n"
        "  </div>\n\n"

        "  <div class=\"result-sticky\">\n"
        "    <div class=\"card\">\n"
        "      <div class=\"card-header\">\n"
        "        <div class=\"card-icon\">&#9733;</div>\n"
        "        <div>\n"
        "          <div class=\"card-title\">命名预览</div>\n"
        '          <div class="card-desc">字段用 <code>_</code> 连接，字段内空格保持原样</div>\n'
        "        </div>\n"
        "      </div>\n\n"

        '      <div class="result-dark">\n'
        '        <div class="result-dark-label">生成结果</div>\n'
        '        <div class="result-name-display" id="preview-output">\n'
        '          <span class="seg-empty">加载中…</span>\n'
        "        </div>\n"
        '        <div class="result-len" id="preview-len"></div>\n'
        "      </div>\n"

        '      <button class="btn-primary" id="copy-main-btn" onclick="copyMain()">\n'
        '        [COPY] 复制命名\n'
        "      </button>\n"
        '      <div class="btn-row">\n'
        '        <button class="btn-secondary" onclick="exportTXT()">[TXT] 导出 TXT</button>\n'
        '        <button class="btn-secondary" onclick="exportCSV()">[CSV] 导出 CSV</button>\n'
        "      </div>\n"
        '      <button class="btn-download" onclick="downloadTool()">[DL] 下载此工具（离线 / 发给同事）</button>\n'
        "    </div>\n\n"

        "    <div class=\"card\" style=\"margin-top:14px\">\n"
        '      <div class="card-header" style="margin-bottom:14px;padding-bottom:12px">\n'
        '        <div class="card-icon">&#128269;</div>\n'
        '        <div class="card-title">字段解析</div>\n'
        "      </div>\n"
        '      <div class="breakdown-list" id="breakdown-list">\n'
        '        <div class="empty-hint">填写字段后自动解析…</div>\n'
        "      </div>\n"
        "    </div>\n"

        "    <div class=\"card\" style=\"margin-top:14px\">\n"
        '      <div class="card-header" style="margin-bottom:14px;padding-bottom:12px;justify-content:space-between">\n'
        '        <div style="display:flex;align-items:center;gap:10px">\n'
        '          <div class="card-icon">&#128337;</div>\n'
        '          <div class="card-title">复制历史</div>\n'
        "        </div>\n"
        '        <span style="font-size:12px;color:var(--muted);cursor:pointer" onclick="clearHistory()">清除</span>\n'
        "      </div>\n"
        '      <div class="history-scroller" id="history-scroller">\n'
        '        <div class="empty-hint">暂无记录…</div>\n'
        "      </div>\n"
        "    </div>\n"
        "  </div>\n"
        "</div>\n\n"

        f"<script>\n{clean_js}\n</script>\n</body>\n</html>"
    )
    return html


def main():
    print(f"[READ] Excel (fallback): {EXCEL_PATH}")
    if not EXCEL_PATH.exists():
        print(f"[ERR] File not found: {EXCEL_PATH}")
        exit(1)

    fallback_data = load_data_from_excel()
    total = sum(len(v) for v in fallback_data.values())
    print(f"   发现 {len(fallback_data)} 个字段，共 {total} 个枚举值（作为离线备用数据嵌入）")
    icons = {"dropdown": "[DROP]", "tag": "[TAG]", "custom": "[IN]"}
    for k, v in fallback_data.items():
        ft = FIELD_TYPES.get(k, "tag")
        print(f"   [{icons.get(ft,'?')}] {k}: {len(v)} 个值")

    print(f"\n[GEN] Generating HTML (Google Sheets mode)...")
    print(f"   Sheet URL: {SHEET_CSV_URL}")
    html = generate_html(fallback_data)

    OUTPUT_HTML.write_text(html, encoding="utf-8")
    size_kb = OUTPUT_HTML.stat().st_size / 1024
    print(f"[OK] Output: {OUTPUT_HTML} ({size_kb:.0f} KB)")
    print(f"\n[TIP] 维护流程：")
    print(f"   1. 编辑 Google Sheet → 刷新页面即可生效")
    print(f"   2. 如需更新离线备用数据 → 运行 python generate_tool.py")


if __name__ == "__main__":
    main()
