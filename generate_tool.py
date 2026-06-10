"""
广告命名工具生成器（GitHub Pages 版）v9.1 - 级联映射稳健版
用法：python generate_tool.py
读取 Excel → 生成 index.html + data.json
更新点：
1. 精准嵌入 product-code-map.json 的矩阵数据，保证生成后的 index.html 完美具备联动筛选功能。
2. 当选择不同的 Product Type 时，Product Code 的下拉列表会自动呈现对应的过滤子集。
3. 切换类别导致当前型号不匹配时，下级型号框会自动干净清空，防止出现“脏命名”。
"""

import json, os, sys, re
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')

try:
    import pandas as pd
except ImportError:
    print("需要安装依赖：pip install pandas openpyxl")
    exit(1)

# ===== 配置 =====
EXCEL_PATH = Path(r"D:/serein/工作/媒介建设/ad campaign namebuilder.xlsx")
OUTPUT_DIR = Path(__file__).parent
OUTPUT_HTML = OUTPUT_DIR / "index.html"
OUTPUT_DATA = OUTPUT_DIR / "data.json"

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

PAGE_TITLE = "Ad Campaign 命名生成器 · Campaign Name Builder"
PAGE_SUBTITLE = "规范快速生成广告系列命名，支持复制 / 导出 · 数据来源：data.json（GitHub 在线维护）"

# ===== 你的 product-code-map.json 数据结构 =====
PRODUCT_RELATION_MAP = {
  "all": ["all"],
  "poe": ["all", "b1200", "b400", "b500", "b800", "c1", "c2 pro", "cx410", "cx410c", "cx810", "d1200", "d400", "d500", "d800", "e1 outdoor poe", "e1 outdoor se poe", "fe-p", "nvs8-5kb4", "nvs8-5kd4", "duo 2 poe", "duo 2v poe", "duo 3 poe", "duo 3t poe", "duo 3v poe", "duo floodlight poe", "duo poe", "trackmix poe", "rlc-1010a", "rlc-1020a", "rlc-1210a", "rlc-1212a", "rlc-1220a", "rlc-1224a", "rlc-1240a", "rlc-410", "rlc-410s", "rlc-411", "rlc-420", "rlc-422", "rlc-423", "rlc-510a", "rlc-511", "rlc-520", "rlc-520a", "rlc-522", "rlc-540a", "rlc-810a", "rlc-810wa", "rlc-811a", "rlc-811wa", "rlc-812a", "rlc-81ma", "rlc-81pa", "rlc-820a", "rlc-822a", "rlc-823a", "rlc-823a 16x", "rlc-823s1", "rlc-823s2", "rlc-824a", "rlc-830a", "rlc-833a", "rlc-840a", "rlc-842a", "rlc-843a", "rlk16-1200b8", "rlk16-1200d8", "rlk16-410b4d4", "rlk16-410b8", "rlk16-520b4d4", "rlk16-800b8", "rlk16-800d8", "rlk16-800pt8", "rlk16-810b8", "rlk16-812b8", "rlk16-820d8", "rlk16-833d8", "rlk16-843v8", "rlk4-410b4", "rlk8-1200b4", "rlk8-1200d4", "rlk8-1200v4", "rlk8-1210b4", "rlk8-410b2d2", "rlk8-410b4", "rlk8-410b6", "rlk8-420d4", "rlk8-500v4", "rlk8-510b4", "rlk8-520b2d2", "rlk8-520d4", "rlk8-800b2d2", "rlk8-800b4", "rlk8-800b6", "rlk8-800d4", "rlk8-800pt4", "rlk8-800tm4", "rlk8-800v4", "rlk8-810b2d2", "rlk8-810b4", "rlk8-810b6", "rlk8-811b4", "rlk8-812b4", "rlk8-820d4", "rlk8-824d4", "rlk8-833d4", "rlk8-842d4", "rlk8-843v4", "rlk8-cx410b4", "v1200", "v500", "v800", "cx820", "elite xpro poe", "elite pro floodlight poe", "rlk16-811b8", "rp-pct16md", "rp-pct8mz"],
  "wifi": ["all", "b500w", "b800w", "cx410w", "e1", "e1 outdoor", "e1 outdoor cx", "e1 outdoor pro", "e1 outdoor s", "e1 pro", "e1 zoom", "fe-w", "e1 outdoor pro+hub pro", "duo 2 wifi", "duo floodlight wifi", "duo wifi", "lumus", "trackmix wifi", "rlc-210w", "rlc-410w", "rlc-410ws", "rlc-411ws", "rlc-422w", "rlc-510wa", "rlc-511w", "rlc-511wa", "rlc-523wa", "rlc-542wa", "rlc-810wa", "rlc-811wa", "rlc-823s1w", "rlc-840wa", "rlc-843wa", "rlc-843wa-c", "rlk12-500wb4", "rlk12-800wb4", "rlk12-800wpt4", "rlk12-800wtm2", "rlk12-800wv4", "rlk4-210wb2", "rlk4-210wb4", "rlk4-211wb4", "v800w", "duo 3 wifi", "elite wifi", "elite floodlight wifi", "trackflex floodlight wifi"],
  "battery wifi": ["all", "argus 3 ultra", "argus 3e", "argus 4", "argus 4 pro", "argus eco pro", "argus eco ultra", "argus pt lite", "argus pt ultra", "argus track", "argus 3 ultra+hub", "argus 3e+hub", "argus 4+hub", "argus 4 pro+hub", "argus eco ultra+hub", "argus pt ultra+hub", "altas pt", "altas pt ultra", "argus", "argus 2", "argus 2e", "argus 3", "argus 3 pro", "argus eco", "argus pro", "argus pt", "duo", "duo 2", "trackmix", "altas", "home hub mini+argus pt", "altas pt ultra+hub"],
  "battery 4g": ["all", "keen ranger pt", "duo 2 lte", "duo 4g", "go", "go plus", "go pt", "go pt plus", "go pt ultra", "go ranger pt", "go ultra", "trackmix lte", "trackmix lte plus", "trackmix wired lte", "talon pro", "trackmix lte plus 2"],
  "doorbell": ["all", "doorbell battery", "doorbell battery+hub", "doorbell poe", "doorbell wifi", "home hub mini+doorbell"],
  "nvr": ["all", "home hub", "home hub pro", "rln12w", "rln16-410", "rln36", "rln8-410"]
}


def js_key(name):
    return name.lower().replace(" ", "-")


def load_data_from_excel():
    df = pd.read_excel(EXCEL_PATH, sheet_name=0, header=0)
    data = {}
    for col in df.columns:
        vals = df[col].dropna().astype(str).str.strip().tolist()
        data[str(col)] = vals
    return data


# ===== CSS =====
CSS = r"""
:root{--primary:#4F46E5;--primary-light:#EEF2FF;--primary-hover:#4338CA;--success:#10B981;--danger:#EF4444;--danger-light:#FEF2F2;--border:#E5E7EB;--text:#111827;--muted:#6B7280;--bg:#F3F4F6;--card:#FFF;--shadow:0 1px 3px rgba(0,0,0,.08)}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI','PingFang SC','Microsoft YaHei',sans-serif;background:var(--bg);color:var(--text);min-height:100vh}
.header{background:var(--primary);padding:18px 28px;display:flex;align-items:center;gap:14px}
.header-icon{font-size:26px;line-height:1}
.header-title{font-size:18px;font-weight:700;color:#fff}
.header-sub{font-size:12px;color:rgba(255,255,255,.7);margin-top:2px}
.layout{display:grid;grid-template-columns:1fr 400px;gap:20px;max-width:1200px;margin:24px auto;padding:0 20px 40px;align-items:start}
@media(max-width:960px){.layout{grid-template-columns:1fr}}
.card{background:var(--card);border:1px solid var(--border);border-radius:12px;padding:22px;box-shadow:var(--shadow)}
.card+.card{margin-top:16px}
.card-header{display:flex;align-items:center;gap:10px;margin-bottom:20px;padding-bottom:14px;border-bottom:1px solid var(--border)}
.card-icon{width:32px;height:32px;border-radius:8px;background:var(--primary-light);display:flex;align-items:center;justify-content:center;font-size:16px;flex-shrink:0}
.card-title{font-size:15px;font-weight:700}
.card-desc{font-size:12px;color:var(--muted);margin-top:1px}
.fields-grid{display:grid;grid-template-columns:1fr 1fr;gap:16px}
@media(max-width:600px){.fields-grid{grid-template-columns:1fr}}
.field-full{grid-column:1/-1}
.field-compact{display:contents}
.field-compact .field{margin-bottom:0}
.field-compact .field-label{font-size:11px}
.field-compact .select-search{padding:7px 32px 7px 10px;font-size:13px}
.field{display:flex;flex-direction:column;gap:6px}
.field-label{font-size:12px;font-weight:600;color:var(--muted);text-transform:uppercase;letter-spacing:.05em;display:flex;align-items:center;gap:5px}
.field-label .req{color:var(--danger);font-size:14px;line-height:1}
.field-label .optional{font-size:10px;font-weight:500;text-transform:none;color:var(--muted);background:#F3F4F6;padding:1px 5px;border-radius:4px}
.tag-wrap{display:flex;flex-wrap:wrap;gap:7px;padding:10px;border:1.5px solid var(--border);border-radius:10px;min-height:44px;cursor:default;transition:border-color .15s}
.tag{display:inline-flex;align-items:center;padding:4px 10px;border:1.5px solid var(--border);border-radius:6px;font-size:12px;font-weight:500;cursor:pointer;transition:all .12s;color:var(--muted);background:var(--bg);user-select:none;white-space:nowrap}
.tag:hover{border-color:var(--primary);color:var(--primary);background:var(--primary-light)}
.tag.active{border-color:var(--primary);background:var(--primary);color:#fff;font-weight:600}
.tag.disabled{opacity:.35;pointer-events:none;cursor:not-allowed}
.select-wrap{position:relative}
.select-search{width:100%;padding:9px 36px 9px 12px;border:1.5px solid var(--border);border-radius:10px;font-size:14px;color:var(--text);background:var(--card);transition:border-color .15s,box-shadow .15s;cursor:pointer}
.select-search::placeholder{color:#9CA3AF;font-size:13px}
.select-search:focus{outline:none;border-color:var(--primary);box-shadow:0 0 0 3px rgba(79,70,229,.1)}
.select-search.has-value{color:var(--primary);font-weight:600}
.select-arrow{position:absolute;right:10px;top:50%;transform:translateY(-50%);pointer-events:none;color:var(--muted);font-size:12px}
.dropdown{position:absolute;top:calc(100%+4px);left:0;right:0;background:var(--card);border:1.5px solid var(--border);border-radius:10px;box-shadow:0 8px 24px rgba(0,0,0,.08);max-height:240px;overflow-y:auto;z-index:100;display:none}
.dropdown.open{display:block}
.dropdown-item{padding:9px 14px;font-size:13px;cursor:pointer;transition:background .1s;color:var(--text)}
.dropdown-item:hover{background:var(--primary-light);color:var(--primary)}
.dropdown-item.selected{background:var(--primary-light);color:var(--primary);font-weight:600}
.dropdown-item.no-match{color:var(--muted);font-style:italic;pointer-events:none;padding:12px;text-align:center}
.custom-input-wrap{position:relative}
.custom-input{width:100%;padding:9px 12px;border:1.5px solid var(--border);border-radius:10px;font-size:14px;color:var(--text);transition:border-color .15s,box-shadow .15s}
.custom-input:focus{outline:none;border-color:var(--primary);box-shadow:0 0 0 3px rgba(79,70,229,.1)}
.custom-input.ok{border-color:var(--success);background:#F0FDF4}
.custom-rules{font-size:11px;color:var(--muted);margin-top:5px;line-height:1.5;background:#F9FAFB;padding:8px 12px;border-radius:8px;border:1px dashed var(--border)}
.custom-rules .rule{display:flex;align-items:center;gap:6px;color:var(--success);font-weight:600}
.custom-help{margin-top:12px;border:1px solid var(--border);border-radius:10px;overflow:hidden}
.custom-help-header{display:flex;align-items:center;justify-content:space-between;padding:9px 14px;background:#F3F4F6;cursor:pointer;user-select:none;font-size:12px;font-weight:600;color:var(--muted);transition:background .12s}
.custom-help-header:hover{background:#E5E7EB}
.custom-help-toggle{font-size:11px;transition:transform .2s}
.custom-help-toggle.open{transform:rotate(180deg)}
.custom-help-body{display:none;padding:12px 14px 14px}
.custom-help-body.show{display:block}
.help-group{margin-bottom:12px}
.help-group:last-child{margin-bottom:0}
.help-group-title{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:var(--primary);margin-bottom:7px}
.help-tags{display:flex;flex-wrap:wrap;gap:6px}
.help-tag{padding:4px 10px;border:1px solid var(--border);border-radius:6px;font-size:11px;font-weight:500;cursor:pointer;transition:all .12s;color:var(--text);background:#fff}
.help-tag:hover{border-color:var(--primary);color:var(--primary);background:var(--primary-light)}
.result-sticky{position:sticky;top:20px}
.result-dark{background:#18181B;border-radius:12px;padding:20px 20px 16px;margin-bottom:14px}
.result-dark-label{font-size:11px;text-transform:uppercase;letter-spacing:.08em;color:#71717A;margin-bottom:10px;font-weight:600}
.result-name-display{font-size:15px;font-weight:700;color:#F4F4F5;word-break:break-all;line-height:1.6;min-height:28px}
.seg-part{color:#A78BFA}
.seg-sep{color:#3F3F46}
.seg-empty{color:#52525B;font-style:italic;font-size:13px;font-weight:400}
.seg-custom{color:#34D399}
.result-len{font-size:11px;color:#52525B;margin-top:8px;text-align:right}
.btn-primary{width:100%;padding:11px;border-radius:10px;font-size:14px;font-weight:700;cursor:pointer;border:none;background:var(--primary);color:#fff;transition:background .15s;display:flex;align-items:center;justify-content:center;gap:8px}
.btn-primary:hover{background:var(--primary-hover)}
.btn-primary.copied{background:var(--success)}
.btn-primary:disabled{opacity:.5;cursor:not-allowed}
.btn-row{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-top:10px}
.btn-secondary{padding:9px 10px;border-radius:10px;font-size:13px;font-weight:500;cursor:pointer;border:1.5px solid var(--border);background:var(--card);color:var(--text);transition:all .15s;display:flex;align-items:center;justify-content:center;gap:5px}
.btn-secondary:hover{border-color:var(--primary);color:var(--primary);background:var(--primary-light)}
.breakdown-list{display:flex;flex-direction:column;gap:7px}
.breakdown-row{display:flex;align-items:center;gap:10px;padding:8px 10px;border-radius:8px;background:var(--bg)}
.breakdown-num{width:20px;height:20px;border-radius:5px;background:var(--primary-light);color:var(--primary);font-size:10px;font-weight:700;display:flex;align-items:center;justify-content:center;flex-shrink:0}
.breakdown-field{font-size:12px;color:var(--muted);flex:1}
.breakdown-val{font-size:13px;font-weight:600;color:var(--text)}
.breakdown-val.custom-val{color:var(--success)}
.empty-hint{font-size:13px;color:var(--muted);text-align:center;padding:14px 0}
.validation-warning{background:#FEF2F2;border:1px solid #FECACA;border-radius:8px;padding:10px 14px;margin-top:12px;font-size:12px;color:#991B1B;display:none}
.validation-warning.show{display:block;font-weight:500}
.history-scroller{max-height:200px;overflow-y:auto}
.history-row{display:flex;align-items:center;gap:8px;padding:8px 10px;border-radius:8px;cursor:pointer;transition:background .1s}
.history-row:hover{background:var(--primary-light)}
.history-name{font-size:12px;font-weight:500;flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:var(--primary)}
.history-row:hover .history-name{text-decoration:underline}
.history-time{font-size:10px;color:var(--muted);flex-shrink:0}
.history-copy-btn{width:50px;height:26px;border-radius:6px;border:1px solid var(--border);background:var(--card);font-size:11px;font-weight:600;cursor:pointer;display:flex;align-items:center;justify-content:center;flex-shrink:0;transition:all .1s}
.history-copy-btn:hover{background:var(--primary);border-color:var(--primary);color:#fff}
::-webkit-scrollbar{width:4px;height:4px}::-webkit-scrollbar-thumb{background:#D1D5DB;border-radius:2px}
.datasource-bar{display:flex;align-items:center;gap:8px;padding:8px 14px;border-radius:8px;font-size:11px;margin-bottom:16px}
.datasource-bar.ok{background:#ECFDF5;border:1px solid #A7F3D0;color:#065F46}
.datasource-bar.error{background:#FEF2F2;border:1px solid #FECACA;color:#991B1B}
.datasource-bar.loading{background:#FFFBEB;border:1px solid #FDE68A;color:#92400E}
.datasource-bar .ds-dot{width:8px;height:8px;border-radius:50%;flex-shrink:0}
.datasource-bar.ok .ds-dot{background:#10B981}
.datasource-bar.error .ds-dot{background:#EF4444}
.datasource-bar.loading .ds-dot{background:#F59E0B;animation:pulse-dot 1s infinite}
.datasource-bar .ds-msg{flex:1}
.datasource-bar .ds-refresh{cursor:pointer;font-size:11px;font-weight:600;white-space:nowrap;opacity:.7;transition:opacity .15s}
.datasource-bar .ds-refresh:hover{opacity:1}
@keyframes pulse-dot{0%,100%{opacity:1}50%{opacity:.3}}
.skeleton{background:linear-gradient(90deg,#f0f0f0 25%,#e0e0e0 50%,#f0f0f0 75%);background-size:200% 100%;animation:skeleton-shimmer 1.5s infinite;border-radius:8px}
@keyframes skeleton-shimmer{0%{background-position:-200% 0}100%{background-position:200% 0}}
"""

# ===== JS TEMPLATE =====
JS_TEMPLATE = r'''
// ===== CONFIG =====
const DATA_URL = './data.json';
const FALLBACK_DATA = __FALLBACK_JSON__;
const FIELD_TYPES = __FIELD_TYPES_JSON__;
const PRODUCT_RELATION_MAP = __PRODUCT_RELATION_MAP__;

let FIELDS_ORDER = [];
let FIELD_LABELS = {};
let OPTIONS = {};
let DATA_READY = false;

const REQUIRED_FIELDS = new Set();

const FUNNEL_GOAL_MAP = {
  'awareness':     ['impression','reach'],
  'consideration': ['click','view','pageview','signup','atc','checkout','engagement','dpv','follow'],
  'conversion':    ['sales']
};

const sel = {};
const copyHistory = [];


// ===== DATA LOADING =====
let _loadTimeout = null;

async function loadData() {
  showDatasourceStatus('loading', 'Loading data...');

  if (_loadTimeout) clearTimeout(_loadTimeout);
  _loadTimeout = setTimeout(function() {
    if (!DATA_READY) {
      console.warn('Data loading timed out, using embedded fallback');
      loadFallback();
      showDatasourceStatus('error', 'Loading timed out. Using built-in fallback data. Click [REFRESH] to retry.');
    }
  }, 4000);

  try {
    const resp = await fetch(DATA_URL + '?t=' + Date.now());
    if (!resp.ok) throw new Error('HTTP ' + resp.status);
    const json = await resp.json();

    FIELDS_ORDER = [];
    FIELD_LABELS = {};
    OPTIONS = {};
    REQUIRED_FIELDS.clear();

    for (const [colName, vals] of Object.entries(json)) {
      const key = colName.toLowerCase().replace(/ /g, '-');
      FIELDS_ORDER.push(key);
      FIELD_LABELS[key] = colName;
      OPTIONS[key] = vals;
      if (colName !== 'Custom') REQUIRED_FIELDS.add(key);
    }

    renderAllFields();
    DATA_READY = true;
    if (_loadTimeout) { clearTimeout(_loadTimeout); _loadTimeout = null; }
    const totalVals = Object.values(OPTIONS).reduce((a,b)=>a+b.length,0);
    showDatasourceStatus('ok', 'Data loaded (' + Object.keys(OPTIONS).length + ' fields, ' + totalVals + ' values)');
  } catch(err) {
    console.warn('Data fetch failed, using embedded fallback:', err);
    if (_loadTimeout) { clearTimeout(_loadTimeout); _loadTimeout = null; }
    loadFallback();
    showDatasourceStatus('error', 'Failed to load data.json. Using built-in fallback.');
  }
}

function loadFallback() {
  FIELDS_ORDER = [];
  FIELD_LABELS = {};
  OPTIONS = {};
  REQUIRED_FIELDS.clear();
  for (const [colName, vals] of Object.entries(FALLBACK_DATA)) {
    const key = colName.toLowerCase().replace(/ /g, '-');
    FIELDS_ORDER.push(key);
    FIELD_LABELS[key] = colName;
    OPTIONS[key] = vals;
    if (colName !== 'Custom') REQUIRED_FIELDS.add(key);
  }
  renderAllFields();
  DATA_READY = true;
}

function showDatasourceStatus(state, msg) {
  const el = document.getElementById('datasource-status');
  if(!el) return;
  el.className = 'datasource-bar ' + state;
  el.innerHTML = '<span class="ds-dot"></span><span class="ds-msg">' + msg + '</span><span class="ds-refresh" onclick="loadData()">[REFRESH]</span>';
}


// ===== RENDER FIELDS =====
function renderAllFields() {
  const container = document.getElementById('fields-container');
  let html = '';

  for (let i = 0; i < FIELDS_ORDER.length; i++) {
    const f = FIELDS_ORDER[i];
    if (f === 'market') {
      html += '<div class="field-compact">';
    }

    const ft = FIELD_TYPES[f] || 'tag';
    const label = FIELD_LABELS[f];
    const opts = OPTIONS[f] || [];

    if (ft === 'dropdown') {
      const isFullWidth = (f === 'product-code' || opts.length > 30);
      let ph = 'Search ' + label + '...';
      if(f === 'market') ph = 'Search or select market...';
      if(f === 'product-type') ph = 'Search product type...';
      if(f === 'product-code') ph = 'Search product model (e.g. argus 4, rlc-410)...';

      html += '<div class="field' + (isFullWidth ? ' field-full' : '') + '">';
      html += '<div class="field-label"><span class="req">*</span> ' + escHtml(label) + '</div>';
      html += '<div class="select-wrap" id="wrap-' + f + '">';
      html += '<input class="select-search" id="search-' + f + '" placeholder="' + escHtml(ph) + '" autocomplete="off" ';
      html += 'onfocus="openDropdown(\'' + f + '\')" oninput="filterDropdown(\'' + f + '\',this.value)" onblur="blurDropdown(\'' + f + '\')">';
      html += '<span class="select-arrow">&#9662;</span>';
      html += '<div class="dropdown" id="drop-' + f + '"></div>';
      html += '</div></div>';
    } else if (ft === 'tag') {
      const tags = renderTagOptions(f, opts);
      html += '<div class="field"><div class="field-label"><span class="req">*</span> ' + escHtml(label) + '</div>';
      html += '<div class="tag-wrap" id="tags-' + f + '">' + tags + '</div></div>';
    } else if (ft === 'custom') {
      if (FIELDS_ORDER.includes('market')) {
        html += '</div>'; 
      }
      html += makeCustomHTML();
    }

    if (f === 'product-type') {
      html += '</div>'; 
    }
  }

  container.innerHTML = '<div class="fields-grid">' + html + '</div>';
  applyFunnelFilter();
}

function renderTagOptions(fieldKey, opts) {
  let disabledGoals = null;
  if (fieldKey === 'bidding-goal' && sel['media-funnel']) {
    disabledGoals = FUNNEL_GOAL_MAP[sel['media-funnel']] || null;
  }

  let html = '';
  for(let i=0; i<opts.length; i++) {
    const v = opts[i];
    const isDisabled = disabledGoals && !disabledGoals.includes(v);
    const cName = 'tag' + (isDisabled ? ' disabled' : '') + (sel[fieldKey]===v ? ' active' : '');
    const clickEvt = isDisabled ? '' : 'selectTag(\'' + fieldKey + '\', \'' + escJs(v) + '\')';
    html += '<div class="' + cName + '" data-val="' + escAttr(v) + '" onclick="' + clickEvt + '">' + escHtml(v) + '</div>';
  }
  return html;
}

function applyFunnelFilter() {
  const funnelVal = sel['media-funnel'];
  const goalContainer = document.getElementById('tags-bidding-goal');
  if (!goalContainer) return;

  const goalOpts = OPTIONS['bidding-goal'] || [];
  goalContainer.innerHTML = renderTagOptions('bidding-goal', goalOpts);

  if (sel['bidding-goal']) {
    const allowed = FUNNEL_GOAL_MAP[funnelVal];
    if (allowed && !allowed.includes(sel['bidding-goal'])) {
      delete sel['bidding-goal'];
      updatePreview();
    } else {
      goalContainer.querySelectorAll('.tag').forEach(function(t) {
        t.classList.toggle('active', t.dataset.val === sel['bidding-goal']);
      });
    }
  } else {
    updatePreview();
  }
}


// ===== CUSTOM SECTION =====
function makeCustomHTML() {
  const searchAds = ['dsa','dsa-brand','dsa-product','dsa-generic','brand','product','generic','competitor'];
  const pmaxAds = ['brand-feed','brand-asset','nonbrand-feed','nonbrand-asset','brand-all','nonbrand-all'];
  
  let sHtml = '';
  for(let i=0; i<searchAds.length; i++) {
    sHtml += '<span class="help-tag" onclick="pickCustom(\'' + searchAds[i] + '\')">' + searchAds[i] + '</span>';
  }
  let pHtml = '';
  for(let i=0; i<pmaxAds.length; i++) {
    pHtml += '<span class="help-tag" onclick="pickCustom(\'' + pmaxAds[i] + '\')">' + pmaxAds[i] + '</span>';
  }

  let groups = '<div class="help-group"><div class="help-group-title">Search Ad</div><div class="help-tags">' + sHtml + '</div></div>';
  groups += '<div class="help-group"><div class="help-group-title">Pmax Ad</div><div class="help-tags">' + pHtml + '</div></div>';

  let html = '<div class="field field-full">';
  html += '<div class="field-label">Custom<span class="optional">Optional</span></div>';
  html += '<div class="custom-input-wrap"><input type="text" class="custom-input" id="custom-input" placeholder="Auto-fixes uppercase, spaces & underscores..." oninput="handleCustomInput(this)" maxlength="50"></div>';
  html += '<div class="custom-rules" id="custom-rules-hint"><div class="rule">&#9888; Smart Guard Active: Auto-converts to lowercase & hyphens. Cleans illegal tags.</div></div>';
  html += '<div class="custom-help">';
  html += '<div class="custom-help-header" onclick="toggleHelp()">💡 Common Custom Values Reference<span class="custom-help-toggle open" id="help-toggle">&#9662;</span></div>';
  html += '<div class="custom-help-body" id="help-body">' + groups + '</div>';
  html += '</div></div>';
  return html;
}

function handleCustomInput(inp) {
  let v = inp.value;
  v = v.toLowerCase();
  v = v.replace(/[_\s]+/g, '-');
  v = v.replace(/[^a-z0-9\-+]/g, '');
  inp.value = v;
  
  if(v !== '') inp.classList.add('ok'); else inp.classList.remove('ok');
  updatePreview();
}

function toggleHelp() {
  const body = document.getElementById('help-body');
  const toggle = document.getElementById('help-toggle');
  if(body.style.display === 'none' || body.classList.contains('collapsed')) {
    body.style.display = 'block';
    body.classList.remove('collapsed');
    toggle.classList.add('open');
  } else {
    body.style.display = 'none';
    body.classList.add('collapsed');
    toggle.classList.remove('open');
  }
}

function escHtml(s) { return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }
function escAttr(s) { return s.replace(/&/g,'&amp;').replace(/"/g,'&quot;').replace(/'/g,"&#39;"); }
function escJs(s) { return s.replace(/\\/g,'\\\\').replace(/'/g,"\\'").replace(/"/g,'\\"').replace(/\n/g,'\\n'); }


// ===== DROPDOWN ENGINE (联动交叉拦截) =====
function getFilteredProductCodes() {
  const allCodes = OPTIONS['product-code'] || [];
  const currentType = sel['product-type'];
  if (!currentType) return allCodes;
  
  const allowedSubset = PRODUCT_RELATION_MAP[currentType] || [];
  return allCodes.filter(function(code) {
    return allowedSubset.includes(code.toLowerCase());
  });
}

function buildDropdown(field) {
  const drop = document.getElementById('drop-' + field);
  if (!drop) return;
  
  const opts = (field === 'product-code') ? getFilteredProductCodes() : (OPTIONS[field] || []);
  
  let html = '';
  for(let i=0; i<opts.length; i++) {
    const v = opts[i];
    const isSelected = sel[field]===v ? ' selected' : '';
    html += '<div class="dropdown-item' + isSelected + '" data-val="' + escAttr(v) + '" ';
    html += 'onmousedown="pickDropdown(\'' + field.replace(/'/g,"\\'") + '\', \'' + escJs(v) + '\', event)">' + escHtml(v) + '</div>';
  }
  drop.innerHTML = html;
}

function openDropdown(field) {
  buildDropdown(field);
  document.getElementById('drop-' + field).classList.add('open');
  document.getElementById('search-' + field).select();
}

function filterDropdown(field, q) {
  const drop = document.getElementById('drop-' + field);
  if (!drop) return;
  drop.classList.add('open');
  
  const baseOpts = (field === 'product-code') ? getFilteredProductCodes() : (OPTIONS[field] || []);
  const filtered = q.trim() === '' ? baseOpts : baseOpts.filter(function(v) { return v.toLowerCase().includes(q.toLowerCase()); });
  
  if (filtered.length === 0) {
    drop.innerHTML = '<div class="dropdown-item no-match">No matches found</div>';
  } else {
    let html = '';
    for(let i=0; i<filtered.length; i++) {
      const v = filtered[i];
      const isSelected = sel[field]===v ? ' selected' : '';
      html += '<div class="dropdown-item' + isSelected + '" data-val="' + escAttr(v) + '" ';
      html += 'onmousedown="pickDropdown(\'' + field.replace(/'/g,"\\'") + '\', \'' + escJs(v) + '\', event)">' + escHtml(v) + '</div>';
    }
    drop.innerHTML = html;
  }
}

function pickDropdown(field, val, e) {
  if(e) e.preventDefault();
  sel[field] = val;
  const inp = document.getElementById('search-' + field);
  if(inp) {
    inp.value = val;
    inp.classList.add('has-value');
  }
  const drop = document.getElementById('drop-' + field);
  if(drop) drop.classList.remove('open');
  
  if (field === 'product-type') {
    const allowedSubset = PRODUCT_RELATION_MAP[val] || [];
    if (sel['product-code'] && !allowedSubset.includes(sel['product-code'].toLowerCase())) {
      delete sel['product-code'];
      const codeInp = document.getElementById('search-product-code');
      if (codeInp) {
        codeInp.value = '';
        codeInp.classList.remove('has-value');
      }
    }
  }
  updatePreview();
}

function blurDropdown(field) {
  setTimeout(function() {
    const drop = document.getElementById('drop-' + field);
    if (drop) drop.classList.remove('open');
    const inp = document.getElementById('search-' + field);
    if(!inp) return;
    if (sel[field]) { inp.value = sel[field]; inp.classList.add('has-value'); }
    else { inp.value = ''; inp.classList.remove('has-value'); }
  }, 150);
}


// ===== TAG ENGINE =====
function selectTag(field, val) {
  if (sel[field] === val) delete sel[field];
  else sel[field] = val;

  const container = document.getElementById('tags-' + field);
  if (container) {
    container.querySelectorAll('.tag').forEach(function(t) {
      t.classList.toggle('active', t.dataset.val === sel[field]);
    });
  }

  if (field === 'media-funnel') {
    applyFunnelFilter();
  } else {
    updatePreview();
  }
}


// ===== VALIDATION =====
function validateRequired() {
  const missing = [];
  FIELDS_ORDER.forEach(function(f) {
    if (REQUIRED_FIELDS.has(f) && !sel[f]) {
      missing.push(FIELD_LABELS[f] || f);
    }
  });
  return missing;
}

function showValidationWarning(msg) {
  const el = document.getElementById('validation-warning');
  if(el) { el.textContent = msg; el.classList.add('show'); }
}

function hideValidationWarning() {
  const el = document.getElementById('validation-warning');
  if(el) el.classList.remove('show');
}


// ===== PREVIEW & BREAKDOWN =====
function toStr(s){ return s.trim(); }

function buildName() {
  const missing = validateRequired();
  if (missing.length > 0) {
    showValidationWarning('Please fill in all required fields: ' + missing.join(', '));
    return [];
  }
  hideValidationWarning();

  const parts = [];
  for(let i=0; i<FIELDS_ORDER.length; i++) {
    const f = FIELDS_ORDER[i];
    if(f === 'custom'){
      const v = document.getElementById('custom-input');
      const cv = v ? v.value.trim() : '';
      if(cv) parts.push({field:f, val:toStr(cv), isCustom:true});
    } else {
      if(sel[f]) parts.push({field:f, val:toStr(sel[f])});
    }
  }
  return parts;
}

function updatePreview() {
  const parts = buildName();
  const outEl = document.getElementById('preview-output');
  const lenEl = document.getElementById('preview-len');
  if(!outEl) return;

  if (parts.length === 0 && validateRequired().length > 0) {
    outEl.innerHTML = '<span class="seg-empty">Fill in all required fields above...</span>';
    if(lenEl) lenEl.textContent = ''; 
    renderBreakdown([]); 
    return;
  }

  if(parts.length === 0){
    outEl.innerHTML = '<span class="seg-empty">Select fields to generate campaign name...</span>';
    if(lenEl) lenEl.textContent = ''; 
    renderBreakdown([]); 
    return;
  }
  
  let pLinks = '';
  for(let i=0; i<parts.length; i++) {
    const p = parts[i];
    if(i > 0) pLinks += '<span class="seg-sep">_</span>';
    pLinks += '<span class="' + (p.isCustom ? 'seg-custom' : 'seg-part') + '">' + escHtml(p.val) + '</span>';
  }
  outEl.innerHTML = pLinks;

  const fullName = parts.map(function(p) { return p.val; }).join('_');
  if(lenEl) lenEl.textContent = fullName.length + ' chars'; 
  window._currentName = fullName;
  renderBreakdown(parts);
}

function renderBreakdown(parts){
  const el = document.getElementById('breakdown-list');
  if(!el) return;
  if(parts.length === 0){ el.innerHTML = '<div class="empty-hint">Fields will appear here after selection...</div>'; return; }
  
  let bHtml = '';
  for(let i=0; i<parts.length; i++) {
    const p = parts[i];
    bHtml += '<div class="breakdown-row"><div class="breakdown-num">' + (i+1) + '</div>';
    bHtml += '<div class="breakdown-field">' + escHtml(FIELD_LABELS[p.field] || p.field) + '</div>';
    bHtml += '<div class="breakdown-val ' + (p.isCustom ? 'custom-val' : '') + '">' + escHtml(p.val) + '</div></div>';
  }
  el.innerHTML = bHtml;
}


// ===== SNAPSHOT TIMEMACHINE =====
function copyMain() {
  const n = window._currentName; if(!n) return; doClipboard(n);
  const b = document.getElementById('copy-main-btn'); if(!b) return;
  b.textContent = '✅ Copied!'; b.classList.add('copied');
  
  const stateSnapshot = {
    sel: JSON.parse(JSON.stringify(sel)),
    custom: document.getElementById('custom-input') ? document.getElementById('custom-input').value : ''
  };

  setTimeout(function(){ b.innerHTML = '[COPY] Copy Name'; b.classList.remove('copied') }, 2000);
  addHistory(n, stateSnapshot);
}

function doClipboard(s){ navigator.clipboard.writeText(s).catch(function(){const t=document.createElement('textarea');t.value=s;document.body.appendChild(t);t.select();document.execCommand('copy');document.body.removeChild(t)})}

function addHistory(n, snapshot){ 
  copyHistory.unshift({
    name: n, 
    time: new Date().toLocaleTimeString('en-US',{hour:'2-digit',minute:'2-digit'}),
    snapshot: snapshot
  }); 
  renderHistory(); 
}

function renderHistory(){
  const el = document.getElementById('history-scroller');
  if(!el) return;
  if(copyHistory.length === 0){ el.innerHTML = '<div class="empty-hint">No records yet...</div>'; return; }
  
  let hHtml = '';
  for(let idx=0; idx<copyHistory.length; idx++) {
    const h = copyHistory[idx];
    hHtml += '<div class="history-row" onclick="restoreSnapshot(' + idx + ')" title="Click to instantly restore this form state">';
    hHtml += '<div class="history-name">' + escHtml(h.name) + '</div>';
    hHtml += '<div class="history-time">' + h.time + '</div>';
    hHtml += '<div class="history-copy-btn" onclick="event.stopPropagation();doClipboard(\'' + h.name.replace(/'/g,"\\'") + '\')">COPY</div></div>';
  }
  el.innerHTML = hHtml;
}

function restoreSnapshot(idx) {
  const record = copyHistory[idx];
  if(!record || !record.snapshot) return;

  Object.keys(sel).forEach(function(k) { delete sel[k]; });
  Object.assign(sel, record.snapshot.sel);

  FIELDS_ORDER.forEach(function(f) {
    const ft = FIELD_TYPES[f];
    if(ft === 'dropdown') {
      const inp = document.getElementById('search-' + f);
      if(inp) {
        if(sel[f]) { inp.value = sel[f]; inp.classList.add('has-value'); }
        else { inp.value = ''; inp.classList.remove('has-value'); }
      }
    } else if(ft === 'tag') {
      const container = document.getElementById('tags-' + f);
      if (container) {
        container.querySelectorAll('.tag').forEach(function(t) {
          t.classList.toggle('active', t.dataset.val === sel[f]);
        });
      }
    }
  });

  applyFunnelFilter();

  const ci = document.getElementById('custom-input');
  if(ci) {
    ci.value = record.snapshot.custom || '';
    if(ci.value !== '') ci.classList.add('ok'); else ci.classList.remove('ok');
  }

  updatePreview();
}

function clearHistory(){copyHistory.length=0;renderHistory()}

function exportTXT(){
  if(!copyHistory.length){alert('No history records yet.');return}
  dlBlob('Campaign Name Export\nExported: '+new Date().toLocaleString('en-US')+'\n'+String.fromCharCode(8210).repeat(60)+'\n\n'+copyHistory.map(function(h,i){return (i+1)+'. '+h.name;}).join('\n'),'campaign-names-'+fmt()+'.txt','text/plain;charset=utf-8');
}
function exportCSV(){
  if(!copyHistory.length){alert('No history records yet.');return}
  dlBlob('\ufeff'+[['No.','Name','Time'],...copyHistory.map(function(h,i){return [i+1,h.name,h.time];})].map(function(r){return r.map(function(c){return '"'+c+'"';}).join(',');}).join('\n'),'campaign-names-'+fmt()+'.csv','text/csv;charset=utf-8');
}

function dlBlob(content,filename,type){const u=URL.createObjectURL(new Blob([content],{type}));const a=document.createElement('a');a.href=u;a.download=filename;document.body.appendChild(a);a.click();document.body.removeChild(a);URL.revokeObjectURL(u)}
function fmt(){const d=new Date();return d.getFullYear()+String(d.getMonth()+1).padStart(2,'0')+String(d.getDate()).padStart(2,'0')}

function pickCustom(val){
  const i=document.getElementById('custom-input');
  if(i){ i.value=val; i.classList.add('ok'); handleCustomInput(i); }
}

function resetAll() {
  Object.keys(OPTIONS).forEach(function(f){delete sel[f];const i=document.getElementById('search-'+f);if(i){i.value='';i.classList.remove('has-value')}});
  document.querySelectorAll('.tag-wrap[id^="tags-"]').forEach(function(c){const f=c.id.replace('tags-','');delete sel[f];c.querySelectorAll('.tag').forEach(function(t){t.classList.remove('active')})});
  const ci=document.getElementById('custom-input');if(ci){ci.value='';ci.classList.remove('error','ok')}
  hideValidationWarning();
  applyFunnelFilter();
  updatePreview();
}

document.addEventListener('click',function(e){if(!e.target.closest('.select-wrap'))document.querySelectorAll('.dropdown').forEach(function(d){d.classList.remove('open')})});

// ===== INIT =====
loadData();
'''


def generate_html(fallback_data):
    """Generate index.html that fetches ./data.json at runtime"""
    fallback_json = json.dumps(fallback_data, ensure_ascii=False)

    ft_dict = {}
    for col_name, ft in FIELD_TYPES.items():
        ft_dict[js_key(col_name)] = ft
    ft_json = json.dumps(ft_dict, ensure_ascii=False)

    # 将 Python 字典序列化为原生的严格 JSON 字符串
    relation_json = json.dumps(PRODUCT_RELATION_MAP, ensure_ascii=False)

    # 采用完全精确的字符字面量安全替换，防转义失效
    js = JS_TEMPLATE.replace('__FALLBACK_JSON__', fallback_json)
    js = js.replace('__FIELD_TYPES_JSON__', ft_json)
    js = js.replace('__PRODUCT_RELATION_MAP__', relation_json)

    clean_js = re.sub(r'[\ud800-\udfff]', '[EMOJI]', js)

    html = (
        "<!DOCTYPE html>\n<html lang=\"en\">\n<head>\n"
        "  <meta charset=\"UTF-8\">\n"
        "  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">\n"
        f"  <title>{PAGE_TITLE}</title>\n"
        f"  <style>{CSS}\n  </style>\n"
        "</head>\n<body>\n\n"
        "<div class=\"header\">\n"
        "  <div class=\"header-icon\">📋</div>\n"
        "  <div>\n"
        f"    <div class=\"header-title\">{PAGE_TITLE}</div>\n"
        f"    <div class=\"header-sub\">{PAGE_SUBTITLE}</div>\n"
        "  </div>\n"
        "</div>\n\n"

        "<div class=\"layout\">\n\n"
        "  <div>\n"
        "    <div class=\"card\">\n"
        "      <div class=\"card-header\">\n"
        "        <div class=\"card-icon\">📝</div>\n"
        "        <div>\n"
        "          <div class=\"card-title\">Select Fields</div>\n"
        "          <div class=\"card-desc\">Choose enum values for each field, or type in Custom.</div>\n"
        "        </div>\n"
        "      </div>\n\n"

        '      <div id="datasource-status" class="datasource-bar loading">\n'
        '        <span class="ds-dot"></span>\n'
        '        <span class="ds-msg">Loading data...</span>\n'
        '        <span class="ds-refresh" onclick="loadData()">[REFRESH]</span>\n'
        "      </div>\n\n"

        f"      <div id=\"fields-container\">\n"
        '        <div style="padding:40px 0;text-align:center">\n'
        '          <div class="skeleton" style="height:40px;width:100%;margin-bottom:16px"></div>\n'
        '          <div class="skeleton" style="height:40px;width:48%;display:inline-block;margin-right:2%"></div>\n'
        '          <div class="skeleton" style="height:40px;width:48%;display:inline-block"></div>\n'
        '          <div class="skeleton" style="height:40px;width:100%;margin-top:16px"></div>\n'
        "        </div>\n"
        "      </div>\n\n"

        "      <div id=\"validation-warning\" class=\"validation-warning\"></div>\n"
        "    </div>\n\n"

        '    <div style="display:flex;justify-content:flex-end;margin-top:12px">\n'
        '      <button class="btn-secondary" onclick="resetAll()" style="padding:8px 18px;border-radius:8px">\n'
        '        🔁 Reset All Fields\n'
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
        "      </div>\n"

        '      <div class="result-dark">\n'
        '        <div class="result-dark-label">生成结果</div>\n'
        '        <div class="result-name-display" id="preview-output">\n'
        '          <span class="seg-empty">加载中…</span>\n'
        "        </div>\n"
        '        <div class="result-len" id="preview-len"></div>\n'
        "      </div>\n"

        '      <button class="btn-primary" id="copy-main-btn" onclick="copyMain()">\n'
        '        [COPY] Copy Name\n'
        "      </button>\n"
        '      <div class="btn-row">\n'
        '        <button class="btn-secondary" onclick="exportTXT()">[TXT] Export TXT</button>\n'
        '        <button class="btn-secondary" onclick="exportCSV()">[CSV] Export CSV</button>\n'
        "      </div>\n"
        '      <button class="btn-download" onclick="downloadTool()">[DL] 下载此工具（离线 / 发给同事）</button>\n'
        "    </div>\n"

        "    <div class=\"card\" style=\"margin-top:14px\">\n"
        '      <div class="card-header" style="margin-bottom:14px;padding-bottom:12px">\n'
        '        <div class="card-icon">&#128269;</div>\n'
        '        <div class="card-title">Field Breakdown</div>\n'
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
        '        <span style="font-size:12px;color:var(--muted);cursor:pointer" onclick="clearHistory()">Clear</span>\n'
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
    print(f"[READ] Excel: {EXCEL_PATH}")
    if not EXCEL_PATH.exists():
        print(f"[ERR] File not found: {EXCEL_PATH}")
        exit(1)

    data = load_data_from_excel()

    OUTPUT_DATA.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')
    print(f"[OK] data.json -> {OUTPUT_DATA}")

    print(f"\n[GEN] Generating index.html...")
    html = generate_html(data)
    OUTPUT_HTML.write_text(html, encoding='utf-8')

    size_html = OUTPUT_HTML.stat().st_size / 1024
    size_data = OUTPUT_DATA.stat().st_size / 1024
    print(f"[OK] index.html -> {OUTPUT_HTML} ({size_html:.0f} KB)")
    print(f"[OK] data.json  -> {OUTPUT_DATA} ({size_data:.0f} KB)")


if __name__ == "__main__":
    main()
