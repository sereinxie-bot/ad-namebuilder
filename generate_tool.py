"""
广告命名工具生成器（GitHub Pages 版）v10.0 - 矩阵解耦版
用法：python generate_tool.py
读取 Excel → 生成 index.html + data.json
核心改进：
1. HTML 内不再死写任何产品映射矩阵，完全交给外部 product-code-map.json 控制。
2. 前端 JS 升级为多源异步流（Promise.all），同时加载两份 JSON，数据与逻辑彻底解耦。
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
PAGE_SUBTITLE = "规范快速生成广告系列命名，支持复制 / 导出 · 数据来源：data.json & product-code-map.json（GitHub 在线维护）"


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

# ===== JS TEMPLATE (使用 Promise.all 多流异步并载，拒绝任何硬编码) =====
JS_TEMPLATE = r'''
// ===== CONFIG =====
const DATA_URL = './data.json';
const MAP_URL = './product-code-map.json'; // 👈 动态地图源解耦

const FALLBACK_DATA = __FALLBACK_JSON__;
const FIELD_TYPES = __FIELD_TYPES_JSON__;

let FIELDS_ORDER = [];
let FIELD_LABELS = {};
let OPTIONS = {};
let PRODUCT_RELATION_MAP = {}; // 👈 由外部外部异步文件接管
let DATA_READY = false;

const REQUIRED_FIELDS = new Set();

const FUNNEL_GOAL_MAP = {
  'awareness':     ['impression','reach'],
  'consideration': ['click','view','pageview','signup','atc','checkout','engagement','dpv','follow'],
  'conversion':    ['sales']
};

const sel = {};
const copyHistory = [];


// ===== DATA LOADING (双源并行异步加载) =====
let _loadTimeout = null;

async function loadData() {
  showDatasourceStatus('loading', 'Loading layout and relation maps...');

  if (_loadTimeout) clearTimeout(_loadTimeout);
  _loadTimeout = setTimeout(function() {
    if (!DATA_READY) {
      console.warn('Data loading timed out, using embedded fallback');
      loadFallback();
      showDatasourceStatus('error', 'Loading timed out. Click [REFRESH] to retry.');
    }
  }, 4000);

  try {
    // 核心大升级：并行发起两项请求，彻底避免死锁
    const [dataResp, mapResp] = await Promise.all([
      fetch(DATA_URL + '?t=' + Date.now()),
      fetch(MAP_URL + '?t=' + Date.now())
    ]);

    if (!dataResp.ok) throw new Error('Data HTTP ' + dataResp.status);
    if (!mapResp.ok) throw new Error('Map HTTP ' + mapResp.status);

    const dataJson = await dataResp.json();
    PRODUCT_RELATION_MAP = await mapResp.json(); // 👈 动态赋权给矩阵

    FIELDS_ORDER = [];
    FIELD_LABELS = {};
    OPTIONS = {};
    REQUIRED_FIELDS.clear();

    for (const [colName, vals] of Object.entries(dataJson)) {
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
    showDatasourceStatus('ok', 'All data synced (' + Object.keys(OPTIONS).length + ' fields, ' + totalVals + ' codes)');
  } catch(err) {
    console.warn('Data fetch failed, using embedded fallback:', err);
    if (_loadTimeout) { clearTimeout(_loadTimeout); _loadTimeout = null; }
    loadFallback();
    showDatasourceStatus('error', 'Failed to load configuration JSON files. Using built-in local data.');
  }
}

function loadFallback() {
  FIELDS_ORDER = [];
  FIELD_LABELS = {};
  OPTIONS = {};
  REQUIRED_FIELDS.clear();
  PRODUCT_RELATION_MAP = {}; 
  
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

function escHtml(s) { return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/
