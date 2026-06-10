"""
广告命名工具生成器（GitHub Pages 版）v7 - 跨字段联动筛选版
用法：python generate_tool.py
读取 Excel → 生成 index.html + data.json
核心改进：
1. 融入 Product Type -> Product Code 的映射矩阵数据，实现级联联动筛选
2. 切换 Product Type 时，自动清理不匹配的 Product Code，防命名逻辑冲突
3. 保留 v6 版本的智能洗稿以及历史记录时光机功能
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

# ===== 你的 Product Type -> Product Code 级联映射数据 =====
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
    """从本地 Excel 读取数据，生成 data.json 和 HTML 内嵌 fallback"""
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

// 核心级联筛选映射矩阵
const PRODUCT_RELATION_MAP = __PRODUCT_RELATION_MAP__;

let FIELDS_ORDER = [];
let FIELD_LABELS = {};
let OPTIONS = {};
let DATA_READY = false;

const REQUIRED_FIELDS = new Set();

const FUNNEL_GOAL_MAP = {
  "awareness":     ["impression","reach"],
  "consideration":["click","view","pageview","signup","atc","checkout","engagement","dpv","follow"],
  "conversion":   ["sales"],
};

const sel = {};
const copyHistory = [];


// ===== DATA LOADING =====
let _loadTimeout = null;

async function loadData() {
  showDatasourceStatus('loading', 'Loading data...');

  if (_loadTimeout) clearTimeout(_loadTimeout);
  _loadTimeout = setTimeout(() => {
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
  el.innerHTML = '<span class="ds-dot"></span><span class="ds-msg">' + msg + '</span>'
    + '<span class="ds-refresh" onclick="loadData()">[REFRESH]</span>';
}


// ===== RENDER FIELDS =====
function renderAllFields() {
  const container = document.getElementById('fields-container');
  let html = '';

  for (const f of FIELDS_ORDER) {
    if (f === 'market') {
      html += `<div class="field-compact">`;
    }

    const ft = FIELD_TYPES[f] || 'tag';
    const label = FIELD_LABELS[f];
    const opts = OPTIONS[f] || [];

    if (ft === 'dropdown') {
      const isFullWidth = (f === 'product-code' || opts.length > 30);
      const phMap = {'market':'Search or select market...','product-type':'Search product type...','product-code':'Search product model (e.g. argus 4, rlc-410)...'};
      const ph = phMap[f] || 'Search ' + label + '...';
      html += `<div class="field${isFullWidth ? ' field-full' : ''}">
        <div class="field-label"><span class="req">*</span> ${escHtml(label)}</div>
        <div class="select-wrap" id="wrap-${f}">
          <input class="select-search" id="search-${f}" placeholder="${escHtml(ph)}" autocomplete="off"
            onfocus="openDropdown('${f}')" oninput="filterDropdown('${f}',this.value)" onblur="blurDropdown('${f}')">
          <span class="select-arrow">&#9662;</span>
          <div class="dropdown" id="drop-${f}"></div>
        </div></div>`;
    } else if (ft === 'tag') {
      const tags = renderTagOptions(f, opts);
      html += `<div class="field"><div class="field-label"><span class="req">*</span> ${escHtml(label)}</div>
        <div class="tag-wrap" id="tags-${f}">${tags}</div></div>`;
    } else if (ft === 'custom') {
      if (FIELDS_ORDER.includes('market')) {
        html += `</div>`; 
      }
      html += makeCustomHTML();
    }

    if (f === 'product-type') {
      html += `</div>`; 
    }
  }

  container.innerHTML = `<div class="fields-grid">${html}</div>`;
  applyFunnelFilter();
}

function renderTagOptions(fieldKey, opts) {
  let disabledGoals = null;
  if (fieldKey === 'bidding-goal' && sel['media-funnel']) {
    disabledGoals = FUNNEL_GOAL_MAP[sel['media-funnel']] || null;
  }

  return opts.map(v => {
    const isDisabled = disabledGoals && !disabledGoals.includes(v);
    return `<div class="tag${isDisabled ? ' disabled' : ''}${sel[fieldKey]===v ? ' active' : ''}" data-val="${escAttr(v)}" onclick="${isDisabled ? '' : `selectTag('${fieldKey}','${escJs(v)}')`}">${escHtml(v)}</div>`;
  }).join('');
}

function applyFunnelFilter() {
  const funnelVal = sel['media-funnel'];
  const goalContainer = document.getElementById('tags-bidding-goal');
  if (!goalContainer) return;

  const
