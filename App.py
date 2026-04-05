import streamlit as st
import pandas as pd
import io
import os
import time
import plotly.express as px
from ratio_engine1 import RatioEngine, REQUIRED_FIELDS, ValidationResult, EngineConfig

# PDF generation for Formula Guide
import io as _io_module
import re as _re_module
try:
    from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                     Table, TableStyle, HRFlowable)
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.lib import colors as _rl_colors
    _REPORTLAB_OK = True
except ImportError:
    _REPORTLAB_OK = False

try:
    from Dashboard import DashboardGenerator
    DASHBOARD_AVAILABLE = True
except ImportError:
    DASHBOARD_AVAILABLE = False



st.set_page_config(
    page_title="Financial Ratio Analysis",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)


def init_session_state():
    defaults = {
        'raw_df': None, 'mapping': {}, 'engine': None,
        'industry': "Manufacturing", 'tax_rate': 0.29,
        'validation_result': None, 'data_quality': None,
        'file_valid': False, 'file_name': None,
        'analysis_complete': False, 'workflow_step': 1,
        'show_sample_success': False, 'frequency': "Annual",
        'operating_cash_pct': 2.0, 'allow_negative_coverage': True,
        'nopat_tax_benefit': True,
        'ratio_category': "💧 Liquidity",
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value
init_session_state()

st.markdown("""<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');

/* ── Design tokens — Light mode ──────────────── */
:root {
    --navy:        #0F172A;
    --navy-mid:    #1E293B;
    --navy-light:  #334155;
    --indigo:      #6366F1;
    --indigo-dark: #4F46E5;
    --teal:        #14B8A6;
    --teal-dark:   #0D9488;
    --amber:       #F59E0B;
    --red:         #EF4444;
    --green:       #10B981;

    /* Adaptive tokens */
    --bg-page:       #F8FAFC;
    --bg-card:       #FFFFFF;
    --bg-subtle:     #F1F5F9;
    --bg-input:      #FFFFFF;
    --border:        #E2E8F0;
    --border-mid:    #CBD5E1;
    --text-primary:  #0F172A;
    --text-secondary:#475569;
    --text-muted:    #94A3B8;
    --text-inverse:  #FFFFFF;
    --shadow-sm:  0 1px 3px rgba(0,0,0,.06), 0 1px 2px rgba(0,0,0,.04);
    --shadow-md:  0 4px 12px rgba(0,0,0,.08), 0 2px 4px rgba(0,0,0,.04);
    --shadow-lg:  0 12px 28px rgba(0,0,0,.10), 0 4px 8px rgba(0,0,0,.04);
    --radius-sm: 6px; --radius-md: 10px; --radius-lg: 16px; --radius-xl: 22px;
    /* Alert backgrounds */
    --alert-success-bg: #f0fdf4; --alert-success-border: #10B981; --alert-success-text: #065f46;
    --alert-info-bg:    #eef2ff; --alert-info-border:    #6366F1; --alert-info-text:    #3730a3;
    --alert-warning-bg: #fffbeb; --alert-warning-border: #F59E0B; --alert-warning-text: #92400e;
    --alert-error-bg:   #fef2f2; --alert-error-border:   #EF4444; --alert-error-text:   #991b1b;
    /* Banner, ratio description, export cards */
    --ratio-banner-bg:     #eef2ff;
    --ratio-banner-border: #c7d2fe;
    --ratio-banner-text:   #3730a3;
    --export-card-bg:      #f8fafc;
    --export-card-border:  #e2e8f0;
    --export-card-title:   #334155;
    --export-card-body:    #64748b;
    --empty-state-text:    #64748B;
    --empty-state-head:    #334155;
    --breadcrumb-cat:      #94A3B8;
    --breadcrumb-arrow:    #CBD5E1;
    --breadcrumb-title:    #1E293B;
    --prog-label:          #334155;
}

/* ── Dark mode token overrides ───────────────── */
[data-theme="dark"], .stApp[data-theme="dark"] {
    --bg-page:       #0F172A;
    --bg-card:       #1E293B;
    --bg-subtle:     #162032;
    --bg-input:      #1E293B;
    --border:        #334155;
    --border-mid:    #475569;
    --text-primary:  #F1F5F9;
    --text-secondary:#CBD5E1;
    --text-muted:    #64748B;
    /* Alert backgrounds — slightly brighter borders so they're visible on dark bg */
    --alert-success-bg: rgba(16,185,129,.12); --alert-success-border: #34D399; --alert-success-text: #6EE7B7;
    --alert-info-bg:    rgba(99,102,241,.15); --alert-info-border:    #818CF8; --alert-info-text:    #C7D2FE;
    --alert-warning-bg: rgba(245,158,11,.12); --alert-warning-border: #FCD34D; --alert-warning-text: #FDE68A;
    --alert-error-bg:   rgba(239,68,68,.12);  --alert-error-border:   #F87171; --alert-error-text:   #FECACA;
    --ratio-banner-bg:     rgba(99,102,241,.15);
    --ratio-banner-border: rgba(99,102,241,.4);
    --ratio-banner-text:   #C7D2FE;
    --export-card-bg:      #1E293B;
    --export-card-border:  #334155;
    --export-card-title:   #E2E8F0;
    --export-card-body:    #94A3B8;
    --empty-state-text:    #64748B;
    --empty-state-head:    #CBD5E1;
    --breadcrumb-cat:      #64748B;
    --breadcrumb-arrow:    #475569;
    --breadcrumb-title:    #F1F5F9;
    --prog-label:          #E2E8F0;
}

/* Also target the OS-level dark preference for browsers that don't set data-theme */
@media (prefers-color-scheme: dark) {
    :root {
        --bg-page: #0F172A; --bg-card: #1E293B; --bg-subtle: #162032;
        --bg-input: #1E293B; --border: #334155; --border-mid: #475569;
        --text-primary: #F1F5F9; --text-secondary: #CBD5E1; --text-muted: #64748B;
        --alert-success-bg: rgba(16,185,129,.12); --alert-success-border: #34D399; --alert-success-text: #6EE7B7;
        --alert-info-bg: rgba(99,102,241,.15);    --alert-info-border: #818CF8;    --alert-info-text: #C7D2FE;
        --alert-warning-bg: rgba(245,158,11,.12); --alert-warning-border: #FCD34D; --alert-warning-text: #FDE68A;
        --alert-error-bg: rgba(239,68,68,.12);    --alert-error-border: #F87171;   --alert-error-text: #FECACA;
        --ratio-banner-bg: rgba(99,102,241,.15); --ratio-banner-border: rgba(99,102,241,.4); --ratio-banner-text: #C7D2FE;
        --export-card-bg: #1E293B; --export-card-border: #334155;
        --export-card-title: #E2E8F0; --export-card-body: #94A3B8;
        --empty-state-text: #64748B; --empty-state-head: #CBD5E1;
        --breadcrumb-cat: #64748B; --breadcrumb-arrow: #475569; --breadcrumb-title: #F1F5F9;
        --prog-label: #E2E8F0;
    }
}

/* ── Base ─────────────────────────────────────── */
html, body, [class*="css"] { font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif !important; }
.main .block-container { padding-top: 0 !important; padding-bottom: 3rem; max-width: 1480px; }
#MainMenu, footer { visibility: hidden; }

/* ── Header banner ────────────────────────────── */
.app-header {
    background: linear-gradient(135deg, #0F172A 0%, #1E293B 50%, #0F2845 100%);
    padding: 2rem 2.5rem;
    border-radius: 0 0 var(--radius-xl) var(--radius-xl);
    margin: -1rem -1rem 1.5rem -1rem;
    position: relative; overflow: hidden;
    box-shadow: var(--shadow-lg);
}
.app-header::before {
    content: ''; position: absolute; inset: 0;
    background-image: linear-gradient(rgba(99,102,241,.12) 1px,transparent 1px),
                      linear-gradient(90deg,rgba(99,102,241,.12) 1px,transparent 1px);
    background-size: 32px 32px;
}
.app-header::after {
    content: ''; position: absolute; top:-60px; right:-60px;
    width:260px; height:260px;
    background: radial-gradient(circle,rgba(99,102,241,.25) 0%,transparent 70%);
    pointer-events: none;
}
.header-inner { position:relative; z-index:1; display:flex; align-items:center; gap:1.5rem; }
.header-icon {
    width:56px; height:56px;
    background: linear-gradient(135deg,var(--indigo) 0%,var(--teal) 100%);
    border-radius: var(--radius-md); display:flex; align-items:center; justify-content:center;
    font-size:1.75rem; box-shadow:0 4px 14px rgba(99,102,241,.4); flex-shrink:0;
}
.header-text h1 { font-size:1.7rem; font-weight:800; color:white; margin:0; letter-spacing:-.02em; line-height:1.2; }
.header-text p  { color:rgba(255,255,255,.65); font-size:.88rem; margin:.2rem 0 0; }
.header-badges  { margin-left:auto; display:flex; gap:.5rem; align-items:center; }
.badge { padding:.25rem .75rem; border-radius:20px; font-size:.72rem; font-weight:600; letter-spacing:.03em; }
.badge-live    { background:rgba(16,185,129,.2); color:#34D399; border:1px solid rgba(16,185,129,.3); }
.badge-live::before { content:'● '; }
.badge-version { background:rgba(99,102,241,.2); color:#A5B4FC; border:1px solid rgba(99,102,241,.3); }

/* ── Workflow stepper ─────────────────────────── */
.stepper {
    display:flex; align-items:center; padding:1rem 1.5rem;
    background: var(--bg-card); border-radius: var(--radius-lg);
    border:1px solid var(--border); box-shadow: var(--shadow-sm);
    margin-bottom:1.25rem;
}
.step-node { display:flex; align-items:center; flex:1; }
.step-circle {
    width:36px; height:36px; border-radius:50%;
    display:flex; align-items:center; justify-content:center;
    font-size:.85rem; font-weight:700; flex-shrink:0; transition:all .25s;
}
.step-circle.done   { background:linear-gradient(135deg,#10B981,#059669); color:white; box-shadow:0 3px 10px rgba(16,185,129,.35); }
.step-circle.active { background:linear-gradient(135deg,var(--indigo),var(--indigo-dark)); color:white; box-shadow:0 3px 10px rgba(99,102,241,.4); }
.step-circle.pending{ background:var(--bg-subtle); color:var(--text-muted); border:2px solid var(--border); }
.step-label { margin-left:.6rem; }
.step-label .step-name { font-weight:600; font-size:.85rem; color:var(--text-secondary); }
.step-label .step-desc { font-size:.73rem; color:var(--text-muted); }
.step-label .step-name.active-text { color:var(--indigo); }
.step-label .step-name.done-text   { color:#059669; }
.step-connector { flex:1; height:2px; max-width:80px; background:var(--border); margin:0 .5rem; }
.step-connector.done { background:linear-gradient(90deg,#10B981,#059669); }

/* ── Alerts ───────────────────────────────────── */
.alert-success, .alert-info, .alert-warning, .alert-error {
    padding:.85rem 1.1rem; border-radius: var(--radius-md); border-left:4px solid;
    margin:.75rem 0; font-size:.9rem; animation:slideIn .25s ease-out;
}
@keyframes slideIn { from{opacity:0;transform:translateY(-8px)} to{opacity:1;transform:translateY(0)} }
.alert-success { background:var(--alert-success-bg); border-color:var(--alert-success-border); color:var(--alert-success-text); }
.alert-info    { background:var(--alert-info-bg);    border-color:var(--alert-info-border);    color:var(--alert-info-text); }
.alert-warning { background:var(--alert-warning-bg); border-color:var(--alert-warning-border); color:var(--alert-warning-text); }
.alert-error   { background:var(--alert-error-bg);   border-color:var(--alert-error-border);   color:var(--alert-error-text); }

/* ── Section title ────────────────────────────── */
.section-title {
    font-size:1.15rem; font-weight:700; color:var(--text-primary);
    padding-bottom:.6rem; border-bottom:2px solid var(--indigo);
    margin:1.25rem 0 1rem; display:flex; align-items:center; gap:.5rem;
}

/* ── Mapping category badge ───────────────────── */
.cat-badge {
    display:inline-flex; align-items:center; gap:.35rem;
    padding:.22rem .65rem; border-radius:20px;
    font-size:.75rem; font-weight:600;
    background:var(--bg-subtle); color:var(--text-secondary); border:1px solid var(--border);
    margin-bottom:.6rem;
}
.cat-badge.complete { background:rgba(16,185,129,.12); color:#059669; border-color:rgba(16,185,129,.3); }
.cat-badge.partial  { background:rgba(245,158,11,.12);  color:#D97706; border-color:rgba(245,158,11,.3); }

/* ── Progress bar ─────────────────────────────── */
.prog-bar-wrap { background:var(--border); border-radius:4px; height:6px; overflow:hidden; }
.prog-bar-fill { height:100%; border-radius:4px; transition:width .4s ease;
    background:linear-gradient(90deg,var(--indigo),var(--teal)); }

/* ── Ratio info banner ────────────────────────── */
.ratio-banner {
    background:var(--ratio-banner-bg); border:1px solid var(--ratio-banner-border);
    border-radius:10px; padding:.75rem 1rem; margin-bottom:1rem;
    font-size:.87rem; color:var(--ratio-banner-text);
}
/* ── Breadcrumb ───────────────────────────────── */
.breadcrumb {
    display:flex; align-items:center; gap:.5rem; margin-bottom:.5rem;
}
.breadcrumb-cat   { font-size:.75rem; color:var(--breadcrumb-cat); font-weight:600; text-transform:uppercase; letter-spacing:.05em; }
.breadcrumb-arrow { color:var(--breadcrumb-arrow); }
.breadcrumb-title { font-size:1.1rem; font-weight:700; color:var(--breadcrumb-title); }

/* ── Export cards ─────────────────────────────── */
.export-card {
    background:var(--export-card-bg); border:1px solid var(--export-card-border);
    border-radius:12px; padding:1.1rem;
}
.export-card-title { font-weight:700; color:var(--export-card-title); margin-bottom:.3rem; }
.export-card-body  { font-size:.82rem; color:var(--export-card-body); margin-bottom:.8rem; }

/* ── Empty states ─────────────────────────────── */
.empty-state { text-align:center; padding:3rem; color:var(--empty-state-text); }
.empty-state-icon { font-size:3rem; margin-bottom:1rem; }
.empty-state h3 { color:var(--empty-state-head); }

/* ── Streamlit native overrides ───────────────── */
.stTabs [data-baseweb="tab-list"] {
    gap:.35rem; background:var(--bg-subtle); padding:.4rem;
    border-radius:var(--radius-lg); border:1px solid var(--border);
}
.stTabs [data-baseweb="tab"] {
    border-radius:var(--radius-md); padding:.6rem 1.25rem;
    font-weight:600; font-size:.88rem; color:var(--text-secondary) !important; transition:all .2s;
}
.stTabs [data-baseweb="tab"]:hover { background:var(--bg-card); color:var(--indigo) !important; }
.stTabs [aria-selected="true"] {
    background:linear-gradient(135deg,var(--indigo) 0%,var(--indigo-dark) 100%) !important;
    color:white !important; box-shadow:var(--shadow-md);
}
[data-testid="stMetric"] {
    background:var(--bg-card); border:1px solid var(--border);
    border-radius:var(--radius-md); padding:.9rem 1rem;
    box-shadow:var(--shadow-sm); transition:all .2s;
}
[data-testid="stMetric"]:hover { box-shadow:var(--shadow-md); transform:translateY(-2px); }
[data-testid="stMetricLabel"] { font-size:.78rem !important; font-weight:600; color:var(--text-muted) !important; text-transform:uppercase; letter-spacing:.04em; }
[data-testid="stMetricValue"] { font-size:1.45rem !important; font-weight:800; color:var(--text-primary) !important; }
.stButton > button {
    border-radius:var(--radius-md); font-weight:600; font-size:.88rem;
    padding:.45rem 1.1rem; transition:all .2s; border:none;
}
.stButton > button:hover { transform:translateY(-1px); box-shadow:var(--shadow-md); }
.stButton > button[kind="primary"] {
    background:linear-gradient(135deg,var(--indigo) 0%,var(--indigo-dark) 100%); color:white;
}
.stButton > button[kind="primary"]:hover {
    background:linear-gradient(135deg,#818cf8 0%,var(--indigo) 100%);
}
.stDownloadButton > button {
    background:linear-gradient(135deg,var(--teal) 0%,var(--teal-dark) 100%);
    color:white; border:none; border-radius:var(--radius-md); font-weight:600; transition:all .2s;
}
.stDownloadButton > button:hover {
    background:linear-gradient(135deg,#2dd4bf 0%,var(--teal) 100%);
    transform:translateY(-1px); box-shadow:var(--shadow-md);
}
.stDataFrame { border-radius:var(--radius-md); overflow:hidden; box-shadow:var(--shadow-md); }
[data-testid="stFileUploader"] {
    background:var(--bg-subtle); border:2px dashed var(--border-mid);
    border-radius:var(--radius-lg); padding:.75rem; transition:all .2s;
}
[data-testid="stFileUploader"]:hover { border-color:var(--indigo); background:rgba(99,102,241,.05); }
.streamlit-expanderHeader {
    font-weight:600 !important; font-size:.9rem !important;
    background:var(--bg-subtle) !important; border-radius:var(--radius-md) !important;
    color:var(--text-secondary) !important;
}
.streamlit-expanderHeader:hover { background:var(--bg-card) !important; color:var(--indigo) !important; }

/* ── Sidebar — always dark ────────────────────── */
section[data-testid="stSidebar"] { background:var(--navy) !important; border-right:1px solid var(--navy-light) !important; }
section[data-testid="stSidebar"] .stMarkdown,
section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] p { color:rgba(255,255,255,.8) !important; }
section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3 { color:white !important; }
section[data-testid="stSidebar"] [data-testid="stMetric"]      { background:var(--navy-mid) !important; border-color:var(--navy-light) !important; }
section[data-testid="stSidebar"] [data-testid="stMetricLabel"] { color:rgba(255,255,255,.5) !important; }
section[data-testid="stSidebar"] [data-testid="stMetricValue"] { color:white !important; }
section[data-testid="stSidebar"] .stSelectbox > div > div,
section[data-testid="stSidebar"] .stMultiSelect > div > div   { background:var(--navy-mid) !important; border-color:var(--navy-light) !important; color:white !important; }
section[data-testid="stSidebar"] .stCheckbox label,
section[data-testid="stSidebar"] .stRadio label { color:rgba(255,255,255,.8) !important; }
.sidebar-divider { border-top:1px solid rgba(255,255,255,.1); margin:.75rem 0; }
.sidebar-section-label { font-size:.68rem; font-weight:700; text-transform:uppercase; letter-spacing:.1em; color:rgba(255,255,255,.35) !important; padding:.4rem 0 .1rem; }

/* ── App footer ───────────────────────────────── */
.app-footer { text-align:center; color:var(--text-muted); padding:1.5rem 0; margin-top:2rem; border-top:1px solid var(--border); font-size:.8rem; }
.app-footer strong { color:var(--indigo); }

@media (max-width:768px) {
    .app-header { padding:1.25rem; border-radius:0 0 var(--radius-lg) var(--radius-lg); }
    .header-text h1 { font-size:1.25rem; }
    .header-badges { display:none; }
}
@keyframes fadeIn { from{opacity:0} to{opacity:1} }
</style>""", unsafe_allow_html=True)

# ── Constants ──────────────────────────────────────────────────────────────────
RATIO_EXPLANATIONS = {
    "Current Ratio": "Measures ability to meet short-term obligations with current assets. Standard threshold: 1.5–3.0",
    "Quick Ratio": "Measures ability to meet short-term obligations using liquid assets only (excludes inventory). Standard threshold: ≥1.0",
    "Cash Ratio": "Most conservative liquidity measure: cash and equivalents relative to current liabilities",
    "Defensive Interval (Days)": "Number of days a company can operate using only its liquid assets, without additional revenue",
    "NWC to Assets": "Proportion of total assets funded by net working capital. Indicates liquidity cushion",
    "Inventory Turnover": "Number of times inventory is sold and replenished during a period. Higher values indicate efficiency",
    "Days Inventory (DIO)": "Average number of days inventory is held before sale. Lower values indicate faster turnover",
    "Receivables Turnover": "Number of times receivables are collected during a period. Higher values indicate efficient collection",
    "Days Sales Outstanding (DSO)": "Average number of days to collect payment after a sale. Standard target: 30–60 days",
    "Payables Turnover": "Number of times payables are settled during a period",
    "Days Payables (DPO)": "Average number of days to pay suppliers after receiving goods or services",
    "Cash Conversion Cycle": "DIO + DSO − DPO. Measures the time (in days) between cash outflow for inventory and cash inflow from sales",
    "Total Asset Turnover": "Revenue generated per unit of total assets. Measures overall asset utilisation",
    "Fixed Asset Turnover": "Revenue generated per unit of property, plant and equipment. Measures fixed asset productivity",
    "Working Capital Turnover": "Revenue generated per unit of net working capital. ⚠️ Shows N/A when working capital is zero or negative — a negative ratio is misleading because the sign of the denominator inverts the interpretation.",
    "Debt to Equity": "Total debt relative to total equity. Measures financial leverage. Conservative threshold: <1.0",
    "Debt to Assets": "Proportion of total assets financed by debt",
    "Debt to Capital": "Debt as a proportion of total capital (debt plus equity)",
    "Interest Coverage": "Operating income (EBIT) divided by gross interest expense. Measures ability to service debt. Threshold: ≥3.0×",
    "Cash Interest Coverage": "EBITDA divided by gross interest expense. Cash-based debt service capacity",
    "Net Debt to EBITDA": "Net debt (total debt minus cash) divided by EBITDA. Investment-grade threshold: <3.0×",
    "Financial Leverage": "Equity multiplier (total assets ÷ total equity). DuPont decomposition component",
    "Gross Margin": "Gross profit as a percentage of revenue. Reflects pricing power and production efficiency",
    "Operating Margin": "Operating income as a percentage of revenue. Measures core operational efficiency. Benchmark: ≥10%",
    "Net Margin": "Net income as a percentage of revenue. Measures overall profitability after all expenses",
    "EBITDA Margin": "EBITDA as a percentage of revenue. Indicates cash-generation capacity from operations",
    "Pretax Margin": "Pretax income as a percentage of revenue. Indicates core business performance before tax effects",
    "ROA": "Return on assets: net income divided by average total assets. Benchmark: ≥5%",
    "ROE": "Return on equity: net income divided by average total equity. Benchmark: ≥15%",
    "ROIC": "Return on invested capital: NOPAT divided by invested capital. Should exceed the weighted average cost of capital",
    "Cash ROIC": "Free cash flow divided by invested capital. Cash-based return on invested capital",
    "ROE (Normalized)": "NOPAT ÷ Average Total Equity. Removes distortion of interest tax shields. Useful for cross-capital-structure comparisons.",
    "DuPont: Net Margin": "DuPont decomposition component: net income ÷ revenue",
    "DuPont: Asset Turnover": "DuPont decomposition component: revenue ÷ average total assets",
    "DuPont: Equity Multiplier": "DuPont decomposition component: average total assets ÷ average total equity",
    "OCF to Sales": "Operating cash flow as a percentage of revenue. Benchmark: ≥10%",
    "FCF to Sales": "Free cash flow as a percentage of revenue. Indicates free cash margin",
    "Quality of Income": "Operating cash flow ÷ net income. Values >1.0 indicate high earnings quality",
    "Capex Coverage": "Operating cash flow ÷ capital expenditures. Values >1.5 indicate strong internal funding",
    "Dividend Payout": "Dividends paid as a percentage of net income. Sustainable range: 30–50%. ⚠️ Shows N/A when net income is negative.",
    "FCF Conversion": "Free cash flow ÷ EBITDA. Benchmark: ≥0.8×. ⚠️ Shows N/A when EBITDA is negative — ratio direction inverts.",
    "EPS": "Earnings per share: net income attributable to common shareholders ÷ diluted shares outstanding",
    "P/E Ratio": "Price-to-earnings ratio: share price ÷ EPS. Market average range: 15–20×",
    "PEG Ratio": "P/E ratio ÷ earnings growth rate. Values <1.0 may indicate undervaluation relative to growth",
    "Book Value Per Share": "Total equity ÷ shares outstanding. Represents net asset value per share",
    "Price to Book (P/B)": "Share price ÷ book value per share. Compares market valuation to net asset value",
    "Price to Sales": "Market capitalisation ÷ revenue. Revenue-based valuation multiple",
    "EV / EBITDA": "Enterprise value ÷ EBITDA. Standard valuation range: 8–12×",
    "EV / Revenue": "Enterprise value ÷ revenue. Enterprise-level revenue multiple",
    "FCF Yield": "Free cash flow ÷ market capitalisation. Benchmark: ≥5%",
    "Revenue Growth": "Period-over-period revenue growth rate",
    "Gross Profit Growth": "Period-over-period gross profit growth rate",
    "Operating Income Growth": "Period-over-period operating income growth rate. Indicates operating leverage",
    "EBITDA Growth": "Period-over-period EBITDA growth rate",
    "Net Income Growth": "Period-over-period net income growth rate",
    "EPS Growth": "Period-over-period earnings per share growth rate",
    "FCF Growth": "Period-over-period free cash flow growth rate",
    "Altman Z-Score": "Public-company distress predictor (Altman 1968, 5-factor). Safe: >2.99 | Grey: 1.81–2.99 | Distress: <1.81",
    "Altman Z-Score (EM Score)": "Emerging-market / private EM model (4-factor). Safe: >2.6 | Grey: 1.1–2.6 | Distress: <1.1. No market price required.",
    "Piotroski F-Score": "9-point binary signal composite (Piotroski 2000). Profitability (4pts) + Leverage/Liquidity (3pts) + Efficiency (2pts). Score 8–9 = Strong buy signal; 0–2 = Short signal. Each sub-signal scores 1 if favourable, 0 if not.",
    "Beneish M-Score": "Earnings manipulation detector (Beneish 1999). 8-variable logistic model. M > −2.22 = potential manipulation; M < −2.22 = manipulation unlikely. Not a guarantee — use as a flag for deeper investigation.",
    "Ohlson Bankruptcy Prob": "Probability of bankruptcy within one year (Ohlson 1980 O-Score logistic model). 0% = minimal risk; values above 50% indicate serious distress. Does not require market price data.",
    "Sloan Accrual Ratio": "(Net Income − Operating Cash Flow) / Average Total Assets. Positive = accounting profits exceed cash (low quality). Negative = cash exceeds profits (high quality). Values above +5% are a red flag. Source: Sloan (1996).",
    "CapEx to D&A": "Capital Expenditures ÷ Depreciation & Amortisation. >1 = net investment (company growing asset base); ~1 = maintenance mode; <1 = asset harvesting (under-investing).",
    "Degree of Operating Leverage": "(%ΔEBIT) ÷ (%ΔRevenue). Measures fixed-cost intensity. High DOL (>3) means small revenue changes cause large profit swings — high operational risk in downturns but high operational gearing in upturns.",
    "Net Debt to Equity": "(Total Debt − Cash) ÷ Total Equity. Unlike D/E, explicitly nets off cash holdings. Negative = net cash company (more cash than debt). More informative than D/E for capital allocation analysis.",
    "Earnings Yield (EBIT/EV)": "EBIT ÷ Enterprise Value. Greenblatt's Magic Formula numerator. The earnings-based alternative to E/P that strips out capital structure effects. Higher = cheaper relative to earnings power.",
    "ROIC Spread (vs 10% Hurdle)": "ROIC minus a 10% cost-of-capital proxy. Positive = value creation above hurdle; negative = value destruction. The 10% is a rough universal hurdle — adjust interpretation for capital-intensive sectors.",
    "Interest Expense Ratio": "Interest Expense ÷ Revenue. Shows the fraction of every revenue dollar consumed by debt service before tax. Values above 5% in most industries warrant attention.",
    "Reinvestment Rate": "(CapEx − D&A + ΔNWC) ÷ NOPAT. What fraction of after-tax operating profit is reinvested. Pairs with ROIC: high reinvestment + high ROIC = fast compounding. Undefined when NOPAT ≤ 0.",
    "Sustainable Growth Rate": "ROIC × Reinvestment Rate. Maximum growth achievable without external financing, assuming returns and reinvestment rates hold constant. Theoretical ceiling on organic growth.",
}

# Ratio categories for the Step 3 browser
RATIO_CATEGORIES = {
    "💧 Liquidity":     ["Current Ratio","Quick Ratio","Cash Ratio","Defensive Interval (Days)","NWC to Assets"],
    "⚙️ Efficiency":    ["Inventory Turnover","Days Inventory (DIO)","Receivables Turnover","Days Sales Outstanding (DSO)",
                         "Payables Turnover","Days Payables (DPO)","Cash Conversion Cycle",
                         "Total Asset Turnover","Fixed Asset Turnover","Working Capital Turnover",
                         "CapEx to D&A","Degree of Operating Leverage","Reinvestment Rate","Sustainable Growth Rate"],
    "🏛️ Solvency":     ["Debt to Equity","Net Debt to Equity","Debt to Assets","Debt to Capital",
                         "Interest Coverage","Cash Interest Coverage","Net Debt to EBITDA",
                         "Financial Leverage","Interest Expense Ratio"],
    "📈 Profitability": ["Gross Margin","Operating Margin","Net Margin","EBITDA Margin","Pretax Margin"],
    "💰 Returns":       ["ROA","ROE","ROIC","Cash ROIC","ROE (Normalized)",
                         "ROIC Spread (vs 10% Hurdle)","Earnings Yield (EBIT/EV)",
                         "DuPont: Net Margin","DuPont: Asset Turnover","DuPont: Equity Multiplier"],
    "💵 Cash Flow":     ["OCF to Sales","FCF to Sales","Quality of Income","Capex Coverage",
                         "Dividend Payout","FCF Conversion","Sloan Accrual Ratio"],
    "📊 Valuation":     ["EPS","P/E Ratio","PEG Ratio","Book Value Per Share","Price to Book (P/B)",
                         "Price to Sales","EV / EBITDA","EV / Revenue","FCF Yield"],
    "📉 Growth":        ["Revenue Growth","Gross Profit Growth","Operating Income Growth",
                         "EBITDA Growth","Net Income Growth","EPS Growth","FCF Growth"],
    "⚠️ Risk & Quality":["Altman Z-Score","Altman Z-Score (EM Score)",
                         "Piotroski F-Score","Beneish M-Score",
                         "Ohlson Bankruptcy Prob","Sloan Accrual Ratio"],
}

# ── Helpers ────────────────────────────────────────────────────────────────────

def show_alert(message: str, alert_type: str = "info"):
    st.markdown(f'<div class="alert-{alert_type}">{message}</div>', unsafe_allow_html=True)

def _user_friendly_error(e: Exception) -> str:
    msg = str(e)
    if "Validation failed" in msg:
        return "Calculation cannot proceed: one or more required fields are missing or contain no valid data. Review field mappings in Step 2."
    if "DataFrame is empty" in msg or "File is empty" in msg:
        return "The uploaded file contains no data. Verify the file is not empty and follows the required format."
    if "No valid companies" in msg or "No companies found" in msg:
        return "No company identifiers found in the first column. Verify the file structure."
    if "No year columns" in msg or "No period" in msg:
        return "No period columns detected. Ensure columns 3+ contain period data."
    if "No companies successfully calculated" in msg:
        return "Calculations could not be completed for any company. Review data completeness and field mappings."
    if "at least" in msg.lower() and "column" in msg.lower():
        return "The file does not contain enough columns. Minimum: Company, Field Name, one period column."
    return "An unexpected error occurred. Review the input data and try again."

def validate_file_structure(df: pd.DataFrame) -> tuple:
    if df is None or df.empty:
        return False, "File is empty"
    if df.shape[1] < 3:
        return False, f"Need ≥3 columns, found {df.shape[1]}"
    companies = df.iloc[:, 0].dropna()
    if len(companies) == 0:
        return False, "No companies found"
    fields = df.iloc[:, 1].dropna()
    if len(fields) == 0:
        return False, "No field names found"
    numeric_cols = df.iloc[:, 2:].apply(pd.to_numeric, errors='coerce')
    if numeric_cols.isna().all().all():
        return False, "No numeric data found"
    return True, f"✓ {companies.nunique()} companies · {fields.nunique()} fields · {df.shape[1]-2} periods"

def smart_field_mapping(field_name: str, options: list) -> int:
    field_clean = field_name.lower().replace(" ","").replace("_","").replace("&","").replace("-","")
    for i, opt in enumerate(options[1:], 1):
        opt_clean = opt.lower().replace(" ","").replace("_","").replace("&","").replace("-","")
        if opt_clean == field_clean: return i
    for i, opt in enumerate(options[1:], 1):
        opt_clean = opt.lower().replace(" ","").replace("_","").replace("&","").replace("-","")
        if field_clean in opt_clean or opt_clean in field_clean: return i
    synonym_map = {
        'revenue': ['sales','turnover'], 'costofrevenue': ['cogs'],
        'totalassets': ['assets'], 'totalequity': ['equity','shareholdersequity'],
        'totaldebt': ['debt'], 'cash': ['cashandequivalents'],
        'inventory': ['stock'], 'accountsreceivable': ['receivables','debtors'],
        'accountspayable': ['payables','creditors'], 'netincome': ['profit','earnings']
    }
    for i, opt in enumerate(options[1:], 1):
        opt_clean = opt.lower().replace(" ","").replace("_","").replace("&","").replace("-","")
        for key, synonyms in synonym_map.items():
            if key in field_clean and any(s in opt_clean for s in synonyms): return i
            if any(s in field_clean for s in synonyms) and key in opt_clean: return i
    return 0

def format_dataframe_preview(df: pd.DataFrame, max_rows: int = 20) -> pd.DataFrame:
    preview = df.head(max_rows).copy()
    cols = list(preview.columns)
    cols[0], cols[1] = "Company Name", "Financial Metric"
    for i in range(2, len(cols)):
        original = str(cols[i]).strip()
        cols[i] = f"FY {original}" if original.isdigit() and len(original) == 4 else f"Period {i-1}"
    preview.columns = cols
    for col in preview.columns[2:]:
        preview[col] = preview[col].apply(lambda x: f"{x:,.2f}" if pd.notna(x) else "—")
    return preview

def load_sample_file(sample_type: str) -> tuple:
    try:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        filename = "sample_CEMENT_companies_yearly.xlsx" if "Cement" in sample_type else "sample_PHARMA_companies_yearly.xlsx"
        filepath = os.path.join(script_dir, filename)
        df_raw = pd.read_excel(filepath, sheet_name=0, dtype=object)
        df_raw.columns = [str(x).strip() for x in df_raw.columns]
        df = df_raw.copy()
        for col_idx in range(2, len(df.columns)):
            df.iloc[:, col_idx] = pd.to_numeric(df.iloc[:, col_idx], errors='coerce')
        is_valid, msg = validate_file_structure(df)
        if not is_valid:
            return None, False, msg, {}
        field_names = [str(x) for x in df.iloc[:, 1].dropna().unique()]
        return df, True, msg, {field: field for field in field_names}
    except FileNotFoundError:
        return None, False, f"File '{filename}' not found in {script_dir}", {}
    except Exception as e:
        return None, False, f"Error: {str(e)}", {}




def generate_formula_pdf(md_text: str) -> bytes:
    """
    Convert formula_guide.md Markdown to a styled PDF using reportlab.
    Returns raw PDF bytes ready for st.download_button.
    Falls back to returning the Markdown as UTF-8 bytes if reportlab unavailable.
    """
    if not _REPORTLAB_OK:
        return md_text.encode("utf-8")

    buf = _io_module.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=2*cm, rightMargin=2*cm,
        topMargin=2.5*cm, bottomMargin=2.5*cm,
        title="Financial Ratio Analysis — Formula Guide",
        author="Financial Ratio Analysis Engine v3.0",
    )
    base = getSampleStyleSheet()
    S = {
        'h1': ParagraphStyle('h1', parent=base['Normal'],
            fontSize=22, fontName='Helvetica-Bold', spaceAfter=12,
            textColor=_rl_colors.HexColor('#0F172A'), leading=28),
        'h2': ParagraphStyle('h2', parent=base['Normal'],
            fontSize=14, fontName='Helvetica-Bold', spaceAfter=6, spaceBefore=16,
            textColor=_rl_colors.HexColor('#6366F1'), leading=20),
        'h3': ParagraphStyle('h3', parent=base['Normal'],
            fontSize=11, fontName='Helvetica-Bold', spaceAfter=4, spaceBefore=10,
            textColor=_rl_colors.HexColor('#1E293B'), leading=16),
        'body': ParagraphStyle('body', parent=base['Normal'],
            fontSize=9.5, fontName='Helvetica', spaceAfter=4,
            textColor=_rl_colors.HexColor('#374151'), leading=14),
        'bullet': ParagraphStyle('bullet', parent=base['Normal'],
            fontSize=9.5, fontName='Helvetica', spaceAfter=3,
            textColor=_rl_colors.HexColor('#374151'), leading=14,
            leftIndent=16, bulletIndent=6),
    }

    def _esc(t):
        return t.replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')

    def _fmt(t):
        t = _esc(t)
        t = _re_module.sub(r'\*\*(.+?)\*\*', r'<b>\1</b>', t)
        t = _re_module.sub(r'`(.+?)`', r'<font name="Courier" size="9">\1</font>', t)
        t = _re_module.sub(r'\*(.+?)\*', r'<i>\1</i>', t)
        return t

    def _flush_table(rows, doc_width):
        if not rows: return None
        nc = max(len(r) for r in rows)
        for r in rows:
            while len(r) < nc: r.append(Paragraph('', base['Normal']))
        cw = doc_width / nc
        t = Table(rows, colWidths=[cw]*nc, repeatRows=1)
        t.setStyle(TableStyle([
            ('BACKGROUND', (0,0),(-1,0), _rl_colors.HexColor('#6366F1')),
            ('TEXTCOLOR',  (0,0),(-1,0), _rl_colors.white),
            ('FONTNAME',   (0,0),(-1,0), 'Helvetica-Bold'),
            ('FONTSIZE',   (0,0),(-1,-1), 9),
            ('ROWBACKGROUNDS',(0,1),(-1,-1),
             [_rl_colors.HexColor('#F8FAFC'), _rl_colors.white]),
            ('GRID',       (0,0),(-1,-1), 0.5, _rl_colors.HexColor('#E2E8F0')),
            ('PADDING',    (0,0),(-1,-1), 5),
            ('VALIGN',     (0,0),(-1,-1), 'TOP'),
        ]))
        return t

    story = []
    trows = []
    in_tbl = False
    para_s = ParagraphStyle('tc', parent=base['Normal'], fontSize=9, leading=13)

    for line in md_text.split('\n'):
        s = line.strip()

        # Flush table on non-table line
        if in_tbl and (not s.startswith('|')):
            tbl = _flush_table(trows, doc.width)
            if tbl: story.append(tbl); story.append(Spacer(1,8))
            trows = []; in_tbl = False

        if not s:
            story.append(Spacer(1, 4)); continue

        if s.startswith('# ') and not s.startswith('## '):
            story.append(Paragraph(_fmt(s[2:]), S['h1']))
            story.append(HRFlowable(width='100%', thickness=2,
                color=_rl_colors.HexColor('#6366F1'), spaceAfter=8))
            continue
        if s.startswith('## '):
            story.append(Paragraph(_fmt(s[3:]), S['h2']))
            story.append(HRFlowable(width='100%', thickness=0.5,
                color=_rl_colors.HexColor('#E2E8F0'), spaceAfter=4))
            continue
        if s.startswith('### '):
            story.append(Paragraph(_fmt(s[4:]), S['h3'])); continue
        if s in ('---','***','___'):
            story.append(HRFlowable(width='100%', thickness=0.5,
                color=_rl_colors.HexColor('#CBD5E1'), spaceBefore=4, spaceAfter=4))
            continue
        if s.startswith('|'):
            cells = [c.strip() for c in s.split('|')[1:-1]]
            if all(_re_module.match(r'^[-:]+$', c) for c in cells if c):
                continue
            if not trows:
                pcells = [Paragraph(f'<b>{_esc(c)}</b>', para_s) for c in cells]
            else:
                pcells = [Paragraph(_fmt(c), para_s) for c in cells]
            trows.append(pcells); in_tbl = True; continue
        if s.startswith('- ') or s.startswith('* '):
            story.append(Paragraph(f'• {_fmt(s[2:])}', S['bullet'])); continue
        m = _re_module.match(r'^(\d+)\.\s+(.*)', s)
        if m:
            story.append(Paragraph(f'{m.group(1)}. {_fmt(m.group(2))}', S['bullet']))
            continue
        story.append(Paragraph(_fmt(s), S['body']))

    if in_tbl and trows:
        tbl = _flush_table(trows, doc.width)
        if tbl: story.append(tbl)

    doc.build(story)
    buf.seek(0)
    return buf.read()

def reset_analysis():
    st.session_state.engine = None
    st.session_state.validation_result = None
    st.session_state.data_quality = None
    st.session_state.analysis_complete = False
    st.session_state.excel_payload = None
    st.session_state.dash_payloads = {}

def _load_dataframe_from_excel(source, name: str):
    """Read an Excel file from either a path string or an UploadedFile object."""
    df_str = pd.read_excel(source, sheet_name=0, dtype=object)
    df_str.columns = [str(c).strip() for c in df_str.columns]
    df = df_str.copy()
    for i in range(2, len(df.columns)):
        df.iloc[:, i] = pd.to_numeric(df.iloc[:, i], errors='coerce')
    return df

# ── Header ─────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="app-header">
  <div class="header-inner">
    <div class="header-icon">📊</div>
    <div class="header-text">
      <h1>Financial Ratio Analysis</h1>
      <p>Multi-company ratio analysis &amp; peer benchmarking</p>
    </div>
  </div>
</div>
""", unsafe_allow_html=True)


# ── Resources bar — globally visible, no sidebar needed ───────────────────────
_script_dir_res = os.path.dirname(os.path.abspath(__file__))
_guide_path_res  = os.path.join(_script_dir_res, "formula_guide.md")
_cement_path_res = os.path.join(_script_dir_res, "sample_CEMENT_companies_yearly.xlsx")
_pharma_path_res = os.path.join(_script_dir_res, "sample_PHARMA_companies_yearly.xlsx")

_res_col1, _res_col2, _res_col3, _res_spacer = st.columns([1.4, 1.4, 1.4, 3])

with _res_col1:
    if os.path.isfile(_guide_path_res) and _REPORTLAB_OK:
        # Generate PDF on each session (cached via session_state)
        if 'formula_pdf_cache' not in st.session_state:
            try:
                with open(_guide_path_res, 'r', encoding='utf-8') as _gf:
                    _guide_md = _gf.read()
                st.session_state.formula_pdf_cache = generate_formula_pdf(_guide_md)
            except Exception:
                st.session_state.formula_pdf_cache = None
        if st.session_state.get('formula_pdf_cache'):
            st.download_button(
                "📖 Formula Guide (PDF)",
                data=st.session_state.formula_pdf_cache,
                file_name="Financial_Ratio_Formula_Guide_v3.pdf",
                mime="application/pdf",
                use_container_width=True,
                help="All 60+ ratio formulas, edge cases, and calculation nuances — PDF format"
            )
        else:
            st.caption("📖 Formula Guide unavailable")
    elif os.path.isfile(_guide_path_res):
        # reportlab not available — offer Markdown as fallback
        with open(_guide_path_res, 'r', encoding='utf-8') as _gf:
            _guide_md = _gf.read()
        st.download_button(
            "📖 Formula Guide (MD)",
            data=_guide_md.encode('utf-8'),
            file_name="Financial_Ratio_Formula_Guide_v3.md",
            mime="text/markdown",
            use_container_width=True,
            help="All 60+ ratio formulas — Markdown format (reportlab not installed for PDF)"
        )
    else:
        st.caption("📖 Formula Guide not found")

with _res_col2:
    if os.path.isfile(_cement_path_res):
        with open(_cement_path_res, 'rb') as _cf:
            _cement_bytes = _cf.read()
        st.download_button(
            "📊 Sample: Cement Data",
            data=_cement_bytes,
            file_name="sample_CEMENT_companies_yearly.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            help="Sample dataset: 4 cement companies, annual data. Use as a template for your own data."
        )
    else:
        st.caption("📊 Cement sample not found")

with _res_col3:
    if os.path.isfile(_pharma_path_res):
        with open(_pharma_path_res, 'rb') as _pf:
            _pharma_bytes = _pf.read()
        st.download_button(
            "💊 Sample: Pharma Data",
            data=_pharma_bytes,
            file_name="sample_PHARMA_companies_yearly.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            help="Sample dataset: 4 pharma companies, annual data. Use as a template for your own data."
        )
    else:
        st.caption("💊 Pharma sample not found")

st.markdown('<div style="height:4px;"></div>', unsafe_allow_html=True)

# ── Workflow stepper ───────────────────────────────────────────────────────────
def show_stepper():
    current = st.session_state.workflow_step
    steps = [
        ("1", "Upload", "Load your Excel data"),
        ("2", "Map Fields", "Connect to standard metrics"),
        ("3", "Results", "Explore 60+ ratios"),
    ]
    parts = []
    for i, (num, name, desc) in enumerate(steps):
        s = i + 1
        if s < current:
            circle_cls = "done"; name_cls = "done-text"; icon = "✓"
            conn_cls = "done"
        elif s == current:
            circle_cls = "active"; name_cls = "active-text"; icon = num
            conn_cls = ""
        else:
            circle_cls = "pending"; name_cls = ""; icon = num
            conn_cls = ""

        node = f"""
        <div class="step-node">
          <div class="step-circle {circle_cls}">{icon}</div>
          <div class="step-label">
            <div class="step-name {name_cls}">{name}</div>
            <div class="step-desc">{desc}</div>
          </div>
        </div>"""
        parts.append(node)
        if i < len(steps) - 1:
            parts.append(f'<div class="step-connector {conn_cls}"></div>')

    st.markdown(f'<div class="stepper">{"".join(parts)}</div>', unsafe_allow_html=True)

show_stepper()

# ── Tabs ───────────────────────────────────────────────────────────────────────
tab1, tab2, tab3 = st.tabs([
    "📂  Step 1 · Data Upload",
    "🔗  Step 2 · Field Mapping",
    "📊  Step 3 · Ratio Analysis",
])



# ══════════════════════════════════════════════════════════════════════════════
# STEP 1 — Upload
# ══════════════════════════════════════════════════════════════════════════════
with tab1:
    if st.session_state.show_sample_success:
        show_alert("✅ <strong>Sample data loaded.</strong> All fields pre-mapped. Head to Step 2 to review or Step 3 to calculate.", "success")
        st.session_state.show_sample_success = False

    left_col, right_col = st.columns([1, 2], gap="large")

    # ── Left: settings panel ──────────────────────────
    with left_col:
        st.markdown('<div class="section-title">⚙️ Analysis Settings</div>', unsafe_allow_html=True)

        st.session_state.industry = st.radio(
            "Industry Type", ["Manufacturing", "Service"],
            help="Manufacturing: includes inventory ratios. Service: payables denominator uses OpEx proxy."
        )
        st.session_state.frequency = st.selectbox(
            "Data Frequency", ["Annual", "Quarterly", "Monthly"],
            help="Adjusts turnover, days, and trailing-sum calculations."
        )
        st.session_state.tax_rate = st.number_input(
            "Corporate Tax Rate (%)", 0.0, 100.0, 29.0, 0.5,
            help="Used in NOPAT = Operating Income × (1 − t). Impacts ROIC."
        )
        if st.session_state.tax_rate == 0:
            st.info("ℹ️ At 0% tax, NOPAT equals Operating Income.")

        with st.expander("🔬 Advanced Parameters"):
            st.session_state.operating_cash_pct = st.number_input(
                "Operating Cash as % of Revenue", 0.0, 20.0, 2.0, 0.5,
                help="Minimum cash required for operations. Excess excluded from Invested Capital for ROIC."
            )
            st.caption("⚠️ Materially impacts ROIC. Default 2% is an estimate — adjust per industry.")
            st.session_state.allow_negative_coverage = st.checkbox(
                "Show Negative Interest Coverage", value=True,
                help="When enabled, distressed companies show negative coverage instead of N/A."
            )
            st.session_state.nopat_tax_benefit = st.checkbox(
                "Tax Benefit on Operating Losses", value=True,
                help="Standard: NOPAT = OpInc × (1−t) always. Uncheck = conservative (no benefit on losses)."
            )

        with st.expander("📖 Required File Format"):
            st.markdown("""
**Excel layout — one row per company-metric:**

| Company | Field | 2021 | 2022 | 2023 |
|---|---|---|---|---|
| ABC Ltd | Revenue | 1000 | 1200 | 1500 |
| ABC Ltd | Net Income | 100 | 120 | 150 |
| ABC Ltd | Total Assets | 2000 | 2200 | 2500 |
| XYZ Corp | Revenue | 5000 | 5500 | 6000 |

**Rules:**
- Col 1: Company name (repeat for each metric)
- Col 2: Field name  
- Col 3+: Period data (all numeric)
- No merged cells · No trailing notes/empty columns
""")

    # ── Right: data source ────────────────────────────
    with right_col:
        st.markdown('<div class="section-title">📤 Data Source</div>', unsafe_allow_html=True)

        # ── Source 1: Manual Upload ───────────────────
        with st.expander("📁  Upload Excel File", expanded=True):
            uploaded = st.file_uploader(
                "Drag & drop or click to browse",
                type=['xlsx', 'xls'], label_visibility="collapsed"
            )
            if uploaded:
                if st.session_state.file_name != uploaded.name:
                    st.session_state.file_name = uploaded.name
                    reset_analysis()
                    st.session_state.mapping = {}
                    st.session_state.workflow_step = 1
                try:
                    with st.spinner("Reading file…"):
                        df = _load_dataframe_from_excel(uploaded, uploaded.name)
                    is_valid, message = validate_file_structure(df)
                    if not is_valid:
                        show_alert(f"❌ <strong>Invalid structure:</strong> {message}", "error")
                        st.session_state.raw_df = None
                        st.session_state.file_valid = False
                    else:
                        st.session_state.raw_df = df
                        st.session_state.file_valid = True
                        st.session_state.workflow_step = 2
                        show_alert(f"✅ <strong>{uploaded.name}</strong> — {message}. Proceed to Step 2.", "success")
                        c1, c2, c3, c4 = st.columns(4)
                        c1.metric("Rows",      f"{df.shape[0]:,}")
                        c2.metric("Columns",   f"{df.shape[1]}")
                        c3.metric("Companies", f"{df.iloc[:,0].dropna().nunique()}")
                        c4.metric("Periods",   f"{df.shape[1]-2}")
                        st.markdown("**Preview**")
                        st.dataframe(format_dataframe_preview(df, 25), use_container_width=True, height=400)
                except Exception as e:
                    show_alert(f"❌ <strong>Error reading file:</strong> {_user_friendly_error(e)}", "error")
                    st.session_state.raw_df = None
                    st.session_state.file_valid = False

        # ── Source 2: Sample Data ─────────────────────
        with st.expander("🧪  Load or Download Sample Dataset"):
            st.caption(
                "Use these sample files to understand the required data format, "
                "or load them directly into the app to explore the engine."
            )
            sc1, sc2, sc3 = st.columns([2, 1, 1])
            with sc1:
                sample_type = st.selectbox(
                    "Sample", ["📊 Cement (4 Companies)", "💊 Pharma (4 Companies)"],
                    label_visibility="collapsed"
                )
            with sc2:
                if st.button("⚡ Load into App", use_container_width=True, type="primary"):
                    with st.spinner("Loading sample…"):
                        df, valid, msg, mapping = load_sample_file(sample_type)
                        if valid:
                            st.session_state.raw_df = df
                            st.session_state.file_valid = True
                            st.session_state.file_name = sample_type
                            st.session_state.mapping = mapping
                            st.session_state.industry = "Manufacturing"
                            st.session_state.workflow_step = 2
                            st.session_state.show_sample_success = True
                            reset_analysis()
                            time.sleep(0.3)
                            st.rerun()
                        else:
                            st.error(msg)
            with sc3:
                # Download the raw .xlsx file so users can study the format
                _script_dir_s = os.path.dirname(os.path.abspath(__file__))
                _is_cement = "Cement" in sample_type
                _sfile = ("sample_CEMENT_companies_yearly.xlsx" if _is_cement
                          else "sample_PHARMA_companies_yearly.xlsx")
                _spath = os.path.join(_script_dir_s, _sfile)
                if os.path.isfile(_spath):
                    with open(_spath, 'rb') as _sf:
                        _sbytes = _sf.read()
                    st.download_button(
                        "📥 Download .xlsx",
                        data=_sbytes,
                        file_name=_sfile,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        help="Download the sample Excel file to use as a template for your own data"
                    )
                else:
                    st.caption("File not found")

# ══════════════════════════════════════════════════════════════════════════════
# STEP 2 — Field Mapping
# ══════════════════════════════════════════════════════════════════════════════
with tab2:
    if not st.session_state.file_valid or st.session_state.raw_df is None:
        st.markdown("""
        <div class="empty-state">
          <div class="empty-state-icon">🔗</div>
          <h3>No Data Loaded</h3>
          <p>Upload your Excel file in <strong>Step 1</strong> first, then return here to map your fields.</p>
        </div>""", unsafe_allow_html=True)
    else:
        available_fields = sorted([str(x) for x in st.session_state.raw_df.iloc[:,1].dropna().unique()])
        field_options = ["(Not Mapped)"] + available_fields

        mapping_categories = {
            "💰 Income Statement": {
                "fields": ["Revenue","Cost of Revenue","Gross Profit","Operating Expenses",
                           "Operating Income","Interest Expense","Net Income","EBITDA","D&A"],
                "critical": ["Revenue"]
            },
            "🏦 Balance Sheet — Assets": {
                "fields": ["Cash & Equivalents","Short-Term Investments","Accounts Receivable",
                           "Inventory","Total Current Assets","PP&E (Net)","Total Assets"],
                "critical": ["Total Assets"]
            },
            "📊 Balance Sheet — Liabilities & Equity": {
                "fields": ["Accounts Payable","Short-Term Debt","Total Current Liabilities",
                           "Long-Term Debt","Total Debt","Total Liabilities","Total Equity",
                           "Retained Earnings","Minority Interest"],
                "critical": ["Total Equity"]
            },
            "💵 Cash Flow": {
                "fields": ["Operating Cash Flow","Capital Expenditures","Free Cash Flow","Dividends Paid"],
                "critical": []
            },
            "📈 Market Data": {
                "fields": ["Share Price","Shares Outstanding (Basic)","Shares Outstanding (Diluted)","Market Cap"],
                "critical": []
            },
        }

        mapped_count = required_count = 0

        # Count pass first
        for cat_data in mapping_categories.values():
            for field in cat_data['fields']:
                key = f"field_{field}"
                if key in st.session_state:
                    val = st.session_state[key]
                    if val and val != "(Not Mapped)":
                        mapped_count += 1
                        if field in cat_data['critical']:
                            required_count += 1

        # ── Progress bar ──────────────────────────────
        total_fields = sum(len(v['fields']) for v in mapping_categories.values())
        pct = int(mapped_count / total_fields * 100) if total_fields else 0
        pct_color = "#10B981" if pct >= 70 else ("#F59E0B" if pct >= 40 else "#6366F1")

        col_prog, col_stats = st.columns([2, 1])
        with col_prog:
            st.markdown(f"""
            <div style="margin-bottom:1rem;">
              <div style="display:flex;justify-content:space-between;margin-bottom:.4rem;">
                <span style="font-weight:600;font-size:.88rem;color:var(--prog-label);">Mapping Progress</span>
                <span style="font-weight:700;font-size:.88rem;color:{pct_color};">{pct}% · {mapped_count}/{total_fields} fields</span>
              </div>
              <div class="prog-bar-wrap">
                <div class="prog-bar-fill" style="width:{pct}%;"></div>
              </div>
            </div>""", unsafe_allow_html=True)

        with col_stats:
            req_col, rec_col = st.columns(2)
            recommended = ["Net Income","Operating Income","Total Debt","Operating Cash Flow","Shares Outstanding (Basic)"]
            rec_count = sum(1 for f in recommended if f in st.session_state.mapping)
            req_col.metric("Required", f"{required_count}/3",
                          delta="✅" if required_count == 3 else "❌ Missing")
            rec_col.metric("Recommended", f"{rec_count}/5")

        st.markdown("---")

        # ── Mapping expanders ─────────────────────────
        mapped_count = required_count = 0  # reset for actual widget pass

        for category, cat_data in mapping_categories.items():
            cat_total = len(cat_data['fields'])
            cat_mapped = sum(
                1 for f in cat_data['fields']
                if f"field_{f}" in st.session_state
                and st.session_state[f"field_{f}"] not in [None,"(Not Mapped)"]
            )
            # st.expander labels are PLAIN TEXT only — HTML is never rendered there.
            # Show completion status as plain text in the label, badge inside.
            status_icon = "✅" if cat_mapped == cat_total else ("🔶" if cat_mapped > 0 else "⬜")
            expander_label = f"{status_icon} {category}  ({cat_mapped}/{cat_total} mapped)"

            badge_cls = "complete" if cat_mapped == cat_total else ("partial" if cat_mapped > 0 else "")
            badge_html = f'<span class="cat-badge {badge_cls}">{cat_mapped} of {cat_total} fields mapped</span>'

            with st.expander(expander_label,
                             expanded=(category == "💰 Income Statement")):
                st.markdown(badge_html, unsafe_allow_html=True)
                cols = st.columns(2)
                for idx, field in enumerate(cat_data['fields']):
                    with cols[idx % 2]:
                        is_critical = field in cat_data['critical']
                        label = (f"🔴 **{field}** *(required)*" if is_critical else field)
                        help_text = REQUIRED_FIELDS.get(field, "")
                        if is_critical:
                            help_text += " | ⚠️ REQUIRED for calculation"
                        default_idx = smart_field_mapping(field, field_options)
                        selection = st.selectbox(
                            label, field_options,
                            index=default_idx, key=f"field_{field}", help=help_text
                        )
                        if selection != "(Not Mapped)":
                            st.session_state.mapping[field] = selection
                            mapped_count += 1
                            if is_critical:
                                required_count += 1
                        elif field in st.session_state.mapping:
                            del st.session_state.mapping[field]

        # ── Review + run ──────────────────────────────
        if mapped_count > 0:
            with st.expander("🔍 Review All Mappings"):
                map_df = pd.DataFrame([
                    {"Standard Field": k, "Your File Field": v}
                    for k, v in sorted(st.session_state.mapping.items())
                ])
                st.dataframe(map_df, use_container_width=True, hide_index=True)

        st.markdown("---")
        can_analyze = required_count == 3 and mapped_count >= 5

        run_col, status_col = st.columns([1, 2])
        with run_col:
            if st.button("🚀 Calculate Ratios",
                         disabled=not can_analyze,
                         use_container_width=True, type="primary"):
                with st.spinner("⚙️ Computing 60+ financial ratios…"):
                    try:
                        config = EngineConfig(
                            operating_cash_pct=st.session_state.operating_cash_pct / 100.0,
                            allow_negative_interest_coverage=st.session_state.allow_negative_coverage,
                            nopat_tax_benefit_on_losses=st.session_state.get('nopat_tax_benefit', True)
                        )
                        engine = RatioEngine(
                            st.session_state.raw_df,
                            st.session_state.mapping,
                            st.session_state.industry,
                            st.session_state.tax_rate / 100.0,
                            frequency=st.session_state.frequency,
                            config=config
                        )
                        engine.run_calculation()
                        st.session_state.engine = engine
                        st.session_state.excel_payload = None
                        st.session_state.dash_payloads = {}
                        st.session_state.analysis_complete = True
                        st.session_state.workflow_step = 3
                        st.session_state.data_quality = engine.get_data_quality_report()
                        st.balloons()
                        show_alert("🎉 <strong>Calculation complete!</strong> Switch to Step 3 to explore results.", "success")
                        q1, q2, q3 = st.columns(3)
                        q1.metric("Companies", len(engine.companies))
                        q2.metric("Periods",   len(engine.years))
                        q3.metric("Ratios",    len(engine.generate_peer_matrix()))
                    except Exception as e:
                        show_alert(f"❌ <strong>Error:</strong> {_user_friendly_error(e)}", "error")

        with status_col:
            if not can_analyze:
                if required_count < 3:
                    show_alert("⚠️ <strong>Missing required fields.</strong> Map: Revenue 🔴, Total Assets 🔴, Total Equity 🔴", "warning")
                else:
                    show_alert("⚠️ Map at least <strong>5 fields total</strong> before calculating.", "warning")
            else:
                show_alert(f"✅ <strong>Ready.</strong> {mapped_count} fields mapped. Click Calculate to proceed.", "success")

# ══════════════════════════════════════════════════════════════════════════════
# STEP 3 — Results
# ══════════════════════════════════════════════════════════════════════════════
with tab3:
    if not st.session_state.analysis_complete or st.session_state.engine is None:
        st.markdown("""
        <div class=\"empty-state\">
          <div class=\"empty-state-icon\">📊</div>
          <h3>No Results Yet</h3>
          <p>Complete <strong>Step 1</strong> (upload) → <strong>Step 2</strong> (map &amp; calculate).</p>
        </div>""", unsafe_allow_html=True)
    else:
        engine = st.session_state.engine
        try:
            matrices = engine.generate_peer_matrix()
            ratio_names_all = sorted(matrices.keys())
        except Exception as e:
            show_alert(f'❌ <strong>Error:</strong> {_user_friendly_error(e)}', 'error')
            ratio_names_all = []

        if ratio_names_all:
            # ── Sidebar Analysis Controls ─────────────────────────────────
            all_companies = engine.companies
            with st.sidebar:
                st.markdown('<div class="sidebar-section-label" style="font-size:.85rem;color:white!important;margin-top:1rem;">🎛️ ANALYSIS CONTROLS</div>', unsafe_allow_html=True)
                st.markdown('<div class="sidebar-divider"></div>', unsafe_allow_html=True)
                
                _cat_names = list(RATIO_CATEGORIES.keys())
                _prev_cat = st.session_state.get('sb_cat', _cat_names[0])
                _cat_idx = _cat_names.index(_prev_cat) if _prev_cat in _cat_names else 0
                selected_cat = st.selectbox('📁 Category', _cat_names, index=_cat_idx, key='sb_cat')
                
                cat_ratios = [r for r in RATIO_CATEGORIES.get(selected_cat, []) if r in matrices]
                if not cat_ratios:
                    cat_ratios = sorted(ratio_names_all)
                _prev_ratio = st.session_state.get('sb_ratio', cat_ratios[0])
                _ratio_idx = cat_ratios.index(_prev_ratio) if _prev_ratio in cat_ratios else 0
                selected_ratio = st.selectbox('📊 Ratio', cat_ratios, index=_ratio_idx, key='sb_ratio')
                
                chart_style = st.radio('Chart Style', ['📈 Line', '📊 Bar', '🏔️ Area'], key='sb_chart')
                
                st.markdown('<div class="sidebar-divider"></div>', unsafe_allow_html=True)
                
                _prev_cos = st.session_state.get('sb_companies',
                    all_companies[:6] if len(all_companies) > 6 else all_companies)
                _prev_cos = [c for c in (_prev_cos or []) if c in all_companies] or (
                    all_companies[:6] if len(all_companies) > 6 else all_companies)
                selected_companies = st.multiselect('🏢 Companies', all_companies,
                    default=_prev_cos, key='sb_companies')
                
                show_mean = st.checkbox('Peer Mean', value=True, key='sb_mean')
                show_median = st.checkbox('Peer Median', value=False, key='sb_median')
                st.markdown('<div class="sidebar-divider"></div>', unsafe_allow_html=True)

            # Validate selections
            if selected_ratio not in matrices:
                selected_ratio = cat_ratios[0] if cat_ratios else ratio_names_all[0]
            selected_companies = [c for c in selected_companies if c in engine.companies]
            if not selected_companies:
                selected_companies = all_companies[:6] if len(all_companies) > 6 else all_companies

            if not selected_companies:
                show_alert('⚠️ <strong>No companies selected.</strong> Choose at least one from the Companies dropdown above.', 'warning')
            else:
                if selected_ratio in RATIO_EXPLANATIONS:
                    st.markdown(
                        f'<div class="ratio-banner"><strong>{selected_ratio}</strong>'
                        f' — {RATIO_EXPLANATIONS[selected_ratio]}</div>',
                        unsafe_allow_html=True
                    )
                st.markdown(
                    f'<div class="breadcrumb">'
                    f'<span class="breadcrumb-cat">{selected_cat}</span>'
                    f'<span class="breadcrumb-arrow">›</span>'
                    f'<span class="breadcrumb-title">{selected_ratio}</span></div>',
                    unsafe_allow_html=True
                )

                try:
                    df_ratio = matrices[selected_ratio]
                    companies_to_show = [c for c in selected_companies if c in df_ratio.index]
                    if show_mean   and 'Mean'   in df_ratio.index: companies_to_show.append('Mean')
                    if show_median and 'Median' in df_ratio.index: companies_to_show.append('Median')

                    if not companies_to_show:
                        show_alert('❌ No data for selected companies.', 'error')
                    else:
                        df_display = df_ratio.loc[companies_to_show]
                        if df_display.isna().all().all():
                            show_alert(
                                f'⚠️ <strong>{selected_ratio}</strong> requires fields not currently mapped.',
                                'warning'
                            )
                        else:
                            pct_keywords = [
                                'Margin','ROA','ROE','ROIC','Growth','Return','Yield',
                                'Payout','OCF to','FCF to','NWC to','Debt to Assets',
                                'Debt to Capital','DuPont: Net','Accrual','Prob',
                                'Reinvestment','Sustainable','Expense Ratio'
                            ]
                            is_percentage = any(kw in selected_ratio for kw in pct_keywords)

                            def fmt(val):
                                if pd.isna(val): return '—'
                                return f'{val:.2%}' if is_percentage else f'{val:,.2f}'

                            st.markdown('**📐 Statistics**')
                            stat_labels = ['Mean', 'Median', 'Std Dev']
                            company_only = df_display.drop(
                                index=[i for i in df_display.index if i in stat_labels],
                                errors='ignore'
                            )
                            latest_col  = company_only.columns[-1]
                            latest_vals = company_only[latest_col].dropna()
                            all_vals    = company_only.values.flatten()
                            all_vals    = all_vals[~pd.isna(all_vals)]

                            def fv(v):
                                return f'{v:.2%}' if is_percentage else f'{v:,.2f}'

                            c1, c2, c3, c4 = st.columns(4)
                            with c1:
                                if not latest_vals.empty:
                                    top = latest_vals.idxmax()
                                    st.metric(f'🏆 Best ({latest_col})', top, fv(latest_vals.max()))
                            with c2:
                                if len(all_vals) > 0:
                                    st.metric('⌀ Average', fv(all_vals.mean()))
                            with c3:
                                if len(all_vals) > 1:
                                    st.metric('σ Std Dev',  fv(all_vals.std()))
                            with c4:
                                if len(all_vals) > 1:
                                    st.metric('↕ Range',    fv(all_vals.max() - all_vals.min()))
                            
                            st.markdown('<br>', unsafe_allow_html=True)
                            st.markdown('**📋 Data Table**')
                            valid_data = df_display.values.flatten()
                            valid_data = valid_data[~pd.isna(valid_data)]
                            if len(valid_data) > 0:
                                q1v, q3v = pd.Series(valid_data).quantile([0.25, 0.75])
                                iqr  = q3v - q1v
                                vmin = max(valid_data.min(), q1v - 1.5 * iqr)
                                vmax = min(valid_data.max(), q3v + 1.5 * iqr)
                                # Ratios where LOWER values are better — invert colormap
                                _inverted_ratios = {
                                    'Beneish M-Score', 'Sloan Accrual Ratio',
                                    'Ohlson Bankruptcy Prob',
                                    'Debt to Equity', 'Net Debt to Equity',
                                    'Debt to Assets', 'Debt to Capital',
                                    'Net Debt to EBITDA', 'Financial Leverage',
                                    'Interest Expense Ratio',
                                    'P/E Ratio', 'PEG Ratio',
                                    'Price to Book (P/B)', 'Price to Sales',
                                    'EV / EBITDA', 'EV / Revenue',
                                    'Days Sales Outstanding (DSO)',
                                    'Days Inventory (DIO)',
                                    'Cash Conversion Cycle',
                                    'Dividend Payout',
                                }
                                _cmap = ('RdYlGn_r' if selected_ratio in _inverted_ratios
                                         else 'RdYlGn')
                                st.dataframe(
                                    df_display.style
                                        .background_gradient(cmap=_cmap, vmin=vmin, vmax=vmax)
                                        .format(fmt),
                                    use_container_width=True,
                                    height=min(420, 60 + len(companies_to_show) * 40)
                                )
                            else:
                                st.dataframe(df_display.map(fmt), use_container_width=True)
                            
                            st.markdown('<br>', unsafe_allow_html=True)

                            st.markdown('**📈 Trend Chart**')
                            chart_data = df_display.T.reset_index().melt(
                                id_vars='index', var_name='Company', value_name='Value'
                            )
                            chart_data.rename(columns={'index': 'Period'}, inplace=True)
                            chart_data = chart_data.dropna(subset=['Value'])

                            if not chart_data.empty:
                                style = chart_style.split(' ')[1]
                                if style == 'Line':
                                    fig = px.line(chart_data, x='Period', y='Value',
                                                  color='Company', markers=True,
                                                  title=f'{selected_ratio} — Trend')
                                elif style == 'Bar':
                                    fig = px.bar(chart_data, x='Period', y='Value',
                                                 color='Company', barmode='group',
                                                 title=f'{selected_ratio} — Comparison')
                                else:
                                    fig = px.area(chart_data, x='Period', y='Value',
                                                  color='Company',
                                                  title=f'{selected_ratio} — Cumulative')
                                fig.update_layout(
                                    template='plotly', height=480, hovermode='x unified',
                                    font=dict(family='Inter, sans-serif', size=12),
                                    title_font_size=15,
                                    legend=dict(orientation='h', yanchor='bottom',
                                                y=1.02, xanchor='right', x=1),
                                    xaxis=dict(type='category', categoryorder='array',
                                               categoryarray=engine.years),
                                    plot_bgcolor='rgba(0,0,0,0)',
                                    paper_bgcolor='rgba(0,0,0,0)',
                                    margin=dict(t=50, b=40, l=50, r=30)
                                )
                                if is_percentage:
                                    fig.update_yaxes(tickformat='.1%')
                                fig.update_xaxes(showgrid=False)
                                fig.update_yaxes(showgrid=True, gridcolor='rgba(148,163,184,.25)')
                                st.plotly_chart(fig, use_container_width=True)
                            else:
                                show_alert(
                                    f'⚠️ No plottable data for <strong>{selected_ratio}</strong>'
                                    f' with the selected companies.', 'warning'
                                )

                except Exception as e:
                    show_alert(f'❌ <strong>Error:</strong> {str(e)}', 'error')

            # ── Export ──────────────────────────────────────────────────
            st.markdown('<div class="section-title">📥 Export Results</div>', unsafe_allow_html=True)
            exp1, exp2 = st.columns(2, gap='large')

            with exp1:
                st.markdown('<div class="export-card"><div class="export-card-title">📊 Excel Workbook</div>'
                    '<div class="export-card-body">All ratios by category across separate sheets, '
                    'with peer mean and median rows.</div>', unsafe_allow_html=True)
                
                if 'excel_payload' not in st.session_state:
                    st.session_state.excel_payload = None
                
                if st.button("⚙️ Generate Excel Report", use_container_width=True, key="btn_gen_excel"):
                    try:
                        with st.spinner("⏳ Rendering Excel Data..."):
                            time.sleep(1.0)
                            buffer = io.BytesIO()
                            engine.export_excel(buffer)
                            st.session_state.excel_payload = buffer.getvalue()
                    except Exception as e:
                        st.error(f'Export error: {_user_friendly_error(e)}')

                if st.session_state.excel_payload:
                    st.success("✅ Excel Workbook Ready!")
                    filename = f"RatioAnalysis_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.xlsx"
                    st.download_button(
                        '📥 Download Excel Report',
                        data=st.session_state.excel_payload, file_name=filename,
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        use_container_width=True
                    )
                st.markdown('</div>', unsafe_allow_html=True)

            with exp2:
                st.markdown('<div class="export-card"><div class="export-card-title">🌐 HTML Dashboard</div>'
                    '<div class="export-card-body">Interactive offline dashboard with 9 charts, '
                    'peer comparison, and automated insights per company.</div>', unsafe_allow_html=True)
                if DASHBOARD_AVAILABLE:
                    dashboard_company = st.selectbox(
                        'Company', engine.companies, label_visibility='collapsed'
                    )
                    
                    if 'dash_payloads' not in st.session_state:
                        st.session_state.dash_payloads = {}

                    if st.button(f"⚙️ Generate Dashboard", use_container_width=True, key=f"btn_gen_dash_{dashboard_company.replace(' ','_')}"):
                        try:
                            with st.spinner(f"⏳ Rendering Dashboard for {dashboard_company}..."):
                                time.sleep(1.0)
                                dash_gen = DashboardGenerator(engine)
                                html_content = dash_gen.generate_html(dashboard_company)
                                st.session_state.dash_payloads[dashboard_company] = html_content
                        except Exception as e:
                            st.error(f'Dashboard error: {_user_friendly_error(e)}')

                    if dashboard_company in st.session_state.dash_payloads:
                        st.success(f"✅ Dashboard for {dashboard_company} Ready!")
                        filename = f"Dashboard_{dashboard_company.replace(' ','_')}_{pd.Timestamp.now().strftime('%Y%m%d')}.html"
                        st.download_button(
                            '📥 Download Dashboard',
                            data=st.session_state.dash_payloads[dashboard_company], file_name=filename,
                            mime='text/html', use_container_width=True,
                            key=f"dl_dash_{dashboard_company.replace(' ','_')}"
                        )
                else:
                    st.warning('⚠️ Dashboard module unavailable (Dashboard.py not found).')
                st.markdown('</div>', unsafe_allow_html=True)



# ── Footer ─────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="app-footer">
  <strong>Financial Ratio Analysis</strong>
  <span style="font-size:.75rem;opacity:.7;margin-top:.3rem;display:block;">
    For educational and analytical purposes only. Not financial advice.
    All calculations are estimates — verify independently before decisions.
  </span>
</div>
""", unsafe_allow_html=True)