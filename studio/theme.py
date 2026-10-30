"""BPM Studio look and feel: the platform logo and the shared stylesheet.

Everything here is used by every tool, so a change here changes the whole platform.
Tool-specific styling lives next to the tool (for example tools/testx/style.py).
"""
from __future__ import annotations

import streamlit as st

PLATFORM = "BPM Studio"
PLATFORM_TAGLINE = "From Process to Possibilities"
VERSION = "v1.1"


def bpm_logo(size=40) -> str:
    """Hexagon overlapping a circle: raw information + continuous process."""
    return f'''<svg width="{size}" height="{size}" viewBox="0 0 48 48" xmlns="http://www.w3.org/2000/svg" aria-label="BPM Studio logo">
<defs><linearGradient id="bpmH" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#0879D9"/><stop offset="1" stop-color="#00A6B2"/></linearGradient>
<linearGradient id="bpmC" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#00A6B2"/><stop offset="1" stop-color="#19C995"/></linearGradient></defs>
<polygon points="17,4 30,11.5 30,26.5 17,34 4,26.5 4,11.5" fill="url(#bpmH)"/>
<circle cx="30" cy="29" r="13" fill="url(#bpmC)" fill-opacity=".9" stroke="#fff" stroke-width="1.6"/></svg>'''



SHARED_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
html, body, [class*="css"], .stMarkdown, button, input, textarea { font-family: 'Inter','Segoe UI',system-ui,sans-serif !important; }
.block-container { padding-top: 1.1rem; padding-bottom: 1.5rem; max-width: 1180px; }
#MainMenu, footer { visibility: hidden; }
header[data-testid="stHeader"] { background: transparent; height: 0; }
h1,h2,h3,h4 { color:#0B2D55; letter-spacing:-.01em; }

/* ---- sidebar (deep navy) ---- */
[data-testid="stSidebar"] { background:#0B2D55; border-right:0; }
[data-testid="stSidebar"] * { color:#DCE7F5; }
[data-testid="stSidebar"] .stButton button > div { justify-content:flex-start !important; width:100%; text-align:left; }
[data-testid="stSidebar"] .stButton button { justify-content:flex-start; text-align:left; border-radius:10px; font-weight:600; padding:.55rem .8rem; border:0; }
[data-testid="stSidebar"] .stButton button[kind="tertiary"] { background:transparent; color:#C5D5EA; }
[data-testid="stSidebar"] .stButton button[kind="tertiary"]:hover { background:rgba(255,255,255,.08); color:#fff; }
[data-testid="stSidebar"] .stButton button[kind="primary"] { background:#0879D9; color:#fff; box-shadow:none; }
[data-testid="stSidebar"] hr { border-color: rgba(255,255,255,.12); }
.sb-brand { display:flex; align-items:center; gap:11px; padding: 4px 2px 2px 2px; }
.sb-brand .n { font-size:19px; font-weight:800; color:#fff !important; letter-spacing:-.01em; line-height:1.1; }
.sb-brand .t { font-size:11px; color:#8FB2D8 !important; letter-spacing:.02em; margin-top:2px; }
.sb-label { font-size:11px; font-weight:700; letter-spacing:.12em; text-transform:uppercase; color:#7FA3CC !important; margin: 6px 0 4px 4px; }
.sb-soon { border:1px dashed rgba(255,255,255,.22); border-radius:10px; padding:10px 12px; font-size:12.5px; color:#9DB9D9 !important; line-height:1.45; }
.sb-foot { font-size:11.5px; color:#7FA3CC !important; line-height:1.5; }

/* ---- header ---- */
.crumb { font-size:12.5px; color:#63758B; font-weight:500; }
.crumb b { color:#0B2D55; }
.page-title { font-size:26px; font-weight:800; color:#0B2D55; margin:2px 0 0 0; letter-spacing:-.02em; }
.id-chip { display:inline-flex; align-items:center; gap:8px; font-size:12.5px; font-weight:700; color:#0B2D55; padding:5px 10px 5px 6px;
  background:#fff; border:1px solid #E1E8F0; border-radius:999px; }

/* ---- cards ---- */
div[class*="st-key-card"] { background:#fff; border:1px solid #E1E8F0; border-radius:16px; padding:22px 24px; box-shadow:0 1px 2px rgba(11,45,85,.04), 0 6px 18px rgba(11,45,85,.04); }
.muted { color:#63758B; font-size:14px; line-height:1.55; }
.h-sec { font-size:17px; font-weight:700; color:#0B2D55; margin:0 0 2px 0; }
.tool-head { display:flex; align-items:center; gap:14px; margin-bottom:6px; }
.tool-head .tn { font-size:22px; font-weight:800; color:#0B2D55; line-height:1.1; }
.tool-head .td { font-size:13px; color:#63758B; }
.tag { display:inline-block; font-size:12px; font-weight:600; color:#0B6A9A; background:#E8F4FB; border:1px solid #CBE6F6; padding:4px 10px; border-radius:999px; margin:0 6px 6px 0; }
.tag.g { color:#0C7A5A; background:#E6F8F1; border-color:#BDEBDB; }
.chip { display:inline-block; font-size:12px; font-weight:600; padding:4px 10px; border-radius:8px; margin:0 6px 6px 0; }
.chip.ok { color:#0C7A5A; background:#E6F8F1; border:1px solid #BDEBDB; }
.chip.no { color:#B42318; background:#FEF0EE; border:1px solid #F8CFC9; }
.note { border-left:4px solid #00A6B2; background:#EFFAFB; border-radius:8px; padding:11px 14px; color:#0B4F63; font-size:13.5px; line-height:1.5; }
.note.warn { border-color:#E8A317; background:#FFF8E6; color:#6B4A00; }

/* ---- hero (home) ---- */
.hero { border-radius:20px; padding:34px 38px; background:linear-gradient(135deg,#0B2D55 0%,#0A4F8F 60%,#00859A 100%); position:relative; overflow:hidden; }
.hero::after { content:""; position:absolute; right:-60px; top:-60px; width:260px; height:260px; border-radius:50%; background:radial-gradient(circle,rgba(25,201,149,.35),transparent 70%); }
.hero h1 { color:#fff; font-size:36px; font-weight:800; margin:10px 0 8px 0; letter-spacing:-.02em; }
.hero p { color:#CFE0F3; font-size:16px; max-width:680px; line-height:1.6; margin:0; }
.hero .tl { display:inline-block; margin-top:14px; font-size:12.5px; font-weight:700; letter-spacing:.14em; text-transform:uppercase; color:#7EE6C5; }

/* ---- stepper ---- */
.stepper { display:flex; align-items:center; gap:0; margin: 4px 0 20px 0; }
.stp { display:flex; align-items:center; gap:10px; flex:1; }
.stp .c { width:30px; height:30px; flex:0 0 30px; border-radius:50%; display:flex; align-items:center; justify-content:center; font-size:13px; font-weight:800;
  background:#fff; border:2px solid #CFDAE8; color:#63758B; }
.stp .l { font-size:13px; font-weight:600; color:#63758B; line-height:1.25; }
.stp .l small { display:block; font-weight:500; font-size:11.5px; color:#8A9BB0; }
.stp .bar { flex:1; height:2px; background:#DCE5F0; margin:0 12px; min-width:18px; }
.stp.done .c { background:#19C995; border-color:#19C995; color:#fff; }
.stp.done .bar { background:#19C995; }
.stp.done .l { color:#0B2D55; }
.stp.active .c { background:linear-gradient(135deg,#0879D9,#00A6B2); border-color:transparent; color:#fff; box-shadow:0 0 0 4px rgba(8,121,217,.15); }
.stp.active .l { color:#0B2D55; font-weight:700; }
.stp:last-child { flex:0 0 auto; }
.stp:last-child .bar { display:none; }

/* ---- KPIs ---- */
.kpis { display:grid; grid-template-columns:repeat(4,1fr); gap:12px; margin:6px 0 14px 0; }
.kpi { border-radius:14px; padding:14px 16px; background:#F5F8FC; border:1px solid #E1E8F0; }
.kpi .v { font-size:26px; font-weight:800; color:#0B2D55; line-height:1.1; }
.kpi .l { font-size:11.5px; letter-spacing:.08em; text-transform:uppercase; color:#63758B; font-weight:600; margin-top:4px; }
.kpi.warn .v { color:#B7791F; }

/* ---- buttons / inputs ---- */
.stButton > button[kind="primary"], .stDownloadButton > button[kind="primary"] { background:#0879D9; border:0; color:#fff; font-weight:700; border-radius:10px; padding:.6rem 1.1rem; }
.stButton > button[kind="primary"]:hover, .stDownloadButton > button[kind="primary"]:hover { background:#0668BD; }
.stButton > button[kind="secondary"], .stDownloadButton > button[kind="secondary"] { border-radius:10px; font-weight:600; border:1px solid #CFDAE8; background:#fff; color:#0B2D55; }
.stButton > button[kind="secondary"]:hover, .stDownloadButton > button[kind="secondary"]:hover { border-color:#0879D9; color:#0879D9; }
[data-testid="stFileUploaderDropzone"] { border:2px dashed #9CC4EA; border-radius:16px; background:#F3F8FE; padding:34px 20px; }
[data-testid="stFileUploaderDropzone"] button { background:#0879D9; color:#fff; border:0; border-radius:10px; font-weight:700; }
[data-baseweb="input"] > div, [data-baseweb="select"] > div, [data-baseweb="textarea"] { border-radius:10px !important; border-color:#E1E8F0 !important; background:#fff !important; }
.stTabs [data-baseweb="tab"] { font-weight:600; color:#63758B; } .stTabs [aria-selected="true"] { color:#0879D9; }
.app-footer { text-align:center; margin-top:34px; padding:18px 8px 4px 8px; border-top:1px solid #E1E8F0; color:#63758B; font-size:12.5px; }
.app-footer b { color:#0B2D55; letter-spacing:.02em; }
@media (max-width:900px){ .kpis{grid-template-columns:repeat(2,1fr);} .stp .l small{display:none;} }
</style>
"""


def inject_css() -> None:
    st.markdown(SHARED_CSS, unsafe_allow_html=True)
