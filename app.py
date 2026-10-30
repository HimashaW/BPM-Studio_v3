"""BPM Studio - internal BPM automation platform.

This file is only the platform shell: page setup, sidebar, header, Home page and the router.
Each tool lives in its own folder under tools/ and is listed in tools/__init__.py, so adding a
tool never means editing this file.

Run locally:   streamlit run app.py
Deploy:        see README.md ("Hosting").
"""
from __future__ import annotations

import streamlit as st

from studio.state import S, go_page
from studio.theme import PLATFORM, PLATFORM_TAGLINE, VERSION, bpm_logo, inject_css
from tools import TOOLS

TOOL_BY_ID = {t.id: t for t in TOOLS}

S.setdefault("page", "home")
if S["page"] not in TOOL_BY_ID:
    S["page"] = "home"
current = TOOL_BY_ID.get(S["page"])

st.set_page_config(page_title=f"{PLATFORM} · {current.name}" if current else PLATFORM,
                   page_icon=":material/hexagon:", layout="wide", initial_sidebar_state="expanded")
inject_css()
S.setdefault("history", [])            # this session's generated files (real data only)

# ----------------------------------------------------------------- sidebar --
with st.sidebar:
    st.markdown(f'<div class="sb-brand">{bpm_logo(38)}<div><div class="n">{PLATFORM}</div>'
                f'<div class="t">{PLATFORM_TAGLINE}</div></div></div>', unsafe_allow_html=True)
    st.markdown("---")
    st.markdown('<div class="sb-label">Workspace</div>', unsafe_allow_html=True)
    st.button("Home", key="nav_home", icon=":material/home:", width="stretch",
              type="primary" if S["page"] == "home" else "tertiary", on_click=go_page, kwargs={"page": "home"})
    for t in TOOLS:
        st.button(t.name, key=f"nav_{t.id}", icon=t.nav_icon, width="stretch",
                  type="primary" if S["page"] == t.id else "tertiary", on_click=go_page, kwargs={"page": t.id})
    st.markdown("---")
    st.markdown('<div class="sb-label">More Tools</div>', unsafe_allow_html=True)
    st.markdown('<div class="sb-soon">Additional BPM tools will appear here as they are released.</div>', unsafe_allow_html=True)
    st.markdown("---")
    st.markdown(f'<div class="sb-foot">Uploaded files are processed in a temporary folder and discarded after each run.'
                f'<br><br>{PLATFORM} · {VERSION}</div>', unsafe_allow_html=True)

# ------------------------------------------------------------------ header --
h1, h2 = st.columns([2.8, 2.6], vertical_alignment="center")
with h1:
    crumb = f'<b>{PLATFORM}</b>' if not current else f'{PLATFORM} &nbsp;›&nbsp; <b>{current.name}</b>'
    st.markdown(f'<div class="crumb">{crumb}</div><div class="page-title">{current.name if current else "Home"}</div>',
                unsafe_allow_html=True)
with h2:
    r1, r2 = st.columns([1.35, 1], vertical_alignment="center")
    with r1:
        st.markdown(f'<div class="id-chip">{bpm_logo(20)} {PLATFORM}</div>', unsafe_allow_html=True)
    with r2:
        with st.popover("Help", icon=":material/help:", width="stretch"):
            if not current:
                names = " and ".join(f"**{t.name}**" for t in TOOLS)
                st.markdown(f"**{PLATFORM}** brings BPM automation tools together in one place. Open {names} from the "
                            "sidebar or the cards below.")
            else:
                st.markdown(current.help_md)
st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)


# ---------------------------------------------------------------- HOME page --
def page_home():
    st.markdown(f"""
<div class="hero">{bpm_logo(54)}
<h1>Welcome to {PLATFORM}</h1>
<p>Transform business information into structured processes and practical documentation with a growing collection of BPM automation tools.</p>
<span class="tl">{PLATFORM_TAGLINE}</span></div>""", unsafe_allow_html=True)
    st.markdown("<div style='height:18px'></div>", unsafe_allow_html=True)
    st.markdown('<div class="h-sec">Available tools</div>', unsafe_allow_html=True)
    for t in TOOLS:
        with st.container(key=f"card_{t.id}"):
            c1, c2 = st.columns([4, 1.2], vertical_alignment="center")
            with c1:
                tags = "".join(f'<span class="tag{" g" if style else ""}">{label}</span>' for label, style in t.tags)
                st.markdown(f"""
<div class="tool-head">{t.icon_svg(52)}<div><div class="tn">{t.name}</div><div class="td">{t.description}</div></div></div>
<div class="muted" style="margin:6px 0 10px 0">{t.blurb}</div>{tags}""", unsafe_allow_html=True)
            with c2:
                st.button(f"Open {t.name}", key=f"open_{t.id}", type="primary", icon=":material/arrow_forward:",
                          width="stretch", on_click=go_page, kwargs={"page": t.id})
        st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
    st.markdown('<div class="h-sec">Recent activity</div>', unsafe_allow_html=True)
    with st.container(key="card_recent"):
        if S["history"]:
            for h in reversed(S["history"][-5:]):
                st.markdown(f"**{h['label']}** — {h['tool']} · {h['detail']} · {h['time']}")
            st.caption("Shown for this browser session only; files are not stored.")
        else:
            st.markdown('<div class="muted">Nothing has been generated in this session yet. '
                        'Open a tool to get started.</div>', unsafe_allow_html=True)


# ------------------------------------------------------------------ router --
if current:
    current.render()
else:
    page_home()

footer_name = f"{current.name} — {current.description}" if current else " · ".join(t.name for t in TOOLS)
st.markdown(f'<div class="app-footer"><b>{PLATFORM_TAGLINE}</b><br>{PLATFORM} · {footer_name} · {VERSION}</div>',
            unsafe_allow_html=True)
