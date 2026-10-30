"""Styling that only the TestX page needs (kawaii hero, speech bubble). Shared styling is in studio/theme.py."""
from __future__ import annotations

import streamlit as st

CSS = """
<style>
.tx-hero { position:relative; overflow:hidden; display:flex; align-items:center; justify-content:space-between; gap:18px;
  border-radius:20px; padding:22px 30px 18px 30px; margin-bottom:14px;
  background:linear-gradient(135deg,#E6F8F1 0%,#E3F0FD 55%,#EEF3FF 100%); border:1px solid #CBE9DD; }
.tx-hero::before { content:"✦"; position:absolute; left:46%; top:14px; font-size:20px; color:#FFD86B; }
.tx-hero::after { content:"✧"; position:absolute; left:58%; bottom:16px; font-size:24px; color:#8FD3F0; }
.tx-hero .tx-copy { max-width:520px; z-index:1; }
.tx-hero h2 { font-size:30px; font-weight:800; color:#0B2D55; margin:8px 0 4px 0; letter-spacing:-.02em; }
.tx-hero p { color:#2D4C70; font-size:15px; line-height:1.55; margin:0 0 10px 0; }
.tx-hero .tx-art { flex:0 0 auto; z-index:1; }
.tx-bob { animation: tx-bob 3.4s ease-in-out infinite; transform-origin:50% 100%; }
@keyframes tx-bob { 0%,100% { transform:translateY(0); } 50% { transform:translateY(-5px); } }
@media (prefers-reduced-motion: reduce) { .tx-bob { animation:none; } }

.tx-say { display:flex; align-items:center; gap:6px; margin:0 0 12px 0; }
.tx-say .tx-art { flex:0 0 auto; }
.tx-bubble { position:relative; background:#fff; border:1.5px solid #CDE3F5; border-radius:18px; padding:11px 16px; max-width:640px;
  color:#0B2D55; font-size:14px; line-height:1.5; box-shadow:0 4px 14px rgba(11,45,85,.06); }
.tx-bubble::before { content:""; position:absolute; left:-9px; top:50%; width:15px; height:15px; background:#fff;
  border-left:1.5px solid #CDE3F5; border-bottom:1.5px solid #CDE3F5; transform:translateY(-50%) rotate(45deg); }
.tx-bubble .who { display:block; font-size:12px; font-weight:800; color:#00A6B2; margin-bottom:2px; }
.tx-pill { display:inline-block; font-size:12px; font-weight:700; color:#0C7A5A; background:#fff; border:1px solid #BDEBDB;
  padding:3px 10px; border-radius:999px; margin:0 6px 4px 0; }
.tx-rule-note { font-size:12.5px; color:#63758B; }
.tx-bugline { display:flex; align-items:center; gap:10px; font-size:13px; color:#63758B; }
@media (max-width:900px){ .tx-hero { flex-direction:column; align-items:flex-start; } }
</style>
"""


def inject() -> None:
    st.markdown(CSS, unsafe_allow_html=True)
