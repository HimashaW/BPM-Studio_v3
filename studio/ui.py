"""Small UI pieces shared by tools that use the step-by-step layout."""
from __future__ import annotations

import streamlit as st


def stepper(active: int, steps) -> None:
    """steps = [(title, subtitle), ...]; `active` is 1-based."""
    html = ['<div class="stepper">']
    for i, (t, s) in enumerate(steps, start=1):
        cls = "done" if i < active else ("active" if i == active else "")
        mark = "✓" if i < active else str(i)
        html.append(f'<div class="stp {cls}"><div class="c">{mark}</div><div class="l">{t}<small>{s}</small></div>'
                    f'<div class="bar"></div></div>')
    html.append("</div>")
    st.markdown("".join(html), unsafe_allow_html=True)
