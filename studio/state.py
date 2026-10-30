"""Small helpers every tool uses. Keeps tools independent from app.py."""
from __future__ import annotations

import streamlit as st

S = st.session_state


def go_page(page: str) -> None:
    """Switch the platform to another page ("home" or a tool id). Used as a button callback."""
    S["page"] = page


def set_state(**values) -> None:
    """Button callback that sets session-state values, e.g. on_click=set_state, kwargs={"tx_step": 2}."""
    for k, v in values.items():
        S[k] = v


def keep_alive(keys) -> None:
    """Streamlit forgets a widget's value when the widget is not drawn on a rerun.
    Re-assigning the value keeps form fields alive while the user moves between steps."""
    for k in keys:
        if k in S:
            S[k] = S[k]


def human_size(n: int) -> str:
    return f"{n/1024:.0f} KB" if n < 1024 * 1024 else f"{n/1024/1024:.1f} MB"


def log_activity(tool: str, label: str, detail: str) -> None:
    """Record something a tool produced so the Home page can list it (this browser session only)."""
    import datetime
    S.setdefault("history", []).append(
        {"tool": tool, "label": label, "detail": detail, "time": datetime.datetime.now().strftime("%H:%M")})
