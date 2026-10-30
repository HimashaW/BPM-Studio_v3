"""TestX - Test Case Generator (the BPM Studio page).

Four steps: upload the product workbook, review the rules TestX found, choose settings,
download the Excel test plan. The thinking is in engine.py / rules.py / plan.py (no Streamlit there).
"""
from __future__ import annotations

import dataclasses
import hashlib
import html
import re
from math import prod
from pathlib import Path

import streamlit as st

from studio.state import S, set_state, keep_alive, human_size, log_activity
from studio.ui import stepper
from tools.testx import style
from tools.testx.engine import TooManyCombinations
from tools.testx.export import build_workbook
from tools.testx.icon import testx_icon
from tools.testx.mascot import tess
from tools.testx.model import BLANK, BLANK_LABEL, Cond
from tools.testx.plan import Settings, generate
from tools.testx.rules import Catalog, detect_rules, manual_rule, _mandatory_flag
from tools.testx.source import read_source

TOOL = "TestX"
TOOL_DESC = "Test Case Generator"

STEPS = [("Upload Source File", "Product workbook"), ("Review Rules", "What TestX found"),
         ("Choose Settings", "Technique and scope"), ("Download Test Plan", "Excel for UAT")]

PERSIST_KEYS = ("tx_product", "tx_prepared", "tx_technique", "tx_scope", "tx_components", "tx_blank", "tx_neg",
                "tx_order", "tx_limit")
MANUAL_KEYS = ("tx_m_wa", "tx_m_wop", "tx_m_wv", "tx_m_ta", "tx_m_top", "tx_m_tv")


# ------------------------------------------------------------------ state
def init_state() -> None:
    S.setdefault("tx_step", 1)
    S.setdefault("tx_file", None)            # {"name","size","bytes","id"}
    S.setdefault("tx_up_n", 0)               # bump to reset the uploader widget
    S.setdefault("tx_replacing", False)
    S.setdefault("tx_parsed", None)          # {"id","src","cat","rules","text_only"}
    S.setdefault("tx_manual", [])            # rules the user added
    S.setdefault("tx_enabled", {})           # rule id -> bool (the "Use" tick boxes)
    S.setdefault("tx_ed_key", 0)             # bump to give the rules table a fresh state
    S.setdefault("tx_table_base", None)
    S.setdefault("tx_result", None)
    keep_alive(PERSIST_KEYS)


def reset_editor() -> None:
    S["tx_ed_key"] += 1
    S["tx_table_base"] = None


def enter_step(step: int) -> None:
    """Button callback: go to a step; entering the review step rebuilds the rules table."""
    S["tx_step"] = step
    if step == 2:
        reset_editor()


def clear_file() -> None:
    for k in ("tx_file", "tx_parsed", "tx_result", "tx_table_base"):
        S[k] = None
    S["tx_step"] = 1
    S["tx_replacing"] = False
    S["tx_up_n"] += 1
    S["tx_manual"], S["tx_enabled"] = [], {}
    for k in ("tx_product", "tx_components") + MANUAL_KEYS:
        S.pop(k, None)


def all_rules():
    """Detected + manual rules with the user's tick boxes applied."""
    p = S["tx_parsed"]
    out = []
    for r in list(p["rules"]) + list(S["tx_manual"]):
        on = S["tx_enabled"].get(r.id, r.enabled)
        out.append(dataclasses.replace(r, enabled=bool(on) and not r.issues))
    return out


def ensure_parsed():
    f = S["tx_file"]
    p = S["tx_parsed"]
    if p and p["id"] == f["id"]:
        return p
    with st.spinner("Tess is reading the workbook…"):
        src = read_source(f["bytes"])
        cat = Catalog.from_nodes(src.nodes)
        rules, text_only = detect_rules(src.blocks, cat)
    S["tx_parsed"] = {"id": f["id"], "src": src, "cat": cat, "rules": rules, "text_only": text_only}
    S["tx_manual"] = []
    S["tx_enabled"] = {r.id: r.enabled for r in rules}
    S["tx_result"] = None
    for k in ("tx_product", "tx_components") + MANUAL_KEYS:
        S.pop(k, None)
    reset_editor()
    return S["tx_parsed"]


# ------------------------------------------------------------------ small html helpers
def tool_head() -> str:
    return (f'<div class="tool-head">{testx_icon(46)}<div><div class="tn">{TOOL}</div>'
            f'<div class="td">{TOOL_DESC}</div></div></div>')


def say(mood: str, text: str, size: int = 112) -> None:
    st.markdown(f'<div class="tx-say"><div class="tx-art tx-bob">{tess(mood, size)}</div>'
                f'<div class="tx-bubble"><span class="who">Tess</span>{text}</div></div>', unsafe_allow_html=True)


def safe_name(text: str) -> str:
    return re.sub(r"[^A-Za-z0-9_-]+", "_", text).strip("_") or "Product"


def fmt(n: int) -> str:
    return f"{n:,}"


def kpis(items) -> None:
    cells = "".join(f'<div class="{"kpi warn" if warn else "kpi"}"><div class="v">{v}</div><div class="l">{l}</div></div>'
                    for v, l, warn in items)
    st.markdown(f'<div class="kpis">{cells}</div>', unsafe_allow_html=True)


# ------------------------------------------------------------------ step 1
def step1() -> None:
    f = S["tx_file"]
    st.markdown(f'''<div class="tx-hero"><div class="tx-copy">{testx_icon(52)}
<h2>Find the bugs before your customers do</h2>
<p>TestX reads your product configuration workbook, works out which combinations of attributes really matter,
respects your business rules, and hands you an Excel test plan ready for UAT.</p>
<span class="tx-pill">Pairwise coverage</span><span class="tx-pill">Business rules built in</span><span class="tx-pill">Negative tests</span><span class="tx-pill">Excel for UAT</span></div>
<div class="tx-art tx-bob">{tess("hello", 190)}</div></div>''', unsafe_allow_html=True)
    with st.container(key="card_tx_upload"):
        st.markdown(tool_head(), unsafe_allow_html=True)
        st.markdown('<div class="muted" style="margin-bottom:12px">Upload the same product configuration workbook you use for GuideX. '
                    'TestX uses the product hierarchy (attributes and values), the special requirements (rules) '
                    'and, optionally, the order scenarios.</div>', unsafe_allow_html=True)
        if f and not S["tx_replacing"]:
            c1, c2, c3 = st.columns([5, 1.1, 1.1], vertical_alignment="center")
            with c1:
                st.markdown(f'<div style="display:flex;align-items:center;gap:12px"><div style="font-size:30px">📗</div>'
                            f'<div><b style="color:#0B2D55">{html.escape(f["name"])}</b><div class="muted" style="font-size:12.5px">'
                            f'{human_size(f["size"])} · Excel workbook · <span style="color:#0C7A5A;font-weight:600">Uploaded</span></div></div></div>',
                            unsafe_allow_html=True)
            with c2:
                if st.button("Replace", key="tx_replace_file", width="stretch"):
                    S["tx_replacing"] = True
                    st.rerun()
            with c3:
                st.button("Remove", key="tx_remove_file", width="stretch", on_click=clear_file)
        else:
            up = st.file_uploader("Upload your product configuration workbook.", type=["xlsx"],
                                  key=f"tx_uploader_{S['tx_up_n']}", help="Excel .xlsx · up to 25 MB.")
            if up is not None:
                data = up.getvalue()
                fid = hashlib.md5(data).hexdigest()
                S["tx_file"] = {"name": up.name, "size": len(data), "bytes": data, "id": fid}
                S["tx_replacing"] = False
                st.rerun()
            if S["tx_replacing"]:
                st.button("Cancel replace", key="tx_cancel_replace", on_click=lambda: S.update(tx_replacing=False))
        st.markdown('<div class="note" style="margin-top:14px"><b>Your workbook stays the source of truth.</b> TestX only combines '
                    'attributes and values that are written in it, and only enforces rules it can read; anything else is '
                    'listed for you to decide.</div>', unsafe_allow_html=True)
    if S["tx_file"] and not S["tx_replacing"]:
        st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)
        _, cn = st.columns([4, 1.4])
        with cn:
            st.button("Continue", key="tx_s1_next", type="primary", icon=":material/arrow_forward:", icon_position="right",
                      width="stretch", on_click=enter_step, kwargs={"step": 2})


# ------------------------------------------------------------------ step 2
def _component_rows(src):
    rows = []
    for n in src.nodes:
        withv = [a for a in n.attributes if a.values]
        sizes = [len(a.values) + (1 if _mandatory_flag(a.mandatory) is False else 0) for a in withv]
        rows.append({"Component": n.name, "Type": n.kind, "Attributes": len(n.attributes),
                     "With a value list": len(withv), "Values": sum(len(a.values) for a in withv),
                     "All combinations": fmt(prod(sizes)) if sizes else "-"})
    return rows


def _rules_table(rules):
    base = []
    for r in rules:
        base.append({"Use": bool(S["tx_enabled"].get(r.id, r.enabled)) and not r.issues, "Rule": r.id,
                     "From": {"table": "Workbook table", "sentence": "Workbook sentence", "manual": "Added by you"}[r.source],
                     "What it says": r.text, "Status": r.status, "Notes": "; ".join(r.issues + r.notes)})
    return base


def add_manual() -> None:
    cat = S["tx_parsed"]["cat"]
    wa, ta = S.get("tx_m_wa"), S.get("tx_m_ta")
    wv, tv = S.get("tx_m_wv") or [], S.get("tx_m_tv") or []
    if not (wa and ta and wv and tv):
        S["tx_m_msg"] = "Pick both attributes and at least one value on each side."
        return
    conv = lambda vals: frozenset(BLANK if v == BLANK_LABEL else v for v in vals)
    idx = len(S["tx_manual"]) + 1
    existing = {r.id for r in S["tx_manual"]}
    while f"M{idx:02d}" in existing:
        idx += 1
    rule = manual_rule(idx, Cond(wa, conv(wv), negate=S.get("tx_m_wop") == "is not"),
                       Cond(ta, conv(tv), negate=S.get("tx_m_top") == "must not be"))
    S["tx_manual"].append(rule)
    S["tx_enabled"][rule.id] = True
    S["tx_m_msg"] = f"Added {rule.id}."
    for k in MANUAL_KEYS:
        S.pop(k, None)
    reset_editor()


def remove_manual(rid: str) -> None:
    S["tx_manual"] = [r for r in S["tx_manual"] if r.id != rid]
    S["tx_enabled"].pop(rid, None)
    reset_editor()


def step2() -> None:
    p = ensure_parsed()
    src, cat = p["src"], p["cat"]
    usable = [n for n in src.nodes if any(a.values for a in n.attributes)]
    if not usable:
        say("oops", "I couldn't find any attribute with a list of values. TestX needs the "
                    "<b>2. Product Hierarchy</b> sheet with <i>Attribute</i> and <i>Value</i> rows.")
    else:
        say("think", "Here is what I found. I read the rules from your workbook; anything I wasn't sure about "
                     "is flagged so <b>you</b> decide what gets used.")
    with st.container(key="card_tx_review"):
        st.markdown('<div class="h-sec">Review what TestX found</div>', unsafe_allow_html=True)
        st.markdown('<div class="muted" style="margin-bottom:10px">Check the attributes, then decide which rules to apply.</div>',
                    unsafe_allow_html=True)
        chips = [("2.", "Product Hierarchy", bool(src.nodes)), ("3.", "Special Requirements", bool(src.blocks)),
                 ("5.1", "Order Scenarios", bool(src.order_types))]
        st.markdown("".join(f'<span class="chip {"ok" if ok else "no"}">{"✓" if ok else "✕"} {a} {b}</span>' for a, b, ok in chips),
                    unsafe_allow_html=True)
        rules = all_rules()
        ready = sum(1 for r in rules if r.enabled)
        review = sum(1 for r in rules if r.issues)
        kpis([(len(src.nodes), "Components", False), (src.attribute_count, "Attributes", False),
              (src.value_count, "Values", False), (f"{ready}", "Rules in use", False)]
             if not review else
             [(len(src.nodes), "Components", False), (src.attribute_count, "Attributes", False),
              (f"{ready}", "Rules in use", False), (review, "Rules to review", True)])
        if usable:
            st.markdown("**Components**")
            st.dataframe(_component_rows(src), hide_index=True, width="stretch")
            st.caption("“All combinations” is what you would have to test without a technique. TestX will cut it down.")
        for w in src.warnings:
            st.warning(w)

        st.markdown("**Business rules**")
        if rules:
            if S["tx_table_base"] is None:
                S["tx_table_base"] = _rules_table(rules)
            edited = st.data_editor(
                S["tx_table_base"], key=f"tx_rules_editor_{S['tx_ed_key']}", hide_index=True, width="stretch",
                disabled=["Rule", "From", "What it says", "Status", "Notes"],
                column_config={"Use": st.column_config.CheckboxColumn("Use", width="small"),
                               "Rule": st.column_config.TextColumn("Rule", width="small"),
                               "From": st.column_config.TextColumn("From", width="small"),
                               "What it says": st.column_config.TextColumn("What it says", width="large"),
                               "Status": st.column_config.TextColumn("Status", width="small"),
                               "Notes": st.column_config.TextColumn("Notes", width="medium")})
            records = edited if isinstance(edited, list) else edited.to_dict("records")
            by_id = {r.id: r for r in rules}
            S["tx_enabled"] = {row["Rule"]: bool(row["Use"]) and not by_id[row["Rule"]].issues for row in records}
            st.markdown('<div class="tx-rule-note">Tick “Use” to include a rule. Rules marked <b>Needs review</b> mention '
                        'something the workbook does not list (or an attribute without a value list), so they stay off until the '
                        'workbook is fixed.</div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="muted">No machine-readable rules were found. You can still generate a plan - or add rules below.</div>',
                        unsafe_allow_html=True)

        if p["text_only"]:
            with st.expander(f"Requirements TestX could not turn into rules · {len(p['text_only'])}"):
                for t in p["text_only"]:
                    st.markdown(f"- {html.escape(t.text[:260])}  \n  <span class='tx-rule-note'>{html.escape(t.reason)}</span>",
                                unsafe_allow_html=True)

        with st.expander("Add your own rule"):
            names = [i.name for i in cat.attrs.values() if i.values]
            if names:
                st.markdown('<div class="muted">Example: <b>If</b> CPE Type <b>is</b> SonicWall <b>then</b> CPE Model <b>must not be</b> FortiF.</div>',
                            unsafe_allow_html=True)
                c1, c2, c3 = st.columns([2, 1, 3])
                c1.selectbox("If attribute", names, index=None, key="tx_m_wa", placeholder="Choose…")
                c2.selectbox("is / is not", ["is", "is not"], key="tx_m_wop")
                wopts = (cat.get(S["tx_m_wa"]).values if S.get("tx_m_wa") else [])
                c3.multiselect("Value(s)", wopts, key="tx_m_wv")
                d1, d2, d3 = st.columns([2, 1, 3])
                d1.selectbox("Then attribute", names, index=None, key="tx_m_ta", placeholder="Choose…")
                d2.selectbox("must be / must not be", ["must be", "must not be"], key="tx_m_top")
                topts = (cat.get(S["tx_m_ta"]).values + [BLANK_LABEL]) if S.get("tx_m_ta") else []
                d3.multiselect("Value(s) ", topts, key="tx_m_tv")
                st.button("Add rule", key="tx_m_add", icon=":material/add:", on_click=add_manual)
                if S.get("tx_m_msg"):
                    st.caption(S.pop("tx_m_msg"))
            for r in S["tx_manual"]:
                a, b = st.columns([6, 1], vertical_alignment="center")
                a.markdown(f"**{r.id}** · {html.escape(r.text)}", unsafe_allow_html=True)
                b.button("Remove", key=f"tx_rm_{r.id}", on_click=remove_manual, kwargs={"rid": r.id})
    st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)
    b1, _, b2 = st.columns([1.2, 3, 1.6])
    with b1:
        st.button("Back", key="tx_s2_back", icon=":material/arrow_back:", width="stretch",
                  on_click=enter_step, kwargs={"step": 1})
    with b2:
        st.button("Continue", key="tx_s2_next", type="primary", disabled=not usable, icon=":material/arrow_forward:",
                  icon_position="right", width="stretch", on_click=enter_step, kwargs={"step": 3})


# ------------------------------------------------------------------ step 3
def step3() -> None:
    f, p = S["tx_file"], ensure_parsed()
    src = p["src"]
    names = [n.name for n in src.nodes if any(a.values for a in n.attributes)]
    S.setdefault("tx_product", src.product_name or Path(f["name"]).stem)
    S.setdefault("tx_prepared", "")
    S.setdefault("tx_technique", "pairwise")
    S.setdefault("tx_scope", "component")
    S.setdefault("tx_components", list(names))
    S.setdefault("tx_blank", True)
    S.setdefault("tx_neg", True)
    S.setdefault("tx_order", False)
    S.setdefault("tx_limit", 5000)
    S["tx_components"] = [c for c in S["tx_components"] if c in names]
    say("work", "Almost there. <b>Pairwise</b> is my favourite: every pair of values is tested together at least once, "
                "with far fewer rows than testing every combination.")
    with st.container(key="card_tx_config"):
        st.markdown('<div class="h-sec">Choose settings</div>', unsafe_allow_html=True)
        st.markdown('<div class="muted" style="margin-bottom:12px">These choices decide how many tests you get and how they are laid out.</div>',
                    unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        with c1:
            st.text_input("Product / service name *", key="tx_product")
        with c2:
            st.text_input("Prepared by", key="tx_prepared", placeholder="optional - appears on the Summary sheet")
        t1, t2 = st.columns(2)
        with t1:
            st.radio("Technique", ["pairwise", "full"], key="tx_technique",
                     format_func=lambda x: "Pairwise (recommended)" if x == "pairwise" else "All valid combinations",
                     help="Pairwise: every pair of attribute values appears in at least one test. "
                          "All valid combinations: every combination the rules allow - can be very large.")
        with t2:
            st.radio("Layout", ["component", "together"], key="tx_scope",
                     format_func=lambda x: "One sheet per component" if x == "component" else "All components together",
                     help="Together lets rules that link attributes of different components be tested.")
        st.multiselect("Components to include", names, key="tx_components")
        st.markdown("**Extras**")
        st.toggle("Let optional attributes stay blank in some tests", key="tx_blank",
                  help="Attributes marked Mandatory? = No get a blank option.")
        st.toggle("Add one “must be refused” test for every rule", key="tx_neg",
                  help="Each negative test picks a value the rule forbids; the expected result says it must be refused.")
        st.toggle("Also vary the Order Type (from sheet 5.1)", key="tx_order", disabled=not src.order_types,
                  help="Adds Order Type as one more attribute: " + (", ".join(src.order_types) if src.order_types else "no order scenarios found"))
        if S["tx_technique"] == "full":
            st.number_input("Largest number of tests allowed", min_value=100, max_value=100000, step=500, key="tx_limit",
                            help="TestX stops and tells you if the rules still allow more combinations than this.")
        n_rules = sum(1 for r in all_rules() if r.enabled)
        st.markdown(f'<div class="tx-rule-note">{n_rules} rule(s) will be applied.</div>', unsafe_allow_html=True)
    st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)
    b1, _, b2 = st.columns([1.2, 3, 2.2])
    ready = bool(S["tx_product"].strip()) and bool(S["tx_components"])
    with b1:
        st.button("Back", key="tx_s3_back", icon=":material/arrow_back:", width="stretch",
                  on_click=enter_step, kwargs={"step": 2})
    with b2:
        go_gen = st.button("Generate Test Plan", key="tx_s3_gen", type="primary", disabled=not ready,
                           icon=":material/auto_awesome:", width="stretch")
    if not S["tx_product"].strip():
        st.info("Enter the product / service name to continue.")
    if not S["tx_components"]:
        st.info("Choose at least one component.")
    if go_gen and ready:
        settings = Settings(technique=S["tx_technique"], scope=S["tx_scope"], components=list(S["tx_components"]),
                            include_blank=S["tx_blank"], negatives=S["tx_neg"],
                            vary_order_type=bool(S["tx_order"] and src.order_types), full_limit=int(S["tx_limit"]))
        product = S["tx_product"].strip()
        with st.status("Tess is building your test plan…", expanded=True) as status:
            try:
                status.write("Combining attributes and applying the rules …")
                plan = generate(src, all_rules(), settings, product)
                status.write("Writing the Excel test plan …")
                xlsx = build_workbook(plan, f["name"], S["tx_prepared"].strip())
                S["tx_result"] = {"plan": plan, "xlsx": xlsx, "name": f"{safe_name(product)}_TestX_Test_Plan.xlsx", "file": f["id"]}
                total = plan.positive + plan.negative
                log_activity(TOOL, product, f"test plan · {total} test cases")
                status.update(label="Done — your test plan is ready", state="complete", expanded=False)
                S["tx_step"] = 4
                st.rerun()
            except TooManyCombinations as exc:
                status.update(label="Too many combinations", state="error")
                st.error(f"The rules still allow more than {exc.limit:,} combinations. Use Pairwise, include fewer components, "
                         "or raise the limit.")
            except Exception as exc:                                         # noqa: BLE001
                status.update(label="Something went wrong", state="error")
                st.error("Test plan generation failed. Please check the workbook and try again.")
                with st.expander("Technical details"):
                    st.code(f"{type(exc).__name__}: {exc}")


# ------------------------------------------------------------------ step 4
def step4() -> None:
    res = S["tx_result"]
    if not res:
        st.info("Nothing generated yet.")
        st.button("Back to settings", on_click=enter_step, kwargs={"step": 3})
        return
    plan = res["plan"]
    total = plan.positive + plan.negative
    if not plan.groups or not total:
        say("oops", "I couldn't build any tests - none of the chosen components has an attribute with a value list.")
    else:
        say("cheer", f"Done! <b>{fmt(total)}</b> test case(s) are ready for UAT. "
                     "Open the Excel file and start ticking off Pass / Fail.")
    with st.container(key="card_tx_results"):
        st.markdown(f'<div class="h-sec">Your test plan · {html.escape(plan.product)}</div>', unsafe_allow_html=True)
        full = plan.full_total
        cut = f"{max(0.0, (1 - plan.positive / full) * 100):.1f}%" if full and plan.positive else "-"
        kpis([(fmt(total), "Test cases", False), (f"{fmt(plan.positive)} + {fmt(plan.negative)}", "Positive + negative", False),
              (fmt(full), "All combinations", False), (cut, "Fewer positive tests", False)])
        d1, d2 = st.columns([1.4, 3])
        with d1:
            st.download_button("Test plan (.xlsx)", res["xlsx"], res["name"],
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary",
                               icon=":material/download:", width="stretch")
        notes = list(plan.warnings)
        for g in plan.groups:
            notes += [f"{g.name}: {n}" for n in g.notes]
        if notes:
            with st.expander(f"{len(notes)} note(s) about this plan", expanded=False):
                for n in notes:
                    st.markdown(f"- {n}")
        else:
            st.success("No notes - every enabled rule could be applied.")
    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
    with st.container(key="card_tx_preview"):
        st.markdown('<div class="h-sec">Preview</div>', unsafe_allow_html=True)
        if plan.groups:
            for tab, g in zip(st.tabs([g.name for g in plan.groups]), plan.groups):
                with tab:
                    keys = [f.key for f in g.factors]
                    rows = [{"Test ID": c.id, "Type": c.kind,
                             **{k: (BLANK_LABEL if c.values.get(k) is BLANK else c.values.get(k)) for k in keys},
                             "Rules": ", ".join(c.rules)} for c in g.cases[:300]]
                    st.dataframe(rows, hide_index=True, width="stretch")
                    if len(g.cases) > 300:
                        st.caption(f"Showing the first 300 of {len(g.cases)} - the Excel file has them all.")
        else:
            st.info("Nothing to preview.")
    st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)
    b1, _, b2 = st.columns([1.6, 3, 1.8])
    with b1:
        st.button("Back to settings", key="tx_s4_back", icon=":material/arrow_back:", width="stretch",
                  on_click=enter_step, kwargs={"step": 3})
    with b2:
        st.button("Start a new plan", key="tx_s4_new", width="stretch", icon=":material/refresh:", on_click=clear_file)


def render() -> None:
    init_state()
    style.inject()
    if S["tx_step"] > 1 and not S["tx_file"]:
        S["tx_step"] = 1
    stepper(S["tx_step"], STEPS)
    {1: step1, 2: step2, 3: step3, 4: step4}[S["tx_step"]]()
