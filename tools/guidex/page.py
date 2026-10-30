"""GuideX - Service Delivery Guide Generator (the BPM Studio page).

The generation logic is NOT here: it lives in app_service.py / pipeline_core.py and is
unchanged. This file only draws the four steps (upload, review, configure, download).
"""
from __future__ import annotations

import datetime
import hashlib

import streamlit as st

import app_service as svc
from render.task_flow import load_flow_rules  # noqa: F401  (kept: warms the flow rules import)
from studio.state import S, set_state, keep_alive, human_size, log_activity
from tools.guidex.icon import guidex_icon

TOOL = "GuideX"
TOOL_DESC = "Service Delivery Guide Generator"

# form values that must survive while their widgets are not on screen (other steps)
PERSIST_KEYS = ("gx_product", "gx_doc_no", "gx_issue_no", "gx_issued_by", "gx_rev_no", "gx_rev_date", "gx_pptx",
                "gx_c_name", "gx_c_des", "gx_c_tel", "gx_c_mob", "gx_c_ext", "gx_c_mail", "gx_use_ai", "gx_ai_key",
                "gx_sheet")


def init_state() -> None:
    S.setdefault("gx_step", 1)
    S.setdefault("gx_file", None)          # {"name","size","bytes","id"}
    S.setdefault("gx_up_n", 0)             # bump to reset the uploader widget
    S.setdefault("gx_replacing", False)
    S.setdefault("gx_preview", None)       # {"key":..., "data":...}
    S.setdefault("result", None)
    keep_alive(PERSIST_KEYS)


def clear_file():
    S["gx_file"] = None
    S["gx_preview"] = None
    S["result"] = None
    S["gx_step"] = 1
    S["gx_replacing"] = False
    S["gx_up_n"] += 1
    for k in ("gx_product", "gx_sheet"):
        S.pop(k, None)


# --------------------------------------------------------------- GuideX page -
STEP_META = [("Upload Source File", "Product workbook"), ("Review Extracted Information", "What was read"),
             ("Configure and Generate", "Document details"), ("Review and Download", "Your outputs")]


def stepper(active: int):
    html = ['<div class="stepper">']
    for i, (t, s) in enumerate(STEP_META, start=1):
        cls = "done" if i < active else ("active" if i == active else "")
        mark = "✓" if i < active else str(i)
        html.append(f'<div class="stp {cls}"><div class="c">{mark}</div><div class="l">{t}<small>{s}</small></div>'
                    f'<div class="bar"></div></div>')
    html.append("</div>")
    st.markdown("".join(html), unsafe_allow_html=True)


def tool_head():
    st.markdown(f'<div class="tool-head">{guidex_icon(46)}<div><div class="tn">{TOOL}</div>'
                f'<div class="td">{TOOL_DESC}</div></div></div>', unsafe_allow_html=True)


def step1():
    with st.container(key="card_upload"):
        tool_head()
        st.markdown('<div class="muted" style="margin-bottom:12px">Upload your product configuration workbook. GuideX reads it '
                    'and builds the Service Delivery Guide, swimlane flowcharts and an editable PowerPoint.</div>',
                    unsafe_allow_html=True)
        f = S["gx_file"]
        if f and not S["gx_replacing"]:
            c1, c2, c3 = st.columns([5, 1.1, 1.1], vertical_alignment="center")
            with c1:
                st.markdown(f'<div style="display:flex;align-items:center;gap:12px"><div style="font-size:30px">📗</div>'
                            f'<div><b style="color:#0B2D55">{f["name"]}</b><div class="muted" style="font-size:12.5px">'
                            f'{human_size(f["size"])} · Excel workbook · <span style="color:#0C7A5A;font-weight:600">Uploaded</span></div></div></div>',
                            unsafe_allow_html=True)
            with c2:
                if st.button("Replace", key="replace_file", width="stretch"):
                    S["gx_replacing"] = True
                    st.rerun()
            with c3:
                st.button("Remove", key="remove_file", width="stretch", on_click=clear_file)
        else:
            up = st.file_uploader("Upload your product configuration workbook.", type=["xlsx"],
                                  key=f"gx_uploader_{S['gx_up_n']}",
                                  help="Excel .xlsx · up to 25 MB. Drag and drop or browse.")
            if up is not None:
                data = up.getvalue()
                fid = hashlib.md5(data).hexdigest()
                if not S["gx_file"] or S["gx_file"]["id"] != fid:
                    S["gx_preview"], S["result"] = None, None
                    for k in ("gx_product", "gx_sheet"):
                        S.pop(k, None)
                S["gx_file"] = {"name": up.name, "size": len(data), "bytes": data, "id": fid}
                S["gx_replacing"] = False
                st.rerun()
            if S["gx_replacing"]:
                st.button("Cancel replace", key="cancel_replace", on_click=lambda: S.update(gx_replacing=False))
        st.markdown('<div class="note" style="margin-top:14px"><b>The workbook is the source of truth.</b> The generated '
                    'document only contains what your workbook contains. AI, if you choose to use it, only helps word the '
                    'Introduction.</div>', unsafe_allow_html=True)
    if S["gx_file"] and not S["gx_replacing"]:
        st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)
        _, cn = st.columns([4, 1.4])
        with cn:
            st.button("Continue", key="s1_next", type="primary", icon=":material/arrow_forward:", icon_position="right",
                      width="stretch", on_click=set_state, kwargs={"gx_step": 2})


def step2():
    f = S["gx_file"]
    info = svc.inspect_workbook(f["bytes"])
    with st.container(key="card_review"):
        st.markdown('<div class="h-sec">Review extracted information</div>', unsafe_allow_html=True)
        st.markdown('<div class="muted" style="margin-bottom:10px">This is what GuideX read from your workbook. Check it looks right before generating.</div>',
                    unsafe_allow_html=True)
        chips = "".join(f'<span class="chip {"ok" if c["found"] else "no"}">{"✓" if c["found"] else "✕"} {c["prefix"]} {c["label"]}</span>'
                        for c in info["checks"])
        st.markdown(chips, unsafe_allow_html=True)
        sheets = info["sheets"]
        if "gx_sheet" not in S or S["gx_sheet"] not in sheets:
            S["gx_sheet"] = info["detected"] if info["detected"] in sheets else sheets[0]
        if not info["detected"]:
            st.error("No sheet starting with “6.3” was found - that sheet holds the order-type task lists.")
        st.selectbox("Order-type task sheet", sheets, key="gx_sheet",
                     help="Detected automatically; change it only if your sheet has a different name.")
        pkey = (f["id"], S["gx_sheet"])
        if not S["gx_preview"] or S["gx_preview"]["key"] != pkey:
            with st.spinner("Reading the workbook…"):
                try:
                    S["gx_preview"] = {"key": pkey, "data": svc.preview_workbook(
                        f["bytes"], S["gx_sheet"], S.get("gx_product", ""))}
                except Exception as exc:                                  # noqa: BLE001
                    S["gx_preview"] = {"key": pkey, "error": f"{type(exc).__name__}: {exc}"}
        pv = S["gx_preview"]
        if "error" in pv:
            st.error("The workbook could not be read with this sheet selection.")
            with st.expander("Technical details"):
                st.code(pv["error"])
        else:
            d = pv["data"]
            warn_cls = "kpi warn" if d["warnings"] else "kpi"
            st.markdown(f"""<div class="kpis">
<div class="kpi"><div class="v">{len(d["order_types"])}</div><div class="l">Order types</div></div>
<div class="kpi"><div class="v">{d["total_tasks"]}</div><div class="l">Tasks</div></div>
<div class="kpi"><div class="v">{len(d["hierarchy"])}</div><div class="l">Product components</div></div>
<div class="{warn_cls}"><div class="v">{len(d["warnings"])}</div><div class="l">Items to review</div></div></div>""",
                        unsafe_allow_html=True)
            st.markdown("**Order types**")
            st.dataframe(
                [{"Order type": o["name"], "Tasks in workbook": o["tasks"], "Tasks in flowchart": o["in_flowchart"],
                  "Swimlanes": o["lanes"], "Decisions": o["decisions"]} for o in d["order_types"]],
                hide_index=True, width="stretch")
            st.caption("Some tasks (for example External/System tasks and rework helpers) are left out of the flowchart "
                       "by design; they remain in the task tables of the guide.")
            e1, e2 = st.columns(2)
            with e1:
                with st.expander(f"Product hierarchy · {len(d['hierarchy'])} node(s)"):
                    for n in d["hierarchy"]:
                        st.markdown(f"- **{n['name']}** · {n['kind']} · {n['attributes']} attribute(s)")
            with e2:
                with st.expander("Other content found"):
                    st.markdown(f"- Introduction paragraphs: **{d['intro_paragraphs']}**\n"
                                f"- Business rules: **{d['business_rules']}** ({d['business_rule_tables']} with tables)\n"
                                f"- Order scenarios: **{d['order_scenarios']}**")
            if d["warnings"]:
                with st.expander(f"⚠️ {len(d['warnings'])} item(s) to review in the workbook", expanded=True):
                    for w in d["warnings"]:
                        st.markdown(f"- {w}")
    st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)
    b1, _, b2 = st.columns([1.2, 3, 1.6])
    with b1:
        st.button("Back", key="s2_back", icon=":material/arrow_back:", width="stretch",
                  on_click=set_state, kwargs={"gx_step": 1})
    with b2:
        ok = bool(S["gx_preview"]) and "error" not in S["gx_preview"]
        st.button("Continue", key="s2_next", type="primary", disabled=not ok, icon=":material/arrow_forward:",
                  icon_position="right", width="stretch", on_click=set_state, kwargs={"gx_step": 3})


def _gemini_key_from_secrets():
    try:
        return st.secrets.get("GEMINI_API_KEY")
    except Exception:
        return None


def step3():
    f = S["gx_file"]
    S.setdefault("gx_product", svc.default_product_name(f["name"]))
    S.setdefault("gx_issued_by", "Business Process Management Section")
    S.setdefault("gx_rev_no", "0")
    S.setdefault("gx_issue_no", "0")
    S.setdefault("gx_rev_date", datetime.date.today())
    S.setdefault("gx_pptx", True)
    S.setdefault("gx_c_des", "DGM/BPM")
    with st.container(key="card_config"):
        st.markdown('<div class="h-sec">Configure and generate</div>', unsafe_allow_html=True)
        st.markdown('<div class="muted" style="margin-bottom:12px">These details appear on the cover and the document control pages.</div>',
                    unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        with c1:
            st.text_input("Product / service name *", key="gx_product", placeholder="e.g. CYS_EdgeDefend")
            st.text_input("Issued by", key="gx_issued_by")
            st.text_input("Revision number", key="gx_rev_no")
        with c2:
            st.text_input("Document number", key="gx_doc_no", placeholder="e.g. SDG-EDGE-001")
            st.text_input("Issue number", key="gx_issue_no")
            st.date_input("Revision date", key="gx_rev_date", format="DD/MM/YYYY")
        st.markdown("**Outputs**")
        st.toggle("Also create the editable PowerPoint with all flowcharts", key="gx_pptx",
                  help="Native shapes and connectors - edit any diagram you are not happy with.")
        with st.expander("Contact for inquiries (optional)"):
            k1, k2 = st.columns(2)
            with k1:
                st.text_input("Name", key="gx_c_name")
                st.text_input("Designation", key="gx_c_des")
                st.text_input("Telephone", key="gx_c_tel")
            with k2:
                st.text_input("Mobile", key="gx_c_mob")
                st.text_input("Extension", key="gx_c_ext")
                st.text_input("E-mail", key="gx_c_mail")
        server_key = _gemini_key_from_secrets()
        with st.expander("AI-assisted introduction (optional)"):
            st.markdown('<div class="note">AI only helps word the <b>Introduction</b>. It must not add specifications, systems, '
                        'process steps or rules that are not in the workbook. The flowcharts never depend on AI.</div>',
                        unsafe_allow_html=True)
            S.setdefault("gx_use_ai", bool(server_key))
            st.toggle("Use Gemini for the Introduction", key="gx_use_ai")
            if S["gx_use_ai"] and not server_key:
                st.text_input("Gemini API key", type="password", key="gx_ai_key",
                              help="Used for this run only; never stored.")

    st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)
    b1, _, b2 = st.columns([1.2, 3, 2.2])
    with b1:
        st.button("Back", key="s3_back", icon=":material/arrow_back:", width="stretch",
                  on_click=set_state, kwargs={"gx_step": 2})
    ready = bool(S["gx_product"].strip()) and bool(S.get("gx_sheet"))
    with b2:
        go_gen = st.button("Generate Guide", key="s3_gen", type="primary", disabled=not ready,
                           icon=":material/auto_awesome:", width="stretch")
    if not S["gx_product"].strip():
        st.info("Enter the product / service name to continue.")

    if go_gen and ready:
        gemini_key = (server_key if S["gx_use_ai"] else None) or ((S.get("gx_ai_key") or None) if S["gx_use_ai"] else None)
        params = svc.GenParams(
            product=S["gx_product"].strip(), sheet=S["gx_sheet"], doc_no=S.get("gx_doc_no", "").strip(),
            issue_no=S["gx_issue_no"].strip(), issued_by=S["gx_issued_by"].strip(), revision_no=S["gx_rev_no"].strip(),
            revision_date=S["gx_rev_date"].strftime("%d.%m.%Y"), export_pptx=S["gx_pptx"], gemini_key=gemini_key,
            contacts=dict(contact_name=S.get("gx_c_name", ""), contact_designation=S.get("gx_c_des", ""),
                          contact_telephone=S.get("gx_c_tel", ""), contact_mobile=S.get("gx_c_mob", ""),
                          contact_extension=S.get("gx_c_ext", ""), contact_email=S.get("gx_c_mail", "")),
        )
        with st.status("Generating your Service Delivery Guide…", expanded=True) as status:
            try:
                out = svc.generate(f["bytes"], params, progress=lambda m: status.write(m))
                S["result"] = {"out": out, "file": f["id"], "product": params.product}
                log_activity(TOOL, params.product, f"guide generated · {len(out.order_types)} order types")
                status.update(label="Done — your files are ready", state="complete", expanded=False)
                S["gx_step"] = 4
                st.rerun()
            except ValueError as exc:
                status.update(label="Could not read the workbook", state="error")
                st.error(f"The workbook could not be processed: {exc}")
            except Exception as exc:                                      # noqa: BLE001
                status.update(label="Something went wrong", state="error")
                st.error("Generation failed. Please check the workbook and try again.")
                with st.expander("Technical details"):
                    st.code(f"{type(exc).__name__}: {exc}")


def step4():
    res = S["result"]
    if not res:
        st.info("Nothing generated yet.")
        st.button("Back to configuration", on_click=set_state, kwargs={"gx_step": 3})
        return
    out = res["out"]
    total_tasks = sum(o["tasks"] for o in out.order_types)
    n_imgs = sum(o["images"] for o in out.order_types)
    with st.container(key="card_results"):
        st.markdown(f'<div class="h-sec">Your files · {res["product"]}</div>', unsafe_allow_html=True)
        warn_cls = "kpi warn" if out.warnings else "kpi"
        st.markdown(f"""<div class="kpis">
<div class="kpi"><div class="v">{len(out.order_types)}</div><div class="l">Order types</div></div>
<div class="kpi"><div class="v">{total_tasks}</div><div class="l">Tasks processed</div></div>
<div class="kpi"><div class="v">{n_imgs}</div><div class="l">Flowchart images</div></div>
<div class="{warn_cls}"><div class="v">{len(out.warnings)}</div><div class="l">Warnings to review</div></div></div>""",
                    unsafe_allow_html=True)
        d1, d2, d3, d4 = st.columns(4)
        with d1:
            st.download_button("Guide (.docx)", out.docx_bytes, out.docx_name,
                               "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                               type="primary", icon=":material/download:", width="stretch")
        with d2:
            if out.pptx_bytes:
                st.download_button("Editable (.pptx)", out.pptx_bytes, out.pptx_name,
                                   "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                                   icon=":material/download:", width="stretch")
            else:
                st.button("PowerPoint not created", disabled=True, width="stretch")
        with d3:
            st.download_button("Images (.zip)", out.flow_zip_bytes, out.flow_zip_name, "application/zip",
                               icon=":material/download:", width="stretch")
        with d4:
            st.download_button("Everything (.zip)", out.all_zip_bytes, out.all_zip_name, "application/zip",
                               icon=":material/download:", width="stretch")
        st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
        if out.warnings:
            with st.expander(f"⚠️ {len(out.warnings)} item(s) to review before sharing the guide", expanded=True):
                for w in out.warnings:
                    st.markdown(f"- {w}")
        else:
            st.success("No warnings — the workbook looks clean.")
    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
    with st.container(key="card_preview"):
        st.markdown('<div class="h-sec">Flowchart preview</div>', unsafe_allow_html=True)
        names = [o["name"] for o in out.order_types if out.flow_images.get(o["name"])]
        if names:
            for tab, nm in zip(st.tabs(names), names):
                with tab:
                    imgs = out.flow_images[nm]
                    cols = st.columns(len(imgs)) if len(imgs) > 1 else [st.container()]
                    for i, (c, png) in enumerate(zip(cols, imgs), start=1):
                        with c:
                            st.image(png, caption=f"{nm} · part {i} of {len(imgs)}" if len(imgs) > 1 else nm,
                                     width="stretch")
        else:
            st.info("No flowcharts were produced — every task was excluded or had no lane.")
    st.markdown("<div style='height:6px'></div>", unsafe_allow_html=True)
    b1, _, b2 = st.columns([1.6, 3, 1.8])
    with b1:
        st.button("Back to settings", key="s4_back", icon=":material/arrow_back:", width="stretch",
                  on_click=set_state, kwargs={"gx_step": 3})
    with b2:
        st.button("Start a new guide", key="s4_new", width="stretch", icon=":material/refresh:",
                  on_click=clear_file)


def render():
    init_state()
    # GuideX needs a file for steps 2-4
    if S["gx_step"] > 1 and not S["gx_file"]:
        S["gx_step"] = 1
    stepper(S["gx_step"])
    {1: step1, 2: step2, 3: step3, 4: step4}[S["gx_step"]]()
