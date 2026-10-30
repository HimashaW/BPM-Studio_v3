"""Scenario-style test cases (the format a tester writes by hand), generated from the workbook.

Each scenario is one row of the UAT sheet:
    Order Type | IT System | Test Scenario (with numbered steps) | Expected in CRM | OSS | BSS

Nothing is combined or multiplied: every scenario comes from ONE fact in the workbook
(an attribute rule, a task, a price plan, a requirement ...) and carries a reference to it.
`core=True` marks the scenarios a tester would normally write; the rest are optional extras that
the user can tick on the Review step.

The templates are generic - they look for patterns in the data (conditional visibility comments,
hidden SR tasks, cancellable work orders, override price lines ...), not for product names.
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

import yaml

LIBRARY = Path(__file__).with_name("standard_scenarios.yaml")

# ordering of the groups inside the sheet
G_ATTR, G_MAP, G_COND, G_CHILD, G_DEFAULT, G_SERVCAT, G_PRICE, G_MODIFY, G_EXTRA = 10, 11, 12, 13, 14, 15, 16, 17, 18
G_OSS, G_FAULT, G_STD, G_REQ, G_CUSTOM = 20, 30, 35, 40, 50

CATEGORIES = {
    "attr": "Attributes", "map": "Mapping & copy", "cond": "Conditional logic", "child": "Child products",
    "default": "Defaults & formats", "servcat": "System mapping", "price": "Pricing", "modify": "Modification",
    "flow": "Order flow", "sr": "Service requests", "task": "Task rules", "fault": "Fault handling",
    "std": "Standard", "req": "Special requirements", "custom": "Your own",
}


@dataclass
class Scenario:
    key: str
    category: str                # one of CATEGORIES values
    order_type: str
    it_system: str
    text: str                    # title line, then numbered steps on following lines
    crm: str = ""
    oss: str = ""
    bss: str = ""
    core: bool = True
    basis: str = ""              # where in the workbook it comes from
    group: int = 99
    order_idx: int = 0
    seq: float = 0.0

    @property
    def title(self) -> str:
        return self.text.split("\n", 1)[0]

    def sort_key(self):
        return (self.group, self.order_idx, self.seq, self.key)


# ------------------------------------------------------------------------------- helpers
def _norm(t) -> str:
    return re.sub(r"[^a-z0-9]+", " ", str(t or "").lower()).strip()


def _clean(t) -> str:
    """One readable paragraph from a spreadsheet cell: no bullets, no leading underscores, single spaces."""
    lines = [re.sub(r"^[\s*_\-•]+", "", ln).strip() for ln in str(t or "").replace("\r", "").split("\n")]
    return " ".join(ln for ln in lines if ln)


def _sentences(t) -> list[str]:
    """Each bullet / line of a cell as a sentence ending in a full stop; 'this is a hidden task' lines dropped."""
    out = []
    for ln in str(t or "").replace("\r", "").split("\n"):
        ln = re.sub(r"^[\s*_\-•]+", "", ln).strip()
        if not ln or re.match(r"this is a hidden task", ln, re.I):
            continue
        out.append(ln if ln[-1] in ".!?" else ln + ".")
    return out


def _first_sentences(t: str, n: int = 2, limit: int = 260) -> str:
    out = " ".join(_sentences(t)[:n]).strip()
    out = re.sub(r"\b[Hh]as to trigger\b", "triggers", out)
    return out if len(out) <= limit else out[: limit - 1].rstrip() + "…"


def _title_of(text: str, limit: int = 170) -> str:
    """First sentence of a paragraph, cut at a word boundary if it is very long."""
    t = _clean(text)
    first = re.split(r"(?<=[.!?])\s+(?=[A-Z‘'(])", t)[0].strip()
    if len(first) <= limit:
        return first
    return first[:limit].rsplit(" ", 1)[0].rstrip(",;:") + "…"


def short_order_type(title: str) -> str:
    """'Comprehensive Task List for X - Order Type CREATE SO' -> 'CREATE'; 'DELETE SERVICE SO' -> 'DELETE'."""
    t = " ".join(str(title).split())
    m = re.search(r"order type\s+(.+?)\s*$", t, re.I) or re.search(r"-\s*([A-Za-z][A-Za-z\- ]*?)\s+order type", t, re.I)
    name = (m.group(1) if m else t).strip()
    name = re.sub(r"\s+SO$", "", name, flags=re.I).strip().upper()
    base = name.replace(" SERVICE", "")
    return base if base in {"DELETE", "SUSPEND", "RESUME", "CREATE"} else name


def _join(items, last="and") -> str:
    items = [str(i) for i in items if str(i)]
    if len(items) <= 1:
        return "".join(items)
    return ", ".join(items[:-1]) + f" {last} " + items[-1]


def _quote(items) -> str:
    return _join([f"'{i}'" for i in items])


def _human_condition(cond: str) -> str:
    c = str(cond or "").strip()
    c = re.sub(r"^\s*if\s*", "", c, flags=re.I)
    c = c.replace("[", "").replace("]", "").replace("'", "")
    c = re.sub(r"\s*<>\s*<Blank>", " is not blank", c, flags=re.I)
    c = re.sub(r"\s*=\s*", " = ", c)
    return " ".join(c.split()).rstrip(".")


_TEST_COND = re.compile(r"=\s*'?TEST'?(?![\w ]*(?:CONFIRM|REQ))", re.I)


def _is_hidden(task) -> bool:
    return "hidden" in (task.work_group or "").lower()


def _visible_prev(tasks, idx):
    """The nearest earlier task a person actually works on (not hidden, not automatic, not conditional)."""
    for j in range(idx - 1, -1, -1):
        t = tasks[j]
        wg = (t.work_group or "").lower()
        if _is_hidden(t) or wg.startswith("auto") or (t.it_system or "").upper() == "SOA" or (t.task_condition or "").strip():
            continue
        return t
    return None


def _values(attr, n: int = 8) -> str:
    vals = list(attr.values)
    if not vals:
        return ""
    return _join(vals, "and") if len(vals) <= n else f"{len(vals)} values"


class _Ctx:
    def __init__(self, src, product: str):
        self.src, self.ex = src, src.extra
        self.product = product or src.product_name
        self.nodes = src.nodes
        self.root = next((n for n in src.nodes if n.kind.strip().lower() == "product"), src.nodes[0] if src.nodes else None)
        self.sections = [(short_order_type(s.order_type), s) for s in src.sections]
        self.ot_index = {ot: i for i, (ot, _) in enumerate(self.sections)}
        self.attrs = {}
        for n in src.nodes:
            for a in n.attributes:
                self.attrs.setdefault(_norm(a.entity_name), (n, a))

    def attr(self, name: str):
        hit = self.attrs.get(_norm(name))
        return hit[1] if hit else None

    def mapping(self, a) -> str:
        return self.ex.hier.get(a.source_row).mapping if self.ex.hier.get(a.source_row) else ""

    def in_system(self, a, system: str) -> bool:
        return system.lower() in self.mapping(a).lower()

    def comment(self, a) -> str:
        e = self.ex.hier.get(a.source_row)
        return (e.comments if e and e.comments else (a.comments or "")).strip()

    def mandatory(self, a) -> bool:
        m = (a.mandatory or "").strip().lower()
        return m not in {"no", "n", "optional", "false"}

    def oi(self, ot: str) -> int:
        return self.ot_index.get(ot, 99)


# ----------------------------------------------------------------------------- templates
def _attribute_catalogue(c: _Ctx):
    if not c.root or not c.root.attributes:
        return []
    mand = [a.entity_name for a in c.root.attributes if c.mandatory(a)]
    opt = [a.entity_name for a in c.root.attributes if not c.mandatory(a)]
    lov = [f"{a.entity_name} ({_values(a)})" for a in c.root.attributes if a.values]
    to_oss = [a.entity_name for a in c.root.attributes if c.in_system(a, "OSS")]
    to_bss = [a.entity_name for a in c.root.attributes if c.in_system(a, "BSS")]
    crm = f"All {len(c.root.attributes)} attributes are available in CRM with the relevant values."
    crm += f"\nMandatory: {_join(mand)}." + (f"\nOptional: {_join(opt)}." if opt else "")
    if lov:
        crm += "\nValue lists: " + "; ".join(lov) + "."
    return [Scenario("attr:catalogue", "Attributes", "CREATE", "CRM",
                     f"Verify all attributes of {c.root.name} are present in CRM and the relevant attributes are mandatory.",
                     crm,
                     "After a valid selection, OSS displays the attributes with the selected values"
                     + (f" ({len(to_oss)} attributes are mapped to OSS)." if to_oss else "."),
                     (f"Attributes mapped to BSS: {_join(to_bss)}." if to_bss else ""),
                     True, "2. Product Hierarchy", G_ATTR)]


def _copy_targets(comment: str) -> list[str]:
    m = re.search(r"copied to(?: the)?\s+(.+?)\s+child products?", comment, re.I)
    return re.findall(r"'([^']+)'", m.group(1)) if m else []


def _mandatory_and_mapping(c: _Ctx):
    if not c.root:
        return []
    cands = [a for a in c.root.attributes if c.mandatory(a) and len(a.values) >= 2]
    cands.sort(key=lambda a: (not c.in_system(a, "BSS"), not bool(_copy_targets(c.comment(a)))))
    pick = cands[:2]
    if len(pick) < 2:
        return []
    a, b = pick
    va, vb = a.values[0], b.values[0]
    kids = sorted({k for x in pick for k in _copy_targets(c.comment(x))})
    text = (f"Verify mandatory validation and successful mapping of {_quote([a.entity_name, b.entity_name])}.\n"
            f"1. Leave '{a.entity_name}' and '{b.entity_name}' blank and try to submit the order.\n"
            f"2. Choose '{a.entity_name}' = {va}.\n3. Choose '{b.entity_name}' = {vb}.\n4. Submit the order.")
    bss = ""
    if kids:
        bss = f"{_quote([a.entity_name, b.entity_name])} are copied to the child products ({_join(kids)}) and used for charging."
    return [Scenario("map:mandatory", "Mapping & copy", "CREATE", "CRM", text,
                     "The system blocks submission while either attribute is blank and allows it only when both are populated.",
                     f"'{a.entity_name}' is displayed as '{va}' and '{b.entity_name}' is displayed as '{vb}' without any mismatch.",
                     bss, True, "2. Product Hierarchy (mandatory, copy-to-child comments)", G_MAP)]


_COND_RE = re.compile(r"(visible(?:\s+and\s+mandatory)?|mandatory|populated)\s+only\s+when\s+'([^']+)'\s+"
                      r"(?:attribute\s+value\s+)?is\s+(not\s+)?(?:set\s+to\s+)?'([^']+)'", re.I)


def _conditional_attributes(c: _Ctx):
    out = []
    for n in c.nodes:
        for a in n.attributes:
            cm = c.comment(a)
            m = _COND_RE.search(cm)
            if not m:
                continue
            kind, ctl_name, neg, val = m.group(1).lower(), m.group(2), bool(m.group(3)), m.group(4)
            ctl = c.attr(ctl_name)
            others = [v for v in (ctl.values if ctl else []) if _norm(v) != _norm(val)]
            other = others[0] if others else "any other value"
            shown_when, hidden_when = (other, val) if neg else (val, other)
            extra_rules = []
            gt = re.search(r"greater than\s+(\d+)", cm, re.I)
            if re.search(r"numerical|numeric", cm, re.I):
                extra_rules.append("non-numeric entries are rejected")
            if gt:
                extra_rules.append(f"values not greater than {gt.group(1)} are rejected")
            steps = [f"1. Select '{ctl_name}' = '{hidden_when}' and check '{a.entity_name}'.",
                     f"2. Select '{ctl_name}' = '{shown_when}' and check '{a.entity_name}'."]
            if extra_rules:
                steps.append(f"3. Enter {'a non-numeric value, then a value not greater than ' + gt.group(1) if gt else 'an invalid value'}, then a valid value.")
            kids = _copy_targets(cm)
            out.append(Scenario(
                f"cond:{_norm(a.entity_name)}", "Conditional logic", "CREATE", "CRM",
                f"Verify conditional handling of '{a.entity_name}' based on '{ctl_name}'.\n" + "\n".join(steps),
                f"'{a.entity_name}' is {'visible and mandatory' if 'mandatory' in kind else 'visible'} only when '{ctl_name}' "
                f"{'is not' if neg else '='} '{val}'; it is hidden otherwise."
                + (f" Invalid entries: {_join(extra_rules)}." if extra_rules else ""),
                f"'{a.entity_name}' and '{ctl_name}' are passed to OSS correctly.",
                (f"The value is copied to {_join(kids)} and used for charging." if kids else ""),
                True, f"2. Product Hierarchy - {a.entity_name} comment", G_COND))
    return out


def _child_products(c: _Ctx):
    """Child products that are 'Conditional' and only visible for certain values of an attribute."""
    groups: dict[str, list] = {}
    for n in c.nodes:
        e = c.ex.hier.get(n.source_row)
        if not e or not e.comments:
            continue
        m = _COND_RE.search(e.comments)
        if m:
            groups.setdefault(m.group(2), []).append((n, e, bool(m.group(3)), m.group(4)))
    out = []
    for ctl_name, items in groups.items():
        ctl = c.attr(ctl_name)
        steps, expect, bss = [], [], []
        for i, (n, e, neg, val) in enumerate(items, start=1):
            pick = next((v for v in (ctl.values if ctl else []) if (_norm(v) != _norm(val)) == neg), val if not neg else "")
            steps.append(f"{i}. Select '{ctl_name}' = '{pick}'.")
            expect.append(f"'{n.name}' is added" + (" as mandatory" if "mandatory" in e.comments.lower() else "")
                          + f" when '{ctl_name}' {'is not' if neg else '='} '{val}'.")
            bss.append(f"'{n.name}'" + (f" ({e.charge})" if e.charge else ""))
        steps.append(f"{len(items) + 1}. Check the child products on the order.")
        out.append(Scenario(
            f"child:{_norm(ctl_name)}", "Child products", "CREATE", "CRM",
            f"Verify child product selection based on '{ctl_name}'.\n" + "\n".join(steps), " ".join(expect),
            "Only the applicable child product and its attributes are passed to OSS.",
            f"Charges are created for {_join(bss)} only when applicable.", True, "2. Product Hierarchy (child product comments)", G_CHILD))
    return out


_DEFAULT_RE = re.compile(r"de\w*ult\s+value\s*(?:is|:)?\s*'?([^'\n]+?)'?\s*\.?\s*$", re.I)


def _defaults(c: _Ctx):
    items = []
    for n in c.nodes:
        for a in n.attributes:
            m = _DEFAULT_RE.search(_clean(c.comment(a)))
            if m:
                items.append((n.name, a.entity_name, m.group(1).strip()))
    out = []
    if items:
        lines = [f"{i}. Open a new order and check '{a}'" + (f" of '{p}'" if p != c.root.name else "") + "." for i, (p, a, _) in enumerate(items, start=1)]
        out.append(Scenario("default:static", "Defaults & formats", "CREATE", "CRM",
                            "Verify the default values of attributes on a new order.\n" + "\n".join(lines),
                            "Defaults are pre-selected: " + "; ".join(f"{a} = {v}" for _, a, v in items) + ".",
                            "Default values are passed to OSS unchanged.", "", False, "2. Product Hierarchy (default value comments)", G_EXTRA))
    # defaults that depend on another attribute (comment written on a Value row)
    by_attr: dict[tuple, list] = {}
    for vn in c.ex.value_notes:
        m = re.search(r"default value when\s+'([^']+)'\s+is\s+set\s+to\s+'([^']+)'", vn.comment, re.I)
        if m:
            by_attr.setdefault((vn.product, vn.attribute, m.group(1)), []).append((m.group(2), vn.value))
    for (prod, attr, ctl), pairs in by_attr.items():
        steps = [f"{i}. Select '{ctl}' = '{cv}' and check '{attr}'." for i, (cv, _) in enumerate(pairs, start=1)]
        out.append(Scenario(
            f"default:{_norm(attr)}", "Conditional logic", "CREATE", "CRM",
            f"Verify the default '{attr}' value changes with '{ctl}'" + (f" (child product '{prod}')" if prod and prod != c.root.name else "") + ".\n" + "\n".join(steps),
            "; ".join(f"'{ctl}' = {cv} → '{attr}' defaults to {v}" for cv, v in pairs) + ".",
            "", "", True, f"2. Product Hierarchy - {attr} value comments", G_DEFAULT))
    return out


def _formats(c: _Ctx):
    items = []
    for n in c.nodes:
        for a in n.attributes:
            cm = _clean(c.comment(a))
            if re.search(r"date format|DD/MM/YYYY", cm, re.I):
                items.append((a.entity_name.replace("\n", " "), "date in DD/MM/YYYY format"))
            elif re.search(r"character limit", cm, re.I):
                items.append((a.entity_name, "free text within the character limit"))
    if not items:
        return []
    lines = [f"{i}. Enter an invalid and then a valid value in '{a}'." for i, (a, _) in enumerate(items, start=1)]
    return [Scenario("default:formats", "Defaults & formats", "CREATE", "CRM",
                     "Verify input formats of attributes with a defined format.\n" + "\n".join(lines),
                     "; ".join(f"'{a}' accepts {f}" for a, f in items) + ". Invalid entries are rejected.",
                     "Valid values are passed to OSS unchanged.", "", False, "2. Product Hierarchy (format comments)", G_EXTRA)]


def _referenced_attributes(c: _Ctx):
    """Attributes that the order tasks themselves read or change (e.g. SERVICE CATEGORY)."""
    texts = []
    for ot, s in c.sections:
        for t in s.tasks:
            texts.append((ot, t, " ".join(filter(None, [t.task_condition, t.special_configurations, t.mandatory_attributes]))))
    out, done = [], set()
    for n in c.nodes:
        for a in n.attributes:
            k = _norm(a.entity_name)
            if len(k) < 6 or k in done or not (c.in_system(a, "CRM") and c.in_system(a, "OSS")):
                continue
            hits = [(ot, t) for ot, t, tx in texts if k in _norm(tx)]
            if not hits:
                continue
            done.add(k)
            tasks = _join(sorted({t.task_name for _, t in hits}))
            out.append(Scenario(
                f"servcat:{k}", "System mapping", "CREATE", "CRM",
                f"Verify '{a.entity_name}' attribute presence in CRM and its mapping to OSS" + (" and BSS" if c.in_system(a, "BSS") else "") + ".\n"
                f"1. Open a new order and select a value for '{a.entity_name}'.\n2. Submit and check the order in OSS.",
                f"'{a.entity_name}' is present and allows selection" + (f" ({_values(a)})." if a.values else "."),
                f"'{a.entity_name}' is correctly mapped and displayed in OSS; it is used by the tasks {tasks}.",
                f"'{a.entity_name}' is available in BSS." if c.in_system(a, "BSS") else "", True,
                "2. Product Hierarchy (Entity Mapping) + task list", G_SERVCAT))
    return out


def _pricing(c: _Ctx):
    lines = [p for p in c.ex.pricing if p.override]
    groups: dict[tuple, list] = {}
    for p in lines:
        groups.setdefault((p.charge, _norm(p.related)), []).append(p)
    if not groups:
        return []
    mrc = [k for k in groups if k[0] == "MRC"]
    best_mrc = max(mrc, key=lambda k: len(groups[k])) if mrc else None
    out = []
    for (charge, rel), items in groups.items():
        lows, highs = {p.low for p in items}, {p.high for p in items}
        low, high = min(lows, key=lambda x: float(x) if x.replace(".", "", 1).isdigit() else 0), max(
            highs, key=lambda x: float(x) if x.replace(".", "", 1).isdigit() else 0)
        related = items[0].related
        names = [r.strip() for r in related.split("/") if r.strip() and r.strip().upper() != "NA"]
        steps = []
        for r in names:
            a = c.attr(r)
            steps.append(f"Select '{r}'" + (f" = {a.values[0]}" if a and a.values else "") + ".")
        prods = _join(sorted({p.product for p in items}))
        steps += [f"Override the price to a value inside the allowed range ({low} – {high}).",
                  "Try to override the price to a value outside the range."]
        steps = [f"{i}. {s}" for i, s in enumerate(steps, start=1)]
        period = items[0].period
        out.append(Scenario(
            f"price:{charge}:{rel}", "Pricing", "CREATE", "CRM",
            f"Verify price override of {charge} charges based on {_join(names) or 'the selected product'} ({prods}).\n" + "\n".join(steps),
            f"The system allows a price override only within the configured minimum ({low}) and maximum ({high}).",
            "The overridden price is passed to OSS without mismatch.",
            f"The overridden price is updated correctly in BSS ({charge}{', ' + period if period else ''}"
            f"{', prorated' if any(p.prorated.lower().startswith('y') for p in items) else ''}).",
            (charge, rel) == best_mrc, "4. Pricing", G_PRICE))
    return out


def _modification(c: _Ctx):
    rows = [r for r in c.ex.modify_rows if r.action == "UPDATE" and r.order_type not in ("", "NA")]
    out = []
    by_ot: dict[str, list] = {}
    for r in rows:
        by_ot.setdefault(r.order_type, []).append(r)
    for ot, items in by_ot.items():
        attrs = [r for r in items if r.kind == "ATTRIBUTE"]
        feats = [r for r in items if r.kind == "FEATURE"]
        steps, changes = [], []
        for r in attrs:
            a = c.attr(r.name)
            if a and len(a.values) >= 2:
                changes.append((r.name, a.values[0], a.values[1]))
            elif a and not a.values:
                changes.append((r.name, "an existing value", "a new value"))
        if not attrs:
            continue
        sub = next((o.condition for o in c.ex.order_rows if o.order_type == ot), "")
        steps.append("1. Take an active order with " + _join([f"'{n}' = {a}" for n, a, _ in changes[:2]]) + ".")
        steps.append(f"2. Create a {sub or ot} request.")
        steps.append("3. Change " + _join([f"'{n}' to {b}" for n, _, b in changes[:3]]) + ".")
        names = [r.name for r in attrs]
        out.append(Scenario(
            f"modify:{ot}", "Modification", ot, "CRM",
            f"Verify modification of {_quote(names)} ({sub or ot}).\n" + "\n".join(steps),
            f"CRM allows modification of {_quote(names)}.",
            "The modified values are passed to OSS and displayed correctly." + (" The ACTUAL DSP DATE is mandatory to complete the modification task." if any(
                "ACTUAL DSP" in (t.mandatory_attributes or "").upper() for o, s in c.sections if o == ot for t in s.tasks) else ""),
            (f"Feature(s) {_quote([f.name for f in feats])} are updated in BSS." if feats else ""),
            True, "5.2 MDM_Attribute vs Order Scenario", G_MODIFY))
        back = [r.name for r in items + [x for x in c.ex.modify_rows if x.back_update and x not in items] if r.back_update]
        if back:
            out.append(Scenario(
                f"modify:back:{ot}", "Modification", ot, "OSS",
                f"Verify back update of {_quote(back)} from OSS to CRM after the {ot} order.\n"
                f"1. Complete a {ot} order in which these attributes are changed in OSS.\n2. Check the order in CRM.",
                f"{_quote(back)} are updated in CRM from OSS.", "The values changed in OSS are sent back to CRM.", "",
                False, "5.2 (Back update required = Yes)", G_EXTRA))
    return out


def _tasks(c: _Ctx):
    out = []
    cancel_done = confirm_done = False
    for ot, sec in c.sections:
        tasks = sec.tasks
        oi = c.oi(ot)
        # ---- SR creation scenarios
        for i, t in enumerate(tasks):
            spec = _clean(t.special_configurations)
            if not (re.search(r"SR[_ ]?REQ", t.task_name, re.I) or re.search(r"send a SR to CRM", spec, re.I)):
                continue
            prev = _visible_prev(tasks, i)
            cond = _human_condition(t.task_condition)
            chain = []
            for u in tasks[i + 1:]:
                if (u.task_condition or "").strip() and (u.task_condition == t.task_condition or _norm(u.task_name).split(" ")[0] in ("wait", "add", "set")):
                    chain.append(u.task_name)
                else:
                    break
            at = f" during the '{prev.task_name}' task" if prev else ""
            seqf = float(re.match(r"[\d.]+", t.seq or "0").group(0)) if re.match(r"[\d.]+", t.seq or "") else 0.0
            out.append(Scenario(
                f"sr:{ot}:{_norm(t.task_name)}", "Service requests", ot, "OSS",
                f"Verify SR creation ({t.task_name}) in CRM from OSS{at} of the {ot} order when {cond or 'its condition is met'}.\n"
                f"1. Create a {ot} order and progress it to the '{prev.task_name if prev else 'previous'}' task.\n"
                f"2. Make the condition true: {cond}.\n3. Complete the task.",
                "SR is passed to CRM with the correct attributes.",
                (_first_sentences(t.special_configurations, 1) or f"OSS sends SR '{t.task_name}' to CRM via SOA.") + f" Condition: {cond}."
                + (f" Follow-up tasks: {_join(chain)}." if chain else ""),
                "", True, f"{c.src.extra.task_sheet.strip()} - {t.task_name}", G_OSS + 1, oi, seqf))
        # ---- customer confirmation tasks are separate tasks
        conf = [t for t in tasks if re.search(r"confirm", t.task_name, re.I) and not re.search(r"SR[_ ]?REQ|WAIT", t.task_name, re.I)]
        if len(conf) >= 2:
            wgs = sorted({t.work_group for t in conf if t.work_group})
            out.append(Scenario(
                f"task:confirm:{ot}", "Task rules", ot, "OSS",
                f"Verify separate OSS tasks exist for customer confirmation ({_quote([t.task_name for t in conf])}) in the {ot} order.\n"
                f"1. Progress a {ot} order to the first confirmation task and complete it.\n"
                f"2. Trigger the condition of the second confirmation task ({_human_condition(conf[-1].task_condition) or 'see task list'}).",
                "", f"A dedicated, distinct OSS task is created for each customer confirmation" + (f" for the same work group ({_join(wgs)})." if len(wgs) == 1 else "."),
                "", not confirm_done, f"{c.src.extra.task_sheet.strip()} - {ot}", G_OSS, oi, 1.0))
            confirm_done = True
        # ---- test service path
        test_tasks = [t for t in tasks if _TEST_COND.search(t.task_condition or "")]
        if test_tasks:
            close = next((t for t in tasks if re.search(r"close order", t.task_name, re.I)), None)
            texts = [f"{t.task_name}: " + _first_sentences(t.special_configurations, 2)
                     for t in test_tasks + ([close] if close else []) if t.special_configurations]
            bss_has = bool(close and close.special_configurations and "BSS" in close.special_configurations)
            if bss_has:
                texts = texts[:-1]
            bss_txt = next((f"{t.task_name}: " + _first_sentences(t.special_configurations, 2) for t in ([close] if close else [])
                            if t.special_configurations and "BSS" in t.special_configurations), "")
            out.append(Scenario(
                f"task:test:{ot}", "Task rules", ot, "OSS",
                f"Verify the test-service path of the {ot} order ({_quote([t.task_name for t in test_tasks])} when {_human_condition(test_tasks[0].task_condition)}).\n"
                f"1. Submit a {ot} order with {_human_condition(test_tasks[0].task_condition)}.\n2. Progress the order and complete {'CLOSE ORDER' if close else 'the last task'}.",
                "The order is accepted with the test service category.",
                " ".join(texts), bss_txt,
                True, f"{c.src.extra.task_sheet.strip()} - {test_tasks[0].task_name}", G_OSS + 2, oi, 1.0))
        # ---- mandatory attributes before completing a task
        mand = [(t.task_name, t.mandatory_attributes) for t in tasks if (t.mandatory_attributes or "").strip()]
        if mand:
            attrs = sorted({m.strip().title() if m.strip().isupper() else m.strip() for _, m in mand}, key=str.lower)
            attrs = list({a.lower(): a for a in attrs}.values())
            out.append(Scenario(
                f"task:mand:{ot}", "Task rules", ot, "OSS",
                f"Verify mandatory task attribute(s) {_quote(attrs)} are enforced in the {ot} order ({_quote([n for n, _ in mand])}).\n"
                "1. Open each task listed and try to complete it without the attribute.\n2. Enter the attribute and complete the task.",
                "", f"The task cannot be completed until {_quote(attrs)} is populated; it completes once the value is entered.", "",
                False, f"{c.src.extra.task_sheet.strip()} - {ot}", G_OSS + 3, oi, 1.0))
        # ---- cancellation
        canc = [t for t in tasks if str(t.cancellation_wo or "").strip().lower() in ("yes", "y")]
        if canc:
            out.append(Scenario(
                f"task:cancel:{ot}", "Task rules", ot, "OSS",
                f"Verify the OSS workflow cancellation path of the {ot} order.\n1. Submit a {ot} order.\n"
                f"2. Initiate the cancellation request while {'any of ' if len(canc) > 1 else ''}{_quote([t.task_name for t in canc])} {'is' if len(canc) == 1 else 'is'} in progress.",
                "", "The workflow follows the defined cancellation steps, updates the order status and triggers the necessary rollback tasks.", "",
                not cancel_done, f"{c.src.extra.task_sheet.strip()} (WO Cancellation = Yes)", G_OSS + 4, oi, 1.0))
            cancel_done = True
        # ---- end-to-end flow
        main = [t for t in tasks if not (t.task_condition or "").strip() and not _is_hidden(t)]
        subs = [o for o in c.ex.order_rows if o.order_type == ot]
        remarks = [o.status_remarks for o in subs if o.status_remarks]
        sub_txt = (f" ({_join([o.condition for o in subs])})" if len(subs) > 1 else "")
        steps = [f"1. Raise a {ot} order in CRM" + (f" for each scenario: {_join([o.condition for o in subs])}." if len(subs) > 1 else
                                                    (f" ('{subs[0].condition}')." if subs else "."))]
        steps.append("2. Progress the order through " + " → ".join(t.task_name for t in main[:-1] or main) + ".")
        steps.append(f"3. Complete {main[-1].task_name} and check that the order is closed." if main else "3. Complete the last task.")
        wgs = _join(sorted({t.work_group for t in main if t.work_group and not t.work_group.lower().startswith("auto")}))
        out.append(Scenario(
            f"flow:{ot}", "Order flow", ot, "OSS", f"Verify the {ot} order flow end-to-end{sub_txt}.\n" + "\n".join(steps),
            f"The {ot} order is created in CRM and sent to OSS.",
            f"OSS creates the tasks in sequence ({len(main)} tasks" + (f"; work groups: {wgs}" if wgs else "") + ") and the order closes successfully.",
            (f"The product status in BSS is updated with the status remark(s): {_join(remarks)}." if remarks else ""),
            True, f"{c.src.extra.task_sheet.strip()} - {ot}" + (" + 5.1 sub scenarios" if subs else ""), G_OSS + 5, oi, 1.0))
    return out


def _faults(c: _Ctx):
    out = []
    for prod in dict.fromkeys(f.product for f in c.ex.faults):
        rows = [f for f in c.ex.faults if f.product == prod]
        f0 = rows[0]
        natures = [r.nature for r in rows if r.nature]
        lev = [f"Level 1: {f0.l1}" if f0.l1 else "", f"Level 2: {f0.l2}" if f0.l2 and f0.l2.upper() != "NA" else "",
               f"Level 3: {f0.l3}" if f0.l3 and f0.l3.upper() != "NA" else ""]
        out.append(Scenario(
            f"fault:{_norm(prod)}", "Fault handling", "-", "CRM",
            f"Verify fault handling and assignment to the correct work groups for {prod}.\n"
            f"1. Raise a fault for {prod}" + (f" with 'Nature of Fault' = {_quote(natures[:3])}" if natures else "") + ".\n"
            "2. Check the work group the fault is routed to.\n3. Escalate the fault through the handling levels.",
            "Fault is created successfully" + (" with the selected 'Nature of Fault'." if natures else "."),
            f"Fault is directed to the default work group {f0.work_group}" + (f" (OLA {f0.ola})" if f0.ola else "") + "; "
            + "; ".join(x for x in lev if x) + ".", "", True, "10.0 Fault Handling", G_FAULT))
    return out


_ORDER_WORDS = [(r"new connection", "CREATE"), (r"suspen", "SUSPEND"), (r"resum", "RESUME"), (r"terminat|disconnect|\bdelete\b", "DELETE")]


def _requirements(c: _Ctx):
    out = []
    for r in c.ex.requirements:
        low = r.text.lower()
        ots = [ot for pat, ot in _ORDER_WORDS if re.search(pat, low)]
        order = ", ".join(ots) if ots else "ALL"
        systems = [s for s in ("CRM", "OSS", "BSS") if s in r.systems.upper().replace("SOA", "OSS")]
        primary = systems[0] if systems else "CRM"
        body = _clean(r.text)
        title = _title_of(r.text)
        exp = {s: "" for s in ("CRM", "OSS", "BSS")}
        exp[primary] = body if len(body) <= 600 else body[:597] + "…"
        for s in systems[1:]:
            exp[s] = f"See requirement {r.rid}: the behaviour above is applied in {s}."
        out.append(Scenario(
            f"req:{r.rid}", "Special requirements", order, primary,
            f"Verify requirement {r.rid} ({r.rtype or 'Special requirement'}): {title}", exp["CRM"], exp["OSS"], exp["BSS"],
            False, f"3. Special Requirements - {r.rid}", G_REQ, 0, float(re.sub(r"\D", "", r.rid) or 0)))
    return out


def _library(c: _Ctx):
    try:
        entries = yaml.safe_load(LIBRARY.read_text(encoding="utf-8")) or []
    except Exception:                                                     # noqa: BLE001 - optional file
        return []
    has_bss = any("bss" in (e.mapping.lower()) for e in c.ex.hier.values()) or bool(c.ex.pricing)
    has_test = any(_TEST_COND.search(t.task_condition or "") for _, s in c.sections for t in s.tasks)
    flags = {"always": True, "bss": has_bss, "test_service": has_test, "pricing": bool(c.ex.pricing)}
    out = []
    for i, e in enumerate(entries):
        if not flags.get(str(e.get("when", "always")), False):
            continue
        fmt = lambda k: str(e.get(k, "") or "").replace("{product}", c.product).strip()
        out.append(Scenario(f"std:{i}", "Standard", str(e.get("order_type", "CREATE")), str(e.get("it_system", "OSS")),
                            fmt("scenario"), fmt("crm"), fmt("oss"), fmt("bss"), bool(e.get("core", True)),
                            "standard_scenarios.yaml", G_STD, 0, float(i)))
    return out


def build_scenarios(src, product: str = "") -> list[Scenario]:
    """All scenarios the workbook supports, sorted in a sensible reading order."""
    for n in src.nodes:
        for a in n.attributes:
            a.entity_name = " ".join(a.entity_name.split())
    c = _Ctx(src, product)
    if not c.root:
        return []
    found: list[Scenario] = []
    for fn in (_attribute_catalogue, _mandatory_and_mapping, _conditional_attributes, _child_products, _defaults, _formats,
               _referenced_attributes, _pricing, _modification, _tasks, _faults, _requirements, _library):
        try:
            found += fn(c)
        except Exception as exc:                                          # noqa: BLE001 - one template must never kill the plan
            src.warnings.append(f"[Scenarios] {fn.__name__.strip('_')} skipped: {type(exc).__name__}: {exc}")
    found.sort(key=lambda s: s.sort_key())
    seen, uniq = set(), []
    for s in found:
        if s.key in seen:
            continue
        seen.add(s.key)
        uniq.append(s)
    return uniq


def coverage_notes(src) -> list[str]:
    """Things in the workbook that no scenario covers - shown to the user so nothing is silently dropped."""
    notes = []
    have = {short_order_type(s.order_type) for s in src.sections}
    listed = [o.order_type for o in src.extra.order_rows]
    missing = sorted({o for o in listed if o not in have})
    if missing:
        notes.append(f"Order type(s) in 5.1 without a task list, so no flow scenario: {', '.join(missing)}.")
    if not src.extra.pricing:
        notes.append("No price plans were found (sheet 4), so no price-override scenarios.")
    if not src.extra.faults:
        notes.append("No fault-handling rows were found (sheet 10.0), so no fault scenario.")
    return notes
