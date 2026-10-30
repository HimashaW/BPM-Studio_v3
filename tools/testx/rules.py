"""Turn the workbook's business rules into constraints the engine can enforce.

What is understood automatically (everything else is listed as "text only" for manual handling):

  1. Look-up tables      "Package Type LOVs | CPE Model | Max Users"  ->  IF Package Type = Pro THEN CPE Model = TZ80 ...
  2. Tick matrices       "Feature Attribute | Standard | Pro | ..."   ->  IF Package Type = Pro THEN Traffic Filtering = No ...
  3. Four sentence forms in sheet 3, e.g.
        When "A" attribute value is 'x', only 'y' value shall be able to select from "B" ...
        When "A" attribute value is 'x', the 'y' value shall not be able to select from "B" ...
        When "A" attribute value is 'x', the "B" attribute shall be hidden.
        "B" attribute shall be populated only when 'A' value is 'x'.

Every detected rule is shown to the user, who can switch it off. A rule that mentions a value the
workbook does not list is switched off automatically and flagged - that is also a data-quality check.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field

from tools.testx.model import BLANK, Cond, Rule

TICKS = {"✓", "✔", "✅", "☑", "√", "x", "yes", "y", "true", "1"}
YES = ("yes", "y", "true", "enabled", "on")
NO = ("no", "n", "false", "disabled", "off")
NA = {"n/a", "na", "-", "—", "not applicable"}


@dataclass
class AttrInfo:
    name: str
    values: list
    mandatory: bool | None


@dataclass
class Catalog:
    """Every attribute name in the workbook with its list of values (merged across components)."""
    attrs: dict = field(default_factory=dict)        # casefold name -> AttrInfo

    @classmethod
    def from_nodes(cls, nodes) -> "Catalog":
        cat = cls()
        for n in nodes:
            for a in n.attributes:
                key = a.entity_name.strip().casefold()
                info = cat.attrs.setdefault(key, AttrInfo(a.entity_name.strip(), [], None))
                for v in a.values:
                    if v not in info.values:
                        info.values.append(v)
                m = _mandatory_flag(a.mandatory)
                if m is not None:
                    info.mandatory = m if info.mandatory is None else (info.mandatory or m)
        return cat

    def get(self, name: str):
        return self.attrs.get((name or "").strip().casefold())

    def canon(self, attr: str, text: str):
        """The workbook's own spelling of a value of `attr`, or None if it is not listed."""
        info = self.get(attr)
        if not info:
            return None
        t = _squash(text).casefold()
        for v in info.values:
            if _squash(v).casefold() == t:
                return v
        return None


def _mandatory_flag(raw) -> bool | None:
    t = (raw or "").strip().casefold() if isinstance(raw, str) else ""
    if t in {"yes", "y", "true", "mandatory", "m", "required"}:
        return True
    if t in {"no", "n", "false", "optional", "o"}:
        return False
    return None


def _squash(t) -> str:
    return " ".join(str(t or "").split())


_Q = {"“": '"', "”": '"', "„": '"', "«": '"', "»": '"', "‘": "'", "’": "'", "‚": "'"}


def _norm(text: str) -> str:
    return _squash("".join(_Q.get(ch, ch) for ch in text or ""))


def _short(items, n=3):
    items = list(items)
    return "; ".join(items[:n]) + (f"; (+{len(items) - n} more)" if len(items) > n else "")


# ------------------------------------------------------------------------- sentences
def _sentence_patterns(cat: Catalog):
    names = sorted((i.name for i in cat.attrs.values()), key=len, reverse=True)
    if not names:
        return None
    alt = "|".join(re.escape(n) for n in names)

    def A(g):
        return rf"[\"']?(?P<{g}>{alt})[\"']?"

    def V(g):
        return rf"[\"'](?P<{g}>[^\"']+)[\"']"

    def VS(g):
        item = r"[\"'][^\"']+[\"']"
        return rf"(?P<{g}>{item}(?:\s*(?:,|or|and|&)\s*{item})*)"

    WHEN = rf"when\s+(?:the\s+)?{A('a')}(?:\s+attribute)?\s+value\s+is\s+{V('v')}"
    SEP = r"\s*,?\s*(?:then\s+)?(?:the\s+)?"
    flags = re.I
    return {
        "only": re.compile(WHEN + SEP + rf"only\s+{VS('x')}\s+values?\s+(?:shall|can|may)\s+be\s+(?:able\s+to\s+)?select(?:ed)?"
                           rf"(?:\s+(?:from|in|for))?\s+(?:the\s+)?{A('b')}", flags),
        "not": re.compile(WHEN + SEP + rf"{VS('x')}\s+values?\s+shall\s+not\s+be\s+(?:able\s+to\s+)?select(?:ed)?"
                          rf"(?:\s+(?:from|in|for))?\s+(?:the\s+)?{A('b')}", flags),
        "hidden": re.compile(WHEN + SEP + rf"{A('b')}\s+attribute\s+shall\s+be\s+hidden", flags),
        "populated": re.compile(rf"^\s*(?:the\s+)?{A('b')}\s+attribute\s+shall\s+be\s+populated\s+only\s+when\s+(?:the\s+)?"
                                rf"{A('a')}(?:\s+attribute)?\s+value\s+is\s+(?P<neg>not\s+)?{V('v')}", flags),
    }


def _values_in(blob: str) -> list[str]:
    return re.findall(r"[\"']([^\"']+)[\"']", blob)


def _check_attr(cat: Catalog, name: str, issues: list):
    info = cat.get(name)
    if info is None:
        issues.append(f"'{name}' is not an attribute in the workbook")
    elif not info.values:
        issues.append(f"'{info.name}' has no list of values, so it cannot be combined")
    return info


def _canon_values(cat, attr, raw_values, issues) -> list:
    out = []
    for raw in raw_values:
        c = cat.canon(attr, raw)
        if c is None:
            issues.append(f"'{raw}' is not a listed value of '{attr}'")
        else:
            out.append(c)
    return out


def _parse_sentence(text: str, cat: Catalog, mk_id):
    pats = _sentence_patterns(cat)
    if pats is None:
        return []
    t = _norm(text)
    for kind, pat in pats.items():
        m = pat.search(t)
        if not m:
            continue
        issues: list = []
        a_info = _check_attr(cat, m.group("a"), issues)
        b_info = _check_attr(cat, m.group("b"), issues)
        a_name = a_info.name if a_info else m.group("a")
        b_name = b_info.name if b_info else m.group("b")
        v = _canon_values(cat, a_name, [m.group("v")], issues) if a_info and a_info.values else []
        rules = []
        if kind in ("only", "not"):
            xs = _canon_values(cat, b_name, _values_in(m.group("x")), issues) if b_info and b_info.values else []
            if not issues:
                neg = kind == "not"
                rules.append(Rule(mk_id(), "sentence", _desc(Cond(a_name, frozenset(v)), [Cond(b_name, frozenset(xs), neg)]),
                                  (Cond(a_name, frozenset(v)),), (Cond(b_name, frozenset(xs), neg),)))
        elif kind == "hidden":
            if not issues:
                w, th = Cond(a_name, frozenset(v)), Cond(b_name, frozenset({BLANK}))
                rules.append(Rule(mk_id(), "sentence", _desc(w, [th], "hidden"), (w,), (th,)))
        elif kind == "populated":
            neg = bool(m.group("neg"))
            if not issues:
                w = Cond(a_name, frozenset(v), negate=not neg)          # populated only when A is v  =>  else blank
                th = Cond(b_name, frozenset({BLANK}))
                rules.append(Rule(mk_id(), "sentence", _desc(w, [th], "blank"), (w,), (th,)))
                if b_info and b_info.mandatory:                          # mandatory -> must be filled when allowed
                    w2 = Cond(a_name, frozenset(v), negate=neg)
                    th2 = Cond(b_name, frozenset({BLANK}), negate=True)
                    rules.append(Rule(mk_id(), "sentence", _desc(w2, [th2], "filled"), (w2,), (th2,)))
        if issues:                                                      # keep it visible, but switched off
            w = (Cond(a_name, frozenset(v)),)
            r = Rule(mk_id(), "sentence", f"[{kind}] " + t[:140], w, (Cond(b_name, frozenset()),), enabled=False, issues=issues)
            return [r]
        return rules
    return []


def _desc(when: Cond, thens: list, flavour: str = "") -> str:
    tail = {"hidden": " (attribute hidden)", "blank": " (not populated)", "filled": " (must be populated)"}.get(flavour, "")
    return f"IF {when.describe()} THEN " + " AND ".join(c.describe() for c in thens) + tail


# --------------------------------------------------------------------------- tables
def _strip_lov_suffix(h: str) -> str:
    return re.sub(r"\s+(lovs?|values?|list of values)\s*$", "", _squash(h), flags=re.I)


def _parse_lookup(table, cat: Catalog, mk_id):
    """First column = values of attribute A, other columns = attributes whose value A decides."""
    a_info = cat.get(_strip_lov_suffix(table.headers[0]))
    cols = [(j, cat.get(h)) for j, h in enumerate(table.headers) if j > 0]
    if a_info is None or not cols or any(i is None for _, i in cols):
        return None
    rules = []
    for row in table.rows:
        raw_a = _squash(row[0]) if row else ""
        if not raw_a:
            continue
        issues, notes, thens = [], [], []
        a_val = cat.canon(a_info.name, raw_a)
        if a_val is None:
            issues.append(f"'{raw_a}' is not a listed value of '{a_info.name}'")
        for j, b in cols:
            cell = _squash(row[j]) if j < len(row) else ""
            if not cell:
                continue
            if cell.casefold() in NA and not cat.canon(b.name, cell):
                notes.append(f"'{cell}' for {b.name} ignored (not a listed value)")
                continue
            c = cat.canon(b.name, cell)
            if c is None:
                notes.append(f"'{cell}' is not a listed value of '{b.name}' - ignored")
                continue
            thens.append(Cond(b.name, frozenset({c})))
        if not thens and not issues:
            continue
        w = Cond(a_info.name, frozenset({a_val} if a_val else set()))
        text = _desc(w, thens) if a_val else f"IF {a_info.name} is '{raw_a}' ..."
        rules.append(Rule(mk_id(), "table", text, (w,), tuple(thens), enabled=not issues, issues=issues, notes=notes))
    return rules


def _tick_kind(cell: str) -> str:
    c = _squash(cell).casefold()
    if not c:
        return "empty"
    return "tick" if c in TICKS else "other"


def _pick(cat, attr, choices):
    info = cat.get(attr)
    if not info:
        return None
    for v in info.values:
        if _squash(v).casefold() in choices:
            return v
    return None


def _parse_matrix(table, cat: Catalog, block_text: str, mk_id):
    """Columns = values of one attribute P, rows = attributes set by a tick."""
    heads = [_squash(h) for h in table.headers[1:]]
    if len(heads) < 2 or not table.rows:
        return None
    best, best_hits = None, 0
    for info in cat.attrs.values():
        lov = {_squash(v).casefold() for v in info.values}
        hits = sum(1 for h in heads if h.casefold() in lov)
        if hits > best_hits:
            best, best_hits = info, hits
    if best is None or best_hits < 2 or best_hits < 0.6 * len(heads):
        return None
    known_rows = sum(1 for r in table.rows if r and cat.get(r[0]))
    if known_rows < 1:
        return None
    unknown_rows = len(table.rows) - known_rows
    rest_no = bool(re.search(r"\brest\b.*\bno\b", _norm(block_text), re.I))
    rules = []
    for j, head in enumerate(heads, start=1):
        p_val = cat.canon(best.name, head)
        if p_val is None:
            continue
        thens, notes = [], []
        if unknown_rows:
            notes.append(f"{unknown_rows} row(s) are not attributes in the workbook - ignored")
        for row in table.rows:
            b = cat.get(row[0]) if row else None
            if b is None:
                continue
            cell = row[j] if j < len(row) else ""
            kind = _tick_kind(cell)
            if kind == "tick":
                val = _pick(cat, b.name, YES)
                (thens.append(Cond(b.name, frozenset({val}))) if val else notes.append(f"{b.name} has no Yes value"))
            elif kind == "empty" and rest_no:
                val = _pick(cat, b.name, NO)
                (thens.append(Cond(b.name, frozenset({val}))) if val else notes.append(f"{b.name} has no No value"))
        if thens:
            w = Cond(best.name, frozenset({p_val}))
            rules.append(Rule(mk_id(), "table", _desc(w, thens), (w,), tuple(thens), notes=notes[:5]))
    return rules


# ---------------------------------------------------------------------------- driver
@dataclass
class TextOnly:
    text: str
    reason: str


def detect_rules(blocks, cat: Catalog):
    """blocks: BusinessRuleBlock list (from extract_business_rule_blocks).
    Returns (rules, text_only)."""
    counter = [0]

    def mk_id():
        counter[0] += 1
        return f"R{counter[0]:02d}"

    rules: list[Rule] = []
    text_only: list[TextOnly] = []
    for blk in blocks:
        text = _norm(blk.text)
        got = _parse_sentence(blk.text, cat, mk_id)
        for tbl in blk.tables:
            if not tbl.headers or not tbl.rows:
                continue
            parsed = _parse_lookup(tbl, cat, mk_id)
            if parsed is None:
                parsed = _parse_matrix(tbl, cat, blk.text, mk_id)
            if parsed:
                got += parsed
            else:
                text_only.append(TextOnly(f"[table] {' | '.join(tbl.headers[:6])}", "Table layout not recognised"))
        if got:
            rules += got
            continue
        if not text:
            continue
        if re.search(r"product component .*shall be (populated|restricted)", text, re.I):
            reason = "Component-level rule (whole component present or absent) - design these tests by hand"
        elif blk.tables:
            reason = "Described by the table above"
        else:
            reason = "Free-text requirement - not machine-readable; add a rule below if it should be enforced"
        text_only.append(TextOnly(text, reason))
    # keep ids tidy: numbering is already sequential
    return rules, text_only


def manual_rule(idx: int, when: Cond, then: Cond) -> Rule:
    return Rule(f"M{idx:02d}", "manual", _desc(when, [then]), (when,), (then,))
