"""Build the test plan: choose factors, apply rules, run the engine, label every case."""
from __future__ import annotations

import re
from dataclasses import dataclass, field

from tools.testx.engine import Engine, TooManyCombinations, bind_rules, blank_needed, full_size
from tools.testx.model import BLANK, BLANK_LABEL, Factor, Group, Rule, TestCase
from tools.testx.rules import _mandatory_flag

ORDER_KEY = "Order Type"


@dataclass
class Settings:
    technique: str = "pairwise"          # "pairwise" | "full"
    scope: str = "component"             # "component" | "together"
    components: list = field(default_factory=list)     # node names; empty = all
    include_blank: bool = True           # optional attributes may stay blank
    negatives: bool = True               # one "must be refused" test per rule
    vary_order_type: bool = False        # add Order Type (sheet 5.1) as one more attribute
    full_limit: int = 5000


@dataclass
class Plan:
    product: str
    settings: Settings
    groups: list
    rules: list
    warnings: list = field(default_factory=list)

    @property
    def positive(self) -> int:
        return sum(1 for g in self.groups for c in g.cases if c.kind == "Positive")

    @property
    def negative(self) -> int:
        return sum(1 for g in self.groups for c in g.cases if c.kind == "Negative")

    @property
    def full_total(self) -> int:
        return sum(g.full_size for g in self.groups)


def _factors_for(nodes, include_blank: bool, qualify: bool):
    """(factors, free_entry) for one or more nodes. Attributes without values cannot be combined."""
    raw, free = [], []
    for n in nodes:
        for a in n.attributes:
            mand = _mandatory_flag(a.mandatory)
            if not a.values:
                free.append((a.entity_name, n.name, mand is not False))
                continue
            vals = list(a.values)
            if include_blank and mand is False:
                vals.append(BLANK)
            raw.append((n.name, a.entity_name, tuple(vals), mand is not False))
    counts: dict[str, int] = {}
    for _, name, _, _ in raw:
        counts[name.casefold()] = counts.get(name.casefold(), 0) + 1
    factors = []
    for node, name, vals, mand in raw:
        key = f"{node} › {name}" if qualify and counts[name.casefold()] > 1 else name
        factors.append(Factor(key, name, node, vals, mand))
    return factors, free


def _expected_positive(fired: list[str], by_id: dict) -> str:
    if not fired:
        return "The configuration is accepted exactly as shown."
    return "The configuration is accepted and satisfies: " + "; ".join(f"{i}" for i in fired) + "."


def _expected_negative(rule: Rule, key: str, bad, broken: list[str]) -> str:
    shown = BLANK_LABEL if bad is BLANK else bad
    extra = [i for i in broken if i != rule.id]
    also = f" (this value also conflicts with {', '.join(extra)})" if extra else ""
    txt = rule.text if len(rule.text) <= 200 else rule.text[:197] + "..."
    return f"The system must refuse or prevent {key} = {shown}{also}. {rule.id}: {txt}"


def generate(source, rules: list[Rule], s: Settings, product: str) -> Plan:
    warnings: list[str] = []
    nodes = [n for n in source.nodes if not s.components or n.name in s.components]
    groups: list[Group] = []
    scopes = [(n.name, [n]) for n in nodes] if s.scope == "component" else [("All components", nodes)]
    by_id = {r.id: r for r in rules}

    for gi, (gname, gnodes) in enumerate(scopes, start=1):
        factors, free = _factors_for(gnodes, s.include_blank, qualify=(s.scope == "together"))
        if s.vary_order_type and source.order_types:
            factors.append(Factor(ORDER_KEY, ORDER_KEY, "(order)", tuple(source.order_types), True))
        if len(factors) < 1:
            warnings.append(f"'{gname}' has no attribute with a list of values - nothing to combine, sheet skipped.")
            continue
        bound, skipped = bind_rules(factors, rules)
        need = blank_needed(bound)
        factors = [Factor(f.key, f.name, f.node, f.values + ((BLANK,) if f.key in need and BLANK not in f.values else ()), f.mandatory)
                   for f in factors]
        eng = Engine(factors, bound)
        g = Group(gname, factors, [], free_entry=free, full_size=full_size(factors),
                  rules_applied=[b.rule.id for b in bound])
        if skipped:
            ids = ", ".join(i for i, _ in skipped[:4]) + (" ..." if len(skipped) > 4 else "")
            g.notes.append(f"{len(skipped)} rule(s) not used on this sheet because they involve attributes that are not here ({ids}). "
                           "Choose 'All components together' to test rules that link components.")
        try:
            if s.technique == "full":
                rows = eng.full(s.full_limit)
                g.pairs_total = g.pairs_unreachable = 0
            else:
                rows, g.pairs_total, g.pairs_unreachable = eng.pairwise()
        except TooManyCombinations as exc:
            raise TooManyCombinations(exc.limit) from None
        if g.pairs_unreachable:
            g.notes.append(f"{g.pairs_unreachable} value pair(s) can never occur together because of the rules.")
        for n, row in enumerate(rows, start=1):
            fired = eng.check.fired(row)
            g.cases.append(TestCase(f"TC-{gi:02d}-{n:03d}", "Positive", row, fired, _expected_positive(fired, by_id)))
        if s.negatives and bound:
            negs, why_not = eng.negatives()
            for n, (b, case, key, bad, broken) in enumerate(negs, start=1):
                g.cases.append(TestCase(f"NT-{gi:02d}-{n:03d}", "Negative", case, [b.rule.id],
                                        _expected_negative(b.rule, key, bad, broken), highlight=key))
            if why_not:
                g.notes.append("No negative test for: " + "; ".join(f"{i} ({w})" for i, w in why_not[:5]))
        groups.append(g)
    return Plan(product=product, settings=s, groups=groups, rules=rules, warnings=warnings)


def safe_sheet_name(name: str, used: set) -> str:
    base = re.sub(r"[\[\]\:\*\?/\\]", "_", name).strip()[:28] or "Sheet"
    cand, i = base, 2
    while cand.casefold() in used:
        cand = f"{base[:25]}_{i}"
        i += 1
    used.add(cand.casefold())
    return cand
