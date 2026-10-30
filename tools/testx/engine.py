"""TestX combinatorial engine. Pure Python - no Streamlit, no Excel.

Two techniques:
  * pairwise - every pair of values of any two attributes appears in at least one test
    (typically 10-50x fewer tests than all combinations, and the usual UAT choice)
  * full     - every valid combination (refuses when there are too many)

Business rules are constraints: a combination that breaks an enabled rule is never produced.
The result is deterministic - same input, same test plan - so UAT runs are repeatable.
"""
from __future__ import annotations

from dataclasses import dataclass
from itertools import combinations
from math import prod

from tools.testx.model import BLANK, Cond, Factor, Rule


class TooManyCombinations(Exception):
    def __init__(self, limit: int):
        super().__init__(f"More than {limit:,} valid combinations")
        self.limit = limit


class _Budget(Exception):
    pass


@dataclass
class BoundRule:
    """A Rule whose attribute names are resolved to factor keys of ONE sheet."""
    rule: Rule
    when: tuple          # ((key, Cond), ...)
    then: tuple


# --------------------------------------------------------------------------- rules
def bind_rules(factors: list[Factor], rules: list[Rule]):
    """Resolve rule attribute names to this sheet's factors.

    Returns (bound_rules, skipped) where skipped is [(rule_id, reason)]."""
    by_name: dict[str, list[str]] = {}
    for f in factors:
        by_name.setdefault(f.name.strip().casefold(), []).append(f.key)
    bound, skipped = [], []
    for r in rules:
        if not r.enabled:
            continue
        w, t, why = [], [], None
        for c in r.when:
            keys = by_name.get(c.attr.strip().casefold(), [])
            if len(keys) != 1:
                why = f"'{c.attr}' " + ("is not in this sheet" if not keys else "appears more than once here")
                break
            w.append((keys[0], c))
        if why is None:
            for c in r.then:
                keys = by_name.get(c.attr.strip().casefold(), [])
                if len(keys) == 1:
                    t.append((keys[0], c))
            if not t:
                why = "none of the attributes it controls are in this sheet"
        if why:
            skipped.append((r.id, why))
        else:
            bound.append(BoundRule(r, tuple(w), tuple(t)))
    return bound, skipped


def blank_needed(bound: list[BoundRule]) -> set[str]:
    """Factor keys that must be allowed to stay blank because a rule says so (hidden / not populated)."""
    need = set()
    for b in bound:
        for k, c in b.then:
            if BLANK in c.values and not c.negate:
                need.add(k)
    return need


class _Checker:
    """Answers: does this (partial) test case break a rule?"""

    def __init__(self, bound: list[BoundRule]):
        self.bound = bound
        self.by_key: dict[str, list[BoundRule]] = {}
        for b in bound:
            for k in {k for k, _ in b.when} | {k for k, _ in b.then}:
                self.by_key.setdefault(k, []).append(b)

    @staticmethod
    def _broken(b: BoundRule, case: dict) -> bool:
        for k, c in b.when:
            if k not in case:
                return False                    # not decidable yet
            if not c.holds(case[k]):
                return False                    # rule does not apply
        for k, c in b.then:
            if k in case and not c.holds(case[k]):
                return True
        return False

    def ok_after(self, case: dict, key: str, value) -> bool:
        case[key] = value
        try:
            return not any(self._broken(b, case) for b in self.by_key.get(key, ()))
        finally:
            del case[key]

    def consistent(self, case: dict) -> bool:
        return not any(self._broken(b, case) for b in self.bound)

    def broken_rules(self, case: dict) -> list[str]:
        return [b.rule.id for b in self.bound if self._broken(b, case)]

    def fired(self, case: dict) -> list[str]:
        out = []
        for b in self.bound:
            if all(k in case and c.holds(case[k]) for k, c in b.when):
                out.append(b.rule.id)
        return out


# --------------------------------------------------------------------------- engine
class Engine:
    def __init__(self, factors: list[Factor], bound: list[BoundRule], max_nodes: int = 20000):
        self.factors = factors
        self.keys = [f.key for f in factors]
        self.dom = {f.key: list(f.values) for f in factors}
        self.check = _Checker(bound)
        self.max_nodes = max_nodes

    # -- completing a partial case (backtracking) ---------------------------
    def complete(self, partial: dict, gain=None):
        case = dict(partial)
        if not self.check.consistent(case):
            return None
        nodes = [0]

        def rec(remaining):
            if not remaining:
                return dict(case)
            nodes[0] += 1
            if nodes[0] > self.max_nodes:
                raise _Budget()
            best = None
            for k in remaining:                  # most-constrained factor first
                feas = [v for v in self.dom[k] if self.check.ok_after(case, k, v)]
                if not feas:
                    return None
                if best is None or len(feas) < len(best[1]):
                    best = (k, feas)
            k, feas = best
            if gain:
                feas.sort(key=lambda v: -gain(case, k, v))   # stable: ties keep list order
            rest = [x for x in remaining if x != k]
            for v in feas:
                case[k] = v
                r = rec(rest)
                if r is not None:
                    return r
                del case[k]
            return None

        return rec([k for k in self.keys if k not in case])

    # -- pairwise ------------------------------------------------------------
    def pairwise(self):
        """Returns (cases, pairs_total, pairs_unreachable)."""
        idx = {k: i for i, k in enumerate(self.keys)}
        uncovered: dict = {}
        for a, b in combinations(self.keys, 2):
            for va in self.dom[a]:
                for vb in self.dom[b]:
                    uncovered[(a, va, b, vb)] = None
        total = len(uncovered)

        def pair(ka, va, kb, vb):
            return (ka, va, kb, vb) if idx[ka] < idx[kb] else (kb, vb, ka, va)

        def gain(case, k, v):
            return sum(1 for ok, ov in case.items() if pair(k, v, ok, ov) in uncovered)

        cases, unreachable = [], 0
        order = list(uncovered)                  # deterministic seed order
        pos = 0
        while uncovered:
            while order[pos] not in uncovered:   # skip pairs already covered
                pos += 1
            seed = order[pos]
            try:
                case = self.complete({seed[0]: seed[1], seed[2]: seed[3]}, gain)
            except _Budget:
                case = None
            if case is None:
                del uncovered[seed]
                unreachable += 1
                continue
            cases.append(case)
            items = [(k, case[k]) for k in self.keys]
            for (ka, va), (kb, vb) in combinations(items, 2):
                uncovered.pop((ka, va, kb, vb), None)
        return cases, total, unreachable

    # -- full ----------------------------------------------------------------
    def full(self, limit: int):
        out: list[dict] = []
        case: dict = {}

        def rec(i):
            if i == len(self.keys):
                out.append(dict(case))
                if len(out) > limit:
                    raise TooManyCombinations(limit)
                return
            k = self.keys[i]
            for v in self.dom[k]:
                if self.check.ok_after(case, k, v):
                    case[k] = v
                    rec(i + 1)
                    del case[k]

        rec(0)
        return out

    # -- negative cases ------------------------------------------------------
    def negatives(self):
        """One 'must be refused' case per rule: a valid case where the rule's IF holds but one
        THEN value is replaced by a forbidden one (preferably breaking only that rule).
        Returns [(BoundRule, case, key, bad_value, broken_rule_ids)] and a list of
        (rule_id, reason) for rules that could not produce one."""
        out, why_not = [], []
        for b in self.check.bound:
            base = self._base_for(b)
            if base is None:
                why_not.append((b.rule.id, "its IF part can never be true together with the other rules"))
                continue
            best = None                                   # (number of rules broken, case, key, bad value, broken ids)
            for k, c in b.then:
                for bad in self.dom[k]:
                    if c.holds(bad):
                        continue
                    trial = dict(base)
                    trial[k] = bad
                    broken = self.check.broken_rules(trial)
                    if b.rule.id not in broken:
                        continue
                    if best is None or len(broken) < best[0]:
                        best = (len(broken), trial, k, bad, broken)
                    if len(broken) == 1:
                        break
                if best and best[0] == 1:
                    break
            if best:
                out.append((b, best[1], best[2], best[3], best[4]))
            else:
                why_not.append((b.rule.id, "no forbidden value exists to try"))
        return out, why_not

    def _base_for(self, b: BoundRule):
        options = []
        for k, c in b.when:
            vals = [v for v in self.dom[k] if c.holds(v)]
            if not vals:
                return None
            options.append((k, vals))
        # try the first value of each condition, then alternatives one at a time
        tries = [{k: vals[0] for k, vals in options}]
        for k, vals in options:
            for v in vals[1:4]:
                t = {kk: vv[0] for kk, vv in options}
                t[k] = v
                tries.append(t)
        for t in tries:
            try:
                case = self.complete(t)
            except _Budget:
                case = None
            if case is not None:
                return case
        return None


def full_size(factors: list[Factor]) -> int:
    return prod(len(f.values) for f in factors) if factors else 0
