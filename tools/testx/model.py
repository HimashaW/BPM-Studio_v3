"""Plain data types shared by the TestX modules (no Streamlit, no Excel)."""
from __future__ import annotations

from dataclasses import dataclass, field

BLANK = None                 # "attribute not set / hidden" inside a test case
BLANK_LABEL = "(blank)"


@dataclass(frozen=True)
class Cond:
    """One test on one attribute: its value is (or is not) one of `values`.

    `attr` is the attribute NAME as written in the workbook. `None` inside `values`
    means "left blank / hidden"."""
    attr: str
    values: frozenset
    negate: bool = False

    def holds(self, value) -> bool:
        return (value in self.values) != self.negate

    def describe(self) -> str:
        shown = " or ".join(BLANK_LABEL if v is None else str(v) for v in sorted(self.values, key=lambda x: (x is None, str(x))))
        return f"{self.attr} {'is not' if self.negate else 'is'} {shown}"


@dataclass
class Rule:
    """IF all `when` conditions hold THEN all `then` conditions must hold."""
    id: str
    source: str                          # "table" | "sentence" | "manual"
    text: str                            # readable summary shown to the user
    when: tuple
    then: tuple
    enabled: bool = True                 # the "Use" tick box
    issues: list = field(default_factory=list)   # blocking problems -> rule is switched off
    notes: list = field(default_factory=list)    # non-blocking remarks (a cell was ignored ...)

    @property
    def status(self) -> str:
        return "Needs review" if self.issues else "Ready"


@dataclass(frozen=True)
class Factor:
    """One column of the test matrix: an attribute with a list of values."""
    key: str                             # unique column name inside one sheet
    name: str                            # attribute name as in the workbook
    node: str                            # product / component it belongs to
    values: tuple                        # allowed values (None = blank)
    mandatory: bool = True


@dataclass
class TestCase:
    id: str
    kind: str                            # "Positive" | "Negative"
    values: dict                         # factor key -> value
    rules: list = field(default_factory=list)    # ids of rules this case exercises / breaks
    expected: str = ""
    highlight: str | None = None         # factor key to colour (the deliberately wrong value of a negative test)


@dataclass
class Group:
    """One generated sheet: a component (or the whole product) with its factors and cases."""
    name: str
    factors: list
    cases: list
    free_entry: list = field(default_factory=list)      # attributes without a list of values
    full_size: int = 0                                  # combinations without any technique
    pairs_total: int = 0
    pairs_unreachable: int = 0
    rules_applied: list = field(default_factory=list)   # rule ids
    notes: list = field(default_factory=list)
