"""Read what TestX needs from the product configuration workbook.

It re-uses the existing GuideX extractors, so both tools read the workbook in exactly the same
way (same sheet-name matching, same merged-cell handling, same fixes for messy real data).
"""
from __future__ import annotations

import tempfile
from dataclasses import dataclass, field
from pathlib import Path

from extractors.order_scenarios import extract_order_scenarios
from extractors.product_hierarchy import extract_hierarchy
from extractors.special_requirements import extract_business_rule_blocks


@dataclass
class Source:
    nodes: list = field(default_factory=list)          # ProductComponent list (sheet 2)
    blocks: list = field(default_factory=list)         # BusinessRuleBlock list (sheet 3)
    order_types: list = field(default_factory=list)    # unique Order Type values (sheet 5.1)
    warnings: list = field(default_factory=list)

    @property
    def product_name(self) -> str:
        for n in self.nodes:
            if n.kind.strip().lower() == "product":
                return n.name
        return self.nodes[0].name if self.nodes else ""

    @property
    def attribute_count(self) -> int:
        return sum(len(n.attributes) for n in self.nodes)

    @property
    def value_count(self) -> int:
        return sum(len(a.values) for n in self.nodes for a in n.attributes)


def read_source(excel_bytes: bytes) -> Source:
    src = Source()
    with tempfile.TemporaryDirectory(prefix="testx_") as td:
        path = str(Path(td) / "input.xlsx")
        Path(path).write_bytes(excel_bytes)

        def attempt(label, fn, default):
            try:
                return fn()
            except ValueError as exc:                       # sheet missing / columns not found
                src.warnings.append(f"[{label}] {exc}")
                return default

        src.nodes = attempt("2. Product Hierarchy", lambda: extract_hierarchy(path), [])
        src.blocks = attempt("3. Special Requirements", lambda: extract_business_rule_blocks(path), [])
        scenarios = attempt("5.1 MDM_Order Scenarios", lambda: extract_order_scenarios(path), [])
    seen = []
    for row in scenarios:
        ot = (row.get("oss_service_order") or "").strip()
        if ot and ot not in seen:
            seen.append(ot)
    src.order_types = seen
    return src
