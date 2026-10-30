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
from extractors.stacked_tables import extract_order_types
from tools.testx.facts import Extra, read_extra


@dataclass
class Source:
    nodes: list = field(default_factory=list)          # ProductComponent list (sheet 2)
    blocks: list = field(default_factory=list)         # BusinessRuleBlock list (sheet 3)
    order_types: list = field(default_factory=list)    # unique Order Type values (sheet 5.1)
    sections: list = field(default_factory=list)       # OrderTypeSection list (task sheet: 6.2 / 6.3 "SO Task List")
    extra: Extra = field(default_factory=Extra)        # hierarchy mapping, pricing, faults, 5.1/5.2 rows, requirements
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
        src.extra = attempt("Extra sheets", lambda: read_extra(path, src.nodes), Extra())
        if src.extra.task_sheet:
            src.sections = attempt(src.extra.task_sheet.strip(), lambda: extract_order_types(
                path, sheet_name=src.extra.task_sheet, product=src.product_name or "Product", warnings=[]), [])
        else:
            src.warnings.append("[Task list] No sheet named like '... SO Task List' was found, so no order-flow scenarios can be built.")
    if src.extra.order_rows:                       # 5.1 was read by the tolerant reader, the strict one's warning is noise
        src.warnings = [w for w in src.warnings if not w.startswith("[5.1 MDM_Order Scenarios]")]
    seen = []
    for row in scenarios:
        ot = (row.get("oss_service_order") or "").strip()
        if ot and ot not in seen:
            seen.append(ot)
    if not seen:
        for o in src.extra.order_rows:
            if o.order_type not in seen:
                seen.append(o.order_type)
    src.order_types = seen
    return src
