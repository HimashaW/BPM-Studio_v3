"""
Sanity checks run between extraction and rendering. Each check appends a
human-readable warning to result.warnings rather than crashing, EXCEPT
for checks that would make the output actively misleading (those raise).
"""
from __future__ import annotations
from models.schemas import ExtractionResult

# Order types are discovered from the supplied workbook. Do not hardcode
# a product-specific allow-list here; the extractor and structural checks
# validate the source sections generically.


def run_all_checks(result: ExtractionResult) -> ExtractionResult:
    _check_no_empty_sections(result)
    _check_tasks_have_names(result)
    _check_duplicate_order_types(result)
    return result


def _check_no_empty_sections(result: ExtractionResult) -> None:
    for sec in result.order_types:
        if sec.task_count == 0:
            result.warnings.append(
                f"Order type '{sec.order_type}' (rows {sec.source_rows}) has "
                f"ZERO tasks extracted - likely a parsing issue, check manually."
            )


def _check_tasks_have_names(result: ExtractionResult) -> None:
    for sec in result.order_types:
        for t in sec.tasks:
            if not t.task_name or not t.task_name.strip():
                result.warnings.append(
                    f"Blank task_name in '{sec.order_type}' at "
                    f"{t.source_sheet}!row{t.source_row}"
                )


def _check_duplicate_order_types(result: ExtractionResult) -> None:
    seen = {}
    for sec in result.order_types:
        seen.setdefault(sec.order_type, []).append(sec.source_rows)
    for order_type, rows_list in seen.items():
        if len(rows_list) > 1:
            result.warnings.append(
                f"Order type '{order_type}' appears in multiple sections: "
                f"{rows_list}. Check for duplicate/leftover template blocks."
            )
