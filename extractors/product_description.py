"""
Extractor for sheet "1. Product Description".

ASSUMED STRUCTURE (not yet validated against a real sheet - only sheets
6.3/6.4 have been seen so far): a sheet whose non-empty cells, read
top-to-bottom / left-to-right, form the narrative "Introduction" text as
a sequence of paragraphs. If the real sheet has a different layout
(e.g. a labeled key-value table, or a single big merged text cell),
this needs a small adjustment - flagged clearly so it's easy to spot
when validating against the real file.
"""
from __future__ import annotations
import openpyxl
from extractors.sheet_utils import find_sheet


def extract_product_description(workbook_path: str, sheet_name: str = "1. Product Description") -> list[str]:
    """Returns a list of paragraph strings, in sheet order, skipping blanks."""
    wb = openpyxl.load_workbook(workbook_path, data_only=True)
    resolved = find_sheet(wb, sheet_name)
    ws = wb[resolved]

    paragraphs: list[str] = []
    for row in ws.iter_rows(values_only=True):
        for val in row:
            if val is not None and str(val).strip():
                paragraphs.append(str(val).strip())
    return paragraphs
