"""Extractor for sheet ``3. Special Requirements``.

The source sheet can contain normal requirement text plus one or more tables
belonging to particular requirements.  The previous implementation treated
an entire rectangular range as one table, which incorrectly rendered all
requirements as a single Word table.

This extractor now returns requirement blocks in source order:
    requirement text -> optional table(s) belonging to that requirement

Structured Excel tables are preferred when present.  A conservative
fallback detects rectangular table-like blocks when the source uses ordinary
cell ranges instead of Excel table objects.
"""
from __future__ import annotations

import openpyxl
from openpyxl.utils.cell import range_boundaries, get_column_letter
from extractors.sheet_utils import find_sheet
from models.schemas import ExcelTableData, BusinessRuleBlock

TARGET_COLUMN = "Requirement Description"
MAX_HEADER_SEARCH_ROWS = 100


def _norm(value) -> str:
    if value is None:
        return ""
    return " ".join(str(value).strip().split()).lower()


def _display(value) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _find_target_header(ws):
    for r in range(1, min(MAX_HEADER_SEARCH_ROWS, ws.max_row) + 1):
        for c in range(1, ws.max_column + 1):
            if _norm(ws.cell(r, c).value) == TARGET_COLUMN.lower():
                return r, c
    return None


def _structured_table_for_target(ws, target_header_row: int, target_col: int):
    for table in ws.tables.values():
        min_col, min_row, max_col, max_row = range_boundaries(table.ref)
        if min_row <= target_header_row <= max_row and min_col <= target_col <= max_col:
            return table.ref
    return None


def _capture_range(ws, ref: str, source_sheet: str) -> ExcelTableData:
    min_col, min_row, max_col, max_row = range_boundaries(ref)
    headers = [_display(ws.cell(min_row, c).value) for c in range(min_col, max_col + 1)]
    rows = []
    for r in range(min_row + 1, max_row + 1):
        rows.append([_display(ws.cell(r, c).value) for c in range(min_col, max_col + 1)])
    return ExcelTableData(
        headers=headers,
        rows=rows,
        source_sheet=source_sheet,
        source_range=ref,
    )


def _row_nonempty_cells(ws, row: int, min_col: int = 1, max_col: int | None = None):
    max_col = max_col or ws.max_column
    return [c for c in range(min_col, max_col + 1) if _display(ws.cell(row, c).value)]


def _looks_like_requirement_text(text: str) -> bool:
    """Conservative filter so table headers/cells in the same column are not
    mistaken for new requirement paragraphs."""
    t = _display(text)
    if not t:
        return False
    if len(t) >= 30:
        return True
    # Short requirements are still accepted when they read like a sentence.
    return t.endswith((".", ":", ";")) and len(t.split()) >= 4


def _hmerges_in_row(ws, row: int):
    """Horizontal merged ranges (span > 1 column) that cover ``row``."""
    return [m for m in ws.merged_cells.ranges
            if m.min_row <= row <= m.max_row and m.max_col > m.min_col]


def _capture_rectangular_rows(ws, start_row: int, end_row: int, min_col: int, max_col: int,
                              source_sheet: str, header_levels: int = 1):
    """Capture an ordinary-cell table.

    ``header_levels`` == 2 means the header is two rows: a group row whose
    cells are merged across several columns (e.g. 'Attribute values of
    "Package Type"') above a row holding the real column labels.
    """
    ref = f"{get_column_letter(min_col)}{start_row}:{get_column_letter(max_col)}{end_row}"
    if header_levels == 2:
        merges = {m.min_col: m for m in _hmerges_in_row(ws, start_row)}
        groups: list[list] = []
        headers: list[str] = []
        c = min_col
        while c <= max_col:
            m = merges.get(c)
            if m is not None:
                span = min(m.max_col, max_col) - c + 1
                groups.append([_display(ws.cell(start_row, c).value), span])
                for cc in range(c, c + span):
                    headers.append(_display(ws.cell(start_row + 1, cc).value))
                c += span
            else:
                # Not under a group: the label sits in the top row
                # (usually vertically merged with the row below).
                label = _display(ws.cell(start_row, c).value) or _display(ws.cell(start_row + 1, c).value)
                groups.append(["", 1])
                headers.append(label)
                c += 1
        first_data = start_row + 2
    else:
        groups = []
        headers = [_display(ws.cell(start_row, c).value) for c in range(min_col, max_col + 1)]
        first_data = start_row + 1
    rows = []
    for r in range(first_data, end_row + 1):
        rows.append([_display(ws.cell(r, c).value) for c in range(min_col, max_col + 1)])
    return ExcelTableData(
        headers=headers,
        rows=rows,
        group_headers=groups,
        source_sheet=source_sheet,
        source_range=ref,
    )


def _looks_like_header_row(ws, row: int, min_col: int, max_col: int) -> bool:
    values = [_display(ws.cell(row, c).value) for c in range(min_col, max_col + 1)]
    nonempty = [v for v in values if v]
    if len(nonempty) < 2:
        return False
    # Requirement paragraphs are long free text and should not become table
    # headers. Table headers are normally short labels.
    if any(len(v) > 40 for v in nonempty):
        return False
    # A row containing only checkmarks/numeric values is not a header.
    alpha = sum(any(ch.isalpha() for ch in v) for v in nonempty)
    return alpha >= 2


def _structured_tables_by_row(ws, exclude_ref: str | None):
    result = []
    for table in ws.tables.values():
        if exclude_ref and table.ref == exclude_ref:
            continue
        min_col, min_row, max_col, max_row = range_boundaries(table.ref)
        result.append((min_row, max_row, min_col, max_col, table.ref))
    return sorted(result)


def _fallback_table_blocks(ws, start_row: int, end_row: int, main_ref: str | None, target_col: int | None = None):
    """Find ordinary-cell table blocks in the rows after a requirement.

    A candidate begins at a row with at least two non-empty cells and a
    header-like pattern, then continues while cells occur inside the same
    column span.  Completely blank spacer rows are tolerated once inside a
    block.  Main requirement-table rows are excluded when the workbook uses
    a structured Excel table for the requirements list.
    """
    main_bounds = range_boundaries(main_ref) if main_ref else None
    blocks = []
    r = start_row
    while r <= end_row:
        # Skip rows belonging to the main requirement table.
        if main_bounds:
            _, mmin_row, _, mmax_row = main_bounds
            if mmin_row <= r <= mmax_row:
                r = mmax_row + 1
                continue

        cells = _row_nonempty_cells(ws, r)
        if len(cells) < 2:
            r += 1
            continue

        # A row that carries text in the 'Requirement Description' column is a
        # requirement row (e.g. R10 / R11), never a table header. Without
        # this guard two consecutive short requirements were mistaken for a
        # 'table' and R11 was rendered as a table instead of a paragraph.
        if (target_col is not None and _display(ws.cell(r, target_col).value)
                and min(cells) < target_col):
            r += 1
            continue

        min_col, max_col = min(cells), max(cells)
        # Horizontally merged header cells (e.g. F19:L19) extend the table
        # width beyond the last non-empty cell of the header row.
        row_merges = _hmerges_in_row(ws, r)
        for m in row_merges:
            if m.min_col >= min_col and m.min_col <= max_col:
                max_col = max(max_col, m.max_col)
        if not _looks_like_header_row(ws, r, min_col, max_col):
            r += 1
            continue

        header_levels = 1
        for m in row_merges:
            if m.min_col < min_col or m.max_col > max_col:
                continue
            below = [c for c in range(m.min_col, m.max_col + 1) if _display(ws.cell(r + 1, c).value)]
            if len(below) >= 2:
                header_levels = 2
                break

        block_start = r
        block_end = r
        last_nonempty = r
        blank_run = 0
        rr = r + 1
        while rr <= end_row:
            if main_bounds:
                _, mmin_row, _, mmax_row = main_bounds
                if mmin_row <= rr <= mmax_row:
                    break

            vals = [ws.cell(rr, c).value for c in range(min_col, max_col + 1)]
            has_any = any(_display(v) for v in vals)
            if has_any and target_col is not None:
                target_text = _display(ws.cell(rr, target_col).value)
                if _looks_like_requirement_text(target_text) and len(_row_nonempty_cells(ws, rr)) <= 2:
                    break
            if has_any:
                # Do not absorb a later requirement paragraph into the table.
                target_col_value = _display(ws.cell(rr, 1).value)
                if target_col_value and _looks_like_header_row(ws, rr, min_col, max_col):
                    # still a possible table row; continue
                    pass
                blank_run = 0
                last_nonempty = rr
                block_end = rr
                rr += 1
                continue

            blank_run += 1
            if blank_run >= 2:
                break
            rr += 1

        if block_end > block_start:
            blocks.append((block_start, block_end, min_col, max_col, header_levels))
            r = block_end + 1
        else:
            r += 1

    return blocks


def extract_business_rule_blocks(workbook_path: str, sheet_name: str = "3. Special Requirements") -> list[BusinessRuleBlock]:
    wb = openpyxl.load_workbook(workbook_path, data_only=True)
    resolved = find_sheet(wb, sheet_name)
    ws = wb[resolved]

    target = _find_target_header(ws)
    if target is None:
        raise ValueError(f"Could not find a column named '{TARGET_COLUMN}' in '{resolved}'.")

    header_row, target_col = target
    main_ref = _structured_table_for_target(ws, header_row, target_col)
    main_bounds = range_boundaries(main_ref) if main_ref else None

    # Requirement rows are identified from the Requirement Description column,
    # excluding the header and any rows belonging to an embedded table.
    requirement_rows = []
    if main_bounds:
        # The normal requirement list itself may be an Excel structured table.
        # In that case the requirement rows are INSIDE that table; only the
        # later table-detection logic must exclude those rows.
        scan_start = main_bounds[1] + 1
        scan_end = main_bounds[3]
    else:
        scan_start = header_row + 1
        scan_end = ws.max_row

    for r in range(scan_start, scan_end + 1):
        text = _display(ws.cell(r, target_col).value)
        if _looks_like_requirement_text(text):
            requirement_rows.append((r, text))

    if not requirement_rows:
        return []

    blocks: list[BusinessRuleBlock] = []
    structured_tables = _structured_tables_by_row(ws, main_ref)
    global_fallback_blocks = _fallback_table_blocks(
        ws, header_row + 1, ws.max_row, main_ref, target_col
    )
    embedded_rows = set()
    for b_start, b_end, _min_col, _max_col, _lv in global_fallback_blocks:
        embedded_rows.update(range(b_start, b_end + 1))

    # Re-scan now that embedded table rows are known.
    requirement_rows = []
    for r in range(scan_start, scan_end + 1):
        text = _display(ws.cell(r, target_col).value)
        if r in embedded_rows or not text:
            continue
        if main_bounds or _looks_like_requirement_text(text):
            requirement_rows.append((r, text))

    for idx, (req_row, text) in enumerate(requirement_rows):
        next_req_row = requirement_rows[idx + 1][0] if idx + 1 < len(requirement_rows) else ws.max_row + 1
        tables: list[ExcelTableData] = []

        # 1) Structured Excel tables located between this requirement and the next.
        for min_row, max_row, min_col, max_col, ref in structured_tables:
            if req_row < min_row < next_req_row:
                tables.append(_capture_range(ws, ref, resolved))

        # 2) Ordinary rectangular cell ranges when no structured table was found.
        if not tables:
            for b_start, b_end, min_col, max_col, levels in global_fallback_blocks:
                if req_row < b_start < next_req_row:
                    tables.append(_capture_rectangular_rows(
                        ws, b_start, b_end, min_col, max_col, resolved, levels
                    ))

        blocks.append(BusinessRuleBlock(text=text, tables=tables, source_row=req_row))

    return blocks


def extract_business_rules_and_table(workbook_path: str, sheet_name: str = "3. Special Requirements"):
    blocks = extract_business_rule_blocks(workbook_path, sheet_name)
    rules = [b.text for b in blocks if b.text]
    # Backward compatibility: no longer return the entire requirement table
    # because it was the source of the unwanted all-in-one table rendering.
    return rules, None


def extract_business_rules(workbook_path: str, sheet_name: str = "3. Special Requirements") -> list[str]:
    return [b.text for b in extract_business_rule_blocks(workbook_path, sheet_name)]


def extract_business_rules_table(workbook_path: str, sheet_name: str = "3. Special Requirements") -> ExcelTableData | None:
    # Kept for callers that still import this function. The main requirement
    # table is deliberately not exposed as a single renderable table.
    return None
