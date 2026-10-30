"""
Template placeholder substitution.

Word frequently splits what looks like one continuous piece of text into
several internal <w:r> "runs" (an artifact of how the file was edited),
even within a single word. A naive per-run string replace would silently
miss a placeholder like "{{PRODUCT_NAME}}" if Word happened to split it
as "{{PRO" + "DUCT_NAME}}" across two runs. This module works around
that by replacing at the whole-paragraph level instead.

Trade-off: replacing a placeholder this way collapses the paragraph onto
its first run's formatting (font, size, color, bold, etc.), so anything
AFTER the placeholder in that same paragraph inherits the first run's
style rather than whatever style it originally had. In practice this is
rarely visible for short placeholder tokens sitting inside one span of
consistent formatting (e.g. a whole cover-page field on its own line),
but is worth knowing if a placeholder sits mid-sentence with mixed
formatting around it.
"""
from __future__ import annotations
from docx.document import Document as DocumentObject
from docx.text.paragraph import Paragraph


def _replace_in_paragraph(paragraph: Paragraph, replacements: dict[str, str]) -> bool:
    """Replace every occurrence of every key in `replacements` within this
    paragraph's visible text. Returns True if anything changed."""
    full_text = "".join(run.text for run in paragraph.runs)
    if not full_text:
        return False

    new_text = full_text
    changed = False
    for token, value in replacements.items():
        if token in new_text:
            new_text = new_text.replace(token, value)
            changed = True

    if not changed:
        return False

    # Put the fully-replaced text in the first run, blank out the rest.
    # This is the standard workaround for cross-run placeholder splitting.
    paragraph.runs[0].text = new_text
    for run in paragraph.runs[1:]:
        run.text = ""
    return True


def replace_placeholders(doc: DocumentObject, replacements: dict[str, str]) -> int:
    """
    Walk the entire document - body paragraphs, tables, and every
    section's header/footer (including first-page variants) - replacing
    {{TOKEN}}-style placeholders. Returns the number of paragraphs
    changed, so callers can warn if a placeholder was expected but never
    found (a common sign the template's token spelling doesn't match).
    """
    count = 0

    def _walk_paragraphs(paragraphs):
        nonlocal count
        for p in paragraphs:
            if _replace_in_paragraph(p, replacements):
                count += 1

    def _walk_tables(tables):
        for table in tables:
            for row in table.rows:
                for cell in row.cells:
                    _walk_paragraphs(cell.paragraphs)
                    _walk_tables(cell.tables)  # nested tables

    _walk_paragraphs(doc.paragraphs)
    _walk_tables(doc.tables)

    for section in doc.sections:
        header_containers = [section.header, section.first_page_header, section.even_page_header]
        footer_containers = [section.footer, section.first_page_footer, section.even_page_footer]
        for container in header_containers + footer_containers:
            if container is None:
                continue
            _walk_paragraphs(container.paragraphs)
            _walk_tables(container.tables)

    return count


def find_unmatched_placeholders(doc: DocumentObject) -> set[str]:
    """
    Scan the document for anything that still looks like a {{TOKEN}}
    placeholder after substitution - useful for warning the user their
    template has a token the pipeline doesn't know how to fill.
    """
    import re
    pattern = re.compile(r"\{\{[A-Z0-9_]+\}\}")
    found: set[str] = set()

    def _scan(paragraphs):
        for p in paragraphs:
            text = "".join(run.text for run in p.runs)
            found.update(pattern.findall(text))

    def _scan_tables(tables):
        for table in tables:
            for row in table.rows:
                for cell in row.cells:
                    _scan(cell.paragraphs)
                    _scan_tables(cell.tables)

    _scan(doc.paragraphs)
    _scan_tables(doc.tables)
    for section in doc.sections:
        for container in [section.header, section.first_page_header, section.footer, section.first_page_footer]:
            if container is None:
                continue
            _scan(container.paragraphs)
            _scan_tables(container.tables)

    return found
