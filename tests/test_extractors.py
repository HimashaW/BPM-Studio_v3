"""
Run with: python -m pytest tests/ -v
(or: python tests/test_extractors.py  to run without pytest installed)
"""
import sys
import os
import pytest
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from extractors.stacked_tables import extract_order_types
from extractors.sheet_utils import find_sheet

WORKBOOK = os.path.join(os.path.dirname(__file__), "..", "input", "Book1.xlsx")
requires_real_workbook = pytest.mark.skipif(not os.path.exists(WORKBOOK), reason="bundled real-world workbook not included in distributable pipeline")
SHEET = "6.3 OSS SO Task List "

EXPECTED_TASK_COUNTS = {
    "CREATE": 22,
    "MODIFY-PACKAGE": 10,
    "MODIFY-FEATURE": 6,
    "MODIFY-LOCATION": 23,
    "DELETE": 5,
    "SUSPEND": 4,
    "RESUME": 4,
    "CREATE-OR": 23,
    "DELETE-OR": 6,
}


@requires_real_workbook
def test_extracts_all_nine_order_types():
    sections = extract_order_types(WORKBOOK, SHEET, product="CYS_EdgeDefend")
    found = {s.order_type for s in sections}
    expected = set(EXPECTED_TASK_COUNTS.keys())
    assert found == expected, f"Mismatch. Expected {expected}, got {found}"


@requires_real_workbook
def test_task_counts_match_known_good_values():
    sections = extract_order_types(WORKBOOK, SHEET, product="CYS_EdgeDefend")
    by_type = {s.order_type: s for s in sections}
    for order_type, expected_count in EXPECTED_TASK_COUNTS.items():
        actual = by_type[order_type].task_count
        assert actual == expected_count, (
            f"{order_type}: expected {expected_count} tasks, got {actual}"
        )


@requires_real_workbook
def test_every_task_has_a_name():
    sections = extract_order_types(WORKBOOK, SHEET, product="CYS_EdgeDefend")
    for sec in sections:
        for t in sec.tasks:
            assert t.task_name and t.task_name.strip(), (
                f"Blank task_name in {sec.order_type} at row {t.source_row}"
            )


@requires_real_workbook
def test_raises_on_missing_sheet():
    try:
        extract_order_types(WORKBOOK, "NOT A REAL SHEET", product="CYS_EdgeDefend")
        assert False, "Expected ValueError for missing sheet"
    except ValueError as e:
        assert "not found" in str(e)


@requires_real_workbook
def test_product_name_is_stamped_on_every_section():
    sections = extract_order_types(WORKBOOK, SHEET, product="AcmeWidget")
    assert all(s.product == "AcmeWidget" for s in sections), (
        "product name should propagate to every extracted section unchanged"
    )


@requires_real_workbook
def test_sheet_resolver_tolerates_typo_in_expected_name():
    """Regression test for the real bug: workbook had '6.3 OSS SO Task List '
    with a trailing space (exact match still required to work), and this
    guards against future typo/whitespace mismatches the same way."""
    import openpyxl
    wb = openpyxl.load_workbook(WORKBOOK)
    # exact match still works
    assert find_sheet(wb, SHEET) == SHEET
    # trimmed/whitespace-normalized match works
    assert find_sheet(wb, SHEET.strip()) == SHEET


@requires_real_workbook
def test_sheet_resolver_raises_with_available_sheets_listed():
    import openpyxl
    wb = openpyxl.load_workbook(WORKBOOK)
    try:
        find_sheet(wb, "9. Definitely Not A Real Sheet")
        assert False, "Expected ValueError"
    except ValueError as e:
        assert "Available sheets" in str(e)


def test_sheet_resolver_rejects_unrelated_sheet_sharing_number_prefix():
    """Regression test: a real workbook has '9. Test Scenarios'. Asking
    for an unrelated '9. <something else>' must NOT resolve to it just
    because they share the leading number - text similarity must also
    clear the threshold, or it should raise."""
    import openpyxl
    wb = openpyxl.Workbook()
    wb.active.title = "9. Test Scenarios"
    from extractors.sheet_utils import find_sheet as _find
    try:
        result = _find(wb, "9. Completely Unrelated Content")
        assert False, f"Should have raised, but incorrectly resolved to '{result}'"
    except ValueError:
        pass  # correct - no confident match should be found


def test_attribute_values_are_grouped_under_their_attribute():
    """Regression test for the 'Attribute List of Values' feature:
    'Value' rows (Entity Type = 'Value') following an 'Attribute' row
    must attach to THAT attribute, and an attribute with no Value rows
    must end up with an empty values list (not crash, not merge into
    the wrong attribute)."""
    from extractors.product_hierarchy import extract_hierarchy
    fixture = os.path.join(os.path.dirname(__file__), "fixtures", "hierarchy_sample.xlsx")
    components = extract_hierarchy(fixture)
    assert len(components) == 1
    attrs_by_name = {a.entity_name: a for a in components[0].attributes}

    assert attrs_by_name["Transfer Category"].values == [
        "Transfer among same Cus.", "Transfer to new Customer",
    ]
    assert attrs_by_name["Effective Date"].values == []


def test_business_rules_survive_blank_spacer_rows():
    """Regression test: extraction used to stop after 3 consecutive blank
    rows in the 'Requirement Description' column, assuming that meant
    end-of-table. Real sheets can have blank spacer/formatting rows
    between entries, which was cutting extraction off after just the
    first requirement. Must now read every non-blank cell down to the
    bottom of the sheet, regardless of blank runs in between."""
    from extractors.special_requirements import extract_business_rules
    fixture = os.path.join(os.path.dirname(__file__), "fixtures", "requirements_sample.xlsx")
    rules = extract_business_rules(fixture)
    assert rules == [
        "All orders must be logged in the CRM system.",
        "Customer identity must be verified before CREATE.",
        "All DELETE orders require manager approval.",
    ], f"Expected 3 rules surviving the blank spacer rows, got: {rules}"


def test_numeric_value_in_text_column_is_coerced_not_rejected():
    """Regression test: a real workbook had a plain number (e.g. 14) in
    the 'Cancellation WO' column, which the Task schema expects to be
    text (e.g. 'YES'/'NO'). This used to raise a pydantic ValidationError
    and abort the whole extraction. Must now coerce to string and warn,
    the same way a bad 'duration' value is handled."""
    fixture = os.path.join(os.path.dirname(__file__), "fixtures", "numeric_in_text_column.xlsx")
    warnings: list[str] = []
    sections = extract_order_types(
        fixture, "6.2 OSS SO Task List ", product="D_SASE", warnings=warnings,
    )
    tasks = sections[0].tasks
    assert tasks[0].cancellation_wo == "14"
    assert tasks[1].cancellation_wo == "NO"
    assert any("cancellation_wo" in w for w in warnings)


@requires_real_workbook
def test_supplied_template_is_used_as_document_master():
    """The supplied Word template owns the common document format. The
    renderer fills its metadata/control fields and inserts generated content
    after the template's Introduction anchor instead of rebuilding the
    front matter."""
    import docx
    import pipeline_core as core
    from models.schemas import ExtractionResult
    from extractors.stacked_tables import extract_order_types

    workbook = os.path.join(os.path.dirname(__file__), "..", "input", "Book1.xlsx")
    template = os.path.join(os.path.dirname(__file__), "..", "assets", "process_guide_template.docx")
    out_path = os.path.join(os.path.dirname(__file__), "..", "output", "_test_template_master.docx")

    order_types = extract_order_types(workbook, "6.3 OSS SO Task List ", product="TestProduct")
    result = ExtractionResult(workbook_name="x", product_name="TestProduct", order_types=order_types)

    core.render(
        result,
        out_path,
        doc_no="DOC-1",
        issue_no="002",
        issue_date="10.08.2026",
        revision_no="1",
        revision_date="10.08.2026",
        issued_by="BPM Section",
        contact_name="Test User",
        template_path=template,
    )

    doc = docx.Document(out_path)

    # The template cover/control structure remains present.
    assert any(p.text.strip() == "Document Control Sheet" for p in doc.paragraphs)
    assert any(p.text.strip() == "Introduction" for p in doc.paragraphs)

    # Generated content was inserted after the template's Introduction anchor.
    intro_index = next(i for i, p in enumerate(doc.paragraphs) if p.text.strip() == "Introduction")
    following = [p.text for p in doc.paragraphs[intro_index + 1:intro_index + 8]]
    assert any("Product-related Business Rules" in x for x in following)

    # Revision history is populated from the supplied metadata.
    revision_table = next(
        table for table in doc.tables
        if table.cell(0, 0).text.strip() == "Issue Number"
    )
    assert revision_table.cell(1, 0).text == "002"
    assert revision_table.cell(1, 1).text == "10.08.2026"
    assert revision_table.cell(1, 2).text == "BPM Section"

    # The template's regular header remains the source of the visual layout.
    header_xml = doc.sections[0].header._element.xml
    assert "TestProduct" in header_xml
    assert "DOC-1" in header_xml

    os.remove(out_path)


@requires_real_workbook
def test_template_without_tokens_still_accepts_metadata_by_template_labels():
    """The real supplied template has no {{TOKENS}}. Its existing Word
    labels/shapes are therefore populated directly without requiring the
    template author to add placeholders."""
    import zipfile
    import pipeline_core as core
    from models.schemas import ExtractionResult
    from extractors.stacked_tables import extract_order_types

    workbook = os.path.join(os.path.dirname(__file__), "..", "input", "Book1.xlsx")
    template = os.path.join(os.path.dirname(__file__), "..", "assets", "process_guide_template.docx")
    out_path = os.path.join(os.path.dirname(__file__), "..", "output", "_test_static_template.docx")

    order_types = extract_order_types(workbook, "6.3 OSS SO Task List ", product="TestProduct")
    result = ExtractionResult(workbook_name="x", product_name="TestProduct", order_types=order_types)

    core.render(result, out_path, template_path=template, doc_no="DOC-2")

    with zipfile.ZipFile(out_path) as z:
        document_xml = z.read("word/document.xml").decode("utf-8")
        header_xml = z.read("word/header2.xml").decode("utf-8")

    assert "BPMS XXX XXX" not in document_xml
    assert "DOC-2" in document_xml
    assert "TestProduct" in header_xml

    os.remove(out_path)


if __name__ == "__main__":
    # Minimal runner if pytest isn't available
    tests = [v for k, v in list(globals().items()) if k.startswith("test_")]
    passed, failed = 0, 0
    for t in tests:
        try:
            t()
            print(f"PASS: {t.__name__}")
            passed += 1
        except AssertionError as e:
            print(f"FAIL: {t.__name__} -> {e}")
            failed += 1
    print(f"\n{passed} passed, {failed} failed")
    sys.exit(1 if failed else 0)


def test_special_requirements_render_tables_under_specific_requirement_only():
    from extractors.special_requirements import extract_business_rule_blocks
    fixture = os.path.join(os.path.dirname(__file__), "fixtures", "special_requirements_with_embedded_tables.xlsx")
    blocks = extract_business_rule_blocks(fixture)
    assert len(blocks) == 2
    assert blocks[0].text.startswith("The LOVs of the attribute")
    assert len(blocks[0].tables) == 1
    assert blocks[0].tables[0].headers == ["Package Type LOVs", "CPE Model", "Max Users", "Concurrent Users"]
    assert blocks[0].tables[0].rows[0] == ["Standard", "FortiF", "50", "25"]
    assert blocks[1].text.startswith("The attribute value of the following feature attributes")
    assert len(blocks[1].tables) == 1
    assert blocks[1].tables[0].headers == ["Feature Attribute", "Standard", "Pro", "Premium"]


def test_duplicate_attribute_rows_are_consolidated():
    from extractors.product_hierarchy import extract_hierarchy
    fixture = os.path.join(os.path.dirname(__file__), "fixtures", "hierarchy_duplicate_attributes.xlsx")
    components = extract_hierarchy(fixture)
    attrs = components[0].attributes
    assert [a.entity_name for a in attrs] == ["Priority", "SLA_Maintenance Window"]
    assert attrs[0].values == ["1", "2", "3", "4", "5"]
    assert attrs[1].values == ["24 x 7", "STANDARD", "NONE"]
