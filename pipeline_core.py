"""
Shared pipeline core: extract -> validate -> render, as a set of plain
functions with no CLI/argparse or Streamlit dependency. Both run.py (CLI)
and app.py (web UI) call into this, so there is exactly one implementation
of the pipeline logic to maintain.
"""
from __future__ import annotations
from pathlib import Path

DEFAULT_TEMPLATE_PATH = Path(__file__).resolve().parent / "assets" / "process_guide_template.docx"

from extractors.stacked_tables import extract_order_types
from extractors.product_description import extract_product_description
from extractors.special_requirements import extract_business_rules, extract_business_rules_table, extract_business_rule_blocks
from extractors.order_scenarios import extract_order_scenarios
from extractors.product_hierarchy import extract_hierarchy
from extractors.hierarchy_images import extract_hierarchy_images
from extractors.task_descriptions import extract_task_descriptions, attach_task_descriptions
from models.schemas import ExtractionResult, OrderScenarioRow
from validation.checks import run_all_checks
from render.build_doc import build

try:
    from ai.introduction import generate_introduction
except ImportError:
    generate_introduction = None

try:
    import yaml
except ImportError:
    yaml = None


def best_effort(label: str, fn, warnings: list[str], default):
    """Run an extractor; on failure, log a warning and fall back to
    `default` instead of aborting the whole pipeline. Used for sections
    whose source sheets are optional / not yet confirmed."""
    try:
        return fn()
    except ValueError as e:
        warnings.append(f"[{label}] {e}")
        return default


def load_order_scenarios(path: Path) -> list[OrderScenarioRow]:
    if not path.exists() or yaml is None:
        return []
    with open(path) as f:
        rows = yaml.safe_load(f) or []
    return [OrderScenarioRow(**r) for r in rows]


def extract_all(
    input_path: str,
    product: str,
    sheet: str = "6.3 OSS SO Task List ",
    order_scenarios_path: str | Path = "input/order_scenarios.yaml",
    progress_cb=None,
    ai_enhance: bool = False,
    ai_model: str = "gemini-3.5-flash",
    gemini_api_key: str | None = None,
) -> ExtractionResult:
    """Run every extractor against `input_path` and return a validated
    ExtractionResult. `progress_cb`, if given, is called with a short
    status string before each extraction stage (used by the Streamlit UI
    to show progress; the CLI passes `print`)."""
    def _report(msg):
        if progress_cb:
            progress_cb(msg)

    warnings: list[str] = []

    _report(f"Extracting order types for '{product}' from '{sheet}' ...")
    order_types = extract_order_types(
        str(input_path), sheet_name=sheet, product=product, warnings=warnings,
    )

    _report("Extracting Introduction (1. Product Description) ...")
    introduction_text = best_effort(
        "1. Product Description",
        lambda: extract_product_description(str(input_path)),
        warnings, default=[],
    )

    _report("Extracting Business Rules (3. Special Requirements) ...")
    business_rules = best_effort(
        "3. Special Requirements",
        lambda: extract_business_rules(str(input_path)),
        warnings, default=[],
    )
    business_rules_table = None
    business_rule_blocks = best_effort(
        "3. Special Requirements (requirement blocks)",
        lambda: extract_business_rule_blocks(str(input_path)),
        warnings, default=[],
    )

    _report("Extracting Product Hierarchy (2. Product Hierarchy) ...")
    hierarchy = best_effort(
        "2. Product Hierarchy",
        lambda: extract_hierarchy(str(input_path)),
        warnings, default=[],
    )

    _report("Extracting images from 1. Product Description ...")
    hierarchy_images = best_effort(
        "1. Product Description (images)",
        lambda: extract_hierarchy_images(str(input_path), sheet_name="1. Product Description"),
        warnings, default=[],
    )

    _report("Extracting Task Descriptions (6.4) ...")
    descriptions = best_effort(
        "6.4 Task Descriptions",
        lambda: extract_task_descriptions(str(input_path)),
        warnings, default={},
    )
    matched_descriptions = attach_task_descriptions(order_types, descriptions)
    if descriptions:
        _report(f"Matched {matched_descriptions} task descriptions from 6.4")

    _report("Extracting Order Scenarios (5.1 MDM_Order Scenarios) ...")
    excel_order_scenarios = best_effort(
        "5.1 MDM_Order Scenarios",
        lambda: extract_order_scenarios(str(input_path)),
        warnings, default=[],
    )

    # Excel is now the primary source. YAML remains as an optional fallback
    # for older/manual workflows when the 5.1 sheet is unavailable.
    if excel_order_scenarios:
        order_scenarios = [OrderScenarioRow(**row) for row in excel_order_scenarios]
        _report(f"Loaded {len(order_scenarios)} order-scenario rows from 5.1")
    else:
        order_scenarios = load_order_scenarios(Path(order_scenarios_path))
        if order_scenarios:
            _report(f"Loaded {len(order_scenarios)} order-scenario rows from YAML fallback")

    result = ExtractionResult(
        workbook_name=Path(input_path).name,
        product_name=product,
        introduction_text=introduction_text,
        business_rules=business_rules,
        business_rules_table=business_rules_table,
        business_rule_blocks=business_rule_blocks,
        hierarchy=hierarchy,
        hierarchy_images=hierarchy_images,
        order_scenarios=order_scenarios,
        order_types=order_types,
        warnings=warnings,
    )

    if ai_enhance:
        _report(f"Generating AI-enhanced Introduction with Gemini ({ai_model}) ...")
        try:
            if generate_introduction is None:
                raise RuntimeError("AI module is unavailable. Install the pipeline requirements including google-genai.")
            result.ai_introduction = generate_introduction(
                result, api_key=gemini_api_key, model=ai_model
            )
            _report("AI-enhanced Introduction generated and validated.")
        except Exception as e:
            result.warnings.append(f"[Gemini Introduction] {e}. Falling back to Excel Introduction text.")
            _report("Gemini enhancement failed; Excel Introduction will be used as fallback.")

    _report("Running validation checks ...")
    result = run_all_checks(result)
    return result


def render(result: ExtractionResult, output_path: str, **doc_kwargs) -> str:
    """Thin wrapper around render.build_doc.build(), resolving a bad/missing
    logo path (or template path) to a safe fallback with a warning instead
    of failing."""
    logo_path = doc_kwargs.get("logo_path")
    if logo_path and not Path(logo_path).exists():
        result.warnings.append(
            f"--logo path '{logo_path}' does not exist - fell back to the placeholder logo."
        )
        doc_kwargs["logo_path"] = None

    template_path = doc_kwargs.get("template_path")
    if not template_path:
        # The supplied Word format is now the default document master.
        # A caller can still pass --template to override it explicitly.
        if DEFAULT_TEMPLATE_PATH.exists():
            template_path = str(DEFAULT_TEMPLATE_PATH)
            doc_kwargs["template_path"] = template_path
        else:
            result.warnings.append(
                f"Default Word template not found at '{DEFAULT_TEMPLATE_PATH}'. "
                f"Falling back to the legacy code-generated layout."
            )
            doc_kwargs["template_path"] = None
    elif not Path(template_path).exists():
        result.warnings.append(
            f"--template path '{template_path}' does not exist - fell back to "
            f"the default bundled template."
        )
        doc_kwargs["template_path"] = (
            str(DEFAULT_TEMPLATE_PATH) if DEFAULT_TEMPLATE_PATH.exists() else None
        )

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    return build(result, output_path, **doc_kwargs)


def export_flowcharts_pptx(result: ExtractionResult, output_path: str, product: str | None = None) -> dict:
    """Editable PowerPoint with every order type's swimlane flowchart.

    One slide per diagram panel, built from the same layout as the images in
    the Word guide, using native shapes and connectors the user can edit.
    Returns {"slides": n, "path": ..., "warnings": [...]}.
    """
    from render.task_flow_pptx import build_flowcharts_pptx
    return build_flowcharts_pptx(result.order_types, output_path, product or result.product_name)


def normalize_attribute_names(result: ExtractionResult) -> ExtractionResult:
    """Merge hierarchy attributes that repeat the same name into one entry
    (values are combined), so each attribute name appears once."""
    for node in result.hierarchy:
        unique, lookup = [], {}
        for attr in node.attributes:
            key = (attr.entity_name or "").strip().casefold()
            if key in lookup:
                ex = lookup[key]
                for v in (attr.values or []):
                    if v not in ex.values:
                        ex.values.append(v)
                if not ex.mandatory and attr.mandatory:
                    ex.mandatory = attr.mandatory
                if not ex.comments and attr.comments:
                    ex.comments = attr.comments
            else:
                lookup[key] = attr
                unique.append(attr)
        node.attributes = unique
    return result
