#!/usr/bin/env python3
"""
Pipeline entrypoint (CLI).

Usage:
    python run.py --input input/Book1.xlsx --output output/ServiceDeliveryGuide.docx --product "Your Product"

Stages: extract -> validate -> render -> report.
Fails loudly with specific messages rather than silently producing a
half-empty document. This is a thin wrapper around pipeline_core.py -
the web app (app.py) uses the exact same core logic.
"""
from __future__ import annotations
import argparse
import sys
import os
from pathlib import Path

import pipeline_core as core


def main():
    parser = argparse.ArgumentParser(description="Generate Service Delivery Guide from BPM config Excel")
    parser.add_argument("--input", required=True, help="Path to source .xlsx")
    parser.add_argument("--output", required=True, help="Path to write .docx")
    parser.add_argument("--product", required=True,
                         help="Product/service name, e.g. 'CYS_EdgeDefend' - used to title the "
                              "document and fill product-specific headings throughout")
    parser.add_argument("--sheet", default="6.3 OSS SO Task List ",
                         help="Sheet name for order-type task lists (note trailing space in source file)")
    parser.add_argument("--version", default=None, help="Document version label, e.g. '1.0'")

    # Document Control Sheet / running header metadata - all optional,
    # default to "[TBD]" placeholders so the pipeline still runs without them.
    parser.add_argument("--doc-no", default="[TBD]", help="Document number shown in the running header")
    parser.add_argument("--issue-no", default="[TBD]", help="Issue number")
    parser.add_argument("--issue-date", default=None, help="Date of issue (defaults to today)")
    parser.add_argument("--revision-no", default="[TBD]", help="Revision number")
    parser.add_argument("--revision-date", default="[TBD]", help="Date of revision")
    parser.add_argument("--issued-by", default="BPM Section", help="Front page: 'Issued By' field")
    parser.add_argument("--logo", default=None, help="Path to a company logo image for the running header")
    parser.add_argument("--template", default=None,
                         help="Optional override for the bundled Word document master. "
                              "If omitted, assets/process_guide_template.docx is used.")
    parser.add_argument("--ai-enhance", action="store_true",
                         help="Use Gemini to generate a source-grounded professional Introduction. Requires GEMINI_API_KEY.")
    parser.add_argument("--ai-model", default="gemini-3.5-flash",
                         help="Gemini model used when --ai-enhance is enabled.")
    parser.add_argument("--no-flowchart", action="store_true",
                         help="Do not generate editable swimlane task-flow diagrams. "
                              "By default a diagram is generated for every order-type task flow.")
    parser.add_argument("--contact-name", default="", help="Document Control Sheet: contact name")
    parser.add_argument("--contact-designation", default="DGM/BPM", help="Document Control Sheet: contact designation")
    parser.add_argument("--contact-telephone", default="", help="Document Control Sheet: contact telephone")
    parser.add_argument("--contact-mobile", default="", help="Document Control Sheet: contact mobile")
    parser.add_argument("--contact-extension", default="", help="Document Control Sheet: contact extension")
    parser.add_argument("--contact-email", default="", help="Document Control Sheet: contact email")
    args = parser.parse_args()

    input_path = Path(args.input)
    if not input_path.exists():
        print(f"ERROR: input file not found: {input_path}", file=sys.stderr)
        sys.exit(1)

    print("[1/3] Extracting ...")
    try:
        result = core.extract_all(
            str(input_path), product=args.product, sheet=args.sheet,
            progress_cb=lambda msg: print(f"      {msg}"),
            ai_enhance=args.ai_enhance,
            ai_model=args.ai_model,
            gemini_api_key=os.getenv("GEMINI_API_KEY"),
        )
    except ValueError as e:
        print(f"ERROR during extraction: {e}", file=sys.stderr)
        sys.exit(1)

    print(f"      -> {len(result.order_types)} order type sections, "
          f"{sum(s.task_count for s in result.order_types)} total tasks, "
          f"{len(result.hierarchy)} hierarchy nodes")

    print(f"[2/3] Rendering document to {args.output} ...")
    core.render(
        result, args.output, doc_version=args.version,
        doc_no=args.doc_no, issue_no=args.issue_no, issue_date=args.issue_date,
        revision_no=args.revision_no, revision_date=args.revision_date,
        issued_by=args.issued_by, logo_path=args.logo, template_path=args.template,
        contact_name=args.contact_name, contact_designation=args.contact_designation,
        contact_telephone=args.contact_telephone, contact_mobile=args.contact_mobile,
        contact_extension=args.contact_extension, contact_email=args.contact_email,
        include_task_flowchart=not args.no_flowchart,
    )

    # Printed AFTER rendering, not before - some warnings (e.g. an unmatched
    # template placeholder) can only be discovered during rendering itself,
    # so summarizing warnings any earlier would silently miss them.
    if result.warnings:
        print(f"[3/3] {len(result.warnings)} warning(s):")
        for w in result.warnings:
            print(f"        - {w}")
    else:
        print("[3/3] No warnings.")

    print("\nDone.")
    print(f"Output: {args.output}")
    if result.warnings:
        print(f"NOTE: {len(result.warnings)} warning(s) were raised above - "
              f"review before treating this as final.")


if __name__ == "__main__":
    main()
