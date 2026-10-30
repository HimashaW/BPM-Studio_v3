"""GuideX tool package. SPEC is what BPM Studio reads to show the tool."""
from studio.registry import ToolSpec
from tools.guidex.icon import guidex_icon
from tools.guidex.page import render

SPEC = ToolSpec(
    id="guidex",
    name="GuideX",
    description="Service Delivery Guide Generator",
    blurb="Generate structured Service Delivery Guides from product configuration and process data.",
    icon_svg=guidex_icon,
    nav_icon=":material/description:",
    render=render,
    tags=(("Excel Input", ""), ("AI-Assisted Content (optional)", ""), ("Automated Flowcharts", ""),
          ("Word Export", ""), ("Editable PowerPoint", "g")),
    help_md=("**GuideX in four steps**\n\n1. Upload the product configuration workbook (.xlsx)\n"
             "2. Review what was read from it\n3. Enter the document details and generate\n"
             "4. Review the warnings and download your files\n\n"
             "The workbook is the source of truth: nothing is added that it does not contain."),
)
