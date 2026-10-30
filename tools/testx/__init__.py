"""TestX tool package. SPEC is what BPM Studio reads to show the tool."""
from studio.registry import ToolSpec
from tools.testx.icon import testx_icon
from tools.testx.page import render

SPEC = ToolSpec(
    id="testx",
    name="TestX",
    description="Test Case Generator",
    blurb="Turn attributes, values and business rules into a UAT-ready Excel test plan - without testing every combination.",
    icon_svg=testx_icon,
    nav_icon=":material/fact_check:",
    render=render,
    tags=(("Excel Input", ""), ("Pairwise Coverage", ""), ("Business Rules", ""), ("Negative Tests", ""), ("Excel Test Plan", "g")),
    help_md=("**TestX in four steps**\n\n1. Upload the product configuration workbook (.xlsx)\n"
             "2. Review the attributes and the rules TestX found - switch off what you do not want\n"
             "3. Choose the technique (pairwise is recommended) and the layout\n"
             "4. Download the Excel test plan and run it in UAT\n\n"
             "TestX only uses attributes and values that are in your workbook."),
)
