"""TestX tool package. SPEC is what BPM Studio reads to show the tool."""
from studio.registry import ToolSpec
from tools.testx.icon import testx_icon
from tools.testx.page import render

SPEC = ToolSpec(
    id="testx",
    name="TestX",
    description="Test Case Generator",
    blurb="Turn the product workbook into UAT test cases (attributes, rules, pricing, order flows, SRs) with expected results in CRM, OSS and BSS.",
    icon_svg=testx_icon,
    nav_icon=":material/fact_check:",
    render=render,
    tags=(("Excel Input", ""), ("Scenario Test Cases", ""), ("CRM · OSS · BSS Results", ""), ("You Choose What Is Included", ""), ("Excel Test Plan", "g")),
    help_md=("**TestX in four steps**\n\n1. Upload the product configuration workbook (.xlsx)\n"
             "2. Tick the test cases you want (the usual ones are pre-ticked) or add your own\n"
             "3. Enter the product name\n"
             "4. Download the Excel test cases and run them in UAT\n\n"
             "Advanced: choose 'Combination matrix' on step 2 for pairwise / full attribute combinations.\n"
             "TestX only uses facts that are in your workbook."),
)
