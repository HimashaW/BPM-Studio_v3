"""What a tool has to provide to appear in BPM Studio.

To add a tool you create a package under tools/ that exposes a ToolSpec called SPEC and
add one line to tools/__init__.py. The sidebar, the Home cards, the page title, the help
popover and the router are all built from this list - app.py never changes.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Callable


@dataclass(frozen=True)
class ToolSpec:
    id: str                              # short, unique, used in session state: "guidex"
    name: str                            # shown everywhere: "GuideX"
    description: str                     # one line: "Service Delivery Guide Generator"
    blurb: str                           # sentence on the Home card
    icon_svg: Callable[[int], str]       # icon_svg(size) -> inline SVG string
    nav_icon: str                        # sidebar icon, e.g. ":material/description:"
    render: Callable[[], None]           # draws the tool's page
    help_md: str                         # text of the Help popover on the tool's page
    tags: tuple = field(default_factory=tuple)   # (("Excel Input", ""), ("Editable PowerPoint", "g"))
    version: str = "v1.0"
