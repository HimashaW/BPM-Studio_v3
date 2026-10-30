"""The tool list. To add a tool to BPM Studio: create tools/<name>/ with a SPEC (see studio/registry.py)
and add it below - that is the only line outside the tool's own folder that has to change."""
from tools import guidex, testx

TOOLS = [guidex.SPEC, testx.SPEC]
