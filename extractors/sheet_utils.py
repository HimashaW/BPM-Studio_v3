"""
Shared helper: resolve a sheet name against the actual workbook, tolerating
minor differences (extra whitespace, case, small typos) so a misspelled
sheet name in the source Excel doesn't silently produce an empty stub.
"""
from __future__ import annotations
import re
import difflib

# Minimum text-similarity ratio (0-1) required before accepting a fuzzy
# match on the leading "N." prefix. This must be high enough that a typo
# like "Descripton" vs "Description" passes, but a genuinely different
# sheet that happens to share a number prefix (e.g. "9. Test Scenarios"
# vs a mistakenly-requested "9. Nonexistent Sheet") is rejected.
SIMILARITY_THRESHOLD = 0.7


def find_sheet(wb, expected_name: str) -> str:
    """Return the real sheet name in `wb` that best matches `expected_name`.

    Matching order:
      1. Exact match
      2. Match ignoring case/extra whitespace
      3. Among sheets sharing the same leading "N." number prefix, the
         one with the highest text-similarity ratio to `expected_name`
         - accepted ONLY if that ratio clears SIMILARITY_THRESHOLD, so a
         shared number prefix alone is never enough (e.g. "9. Test
         Scenarios" must not match an unrelated "9. Something Else").

    Raises ValueError with the full list of available sheet names if no
    confident match is found, so the failure is always inspectable.
    """
    if expected_name in wb.sheetnames:
        return expected_name

    norm = lambda s: " ".join(s.strip().lower().split())
    target = norm(expected_name)
    for name in wb.sheetnames:
        if norm(name) == target:
            return name

    m = re.match(r"^\s*(\d+)\.", expected_name)
    if m:
        prefix = m.group(1) + "."
        candidates = [n for n in wb.sheetnames if n.strip().startswith(prefix)]
        if candidates:
            best = max(
                candidates,
                key=lambda n: difflib.SequenceMatcher(None, norm(n), target).ratio(),
            )
            best_ratio = difflib.SequenceMatcher(None, norm(best), target).ratio()
            if best_ratio >= SIMILARITY_THRESHOLD:
                return best

    raise ValueError(
        f"Could not find a sheet matching '{expected_name}'. "
        f"Available sheets: {wb.sheetnames}"
    )
