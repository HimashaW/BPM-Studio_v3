"""GuideX icon: a document with a forward arrow."""
from __future__ import annotations


def guidex_icon(size=44) -> str:
    """Document with a forward arrow."""
    return f'''<svg width="{size}" height="{size}" viewBox="0 0 48 48" xmlns="http://www.w3.org/2000/svg" aria-label="GuideX icon">
<defs><linearGradient id="gxG" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#00A6B2"/><stop offset="1" stop-color="#0879D9"/></linearGradient></defs>
<rect width="48" height="48" rx="12" fill="url(#gxG)"/>
<path d="M15 9h13l8 8v21a2 2 0 0 1-2 2H15a2 2 0 0 1-2-2V11a2 2 0 0 1 2-2z" fill="#fff"/>
<path d="M28 9v8h8z" fill="#BFE9EE"/>
<path d="M18 29h11m-4-4 4 4-4 4" stroke="#0879D9" stroke-width="2.4" fill="none" stroke-linecap="round" stroke-linejoin="round"/></svg>'''
