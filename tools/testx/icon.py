"""TestX icon: a checklist clipboard with a tiny bug - same rounded-square family as GuideX."""
from __future__ import annotations


def testx_icon(size=44) -> str:
    return f'''<svg width="{size}" height="{size}" viewBox="0 0 48 48" xmlns="http://www.w3.org/2000/svg" aria-label="TestX icon">
<defs><linearGradient id="txG" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#0879D9"/><stop offset="1" stop-color="#19C995"/></linearGradient></defs>
<rect width="48" height="48" rx="12" fill="url(#txG)"/>
<rect x="12" y="9" width="24" height="31" rx="3.5" fill="#fff"/>
<rect x="18" y="6.5" width="12" height="6.5" rx="2.4" fill="#0B2D55"/>
<path d="M16.5 20.5l2.4 2.4 4.2-4.6" stroke="#19C995" stroke-width="2.4" fill="none" stroke-linecap="round" stroke-linejoin="round"/>
<path d="M26 21h6" stroke="#B9C9DD" stroke-width="2.4" stroke-linecap="round"/>
<path d="M16.5 29.5l2.4 2.4 4.2-4.6" stroke="#19C995" stroke-width="2.4" fill="none" stroke-linecap="round" stroke-linejoin="round"/>
<path d="M26 30h6" stroke="#B9C9DD" stroke-width="2.4" stroke-linecap="round"/>
<g transform="translate(37.5 38)"><circle r="6.4" fill="#FF7A7A" stroke="#fff" stroke-width="1.6"/><path d="M0 -6 V-0.5" stroke="#B83A3A" stroke-width="1"/>
<circle cx="-2.4" cy="-2.2" r="1.1" fill="#7A2A2A"/><circle cx="2.4" cy="-2.2" r="1.1" fill="#7A2A2A"/>
<circle cx="-2" cy="1.2" r="1.7" fill="#fff"/><circle cx="2" cy="1.2" r="1.7" fill="#fff"/><circle cx="-1.8" cy="1.5" r=".8" fill="#0B2D55"/><circle cx="2.2" cy="1.5" r=".8" fill="#0B2D55"/></g>
</svg>'''
