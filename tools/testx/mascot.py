"""Tess - the TestX mascot, an original chibi tester drawn as inline SVG (no images, no external files).

Tess has mint hair, a headset, a clipboard and a bug-catching magnifier. Her sidekick Bugsy is the
little ladybug-like bug she keeps finding. Both are original characters made for BPM Studio.

    tess("hello" | "think" | "work" | "cheer" | "oops", size=180)  ->  SVG string
"""
from __future__ import annotations

NAVY, BLUE, DBLUE, TEAL = "#0B2D55", "#0879D9", "#0668BD", "#00A6B2"
MINT, DMINT, LMINT = "#19C995", "#0E9F78", "#8FF0CF"
SKIN, SHADE, BLUSH = "#FFE2CF", "#F4C5AA", "#FF8FA3"
BUG, DBUG = "#FF7A7A", "#B83A3A"

MOODS = ("hello", "think", "work", "cheer", "oops")


def _sparkle(x, y, r, color="#FFD86B", op=1):
    return (f'<path d="M{x},{y-r} Q{x+r*0.18},{y-r*0.18} {x+r},{y} Q{x+r*0.18},{y+r*0.18} {x},{y+r} '
            f'Q{x-r*0.18},{y+r*0.18} {x-r},{y} Q{x-r*0.18},{y-r*0.18} {x},{y-r}Z" fill="{color}" opacity="{op}"/>')


def bugsy(x, y, s=1.0, happy=True, flip=False):
    """The ladybug-like bug. (x, y) = centre of its body."""
    fx = -1 if flip else 1
    mouth = ('<path d="M-3.2,3.4 Q0,6.4 3.2,3.4" stroke="#5A1D1D" stroke-width="1.5" fill="none" stroke-linecap="round"/>' if happy else
             '<path d="M-3,5.2 Q0,2.6 3,5.2" stroke="#5A1D1D" stroke-width="1.5" fill="none" stroke-linecap="round"/>')
    return f'''<g transform="translate({x},{y}) scale({s*fx},{s})">
<path d="M-6,-13 Q-9,-21 -14,-22" stroke="{DBUG}" stroke-width="1.8" fill="none" stroke-linecap="round"/><circle cx="-14.5" cy="-22.5" r="2.1" fill="{DBUG}"/>
<path d="M6,-13 Q9,-21 14,-22" stroke="{DBUG}" stroke-width="1.8" fill="none" stroke-linecap="round"/><circle cx="14.5" cy="-22.5" r="2.1" fill="{DBUG}"/>
<path d="M-12,8 l-6,5 M12,8 l6,5 M-13,1 l-7,1 M13,1 l7,1" stroke="{DBUG}" stroke-width="2" stroke-linecap="round"/>
<ellipse cx="0" cy="0" rx="14" ry="13" fill="{BUG}" stroke="{DBUG}" stroke-width="1.6"/>
<path d="M0,-12.5 V-2" stroke="{DBUG}" stroke-width="1.4"/>
<circle cx="-8" cy="-6" r="2.2" fill="#7A2A2A" opacity=".75"/><circle cx="8" cy="-6" r="2.2" fill="#7A2A2A" opacity=".75"/>
<ellipse cx="-5" cy="0.5" rx="3.6" ry="4.2" fill="#fff"/><ellipse cx="5" cy="0.5" rx="3.6" ry="4.2" fill="#fff"/>
<circle cx="-4.6" cy="1.2" r="2" fill="{NAVY}"/><circle cx="5.4" cy="1.2" r="2" fill="{NAVY}"/>
<circle cx="-4" cy="0.4" r=".8" fill="#fff"/><circle cx="6" cy="0.4" r=".8" fill="#fff"/>
{mouth}
<ellipse cx="-9.5" cy="5" rx="2.4" ry="1.5" fill="{BLUSH}" opacity=".6"/><ellipse cx="9.5" cy="5" rx="2.4" ry="1.5" fill="{BLUSH}" opacity=".6"/>
</g>'''


def _eyes(mood):
    L, R = 90, 130            # eye centres x
    cy = 96
    def big(cx, look=(0, 0)):
        lx, ly = look
        return f'''<ellipse cx="{cx}" cy="{cy}" rx="10.5" ry="13.5" fill="{NAVY}"/>
<ellipse cx="{cx+lx}" cy="{cy+3+ly}" rx="7.6" ry="9.8" fill="url(#tsIris)"/>
<ellipse cx="{cx+lx}" cy="{cy+5+ly}" rx="4.6" ry="5.4" fill="{NAVY}" opacity=".55"/>
<circle cx="{cx-3.4+lx}" cy="{cy-4+ly}" r="4.1" fill="#fff"/><circle cx="{cx+3.6+lx}" cy="{cy+6.5+ly}" r="2" fill="#fff"/>
<path d="M{cx-11.5},{cy-6} Q{cx},{cy-17} {cx+11.5},{cy-6}" stroke="{NAVY}" stroke-width="3.4" fill="none" stroke-linecap="round"/>'''
    def happy(cx):
        return f'<path d="M{cx-9},{cy+3} Q{cx},{cy-9} {cx+9},{cy+3}" stroke="{NAVY}" stroke-width="3.6" fill="none" stroke-linecap="round"/>'
    def squint(cx, sign):
        return f'<path d="M{cx-7*sign},{cy-6} L{cx+6*sign},{cy+1} L{cx-7*sign},{cy+8}" stroke="{NAVY}" stroke-width="3.6" fill="none" stroke-linecap="round" stroke-linejoin="round"/>'
    def star(cx):
        pts = []
        import math
        for i in range(10):
            r = 11 if i % 2 == 0 else 4.6
            a = -math.pi / 2 + i * math.pi / 5
            pts.append(f"{cx + r*math.cos(a):.1f},{cy + 2 + r*math.sin(a):.1f}")
        return (f'<polygon points="{" ".join(pts)}" fill="#FFD86B" stroke="{NAVY}" stroke-width="2.2" stroke-linejoin="round"/>')
    if mood == "hello":
        return big(L) + big(R)
    if mood == "think":
        return big(L, (2.4, -3)) + big(R, (2.4, -3))
    if mood == "work":
        return squint(L, 1) + squint(R, -1)
    if mood == "cheer":
        return star(L) + star(R)
    return big(L, (0, 1.5)) + big(R, (0, 1.5))        # oops


def _mouth(mood):
    if mood == "hello":
        return f'<path d="M100,113 Q110,126 120,113 Z" fill="#FF6B81" stroke="{NAVY}" stroke-width="2.4" stroke-linejoin="round"/><path d="M104,119 Q110,123 116,119" stroke="#fff" stroke-width="2" fill="none" opacity=".7"/>'
    if mood == "think":
        return f'<path d="M104,116 Q110,119 116,115" stroke="{NAVY}" stroke-width="2.6" fill="none" stroke-linecap="round"/>'
    if mood == "work":
        return f'<path d="M102,114 Q110,122 118,114" stroke="{NAVY}" stroke-width="2.6" fill="none" stroke-linecap="round"/>'
    if mood == "cheer":
        return f'<path d="M97,111 Q110,134 123,111 Z" fill="#FF6B81" stroke="{NAVY}" stroke-width="2.6" stroke-linejoin="round"/><path d="M103,121 Q110,128 117,121 Q110,117 103,121Z" fill="#fff" opacity=".85"/>'
    return f'<path d="M101,119 q3.2,-5 6.4,0 t6.4,0" stroke="{NAVY}" stroke-width="2.6" fill="none" stroke-linecap="round"/>'


def _clipboard():
    return f'''<g transform="rotate(-10 52 168)">
<rect x="30" y="140" width="40" height="52" rx="6" fill="#fff" stroke="{NAVY}" stroke-width="3"/>
<rect x="42" y="135" width="16" height="10" rx="3.5" fill="{NAVY}"/>
<path d="M37,157 l3.2,3.2 l5.8,-6.4" stroke="{MINT}" stroke-width="3" fill="none" stroke-linecap="round" stroke-linejoin="round"/><path d="M50,157 h13" stroke="#B9C9DD" stroke-width="3" stroke-linecap="round"/>
<path d="M37,170 l3.2,3.2 l5.8,-6.4" stroke="{MINT}" stroke-width="3" fill="none" stroke-linecap="round" stroke-linejoin="round"/><path d="M50,170 h13" stroke="#B9C9DD" stroke-width="3" stroke-linecap="round"/>
<path d="M37,183 l3.2,3.2 l5.8,-6.4" stroke="{BUG}" stroke-width="3" fill="none" stroke-linecap="round" stroke-linejoin="round" opacity=".0"/><path d="M50,183 h13" stroke="#B9C9DD" stroke-width="3" stroke-linecap="round"/><circle cx="40.5" cy="182.5" r="3.2" fill="none" stroke="#B9C9DD" stroke-width="2"/>
</g>'''


def _magnifier():
    return f'''<g>
<path d="M146,166 L160,134" stroke="{NAVY}" stroke-width="6" stroke-linecap="round"/>
<circle cx="170" cy="118" r="22" fill="#CFEFFF" fill-opacity=".55" stroke="{NAVY}" stroke-width="5"/>
<path d="M156,108 Q160,98 171,96" stroke="#fff" stroke-width="3.4" fill="none" stroke-linecap="round" opacity=".9"/>
{bugsy(171, 124, 0.62)}
</g>'''


def tess(mood: str = "hello", size: int = 180) -> str:
    mood = mood if mood in MOODS else "hello"
    uid = f"ts{mood}{size}"
    arms_up = mood == "cheer"
    right_arm = {
        "hello": "",
        "think": f'<path d="M142,152 Q148,160 146,166" stroke="{BLUE}" stroke-width="15" fill="none" stroke-linecap="round"/>' + _magnifier() + f'<circle cx="146" cy="166" r="9" fill="{SKIN}"/>',
        "work": f'<path d="M142,150 Q158,158 156,172" stroke="{BLUE}" stroke-width="15" fill="none" stroke-linecap="round"/><circle cx="156" cy="175" r="9" fill="{SKIN}"/>',
        "cheer": "",
        "oops": f'<path d="M142,150 Q150,126 134,118" stroke="{BLUE}" stroke-width="15" fill="none" stroke-linecap="round"/><circle cx="131" cy="116" r="9" fill="{SKIN}"/>',
    }[mood]
    left_arm = ("" if arms_up else
                f'<path d="M78,150 Q62,160 56,170" stroke="{BLUE}" stroke-width="15" fill="none" stroke-linecap="round"/>')
    prop_left = "" if arms_up else _clipboard()
    hand_left = "" if arms_up else f'<circle cx="64" cy="180" r="8.5" fill="{SKIN}"/>'
    extras = {
        "hello": _sparkle(184, 60, 8) + _sparkle(40, 52, 6, "#9ADFF5"),
        "think": f'<text x="22" y="60" font-family="Inter,Segoe UI,sans-serif" font-size="26" font-weight="800" fill="{TEAL}">?</text>' + _sparkle(188, 62, 6, "#9ADFF5"),
        "work": _sparkle(30, 70, 9) + _sparkle(190, 74, 7, "#9ADFF5") + _sparkle(176, 40, 5) ,
        "cheer": "".join(f'<rect x="{x}" y="{y}" width="6" height="10" rx="2" fill="{c}" transform="rotate({r} {x+3} {y+5})"/>'
                         for x, y, c, r in [(22, 40, MINT, 20), (60, 18, "#FFD86B", -25), (160, 16, BLUE, 30), (196, 44, "#FF8FA3", -15),
                                            (36, 108, TEAL, 40), (190, 112, "#FFD86B", -35)]) + _sparkle(98, 14, 7) + _sparkle(206, 80, 6, "#9ADFF5"),
        "oops": _sparkle(30, 66, 5, "#9ADFF5"),
    }[mood]
    bug = {
        "hello": bugsy(176, 214, 0.9, True, True),
        "think": bugsy(34, 212, 0.8, True),
        "work": bugsy(176, 214, 0.9, True, True),
        "cheer": bugsy(176, 212, 0.9, True, True),
        "oops": bugsy(166, 206, 1.15, False, True),
    }[mood]
    front = {
        "hello": f'<path d="M144,152 Q180,144 182,112" stroke="{BLUE}" stroke-width="15" fill="none" stroke-linecap="round"/><circle cx="182" cy="105" r="9.5" fill="{SKIN}"/>'
                 f'<path d="M172,88 q-2,-7 1,-11 M182,84 q0,-7 4,-10 M192,90 q3,-6 7,-7" stroke="{TEAL}" stroke-width="2.6" fill="none" stroke-linecap="round"/>',
        "cheer": f'<path d="M80,150 Q34,140 34,84" stroke="{BLUE}" stroke-width="15" fill="none" stroke-linecap="round"/><circle cx="34" cy="76" r="9.5" fill="{SKIN}"/>'
                 f'<path d="M142,150 Q186,140 186,84" stroke="{BLUE}" stroke-width="15" fill="none" stroke-linecap="round"/><circle cx="186" cy="76" r="9.5" fill="{SKIN}"/>',
    }.get(mood, "")
    sweat = f'<path d="M72,72 q-4,8 0,12 q4,-4 0,-12z" fill="#8ED8F8" stroke="{BLUE}" stroke-width="1.4"/>' if mood == "oops" else ""
    return f'''<svg width="{size}" height="{round(size*240/220)}" viewBox="0 0 220 240" xmlns="http://www.w3.org/2000/svg" role="img" aria-label="Tess, the TestX mascot ({mood})">
<defs>
<radialGradient id="tsIris" cx=".4" cy=".35" r=".8"><stop offset="0" stop-color="#5CC8FF"/><stop offset=".6" stop-color="{BLUE}"/><stop offset="1" stop-color="{TEAL}"/></radialGradient>
<linearGradient id="{uid}h" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="{LMINT}"/><stop offset=".55" stop-color="{MINT}"/><stop offset="1" stop-color="{DMINT}"/></linearGradient>
</defs>
<ellipse cx="110" cy="226" rx="52" ry="7" fill="{NAVY}" opacity=".12"/>
<!-- twin tails + back hair -->
<path d="M62,92 Q28,98 26,150 Q26,176 44,190 Q40,160 56,134 Z" fill="url(#{uid}h)"/>
<path d="M158,92 Q192,98 194,150 Q194,176 176,190 Q180,160 164,134 Z" fill="url(#{uid}h)"/>
<circle cx="60" cy="108" r="7" fill="{BLUE}"/><circle cx="160" cy="108" r="7" fill="{BLUE}"/>
<ellipse cx="110" cy="88" rx="58" ry="56" fill="{DMINT}"/>
<!-- legs + shoes -->
<rect x="92" y="184" width="13" height="26" rx="6" fill="{SKIN}"/><rect x="115" y="184" width="13" height="26" rx="6" fill="{SKIN}"/>
<path d="M86,214 q0,-9 11,-9 h12 q0,9 -6,11 h-17z" fill="{NAVY}"/><path d="M134,214 q0,-9 -11,-9 h-12 q0,9 6,11 h17z" fill="{NAVY}"/>
<!-- body: hoodie -->
<path d="M76,150 Q76,130 110,130 Q144,130 144,150 L148,192 Q110,200 72,192 Z" fill="{BLUE}"/>
<path d="M88,134 Q110,152 132,134" stroke="{DBLUE}" stroke-width="5" fill="none" stroke-linecap="round"/>
<path d="M104,142 v14 M116,142 v14" stroke="#fff" stroke-width="2.4" stroke-linecap="round"/>
<circle cx="110" cy="170" r="9.5" fill="#fff"/><path d="M105,170 l3.4,3.6 l6.2,-7" stroke="{MINT}" stroke-width="3" fill="none" stroke-linecap="round" stroke-linejoin="round"/>
{left_arm}{prop_left}{hand_left}
{right_arm}
<!-- head -->
<ellipse cx="110" cy="96" rx="50" ry="44" fill="{SKIN}"/>
<path d="M62,108 Q66,134 96,138" stroke="{SHADE}" stroke-width="3" fill="none" opacity=".0"/>
<ellipse cx="70" cy="112" rx="9" ry="5.4" fill="{BLUSH}" opacity=".5"/><ellipse cx="150" cy="112" rx="9" ry="5.4" fill="{BLUSH}" opacity=".5"/>
{_eyes(mood)}
{_mouth(mood)}
{sweat}
<!-- fringe -->
<path d="M58,92 C54,48 84,36 110,36 C138,36 168,50 162,92 C156,76 148,70 140,76 C134,62 124,60 116,72 C110,58 98,60 92,72 C84,64 70,68 58,92 Z" fill="url(#{uid}h)"/>
<path d="M80,48 Q96,38 118,40" stroke="{LMINT}" stroke-width="4" fill="none" stroke-linecap="round" opacity=".85"/>
<path d="M132,38 q10,-16 22,-12 q-2,12 -16,16" fill="{MINT}" stroke="{DMINT}" stroke-width="1.4"/>
<!-- headset -->
<path d="M56,96 Q54,40 110,34 Q166,40 164,96" stroke="{NAVY}" stroke-width="6" fill="none" stroke-linecap="round"/>
<rect x="46" y="86" width="14" height="26" rx="7" fill="{NAVY}"/><rect x="160" y="86" width="14" height="26" rx="7" fill="{NAVY}"/>
<rect x="49" y="92" width="5" height="14" rx="2.5" fill="{TEAL}"/><rect x="166" y="92" width="5" height="14" rx="2.5" fill="{TEAL}"/>
<path d="M168,112 Q168,130 140,126" stroke="{NAVY}" stroke-width="3.4" fill="none" stroke-linecap="round"/><circle cx="137" cy="125.5" r="4.6" fill="{NAVY}"/><circle cx="137" cy="125.5" r="1.8" fill="{TEAL}"/>
{front}
{extras}
{bug}
</svg>'''


def bugsy_svg(size: int = 56, happy: bool = True) -> str:
    return (f'<svg width="{size}" height="{size}" viewBox="-26 -30 52 52" xmlns="http://www.w3.org/2000/svg" role="img" aria-label="Bugsy the bug">'
            f'{bugsy(0, 0, 1.0, happy)}</svg>')
