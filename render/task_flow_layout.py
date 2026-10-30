"""Geometry for the vertical-swimlane flowcharts (renderer independent).

``build_layouts`` turns the task list of ONE order type into one or more
``Layout`` objects (panels). Both the PNG renderer (task_flow_image.py) and
the editable PowerPoint renderer (task_flow_pptx.py) draw from the same
Layout, so the picture in Word and the slides always match.

Rules
-----
* Tasks follow the Task Seq order strictly; every step connects to the next.
* Swimlane = Work Group (blank -> IT System); see mapping/flow_rules.yaml.
* Work Activity (same Seq, name starts with "WA-"): attached to its task with a DOUBLE-HEADED arrow.
* Parallel tasks (same Seq, names NOT starting with "WA-"): the flow forks into all of them and
  joins again before the next step.
* A Task Condition becomes a decision diamond (Yes -> task, No -> next step).
* Excluded tasks (External/System, ADD REWORK ACTIVITYn ...) are omitted.
"""
from __future__ import annotations

import textwrap
from dataclasses import dataclass, field
from typing import Iterable

from models.schemas import Task
from render import task_flow as tf

# ---- geometry (inches) -----------------------------------------------------
IMG_W = 6.2
TITLE_H = 0.36
LANE_HEAD_H = 0.5
TOP_PAD = 0.72
SLOT_H = 0.82
H_AVAIL = 8.3                      # drawable height of one Word page below the headings
MIN_SCALE = 0.85                   # never shrink a diagram below this
BOX_H = 0.5
DIA_H = 0.68
MAX_ROWS_PER_IMAGE = 9
TOP = TITLE_H + LANE_HEAD_H          # y where the lane body starts


@dataclass(eq=False)
class LNode:
    kind: str                 # task | decision | oval
    label: str
    lane: str
    row: int
    cx: float = 0.0
    cy: float = 0.0
    w: float = 0.0
    h: float = 0.0
    font: float = 6.4
    task_key: str = ""

    top = property(lambda s: s.cy - s.h / 2)
    bottom = property(lambda s: s.cy + s.h / 2)
    left = property(lambda s: s.cx - s.w / 2)
    right = property(lambda s: s.cx + s.w / 2)


@dataclass
class LEdge:
    points: list
    heads: str = "end"        # end | both | none
    label: str | None = None
    label_xy: tuple = (0.0, 0.0)
    label_below: bool = False


@dataclass
class Layout:
    width: float
    height: float
    title: str
    lanes: list
    lane_w: float
    nodes: list = field(default_factory=list)
    edges: list = field(default_factory=list)
    panel_no: int = 1
    total: int = 1
    scale: float = 1.0               # display scale in the Word page (1.0 = full width)


# ---------------------------------------------------------------- helpers ---
def wrap_text(label, width_in, base=6.4):
    chars = max(9, int(width_in * 17))
    lines = textwrap.wrap(label, chars, break_long_words=True) or [""]
    size = base
    if len(lines) > 3:
        size = base - 0.8
        lines = textwrap.wrap(label, max(10, int(width_in * 20)), break_long_words=True)
    if len(lines) > 4:
        lines = lines[:4]
        lines[-1] = lines[-1][:-1] + "…"
    return "\n".join(lines), size


def group_tasks(flow_tasks):
    groups = []
    for t in flow_tasks:
        if groups and groups[-1][0].seq == t.seq:
            groups[-1].append(t)
        else:
            groups.append([t])
    return groups


def classify(group):
    """-> (mains, work_activities). Same-Seq tasks whose name starts with WA- are work
    activities; the others are the main tasks (several = parallel tasks)."""
    mains = [t for t in group if not t.is_wa]
    was = [t for t in group if t.is_wa]
    if not mains:                       # all are WA: the first one is the main task
        mains, was = [group[0]], list(group[1:])
    return mains, was


def _rows_needed(group):
    counts = {}
    for t in group:
        counts[t.lane] = counts.get(t.lane, 0) + 1
    return max(counts.values())


def _weight(g):
    return _rows_needed(g) + (1 if g[0].condition else 0)


def split_balanced(groups, k):
    """Split the step groups into k consecutive parts of similar height."""
    total = sum(_weight(g) for g in groups)
    target = total / k
    parts, cur, acc = [], [], 0.0
    for g in groups:
        w = _weight(g)
        remaining_parts = k - len(parts)
        if cur and len(parts) < k - 1 and acc + w / 2 > target * (len(parts) + 1) - 1e-9 and remaining_parts > 1:
            parts.append(cur)
            cur = []
        cur.append(g)
        acc += w
    if cur:
        parts.append(cur)
    return parts


def lane_order(groups):
    rules = tf.load_flow_rules()
    present = []
    for g in groups:
        for t in g:
            if t.lane not in present:
                present.append(t.lane)
    ordered = []
    for want in rules["lane_order"]:
        for p in present:
            if p.casefold() == str(want).casefold() and p not in ordered:
                ordered.append(p)
    return ordered + [p for p in present if p not in ordered]


def _exit_index(group, order_type):
    for o in tf.load_flow_rules()["exit_task_overrides"]:
        if str(o.get("order_type", "")).upper() == order_type.upper() and str(o.get("seq")) == group[0].seq:
            for i, t in enumerate(group):
                if tf._norm_ws(t.task_name).casefold() == tf._norm_ws(o.get("task", "")).casefold():
                    return i
    return 0


# ----------------------------------------------------------------- layout ---
def build_panel_layout(groups, panel_no, total, order_type, title, lanes, slot_h=SLOT_H):
    """Geometry of one diagram part. `lanes` = ALL swimlanes of the order type, so every
    part of a split diagram shows the same lanes."""
    lane_w = IMG_W / len(lanes)
    box_w = min(1.25, lane_w - 0.26)
    lane_x = lambda name: lanes.index(name) * lane_w + lane_w / 2
    li = {ln: k for k, ln in enumerate(lanes)}
    n = len(groups)
    cls = [classify(g) for g in groups]                    # (mains, work activities)
    has_dec = [bool(g[0].condition) for g in groups]

    # tasks of one step go in the same row when their lanes differ
    row_tasks = []
    for g in groups:
        rows = []
        for t in g:
            for r in rows:
                if all(x.lane != t.lane for x in r):
                    r.append(t)
                    break
            else:
                rows.append([t])
        row_tasks.append(rows)

    # Decision placement. Preferred (reference charts): diamond in the lane of the NEXT step,
    # on the task's row: "Yes" runs sideways to the task, "No" drops to the next step.
    # Otherwise (and always for parallel tasks) the diamond sits on its own row above the
    # task(s) and "No" loops round the lane edge.
    entry_lane, dec_lane, dec_ref = [None] * n, [None] * n, [False] * n
    for i in range(n - 1, -1, -1):
        mains = cls[i][0]
        primary = mains[0]
        if has_dec[i]:
            cand = entry_lane[i + 1] if i + 1 < n else None
            ok = False
            if cand is not None and len(mains) == 1:
                first_row = row_tasks[i][0]
                lo, hi = sorted((li[cand], li[primary.lane]))
                ok = (all(x.lane != cand for x in first_row)
                      and not any(lo < li[x.lane] < hi for x in first_row))
            dec_ref[i] = ok
            dec_lane[i] = cand if ok else primary.lane
            entry_lane[i] = dec_lane[i]
        else:
            entry_lane[i] = primary.lane

    nodes, tnode = [], {}
    dec_nodes = [None] * n
    entries, exits = [None] * n, [None] * n
    row = 0
    for i, g in enumerate(groups):
        mains, was = cls[i]
        rows = row_tasks[i]
        dec_row = row
        if has_dec[i] and not dec_ref[i]:
            row += 1
        for r_i, r in enumerate(rows):
            for t in r:
                label = t.label
                if t is not g[0] and t.condition:
                    label = f"{t.label}\n({tf.condition_label(t.condition)[:-1]})"
                nd = LNode("task", label, t.lane, row + r_i, task_key=f"{t.seq}|{t.task_name}")
                tnode[id(t)] = nd
                nodes.append(nd)
        if has_dec[i]:
            dn = LNode("decision", tf.condition_label(g[0].condition), dec_lane[i], dec_row)
            dec_nodes[i] = dn
            nodes.append(dn)
        row += len(rows)
        mnodes = [tnode[id(m)] for m in mains]
        entries[i] = [dec_nodes[i]] if dec_nodes[i] else mnodes
        if len(mains) > 1:                                  # parallel tasks: all of them exit
            exits[i] = mnodes
        else:
            ex = tnode[id(g[_exit_index(g, order_type)])]
            same = [tnode[id(t)] for t in g if tnode[id(t)].lane == ex.lane]
            exits[i] = [max(same, key=lambda nd: nd.row)]
    end_row = row
    height = TOP + TOP_PAD + (end_row + 1) * slot_h + 0.15

    for nd in nodes:
        nd.cx = lane_x(nd.lane)
        nd.cy = TOP + TOP_PAD + nd.row * slot_h + slot_h / 2 - 0.12
        nd.w = box_w
        nd.h = DIA_H if nd.kind == "decision" else BOX_H
    for i in range(n):                                     # centre a diamond over its parallel tasks
        if dec_nodes[i] and len(cls[i][0]) > 1:
            dec_nodes[i].cx = sum(tnode[id(m)].cx for m in cls[i][0]) / len(cls[i][0])
    for nd in nodes:
        if nd.kind == "decision":
            nd.label, nd.font = wrap_text(nd.label, nd.w * 0.62, 6.0)
        else:
            nd.label, nd.font = wrap_text(nd.label, nd.w - 0.1)

    first = entries[0]
    start = LNode("oval", "Start" if panel_no == 1 else "Continue", first[0].lane, -1)
    start.cx = sum(e.cx for e in first) / len(first) if len(first) > 1 else first[0].cx
    start.cy, start.w, start.h = TOP + 0.4, 0.95, 0.34
    last = exits[-1]
    end = LNode("oval", "End" if panel_no == total else "Continue", last[0].lane, end_row)
    end.cx = sum(e.cx for e in last) / len(last)
    end.cy = TOP + TOP_PAD + end_row * slot_h + slot_h / 2 - 0.12
    end.w, end.h = 0.95, 0.34
    for o in (start, end):
        o.font = 6.4
    nodes += [start, end]

    lay = Layout(IMG_W, height, title, lanes, lane_w, nodes, [], panel_no, total)

    def hits_v(x, y0, y1, ignore):
        lo, hi = min(y0, y1), max(y0, y1)
        return any(b not in ignore and b.left - 0.02 <= x <= b.right + 0.02 and not (hi < b.top or lo > b.bottom)
                   for b in nodes)

    def draw(points, heads="end", label=None, label_at=0):
        e = LEdge(points, heads, label)
        if label:
            px, py = points[label_at]
            going_down = abs(points[1][1] - points[0][1]) > 0.04
            e.label_below = going_down
            e.label_xy = (px + 0.06, py + 0.07) if going_down else (px + 0.05, py - 0.03)
        lay.edges.append(e)

    def lane_edge_x(nd, right=True):
        k = lanes.index(nd.lane)
        return (k + 1) * lane_w - 0.05 if right else k * lane_w + 0.05

    def channel(src, dst, y_in, label=None):
        gx = lane_edge_x(src, True)
        draw([(src.right, src.cy), (gx, src.cy), (gx, y_in), (dst.right, y_in)], "end", label, 0)

    def route(src, dst, join=False, label=None):
        sx, sb = src.cx, src.bottom
        if join:
            jy = dst.cy - 0.15 if dst.kind == "task" else dst.cy
            if abs(sx - dst.cx) < 1e-6 or hits_v(sx, sb, jy, {src, dst}):
                return channel(src, dst, jy, label)
            ex = dst.right if sx > dst.cx else dst.left
            return draw([(sx, sb), (sx, jy), (ex, jy)], "end", label, 0)
        ex, et = dst.cx, dst.top
        if abs(sx - ex) < 1e-6:
            if not hits_v(sx, sb, et, {src, dst}):
                return draw([(sx, sb), (ex, et)], "end", label, 0)
            return channel(src, dst, dst.cy - 0.15, label)
        ys = sb + 0.17
        if ys > et - 0.06:
            ys = (sb + et) / 2
        if hits_v(ex, ys, et, {src, dst}):
            return channel(src, dst, dst.cy - 0.15, label)
        draw([(sx, sb), (sx, ys), (ex, ys), (ex, et)], "end", label, 1 if label else 0)

    def bus_connect(srcs, dsts, label=None):
        """Fork / join: every src drops to a common bus line, which feeds every dst."""
        ymax = max(s.bottom for s in srcs)
        ymin = min(d.top for d in dsts)
        yj = ymax + 0.17 if ymin - ymax > 0.45 else (ymax + ymin) / 2
        blocked = any(hits_v(s.cx, s.bottom, yj, {s}) for s in srcs) or \
            any(hits_v(d.cx, yj, d.top, {d}) for d in dsts) or yj > ymin - 0.04
        if blocked:                                          # fall back to one elbow per pair
            for s_ in srcs:
                for d_ in dsts:
                    route(s_, d_, False, label if (s_ is srcs[0] and d_ is dsts[0]) else None)
            return
        for k, s_ in enumerate(srcs):
            draw([(s_.cx, s_.bottom), (s_.cx, yj)], "none", label if k == 0 else None, 0)
        xs = [s_.cx for s_ in srcs] + [d_.cx for d_ in dsts]
        if max(xs) - min(xs) > 1e-6:
            draw([(min(xs), yj), (max(xs), yj)], "none")
        for d_ in dsts:
            draw([(d_.cx, yj), (d_.cx, d_.top)], "end")

    def connect(srcs, dsts, join=False, label=None):
        if len(srcs) == 1 and len(dsts) == 1:
            route(srcs[0], dsts[0], join, label)
        else:
            bus_connect(srcs, dsts, label)

    def no_channel(d, entry, label="No"):
        """'No' branch: leave the diamond sideways, run down a clear channel beside every box
        between the diamond and the next step, then enter that step from the side."""
        lo_r, hi_r = sorted((d.row, entry.row))
        between = [n_ for n_ in nodes if n_.kind != "oval" and lo_r <= n_.row <= hi_r and n_ is not d]
        right_x = max([d.right, entry.right] + [n_.right for n_ in between]) + 0.08
        left_x = min([d.left, entry.left] + [n_.left for n_ in between]) - 0.08
        right_x = min(right_x, IMG_W - 0.04)
        left_x = max(left_x, 0.04)
        if abs(entry.cx - d.cx) > 0.05:
            go_right = entry.cx > d.cx
        else:
            go_right = (right_x - d.right) <= (d.left - left_x)
        if go_right:
            draw([(d.right, d.cy), (right_x, d.cy), (right_x, entry.cy), (entry.right, entry.cy)], "end", label, 0)
        else:
            draw([(d.left, d.cy), (left_x, d.cy), (left_x, entry.cy), (entry.left, entry.cy)], "end", label, 0)

    # nodes that receive a side "join"/"No" line: keep the double arrow clear of it
    busy = {id(t) for i in range(1, n) if dec_nodes[i - 1] for t in entries[i]}

    connect([start], entries[0])
    for i, g in enumerate(groups):
        mains, was = cls[i]
        pos = {id(t): k for k, t in enumerate(g)}
        for w in was:                                        # Work Activity: DOUBLE-HEADED arrow
            before = [m for m in mains if pos[id(m)] < pos[id(w)]]
            anchor = tnode[id(before[-1] if before else mains[0])]
            wn = tnode[id(w)]
            if wn is anchor:
                continue
            if anchor.row == wn.row:
                y = anchor.cy + (0.14 if (id(anchor) in busy or id(wn) in busy) else 0.0)
                left, right = (anchor, wn) if anchor.cx < wn.cx else (wn, anchor)
                draw([(left.right, y), (right.left, y)], "both")
            elif anchor.lane == wn.lane:
                a_, b_ = (anchor, wn) if anchor.row < wn.row else (wn, anchor)
                draw([(a_.cx, a_.bottom), (b_.cx, b_.top)], "both")
            else:
                a_, b_ = (anchor, wn) if anchor.row < wn.row else (wn, anchor)
                ys = a_.bottom + 0.17
                draw([(a_.cx, a_.bottom), (a_.cx, ys), (b_.cx, ys), (b_.cx, b_.top)], "both")
        if dec_nodes[i]:                                     # Yes branch
            d = dec_nodes[i]
            mnodes = [tnode[id(m)] for m in mains]
            if dec_ref[i]:
                t0 = mnodes[0]
                if t0.cx > d.cx:
                    draw([(d.right, d.cy), (t0.left, t0.cy)], "end", "Yes", 0)
                else:
                    draw([(d.left, d.cy), (t0.right, t0.cy)], "end", "Yes", 0)
            else:
                for k, m in enumerate(mnodes):               # fork to every parallel task
                    draw([(d.cx, d.bottom), (m.cx, m.top)], "end", "Yes" if k == 0 else None, 0)

    for i in range(1, n):
        prev_exits, prev_dec, ent = exits[i - 1], dec_nodes[i - 1], entries[i]
        if prev_dec and dec_ref[i - 1]:
            connect([prev_dec], ent, False, "No")            # No drops straight to the next step
            connect(prev_exits, ent, True)                   # the task joins from the side
        elif prev_dec:
            connect(prev_exits, ent, False)
            for e_ in ent:
                no_channel(prev_dec, e_)                     # No loops round the lane edge
        else:
            connect(prev_exits, ent, False)
    if dec_nodes[-1]:
        no_channel(dec_nodes[-1], end)
    connect(exits[-1], [end], False)
    return lay


def build_layouts(tasks: Iterable[Task], order_type: str, product: str = ""):
    """Return (layouts, warnings) for one order type.

    The diagram is fitted to ONE Word page when it can be (tighter rows, then a slight
    scale-down); otherwise it is split into balanced parts. Every part shows ALL swimlanes.
    """
    flow_tasks, warnings = tf.prepare_flow_tasks_ex(tasks)
    if not flow_tasks:
        return [], warnings
    groups = group_tasks(flow_tasks)
    lanes = lane_order(groups)
    base = f"{product}-{order_type} workflow" if product else f"{order_type} workflow"

    def one(gs, i, k, pitch):
        title = base + (f" (part {i} of {k})" if k > 1 else "")
        return build_panel_layout(gs, i, k, order_type, title, lanes, pitch)

    full = one(groups, 1, 1, SLOT_H)
    if full.height <= H_AVAIL:
        return [full], warnings
    tight = one(groups, 1, 1, 0.70)
    if tight.height <= H_AVAIL:
        return [tight], warnings
    if H_AVAIL / tight.height >= MIN_SCALE:
        tight.scale = H_AVAIL / tight.height
        return [tight], warnings
    for k in range(2, len(groups) + 1):
        for pitch in (0.78, 0.70):
            parts = split_balanced(groups, k)
            if len(parts) != k:
                continue
            lays = [one(gs, i, k, pitch) for i, gs in enumerate(parts, start=1)]
            if all(l.height <= H_AVAIL for l in lays):
                return lays, warnings
    lays = [one(gs, i, len(parts), 0.70) for i, gs in enumerate(split_balanced(groups, len(groups)), start=1)]
    return lays, warnings
