"""Shape-built chart exhibits: waterfall, funnel, mekko, bar-mekko,
tornado, football-field.

Split out of builders_consulting.py (file-size budget); registered into
builders.LAYOUT_MAP via builders_consulting.LAYOUTS at import time.
All geometry is in design inches on the 13.33 x 7.5 canvas; builders'
wrappers rescale to the actual slide size.
"""
import re

from pptx.enum.shapes import MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

import builders as B
import shape_groups
from charts import round_to_sum
from helpers import label_color_on as _label_color_on
from numfmt import (STYLE_COMPACT, fmt_abs, fmt_pct, label_decimals,
                    pct_decimals, round_consistent)
from helpers import warn

NEG_HEX = "D9655B"   # waterfall decreases / down-trends (readable on all palettes)
POS_HEX = None       # increases use pal accent2

_NUM_RX = re.compile(r"-?(?:\d{1,3}(?:,\d{3})*|\d+)(?:\.\d+)?")


BREAK_THRESHOLD = 0.75   # break only when the detail window starts this high
                         # (i.e. the bottom 75%+ of the axis carries no detail)
BREAK_ZONE_H = 0.5       # compressed 0..b zone height (design inches)


def _broken_scale(top, bottom, peak, lo_detail, enabled):
    """(vy, break_y): vy maps a value to its y coordinate; break_y is the
    boundary of the compressed zone, or None when no break applies.

    With '- Axis-Break: on' and a detail window that starts above
    BREAK_THRESHOLD * peak, the range 0..b (b = 0.9 * lo_detail) compresses
    into a fixed BREAK_ZONE_H strip above the baseline so the differences
    between bars become visible; b..peak gets the remaining height."""
    span = bottom - top
    if (not enabled or peak <= 0 or lo_detail <= BREAK_THRESHOLD * peak
            or span <= BREAK_ZONE_H * 2):
        scale = span / peak if peak else 0
        return (lambda v: bottom - v * scale), None
    b = lo_detail * 0.9
    rest = span - BREAK_ZONE_H

    def vy(v):
        if v <= b:
            return bottom - (v / b) * BREAK_ZONE_H
        return bottom - BREAK_ZONE_H - (v - b) / (peak - b) * rest
    return vy, bottom - BREAK_ZONE_H


def _break_marks(slide, pal, x, w, break_y):
    """Double diagonal slash in bg color across a bar at the break boundary."""
    for dy in (-0.045, 0.045):
        mark = B.add_rect(slide, x - 0.06, break_y + dy - 0.018, w + 0.12,
                          0.036, pal["bg"])
        mark.rotation = 8


def _axis_break_on(p):
    return str(p.get("axis_break", "")).strip().lower() in ("on", "true", "yes")


class _LabelLane:
    """Greedy label collision avoidance for shape-built exhibits.

    Remembers every placed label rect; place() returns the label's y, nudged
    away from already-placed labels in `step` increments (default upward).
    Returns None when no clear position exists within max_shift — the caller
    should then drop the label rather than print text on text."""

    def __init__(self):
        self._rects = []

    EPS = 0.05  # frames touching by less than this aren't a real collision
                # (centered glyphs never reach the frame edge)

    def _hits(self, x, y, w, h):
        e = self.EPS
        return any(x + e < px + pw and px + e < x + w
                   and y + e < py + ph and py + e < y + h
                   for px, py, pw, ph in self._rects)

    def place(self, x, y, w, h, direction=-1, step=0.30, max_shift=0.95):
        shift = 0.0
        while self._hits(x, y, w, h):
            shift += step
            if shift > max_shift:
                return None
            y += direction * step
        self._rects.append((x, y, w, h))
        return y


# ── waterfall ────────────────────────────────────────────────────────────────
def _row_total(v):
    """A delta row's magnitude: multi-series rows sum their segments."""
    return sum(v) if isinstance(v, list) else v


def _shade(hex_color, bg_hex, t):
    """Lerp hex_color toward bg_hex by t (segment shading on stacked bars)."""
    a = [int(hex_color[i:i + 2], 16) for i in (0, 2, 4)]
    b = [int(bg_hex[i:i + 2], 16) for i in (0, 2, 4)]
    return "".join(f"{round(x + (y - x) * t):02X}" for x, y in zip(a, b))


def build_waterfall_slide(prs, p, pal, ctx):
    """Value bridge: first bar = start, middle bars = +/- deltas, a 'total'
    row (or the computed end) closes the bridge.

    Extensions: a '- Label: subtotal' row inserts a full bar at the running
    total; with **Series:** names, '- Label: v1, v2' rows stack segments
    within one delta bar (all segments must share the delta's sign)."""
    slide = B._blank_slide(prs, pal, pal["bg"])
    B._heading(slide, p, pal)
    body_start = len(slide.shapes)
    data = [(lbl, val) for lbl, val in p.get("data", [])]
    if len(data) < 2:
        warn("waterfall has <2 data rows; rendering bullet-list")
        return B.build_bullet_slide(prs, p, pal, ctx)

    has_total = data[-1][1] == "total"
    middle = data[1:-1] if has_total else data[1:]
    start_val = _row_total(data[0][1])
    # bar = (label, lo, hi, color, display_value, kind); endpoints/subtotals
    # show their plain running value, deltas show the SIGNED change
    bars = [(data[0][0], 0, start_val, pal["accent1"], start_val, "end")]
    segments = {}  # bar index -> list of segment values (stacked rows)
    cum = start_val
    for label, v in middle:
        if v == "subtotal":
            bars.append((label, 0, cum, pal["accent1"], cum, "subtotal"))
            continue
        delta = _row_total(v)
        lo, hi = (cum, cum + delta) if delta >= 0 else (cum + delta, cum)
        color = pal["accent2"] if delta >= 0 else NEG_HEX
        if isinstance(v, list) and len(v) > 1:
            segments[len(bars)] = v
        bars.append((label, lo, hi, color, delta, "delta"))
        cum += delta
    end_val = cum
    bars.append(((data[-1][0] if has_total else "Total"), 0, end_val,
                 pal["accent1"], end_val, "end"))

    top, bottom = 1.9, 6.15
    peak = max(max(b[1], b[2]) for b in bars) or 1
    n = len(bars)
    gap = 0.18
    bar_w = (11.9 - gap * (n - 1)) / n

    # Optional axis break: when every interesting level sits high above zero,
    # compress the empty 0..b zone so the deltas become readable.
    detail_levels = [b[2] if b[5] != "delta" else min(b[1], b[2]) for b in bars]
    vy, break_y = _broken_scale(top, bottom, peak, min(detail_levels),
                                _axis_break_on(p))

    # connector between bar i and i+1 sits at the RUNNING TOTAL after bar i
    # (bottom of a decrease bar, top of an increase bar — never floats)
    cum_after, running = [], start_val
    for label, lo, hi, color, shown, kind in bars:
        if kind == "delta":
            running += shown
        cum_after.append(running)

    label_vals = _bridge_labels(bars)
    lane = _LabelLane()
    if segments:  # legend for the stacked segments, named by **Series:**
        names = [s.strip() for s in p.get("series", "").split(",") if s.strip()]
        lx = 0.75
        for si, name in enumerate(names):
            text_w = min(len(name) * 0.085 + 0.3, 2.0)
            B.add_rect(slide, lx, 1.72, 0.18, 0.18,
                       _shade(pal["accent2"], pal["bg"], 0.45 * (si % 2)))
            B.add_tb(slide, name, lx + 0.26, 1.64, text_w, 0.32, size=12,
                     color=pal["text_muted"], font=pal["font_label"])
            lane.place(lx, 1.64, 0.26 + text_w, 0.32)  # block value labels
            lx += 0.32 + text_w
    for i, (label, lo, hi, color, shown, kind) in enumerate(bars):
        x = 0.7 + i * (bar_w + gap)
        y0 = vy(hi)
        h = max(vy(lo) - vy(hi), 0.04)
        if i in segments:
            # stack same-sign segments between lo and hi, alternating shade
            base = lo if shown >= 0 else hi
            direction = 1 if shown >= 0 else -1
            level = base
            for si, seg in enumerate(segments[i]):
                seg_lo, seg_hi = sorted((level, level + direction * abs(seg)))
                B.add_rect(slide, x, vy(seg_hi), bar_w,
                           max(vy(seg_lo) - vy(seg_hi), 0.02),
                           _shade(color, pal["bg"], 0.45 * (si % 2)),
                           line_hex=pal["bg"], line_pt=0.75)
                level += direction * abs(seg)
        else:
            B.add_rect(slide, x, y0, bar_w, h, color)
        if break_y is not None and lo < min(detail_levels) * 0.9 < hi:
            _break_marks(slide, pal, x, bar_w, break_y)
        ly = lane.place(x - 0.1, y0 - 0.34, bar_w + 0.2, 0.3)
        if ly is None:
            warn(f"waterfall: value label for {label!r} suppressed — "
                 "no collision-free position")
        else:
            B.add_tb(slide, _fmt_num(label_vals[i], signed=kind == "delta"),
                     x - 0.1, ly, bar_w + 0.2, 0.3, size=12, bold=True,
                     color=pal["text"], align=PP_ALIGN.CENTER,
                     font=pal["font_body"])
        B.add_tb(slide, label, x - 0.12, bottom + 0.10, bar_w + 0.24, 0.75,
                 size=11, color=pal["text_muted"], align=PP_ALIGN.CENTER,
                 font=pal["font_label"])
        if i > 0:
            level_y = vy(cum_after[i - 1])
            B.add_rect(slide, x - gap, level_y - 0.009, gap, 0.018,
                       pal["text_muted"])
    if p.get("bracket"):
        _bracket_on_waterfall(slide, pal, p, bars, vy, bar_w, gap)
    B.add_rect(slide, 0.7, bottom, 11.9, 0.02, pal["surface"])
    shape_groups.group_shapes(slide, list(slide.shapes)[body_start:])
    return slide


def _delta_bars(start, deltas, pal):
    """Legacy helper (kept for imports): simple delta bars, no stacking."""
    bars, cum = [], start
    for label, v in deltas:
        lo, hi = (cum, cum + v) if v >= 0 else (cum + v, cum)
        bars.append((label, lo, hi,
                     pal["accent2"] if v >= 0 else NEG_HEX, v))
        cum += v
    return bars


def _bridge_labels(bars):
    """Display values for a waterfall's bars, deltas rounded think-cell style.

    The delta labels are rounded together via numfmt.round_consistent so the
    printed +/- steps add up to the printed end-minus-start change; endpoint
    and subtotal bars keep their own running value.  Geometry still uses the
    unrounded values, so a bar is never redrawn to match its label.
    """
    idx = [i for i, b in enumerate(bars) if b[5] == "delta"]
    shown = [b[4] for b in bars]
    if not idx:
        return shown
    deltas = [bars[i][4] for i in idx]
    rounded = round_consistent(deltas, sum(deltas),
                               label_decimals(deltas, STYLE_COMPACT))
    for i, value in zip(idx, rounded):
        shown[i] = value
    return shown


def _fmt_num(v, signed=False):
    """Compact bar-label number under the deck-wide number-format policy."""
    return fmt_abs(v, signed=signed, style=STYLE_COMPACT)


def _bracket_on_waterfall(slide, pal, p, bars, vy, bar_w, gap):
    """Difference bracket between two named bars: ┌────┐ + centered label.
    `vy` maps a value to its y coordinate (axis-break aware)."""
    spec = p.get("bracket", "")
    parts = [s.strip().strip('"') for s in spec.split(",") if s.strip()]
    if len(parts) < 2:
        warn(f"Bracket needs two bar labels, got: {spec!r}")
        return
    labels = [b[0].lower() for b in bars]
    try:
        ia, ib = labels.index(parts[0].lower()), labels.index(parts[1].lower())
    except ValueError:
        warn(f"Bracket labels not found among bars: {spec!r}")
        return
    (ha, va), (hb, vb) = (bars[ia][2], bars[ia][4]), (bars[ib][2], bars[ib][4])
    xa = 0.7 + ia * (bar_w + gap) + bar_w / 2
    xb = 0.7 + ib * (bar_w + gap) + bar_w / 2
    y = max(min(vy(ha), vy(hb)) - 0.66, 1.55)
    label = parts[2] if len(parts) > 2 else (
        fmt_pct((vb - va) / va * 100, signed=True) if va
        else _fmt_num(vb - va, signed=True))
    B.add_rect(slide, xa, y, xb - xa, 0.018, pal["text_muted"])      # beam
    B.add_rect(slide, xa, y, 0.018, 0.14, pal["text_muted"])          # left tick
    B.add_rect(slide, xb - 0.018, y, 0.018, 0.14, pal["text_muted"])  # right tick
    B.add_tb(slide, label, xa, y - 0.36, xb - xa, 0.32, size=13, bold=True,
             color=pal["text"], align=PP_ALIGN.CENTER, font=pal["font_body"])



# ── funnel ───────────────────────────────────────────────────────────────────
def build_funnel_slide(prs, p, pal, ctx):
    data = [(lbl, v) for lbl, v in p.get("data", [])
            if isinstance(v, (int, float))]
    if len(data) < 2:
        warn("funnel needs a **Data:** block with 2+ numeric rows")
        return B.build_bullet_slide(prs, p, pal, ctx)
    slide = B._blank_slide(prs, pal, pal["bg"])
    B._heading(slide, p, pal)
    body_start = len(slide.shapes)
    n = len(data)
    top_v = data[0][1] or 1
    row_h = min(0.85, 4.7 / n)
    label_w = 3.0  # stage labels live left of the bars (never overflow them)
    bar_area = 11.9 - label_w - 1.2
    for i, (label, v) in enumerate(data):
        w = max(1.0, bar_area * (v / top_v))
        x = 0.7 + label_w + 0.2 + (bar_area - w) / 2
        y = 1.95 + i * (row_h + 0.18)
        B.add_tb(slide, label, 0.7, y + row_h / 2 - 0.18, label_w, 0.4,
                 size=13, bold=i == 0, color=pal["text"],
                 align=PP_ALIGN.RIGHT, font=pal["font_body"])
        B.add_rect(slide, x, y, w, row_h,
                   pal["accent1"] if i == 0 else pal["surface"],
                   line_hex=pal["accent1"], line_pt=0 if i == 0 else 1.0)
        B.add_tb(slide, _fmt_plain(v), x, y + row_h / 2 - 0.18, w, 0.4, size=14,
                 bold=True, color=pal["bg_deep"] if i == 0 else pal["accent1"],
                 align=PP_ALIGN.CENTER, font=pal["font_body"])
        if i:
            pct = v / data[i - 1][1] * 100 if data[i - 1][1] else 0
            B.add_tb(slide, f"↳ {fmt_pct(pct)}", x + w + 0.15, y + row_h / 2 - 0.16,
                     1.1, 0.32, size=11, color=pal["text_muted"],
                     font=pal["font_label"])
    shape_groups.group_shapes(slide, list(slide.shapes)[body_start:])
    return slide


def _fmt_plain(v):
    """Plain value label under the deck-wide number-format policy."""
    return fmt_abs(v)



# ── mekko (marimekko) ────────────────────────────────────────────────────────
def build_mekko_slide(prs, p, pal, ctx):
    """Variable-width 100%-stacked columns: column width ∝ column total,
    segment height = share of column. Needs multi-series **Data:** rows."""
    names = [n.strip() for n in p.get("series", "").split(",") if n.strip()]
    data = [(lbl, v) for lbl, v in p.get("data", []) if isinstance(v, list)]
    if not names or not data:
        warn("mekko needs **Series:** names and multi-series **Data:** rows")
        return B.build_bullet_slide(prs, p, pal, ctx)
    slide = B._blank_slide(prs, pal, pal["bg"])
    B._heading(slide, p, pal)
    body_start = len(slide.shapes)

    accents = [pal["accent1"], pal["accent2"], pal["accent3"], NEG_HEX]
    # legend
    lx = 0.75
    for i, name in enumerate(names):
        B.add_rect(slide, lx, 1.72, 0.18, 0.18, accents[i % len(accents)])
        B.add_tb(slide, name, lx + 0.26, 1.64, 1.9, 0.32, size=12,
                 color=pal["text_muted"], font=pal["font_label"])
        lx += 0.32 + min(len(name) * 0.085 + 0.3, 2.0)

    top, bottom = 2.25, 6.05
    height = bottom - top
    gap = 0.06
    totals = [sum(v) for _, v in data]
    grand = sum(totals) or 1
    plot_w = 11.9 - gap * (len(data) - 1)
    x = 0.7
    for (label, vals), total in zip(data, totals):
        col_w = plot_w * (total / grand)
        # largest-remainder so each column's segment labels sum to 100
        pcts = (round_to_sum([v / total * 100 for v in vals],
                             decimals=pct_decimals())
                if total else [0] * len(vals))
        y = bottom
        for i, v in enumerate(vals):
            seg_h = height * (v / total) if total else 0
            y -= seg_h
            fill = accents[i % len(accents)]
            B.add_rect(slide, x, y, col_w, seg_h, fill,
                       line_hex=pal["bg"], line_pt=0.75)
            if seg_h >= 0.34 and col_w >= 0.9:
                B.add_tb(slide, fmt_pct(pcts[i]), x, y + seg_h / 2 - 0.16,
                         col_w, 0.32, size=14, bold=True,
                         color=_label_color_on(fill, pal),
                         align=PP_ALIGN.CENTER, font=pal["font_body"])
        B.add_tb(slide, _fmt_plain(total), x - 0.1, top - 0.36, col_w + 0.2, 0.3,
                 size=12, bold=True, color=pal["text"],
                 align=PP_ALIGN.CENTER, font=pal["font_body"])
        B.add_tb(slide, label, x - 0.12, bottom + 0.08, col_w + 0.24, 0.66,
                 size=11, color=pal["text_muted"], align=PP_ALIGN.CENTER,
                 font=pal["font_label"])
        x += col_w + gap
    from chart_options import draw_total_line, switch_on
    if switch_on(p.get("total_line")):   # `- Total-Line: on` (Cycle 7)
        sx, sy, sw, _ = B._sc(0.7, top, 11.9, 0)
        draw_total_line(slide, pal, sx, sy, sw)
    shape_groups.group_shapes(slide, list(slide.shapes)[body_start:])
    return slide



# ── bar mekko (profit pool: width = size, height = value) ───────────────────
def build_bar_mekko_slide(prs, p, pal, ctx):
    slide = B._blank_slide(prs, pal, pal["bg"])
    B._heading(slide, p, pal)
    body_start = len(slide.shapes)
    rows = []
    for b in p.get("bars", []):
        try:
            size, value = float(b["size"]), float(b["value"])
        except (KeyError, ValueError):
            warn(f"bar-mekko row needs numeric Size and Value: {b}")
            continue
        if size <= 0 or value < 0:
            warn(f"bar-mekko row needs Size>0 and Value>=0: {b}")
            continue
        rows.append((b.get("label", "?"), size, value))
    if len(rows) < 2:
        warn("bar-mekko has <2 valid bars; rendering bullet-list")
        return B.build_bullet_slide(prs, p, pal, ctx)

    L, bottom, total_w, top = 0.7, 6.15, 11.9, 2.25
    gap = 0.08
    total_size = sum(s for _, s, _ in rows) or 1
    vmax = max(v for _, _, v in rows) or 1
    vmin = min(v for _, _, v in rows)
    vy, break_y = _broken_scale(top, bottom, vmax, vmin, _axis_break_on(p))
    accents = [pal["accent1"], pal["accent2"], pal["accent3"]]

    x = L
    lane = _LabelLane()
    for i, (label, size, value) in enumerate(rows):
        w = (total_w - gap * (len(rows) - 1)) * size / total_size
        h = max(bottom - vy(value), 0.04)
        B.add_rect(slide, x, bottom - h, w, h, accents[i % len(accents)])
        if break_y is not None:
            _break_marks(slide, pal, x, w, break_y)
        # Cap the label frame near the bar's center — a full-width frame on a
        # wide bar would otherwise force needless nudges on its neighbors.
        lw = min(w + 0.2, 1.4)
        lx_label = x + w / 2 - lw / 2
        ly = lane.place(lx_label, bottom - h - 0.34, lw, 0.3)
        if ly is None:
            warn(f"bar-mekko: value label {_fmt_num(value)!r} suppressed — "
                 "no collision-free position")
        else:
            B.add_tb(slide, _fmt_num(value), lx_label, ly, lw,
                     0.3, size=12, bold=True, color=pal["text"],
                     align=PP_ALIGN.CENTER, font=pal["font_body"])
        B.add_tb(slide, f"{label}\n{_fmt_num(size)}", x - 0.1, bottom + 0.10,
                 w + 0.2, 0.75, size=11, color=pal["text_muted"],
                 align=PP_ALIGN.CENTER, font=pal["font_label"])
        x += w + gap
    B.add_rect(slide, L, bottom, total_w, 0.02, pal["surface"])
    shape_groups.group_shapes(slide, list(slide.shapes)[body_start:])
    return slide



# ── tornado (sensitivity) ────────────────────────────────────────────────────
def build_tornado_slide(prs, p, pal, ctx):
    """Sensitivity tornado built from shapes. Requires **Series:** Low, High
    (2 names) so '- Driver: -12, +18' rows parse as 2-lists; bars hang off a
    central label gutter — left = values[0] (accent2), right = values[1]
    (accent1), shared symmetric scale, value labels at the outer ends. Rows
    sort by |left|+|right| descending unless '- Sort: off'."""
    rows = [(lbl, v) for lbl, v in p.get("data", [])
            if isinstance(v, list) and len(v) == 2]
    if len(rows) < 2:
        warn("tornado needs **Series:** Low, High and 2+ "
             "'- Driver: low, high' **Data:** rows")
        return B.build_bullet_slide(prs, p, pal, ctx)
    if p.get("sort", "").strip().lower() not in ("off", "no", "false"):
        rows.sort(key=lambda r: abs(r[1][0]) + abs(r[1][1]), reverse=True)
    slide = B._blank_slide(prs, pal, pal["bg"])
    B._heading(slide, p, pal)
    body_start = len(slide.shapes)

    top, bottom = 2.0, 6.3
    cx, gutter = 6.65, 1.25   # spine center + half-width of the label gutter
    val_w = 0.72              # outer value-label zone per side
    half_w = cx - gutter - 0.7 - val_w
    vmax = max(max(abs(v[0]), abs(v[1])) for _, v in rows) or 1
    row_h = min(0.78, (bottom - top) / len(rows))
    bar_h = max(0.04, min(0.42, row_h - 0.16))

    names = [s.strip() for s in p.get("series", "").split(",") if s.strip()]
    if len(names) >= 2:  # side headers above the two bar areas
        B.add_tb(slide, names[0], 0.7, top - 0.42, cx - gutter - 0.75, 0.3,
                 size=11, bold=True, color=pal["text_muted"],
                 align=PP_ALIGN.RIGHT, font=pal["font_label"])
        B.add_tb(slide, names[1], cx + gutter + 0.05, top - 0.42,
                 12.6 - cx - gutter, 0.3, size=11, bold=True,
                 color=pal["text_muted"], font=pal["font_label"])
    chart_h = row_h * len(rows)
    for gx in (cx - gutter, cx + gutter):  # spine edges the bars hang off
        B.add_rect(slide, gx - 0.009, top, 0.018, chart_h, pal["text_muted"])
    for i, (label, (lo, hi)) in enumerate(rows):
        y = top + i * row_h + (row_h - bar_h) / 2
        wl = max(half_w * abs(lo) / vmax, 0.02)
        wr = max(half_w * abs(hi) / vmax, 0.02)
        B.add_rect(slide, cx - gutter - wl, y, wl, bar_h, pal["accent2"])
        B.add_rect(slide, cx + gutter, y, wr, bar_h, pal["accent1"])
        B.add_tb(slide, label, cx - gutter + 0.05, y + bar_h / 2 - 0.16,
                 2 * gutter - 0.1, 0.35, size=12, bold=True, color=pal["text"],
                 align=PP_ALIGN.CENTER, font=pal["font_body"])
        B.add_tb(slide, _fmt_num(lo), cx - gutter - wl - val_w - 0.05,
                 y + bar_h / 2 - 0.16, val_w, 0.32, size=12,
                 color=pal["text_muted"], align=PP_ALIGN.RIGHT,
                 font=pal["font_label"])
        B.add_tb(slide, _fmt_num(hi, signed=hi > 0), cx + gutter + wr + 0.05,
                 y + bar_h / 2 - 0.16, val_w, 0.32, size=12,
                 color=pal["text_muted"], font=pal["font_label"])
    shape_groups.group_shapes(slide, list(slide.shapes)[body_start:])
    return slide



# ── football field (valuation ranges) ────────────────────────────────────────
def _parse_marker(spec):
    """'Current price, 47' -> ('Current price', 47.0) or None."""
    label, sep, val = (spec or "").rpartition(",")
    m = _NUM_RX.search(val)
    if not sep or not m or not label.strip():
        return None
    return label.strip().strip('"'), float(m.group().replace(",", ""))


def build_football_field_slide(prs, p, pal, ctx):
    """Valuation football field. Requires **Series:** Low, High so
    '- Method: low, high' rows parse as 2-lists; each row is a floating
    rounded bar low->high on a shared value axis with nice-interval
    gridlines. Optional '- Marker: label, value' draws a dashed vertical
    reference line (skipped with a warning outside the data range)."""
    rows = [(lbl, v[0], v[1]) for lbl, v in p.get("data", [])
            if isinstance(v, list) and len(v) == 2 and v[0] < v[1]]
    if len(rows) < 2:
        warn("football-field needs **Series:** Low, High and 2+ "
             "'- Method: low, high' **Data:** rows (low < high)")
        return B.build_bullet_slide(prs, p, pal, ctx)
    slide = B._blank_slide(prs, pal, pal["bg"])
    B._heading(slide, p, pal)
    body_start = len(slide.shapes)

    import math
    from charts import _nice_ceil
    gmin = min(lo for _, lo, _ in rows)
    gmax = max(hi for _, _, hi in rows)
    step = _nice_ceil((gmax - gmin) / 5 or 1)
    a0 = math.floor(gmin / step) * step
    a1 = math.ceil(gmax / step) * step
    if a1 <= a0:
        a1 = a0 + step

    # axis labels must stay above the footer zone (top <= 6.0in) or
    # pptx_lint's page-number heuristics flag the digit-only tick labels
    L, R, top, bottom = 3.5, 12.5, 2.0, 5.85

    def vx(v):
        return L + (v - a0) / (a1 - a0) * (R - L)

    for k in range(int(round((a1 - a0) / step)) + 1):  # gridlines + axis
        v = a0 + k * step
        B.add_rect(slide, vx(v) - 0.0075, top, 0.015, bottom - top,
                   pal["surface"])
        B.add_tb(slide, _fmt_plain(v), vx(v) - 0.6, bottom + 0.08, 1.2, 0.3,
                 size=11, color=pal["text_muted"], align=PP_ALIGN.CENTER,
                 font=pal["font_label"])

    from palettes import hex_rgb
    row_h = min(0.9, (bottom - top) / len(rows))
    bar_h = max(0.04, min(0.46, row_h - 0.22))
    for i, (label, lo, hi) in enumerate(rows):
        y = top + i * row_h + (row_h - bar_h) / 2
        x0, x1 = vx(lo), vx(hi)
        sl, st, sw, sh = B._sc(x0, y, x1 - x0, bar_h)
        bar = slide.shapes.add_shape(5, Inches(sl), Inches(st),
                                     Inches(sw), Inches(sh))  # 5 = ROUNDED_RECT
        bar.fill.solid()
        bar.fill.fore_color.rgb = hex_rgb(pal["accent1"])
        bar.line.fill.background()
        B.add_tb(slide, label, 0.7, y + bar_h / 2 - 0.17, 2.3, 0.38, size=13,
                 bold=True, color=pal["text"], align=PP_ALIGN.RIGHT,
                 font=pal["font_body"])
        B.add_tb(slide, _fmt_plain(lo), x0 - 0.78, y + bar_h / 2 - 0.16, 0.7,
                 0.32, size=12, color=pal["text_muted"], align=PP_ALIGN.RIGHT,
                 font=pal["font_label"])
        B.add_tb(slide, _fmt_plain(hi), x1 + 0.05, y + bar_h / 2 - 0.16, 0.7,
                 0.32, size=12, color=pal["text_muted"],
                 font=pal["font_label"])
    _football_marker(slide, p, pal, top, bottom, vx, gmin, gmax)
    shape_groups.group_shapes(slide, list(slide.shapes)[body_start:])
    return slide


def _football_marker(slide, p, pal, top, bottom, vx, gmin, gmax):
    """Dashed vertical reference line + label for '- Marker: label, value'."""
    spec = p.get("marker")
    if not spec:
        return
    parsed = _parse_marker(spec)
    if not parsed:
        warn(f"Marker needs 'label, value': {spec!r} — skipped")
        return
    label, value = parsed
    if not gmin <= value <= gmax:
        warn(f"Marker {label!r} value {value:g} outside range "
             f"[{gmin:g}, {gmax:g}] — skipped")
        return
    from pptx.enum.dml import MSO_LINE_DASH_STYLE
    from palettes import hex_rgb
    x = vx(value)
    sl, st, _, h = B._sc(x, top - 0.12, 0, bottom - top + 0.12)
    conn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,
                                      Inches(sl), Inches(st),
                                      Inches(sl), Inches(st + h))
    conn.line.color.rgb = hex_rgb(pal["accent3"])
    conn.line.width = Pt(1.75)
    conn.line.dash_style = MSO_LINE_DASH_STYLE.DASH
    # clamp 2.4in-wide label so it stays within canvas (right edge <= 12.63)
    label_left = max(0.7, min(x - 1.2, 12.63 - 2.4))
    B.add_tb(slide, f"{label} · {_fmt_plain(value)}", label_left, top - 0.46,
             2.4, 0.3, size=11, bold=True, color=pal["text_muted"],
             align=PP_ALIGN.CENTER, font=pal["font_label"])



# ── gantt / swimlane roadmap ─────────────────────────────────────────────────
# build_gantt_slide moved to builders_gantt.py (Cycle 7 §8 — label columns,
# spans, shades, milestone lines, hierarchy, period expansion); re-exported
# here so existing imports and the LAYOUTS registration keep working.
from builders_gantt import build_gantt_slide  # noqa: E402,F401


LAYOUTS = {
    "waterfall": build_waterfall_slide,
    "gantt": build_gantt_slide,
    "funnel": build_funnel_slide,
    "mekko": build_mekko_slide,
    "bar-mekko": build_bar_mekko_slide,
    "tornado": build_tornado_slide,
    "football-field": build_football_field_slide,
}
