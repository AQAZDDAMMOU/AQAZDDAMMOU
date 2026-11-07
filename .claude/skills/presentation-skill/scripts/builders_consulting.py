"""Consulting layout builders: 2x2 matrix, harvey scorecard, process flow,
big number, chart callout, dashboard, quote+evidence, next-steps, agenda,
SCQA executive summary, heatmap-table, driver-tree, stakeholder-map, raci.
Shape-built chart exhibits (waterfall, funnel, mekko, bar-mekko, gantt,
tornado, football-field) live in builders_exhibits.py and are re-exported here.

Registered into builders.LAYOUT_MAP at import time (see bottom of builders.py).
All geometry is in design inches on the 13.33 x 7.5 canvas; builders' wrappers
rescale to the actual slide size.
"""
import re

from pptx.enum.shapes import MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

import builders as B
import shape_groups
from helpers import label_color_on as _label_color_on
from helpers import rel_lum as _rel_lum
from helpers import set_shape_text
from helpers import warn




# ── 2x2 matrix ───────────────────────────────────────────────────────────────
def build_matrix_slide(prs, p, pal, ctx, *, group=True, body_out=None):
    """`group=False` leaves the exhibit ungrouped and `body_out` (a list)
    receives the index the body starts at — how `build_stakeholder_map_slide`
    borrows this engine and still emits exactly one exhibit group covering
    both the matrix and its own move arrows."""
    slide = B._blank_slide(prs, pal, pal["bg"])
    B._heading(slide, p, pal)
    body_start = len(slide.shapes)
    if body_out is not None:
        body_out.append(body_start)
    L, T, W, H = 2.0, 1.85, 9.6, 4.6  # plot area
    # quadrant tints + frame
    B.add_rect(slide, L, T, W, H, pal["surface"])
    B.add_rect(slide, L, T + H / 2, W, 0.02, pal["text_muted"])
    B.add_rect(slide, L + W / 2, T, 0.02, H, pal["text_muted"])
    # axes labels
    B.add_tb(slide, p.get("x_axis", "X →"), L, T + H + 0.12, W, 0.4, size=13,
             bold=True, color=pal["text_muted"], align=PP_ALIGN.CENTER,
             font=pal["font_body"])
    y_label = p.get("y_axis", "Y ↑").replace("→", "↑")  # vertical axis points up
    B.add_tb(slide, y_label, 0.4, T + H / 2 - 0.2, 1.5, 0.8,
             size=13, bold=True, color=pal["text_muted"], font=pal["font_body"])
    # quadrant labels: Q1 TL, Q2 TR, Q3 BL, Q4 BR
    quads = [("q1", L, T), ("q2", L + W / 2, T),
             ("q3", L, T + H / 2), ("q4", L + W / 2, T + H / 2)]
    for key, qx, qy in quads:
        if p.get(key):
            B.add_tb(slide, p[key], qx + 0.15, qy + 0.1, W / 2 - 0.3, 0.4,
                     size=12, bold=True, color=pal["text_muted"],
                     font=pal["font_body"])
    items = p.get("matrix_items", [])[:12]
    sizes = []
    for it in items:
        try:
            s = float(it["size"])
            sizes.append(s if s > 0 else None)
        except (KeyError, ValueError):
            sizes.append(None)
    smax = max((s for s in sizes if s), default=0)
    from helpers import set_fill_alpha
    for item, s in zip(items, sizes):
        try:
            fx, fy = float(item.get("x", 0.5)), float(item.get("y", 0.5))
        except ValueError:
            continue
        cx = L + fx * W
        cy = T + (1 - fy) * H
        d = 0.18 if not (s and smax) else 0.24 + 0.66 * (s / smax) ** 0.5
        cx = max(L + d / 2, min(L + W - d / 2, cx))
        cy = max(T + d / 2, min(T + H - d / 2, cy))
        dot = B.add_circle(slide, cx - d / 2, cy - d / 2, d, pal["accent1"])
        if s and smax:
            set_fill_alpha(dot, 80)  # bubbles overlap; keep grid visible
        off = d / 2 + 0.06
        if cx + off + 2.2 > L + W:  # label would cross the right border
            B.add_tb(slide, item.get("name", ""), cx - off - 2.2, cy - 0.16,
                     2.2, 0.35, size=11, color=pal["text"],
                     align=PP_ALIGN.RIGHT, font=pal["font_body"])
        else:
            B.add_tb(slide, item.get("name", ""), cx + off, cy - 0.16, 2.2,
                     0.35, size=11, color=pal["text"], font=pal["font_body"])
    if group:
        shape_groups.group_shapes(slide, list(slide.shapes)[body_start:])
    return slide


# ── harvey-ball scorecard ────────────────────────────────────────────────────
def _harvey(slide, pal, cx, cy, d, quarters):
    """0-4 quarters filled. Outline circle + pie wedge (or full disc)."""
    ring = B.add_circle(slide, cx, cy, d, pal["bg_deep"] if pal["dark"] else "FFFFFF",
                        line_hex=pal["accent1"], line_pt=1.2)
    if quarters >= 4:
        B.add_circle(slide, cx + 0.03, cy + 0.03, d - 0.06, pal["accent1"])
    elif quarters > 0:
        sl, st, sw, _ = B._sc(cx + 0.03, cy + 0.03, d - 0.06, d - 0.06)
        pie = slide.shapes.add_shape(142, Inches(sl), Inches(st), Inches(sw), Inches(sw))
        pie.fill.solid()
        from palettes import hex_rgb
        pie.fill.fore_color.rgb = hex_rgb(pal["accent1"])
        pie.line.fill.background()
        # OOXML pie angles are 60000ths of a degree; python-pptx adjustments
        # write value*100000 — so multiply degrees by 0.6
        # Keep angles in 0..360 — negative adj values render inconsistently
        # in LibreOffice (start at 12 o'clock = 270 deg, sweep clockwise).
        pie.adjustments[0] = 270.0 * 0.6
        pie.adjustments[1] = ((270.0 + 90.0 * quarters) % 360.0) * 0.6
    return ring


def build_harvey_slide(prs, p, pal, ctx):
    rows = p.get("table_rows", [])
    if len(rows) < 2:
        warn("harvey-scorecard needs a markdown table; rendering as table")
        return B.build_table_slide(prs, p, pal, ctx)
    slide = B._blank_slide(prs, pal, pal["bg"])
    B._heading(slide, p, pal)
    body_start = len(slide.shapes)
    n_rows, n_cols = len(rows), max(len(r) for r in rows)
    first_col_w = 3.6
    col_w = (11.9 - first_col_w) / max(n_cols - 1, 1)
    top = 1.9
    # Table + legend must clear the bottom band: kicker sits at 6.18,
    # the source footnote at 6.80, the footer at 7.08.
    bottom_limit = 7.0
    if p.get("source"):
        bottom_limit = 6.75
    if p.get("kicker"):
        bottom_limit = 6.10
    legend_h = 0.4  # 0.1 gap + 0.3 legend text
    row_h = min(0.62, (bottom_limit - top - legend_h) / n_rows - 0.08)
    for r, row in enumerate(rows):
        y = top + r * (row_h + 0.08)
        if r > 0:
            B.add_rect(slide, 0.7, y, 11.9, row_h,
                       pal["surface"] if r % 2 else pal["bg_deep"])
        for c in range(n_cols):
            cell = row[c] if c < len(row) else ""
            x = 0.7 + (first_col_w + (c - 1) * col_w if c else 0)
            w = col_w if c else first_col_w
            if r == 0:  # header
                B.add_tb(slide, cell, x, y + 0.05, w, row_h, size=14, bold=True,
                         color=pal["accent1"],
                         align=PP_ALIGN.CENTER if c else PP_ALIGN.LEFT,
                         font=pal["font_body"])
            elif c == 0:
                B.add_tb(slide, cell, x + 0.15, y + row_h / 2 - 0.16, w - 0.2, 0.35,
                         size=13, color=pal["text"], font=pal["font_body"])
            elif cell.strip().isdigit():
                d = min(row_h - 0.18, 0.4)
                _harvey(slide, pal, x + w / 2 - d / 2, y + row_h / 2 - d / 2, d,
                        max(0, min(4, int(cell))))
            else:
                B.add_tb(slide, cell, x, y + row_h / 2 - 0.16, w, 0.35, size=12,
                         color=pal["text_muted"], align=PP_ALIGN.CENTER,
                         font=pal["font_body"])
    B.add_tb(slide, "Empty = not met · quarter fills = partial · solid = fully met",
             0.7, top + n_rows * (row_h + 0.08) + 0.1, 11.9, 0.3, size=11,
             color=pal["text_muted"], font=pal["font_label"])
    shape_groups.group_shapes(slide, list(slide.shapes)[body_start:])
    return slide


# ── process flow ─────────────────────────────────────────────────────────────
def build_process_flow_slide(prs, p, pal, ctx):
    steps = p.get("steps", [])[:6]
    if not steps:
        steps = [{"title": B.split_icon(b)[1]}
                 for b in p.get("bullets", [])[:6]]
    slide = B._blank_slide(prs, pal, pal["bg"])
    B._heading(slide, p, pal)
    body_start = len(slide.shapes)
    n = max(len(steps), 1)
    gap = 0.12
    w = (11.9 - gap * (n - 1)) / n
    # Center the chevron block (plus desc zone when present) in the content
    # area instead of pinning it near the title.
    has_desc = any(s.get("desc") for s in steps)
    block = 1.15 + (0.2 + 2.8 if has_desc else 0)
    zone_top, zone_bot = 1.9, B.content_bottom(p)
    chev_y = zone_top + max(0.0, (zone_bot - zone_top - block) / 2)
    for i, step in enumerate(steps):
        x = 0.7 + i * (w + gap)
        sl, st, ww, hh = B._sc(x, chev_y, w, 1.15)
        chev = slide.shapes.add_shape(52, Inches(sl), Inches(st),
                                      Inches(ww), Inches(hh))  # 52 = CHEVRON
        chev.fill.solid()
        from palettes import hex_rgb
        chev.fill.fore_color.rgb = hex_rgb(pal["accent1"] if i == 0 else pal["surface"])
        chev.line.fill.background()
        B.add_tb(slide, step.get("title", ""), x + 0.45, chev_y + 0.32,
                 w - 0.55, 0.6, size=14, bold=True,
                 color=pal["bg_deep"] if i == 0 else pal["text"],
                 font=pal["font_body"])
        if step.get("desc"):
            B.add_tb(slide, step["desc"], x + 0.05, chev_y + 1.35, w - 0.1,
                     min(2.8, zone_bot - chev_y - 1.35), size=12,
                     color=pal["text_muted"], font=pal["font_body"],
                     accent=pal["accent1"])
    shape_groups.group_shapes(slide, list(slide.shapes)[body_start:])
    return slide


# ── big number ───────────────────────────────────────────────────────────────
def build_big_number_slide(prs, p, pal, ctx):
    slide = B._blank_slide(prs, pal, pal["bg_deep"], accent_bar=False, gradient=True)
    B.add_tb(slide, p.get("value", ""), 0.7, 2.0, 11.9, 2.2, size=110, bold=True,
             color=pal["accent1"], align=PP_ALIGN.CENTER, font=pal["font_title"])
    B.add_tb(slide, p.get("label", p.get("heading", "")), 1.5, 4.5, 10.33, 0.9,
             size=24, bold=True, color=pal["text"], align=PP_ALIGN.CENTER,
             font=pal["font_body"])
    if p.get("context"):
        B.add_tb(slide, p["context"], 1.5, 5.45, 10.33, 0.7, size=16,
                 color=pal["text_muted"], align=PP_ALIGN.CENTER,
                 font=pal["font_label"])
    return slide


# ── chart + insight callout ──────────────────────────────────────────────────
def build_chart_callout_slide(prs, p, pal, ctx):
    slide = B._blank_slide(prs, pal, pal["bg"])
    B._heading(slide, p, pal)
    B._place_visual(slide, prs, p, pal, ctx, 0.7, 1.85, 8.4, 4.9)
    box = B.add_rect(slide, 9.4, 2.5, 3.45, 2.8, pal["surface"],
                     line_hex=pal["accent1"], line_pt=1.5)
    from helpers import add_soft_shadow
    add_soft_shadow(box)
    sl, st, sw, sh = B._sc(9.12, 3.6, 0.32, 0.26)
    arrow = slide.shapes.add_shape(34, Inches(sl), Inches(st), Inches(sw), Inches(sh))
    arrow.fill.solid()
    from palettes import hex_rgb
    arrow.fill.fore_color.rgb = hex_rgb(pal["accent1"])
    arrow.line.fill.background()
    B.add_tb(slide, "KEY INSIGHT", 9.65, 2.7, 3.0, 0.35, size=14, bold=True,
             color=pal["accent1"], font=pal["font_label"])
    B.add_tb(slide, p.get("callout", ""), 9.65, 3.1, 3.0, 2.05, size=14,
             color=pal["text"], font=pal["font_body"], accent=pal["accent1"])
    return slide


# ── KPI dashboard ────────────────────────────────────────────────────────────
def build_dashboard_slide(prs, p, pal, ctx):
    tiles = p.get("tiles", [])[:6]
    slide = B._blank_slide(prs, pal, pal["bg"])
    B._heading(slide, p, pal)
    body_start = len(slide.shapes)
    cols = 3 if len(tiles) > 4 else 2
    rows = -(-len(tiles) // cols)
    gap = 0.3
    tw = (11.9 - gap * (cols - 1)) / cols
    th = min(2.15, (4.7 - gap * (rows - 1)) / max(rows, 1))
    from helpers import add_soft_shadow
    for i, tile in enumerate(tiles):
        r, c = divmod(i, cols)
        x = 0.7 + c * (tw + gap)
        y = 1.95 + r * (th + gap)
        card = B.add_rect(slide, x, y, tw, th, pal["surface"])
        add_soft_shadow(card)
        B.add_tb(slide, tile.get("value", ""), x + 0.25, y + 0.18, tw - 0.5, 0.85,
                 size=32, bold=True, color=pal["accent1"], font=pal["font_title"])
        B.add_tb(slide, tile.get("label", ""), x + 0.25, y + 0.98, tw - 0.5,
                 0.45, size=13, color=pal["text"], font=pal["font_body"])
        if tile.get("delta"):
            up = tile.get("trend", "up").lower() != "down"
            # 14pt bold = WCAG large text (3:1), so these hold on white tiles
            B.add_tb(slide, f"{'▲' if up else '▼'} {tile['delta']}",
                     x + 0.25, y + 1.42, tw - 0.5, 0.4, size=14, bold=True,
                     color=pal["accent1"] if up else NEG_HEX,
                     font=pal["font_body"])
    shape_groups.group_shapes(slide, list(slide.shapes)[body_start:])
    return slide


# ── quote + evidence ─────────────────────────────────────────────────────────
def build_quote_evidence_slide(prs, p, pal, ctx):
    slide = B._blank_slide(prs, pal, pal["bg"])
    B.add_tb(slide, "“", 0.6, 1.0, 1.4, 1.6, size=110, bold=True,
             color=pal["accent1"], font=pal["font_title"])
    quote_tb = B.add_tb(slide, p.get("quote", ""), 1.1, 2.3, 6.9, 3.2, size=22,
                        color=pal["text"], font=pal["font_title"],
                        accent=pal["accent1"])
    from pptx.enum.text import MSO_ANCHOR
    quote_tb.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    if p.get("attribution"):
        B.add_tb(slide, f"— {p['attribution']}", 1.1, 5.6, 6.9, 0.6, size=14,
                 bold=True, color=pal["text_muted"], font=pal["font_body"])
    stats = p.get("stats", [])[:1]
    if stats:
        from helpers import add_soft_shadow
        card = B.add_rect(slide, 8.7, 2.2, 4.0, 3.2, pal["surface"],
                          line_hex=pal["accent1"], line_pt=1.5)
        add_soft_shadow(card)
        B.add_tb(slide, stats[0].get("value", ""), 8.9, 2.6, 3.6, 1.3, size=48,
                 bold=True, color=pal["accent1"], align=PP_ALIGN.CENTER,
                 font=pal["font_title"])
        B.add_tb(slide, stats[0].get("label", ""), 8.9, 4.0, 3.6, 1.1, size=14,
                 color=pal["text"], align=PP_ALIGN.CENTER, font=pal["font_body"])
    return slide



# ── next steps (conclusions that stay up during Q&A) ─────────────────────────
def build_next_steps_slide(prs, p, pal, ctx):
    steps = p.get("steps", [])[:6]
    if not steps:  # plain bullets validate too — render them as step cards
        steps = [{"action": B.split_icon(b)[1]}
                 for b in p.get("bullets", [])[:6]]
    slide = B._blank_slide(prs, pal, pal["bg"])
    B._heading(slide, p, pal)
    from helpers import add_soft_shadow
    row_h = min(0.85, 4.8 / max(len(steps), 1))
    body_start = len(slide.shapes)
    for i, s in enumerate(steps):
        y = 1.95 + i * (row_h + 0.15)
        card = B.add_rect(slide, 0.7, y, 11.9, row_h, pal["surface"])
        add_soft_shadow(card)
        # the number is the circle's own text (a text box parked on the
        # circle would be a text-over-shape editability hit)
        badge = B.add_circle(slide, 0.95, y + row_h / 2 - 0.19, 0.38,
                             pal["accent1"])
        set_shape_text(badge, str(i + 1), size=B._fs(15), bold=True,
                       color=pal["bg_deep"], font=pal["font_body"])
        B.add_tb(slide, s.get("action", s.get("title", "")), 1.55,
                 y + row_h / 2 - 0.19, 6.6, 0.42, size=14, bold=True,
                 color=pal["text"], font=pal["font_body"])
        B.add_tb(slide, s.get("owner", ""), 8.3, y + row_h / 2 - 0.19, 2.1, 0.42,
                 size=14, bold=True, color=pal["accent1"], font=pal["font_body"])
        B.add_tb(slide, s.get("when", ""), 10.5, y + row_h / 2 - 0.19, 2.0, 0.42,
                 size=13, color=pal["text_muted"], align=PP_ALIGN.RIGHT,
                 font=pal["font_body"])
    shape_groups.group_shapes(slide, list(slide.shapes)[body_start:])
    return slide


# ── agenda with current-section highlight ────────────────────────────────────
# Moved to builders_agenda.py (Cycle 7: time/page columns, sub-chapters);
# re-exported here so `builders_consulting.build_agenda_slide` keeps resolving.
from builders_agenda import build_agenda_slide  # noqa: E402  (re-export)


# ── SCQA executive summary (Situation | Findings | Recommendation) ───────────
def build_exec_scqa_slide(prs, p, pal, ctx):
    slide = B._blank_slide(prs, pal, pal["bg"])
    B._heading(slide, p, pal)
    body_start = len(slide.shapes)
    from helpers import add_soft_shadow
    cols = [
        ("SITUATION", [p.get("situation", "")], False),
        ("FINDINGS", p.get("bullets", [])[:4], False),
        ("RECOMMENDATION", [p.get("recommendation", "")], True),
    ]
    gap = 0.3
    col_w = (11.9 - 2 * gap) / 3
    card_h = B.content_bottom(p) - 1.9
    for i, (header, items, emphasize) in enumerate(cols):
        x = 0.7 + i * (col_w + gap)
        card = B.add_rect(slide, x, 1.9, col_w, card_h, pal["surface"],
                          line_hex=pal["accent1"] if emphasize else None,
                          line_pt=1.5 if emphasize else 0)
        add_soft_shadow(card)
        B.add_tb(slide, header, x + 0.25, 2.1, col_w - 0.5, 0.4, size=14,
                 bold=True, color=pal["accent1"], font=pal["font_label"])
        single = len([t for t in items if t]) == 1
        for j, item in enumerate([t for t in items if t]):
            # All bodies top-align to the same first line (y=2.6) so the three
            # columns share a baseline; singles just get wrap room below.
            B.add_tb(slide, item, x + 0.25, 2.6 + j * 1.0, col_w - 0.5,
                     0.96 if not single else card_h - 0.9,
                     size=13, bold=emphasize, color=pal["text"],
                     font=pal["font_body"], accent=pal["accent1"])
    shape_groups.group_shapes(slide, list(slide.shapes)[body_start:])
    return slide


LAYOUTS = {
    "matrix-2x2": build_matrix_slide,
    "harvey-scorecard": build_harvey_slide,
    "process-flow": build_process_flow_slide,
    "big-number": build_big_number_slide,
    "chart-callout": build_chart_callout_slide,
    "dashboard": build_dashboard_slide,
    "quote-evidence": build_quote_evidence_slide,
    "next-steps": build_next_steps_slide,
    "agenda": build_agenda_slide,
    "exec-summary-scqa": build_exec_scqa_slide,
}





# ── heatmap table ────────────────────────────────────────────────────────────
_CELL_NUM_RX = re.compile(r"-?(?:\d{1,3}(?:,\d{3})*|\d+)(?:\.\d+)?")


def _cell_value(cell):
    """Numeric value of a table cell ('$1,200', '42%', '-3.5') or None."""
    stripped = str(cell).strip().lstrip("$€£").rstrip("%").strip()
    m = _CELL_NUM_RX.fullmatch(stripped)
    return float(m.group().replace(",", "")) if m else None


def _lerp_hex(a, b, t):
    """Linear blend between two hex colors, t in 0..1."""
    return "".join(
        f"{round(int(a[i:i + 2], 16) + (int(b[i:i + 2], 16) - int(a[i:i + 2], 16)) * t):02X}"
        for i in (0, 2, 4))


def _bw_on(fill_hex):
    """Pure black/white text by fill luminance. The 0.179 threshold is the
    equal-contrast point, so the winner always clears 4.5:1 on any fill
    (palette text colors can dip to ~4.3:1 on mid-tone heat cells)."""
    return "000000" if _rel_lum(fill_hex) > 0.179 else "FFFFFF"


def _rag_fill(v, svals, pal):
    """Tercile thresholds over the column's sorted values -> red/amber/green.

    Note: with exactly 2 rows, t1 == svals[0] and t2 == svals[-1], so the
    function yields rag_bad for the lower value and rag_good for the upper
    (rag_mid is never returned in the 2-row case).
    """
    if svals[0] == svals[-1]:
        return pal["rag_mid"]
    n = len(svals)
    t1, t2 = svals[(n - 1) // 3], svals[(2 * (n - 1)) // 3]
    if v <= t1:
        return pal["rag_bad"]
    return pal["rag_mid"] if v <= t2 else pal["rag_good"]


def build_heatmap_slide(prs, p, pal, ctx):
    """Markdown table whose numeric body cells get heat fills, normalized
    per column: bg at the column min -> accent1 at the max. '- Scale: rag'
    switches to terciled red/amber/green chips. Non-numeric body cells get
    a plain surface fill; cell text picks its color by fill luminance."""
    rows = p.get("table_rows", [])
    if len(rows) < 2:
        warn("heatmap-table needs a markdown table; rendering as table")
        return B.build_table_slide(prs, p, pal, ctx)
    slide = B._blank_slide(prs, pal, pal["bg"])
    B._heading(slide, p, pal)
    body_start = len(slide.shapes)
    rag = p.get("scale", "").strip().lower() == "rag"

    n_cols = max(len(r) for r in rows)
    col_vals = []  # per column: sorted numeric body values
    for c in range(n_cols):
        nums = [_cell_value(r[c]) for r in rows[1:] if c < len(r)]
        col_vals.append(sorted(v for v in nums if v is not None))

    n_body_rows = len(rows) - 1  # exclude header
    if n_body_rows > 12:
        warn(f"heatmap-table with {n_body_rows} rows — consider splitting")

    first_col_w = 3.2
    col_w = (11.9 - first_col_w) / max(n_cols - 1, 1)
    # rows + 0.06 gaps must clear the footer zone (1.9 + total <= 6.6)
    row_h = min(0.72, 4.7 / len(rows) - 0.06)
    top = 1.9
    for r, row in enumerate(rows):
        y = top + r * (row_h + 0.06)
        for c in range(n_cols):
            cell = row[c] if c < len(row) else ""
            x = 0.7 + (first_col_w + (c - 1) * col_w if c else 0)
            w = col_w if c else first_col_w
            if r == 0:  # header
                B.add_tb(slide, cell, x + (0.12 if c == 0 else 0),
                         y + row_h / 2 - 0.16, w, 0.35, size=13, bold=True,
                         color=pal["accent1"],
                         align=PP_ALIGN.CENTER if c else PP_ALIGN.LEFT,
                         font=pal["font_body"])
                continue
            v = _cell_value(cell)
            if v is None or not col_vals[c]:
                fill, text_color = pal["surface"], pal["text"]
            elif rag:
                fill = _rag_fill(v, col_vals[c], pal)
                text_color = _bw_on(fill)
            else:
                lo, hi = col_vals[c][0], col_vals[c][-1]
                t = (v - lo) / (hi - lo) if hi > lo else 0.5
                fill = _lerp_hex(pal["bg"], pal["accent1"], t)
                text_color = _bw_on(fill)
            B.add_rect(slide, x + 0.03, y, w - 0.06, row_h, fill)
            # when rows are very dense (row_h < 0.35) shrink the text box
            # height to match and drop font size to 10 so text doesn't overflow
            cell_tb_h = min(0.35, row_h)
            cell_font = 10 if row_h < 0.35 else 12
            B.add_tb(slide, cell, x + (0.12 if c == 0 else 0),
                     y + row_h / 2 - cell_tb_h / 2, w, cell_tb_h,
                     size=cell_font, color=text_color,
                     align=PP_ALIGN.CENTER if c else PP_ALIGN.LEFT,
                     font=pal["font_body"])
    shape_groups.group_shapes(slide, list(slide.shapes)[body_start:])
    return slide


LAYOUTS["heatmap-table"] = build_heatmap_slide






# ── driver tree (value decomposition) ────────────────────────────────────────
def _tree_leaf_order(nid, children, seen):
    """DFS leaf ids under nid in outline order (cycle-safe)."""
    if nid in seen:
        return []
    seen.add(nid)
    kids = children.get(nid, [])
    if not kids:
        return [nid]
    out = []
    for k in kids:
        out.extend(_tree_leaf_order(k, children, seen))
    return out


def build_driver_tree_slide(prs, p, pal, ctx):
    """Left-to-right value decomposition: root box at left, children in
    depth columns to the right, elbow connectors parent -> child. Boxes show
    Label + Value + optional Delta (green-ish '+', red-ish '-')."""
    nodes = [n for n in p.get("nodes", []) if n.get("id", "").strip()][:12]
    if not nodes:
        warn("driver-tree needs '- Node:' rows; rendering bullet-list")
        return B.build_bullet_slide(prs, p, pal, ctx)
    slide = B._blank_slide(prs, pal, pal["bg"])
    B._heading(slide, p, pal)
    body_start = len(slide.shapes)

    by_id = {n["id"]: n for n in nodes}
    children = {}
    for n in nodes:
        par = n.get("parent", "").strip()
        if par and par in by_id:
            children.setdefault(par, []).append(n["id"])
    roots = [n["id"] for n in nodes if not n.get("parent", "").strip()]
    root = roots[0] if roots else nodes[0]["id"]

    depths, queue = {root: 1}, [root]  # BFS from root (validator caps at 3)
    while queue:
        nid = queue.pop(0)
        for ch in children.get(nid, []):
            if ch not in depths:
                depths[ch] = depths[nid] + 1
                queue.append(ch)
    skipped = [i for i in by_id if i not in depths]
    if skipped:
        warn("driver-tree: nodes unreachable from root skipped: "
             + ", ".join(skipped))
    max_d = max(depths.values())

    top, bottom = 1.95, 6.45
    gap = 0.55  # connector corridor between columns
    col_w = min(3.4, (11.93 - gap * (max_d - 1)) / max_d)
    leaves = _tree_leaf_order(root, children, set())
    row_h = (bottom - top) / max(len(leaves), 1)
    box_h = max(0.38, min(0.95, row_h - 0.10))
    ys = {nid: top + r * row_h + row_h / 2 for r, nid in enumerate(leaves)}

    def center_y(nid):  # parents center on their children
        if nid not in ys:
            ys[nid] = sum(center_y(k) for k in children[nid]) / len(children[nid])
        return ys[nid]
    center_y(root)

    for nid, d in depths.items():
        _tree_box(slide, pal, by_id[nid],
                  0.7 + (d - 1) * (col_w + gap), ys[nid] - box_h / 2,
                  col_w, box_h)
    _tree_connectors(slide, pal, children, depths, ys, col_w, gap)
    shape_groups.group_shapes(slide, list(slide.shapes)[body_start:])
    return slide


def _tree_box(slide, pal, node, x, y, col_w, box_h):
    """Surface card with accent edge stripe: Label bold + Value + Delta."""
    B.add_rect(slide, x, y, col_w, box_h, pal["surface"])
    B.add_rect(slide, x, y, 0.06, box_h, pal["accent1"])
    delta = node.get("delta", "").strip()
    d_color = (pal["rag_good"] if delta.startswith("+")
               else pal["rag_bad"] if delta.startswith("-")
               else pal["text_muted"])
    if box_h >= 0.62:  # two lines: label / value + delta
        B.add_tb(slide, node.get("label", node["id"]), x + 0.18, y + 0.07,
                 col_w - 0.30, 0.32, size=13, bold=True, color=pal["text"],
                 font=pal["font_body"])
        B.add_tb(slide, node.get("value", ""), x + 0.18, y + box_h - 0.36,
                 col_w - 1.2, 0.3, size=12, color=pal["text_muted"],
                 font=pal["font_body"])
        if delta:
            B.add_tb(slide, delta, x + col_w - 1.0, y + box_h - 0.36, 0.88,
                     0.3, size=12, bold=True, color=d_color,
                     align=PP_ALIGN.RIGHT, font=pal["font_body"])
    else:  # compact: single centered row, value + delta right-aligned
        ty = y + box_h / 2 - 0.15
        B.add_tb(slide, node.get("label", node["id"]), x + 0.16, ty,
                 col_w - 1.9, 0.3, size=11, bold=True, color=pal["text"],
                 font=pal["font_body"])
        # label box ends at (x + col_w - 1.74in); value box starts at
        # (x + col_w - 1.78in) — intentional 0.04in overlap so text regions
        # butt against each other without a visible gap at compact density.
        B.add_tb(slide, node.get("value", ""), x + col_w - 1.78, ty, 0.95,
                 0.3, size=11, color=pal["text_muted"], align=PP_ALIGN.RIGHT,
                 font=pal["font_body"])
        if delta:
            B.add_tb(slide, delta, x + col_w - 0.82, ty, 0.7, 0.3, size=11,
                     bold=True, color=d_color, align=PP_ALIGN.RIGHT,
                     font=pal["font_body"])


def _tree_connectors(slide, pal, children, depths, ys, col_w, gap):
    """Elbow connectors parent right edge -> child left edge, text_muted."""
    from palettes import hex_rgb
    for par, kids in children.items():
        if par not in depths:
            continue
        px = 0.7 + (depths[par] - 1) * (col_w + gap) + col_w
        for kid in kids:
            if kid not in depths:
                continue
            kx = 0.7 + (depths[kid] - 1) * (col_w + gap)
            x1, y1, _, _ = B._sc(px, ys[par], 0, 0)
            x2, y2, _, _ = B._sc(kx, ys[kid], 0, 0)
            conn = slide.shapes.add_connector(
                MSO_CONNECTOR.ELBOW, Inches(x1), Inches(y1),
                Inches(x2), Inches(y2))
            conn.line.color.rgb = hex_rgb(pal["text_muted"])
            conn.line.width = Pt(1.25)


LAYOUTS["driver-tree"] = build_driver_tree_slide


# ── stakeholder map (support x influence) ────────────────────────────────────
def build_stakeholder_map_slide(prs, p, pal, ctx):
    """2x2 matrix engine with stakeholder axis defaults; items carrying
    TargetX/TargetY get a thin arrow from current position to a hollow
    (outlined) target circle — the engagement move to make."""
    q = dict(p)
    q.setdefault("x_axis", "Support →")
    q.setdefault("y_axis", "Influence →")
    # one group for the whole exhibit: the matrix is left flat here and the
    # arrows/rings drawn below join it in a single group_shapes() call
    body = []
    slide = build_matrix_slide(prs, q, pal, ctx, group=False, body_out=body)
    body_start = body[0]

    L, T, W, H = 2.0, 1.85, 9.6, 4.6  # build_matrix_slide plot area
    from charts import add_arrowhead_to_connector
    from palettes import hex_rgb

    def place(fx, fy):  # fraction -> clamped plot coordinates
        return (max(L + 0.09, min(L + W - 0.09, L + fx * W)),
                max(T + 0.09, min(T + H - 0.09, T + (1 - fy) * H)))

    for item in p.get("matrix_items", [])[:12]:
        if not (item.get("targetx") or item.get("targety")):
            continue
        try:
            fx, fy = float(item.get("x", "")), float(item.get("y", ""))
            tx, ty = float(item["targetx"]), float(item["targety"])
        except (KeyError, ValueError):
            warn(f"stakeholder-map item {item.get('name', '')!r} needs "
                 "numeric X/Y and both TargetX/TargetY — arrow skipped")
            continue
        x0, y0 = place(fx, fy)
        x1, y1 = place(tx, ty)
        # trim both arrow ends so it clears the dot and the hollow ring
        ax0, ay0, ax1, ay1 = x0, y0, x1, y1
        dist = ((x1 - x0) ** 2 + (y1 - y0) ** 2) ** 0.5
        if dist > 0.3:
            ux, uy = (x1 - x0) / dist, (y1 - y0) / dist
            ax0, ay0 = x0 + ux * 0.13, y0 + uy * 0.13
            ax1, ay1 = x1 - ux * 0.12, y1 - uy * 0.12
        bx, by, _, _ = B._sc(ax0, ay0, 0, 0)
        ex, ey, _, _ = B._sc(ax1, ay1, 0, 0)
        conn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,
                                          Inches(bx), Inches(by),
                                          Inches(ex), Inches(ey))
        conn.line.color.rgb = hex_rgb(pal["text_muted"])
        conn.line.width = Pt(1.5)
        add_arrowhead_to_connector(conn)
        ring = B.add_circle(slide, x1 - 0.09, y1 - 0.09, 0.18, pal["bg"],
                            line_hex=pal["accent1"], line_pt=1.5)
        ring.fill.background()  # hollow: future position, not a second dot
    shape_groups.group_shapes(slide, list(slide.shapes)[body_start:])
    return slide


LAYOUTS["stakeholder-map"] = build_stakeholder_map_slide


# ── raci ─────────────────────────────────────────────────────────────────────
RACI_CHIP_KEYS = {"R": "accent1", "A": "accent2", "C": "accent3",
                  "I": "text_muted"}


def build_raci_slide(prs, p, pal, ctx):
    """RACI chart: activity rows x person columns, colored letter chips
    (R=accent1, A=accent2, C=accent3, I=text_muted; blank = no chip)."""
    rows = p.get("table_rows", [])
    if len(rows) < 2:
        warn("raci needs a markdown table; rendering as table")
        return B.build_table_slide(prs, p, pal, ctx)
    slide = B._blank_slide(prs, pal, pal["bg"])
    B._heading(slide, p, pal)
    body_start = len(slide.shapes)
    if len(rows) - 1 > 12:
        warn(f"raci with {len(rows) - 1} activities — consider splitting")

    n_cols = max(len(r) for r in rows)
    first_col_w = 3.6
    col_w = (11.9 - first_col_w) / max(n_cols - 1, 1)
    row_h = max(0.3, min(0.62, 4.4 / len(rows) - 0.06))
    top = 1.9
    for r, row in enumerate(rows):
        y = top + r * (row_h + 0.06)
        if r > 0:
            B.add_rect(slide, 0.7, y, 11.9, row_h,
                       pal["surface"] if r % 2 else pal["bg_deep"])
        for c in range(n_cols):
            cell = (row[c] if c < len(row) else "").strip()
            x = 0.7 + (first_col_w + (c - 1) * col_w if c else 0)
            w = col_w if c else first_col_w
            if r == 0:  # header: activity column + person names
                B.add_tb(slide, cell, x + (0.15 if c == 0 else 0),
                         y + row_h / 2 - 0.16, w, 0.35, size=14, bold=True,
                         color=pal["accent1"],
                         align=PP_ALIGN.CENTER if c else PP_ALIGN.LEFT,
                         font=pal["font_body"])
            elif c == 0:
                B.add_tb(slide, cell, x + 0.15, y + row_h / 2 - 0.16,
                         w - 0.2, 0.35, size=13, color=pal["text"],
                         font=pal["font_body"])
            elif cell:
                key = RACI_CHIP_KEYS.get(cell.upper())
                if not key:  # validator errors on these; render muted text
                    B.add_tb(slide, cell, x, y + row_h / 2 - 0.16, w, 0.35,
                             size=12, color=pal["text_muted"],
                             align=PP_ALIGN.CENTER, font=pal["font_body"])
                    continue
                d = max(0.2, min(0.42, row_h - 0.14))
                fill = pal[key]
                B.add_circle(slide, x + w / 2 - d / 2, y + row_h / 2 - d / 2,
                             d, fill)
                B.add_tb(slide, cell.upper(), x + w / 2 - d / 2,
                         y + row_h / 2 - 0.15, d, 0.3, size=12, bold=True,
                         color=_label_color_on(fill, pal),
                         align=PP_ALIGN.CENTER, font=pal["font_body"])
    legend_y = min(6.45, top + len(rows) * (row_h + 0.06) + 0.08)
    B.add_tb(slide,
             "R = Responsible · A = Accountable · C = Consulted · I = Informed",
             0.7, legend_y, 11.9, 0.3, size=11, color=pal["text_muted"],
             font=pal["font_label"])
    shape_groups.group_shapes(slide, list(slide.shapes)[body_start:])
    return slide


LAYOUTS["raci"] = build_raci_slide


# ── shape-built chart exhibits live in builders_exhibits.py (file-size split);
#    re-exported here so existing imports and LAYOUT_MAP registration keep working.
import builders_exhibits  # noqa: E402
from builders_exhibits import (  # noqa: E402,F401
    NEG_HEX, build_waterfall_slide, build_funnel_slide, build_mekko_slide,
    build_bar_mekko_slide, build_tornado_slide, build_football_field_slide,
    build_gantt_slide,
    _bracket_on_waterfall, _delta_bars, _fmt_num, _fmt_plain,
    _parse_marker)

LAYOUTS.update(builders_exhibits.LAYOUTS)
