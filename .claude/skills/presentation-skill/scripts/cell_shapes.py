"""cell_shapes.py — multi-state micro-shapes for table cells and bullets.

The shared vocabulary behind per-cell states (Cycle 7): N-segment harvey
balls, traffic lights, percent bars, rating dots and checkboxes.  A cell (or
bullet prefix) declares its state with a token:

    harvey:3        3 of 4 segments filled
    harvey:5/8      5 of 8 segments
    rag:green       traffic light (green | amber | red, aliases g/a/r,
                    yellow == amber)
    bar:64%         percent bar (0-100, '%' optional)
    dots:3/5        3 of 5 rating dots (default total 5: 'dots:3')
    [x] / [ ]       checkbox (also [X])

`parse_cell_token` recognises a token (None for plain text); `draw_cell_token`
draws it centred in a design-inch box.  Geometry is design-space inches —
every draw goes through builders' scaled helpers, so template mode rescales
these like any other shape.  Nothing here imports build_deck.
"""
import re

from pptx.util import Inches

from helpers import warn

TOKEN_RX = re.compile(
    r"^(?:(harvey):(\d+)(?:/(\d+))?"
    r"|(rag):([a-z]+)"
    r"|(bar):(\d{1,3})%?"
    r"|(dots):(\d+)(?:/(\d+))?"
    r"|(\[[xX ]?\]))$")

RAG_ALIASES = {"green": "good", "g": "good", "amber": "mid", "a": "mid",
               "yellow": "mid", "y": "mid", "red": "bad", "r": "bad"}

DOTS_DEFAULT_TOTAL = 5
HARVEY_DEFAULT_TOTAL = 4


def parse_cell_token(text):
    """(kind, args) for a state token, or None for plain text.

    kinds: ('harvey', filled, total) / ('rag', 'good'|'mid'|'bad') /
    ('bar', pct) / ('dots', filled, total) / ('check', bool).
    A malformed but token-shaped value (rag:blue, harvey:9/4) warns and
    returns None so it renders as literal text the author can spot.
    """
    m = TOKEN_RX.match(text.strip())
    if not m:
        return None
    if m.group(1):  # harvey
        filled, total = int(m.group(2)), int(m.group(3) or HARVEY_DEFAULT_TOTAL)
        if not 0 < total <= 12 or filled > total:
            warn(f"cell token {text!r}: need 0 <= filled <= total <= 12 "
                 "— rendered as text")
            return None
        return ("harvey", filled, total)
    if m.group(4):  # rag
        state = RAG_ALIASES.get(m.group(5))
        if state is None:
            warn(f"cell token {text!r}: unknown state (green|amber|red) "
                 "— rendered as text")
            return None
        return ("rag", state)
    if m.group(6):  # bar
        pct = int(m.group(7))
        if pct > 100:
            warn(f"cell token {text!r}: percent over 100 — rendered as text")
            return None
        return ("bar", pct)
    if m.group(8):  # dots
        filled, total = int(m.group(9)), int(m.group(10) or DOTS_DEFAULT_TOTAL)
        if not 0 < total <= 10 or filled > total:
            warn(f"cell token {text!r}: need 0 <= filled <= total <= 10 "
                 "— rendered as text")
            return None
        return ("dots", filled, total)
    return ("check", "x" in m.group(11).lower())


def draw_cell_token(slide, pal, token, left, top, w, h):
    """Draw a parsed token centred in the design-inch box; returns the shapes."""
    kind = token[0]
    cx, cy = left + w / 2, top + h / 2
    if kind == "harvey":
        d = min(h * 0.72, 0.34)
        return draw_harvey(slide, pal, cx - d / 2, cy - d / 2, d,
                           token[1], token[2])
    if kind == "rag":
        d = min(h * 0.62, 0.28)
        return draw_traffic_light(slide, pal, cx - d / 2, cy - d / 2, d,
                                  token[1])
    if kind == "bar":
        bw, bh = min(w * 0.8, 1.4), 0.16
        return draw_percent_bar(slide, pal, cx - bw / 2, cy - bh / 2,
                                bw, bh, token[1])
    if kind == "dots":
        return draw_rating_dots(slide, pal, cx, cy, token[1], token[2])
    if kind == "check":
        d = min(h * 0.62, 0.26)
        return draw_checkbox(slide, pal, cx - d / 2, cy - d / 2, d, token[1])
    return []


def draw_harvey(slide, pal, x, y, d, filled, total=HARVEY_DEFAULT_TOTAL):
    """N-segment harvey ball at (x, y) diameter d: outline ring + pie wedge.
    Generalises the 0-4 quarters harvey to any total (think-cell allows N)."""
    import builders as B
    shapes = [B.add_circle(slide, x, y, d,
                           pal["bg_deep"] if pal["dark"] else "FFFFFF",
                           line_hex=pal["accent1"], line_pt=1.2)]
    if filled >= total:
        shapes.append(B.add_circle(slide, x + 0.03, y + 0.03, d - 0.06,
                                   pal["accent1"]))
    elif filled > 0:
        sl, st, sw, _ = B._sc(x + 0.03, y + 0.03, d - 0.06, d - 0.06)
        pie = slide.shapes.add_shape(142, Inches(sl), Inches(st),
                                     Inches(sw), Inches(sw))
        pie.fill.solid()
        from palettes import hex_rgb
        pie.fill.fore_color.rgb = hex_rgb(pal["accent1"])
        pie.line.fill.background()
        # OOXML pie angles: 60000ths of a degree; python-pptx adjustments
        # write value*100000 -> multiply degrees by 0.6.  Start at 12
        # o'clock (270), sweep clockwise, keep angles in 0..360.
        sweep = 360.0 * filled / total
        pie.adjustments[0] = 270.0 * 0.6
        pie.adjustments[1] = ((270.0 + sweep) % 360.0) * 0.6
        shapes.append(pie)
    return shapes


def draw_traffic_light(slide, pal, x, y, d, state):
    """One RAG lamp: filled circle in the palette's rag_{good,mid,bad}."""
    import builders as B
    return [B.add_circle(slide, x, y, d, pal[f"rag_{state}"],
                         line_hex=pal["text_muted"], line_pt=0.75)]


def draw_percent_bar(slide, pal, x, y, w, h, pct):
    """Track + accent fill proportional to pct (0-100)."""
    import builders as B
    shapes = [B.add_rect(slide, x, y, w, h, pal["surface"])]
    if pct > 0:
        shapes.append(B.add_rect(slide, x, y, w * pct / 100, h, pal["accent1"]))
    return shapes


def draw_rating_dots(slide, pal, cx, cy, filled, total=DOTS_DEFAULT_TOTAL):
    """Row of `total` dots centred on (cx, cy), first `filled` in accent."""
    import builders as B
    d, gap = 0.14, 0.06
    x = cx - (total * d + (total - 1) * gap) / 2
    shapes = []
    for i in range(total):
        if i < filled:
            shapes.append(B.add_circle(slide, x, cy - d / 2, d, pal["accent1"]))
        else:
            ring = B.add_circle(slide, x, cy - d / 2, d, pal["bg"],
                                line_hex=pal["text_muted"], line_pt=1.0)
            ring.fill.background()
            shapes.append(ring)
        x += d + gap
    return shapes


def draw_checkbox(slide, pal, x, y, d, checked):
    """Rounded box; checked adds an accent tick (text run, font-independent)."""
    import builders as B
    from pptx.enum.text import PP_ALIGN
    box = B.add_round_rect(slide, x, y, d, d,
                           pal["surface"] if checked else pal["bg"],
                           line_hex=pal["accent1" if checked else "text_muted"],
                           line_pt=1.2)
    shapes = [box]
    if checked:
        shapes.append(B.add_tb(slide, "✓", x - 0.02, y - 0.05, d + 0.04,
                               d + 0.04, size=12, bold=True,
                               color=pal["accent1"], font=pal["font_label"],
                               align=PP_ALIGN.CENTER))
    return shapes
