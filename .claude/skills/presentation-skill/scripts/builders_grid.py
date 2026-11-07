"""Process-grid layouts: what-how-who.

Registered into builders.LAYOUT_MAP at import time (see bottom of
builders.py).  All geometry is design inches on the 13.33 x 7.5 canvas;
builders' primitives rescale to the actual slide size.  Content must clear
the footer zone (nothing below y = 7.0)."""
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches

import builders as B
import shape_groups
from helpers import warn

MAX_WHW_STEPS = 5
WHW_ROWS = (("WHAT", "what"), ("HOW", "how"), ("WHO", "who"))
_LABEL_W = 1.0   # left gutter for the WHAT / HOW / WHO labels
_GAP = 0.10


def build_what_how_who_slide(prs, p, pal, ctx):
    """Process grid: chevron step header over a striped WHAT/HOW/WHO matrix.
    Row syntax: '- Step N: Title=".." What=".." How=".." Who=".."' (2-5)."""
    steps = p.get("steps", [])
    if len(steps) > MAX_WHW_STEPS:
        warn(f"what-how-who renders the first {MAX_WHW_STEPS} steps — "
             f"dropping {len(steps) - MAX_WHW_STEPS}")
        steps = steps[:MAX_WHW_STEPS]
    if len(steps) < 2:
        warn("what-how-who needs 2+ '- Step N:' rows; rendering bullet-list")
        return B.build_bullet_slide(prs, p, pal, ctx)
    slide = B._blank_slide(prs, pal, pal["bg"])
    B._heading(slide, p, pal)
    body_start = len(slide.shapes)

    from palettes import hex_rgb
    n = len(steps)
    grid_x = 0.7 + _LABEL_W + _GAP
    col_w = (11.9 - _LABEL_W - _GAP * n) / n
    head_y, head_h = 1.9, 0.85
    for i, step in enumerate(steps):
        x = grid_x + i * (col_w + _GAP)
        sl, st, ww, hh = B._sc(x, head_y, col_w, head_h)
        chev = slide.shapes.add_shape(52, Inches(sl), Inches(st),
                                      Inches(ww), Inches(hh))  # 52 = CHEVRON
        chev.fill.solid()
        chev.fill.fore_color.rgb = hex_rgb(
            pal["accent1"] if i == 0 else pal["surface"])
        chev.line.fill.background()
        B.add_tb(slide, step.get("title", f"Step {i + 1}"), x + 0.28,
                 head_y + head_h / 2 - 0.22, col_w - 0.40, 0.44, size=13,
                 bold=True, color=pal["bg_deep"] if i == 0 else pal["text"],
                 align=PP_ALIGN.CENTER, font=pal["font_body"])

    top = head_y + head_h + 0.15
    bottom = B.content_bottom(p)
    row_h = (bottom - top - 2 * 0.08) / 3
    for r, (label, key) in enumerate(WHW_ROWS):
        y = top + r * (row_h + 0.08)
        band = pal["surface"] if r % 2 == 0 else pal["bg_deep"]
        B.add_rect(slide, grid_x, y, 11.9 - _LABEL_W - _GAP, row_h, band)
        B.add_tb(slide, label, 0.7, y + row_h / 2 - 0.16, _LABEL_W, 0.32,
                 size=11, bold=True, color=pal["text_muted"],
                 font=pal["font_label"])
        for i, step in enumerate(steps):
            x = grid_x + i * (col_w + _GAP)
            B.add_tb(slide, step.get(key, ""), x + 0.14, y + 0.10,
                     col_w - 0.28, row_h - 0.20, size=12, color=pal["text"],
                     font=pal["font_body"], accent=pal["accent1"])
    shape_groups.group_shapes(slide, list(slide.shapes)[body_start:])
    return slide


LAYOUTS = {"what-how-who": build_what_how_who_slide}
