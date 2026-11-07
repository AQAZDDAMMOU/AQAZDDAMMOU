#!/usr/bin/env python3
"""
geometry_balance.py — the AeSlides centroid visual-mass balance metric and
the --render-dir local-variance whitespace map, split out of
geometry_report.py so that module stays focused on its overlap/gap/
near-miss findings.

geometry_report.py imports every name defined here and re-exports it, so
`from geometry_report import X` keeps working for callers written against
the pre-split module (qa_summary.run_geometry among them).

Only `_centroid` reaches back into geometry_report (for GRID_STEP_IN, the
rasterization resolution shared with the whitespace-ratio grid); that
import is function-local, mirroring verbatim.py's pattern, so this module
loads cleanly whichever of the two is imported first.
"""

BALANCE_X_TOL = 0.05        # centroid x tolerance (AeSlides), -1..1 scale
BALANCE_Y_TOL = 0.15        # centroid y tolerance (AeSlides), -1..1 scale
TITLE_ZONE_IN = 1.2         # text starting above this = the title shape
FOOTER_ZONE_FRAC = 0.92     # boxes starting below this * height = footer
BALANCE_SKIP_FRAME_FRAC = 0.25  # one chart/picture over this share of the
                            # canvas: a box centroid says nothing useful
BALANCE_MIN_BOXES = 4       # fewer shapes = title/quote/big-number slide,
                            # centred by design (cf. pptx_lint's
                            # MIN_BODY_SHAPES)
RULE_MAX_IN = 0.083         # 6pt — a textless band thinner than this on
                            # either axis is a divider rule, not content
                            # (cf. pptx_lint's LINE_H_MAX_EMU)
RENDER_COLS, RENDER_ROWS = 16, 9   # whitespace-map grid on a rendered PNG
WS_VAR_MIN = 4.0            # cell grey-level variance below this = empty
RENDER_WS_MAX = 0.55        # empty-cell fraction above this = too sparse


def _is_title_chrome(box):
    """The title shape: a plain text frame starting in the title zone.
    A chart, picture or filled shape that merely *starts* high is content
    — slide-3-style exhibits span the title band and the body alike."""
    return (box["t"] < TITLE_ZONE_IN and box["txt"]
            and not box["solid"] and not box["frame"])


def _is_rule(box):
    """A hairline divider — accent bar, timeline spine, card edge stripe.
    Template furniture with negligible ink, on every slide alike."""
    return not box["txt"] and min(box["w"], box["h"]) <= RULE_MAX_IN


def body_boxes(boxes, sh_in):
    """Boxes minus fixed chrome — the title shape, divider rules and
    anything in the footer zone sit in the same place on every slide, so
    leaving them in the balance metric would let the template mask a
    lopsided body."""
    return [b for b in boxes
            if not _is_title_chrome(b) and not _is_rule(b)
            and b["t"] <= FOOTER_ZONE_FRAC * sh_in]


def frame_dominant(boxes, sw_in, sh_in):
    """Is one chart/picture frame more than a quarter of the canvas?"""
    canvas = sw_in * sh_in
    return canvas > 0 and any(
        b["frame"] and b["w"] * b["h"] > BALANCE_SKIP_FRAME_FRAC * canvas
        for b in boxes)


def _centroid(grid, sw_in, sh_in):
    """Centre of mass of the covered cells (each weighted equally),
    normalized to -1..1 from the slide centre.  (0, 0) when nothing is
    covered — an empty slide is not imbalanced."""
    from geometry_report import GRID_STEP_IN
    sum_x = sum_y = cells = 0
    for y, row in enumerate(grid):
        for x, on in enumerate(row):
            if on:
                sum_x += x + 0.5
                sum_y += y + 0.5
                cells += 1
    if not cells:
        return 0.0, 0.0
    cx = ((sum_x / cells) * GRID_STEP_IN - sw_in / 2) / (sw_in / 2)
    cy = ((sum_y / cells) * GRID_STEP_IN - sh_in / 2) / (sh_in / 2)
    return round(cx, 3), round(cy, 3)


def _largest_empty_block(grid):
    """Largest all-empty axis-aligned rectangle, as (cols, rows).  The grid
    is 16x9, so the histogram scan is trivially cheap."""
    if not grid or not grid[0]:
        return 0, 0
    cols = len(grid[0])
    heights = [0] * cols
    best = (0, 0, 0)  # area, cols, rows
    for row in grid:
        for c in range(cols):
            heights[c] = heights[c] + 1 if row[c] else 0
        for c in range(cols):
            run = heights[c]
            for c2 in range(c, cols):
                run = min(run, heights[c2])
                if not run:
                    break
                width = c2 - c + 1
                if width * run > best[0]:
                    best = (width * run, width, run)
    return best[1], best[2]


def render_whitespace(png_path):
    """Empty-cell map of an already-rendered slide PNG (the mirror image of
    the `crowded` check: too much empty space rather than too little).

    A 16x9 cell grid over the greyscale image; a cell whose grey-level
    variance is below WS_VAR_MIN carries no ink.  Returns None — silently,
    it is an optional signal — when Pillow is missing or the file is not a
    readable image.
    """
    try:
        from PIL import Image, ImageStat
    except ImportError:
        return None
    try:
        with Image.open(png_path) as img:
            gray = img.convert("L")
            width, height = gray.size
            empty = [[False] * RENDER_COLS for _ in range(RENDER_ROWS)]
            for r in range(RENDER_ROWS):
                for c in range(RENDER_COLS):
                    cell = gray.crop((
                        width * c // RENDER_COLS, height * r // RENDER_ROWS,
                        width * (c + 1) // RENDER_COLS,
                        height * (r + 1) // RENDER_ROWS))
                    if ImageStat.Stat(cell).var[0] < WS_VAR_MIN:
                        empty[r][c] = True
    except (OSError, ValueError):
        return None
    blank = sum(sum(row) for row in empty)
    block_c, block_r = _largest_empty_block(empty)
    return {
        "render_empty_frac": round(blank / (RENDER_COLS * RENDER_ROWS), 3),
        "render_empty_block": f"{block_c}x{block_r}",
    }
