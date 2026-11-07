"""Banking / IB layouts: tombstone-grid, league-table, logo-wall.

Registered into builders.LAYOUT_MAP at import time (see bottom of
builders.py).  All geometry is design inches on the 13.33 x 7.5 canvas;
builders' primitives rescale to the actual slide size.  Content must clear
the footer zone (nothing below y = 7.0)."""
import re
from pathlib import Path

from pptx.enum.text import PP_ALIGN

import builders as B
import cell_shapes
import shape_groups
from helpers import warn

MAX_TOMBSTONES = 12
MAX_LEAGUE_ROWS = 10   # body rows below the header
MAX_LOGOS = 24

_NUM_RX = re.compile(r"-?\d[\d,]*\.?\d*")


def _num(cell):
    """Numeric value of a cell ('$1,240', '24.1%', '3.2x') or None."""
    s = str(cell).strip().lstrip("$€£").rstrip("%xX×").replace(",", "").strip()
    return float(s) if _NUM_RX.fullmatch(s) else None


# ── tombstone grid (deal credentials) ────────────────────────────────────────
def build_tombstone_grid_slide(prs, p, pal, ctx):
    """Bordered deal-credential cards in an auto grid.  Row syntax:
    '- Tombstone: Client=".." Deal=".." Value=".." Date=".." [Role=".."]'."""
    stones = [t for t in p.get("tombstones", [])
              if t.get("client", "").strip()]
    if len(stones) > MAX_TOMBSTONES:
        warn(f"tombstone-grid renders the first {MAX_TOMBSTONES} tombstones "
             f"— dropping {len(stones) - MAX_TOMBSTONES}")
        stones = stones[:MAX_TOMBSTONES]
    if not stones:
        warn("tombstone-grid needs '- Tombstone: Client=\"..\" Deal=\"..\"' "
             "rows; rendering bullet-list")
        return B.build_bullet_slide(prs, p, pal, ctx)
    slide = B._blank_slide(prs, pal, pal["bg"])
    B._heading(slide, p, pal)
    body_start = len(slide.shapes)

    n = len(stones)
    cols = min(n, 3) if n <= 6 else 4
    rows = -(-n // cols)
    gap = 0.3
    top, bottom = 1.9, B.content_bottom(p)
    card_w = (11.9 - gap * (cols - 1)) / cols
    card_h = (bottom - top - gap * (rows - 1)) / rows
    for i, t in enumerate(stones):
        r, c = divmod(i, cols)
        x = 0.7 + c * (card_w + gap)
        y = top + r * (card_h + gap)
        # bordered card only — no accent strip (a thin accent rect at the
        # card top trips pptx_lint's "accent line under title" AI-tell)
        B.add_rect(slide, x, y, card_w, card_h, pal["bg"],
                   line_hex=pal["text_muted"], line_pt=1.0)
        B.add_tb(slide, t.get("client", ""), x + 0.12, y + 0.06 * card_h,
                 card_w - 0.24, 0.38, size=13, bold=True, color=pal["text"],
                 align=PP_ALIGN.CENTER, font=pal["font_body"])
        if t.get("deal"):
            B.add_tb(slide, t["deal"], x + 0.12, y + 0.28 * card_h,
                     card_w - 0.24, 0.34, size=10.5, color=pal["text_muted"],
                     align=PP_ALIGN.CENTER, font=pal["font_label"])
        if t.get("role"):
            B.add_tb(slide, t["role"], x + 0.12, y + 0.45 * card_h,
                     card_w - 0.24, 0.30, size=10, color=pal["text_muted"],
                     align=PP_ALIGN.CENTER, font=pal["font_label"])
        if t.get("value"):
            B.add_tb(slide, t["value"], x + 0.12, y + 0.62 * card_h,
                     card_w - 0.24, 0.36, size=14, bold=True,
                     color=pal["accent1"], align=PP_ALIGN.CENTER,
                     font=pal["font_body"])
        if t.get("date"):
            B.add_tb(slide, t["date"], x + 0.12, y + 0.80 * card_h,
                     card_w - 0.24, 0.28, size=10, color=pal["text_muted"],
                     align=PP_ALIGN.CENTER, font=pal["font_label"])
    shape_groups.group_shapes(slide, list(slide.shapes)[body_start:])
    return slide


# ── league table (ranked comparison with in-row bars) ────────────────────────
def _league_columns(header, body):
    """(rank_col, bar_col, numeric_cols) — rank = col 0 when every body cell
    is an integer; bar = the first all-numeric non-rank column."""
    n_cols = max(len(r) for r in [header] + body)
    rank_col = None
    col0 = [_num(r[0]) for r in body if len(r) > 0]
    if col0 and all(v is not None and v == int(v) for v in col0):
        rank_col = 0
    numeric_cols, bar_col = set(), None
    for j in range(n_cols):
        vals = [_num(r[j]) for r in body if j < len(r) and str(r[j]).strip()]
        if vals and all(v is not None for v in vals):
            numeric_cols.add(j)
            if bar_col is None and j != rank_col:
                bar_col = j
    return rank_col, bar_col, numeric_cols


def build_league_table_slide(prs, p, pal, ctx):
    """Shape-drawn ranked table (NOT a native pptx table): accent rank
    numbers, an in-row percent bar behind the first all-numeric non-rank
    column (scaled to the column max), right-aligned numerics."""
    t_rows = p.get("table_rows", [])
    if len(t_rows) < 2:
        warn("league-table needs a markdown table (header + ranked rows); "
             "rendering bullet-list")
        return B.build_bullet_slide(prs, p, pal, ctx)
    header, body = t_rows[0], t_rows[1:]
    if len(body) > MAX_LEAGUE_ROWS:
        warn(f"league-table renders the first {MAX_LEAGUE_ROWS} rows — "
             f"dropping {len(body) - MAX_LEAGUE_ROWS}")
        body = body[:MAX_LEAGUE_ROWS]
    slide = B._blank_slide(prs, pal, pal["bg"])
    B._heading(slide, p, pal)
    body_start = len(slide.shapes)

    rank_col, bar_col, numeric_cols = _league_columns(header, body)
    n_cols = max(len(r) for r in [header] + body)
    weights = [0.5 if j == rank_col else
               (1.0 if j in numeric_cols else 2.0) for j in range(n_cols)]
    total_wt = sum(weights)
    col_ws = [11.9 * wt / total_wt for wt in weights]
    col_xs, x = [], 0.7
    for w in col_ws:
        col_xs.append(x)
        x += w
    bar_max = max((v for r in body if bar_col < len(r)
                   for v in [_num(r[bar_col])] if v is not None),
                  default=0) if bar_col is not None else 0

    head_y = 1.95
    for j, cell in enumerate(header[:n_cols]):
        align = PP_ALIGN.RIGHT if j in numeric_cols or j == rank_col \
            else PP_ALIGN.LEFT
        B.add_tb(slide, str(cell), col_xs[j] + 0.08, head_y,
                 col_ws[j] - 0.16, 0.32, size=11, bold=True,
                 color=pal["text_muted"], align=align, font=pal["font_label"])
    B.add_rect(slide, 0.7, head_y + 0.40, 11.9, 0.02, pal["accent1"])

    top = head_y + 0.52
    row_h = min(0.56, (B.content_bottom(p) - top) / len(body))
    for i, row in enumerate(body):
        y = top + i * row_h
        if i:
            B.add_rect(slide, 0.7, y, 11.9, 0.012, pal["surface"])
        for j, cell in enumerate(row[:n_cols]):
            text = str(cell).strip()
            cx, cw = col_xs[j], col_ws[j]
            if j == rank_col:
                B.add_tb(slide, text, cx + 0.08, y + row_h / 2 - 0.17,
                         cw - 0.16, 0.34, size=13, bold=True,
                         color=pal["accent1"], align=PP_ALIGN.RIGHT,
                         font=pal["font_body"])
                continue
            if j == bar_col and bar_max > 0:
                # strip at the row bottom, vertically clear of the value
                # text so the layer-aware contrast check reads the text
                # against the row band, not the accent fill
                val = _num(text) or 0
                cell_shapes.draw_percent_bar(
                    slide, pal, cx + 0.08, y + row_h - 0.14,
                    cw - 0.24, 0.10, min(100, max(0, round(100 * val / bar_max))))
            align = PP_ALIGN.RIGHT if j in numeric_cols else PP_ALIGN.LEFT
            B.add_tb(slide, text, cx + 0.08, y + row_h / 2 - 0.16,
                     cw - 0.16, 0.32, size=12,
                     bold=(j == bar_col), color=pal["text"], align=align,
                     font=pal["font_body"])
    shape_groups.group_shapes(slide, list(slide.shapes)[body_start:])
    return slide


# ── logo wall (client / partner grid) ────────────────────────────────────────
_LOGO_RX = re.compile(r'^(?:(\S+)\s+)?"([^"]+)"$')
_IMG_EXTS = (".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp", ".svg")


def parse_logo(raw):
    """'acme.png "Acme Corp"' -> ('acme.png', 'Acme Corp');
    '"Acme Corp"' / 'Acme Corp' -> (None, 'Acme Corp') (name-only tile);
    'acme-corp.png' -> ('acme-corp.png', 'Acme Corp') (name from the stem)."""
    raw = str(raw).strip()
    m = _LOGO_RX.match(raw)
    if m:
        return m.group(1), m.group(2)
    if raw.lower().endswith(_IMG_EXTS):
        name = Path(raw).stem.replace("-", " ").replace("_", " ").title()
        return raw, name
    return None, raw


def build_logo_wall_slide(prs, p, pal, ctx):
    """Uniform tile grid of contain-fitted logos with muted captions.
    Row syntax: '- Logo: acme.png "Acme Corp"' (path optional -> name tile);
    a missing file warns and falls back to a name tile."""
    from helpers import resolve_image_path
    logos = [parse_logo(s) for s in p.get("logos", []) if str(s).strip()]
    if len(logos) > MAX_LOGOS:
        warn(f"logo-wall renders the first {MAX_LOGOS} logos — "
             f"dropping {len(logos) - MAX_LOGOS}")
        logos = logos[:MAX_LOGOS]
    if not logos:
        warn("logo-wall needs '- Logo: file.png \"Name\"' rows; "
             "rendering bullet-list")
        return B.build_bullet_slide(prs, p, pal, ctx)
    slide = B._blank_slide(prs, pal, pal["bg"])
    B._heading(slide, p, pal)
    body_start = len(slide.shapes)

    n = len(logos)
    cols = min(n, 4) if n <= 8 else (5 if n <= 15 else 6)
    rows = -(-n // cols)
    gap = 0.25
    top, bottom = 1.9, B.content_bottom(p)
    tile_w = (11.9 - gap * (cols - 1)) / cols
    tile_h = (bottom - top - gap * (rows - 1)) / rows
    for i, (path, name) in enumerate(logos):
        r, c = divmod(i, cols)
        x = 0.7 + c * (tile_w + gap)
        y = top + r * (tile_h + gap)
        B.add_rect(slide, x, y, tile_w, tile_h, pal["surface"])
        img = resolve_image_path(path, ctx) if path else None
        if path and not img:
            warn(f"logo-wall: image not found: {path} — name tile fallback")
        if img:
            B.add_picture_contain(slide, img, x + 0.20, y + 0.15,
                                  tile_w - 0.40, tile_h - 0.62, alt=name)
            B.add_tb(slide, name, x + 0.08, y + tile_h - 0.40,
                     tile_w - 0.16, 0.32, size=10, color=pal["text_muted"],
                     align=PP_ALIGN.CENTER, font=pal["font_label"])
        else:
            B.add_tb(slide, name, x + 0.10, y + tile_h / 2 - 0.24,
                     tile_w - 0.20, 0.48, size=12, bold=True,
                     color=pal["text"], align=PP_ALIGN.CENTER,
                     font=pal["font_label"])
    shape_groups.group_shapes(slide, list(slide.shapes)[body_start:])
    return slide


LAYOUTS = {
    "tombstone-grid": build_tombstone_grid_slide,
    "league-table": build_league_table_slide,
    "logo-wall": build_logo_wall_slide,
}
