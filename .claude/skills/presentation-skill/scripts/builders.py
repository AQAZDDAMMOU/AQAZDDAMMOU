"""Layout builders — one function per layout name in the outline vocabulary.

Every documented layout in references/generation-guide.md has a builder here;
LAYOUT_MAP at the bottom is the single source of truth for valid layout names.
Builder signature: build_x(prs, p, pal, ctx) -> slide
  p   = parsed slide dict from build_deck.parse_outline
  ctx = {"outline_dir": Path, "assets_dir": Path}
"""
import re

from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches

import shape_groups
from helpers import (_add_filled_shape, add_picture_cover, add_soft_shadow,
                     parse_visual, parse_visual_opts, resolve_image_path,
                     set_fill_alpha, set_picture_alpha, set_picture_duotone,
                     set_slide_bg_gradient, warn)
from helpers import add_circle as _add_circle
from helpers import add_overlay as _add_overlay
from helpers import add_picture_contain as _add_picture_contain
from helpers import add_rect as _add_rect
from helpers import add_tb as _add_tb

# Builders design on a 13.33 x 7.5 canvas; set_canvas() rescales everything
# to the actual slide size (e.g. a user template that is 10 x 7.5).
SLIDE_W, SLIDE_H = 13.33, 7.5
_SCALE = {"x": 1.0, "y": 1.0, "font": 1.0}

# Density: compact (default — double-density: bullets flow into two columns,
# exec points into a 2-column card grid) vs comfortable (single column).
# Geometry chosen so the last row clears the footer zone (y >= 7.0).
DENSITY_TABLE = {
    "compact": {
        "bullet_max": 14, "bullet_cols": 2, "bullet_step": 0.74,
        "bullet_y": 1.70, "bullet_size": 14,
        "twocol_max": 8, "twocol_step": 0.585, "twocol_size": 13,
        "exec_max": 8, "exec_cols": 2, "exec_h": 1.02, "exec_gap": 0.12,
        "exec_y": 1.75, "exec_size": 13,
        "cmp_max": 9, "cmp_step": 0.46, "cmp_size": 12,
    },
    "comfortable": {
        "bullet_max": 7, "bullet_cols": 1, "bullet_step": 0.74,
        "bullet_y": 1.70, "bullet_size": 16,
        "twocol_max": 5, "twocol_step": 0.88, "twocol_size": 16,
        "exec_max": 5, "exec_cols": 1, "exec_h": 0.95, "exec_gap": 0.10,
        "exec_y": 1.75, "exec_size": 14,
        "cmp_max": 6, "cmp_step": 0.70, "cmp_size": 14,
    },
}
DEFAULT_DENSITY = "compact"
_DENSITY = {"mode": DEFAULT_DENSITY}


def set_density(mode):
    _DENSITY["mode"] = mode if mode in DENSITY_TABLE else DEFAULT_DENSITY


def density():
    return DENSITY_TABLE[_DENSITY["mode"]]


# Bullet text must clear the footer zone (footer/page number at y=7.08).
BULLET_BOTTOM = 7.0

# Documented bullet floor: normAutofit shrink for an overflowing bullet list
# never drops the fontScale below this percent.
BULLET_FLOOR_PCT = 80.0

# Kicker band geometry (slide_chrome.add_kicker_box draws at y=6.18) and the
# standard content floor. Builders that fill toward the bottom must budget
# against content_bottom(p) so a declared kicker never overlays content.
KICKER_TOP = 6.18
CONTENT_BOTTOM = 6.85


def content_bottom(p):
    """Bottom y (design inches) available to slide content: above the kicker
    band when the slide declares one, above the `- Footnote:` block when the
    slide declares footnotes, else the standard content floor."""
    bottom = KICKER_TOP - 0.08 if p.get("kicker") else CONTENT_BOTTOM
    if p.get("footnotes"):
        from slide_chrome import footnote_top
        bottom = min(bottom, footnote_top(p) - 0.06)
    return bottom


def bullet_geometry():
    """bullet-list text geometry, shared with build_deck.validate's overflow
    check so the capacity budget can't drift from the builder."""
    d = density()
    cols = d.get("bullet_cols", 1)
    gap = 0.4
    return {"cols": cols, "gap": gap,
            "col_w": (11.93 - gap * (cols - 1)) / cols,
            "font_pt": d["bullet_size"], "step": d["bullet_step"],
            "avail_h": BULLET_BOTTOM - d["bullet_y"], "max": d["bullet_max"]}


def twocol_geometry():
    """two-column-split bullet geometry (single text column beside the visual)."""
    d = density()
    return {"cols": 1, "gap": 0.0, "col_w": 6.1,
            "font_pt": d["twocol_size"], "step": d["twocol_step"],
            "avail_h": BULLET_BOTTOM - 1.74, "max": d["twocol_max"]}


def _autofit_overflow(boxes, bullets, geo):
    """Bullet block over capacity OR per-slot collision: shrink via normAutofit.

    Triggers when:
      - The column's total height ratio exceeds 1.0 (existing behaviour), OR
      - Any non-terminal bullet's wrapped height exceeds its fixed step slot
        (per-slot collision guard — catches a single long bullet that would
        render on top of the next bullet even when total ratio is under 1.0).

    Scale is floored at BULLET_FLOOR_PCT (the documented bullet floor),
    snapped to PowerPoint's 2.5% fontScale grid via textfit.snap_to_grid,
    and picked as the tightest needed across both checks.
    """
    if not bullets:
        return
    import textfit
    ratio, _, overflowing = textfit.bullets_fit(
        bullets, geo["font_pt"], geo["col_w"],
        geo["avail_h"], cols=geo["cols"], step_in=geo["step"])

    # Scale needed to bring total column height within budget.
    scale_ratio = (textfit.snap_to_grid(100.0 / ratio, floor_pct=BULLET_FLOOR_PCT)
                   if ratio > 1.0 else 100.0)

    # Scale needed to bring each overflowing per-slot bullet within its step.
    scale_slot = 100.0
    if overflowing and geo["step"] > 0:
        line_h = geo["font_pt"] * textfit.LINE_PITCH / 72.0
        for idx in overflowing:
            plain, has_icon = textfit.strip_markup(bullets[idx])
            w = geo["col_w"] - (textfit.ICON_INDENT if has_icon
                                else textfit.MARKER_INDENT)
            lines = max(textfit.estimate_lines(plain, geo["font_pt"], w), 1)
            needed = textfit.snap_to_grid(
                100.0 * geo["step"] / (lines * line_h), floor_pct=BULLET_FLOOR_PCT)
            scale_slot = min(scale_slot, needed)

    scale = min(scale_ratio, scale_slot)
    if scale >= 100:
        return
    for tb in boxes:
        textfit.apply_autofit(tb.text_frame, scale)


def set_canvas(prs):
    _SCALE["x"] = prs.slide_width / Inches(1) / SLIDE_W
    _SCALE["y"] = prs.slide_height / Inches(1) / SLIDE_H
    _SCALE["font"] = min(_SCALE["x"], _SCALE["y"], 1.0)


def _sc(left, top, w, h):
    return (left * _SCALE["x"], top * _SCALE["y"],
            w * _SCALE["x"], h * _SCALE["y"])


def _fs(size):
    return max(round(size * _SCALE["font"]), 10)  # never below 10pt


def add_tb(slide, text, left, top, w, h, size=16, **kw):
    return _add_tb(slide, text, *_sc(left, top, w, h), size=_fs(size), **kw)


def add_rect(slide, left, top, w, h, *args, **kw):
    return _add_rect(slide, *_sc(left, top, w, h), *args, **kw)


def add_round_rect(slide, left, top, w, h, *args, **kw):
    return _add_filled_shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE,
                             *_sc(left, top, w, h), *args, **kw)


def add_overlay(slide, left, top, w, h, *args, **kw):
    return _add_overlay(slide, *_sc(left, top, w, h), *args, **kw)


def add_picture_contain(slide, img_path, left, top, w, h, **kw):
    return _add_picture_contain(slide, img_path, *_sc(left, top, w, h), **kw)


def add_circle(slide, left, top, d, *args, **kw):
    sl, st, sw, _ = _sc(left, top, d, d)
    return _add_circle(slide, sl, st, sw, *args, **kw)


def _blank_slide(prs, pal, bg_hex, accent_bar=True, gradient=False):
    layout = ctx_blank_layout(prs)
    slide = prs.slides.add_slide(layout)
    if gradient and pal["dark"]:
        set_slide_bg_gradient(slide, pal["bg"], pal["bg_deep"])
    else:
        fill = slide.background.fill
        fill.solid()
        from palettes import hex_rgb
        fill.fore_color.rgb = hex_rgb(bg_hex)
    if accent_bar:
        add_rect(slide, 0, 0, SLIDE_W, 0.05, pal["accent1"])
    return slide


# ── Icon bullets ─────────────────────────────────────────────────────────────
ICON_RX = re.compile(r"^icon:([a-z0-9-]+)\s+(.*)$")


def split_icon(text):
    """'icon:rocket Launch fast' -> ('rocket', 'Launch fast'); else (None, text)."""
    m = ICON_RX.match(text.strip())
    return (m.group(1), m.group(2)) if m else (None, text)


def split_cell_token(text):
    """'rag:green On track' -> (('rag', 'good'), 'On track'); else (None, text).
    Multi-state bullet prefix — the grammar is cell_shapes.parse_cell_token
    (never forked here); a bare token with no trailing text stays plain text."""
    import cell_shapes
    stripped = text.strip()
    if stripped.startswith("[ ]"):   # the one token containing a space
        head, rest = "[ ]", stripped[3:]
    else:
        head, _, rest = stripped.partition(" ")
    if not rest.strip():
        return None, text
    token = cell_shapes.parse_cell_token(head)
    return (token, rest.strip()) if token else (None, text)


def _icon_path(name, pal, ctx):
    try:
        from fetch_icon import fetch_icon
        from pathlib import Path
        cache = Path(ctx.get("assets_dir", "assets")) / "icons"
        return fetch_icon(name, pal["accent1"], cache)
    except Exception as e:
        warn(f"icon '{name}': {e}")
        return None


def _draw_marker(slide, pal, ctx, icon, x, y, token=None):
    """Bullet marker: icon-in-circle when given, a cell_shapes state token
    when given, accent square otherwise.  Returns the x offset where the
    bullet text should start."""
    if icon:
        path = _icon_path(icon, pal, ctx)
        if path:
            circle_fill = pal["surface"] if pal["dark"] else pal["bg_deep"]
            add_circle(slide, x, y - 0.06, 0.46, circle_fill)
            add_picture_contain(slide, path, x + 0.10, y + 0.04, 0.26, 0.26)
            return x + 0.66
    if token:
        import cell_shapes
        w = token[2] * 0.20 + 0.06 if token[0] == "dots" else 0.42
        cell_shapes.draw_cell_token(slide, pal, token, x, y - 0.08, w, 0.40)
        return x + w + 0.12
    add_rect(slide, x, y + 0.05, 0.12, 0.12, pal["accent1"])
    return x + 0.30


def ctx_blank_layout(prs):
    """Find the most blank-like layout (fewest placeholders); index 6 in the
    default template, but arbitrary templates differ."""
    return min(prs.slide_layouts, key=lambda lay: len(lay.placeholders))


def _heading(slide, p, pal, w=11.9, size=32, align=PP_ALIGN.LEFT):
    add_tb(slide, p.get("heading", ""), 0.7, 0.55, w, 1.1,
           size=size, bold=True, color=pal["accent1"],
           font=pal["font_title"], align=align)


def _apply_image_opts(pic, opts, pal):
    """Apply parsed visual options (duotone / alpha) to a placed picture."""
    if not opts:
        return
    if opts.get("duotone"):
        set_picture_duotone(pic, pal["bg"], pal["accent1"])
    if "alpha" in opts:
        set_picture_alpha(pic, opts["alpha"])


def _place_visual(slide, prs, p, pal, ctx, left, top, w, h):
    """Fill a box with the slide's visual: chart, image, or styled placeholder."""
    kind, value, opts = parse_visual_opts(p.get("visual", ""), warn_opts=False)
    if kind == "chart" and p.get("data"):
        from chart_visual import place_chart
        place_chart(slide, p, pal, left, top, w, h, value)
        return
    if kind == "image":
        path = resolve_image_path(value, ctx)
        if path:
            pic = add_picture_contain(slide, path, left, top, w, h,
                                      alt=p.get("caption") or p.get("heading"))
            _apply_image_opts(pic, opts, pal)
            return
        warn(f"Image not found: {value}")
    _visual_placeholder(slide, pal, left, top, w, h)


def _visual_placeholder(slide, pal, left, top, w, h):
    add_rect(slide, left, top, w, h, pal["surface"],
             line_hex=pal["accent1"], line_pt=1.0)
    add_tb(slide, "[ Add Visual ]", left, top + h / 2 - 0.3, w, 0.6,
           size=14, color=pal["text_muted"], align=PP_ALIGN.CENTER,
           font=pal["font_label"])


GHOST_KEEP_REAL = {"title", "section-divider", "closing", "agenda"}


def build_ghost_slide(prs, p, pal, ctx):
    """Skeleton slide: real heading + dashed placeholder describing the
    planned exhibit. Used by build_deck --ghost for storyline alignment."""
    slide = _blank_slide(prs, pal, pal["bg"])
    _heading(slide, p, pal)
    kind, value = parse_visual(p.get("visual", ""))
    label = f"[ {p.get('layout', 'bullet-list')} ]"
    if kind:
        label += f"   planned visual — {kind}: {value}"
    elif any(p.get(k) for k in ("data", "bars", "matrix_items", "tiles",
                                "milestones", "table_rows", "steps")):
        label += "   planned exhibit from structured rows"
    box = add_rect(slide, 0.7, 1.9, 11.9, 4.6, pal["surface"],
                   line_hex=pal["text_muted"], line_pt=1.0)
    set_fill_alpha(box, 40)
    add_tb(slide, label, 0.9, 2.05, 11.5, 0.4, size=14, bold=True,
           color=pal["text_muted"], font=pal["font_label"])
    for i, b in enumerate(p.get("bullets", [])[:6]):
        add_tb(slide, f"—  {split_icon(b)[1]}", 1.0, 2.7 + i * 0.55, 11.2,
               0.5, size=12, color=pal["text_muted"], font=pal["font_body"])
    return slide


# ── Layouts ──────────────────────────────────────────────────────────────────
def _slide_title(p, default=""):
    """Title text: explicit - Title: wins, then ## Slide N: heading."""
    return p.get("title") or p.get("heading") or default


def build_title_slide(prs, p, pal, ctx):
    # No accent line under the title — a known hallmark of AI-generated
    # slides; hierarchy comes from size contrast and whitespace instead.
    slide = _blank_slide(prs, pal, pal["bg_deep"], gradient=True)
    add_rect(slide, 0, 0, SLIDE_W, 0.07, pal["accent1"])

    has_hero = False
    kind, value, opts = parse_visual_opts(p.get("visual", ""), warn_opts=False)
    if kind == "image":
        path = resolve_image_path(value, ctx)
        if path:
            pic = add_picture_contain(slide, path, 9.3, 1.4, 3.4, 4.7,
                                      alt=_slide_title(p, "hero image"))
            _apply_image_opts(pic, opts, pal)
            has_hero = True

    title_w = 8.2 if has_hero else 11.7
    add_tb(slide, _slide_title(p, "Presentation Title"),
           0.8, 2.1, title_w, 1.8, size=46, bold=True,
           color=pal["accent1"], font=pal["font_title"])
    if p.get("subtitle"):
        add_tb(slide, p["subtitle"], 0.8, 4.15, title_w, 0.8, size=20,
               color=pal["text_muted"], font=pal["font_label"])
    return slide


def build_section_divider_slide(prs, p, pal, ctx):
    slide = _blank_slide(prs, pal, pal["bg_deep"], accent_bar=False, gradient=True)
    add_rect(slide, 0, 3.0, 0.18, 1.5, pal["accent1"])
    add_tb(slide, p.get("heading", p.get("title", "")), 0.8, 2.9, 11.7, 1.7,
           size=44, bold=True, color=pal["text"], font=pal["font_title"])
    if p.get("subtitle"):
        add_tb(slide, p["subtitle"], 0.8, 4.6, 11.7, 0.8, size=18,
               color=pal["text_muted"], font=pal["font_label"])
    return slide


def build_exec_summary_slide(prs, p, pal, ctx):
    d = density()
    slide = _blank_slide(prs, pal, pal["bg"])
    _heading(slide, p, pal)
    body_start = len(slide.shapes)
    points = p.get("bullets", [])[:d["exec_max"]]
    # 3 points always render as 3 full-width cards in a single column — a
    # compact 2-column grid would leave one card alone in an empty second
    # row (2+1), which trips the geometry balance check. 4+ points keep the
    # density table's exec_cols.
    cols = 1 if len(points) == 3 else d.get("exec_cols", 1)
    per_col = -(-len(points) // cols) if points else 1
    gap_x = 0.3
    card_w = (11.9 - gap_x * (cols - 1)) / cols
    card_h = d["exec_h"]
    # Fit rows into the content zone (shrink if over budget, e.g. a kicker
    # is present); center the block vertically when it underfills the zone.
    avail = content_bottom(p) - d["exec_y"]
    block = per_col * card_h + (per_col - 1) * d["exec_gap"]
    if block > avail:
        card_h = (avail - (per_col - 1) * d["exec_gap"]) / per_col
        block = avail
    y0 = d["exec_y"] + (avail - block) / 2
    for i, point in enumerate(points):
        icon, text = split_icon(point)
        col, row = divmod(i, per_col)  # column-major: read down, then across
        x = 0.7 + col * (card_w + gap_x)
        y = y0 + row * (card_h + d["exec_gap"])
        card = add_rect(slide, x, y, card_w, card_h, pal["surface"])
        add_soft_shadow(card)
        add_rect(slide, x, y, 0.08, card_h, pal["accent1"])
        text_x = x + 0.35
        if icon:
            text_x = _draw_marker(slide, pal, ctx, icon, x + 0.28,
                                  y + card_h / 2 - 0.18) + 0.12
        add_tb(slide, text, text_x, y + card_h / 2 - 0.30,
               x + card_w - text_x - 0.2, 0.62,
               size=d["exec_size"], color=pal["text"], font=pal["font_body"],
               accent=pal["accent1"])
    shape_groups.group_shapes(slide, list(slide.shapes)[body_start:])
    return slide


def build_bullet_slide(prs, p, pal, ctx):
    d = density()
    slide = _blank_slide(prs, pal, pal["bg"])
    _heading(slide, p, pal)
    geo = bullet_geometry()
    bullets = p.get("bullets", [])[:geo["max"]]
    cols = geo["cols"]
    per_col = -(-len(bullets) // cols) if bullets else 1
    col_w = geo["col_w"]
    boxes = []
    for i, bullet in enumerate(bullets):
        icon, text = split_icon(bullet)
        token = None
        if not icon:
            token, text = split_cell_token(text)
        col, row = divmod(i, per_col)  # column-major: read down, then across
        x0 = 0.7 + col * (col_w + geo["gap"])
        y = d["bullet_y"] + row * d["bullet_step"]
        text_x = _draw_marker(slide, pal, ctx, icon, x0, y, token)
        boxes.append(add_tb(slide, text, text_x, y - 0.06, x0 + col_w - text_x,
                            d["bullet_step"] - 0.06,
                            size=d["bullet_size"], color=pal["text"],
                            font=pal["font_body"], accent=pal["accent1"]))
    _autofit_overflow(boxes, bullets, geo)
    return slide


def build_two_column_slide(prs, p, pal, ctx):
    slide = _blank_slide(prs, pal, pal["bg"])
    add_tb(slide, p.get("heading", ""), 0.7, 0.55, 6.0, 1.1,
           size=30, bold=True, color=pal["accent1"], font=pal["font_title"])
    d = density()
    geo = twocol_geometry()
    bullets = p.get("bullets", [])[:geo["max"]]
    boxes = []
    for i, bullet in enumerate(bullets):
        icon, text = split_icon(bullet)
        token = None
        if not icon:
            token, text = split_cell_token(text)
        y = 1.74 + i * d["twocol_step"]
        text_x = _draw_marker(slide, pal, ctx, icon, 0.7, y, token)
        boxes.append(add_tb(slide, text, text_x, y - 0.06, 6.8 - text_x,
                            d["twocol_step"] - 0.08,
                            size=d["twocol_size"], color=pal["text"],
                            font=pal["font_body"], accent=pal["accent1"]))
    _autofit_overflow(boxes, bullets, geo)
    _place_visual(slide, prs, p, pal, ctx, 6.9, 0.9, 6.0, 5.9)
    return slide


def _build_cards(prs, p, pal, ctx, per_row, rows):
    slide = _blank_slide(prs, pal, pal["bg"])
    _heading(slide, p, pal)
    body_start = len(slide.shapes)
    cards = p.get("cards", [])[: per_row * rows]
    gap = 0.4
    card_w = (11.9 - gap * (per_row - 1)) / per_row
    bottom = content_bottom(p)
    card_h = (bottom - 1.9) if rows == 1 else (bottom - 1.9 - 0.35) / 2
    for i, card in enumerate(cards):
        row, col = divmod(i, per_row)
        x = 0.7 + col * (card_w + gap)
        y = 1.9 + row * (card_h + 0.35)
        shape = add_rect(slide, x, y, card_w, card_h, pal["surface"])
        add_soft_shadow(shape)
        y_text = y + 0.3
        if card.get("icon"):
            path = _icon_path(card["icon"], pal, ctx)
            if path:
                circle_fill = pal["bg_deep"] if pal["dark"] else pal["bg"]
                add_circle(slide, x + 0.25, y + 0.25, 0.55, circle_fill)
                add_picture_contain(slide, path, x + 0.37, y + 0.37, 0.31, 0.31)
                y_text = y + 0.3 if rows == 1 else y + 0.25
                x_text = x + 0.95
            else:
                x_text = x + 0.25
        else:
            x_text = x + 0.25
        name = card.get("name") or card.get("title", "")
        role = card.get("title", "") if card.get("name") else card.get("subtitle", "")
        body = card.get("bio") or card.get("body") or card.get("desc", "")
        add_tb(slide, name, x_text, y_text, card_w - (x_text - x) - 0.25, 0.6,
               size=18, bold=True, color=pal["text"], font=pal["font_title"])
        if role:
            add_tb(slide, role, x_text, y_text + 0.55, card_w - (x_text - x) - 0.25,
                   0.45, size=13, bold=True, color=pal["accent1"],
                   font=pal["font_body"])
        if body:
            add_tb(slide, body, x + 0.25, y + 1.4, card_w - 0.5, card_h - 1.6,
                   size=13, color=pal["text_muted"], font=pal["font_body"],
                   accent=pal["accent1"])
    shape_groups.group_shapes(slide, list(slide.shapes)[body_start:])
    return slide


def build_cards3_slide(prs, p, pal, ctx):
    return _build_cards(prs, p, pal, ctx, per_row=3, rows=1)


def build_cards4_slide(prs, p, pal, ctx):
    return _build_cards(prs, p, pal, ctx, per_row=2, rows=2)


def build_stat_callout_slide(prs, p, pal, ctx):
    stats = p.get("stats", [])[:3]
    if not stats:
        warn("stat-callout slide has no stats; rendering as bullet-list")
        return build_bullet_slide(prs, p, pal, ctx)
    slide = _blank_slide(prs, pal, pal["bg_deep"])
    _heading(slide, p, pal, size=30, align=PP_ALIGN.CENTER)
    body_start = len(slide.shapes)
    card_w = 3.8
    total_w = len(stats) * card_w + (len(stats) - 1) * 0.45
    start_x = (SLIDE_W - total_w) / 2
    for i, stat in enumerate(stats):
        x = start_x + i * (card_w + 0.45)
        card = add_rect(slide, x, 1.9, card_w, 4.0, pal["surface"],
                        line_hex=pal["accent1"], line_pt=1.5)
        add_soft_shadow(card)
        add_tb(slide, stat.get("value", ""), x + 0.2, 2.4, card_w - 0.4, 1.5,
               size=52, bold=True, color=pal["accent1"],
               align=PP_ALIGN.CENTER, font=pal["font_title"])
        add_tb(slide, stat.get("label", ""), x + 0.2, 4.1, card_w - 0.4, 0.7,
               size=16, bold=True, color=pal["text"],
               align=PP_ALIGN.CENTER, font=pal["font_body"])
        add_tb(slide, stat.get("sublabel", ""), x + 0.2, 4.85, card_w - 0.4, 0.6,
               size=12, color=pal["text_muted"],
               align=PP_ALIGN.CENTER, font=pal["font_label"])
    shape_groups.group_shapes(slide, list(slide.shapes)[body_start:])
    return slide


def build_timeline_slide(prs, p, pal, ctx):
    slide = _blank_slide(prs, pal, pal["bg"])
    _heading(slide, p, pal)
    body_start = len(slide.shapes)
    items = p.get("items", [])[:5]
    if not items:  # fall back to "Date: text" bullets
        for b in p.get("bullets", [])[:5]:
            date, _, rest = b.partition(":")
            items.append({"date": date.strip(), "title": rest.strip()})
    n = max(len(items), 1)
    # Center the block (dates above rail, titles/desc below) in the content zone.
    has_desc = any(i.get("desc") for i in items)
    below = 2.85 if has_desc else 1.2
    zone_top, zone_bot = 1.74, content_bottom(p)
    line_y = zone_top + 0.85 + (zone_bot - zone_top - 0.85 - below) / 2
    add_rect(slide, 0.9, line_y, 11.5, 0.04, pal["surface"])
    step = 11.5 / n
    for i, item in enumerate(items):
        cx = 0.9 + step * i + step / 2
        add_rect(slide, cx - 0.11, line_y - 0.09, 0.22, 0.22, pal["accent1"])
        add_tb(slide, item.get("date", ""), cx - step / 2 + 0.1, line_y - 0.85,
               step - 0.2, 0.5, size=15, bold=True, color=pal["accent1"],
               align=PP_ALIGN.CENTER, font=pal["font_body"])
        add_tb(slide, item.get("title", ""), cx - step / 2 + 0.1, line_y + 0.4,
               step - 0.2, 0.8, size=14, bold=True, color=pal["text"],
               align=PP_ALIGN.CENTER, font=pal["font_body"])
        if item.get("desc"):
            add_tb(slide, item["desc"], cx - step / 2 + 0.1, line_y + 1.25,
                   step - 0.2, 1.6, size=12, color=pal["text_muted"],
                   align=PP_ALIGN.CENTER, font=pal["font_label"])
    shape_groups.group_shapes(slide, list(slide.shapes)[body_start:])
    return slide


def build_comparison_slide(prs, p, pal, ctx):
    slide = _blank_slide(prs, pal, pal["bg"])
    _heading(slide, p, pal)
    body_start = len(slide.shapes)
    halves = [
        (0.7, p.get("left_label", "Before"), p.get("visual_left"), p.get("left_bullets", [])),
        (6.93, p.get("right_label", "After"), p.get("visual_right"), p.get("right_bullets", [])),
    ]
    panel_h = content_bottom(p) - 2.45
    for x, label, visual, bullets in halves:
        col_w = 5.7
        add_tb(slide, label, x, 1.7, col_w, 0.6, size=20, bold=True,
               color=pal["accent1"], align=PP_ALIGN.CENTER, font=pal["font_title"])
        if visual:
            _, v_value, v_opts = parse_visual_opts(visual, warn_opts=False)
            path = resolve_image_path(v_value, ctx)
            if path:
                pic = add_picture_contain(slide, path, x, 2.45, col_w, panel_h,
                                          alt=label)
                _apply_image_opts(pic, v_opts, pal)
                continue
            warn(f"Comparison image not found: {visual}")
        if bullets:
            d = density()
            panel = add_rect(slide, x, 2.45, col_w, panel_h, pal["surface"])
            add_soft_shadow(panel)
            for i, b in enumerate(bullets[:d["cmp_max"]]):
                add_tb(slide, b, x + 0.3, 2.70 + i * d["cmp_step"],
                       col_w - 0.6, d["cmp_step"] - 0.05,
                       size=d["cmp_size"], color=pal["text"], font=pal["font_body"],
                       accent=pal["accent1"])
        else:
            _visual_placeholder(slide, pal, x, 2.45, col_w, panel_h)
    shape_groups.group_shapes(slide, list(slide.shapes)[body_start:])
    return slide


def build_full_image_slide(prs, p, pal, ctx):
    slide = _blank_slide(prs, pal, pal["bg_deep"], accent_bar=False)
    kind, value, opts = parse_visual_opts(p.get("visual", ""), warn_opts=False)
    placed = False
    if kind == "image":
        path = resolve_image_path(value, ctx)
        if path:
            pic = add_picture_cover(slide, prs, path,
                                    alt=p.get("caption", "full-bleed image"))
            _apply_image_opts(pic, opts, pal)
            placed = True
        else:
            warn(f"Full-image visual not found: {value}")
    caption = p.get("caption", "")
    if caption:
        if placed:
            add_overlay(slide, 0, 5.2, SLIDE_W, 2.3, "000000", alpha_pct=50)
        add_tb(slide, caption, 0.7, 5.5, 11.9, 1.6, size=22,
               color="F1F5F9", font=pal["font_body"])
    return slide


def build_closing_slide(prs, p, pal, ctx):
    slide = _blank_slide(prs, pal, pal["bg_deep"], accent_bar=False, gradient=True)
    add_rect(slide, 0, 0, SLIDE_W, 0.07, pal["accent1"])
    add_rect(slide, 0, 7.43, SLIDE_W, 0.07, pal["accent2"])
    add_tb(slide, _slide_title(p, "Thank You"), 1.0, 2.4, 11.33, 1.9,
           size=54, bold=True, color=pal["accent1"],
           align=PP_ALIGN.CENTER, font=pal["font_title"])
    if p.get("subtitle"):
        add_tb(slide, p["subtitle"], 1.0, 4.6, 11.33, 0.9, size=20,
               color=pal["text_muted"], align=PP_ALIGN.CENTER,
               font=pal["font_label"])
    if p.get("contact"):
        # accent1 + bold: WCAG large-text (14pt bold) and readable on light themes,
        # where accent2 support tones lack contrast
        add_tb(slide, p["contact"], 1.0, 5.6, 11.33, 0.6, size=14, bold=True,
               color=pal["accent1"], align=PP_ALIGN.CENTER,
               font=pal["font_body"])
    return slide


LAYOUT_MAP = {
    "title": build_title_slide,
    "section-divider": build_section_divider_slide,
    "exec-summary": build_exec_summary_slide,
    "bullet-list": build_bullet_slide,
    "bullets": build_bullet_slide,
    "two-column-split": build_two_column_slide,
    "cards-3": build_cards3_slide,
    "cards-4": build_cards4_slide,
    "stat-callout": build_stat_callout_slide,
    "timeline": build_timeline_slide,
    "comparison": build_comparison_slide,
    "full-image": build_full_image_slide,
    "closing": build_closing_slide,
}

# Consulting layouts (waterfall, matrix-2x2, harvey-scorecard, ...) register
# themselves here; builders_consulting imports this module lazily, so the
# import must come after LAYOUT_MAP is defined.
import builders_consulting  # noqa: E402
LAYOUT_MAP.update(builders_consulting.LAYOUTS)

# Structural diagram layouts (architecture-layers, swimlane-process,
# org-chart, value-chain) register themselves here the same way.
import builders_structure  # noqa: E402
LAYOUT_MAP.update(builders_structure.LAYOUTS)

# The table family (table, comps-table) lives in builders_table.py; the old
# names are re-exported so `builders.build_table_slide` (the consulting
# builders' fallback) and the TABLE_* constants keep resolving.
import builders_table  # noqa: E402
LAYOUT_MAP.update(builders_table.LAYOUTS)
from builders_table import (build_table_slide,  # noqa: E402,F401
                            TABLE_CELL_PAD_IN, TABLE_FONT_PT, TABLE_MIN_PT)

# Process-grid layout (what-how-who) registers itself here the same way.
import builders_grid  # noqa: E402
LAYOUT_MAP.update(builders_grid.LAYOUTS)

# Banking layouts (tombstone-grid, league-table, logo-wall) likewise.
import builders_banking  # noqa: E402
LAYOUT_MAP.update(builders_banking.LAYOUTS)
