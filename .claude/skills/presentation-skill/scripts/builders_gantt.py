#!/usr/bin/env python3
"""
builders_gantt.py — the gantt / swimlane roadmap exhibit.

Moved out of builders_exhibits.py (Cycle 7 §8, think-cell parity) and grown:

- right-hand label columns from Bar `Responsible=".."` / `Remark=".."` attrs;
- row-spanning left brackets (`- Span: Rows=".." Label=".."`);
- shaded period bands (`- Shade: Periods="3-4" Label=".."`);
- vertical milestone lines (`- Line: At="2.5" Label=".."`);
- `Row="Parent > Child"` hierarchy (bold group headers, indented children);
- fiscal / month period expansion (`**Periods:** quarters FY26-FY27`,
  `**Periods:** months Jan 2027 - Jun 2027`) via `expand_periods()` — the
  expansion happens here in the builder, so replay consumers keep seeing the
  raw front-matter string.

A gantt that uses none of the new keys renders byte-identical to the
pre-split builder. Registered into builders.LAYOUT_MAP through
builders_exhibits.LAYOUTS (which re-exports `build_gantt_slide`).
All geometry is in design inches on the 13.33 x 7.5 canvas.
"""
import re

from pptx.enum.shapes import MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

import builders as B
import shape_groups
from helpers import label_color_on as _label_color_on
from helpers import set_fill_alpha, warn

# ── geometry constants (design inches) ───────────────────────────────────────
GRID_TOP = 2.35        # grid top with no extra header bands (legacy value)
ROW_BUDGET = 4.3       # vertical space for rows below GRID_TOP (legacy value)
BAND_H = 0.34          # height of one extra header band (year / line labels)
LABEL_COL_W = 1.3      # width of one Responsible/Remark column
MAX_LABEL_COLS = 2
SPAN_SHIFT = 0.35      # row-label shift per bracket nesting level (v1: one)
CHILD_INDENT = 0.25    # hierarchy child indent
SHADE_ALPHA = 40       # opacity % of a shaded period band
MAX_DEPTH = 2          # `Parent > Child` — deeper flattens with a warning

_MONTHS = ("Jan", "Feb", "Mar", "Apr", "May", "Jun",
           "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
_FY_RANGE_RX = re.compile(r"^(FY\d{2,4})(?:\s*[-–]\s*(FY\d{2,4}))?$",
                          re.IGNORECASE)
_MONTH_RANGE_RX = re.compile(
    r"^([A-Za-z]+)\s+(\d{4})(?:\s*[-–]\s*([A-Za-z]+)\s+(\d{4}))?$")
_RANGE_RX = re.compile(r"^(\d+)(?:\s*[-–]\s*(\d+))?$")


# ── period expansion ─────────────────────────────────────────────────────────
def expand_periods(value):
    """Expand a `**Periods:**` declaration into `(labels, year_spans)`.

    `value` is the raw front-matter string or the parser's comma-split list.
    Plain comma lists pass through unchanged with no year header. A single
    `quarters FYa[-FYb]` item expands to Q1..Q4 per fiscal year with one
    year-header span per year; `months Mon YYYY - Mon YYYY` expands to month
    names, with year-header spans only when the range crosses a year boundary.
    `year_spans` is a list of (label, first_column_index, column_count).

    Raises ValueError for an unparsable or reversed quarters/months range
    (outline_validate surfaces this as a build error).
    """
    if isinstance(value, str):
        items = [s.strip() for s in value.split(",") if s.strip()]
    else:
        items = [str(s).strip() for s in (value or []) if str(s).strip()]
    if len(items) != 1:
        return items, []
    text = items[0]
    low = text.lower()
    if low.startswith("quarters"):
        return _expand_quarters(text[len("quarters"):].strip())
    if low.startswith("months"):
        return _expand_months(text[len("months"):].strip())
    return items, []


def _fy_year(token):
    n = int(token[2:])
    return n + 2000 if n < 100 else n


def _expand_quarters(spec):
    m = _FY_RANGE_RX.match(spec)
    if not m:
        raise ValueError(
            f"**Periods:** quarters needs 'FYa' or 'FYa-FYb', got {spec!r}")
    a, b = m.group(1), m.group(2) or m.group(1)
    ya, yb = _fy_year(a), _fy_year(b)
    if yb < ya:
        raise ValueError(
            f"**Periods:** quarters range runs backwards: {spec!r}")
    labels, spans = [], []
    for i, year in enumerate(range(ya, yb + 1)):
        fy = f"FY{str(year)[-2:]}" if len(a) <= 4 else f"FY{year}"
        spans.append((fy, i * 4, 4))
        labels += ["Q1", "Q2", "Q3", "Q4"]
    return labels, spans


def _month_index(name):
    try:
        return [mo.lower() for mo in _MONTHS].index(name.strip()[:3].lower())
    except ValueError:
        raise ValueError(f"**Periods:** unknown month name {name!r}") from None


def _expand_months(spec):
    m = _MONTH_RANGE_RX.match(spec)
    if not m:
        raise ValueError(
            "**Periods:** months needs 'Mon YYYY - Mon YYYY', "
            f"got {spec!r}")
    m0, y0 = _month_index(m.group(1)), int(m.group(2))
    m1 = _month_index(m.group(3)) if m.group(3) else m0
    y1 = int(m.group(4)) if m.group(4) else y0
    if (y1, m1) < (y0, m0):
        raise ValueError(f"**Periods:** months range runs backwards: {spec!r}")
    labels, per_year = [], {}
    y, mo = y0, m0
    while (y, mo) <= (y1, m1):
        labels.append(_MONTHS[mo])
        per_year.setdefault(y, [len(labels) - 1, 0])
        per_year[y][1] += 1
        mo += 1
        if mo == 12:
            mo, y = 0, y + 1
    spans = ([(str(y), start, count) for y, (start, count) in per_year.items()]
             if y1 != y0 else [])
    return labels, spans


# ── row plan (hierarchy) ─────────────────────────────────────────────────────
def _canon_row(path):
    """Canonical row key for a Bar/Milestone/Span row reference. Plain names
    stay byte-exact (legacy behavior); `Parent > Child` paths normalize to
    single-space separators, flattened to MAX_DEPTH with a warning."""
    if ">" not in path:
        return path
    segs = [s.strip() for s in path.split(">") if s.strip()]
    if len(segs) > MAX_DEPTH:
        warn(f"gantt: Row {path!r} deeper than {MAX_DEPTH} levels — flattened")
        segs = [segs[0], " > ".join(segs[1:])]
    return " > ".join(segs)


def _row_plan(bars):
    """Ordered display rows: list of {key, name, kind} where kind is
    'plain', 'header' (a parent with children) or 'child'."""
    order = []          # top-level keys in first appearance order
    children = {}       # parent -> [child keys in order]
    names = {}          # key -> display name
    for bar in bars:
        key = _canon_row(bar.get("row", ""))
        if " > " in key and ">" in bar.get("row", ""):
            parent, child = key.split(" > ", 1)
            if parent not in names:
                names[parent] = parent
                order.append(parent)
            kids = children.setdefault(parent, [])
            if key not in names:
                names[key] = child
                kids.append(key)
        elif key not in names:
            names[key] = key
            order.append(key)
    rows = []
    for key in order:
        kids = children.get(key)
        rows.append({"key": key, "name": names[key],
                     "kind": "header" if kids else "plain"})
        for kid in kids or []:
            rows.append({"key": kid, "name": names[kid], "kind": "child"})
    return rows


# ── label columns (Responsible / Remark) ─────────────────────────────────────
def _label_columns(bars):
    """[(header, {row_key: value})] for the Bar kv attrs that appear."""
    cols = []
    for attr, header in (("responsible", "RESPONSIBLE"), ("remark", "REMARK")):
        if not any(b.get(attr) for b in bars):
            continue
        values = {}
        for b in bars:
            v = b.get(attr)
            if not v:
                continue
            key = _canon_row(b.get("row", ""))
            if key in values:
                if values[key] != v:
                    warn(f"gantt: conflicting {header.capitalize()}= values "
                         f"for row {key!r} — keeping {values[key]!r}")
            else:
                values[key] = v
        cols.append((header, values))
    return cols[:MAX_LABEL_COLS]


def _parse_period_range(spec, n_periods):
    """'3-4' -> (3, 4); '3' -> (3, 3); malformed/reversed/out-of-range -> None."""
    m = _RANGE_RX.match(str(spec or "").strip())
    if not m:
        return None
    a = int(m.group(1))
    b = int(m.group(2)) if m.group(2) else a
    if not 1 <= a <= b <= n_periods:
        return None
    return a, b


# ── the builder ──────────────────────────────────────────────────────────────
def build_gantt_slide(prs, p, pal, ctx):
    """Workstream rows x period columns with phase bars and milestone diamonds.
    Syntax: **Periods:** Q1, Q2... + '- Bar: Row=".." Label=".." Start="1"
    End="2"' + '- Milestone: Row=".." Label=".." At="2.5"'. Cycle 7 adds
    Span/Shade/Line rows, Responsible/Remark columns, `Parent > Child` rows
    and quarters/months period expansion (see the module docstring)."""
    try:
        periods, year_spans = expand_periods(p.get("periods", []))
    except ValueError as exc:
        warn(f"gantt: {exc}")
        periods, year_spans = [], []
    bars = p.get("bars", [])
    if not periods or not bars:
        warn("gantt needs **Periods:** and '- Bar:' rows")
        return B.build_bullet_slide(prs, p, pal, ctx)
    slide = B._blank_slide(prs, pal, pal["bg"])
    B._heading(slide, p, pal)
    body_start = len(slide.shapes)

    rows = _row_plan(bars)
    row_index = {r["key"]: i for i, r in enumerate(rows)}
    label_cols = _label_columns(bars)
    lines = p.get("gantt_lines", [])
    spans = p.get("spans", [])

    label_w, grid_x = 2.3, 3.1
    grid_right = 12.6 - LABEL_COL_W * len(label_cols)
    grid_w = grid_right - grid_x
    col_w = grid_w / len(periods)
    line_band = any(ln.get("label") for ln in lines)
    extra_bands = (1 if year_spans else 0) + (1 if line_band else 0)
    top = GRID_TOP + BAND_H * extra_bands
    row_h = min(0.85, (ROW_BUDGET - BAND_H * extra_bands) / max(len(rows), 1))
    grid_h = row_h * len(rows)

    # shaded period bands — drawn first so bars and separators sit on top
    for sh in p.get("shades", []):
        rng = _parse_period_range(sh.get("periods"), len(periods))
        if rng is None:
            warn(f"gantt: Shade Periods={sh.get('periods', '')!r} not a "
                 f"valid range within 1-{len(periods)} — skipped")
            continue
        a, b = rng
        sx = grid_x + (a - 1) * col_w
        sw = (b - a + 1) * col_w
        band = B.add_rect(slide, sx, top, sw, grid_h, pal["surface"])
        set_fill_alpha(band, SHADE_ALPHA)
        if sh.get("label"):
            B.add_tb(slide, sh["label"], sx, top + 0.04, sw, 0.26, size=10,
                     color=pal["text_muted"], align=PP_ALIGN.CENTER,
                     font=pal["font_label"])

    # spanning year-header row above the period labels
    if year_spans:
        yy = top - 0.42 - BAND_H
        for label, i0, count in year_spans:
            yx = grid_x + i0 * col_w
            yw = count * col_w
            B.add_tb(slide, label, yx, yy, yw, 0.3, size=11, bold=True,
                     color=pal["text_muted"], align=PP_ALIGN.CENTER,
                     font=pal["font_label"])
            B.add_rect(slide, yx + 0.05, yy + 0.3, yw - 0.1, 0.015,
                       pal["surface"])

    for i, period in enumerate(periods):
        B.add_tb(slide, period, grid_x + i * col_w, top - 0.42, col_w, 0.34,
                 size=12, bold=True, color=pal["text_muted"],
                 align=PP_ALIGN.CENTER, font=pal["font_label"])
        if i:
            B.add_rect(slide, grid_x + i * col_w, top, 0.015, grid_h,
                       pal["surface"])
    accents = [pal["accent1"], pal["accent2"], pal["accent3"]]
    label_x = 0.8 + (SPAN_SHIFT if spans else 0)
    for r, row in enumerate(rows):
        y = top + r * row_h
        if r:
            B.add_rect(slide, 0.7, y, 11.9, 0.015, pal["surface"])
        indent = (label_x - 0.8) + (CHILD_INDENT if row["kind"] == "child"
                                    else 0)
        B.add_tb(slide, row["name"], 0.8 + indent, y + row_h / 2 - 0.17,
                 label_w - indent, 0.38,
                 size=12 if row["kind"] == "child" else 13,
                 bold=row["kind"] != "child", color=pal["text"],
                 font=pal["font_body"])

    # Responsible / Remark label columns right of the (shrunk) grid
    for ci, (header, values) in enumerate(label_cols):
        cx = grid_right + ci * LABEL_COL_W
        B.add_tb(slide, header, cx, top - 0.42, LABEL_COL_W, 0.34, size=11,
                 bold=True, color=pal["text_muted"], align=PP_ALIGN.CENTER,
                 font=pal["font_label"])
        for r, row in enumerate(rows):
            value = values.get(row["key"])
            if value:
                B.add_tb(slide, value, cx + 0.05,
                         top + r * row_h + row_h / 2 - 0.15,
                         LABEL_COL_W - 0.1, 0.3, size=11,
                         color=pal["text_muted"], align=PP_ALIGN.CENTER,
                         font=pal["font_label"])

    _draw_spans(slide, spans, rows, row_index, pal, top, row_h)

    for bar in bars:
        try:
            start = float(bar.get("start", 1))
            end = float(bar.get("end", start))
        except ValueError:
            continue
        r = row_index[_canon_row(bar.get("row", ""))]
        fill = accents[r % len(accents)]
        bx = grid_x + (start - 1) * col_w + 0.04
        bw = max((end - start + 1) * col_w - 0.08, 0.25)
        y = top + r * row_h + row_h / 2 - 0.21
        B.add_rect(slide, bx, y, bw, 0.42, fill)
        B.add_tb(slide, bar.get("label", ""), bx + 0.08, y + 0.04, bw - 0.16,
                 0.34, size=14, bold=True, color=_label_color_on(fill, pal),
                 font=pal["font_body"])
    for ms in p.get("milestones", []):
        try:
            at = float(ms.get("at", 1))
        except ValueError:
            continue
        key = _canon_row(ms.get("row", ""))
        if key not in row_index:
            warn(f"gantt: Milestone row {ms.get('row', '')!r} unknown — "
                 "skipped")
            continue
        r = row_index[key]
        cx = grid_x + (at - 0.5) * col_w  # At=N means the middle of period N
        cy = top + r * row_h + row_h / 2
        # if the milestone falls within a bar's span on its row, lift it above
        # the bar so it doesn't cover the bar label
        overlaps = any(
            b.get("row") == ms.get("row")
            and float(b.get("start", 1)) - 0.5 <= at <= float(b.get("end", 1)) + 0.5
            for b in bars if str(b.get("start", "")).replace(".", "").isdigit())
        if overlaps:
            cy = top + r * row_h + 0.13
        sl, st, sw, _ = B._sc(cx - 0.13, cy - 0.13, 0.26, 0.26)
        d = slide.shapes.add_shape(4, Inches(sl), Inches(st), Inches(sw), Inches(sw))
        d.fill.solid()
        from palettes import hex_rgb
        d.fill.fore_color.rgb = hex_rgb(pal["text"])
        d.line.fill.background()
        if ms.get("label"):
            # label sits just above the row, clear of bars and the diamond
            B.add_tb(slide, ms["label"], cx - 0.75,
                     top + r * row_h - 0.30, 1.5, 0.3, size=11, bold=True,
                     color=pal["text"], align=PP_ALIGN.CENTER,
                     font=pal["font_label"])

    _draw_lines(slide, lines, pal, grid_x, col_w, len(periods), top, grid_h,
                bool(year_spans))
    shape_groups.group_shapes(slide, list(slide.shapes)[body_start:])
    return slide


def _draw_spans(slide, spans, rows, row_index, pal, top, row_h):
    """Left square brackets spanning named rows, rotated label alongside.
    One nesting level in v1: overlapping spans warn and are skipped."""
    display = {r["name"].casefold(): i for i, r in enumerate(rows)}
    taken = []
    for sp in spans:
        names = [s.strip() for s in sp.get("rows", "").split(",") if s.strip()]
        if not names:
            warn(f"gantt: Span needs Rows=\"a, b\": {sp} — skipped")
            continue
        idxs = []
        for name in names:
            i = row_index.get(name, display.get(name.casefold()))
            if i is None:
                idxs = None
                warn(f"gantt: Span row {name!r} unknown — span skipped")
                break
            idxs.append(i)
        if not idxs:
            continue
        r0, r1 = min(idxs), max(idxs)
        if any(a <= r1 and r0 <= b for a, b in taken):
            warn(f"gantt: Span {sp.get('label', '')!r} overlaps an earlier "
                 "span — skipped (one nesting level in v1)")
            continue
        taken.append((r0, r1))
        y0 = top + r0 * row_h + 0.06
        y1 = top + (r1 + 1) * row_h - 0.06
        spine_x = 1.02
        B.add_rect(slide, spine_x, y0, 0.022, y1 - y0, pal["text_muted"])
        B.add_rect(slide, spine_x, y0, 0.12, 0.022, pal["text_muted"])
        B.add_rect(slide, spine_x, y1 - 0.022, 0.12, 0.022, pal["text_muted"])
        if sp.get("label"):
            # rotated 270° about its center: the pre-rotation frame must stay
            # on-canvas (qa_check reads the unrotated box), so cap its width
            lw = min(y1 - y0, 1.6)
            cx, cy = 0.85, (y0 + y1) / 2
            tb = B.add_tb(slide, sp["label"], cx - lw / 2, cy - 0.15,
                          lw, 0.3, size=11, bold=True,
                          color=pal["text_muted"], align=PP_ALIGN.CENTER,
                          font=pal["font_label"])
            tb.rotation = 270


def _draw_lines(slide, lines, pal, grid_x, col_w, n_periods, top, grid_h,
                has_year_band):
    """Dashed vertical milestone lines across the grid, labels above it."""
    from pptx.enum.dml import MSO_LINE_DASH_STYLE
    from palettes import hex_rgb
    label_y = top - 0.42 - BAND_H * (1 + (1 if has_year_band else 0))
    for ln in lines:
        try:
            at = float(ln.get("at", ""))
        except ValueError:
            warn(f"gantt: Line At={ln.get('at', '')!r} is not a number — "
                 "skipped")
            continue
        if not 0.5 <= at <= n_periods + 0.5:
            warn(f"gantt: Line At={at:g} outside periods 1-{n_periods} — "
                 "skipped")
            continue
        x = grid_x + (at - 0.5) * col_w  # At=N means the middle of period N
        sl, st, _, sh = B._sc(x, top, 0, grid_h)
        conn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,
                                          Inches(sl), Inches(st),
                                          Inches(sl), Inches(st + sh))
        conn.line.color.rgb = hex_rgb(pal["accent3"])
        conn.line.width = Pt(1.75)
        conn.line.dash_style = MSO_LINE_DASH_STYLE.DASH
        if ln.get("label"):
            label_left = max(0.7, min(x - 0.7, 12.63 - 1.4))
            B.add_tb(slide, ln["label"], label_left, label_y, 1.4, 0.3,
                     size=11, bold=True, color=pal["text_muted"],
                     align=PP_ALIGN.CENTER, font=pal["font_label"])


LAYOUTS = {"gantt": build_gantt_slide}
