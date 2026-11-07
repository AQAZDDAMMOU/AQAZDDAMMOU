"""chart_options.py — the Cycle 7 chart option set (spec §2).

Hooked from `chart_visual.place_chart()` at the contract's fixed steps:

  step 1  `shape_data(p, kind)` — `- Sort:` / `- Ordinal:` reorder/relabel the
          data rows; `- Target:` / `- Error-Bars:` route series out of the
          plotted set (their values survive on private `_target` / `_errbars`
          keys of the *copied* slide dict — the original is never mutated,
          and with none of the new keys present the dict is returned as-is).
  step 5  `apply_native(chart, p, pal, kind)` — native, PowerPoint-editable
          XML: `- Gap-Width:`, `- Label-Rotation:`, `- Label-Skip:`, and the
          custom `c:errBars` whiskers.
  step 6  `resolve_value_line(p, kind)` — the mean|median|wmean pre-resolver
          in front of `charts.add_value_line`; `draw_annotations(...)` —
          shape-drawn marks against the final axis range: target ticks,
          `- Total-Line:`, `- Units:`, `- Category-Break:`, and the
          per-series / multi-span `- CAGR:` variants.

Geometry uses the `charts._plot_box` approximation on the already-scaled box
`place_chart` passes through — the same "confirm placement in visual QA"
caveat as the CAGR arrow.  Unknown names/values warn and skip, never crash.
Nothing here imports build_deck.
"""
import re

from lxml import etree
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.shapes import MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

from charts import _nice_ceil, _plot_box
from helpers import warn
from numfmt import fmt_pct
from palettes import hex_rgb

SORT_KINDS = ("bar", "column", "hbar")
GAP_KINDS = ("bar", "column", "hbar", "stacked")
ERRBAR_KINDS = ("bar", "column", "line")
TARGET_KINDS = ("bar", "column")
TARGET_TICK_PT = 2.5
TARGET_TICK_FRACTION = 0.6   # of the category slot width
GAP_WIDTH_MAX = 500
LABEL_ROTATION_MAX = 90
STAT_LINE_LABELS = {"mean": "Mean", "median": "Median",
                    "wmean": "Weighted mean"}


def switch_on(value):
    return str(value or "").strip().lower() in ("on", "yes", "true")


def _names(p):
    return [n.strip() for n in str(p.get("series", "")).split(",")
            if n.strip()]


def _norm(name):
    """Casefolded name with any `1. ` ordinal prefix stripped, so specs match
    categories both before and after `- Ordinal: on` relabels them."""
    return re.sub(r"^\d+\.\s+", "", str(name)).strip().casefold()


def find_name(seq, name):
    """Index of `name` in `seq` (casefold, ordinal-prefix-insensitive), or
    None."""
    target = _norm(name)
    for i, item in enumerate(seq):
        if _norm(item) == target:
            return i
    return None


def _row_values(row):
    vals = row[1]
    return list(vals) if isinstance(vals, list) else [vals]


def _first_value(row):
    vals = _row_values(row)
    return vals[0] if vals else None


# ── step 1: data shaping ────────────────────────────────────────────────────
def shape_data(p, kind):
    """Apply `- Sort:` / `- Ordinal:` / `- Target:` / `- Error-Bars:` to the
    slide dict before the chart is built.  Returns `p` untouched when none of
    the keys is active (byte-identical builds), else a shallow copy with new
    `series` / `data` and private `_target` / `_errbars` routing keys."""
    sort = str(p.get("sort", "")).strip().lower()
    active = (sort in ("desc", "asc") or switch_on(p.get("ordinal"))
              or p.get("target") or p.get("error_bars"))
    if not active or not p.get("data"):
        return p
    p = dict(p)
    p["data"] = [(lbl, list(v) if isinstance(v, list) else v)
                 for lbl, v in p["data"]]
    _route_target(p, kind)
    _route_error_bars(p, kind)
    _apply_sort(p, kind, sort)
    _apply_ordinal(p, kind)
    return p


def _drop_series(p, names, idx):
    """Remove series `idx` from p['series']/p['data']; returns its values in
    data-row order."""
    removed = []
    new_data = []
    for lbl, vals in p["data"]:
        row = _row_values((lbl, vals))
        removed.append(row[idx] if idx < len(row) else None)
        rest = row[:idx] + row[idx + 1:]
        new_data.append((lbl, rest if len(rest) > 1 else rest[0]))
    p["data"] = new_data
    p["series"] = ", ".join(names[:idx] + names[idx + 1:])
    return removed


def _route_target(p, kind):
    spec = str(p.get("target", "")).strip()
    if not spec:
        return
    names = _names(p)
    if kind not in TARGET_KINDS:
        warn(f"Target only applies to chart:bar/column (chart:{kind}) — "
             "plot unchanged")
        return
    idx = find_name(names, spec)
    if idx is None or len(names) < 2:
        warn(f"Target series {spec!r} not found in **Series:** "
             f"{p.get('series')!r} — plot unchanged")
        return
    values = _drop_series(p, names, idx)
    p["_target"] = (names[idx], values)


def _route_error_bars(p, kind):
    spec = str(p.get("error_bars", "")).strip()
    if not spec:
        return
    parts = [s.strip() for s in spec.split(",") if s.strip()]
    names = _names(p)
    if kind not in ERRBAR_KINDS:
        warn(f"Error-Bars only apply to chart:bar/column/line (chart:{kind})"
             " — plot unchanged")
        return
    if len(parts) != 2:
        warn(f"Error-Bars needs exactly two series names "
             f"('- Error-Bars: Low, High'), got {spec!r} — plot unchanged")
        return
    low_i, high_i = find_name(names, parts[0]), find_name(names, parts[1])
    if low_i is None or high_i is None or low_i == high_i or len(names) < 3:
        warn(f"Error-Bars series {spec!r} must name two distinct **Series:** "
             "names and leave one to plot — plot unchanged")
        return
    lows = [_row_values(r)[low_i] for r in p["data"]]
    highs = [_row_values(r)[high_i] for r in p["data"]]
    for idx in sorted((low_i, high_i), reverse=True):  # high index first
        _drop_series(p, _names(p), idx)
    firsts = [_first_value(r) for r in p["data"]]
    plus, minus = [], []
    for v, lo, hi in zip(firsts, lows, highs):
        v, lo, hi = v or 0, lo or 0, hi or 0
        d_plus, d_minus = hi - v, v - lo
        if d_plus < 0 or d_minus < 0:
            warn(f"Error-Bars value outside [low, high] "
                 f"({lo:g} <= {v:g} <= {hi:g} fails) — delta clamped to 0")
        plus.append(max(d_plus, 0.0))
        minus.append(max(d_minus, 0.0))
    p["_errbars"] = (plus, minus)


def _apply_sort(p, kind, sort):
    if sort not in ("desc", "asc"):
        if sort and sort not in ("", "off", "no", "false"):
            warn(f"Sort {sort!r} not desc|asc — data order unchanged")
        return
    if kind not in SORT_KINDS:
        warn(f"Sort only applies to chart:bar/column/hbar (chart:{kind}) — "
             "data order unchanged")
        return
    order = sorted(range(len(p["data"])),
                   key=lambda i: (_first_value(p["data"][i])
                                  if _first_value(p["data"][i]) is not None
                                  else float("-inf")),
                   reverse=(sort == "desc"))
    p["data"] = [p["data"][i] for i in order]
    if "_target" in p:
        name, vals = p["_target"]
        p["_target"] = (name, [vals[i] for i in order])
    if "_errbars" in p:
        plus, minus = p["_errbars"]
        p["_errbars"] = ([plus[i] for i in order], [minus[i] for i in order])


def _apply_ordinal(p, kind):
    if not switch_on(p.get("ordinal")):
        return
    if kind not in SORT_KINDS:
        warn(f"Ordinal only applies to chart:bar/column/hbar (chart:{kind}) "
             "— labels unchanged")
        return
    p["data"] = [(f"{i}. {lbl}", vals)
                 for i, (lbl, vals) in enumerate(p["data"], 1)]


# ── step 5: native XML options ──────────────────────────────────────────────
def apply_native(chart, p, pal, kind):
    """Gap width, category-label rotation/skip, and custom error bars —
    all native chart XML, editable in PowerPoint."""
    _apply_gap_width(chart, p, kind)
    _apply_label_rotation(chart, p)
    _apply_label_skip(chart, p)
    if p.get("_errbars"):
        _apply_error_bars(chart, p["_errbars"], pal)


def _apply_gap_width(chart, p, kind):
    raw = str(p.get("gap_width", "")).strip()
    if not raw:
        return
    if kind not in GAP_KINDS:
        warn(f"Gap-Width only applies to chart:bar/column/hbar/stacked "
             f"(chart:{kind}) — ignored")
        return
    try:
        value = int(float(raw))
    except ValueError:
        warn(f"Gap-Width not numeric: {raw!r} — ignored")
        return
    chart.plots[0].gap_width = min(max(value, 0), GAP_WIDTH_MAX)


def _cat_ax(chart):
    try:
        return chart.category_axis._element
    except ValueError:
        return None


def _apply_label_rotation(chart, p):
    raw = str(p.get("label_rotation", "")).strip()
    if not raw:
        return
    try:
        deg = float(raw)
    except ValueError:
        warn(f"Label-Rotation not numeric: {raw!r} — ignored")
        return
    deg = min(max(deg, -LABEL_ROTATION_MAX), LABEL_ROTATION_MAX)
    cat_ax = _cat_ax(chart)
    if cat_ax is None:
        warn("Label-Rotation: chart has no category axis — ignored")
        return
    txPr = cat_ax.find(qn("c:txPr"))
    if txPr is None:  # _style_axes normally created it via tick_labels.font
        warn("Label-Rotation: category axis has no text properties — ignored")
        return
    bodyPr = txPr.find(qn("a:bodyPr"))
    if bodyPr is None:
        bodyPr = etree.SubElement(txPr, qn("a:bodyPr"))
        txPr.insert(0, bodyPr)
    bodyPr.set("rot", str(int(round(deg * 60000))))


def _apply_label_skip(chart, p):
    raw = str(p.get("label_skip", "")).strip()
    if not raw:
        return
    try:
        skip = int(raw)
        if skip < 1:
            raise ValueError
    except ValueError:
        warn(f"Label-Skip must be a positive integer, got {raw!r} — ignored")
        return
    cat_ax = _cat_ax(chart)
    if cat_ax is None:
        warn("Label-Skip: chart has no category axis — ignored")
        return
    el = cat_ax.makeelement(qn("c:tickLblSkip"), {"val": str(skip)})
    anchor = cat_ax.find(qn("c:noMultiLvlLbl"))
    if anchor is not None:
        anchor.addprevious(el)
    else:
        cat_ax.append(el)


def _num_lit(tag, deltas):
    lit = etree.SubElement(tag, qn("c:numLit"))
    etree.SubElement(lit, qn("c:formatCode")).text = "General"
    etree.SubElement(lit, qn("c:ptCount")).set("val", str(len(deltas)))
    for i, d in enumerate(deltas):
        pt = etree.SubElement(lit, qn("c:pt"))
        pt.set("idx", str(i))
        etree.SubElement(pt, qn("c:v")).text = f"{d:g}"


def _apply_error_bars(chart, deltas, pal):
    """Custom min/max whiskers (`c:errBars`, errValType cust) on the first
    plotted series — the native, PowerPoint-editable representation."""
    plus, minus = deltas
    ser = chart.plots[0].series[0]._element
    err = etree.SubElement(ser, qn("c:errBars"))
    etree.SubElement(err, qn("c:errDir")).set("val", "y")
    etree.SubElement(err, qn("c:errBarType")).set("val", "both")
    etree.SubElement(err, qn("c:errValType")).set("val", "cust")
    etree.SubElement(err, qn("c:noEndCap")).set("val", "0")
    for tag, vals in ((qn("c:plus"), plus), (qn("c:minus"), minus)):
        _num_lit(etree.SubElement(err, tag), vals)
    spPr = etree.fromstring(
        '<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/'
        'chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/'
        f'main"><a:ln w="12700"><a:solidFill><a:srgbClr val="'
        f'{pal["text_muted"]}"/></a:solidFill></a:ln></c:spPr>')
    err.append(spPr)
    # schema order: errBars sits before c:cat inside c:ser
    cat = ser.find(qn("c:cat"))
    if cat is not None:
        cat.addprevious(err)


# ── step 6: value-line pre-resolver ─────────────────────────────────────────
def resolve_value_line(p, kind):
    """Resolve `- Value-Line: mean|median|wmean ["label"]` into the numeric
    '<label>, <value>' spec `charts.add_value_line` understands.  A spec whose
    first token is not a stat keyword is returned unchanged (the numeric
    `- Value-Line: Target, 40` path keeps today's behavior byte-for-byte).
    Returns None when the stat cannot be computed (warned)."""
    spec = str(p.get("value_line", "")).strip()
    # stat form is ONLY the bare keyword or keyword + quoted label (the
    # parser may have eaten the closing quote); anything else — including
    # 'Median EBITDA, 40' — is a numeric spec and passes through untouched
    m = re.match(r'^(mean|median|wmean)\s*(?:"([^"]*)"?)?$', spec,
                 re.IGNORECASE)
    if not m:
        return spec
    stat = m.group(1).lower()
    label = (m.group(2) or "").strip() or STAT_LINE_LABELS[stat]
    rows = p.get("data") or []
    values = [_first_value(r) for r in rows]
    values = [v for v in values if v is not None]
    if not values:
        warn(f"Value-Line {stat}: no numeric data — line skipped")
        return None
    if stat == "mean":
        value = sum(values) / len(values)
    elif stat == "median":
        ordered = sorted(values)
        mid = len(ordered) // 2
        value = (ordered[mid] if len(ordered) % 2
                 else (ordered[mid - 1] + ordered[mid]) / 2)
    else:  # wmean: series 1 weighted by series 2
        names = _names(p)
        if len(names) != 2:
            warn(f"Value-Line wmean needs exactly 2 series (values, weights) "
                 f"— got {len(names)} — line skipped")
            return None
        pairs = [(_row_values(r) + [None, None])[:2] for r in rows]
        pairs = [(v, w) for v, w in pairs if v is not None and w is not None]
        total_w = sum(w for _, w in pairs)
        if not pairs or not total_w:
            warn("Value-Line wmean: weights sum to zero — line skipped")
            return None
        value = sum(v * w for v, w in pairs) / total_w
    return f"{label}, {value:.10g}"


# ── step 6: shape-drawn annotations ─────────────────────────────────────────
def draw_annotations(slide, chart, p, pal, box, kind):
    """Target ticks, 100% line, units label, category breaks, and the
    per-series / multi-span CAGR variants — drawn against the final axis."""
    if p.get("_target"):
        _draw_target_ticks(slide, chart, p, pal, box)
    if switch_on(p.get("total_line")):
        if kind == "stacked-100":
            px, py, pw, _ = _plot_box(*box)
            draw_total_line(slide, pal, px, py, pw)
        else:
            warn(f"Total-Line only applies to chart:stacked-100 and mekko "
                 f"(chart:{kind}) — ignored")
    if p.get("units"):
        _draw_units(slide, p, pal, box, kind)
    for spec in p.get("category_breaks") or []:
        _draw_category_break(slide, spec, p, pal, box)
    mode = cagr_mode(p.get("cagr"), [d[0] for d in p.get("data") or []])
    if mode == "series":
        _draw_series_cagr(slide, chart, p, pal, box, kind)
    elif mode == "spans":
        _draw_span_cagr(slide, chart, p, pal, box, kind)


def _axis_max(chart, fallback_values, headroom=1.05):
    """The chart's value-axis maximum, setting a nice one (never shrinking an
    existing one) when no earlier step pinned it."""
    va = chart.value_axis
    if va.maximum_scale is not None:
        return va.maximum_scale
    data_max = max((v for v in fallback_values if v is not None), default=0)
    axis_max = _nice_ceil(data_max * headroom) if data_max > 0 else 1
    va.maximum_scale = float(axis_max)
    va.minimum_scale = 0.0
    return axis_max


def _tb(slide, text, x, y, w, h, pal, size=11, bold=False, color=None,
        align=PP_ALIGN.LEFT, font_key="font_label"):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    para = tb.text_frame.paragraphs[0]
    para.alignment = align
    run = para.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = hex_rgb(color or pal["text_muted"])
    run.font.name = pal[font_key]
    return tb


def _connector(slide, pal, x0, y0, x1, y1, color, width_pt, dash=None):
    conn = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT, Inches(x0), Inches(y0), Inches(x1), Inches(y1))
    conn.line.color.rgb = hex_rgb(color)
    conn.line.width = Pt(width_pt)
    if dash is not None:
        conn.line.dash_style = dash
    return conn


def _draw_target_ticks(slide, chart, p, pal, box):
    """Short thick accent3 ticks at the routed Target series' value per
    category, plus a one-line legend chip under the plot."""
    name, values = p["_target"]
    px, py, pw, ph = _plot_box(*box)
    plotted = [v for s in chart.plots[0].series for v in s.values
               if v is not None]
    axis_max = _axis_max(chart, plotted + [v for v in values if v is not None])
    n = len(values) or 1
    slot = pw / n
    tick_w = slot * TARGET_TICK_FRACTION
    for i, v in enumerate(values):
        if v is None:
            continue
        y = py + (1 - v / axis_max) * ph
        x0 = px + slot * (i + 0.5) - tick_w / 2
        _connector(slide, pal, x0, y, x0 + tick_w, y, pal["accent3"],
                   TARGET_TICK_PT)
    _tb(slide, f"— {name}", px, py + ph + 0.06, min(pw, 3.0), 0.25, pal,
        size=10, color=pal["accent3"])


def draw_total_line(slide, pal, x, y, w):
    """Thin muted 100% line + label — shared by chart:stacked-100 and the
    mekko builder (which passes its own scaled column-area geometry)."""
    _connector(slide, pal, x, y, x + w, y, pal["text_muted"], 1.0)
    _tb(slide, "100%", x + w + 0.06, y - 0.13, 0.6, 0.26, pal,
        size=10, bold=True)


def _draw_units(slide, p, pal, box, kind):
    from chart_visual import VERTICAL_KINDS
    if kind not in VERTICAL_KINDS:
        warn(f"Units label only applies to vertical chart kinds "
             f"(chart:{kind}) — ignored")
        return
    px, py, _, _ = _plot_box(*box)
    _tb(slide, str(p["units"]).strip(), px, py - 0.30, 1.8, 0.26, pal,
        size=11)


def _draw_category_break(slide, spec, p, pal, box):
    """Dashed vertical separator after a category, optional 'left | right'
    label split at the top of the plot box."""
    # the parser's outer strip('"') may have eaten the closing quote, so
    # tolerate `FY25 "Actual | Forecast` as well as `FY25 "Actual | Forecast"`
    text = str(spec).strip()
    cat_name, quote, label = text.partition('"')
    cat_name = cat_name.strip()
    label = label.rstrip('"').strip() if quote else ""
    cats = [d[0] for d in p.get("data") or []]
    idx = find_name(cats, cat_name)
    if idx is None:
        warn(f"Category-Break category {cat_name!r} not found — skipped")
        return
    px, py, pw, ph = _plot_box(*box)
    n = len(cats) or 1
    x = px + pw * ((idx + 1) / n)   # boundary after the category's slot
    _connector(slide, pal, x, py, x, py + ph, pal["text_muted"], 1.25,
               dash=MSO_LINE_DASH_STYLE.DASH)
    if not label:
        return
    left_txt, _, right_txt = (s.strip() for s in label.partition("|"))
    if left_txt:
        _tb(slide, left_txt, x - 1.65, py + 0.02, 1.57, 0.24, pal, size=10,
            align=PP_ALIGN.RIGHT)
    if right_txt:
        _tb(slide, right_txt, x + 0.08, py + 0.02, 1.57, 0.24, pal, size=10)


# ── step 6: CAGR variants ───────────────────────────────────────────────────
def cagr_mode(value, cats=None):
    """Route a `- CAGR:` value: 'series' -> per-series labels, a comma list
    of `CatA-CatB` spans -> span arrows, anything else (incl. 'on') -> the
    legacy single arrow, byte-identical to today.  With `cats` (the slide's
    category labels), a span-shaped value routes to 'spans' only when at
    least one token names two known categories — so legacy values that
    merely contain a hyphen ('5-yr', 'FY22-FY26' on other labels) keep
    their arrow instead of silently drawing nothing."""
    text = str(value or "").strip()
    if not text:
        return "legacy"
    if text.lower() == "series":
        return "series"
    if re.fullmatch(r"[^,]+-[^,]+(\s*,\s*[^,]+-[^,]+)*", text):
        if cats is None:
            return "spans"
        return "spans" if any(
            _split_span(t.strip(), list(cats)) is not None
            for t in text.split(",")) else "legacy"
    return "legacy"


def _cagr_text(first, last, periods):
    cagr = (last / first) ** (1 / periods) - 1
    return f"{fmt_pct(cagr * 100, signed=True, legacy_decimals=1)} p.a."


def _draw_series_cagr(slide, chart, p, pal, box, kind):
    """`- CAGR: series` — each series' `<name> +8.2% p.a.` label at its right
    end, in the series color (multi-series line/bar/column)."""
    series_list = list(chart.plots[0].series)
    if kind not in ("line", "bar", "column") or len(series_list) < 2:
        warn("CAGR: series needs a multi-series chart:line/column — skipped")
        return
    px, py, pw, ph = _plot_box(*box)
    plotted = [v for s in series_list for v in s.values if v is not None]
    axis_max = _axis_max(chart, plotted)
    accents = pal.get("chart_series") or [pal["accent1"], pal["accent2"],
                                          pal["accent3"]]
    for i, s in enumerate(series_list):
        values = [v for v in s.values if v is not None]
        if len(values) < 2 or values[0] <= 0 or values[-1] <= 0:
            warn(f"CAGR: series {s.name!r} needs 2+ positive values — "
                 "label skipped")
            continue
        y = py + (1 - values[-1] / axis_max) * ph - 0.12
        _tb(slide, f"{s.name} {_cagr_text(values[0], values[-1], len(values) - 1)}",
            px + pw + 0.06, y, 2.3, 0.26, pal, size=11, bold=True,
            color=accents[i % len(accents)], font_key="font_body")


def _split_span(token, cats):
    """Split 'FY22-FY24' on the '-' that makes both halves known categories
    (category names may themselves contain hyphens)."""
    for m in re.finditer("-", token):
        a, b = token[:m.start()].strip(), token[m.end():].strip()
        ia, ib = find_name(cats, a), find_name(cats, b)
        if ia is not None and ib is not None:
            return ia, ib
    return None


def _draw_span_cagr(slide, chart, p, pal, box, kind):
    """`- CAGR: FY22-FY24, FY24-FY27` — one arrow + `+12.4% p.a.` label per
    span between the named categories (single-series bar/column/line)."""
    series_list = list(chart.plots[0].series)
    if kind not in ("bar", "column", "line") or len(series_list) != 1:
        warn("CAGR spans need a single-series chart:bar/column/line — skipped")
        return
    values = list(series_list[0].values)
    cats = [d[0] for d in p.get("data") or []]
    axis_max = _axis_max(chart, [v for v in values if v is not None])
    px, py, pw, ph = _plot_box(*box)
    n = len(values) or 1
    for token in str(p["cagr"]).split(","):
        span = _split_span(token.strip(), cats)
        if span is None:
            warn(f"CAGR span {token.strip()!r} does not name two known "
                 "categories — span skipped")
            continue
        ia, ib = span
        va, vb = values[ia], values[ib]
        if ia == ib or va is None or vb is None or va <= 0 or vb <= 0:
            warn(f"CAGR span {token.strip()!r} needs two distinct positive "
                 "values — span skipped")
            continue
        x0 = px + pw * ((ia + 0.5) / n)
        x1 = px + pw * ((ib + 0.5) / n)
        y0 = py + (1 - va / axis_max) * ph - 0.22
        y1 = py + (1 - vb / axis_max) * ph - 0.22
        conn = _connector(slide, pal, x0, y0, x1, y1, pal["accent1"], 2.0)
        from charts import add_arrowhead_to_connector
        add_arrowhead_to_connector(conn)
        _tb(slide, _cagr_text(va, vb, abs(ib - ia)),
            (x0 + x1) / 2 - 1.1, min(y0, y1) - 0.34, 2.2, 0.3, pal,
            size=12, bold=True, color=pal["text"], align=PP_ALIGN.CENTER,
            font_key="font_body")
