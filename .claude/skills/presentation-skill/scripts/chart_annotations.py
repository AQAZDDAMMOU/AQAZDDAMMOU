"""chart_annotations.py — think-cell style difference arrows (Cycle 7 §1).

`- Diff-Arrow:` on a chart slide draws the delta between two bars:

  * total difference — `CatA, CatB [abs|pct|both]` (default pct): a thin
    leader from the top of each bar to a column just right of the plot box,
    a vertical double-headed connector between the leaders, and a bubble
    label centred on the connector.  Single series compares the two bar
    values; multi series compares per-category totals (stacked reading).
  * level difference — `SerA@Cat -> SerB@Cat` (`->` or `→`): a vertical
    double-headed connector directly above the category's cluster between
    the two series' bar tops, bubble label beside it.

Labels: abs -> `fmt_abs(delta, signed=True)`; pct -> `fmt_pct` of the delta
over the *from* value (zero/negative base warns and falls back to abs);
both -> `+12 (+24%)`.  Kinds: bar/column/line/stacked/stacked-100 (stacked-100
warns that plotted totals are constant); hbar warns and skips (horizontal
geometry is not v1); anything else warns and skips.  Unknown category or
series names warn and skip that arrow — never crash.

Geometry is the `charts._plot_box` approximation on the already-scaled box
`chart_visual.place_chart` passes through (same visual-QA caveat as the CAGR
arrow).  Axis handling mirrors `add_cagr_arrow`: an existing maximum_scale is
respected — never shrunk — else a `_nice_ceil(data_max * 1.15)` headroom max
is set.  Nothing here imports build_deck.
"""
import re

from lxml import etree
from pptx.enum.shapes import MSO_CONNECTOR
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

from charts import _nice_ceil, _plot_box
from helpers import warn
from numfmt import fmt_abs, fmt_pct
from palettes import hex_rgb

ARROW_KINDS = ("bar", "column", "line", "stacked", "stacked-100")
MODES = ("abs", "pct", "both")
DEFAULT_MODE = "pct"
HEADROOM = 1.15
LEADER_PT = 1.0
CONNECTOR_PT = 1.75
BUBBLE_H = 0.30
ARROW_COLUMN_GAP = 0.10   # plot-box right edge -> first arrow column
ARROW_COLUMN_STEP = 0.42  # between stacked arrow columns

_LEVEL_RE = re.compile(r"^(.+?)@(.+?)\s*(?:->|→)\s*(.+?)@(.+?)$")


def _norm(name):
    """Casefolded name with any `1. ` ordinal prefix stripped (matches
    categories relabelled by `- Ordinal: on`)."""
    return re.sub(r"^\d+\.\s+", "", str(name)).strip().casefold()


def _find(seq, name):
    target = _norm(name)
    for i, item in enumerate(seq):
        if _norm(item) == target:
            return i
    return None


def parse_arrow(spec):
    """Parse one `- Diff-Arrow:` value.

    Returns ("level", ser_a, ser_b, cat, mode) or
            ("total", cat_a, cat_b, mode) or None (malformed)."""
    text = str(spec).strip()
    m = _LEVEL_RE.match(text)
    if m:
        ser_a, cat_a, ser_b, cat_b = (g.strip().strip('"') for g in m.groups())
        mode = DEFAULT_MODE
        tail = cat_b.rsplit(None, 1)
        if len(tail) == 2 and tail[1].lower() in MODES:
            cat_b, mode = tail[0].strip(), tail[1].lower()
        if _norm(cat_a) != _norm(cat_b):
            return None
        return ("level", ser_a, ser_b, cat_a, mode)
    head, sep, rest = text.partition(",")
    if not sep:
        return None
    cat_a, cat_b, mode = head.strip(), rest.strip(), DEFAULT_MODE
    tail = cat_b.rsplit(None, 1)
    if len(tail) == 2 and tail[1].lower() in MODES:
        cat_b, mode = tail[0].strip(), tail[1].lower()
    if not cat_a or not cat_b:
        return None
    return ("total", cat_a.strip('"'), cat_b.strip('"'), mode)


def _delta_label(from_v, to_v, mode, where):
    delta = to_v - from_v
    if mode in ("pct", "both") and from_v <= 0:
        warn(f"{where}: percent of a zero/negative base ({from_v:g}) — "
             "falling back to abs")
        mode = "abs"
    if mode == "abs":
        return fmt_abs(delta, signed=True)
    pct = fmt_pct(delta / from_v * 100, signed=True)
    if mode == "pct":
        return pct
    return f"{fmt_abs(delta, signed=True)} ({pct})"


def _double_arrowheads(conn):
    """Triangle headEnd + tailEnd on a connector (double-headed), via the
    same private `_get_or_add_ln` seam as charts.add_arrowhead_to_connector."""
    getter = getattr(conn.line, "_get_or_add_ln", None)
    if getter is None:
        return
    ln = getter()
    for tag in ("a:headEnd", "a:tailEnd"):
        end = etree.SubElement(ln, qn(tag))
        end.set("type", "triangle")
        end.set("w", "med")
        end.set("len", "med")


def _connector(slide, pal, x0, y0, x1, y1, width_pt):
    conn = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT, Inches(x0), Inches(y0), Inches(x1), Inches(y1))
    conn.line.color.rgb = hex_rgb(pal["text_muted"])
    conn.line.width = Pt(width_pt)
    return conn


def _bubble(slide, pal, cx, cy, label):
    """Rounded-rect bubble (surface fill, 11pt bold text) centred on
    (cx, cy)."""
    from pptx.enum.shapes import MSO_SHAPE
    from helpers import set_shape_text
    w = max(0.5, 0.09 * len(label) + 0.22)
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(cx - w / 2),
        Inches(cy - BUBBLE_H / 2), Inches(w), Inches(BUBBLE_H))
    shape.fill.solid()
    shape.fill.fore_color.rgb = hex_rgb(pal["surface"])
    shape.line.fill.background()
    set_shape_text(shape, label, size=11, bold=True, color=pal["text"],
                   font=pal["font_label"])
    return shape


def _category_values(chart, kind, multi):
    """The per-category height a total arrow reads: the value for a single
    series, the per-category total (stacked reading) for multi series."""
    series_values = [list(s.values) for s in chart.plots[0].series]
    if not series_values:
        return []
    if not multi or len(series_values) == 1:
        return [v if v is not None else 0 for v in series_values[0]]
    return [sum(v or 0 for v in col) for col in zip(*series_values)]


def _ensure_axis_max(chart, values):
    """Mirror add_cagr_arrow's axis handling: respect an existing
    maximum_scale (never shrink), else set a nice 1.15-headroom maximum."""
    va = chart.value_axis
    if va.maximum_scale is not None:
        return va.maximum_scale
    data_max = max((v for v in values if v is not None), default=0)
    axis_max = _nice_ceil(data_max * HEADROOM) if data_max > 0 else 1
    va.maximum_scale = float(axis_max)
    va.minimum_scale = 0.0
    return axis_max


def add_diff_arrows(slide, chart, p, pal, box, kind):
    """Draw every `- Diff-Arrow:` on the built chart (see module docstring)."""
    specs = p.get("diff_arrows") or []
    if not specs:
        return
    if kind == "hbar":
        warn("Diff-Arrow is not supported on chart:hbar (horizontal "
             "geometry) — arrows skipped")
        return
    if kind not in ARROW_KINDS:
        warn(f"Diff-Arrow only applies to chart:"
             f"{'/'.join(ARROW_KINDS)} (chart:{kind}) — arrows skipped")
        return
    cats = [d[0] for d in p.get("data") or []]
    names = [n.strip() for n in str(p.get("series", "")).split(",")
             if n.strip()]
    multi = len(names) > 1
    if kind == "stacked-100" and any(
            parse_arrow(s) and parse_arrow(s)[0] == "total" for s in specs):
        warn("Diff-Arrow on chart:stacked-100 compares segment sums — the "
             "plotted column totals are constant (100%)")
    totals = _category_values(chart, kind, multi)
    # stacked-100's value axis is 0..100% — never widen it; every column top
    # sits at the top of the plot box, only the labels carry the delta.
    axis_max = (None if kind == "stacked-100"
                else _ensure_axis_max(chart, totals))
    px, py, pw, ph = _plot_box(*box)
    n = len(cats) or 1
    total_arrow_i = 0
    for spec in specs:
        parsed = parse_arrow(spec)
        if parsed is None:
            warn(f"Diff-Arrow {spec!r} did not parse — expected "
                 "'CatA, CatB [abs|pct|both]' or 'SerA@Cat -> SerB@Cat' — "
                 "arrow skipped")
            continue
        if parsed[0] == "total":
            _, cat_a, cat_b, mode = parsed
            ia, ib = _find(cats, cat_a), _find(cats, cat_b)
            if ia is None or ib is None:
                warn(f"Diff-Arrow category not found in the data rows: "
                     f"{cat_a if ia is None else cat_b!r} — arrow skipped")
                continue
            _draw_total(slide, pal, totals, ia, ib, mode, axis_max,
                        (px, py, pw, ph), n, total_arrow_i, spec)
            total_arrow_i += 1
        else:
            if kind == "stacked-100":
                warn(f"Diff-Arrow {spec!r}: level arrows are not supported "
                     "on chart:stacked-100 — arrow skipped")
                continue
            _, ser_a, ser_b, cat, mode = parsed
            ja, jb = _find(names, ser_a), _find(names, ser_b)
            ic = _find(cats, cat)
            if ja is None or jb is None or ic is None:
                warn(f"Diff-Arrow {spec!r}: unknown series or category "
                     "name — arrow skipped")
                continue
            series_values = [list(s.values) for s in chart.plots[0].series]
            _draw_level(slide, pal, series_values, ja, jb, ic, mode,
                        axis_max, (px, py, pw, ph), n, spec)


def _bar_top_y(value, axis_max, py, ph):
    if axis_max is None:   # stacked-100: every column top is the plot top
        return py
    return py + (1 - (value or 0) / axis_max) * ph


def _draw_total(slide, pal, totals, ia, ib, mode, axis_max, plot, n,
                arrow_i, spec):
    px, py, pw, ph = plot
    v_a, v_b = totals[ia], totals[ib]
    y_a = _bar_top_y(v_a, axis_max, py, ph)
    y_b = _bar_top_y(v_b, axis_max, py, ph)
    x_a = px + pw * ((ia + 0.5) / n)
    x_b = px + pw * ((ib + 0.5) / n)
    x_col = px + pw + ARROW_COLUMN_GAP + arrow_i * ARROW_COLUMN_STEP
    _connector(slide, pal, x_a, y_a, x_col, y_a, LEADER_PT)
    _connector(slide, pal, x_b, y_b, x_col, y_b, LEADER_PT)
    conn = _connector(slide, pal, x_col, y_a, x_col, y_b, CONNECTOR_PT)
    _double_arrowheads(conn)
    label = _delta_label(v_a, v_b, mode, f"Diff-Arrow {spec!r}")
    _bubble(slide, pal, x_col, (y_a + y_b) / 2, label)


def _draw_level(slide, pal, series_values, ja, jb, ic, mode, axis_max,
                plot, n, spec):
    px, py, pw, ph = plot
    v_a = series_values[ja][ic] if ic < len(series_values[ja]) else None
    v_b = series_values[jb][ic] if ic < len(series_values[jb]) else None
    if v_a is None or v_b is None:
        warn(f"Diff-Arrow {spec!r}: missing value at the category — "
             "arrow skipped")
        return
    x = px + pw * ((ic + 0.5) / n)
    y_a = _bar_top_y(v_a, axis_max, py, ph)
    y_b = _bar_top_y(v_b, axis_max, py, ph)
    conn = _connector(slide, pal, x, y_a, x, y_b, CONNECTOR_PT)
    _double_arrowheads(conn)
    label = _delta_label(v_a, v_b, mode, f"Diff-Arrow {spec!r}")
    _bubble(slide, pal, x + 0.55, (y_a + y_b) / 2, label)
