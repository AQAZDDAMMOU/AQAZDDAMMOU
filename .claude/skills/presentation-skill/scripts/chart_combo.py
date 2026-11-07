#!/usr/bin/env python3
"""
chart_combo.py — combo bar+line exhibits and the XY chart extensions
(Cycle 7 §3, think-cell parity).

  * `**Visual:** chart:combo` — clustered columns with the `- Line-Series:`
    names (default: the last `**Series:**` name) rendered as lines: the
    chart is built as one native column chart and the line series' `c:ser`
    elements are moved into a `c:lineChart` group sharing the same axes.
  * `- Secondary: <names>` — a right-hand secondary value axis (delete=0,
    crosses at max) plus a hidden category axis; the named series' chart
    group is rewired to the new axis ids. chart:combo / chart:line only.
  * `- Trend: linear|exp|log|power|poly2..poly6 ["label"]` — a native
    `c:trendline` on the first plotted series, dashed accent3.
  * Scatter/bubble `[Group]` row-label suffixes bucket single-Y rows into
    one series per group (legend on, palette accents).
  * `- Partition: x=50, y=30` — quadrant tint rects drawn *before* the
    chart plus dashed partition lines; the axis ranges are set explicitly
    so the plot-box mapping holds.
  * `chart:bubble` — native bubble chart from 3-value rows (x, y, size);
    `- Size-Legend: on` draws two reference circles.

Shape-drawn parts reuse `charts._plot_box` — the same "confirm placement
in visual QA" caveat as the CAGR/benchmark annotations. Nothing here
imports build_deck.
"""
import re
from copy import deepcopy

from lxml import etree
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.util import Inches, Pt
from pptx.oxml.ns import nsmap, nsuri, qn

from helpers import add_rect, set_fill_alpha, warn
from numfmt import fmt_abs
from palettes import hex_rgb

_NSMAP_CA = nsmap("c", "a")

LINE_SERIES_W_PT = 2.5       # combo line series weight — the chart:line spec
TREND_W_PT = 1.75            # trendline weight, matching the ref-line style
PARTITION_ALPHA = 35         # quadrant tint opacity (percent)
PARTITION_W_PT = 1.75        # dashed partition-line weight
SIZE_LEGEND_R = (0.10, 0.18)  # min/max reference-circle radii (inches)
TREND_TYPES = {"linear": "linear", "exp": "exp", "log": "log",
               "power": "power"}
TREND_KINDS = ("column", "bar", "line", "scatter", "combo")
SECONDARY_KINDS = ("combo", "line")
UNGROUPED = "Other"          # bucket for suffix-less rows on a grouped chart

GROUP_RX = re.compile(r"\s*\[([^\[\]]+)\]\s*$")


# ── shared grammar ──────────────────────────────────────────────────────────
def parse_names(value):
    """A comma-separated name list field, trimmed and empties dropped."""
    return [n.strip() for n in str(value or "").split(",") if n.strip()]


def split_group(label):
    """('Acme [Tier 1]') -> ('Acme', 'Tier 1'); no suffix -> (label, None)."""
    text = str(label)
    m = GROUP_RX.search(text)
    if not m:
        return text.strip(), None
    return text[:m.start()].strip(), m.group(1).strip()


def data_groups(data):
    """None when no row label carries a `[Group]` suffix; else the ordered
    buckets [(group, [(row_index, clean_label, values), ...])] — suffix-less
    rows land in the UNGROUPED bucket."""
    if not any(split_group(lbl)[1] for lbl, _ in data):
        return None
    order, buckets = [], {}
    for i, (lbl, vals) in enumerate(data):
        base, grp = split_group(lbl)
        grp = grp or UNGROUPED
        if grp not in buckets:
            order.append(grp)
            buckets[grp] = []
        buckets[grp].append((i, base, vals))
    return [(g, buckets[g]) for g in order]


def line_series_names(p, names):
    """The `- Line-Series:` names of a combo slide (default: the last
    **Series:** name). Unknown names warn and are skipped; at least one
    series always stays a column."""
    wanted = parse_names(p.get("line_series")) or names[-1:]
    for n in wanted:
        if n not in names:
            warn(f"Line-Series {n!r} is not a **Series:** name — skipped")
    known = [n for n in wanted if n in names]
    if known and len(set(known)) >= len(names):
        warn("Line-Series names every series — keeping the first series "
             "as columns (an all-line exhibit is chart:line)")
        known = [n for n in known if n != names[0]]
    return known


def parse_partition(spec):
    """'- Partition: x=50, y=30' -> {'x': 50.0, 'y': 30.0} (either key may
    be absent). Unparsable spec -> {} with a warning."""
    out = {}
    for m in re.finditer(r"([xy])\s*=\s*(-?[\d.]+(?:,\d{3})*[\d.]*)",
                         str(spec or ""), re.IGNORECASE):
        try:
            out[m.group(1).lower()] = float(m.group(2).replace(",", ""))
        except ValueError:
            continue
    if not out:
        warn(f"Partition not parsable: {spec!r} "
             "(expected '- Partition: x=50, y=30')")
    return out


def parse_trend(spec):
    """'- Trend: poly3 \"Fit\"' -> ('poly', 3, 'Fit'); plain types ->
    (type, None, label). Unknown type warns and falls back to linear.
    (The outline parser may have stripped an outer quote already, so the
    label's quotes are treated as optional.)"""
    parts = str(spec or "").strip().split(None, 1)
    token = parts[0].lower() if parts else ""
    label = parts[1].strip().strip('"') if len(parts) > 1 else None
    label = label or None
    pm = re.match(r"poly([2-6])$", token)
    if pm:
        return "poly", int(pm.group(1)), label
    if token in TREND_TYPES:
        return TREND_TYPES[token], None, label
    warn(f"Trend type {spec!r} is not linear|exp|log|power|poly2..poly6 — "
         "using linear")
    return "linear", None, label


def _ser_name(ser_el):
    """The cached display name of a <c:ser>, or ''."""
    tx = ser_el.find(qn("c:tx"))
    if tx is None:
        return ""
    node = next(tx.iter(qn("c:v")), None)
    return (node.text or "") if node is not None else ""


def _sub(parent, tag, **attrs):
    el = etree.SubElement(parent, qn(tag))
    for k, v in attrs.items():
        el.set(k, str(v))
    return el


def _plot_area(chart):
    return chart._chartSpace.find(qn("c:chart")).find(qn("c:plotArea"))


def _plot_elements(plot_area):
    return [el for el in plot_area
            if isinstance(el.tag, str) and el.tag.endswith("Chart")]


def _accents(pal):
    return pal.get("chart_series") or [pal["accent1"], pal["accent2"],
                                       pal["accent3"]]


# ── combo: clustered columns + line series ──────────────────────────────────
def place_combo(slide, p, pal, cats, names, series, left, top, w, h):
    """Build `chart:combo`: one native column chart with every series, then
    move the line series into a `c:lineChart` group on the same axes."""
    from charts import add_native_chart
    chart = add_native_chart(slide, pal, "combo", cats, series,
                             left, top, w, h)
    if len(names) < 2:
        warn("chart:combo needs 2+ **Series:** names — built as columns")
        return chart
    _move_line_series(chart, pal, line_series_names(p, names), names)
    return chart


def _move_line_series(chart, pal, line_names, names):
    """Move the named `c:ser` elements out of the column group into a new
    `c:lineChart` sharing the same axis ids, restyled as 2.5pt lines with
    the palette accent their series position already had."""
    if not line_names:
        return
    plot_area = _plot_area(chart)
    bar = plot_area.find(qn("c:barChart"))
    if bar is None:
        return
    ax_ids = [a.get("val") for a in bar.findall(qn("c:axId"))]
    line = etree.Element(qn("c:lineChart"), nsmap=_NSMAP_CA)
    bar.addnext(line)
    _sub(line, "c:grouping", val="standard")
    _sub(line, "c:varyColors", val="0")
    moved = []
    for ser in bar.findall(qn("c:ser")):
        if _ser_name(ser) in line_names:
            bar.remove(ser)
            line.append(ser)
            moved.append(ser)
    _sub(line, "c:marker", val="1")
    for v in ax_ids:
        _sub(line, "c:axId", val=v)
    accents = _accents(pal)
    for ser in moved:
        try:
            idx = names.index(_ser_name(ser))
        except ValueError:
            idx = 0
        _restyle_as_line(ser, accents[idx % len(accents)])


def _restyle_as_line(ser, hex_color):
    """A column-styled <c:ser> becomes a line series: fill dropped, a 2.5pt
    accent line, markers off, smooth on (the chart:line spelling)."""
    for bad in ser.findall(qn("c:invertIfNegative")):
        ser.remove(bad)                     # CT_LineSer does not allow it
    spPr = ser.find(qn("c:spPr"))
    if spPr is None:
        spPr = etree.Element(qn("c:spPr"), nsmap=_NSMAP_CA)
        tx = ser.find(qn("c:tx"))
        (tx if tx is not None else ser[-1]).addnext(spPr)
    for fill in spPr.findall(qn("a:solidFill")):
        spPr.remove(fill)
    for old in spPr.findall(qn("a:ln")):
        spPr.remove(old)
    ln = _sub(spPr, "a:ln", w=str(int(Pt(LINE_SERIES_W_PT))))
    fill = _sub(ln, "a:solidFill")
    _sub(fill, "a:srgbClr", val=hex_color)
    marker = etree.Element(qn("c:marker"), nsmap=_NSMAP_CA)
    _sub(marker, "c:symbol", val="none")
    spPr.addnext(marker)
    _sub(ser, "c:smooth", val="1")


# ── secondary value axis ────────────────────────────────────────────────────
_SECONDARY_AXES = (
    '<root xmlns:c="{c}" xmlns:a="{a}">'
    '<c:catAx><c:axId val="{cat}"/>'
    '<c:scaling><c:orientation val="minMax"/></c:scaling>'
    '<c:delete val="1"/><c:axPos val="b"/><c:majorTickMark val="none"/>'
    '<c:minorTickMark val="none"/><c:tickLblPos val="none"/>'
    '<c:crossAx val="{val}"/></c:catAx>'
    '<c:valAx><c:axId val="{val}"/>'
    '<c:scaling><c:orientation val="minMax"/></c:scaling>'
    '<c:delete val="0"/><c:axPos val="r"/><c:majorTickMark val="out"/>'
    '<c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>'
    '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1200">'
    '<a:solidFill><a:srgbClr val="{muted}"/></a:solidFill>'
    '<a:latin typeface="{font}"/></a:defRPr></a:pPr>'
    '<a:endParaRPr lang="en-US"/></a:p></c:txPr>'
    '<c:crossAx val="{cat}"/><c:crosses val="max"/></c:valAx></root>'
)


def apply_secondary(chart, pal, spec, kind):
    """Rewire the `- Secondary:` series' chart group to a new right-hand
    value axis (hidden category-axis pair). chart:combo / chart:line only;
    a group holding both named and unnamed series is split in two."""
    if kind not in SECONDARY_KINDS:
        warn(f"- Secondary: applies to chart:combo / chart:line — "
             f"ignored on chart:{kind}")
        return
    names = parse_names(spec)
    plot_area = _plot_area(chart)
    targets, found = {}, set()
    for plot in _plot_elements(plot_area):
        for ser in plot.findall(qn("c:ser")):
            if _ser_name(ser) in names:
                targets.setdefault(plot, []).append(ser)
                found.add(_ser_name(ser))
    for n in names:
        if n not in found:
            warn(f"Secondary {n!r} is not a plotted series — skipped")
    if not targets:
        return
    cat_id, val_id = _new_axis_ids(plot_area)
    for plot, sers in targets.items():
        if len(sers) == len(plot.findall(qn("c:ser"))):
            _rewire_axis_ids(plot, cat_id, val_id)
        else:
            _split_plot_group(plot, sers, cat_id, val_id)
    _append_secondary_axes(plot_area, pal, cat_id, val_id)


def _new_axis_ids(plot_area):
    used = set()
    for a in plot_area.iter(qn("c:axId")):
        try:
            used.add(int(a.get("val")))
        except (TypeError, ValueError):
            continue
    base = max(used, default=100) + 1
    return base, base + 1


def _rewire_axis_ids(plot, cat_id, val_id):
    for a, v in zip(plot.findall(qn("c:axId")), (cat_id, val_id)):
        a.set("val", str(v))


def _split_plot_group(plot, sers, cat_id, val_id):
    """Move `sers` out of `plot` into a twin group (same tag and settings)
    wired to the new axis pair, inserted right after the original."""
    twin = deepcopy(plot)
    insert_at = None
    for i, child in enumerate(list(twin)):
        if child.tag == qn("c:ser"):
            if insert_at is None:
                insert_at = i
            twin.remove(child)
    if insert_at is None:
        insert_at = len(twin)
    for ser in reversed(sers):
        plot.remove(ser)
        twin.insert(insert_at, ser)
    _rewire_axis_ids(twin, cat_id, val_id)
    plot.addnext(twin)


def _append_secondary_axes(plot_area, pal, cat_id, val_id):
    root = etree.fromstring(_SECONDARY_AXES.format(
        c=nsuri("c"), a=nsuri("a"), cat=cat_id, val=val_id,
        muted=pal["text_muted"], font=pal["font_body"]))
    anchor = plot_area.find(qn("c:spPr"))   # axes precede plot-area spPr
    for ax in list(root):
        if anchor is not None:
            anchor.addprevious(ax)
        else:
            plot_area.append(ax)


# ── trendline ───────────────────────────────────────────────────────────────
# CT_Trendline anchors: the first of these children of <c:ser> that exists
# is where the trendline element must be inserted before.
_TREND_ANCHORS = ("c:errBars", "c:cat", "c:xVal", "c:val", "c:yVal",
                  "c:bubbleSize", "c:smooth")


def apply_trendline(chart, pal, spec, kind):
    """Native `c:trendline` on the first plotted series: dashed accent3,
    equation/R² off, optional `c:name` from a trailing quoted label."""
    if kind not in TREND_KINDS:
        warn(f"- Trend: not supported on chart:{kind} — skipped "
             f"(kinds: {', '.join(TREND_KINDS)})")
        return
    ttype, order, label = parse_trend(spec)
    plots = _plot_elements(_plot_area(chart))
    ser = plots[0].find(qn("c:ser")) if plots else None
    if ser is None:
        return
    tl = etree.Element(qn("c:trendline"), nsmap=_NSMAP_CA)
    if label:
        name = _sub(tl, "c:name")
        name.text = label
    spPr = _sub(tl, "c:spPr")
    ln = _sub(spPr, "a:ln", w=str(int(Pt(TREND_W_PT))))
    fill = _sub(ln, "a:solidFill")
    _sub(fill, "a:srgbClr", val=pal["accent3"])
    _sub(ln, "a:prstDash", val="dash")
    _sub(tl, "c:trendlineType", val=ttype)
    if order is not None:
        _sub(tl, "c:order", val=order)
    _sub(tl, "c:dispRSqr", val="0")
    _sub(tl, "c:dispEq", val="0")
    for tag in _TREND_ANCHORS:
        anchor = ser.find(qn(tag))
        if anchor is not None:
            anchor.addprevious(tl)
            return
    ser.append(tl)


# ── scatter groups / bubble data ────────────────────────────────────────────
def xy_named_points(p, names, multi_series, default_name, kind):
    """[(series name, [(x, y[, size]), ...])] as the chart will plot (and
    cache) them, or None when the plain shared-X scatter path applies
    (charts.scatter_xy — kept byte-identical for ungrouped scatter)."""
    data = list(p.get("data") or [])
    groups = data_groups(data)
    y_cols = (len(names) - 1 if multi_series else 1)
    if kind == "bubble":
        y_cols -= 1                       # the last column is the size
    if groups is not None and multi_series and y_cols > 1:
        warn("[Group] labels only apply with a single Y column — ignored "
             f"(**Series:** {', '.join(names)})")
        groups = None
        data = [(split_group(lbl)[0], v) for lbl, v in data]
    if kind == "scatter" and groups is None:
        return None                       # plain path (charts.scatter_xy)
    single_name = _xy_single_name(names, default_name, kind)
    buckets = groups or [(single_name,
                          [(i, lbl, v) for i, (lbl, v) in enumerate(data)])]
    from charts import label_number
    all_labels_numeric = all(
        label_number(lbl) is not None
        for _, rows in buckets for _, lbl, _v in rows)
    named = []
    for grp, rows in buckets:
        pts = []
        for i, lbl, vals in rows:
            row = list(vals) if isinstance(vals, list) else [vals]
            pts.append(_xy_point(row, lbl, i, multi_series,
                                 all_labels_numeric, kind))
        named.append((grp, [pt for pt in pts if pt is not None]))
    return named


def _xy_single_name(names, default_name, kind):
    """The single-series name: the Y column's **Series:** name when the
    full X, Y[, Size] header was given, else the fallback."""
    expected = 3 if kind == "bubble" else 2
    if len(names) == expected:
        return names[1]
    return default_name or "Series 1"


def _xy_point(row, label, index, multi_series, labels_numeric, kind):
    """One data row -> (x, y) / (x, y, size), or None when it cannot map."""
    from charts import label_number
    need = 3 if kind == "bubble" else 2
    if multi_series and len(row) >= need:
        return tuple(row[:need])
    if kind == "bubble":
        return tuple(row[:3]) if len(row) >= 3 else None
    x = label_number(label) if labels_numeric else float(index + 1)
    return (x, row[0]) if row else None


def place_xy(slide, p, pal, kind, names, multi_series, default_name,
             left, top, w, h):
    """Build `chart:scatter` / `chart:bubble` into the (already scaled)
    box: partition tint rects first, then the native chart, then explicit
    axis ranges when a partition needs the plot-box mapping to hold."""
    from charts import add_native_chart, scatter_xy, _nice_ceil
    named = xy_named_points(p, names, multi_series, default_name, kind)
    part = parse_partition(p["partition"]) if p.get("partition") else None
    ranges = None
    if part:
        xs, ys = _xy_extent(named, p, names, multi_series)

        def _rng(vals, cut):
            # include 0, every point and the partition value; extend below 0
            # symmetrically so negative data is never clipped off the axis
            vals = list(vals) + ([cut] if cut is not None else [])
            hi = max(vals + [0])
            lo = min(vals + [0])
            hi = float(_nice_ceil(hi * 1.05)) if hi > 0 else 0.0
            lo = -float(_nice_ceil(-lo * 1.05)) if lo < 0 else 0.0
            if lo == hi:   # degenerate (all zero): keep a unit axis
                hi = 1.0
            return (lo, hi)

        ranges = (_rng(xs or [1], part.get("x")),
                  _rng(ys or [1], part.get("y")))
        if "x" in part and "y" in part:
            _draw_quadrants(slide, pal, part, ranges, left, top, w, h)
    if named is None:
        # XY contract (charts.scatter_xy): multi-series -> first column is
        # X, the rest are Y series; single -> numeric labels (or the row
        # index) are X.
        xs, ys = scatter_xy(p["data"], names if multi_series else None)
        chart = add_native_chart(slide, pal, kind, xs, ys, left, top, w, h,
                                 series_name=default_name)
    else:
        chart = add_native_chart(slide, pal, kind, None, named,
                                 left, top, w, h)
    if ranges:
        (x0, x1), (y0, y1) = ranges
        chart.category_axis.minimum_scale = x0   # XY: category axis = X
        chart.category_axis.maximum_scale = x1
        chart.value_axis.minimum_scale = y0
        chart.value_axis.maximum_scale = y1
    return chart


def _xy_extent(named, p, names, multi_series):
    """(xs, ys) of every plotted point, for axis-range computation."""
    if named is not None:
        pts = [pt for _, series_pts in named for pt in series_pts]
        return ([pt[0] for pt in pts if pt[0] is not None],
                [pt[1] for pt in pts if pt[1] is not None])
    from charts import scatter_xy
    xs, ys = scatter_xy(p.get("data") or [], names if multi_series else None)
    flat = ([v for vals in ys.values() for v in vals]
            if isinstance(ys, dict) else list(ys))
    return ([x for x in xs if x is not None],
            [y for y in flat if y is not None])


def _draw_quadrants(slide, pal, part, ranges, left, top, w, h):
    """Four quadrant tint rects, drawn before the chart so markers stay on
    top. Plot-box insets are approximations — confirm in visual QA."""
    from charts import _plot_box
    px, py, pw, ph = _plot_box(left, top, w, h)
    (x0, x1), (y0, y1) = ranges
    fx = px + (part["x"] - x0) / (x1 - x0) * pw
    fy = py + (1 - (part["y"] - y0) / (y1 - y0)) * ph
    fx = min(max(fx, px), px + pw)
    fy = min(max(fy, py), py + ph)
    boxes = ((px, py, fx - px, fy - py),
             (fx, py, px + pw - fx, fy - py),
             (px, fy, fx - px, py + ph - fy),
             (fx, fy, px + pw - fx, py + ph - fy))
    for bx, by, bw, bh in boxes:
        if bw <= 0 or bh <= 0:
            continue
        rect = add_rect(slide, bx, by, bw, bh, pal["surface"])
        set_fill_alpha(rect, PARTITION_ALPHA)


def finish_xy(slide, chart, pal, p, kind, left, top, w, h):
    """Step-6 shape annotations for scatter/bubble, drawn against the final
    axis ranges: dashed partition lines and the bubble size legend."""
    if p.get("partition"):
        _draw_partition_lines(slide, chart, pal, p, left, top, w, h)
    if kind == "bubble" and _flag_on(p.get("size_legend")):
        _draw_size_legend(slide, pal, p, left, top, w, h)


def _flag_on(value):
    text = str(value or "").strip().lower()
    if text in ("on", "yes", "true", "1"):
        return True
    if text and text not in ("off", "no", "false", "0"):
        warn(f"Size-Legend value {text!r} is not on|off — treated as off")
    return False


def _axis_range(axis):
    lo = axis.minimum_scale
    hi = axis.maximum_scale
    if lo is None or hi is None or hi <= lo:
        return None
    return lo, hi


def _draw_partition_lines(slide, chart, pal, p, left, top, w, h):
    from charts import _plot_box
    part = parse_partition(p["partition"])
    px, py, pw, ph = _plot_box(left, top, w, h)
    x_range = _axis_range(chart.category_axis)   # XY: category axis = X
    y_range = _axis_range(chart.value_axis)
    if "x" in part and x_range:
        lo, hi = x_range
        if lo <= part["x"] <= hi:
            fx = px + (part["x"] - lo) / (hi - lo) * pw
            _dashed(slide, pal, fx, py, fx, py + ph)
        else:
            warn(f"Partition x={part['x']:g} is outside the axis range "
                 f"{lo:g}-{hi:g} — line skipped")
    if "y" in part and y_range:
        lo, hi = y_range
        if lo <= part["y"] <= hi:
            fy = py + (1 - (part["y"] - lo) / (hi - lo)) * ph
            _dashed(slide, pal, px, fy, px + pw, fy)
        else:
            warn(f"Partition y={part['y']:g} is outside the axis range "
                 f"{lo:g}-{hi:g} — line skipped")


def _dashed(slide, pal, x0, y0, x1, y1):
    conn = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT, Inches(x0), Inches(y0),
        Inches(x1), Inches(y1))
    conn.line.color.rgb = hex_rgb(pal["accent3"])
    conn.line.width = Pt(PARTITION_W_PT)
    conn.line.dash_style = MSO_LINE_DASH_STYLE.DASH
    return conn


def _bubble_sizes(p):
    sizes = []
    for _, vals in p.get("data") or []:
        row = vals if isinstance(vals, list) else [vals]
        if len(row) >= 3 and isinstance(row[2], (int, float)):
            sizes.append(float(row[2]))
    return sizes


def _draw_size_legend(slide, pal, p, left, top, w, h):
    """Two reference circles (min/max bubble size) bottom-right inside the
    plot box, labelled via fmt_abs."""
    from charts import _plot_box
    sizes = _bubble_sizes(p)
    if not sizes:
        return
    px, py, pw, ph = _plot_box(left, top, w, h)
    r_min, r_max = SIZE_LEGEND_R
    cy = py + ph - r_max - 0.30
    entries = ((min(sizes), r_min, px + pw - 2 * r_max - r_min - 0.75),
               (max(sizes), r_max, px + pw - 2 * r_max - 0.05))
    for value, r, cx in entries:
        shape = slide.shapes.add_shape(
            MSO_SHAPE.OVAL, Inches(cx), Inches(cy + (r_max - r)),
            Inches(2 * r), Inches(2 * r))
        shape.fill.background()
        shape.line.color.rgb = hex_rgb(pal["text_muted"])
        shape.line.width = Pt(1.0)
        tb = slide.shapes.add_textbox(
            Inches(cx - 0.3), Inches(cy + 2 * r_max + 0.02),
            Inches(2 * r + 0.6), Inches(0.22))
        run = tb.text_frame.paragraphs[0].add_run()
        run.text = fmt_abs(value)
        run.font.size = Pt(9.5)
        run.font.color.rgb = hex_rgb(pal["text_muted"])
        run.font.name = pal["font_label"]
        from pptx.enum.text import PP_ALIGN
        tb.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER


# ── refresh_deck pairing ────────────────────────────────────────────────────
def xy_refresh_points(p, names, kind, default_name):
    """[(series name, x, y)] exactly as the built chart caches them —
    grouped rows keyed by group, bubble rows by their Y value (the size
    column is not cached as a series value)."""
    multi = len(names) > 1
    named = xy_named_points(p, names, multi, default_name, kind)
    if named is None:
        from charts import scatter_xy
        xs, ys = scatter_xy(p.get("data") or [], names if multi else None)
        series = ys if isinstance(ys, dict) else {names[0]: ys}
        return [(n, x, y) for n, vals in series.items()
                for x, y in zip(xs, vals)]
    return [(n, pt[0], pt[1]) for n, pts in named for pt in pts
            if pt is not None]
