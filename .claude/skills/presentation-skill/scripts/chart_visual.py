"""chart_visual.py — place an outline slide's `chart:<kind>` visual.

The chart branch of `builders._place_visual`, split out so the chart option
modules (annotations, options, combo, data rows) hook into one seam instead
of growing builders.py.  `place_chart()` builds the native chart from the
slide's **Data:** rows and then applies the per-slide chart fields in a fixed
order — the order is part of the contract because several steps set the
value-axis range and later steps must be drawn against the final range:

  1. data shaping   (`- Sort:` / `- Ordinal:` / series routed to ticks, error
                     bars, data rows — anything that changes what is plotted)
  2. native chart   (`charts.add_native_chart`, or the combo builder)
  3. benchmark / CAGR arrows  (may widen the value axis)
  4. `- Axis-Max:`  (wins over 3 unless the CAGR path already applied it)
  5. native options (labels mode, gap width, label rotation/skip, units,
                     trendline, error bars, secondary axis)
  6. shape-drawn annotations against the final axis range (value line,
     difference arrows, target ticks, 100% line, category breaks, data rows)

Geometry is design-space inches; `builders._sc` rescales to the slide.
Nothing here imports build_deck.
"""
from charts import add_native_chart
from helpers import warn

# Chart kinds whose value axis is vertical and whose categories run left to
# right — the ones the plot-box approximation (charts._plot_box) can place
# shape-drawn annotations on.
VERTICAL_KINDS = ("bar", "column", "line", "area", "stacked", "stacked-100",
                  "combo")


def split_series(p):
    """(names, multi) — the **Series:** names of a slide and whether the
    data rows carry one value per name."""
    series_field = p.get("series", "")
    names = [n.strip() for n in series_field.split(",") if n.strip()]
    return names, bool(series_field and "," in series_field)


def series_dict(p, names):
    """{name: [values]} for a multi-series slide (rows are value lists)."""
    return {
        names[i]: [row[1][i] if isinstance(row[1], list) else row[1]
                   for row in p["data"]]
        for i in range(len(names))
    }


def single_values(p):
    """[value] for a single-series slide (rows may still be 1-lists)."""
    return [row[1][0] if isinstance(row[1], list) else row[1]
            for row in p["data"]]


def place_chart(slide, p, pal, left, top, w, h, kind):
    """Build the slide's native chart for visual `chart:<kind>` into the box
    (design inches) and apply the chart fields.  Returns the pptx Chart."""
    import builders  # late: builders imports this module's caller
    import chart_options
    box = builders._sc(left, top, w, h)
    # step 1: data shaping (Sort/Ordinal reorder; Target/Error-Bars series
    # routed out of the plotted set). Returns p untouched with no new keys.
    p = chart_options.shape_data(p, kind)
    data_rows_plan = None
    if p.get("data_rows"):  # step 1: series -> row routing + chart-box shrink
        from chart_data_rows import route_rows
        p, h, data_rows_plan = route_rows(p, kind, h)
        box = builders._sc(left, top, w, h)
    cats = [d[0] for d in p["data"]]
    names, multi_series = split_series(p)
    series_field = p.get("series", "")
    if kind in ("scatter", "bubble"):
        # XY contract (charts.scatter_xy) plus the chart_combo extensions:
        # [Group] series bucketing, bubble (x, y, size) rows, partition
        # tint rects (drawn before the chart) with explicit axis ranges.
        # Benchmark/CAGR/value-line don't apply.
        from chart_combo import place_xy
        chart = place_xy(slide, p, pal, kind, names, multi_series,
                         series_field or p.get("heading", ""), *box)
    elif kind == "combo":
        # clustered columns + Line-Series lines on shared axes
        from chart_combo import place_combo
        chart = place_combo(slide, p, pal, cats, names,
                            series_dict(p, names), *box)
    elif multi_series:
        chart = add_native_chart(slide, pal, kind, cats, series_dict(p, names),
                                 *box)
    else:
        chart = add_native_chart(slide, pal, kind, cats, single_values(p), *box,
                                 series_name=series_field or p.get("heading", ""))
        if p.get("benchmark") and kind in ("bar", "column", "line", "area"):
            from charts import add_benchmark_line
            add_benchmark_line(slide, chart, pal, p["benchmark"], *box)
        if p.get("cagr") and kind in ("bar", "column", "line") \
                and chart_options.cagr_mode(p["cagr"], cats) == "legacy":
            from charts import add_cagr_arrow
            try:
                user_max = float(p["axis_max"]) if p.get("axis_max") else None
            except (ValueError, TypeError):
                user_max = None
                warn(f"Axis-Max not numeric: {p.get('axis_max')!r} "
                     f"(chart:{kind}, CAGR)")
            add_cagr_arrow(slide, chart, pal, *box, axis_max=user_max)
    # CAGR path passes axis_max into add_cagr_arrow; don't overwrite after.
    cagr_sets_axis = (
        p.get("cagr") and not multi_series
        and kind in ("bar", "column", "line")
        and chart_options.cagr_mode(p["cagr"], cats) == "legacy")
    if p.get("axis_max") and not cagr_sets_axis:
        try:
            chart.value_axis.maximum_scale = float(p["axis_max"])
            chart.value_axis.minimum_scale = 0.0
        except (ValueError, TypeError):
            warn(f"Axis-Max not numeric: {p['axis_max']!r}")
    if p.get("labels_mode") and kind == "stacked-100":
        from charts import LABEL_MODES, add_stacked_100_labels
        if p["labels_mode"] in LABEL_MODES:
            add_stacked_100_labels(chart, p["labels_mode"], pal)
        else:
            warn(f"Labels mode {p['labels_mode']!r} not one of "
                 f"{'|'.join(LABEL_MODES)} — labels skipped")
    # step 5: native options (gap width, label rotation/skip, error bars,
    # secondary axis, trendline) against the final primary-axis range
    if kind not in ("scatter", "bubble"):
        chart_options.apply_native(chart, p, pal, kind)
    if p.get("secondary"):
        from chart_combo import apply_secondary
        apply_secondary(chart, pal, p["secondary"], kind)
    if p.get("trend"):
        from chart_combo import apply_trendline
        apply_trendline(chart, pal, p["trend"], kind)
    # step 6: shape-drawn annotations against the final axis range
    if p.get("value_line") and kind in ("bar", "column", "line"):
        # mean|median|wmean specs resolve to a numeric spec; numeric specs
        # pass through unchanged (today's behavior byte-for-byte)
        spec = chart_options.resolve_value_line(p, kind)
        if spec:
            from charts import add_value_line
            add_value_line(slide, chart, pal, spec, *box)
    if kind not in ("scatter", "bubble"):
        chart_options.draw_annotations(slide, chart, p, pal, box, kind)
    if p.get("diff_arrows"):
        from chart_annotations import add_diff_arrows
        add_diff_arrows(slide, chart, p, pal, box, kind)
    if kind in ("scatter", "bubble") and (p.get("partition")
                                          or p.get("size_legend")):
        from chart_combo import finish_xy
        finish_xy(slide, chart, pal, p, kind, *box)
    if data_rows_plan:  # step 6: draw the routed rows under the chart
        from chart_data_rows import draw_rows
        draw_rows(slide, p, pal, data_rows_plan, left, top, w, h)
    return chart
