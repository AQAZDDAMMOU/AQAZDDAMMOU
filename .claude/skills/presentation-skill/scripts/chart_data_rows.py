"""chart_data_rows.py — think-cell-style data rows printed under a chart.

Outline syntax (comma list; parser maps `- Data-Rows:` to p["data_rows"]):

    - Data-Rows: values, growth
    - Data-Rows: values, Margin %      <- a **Series:** name prints as a row

Row kinds: `values` (each plotted series' values, one row per series, label =
series name), `growth` (period-over-period % of the first plotted series,
first cell an en dash), `cagr` (single row, start->end CAGR under the last
column, label "CAGR p.a."), or a **Series:** name — that series prints as a
row and is removed from the plotted set.

Hooked from chart_visual.place_chart at step 1 (`route_rows`: series routing
+ chart-box shrink — 0.42 in per row, min chart height 2.2 in, excess rows
warn + drop) and step 6 (`draw_rows`: text boxes aligned to category centres
via charts._plot_box, label right-aligned in the left gutter).  Only the
vertical categorical kinds carry rows; others warn + ignore.  When the key
is absent neither hook fires, so existing decks are byte-identical.

Geometry is design-space inches, drawn through builders' scaled helpers.
Nothing here imports build_deck.
"""
from pptx.enum.text import PP_ALIGN

from chart_visual import series_dict, single_values, split_series
from charts import _plot_box
from helpers import warn
from numfmt import EN_DASH, fmt_abs, fmt_pct

# Kinds whose value axis is vertical and categories run left to right —
# "stacked" and "combo" are routed here even before charts.CHART_TYPES
# grows them, so the Cycle-7 chart modules compose.
ROW_KINDS = ("bar", "column", "line", "stacked", "stacked-100", "combo")

ROW_H = 0.42          # inches of chart height each data row claims
MIN_CHART_H = 2.2     # the chart box never shrinks below this
LABEL_MAX_W = 1.9     # row-label gutter cap (inches)
LABEL_MIN_X = 0.3     # labels never start left of this canvas x
LABEL_CHAR_W = 0.085  # ~11pt average glyph width — truncation budget
FONT_PT = 11
CELL_H = 0.30         # text-box height inside each 0.42 row slot
RULE_PT_IN = 0.012    # thin surface rule above the first row


def _values_of(p, name, names):
    """[values] of one plotted series in p (multi or single)."""
    if len(names) > 1:
        return series_dict(p, names)[name]
    return single_values(p)


def _fmt_value(v, series_name):
    """A values-row cell: fmt_pct for a '%' series, fmt_abs otherwise."""
    if "%" in (series_name or ""):
        return fmt_pct(v)
    return fmt_abs(v)


def _growth_cells(values):
    """Period-over-period % of `values`; first (and non-computable) cells
    are an en dash."""
    cells = [EN_DASH]
    for prev, cur in zip(values, values[1:]):
        if prev is None or cur is None or prev == 0:
            cells.append(EN_DASH)
        else:
            cells.append(fmt_pct((cur - prev) / prev * 100, signed=True))
    return cells


def _cagr_cells(values):
    """Blank cells with the start->end CAGR under the last column."""
    cells = [None] * len(values)
    pairs = [(i, v) for i, v in enumerate(values) if v is not None]
    if len(pairs) < 2 or pairs[0][1] <= 0 or pairs[-1][1] <= 0:
        warn("Data-Rows: cagr needs two positive endpoint values — "
             "printed as " + EN_DASH)
        cells[-1] = EN_DASH
        return cells
    (i0, v0), (i1, v1) = pairs[0], pairs[-1]
    pct = ((v1 / v0) ** (1 / (i1 - i0)) - 1) * 100
    cells[-1] = fmt_pct(pct, signed=True, legacy_decimals=1)
    return cells


def _classify(tokens, names):
    """[(kind, series-or-None)] per token; `removed` = series routed out of
    the plotted set. Unknown tokens warn + drop; the last plotted series is
    never removed (warn, row still prints)."""
    specs, removed = [], []
    for tok in tokens:
        key = tok.casefold()
        if key in ("values", "growth", "cagr"):
            specs.append((key, None))
            continue
        match = next((n for n in names if n.casefold() == key), None)
        if match is None:
            warn(f"Data-Rows: {tok!r} is not values/growth/cagr or a "
                 "**Series:** name — row dropped")
        elif len(names) - len(removed) <= 1:
            warn(f"Data-Rows: {match!r} is the only plotted series — "
                 "kept in the chart, printed as a row")
            specs.append(("series-keep", match))
        else:
            removed.append(match)
            specs.append(("series", match))
    return specs, removed


def _cap_specs(specs, removed, names, h):
    """Drop the specs that would shrink the chart below MIN_CHART_H (warn,
    un-removing any named series whose row no longer prints). Returns
    (kept_specs, total_rows)."""
    n_plot = max(1, len(names) - len(removed)) if names else 1
    max_rows = int((h - MIN_CHART_H) / ROW_H + 1e-6)
    kept, total = [], 0
    for spec in specs:
        n = n_plot if spec[0] == "values" else 1
        if total + n > max_rows:
            dropped = specs[len(kept):]
            warn(f"Data-Rows: only {max(max_rows, 0)} row(s) fit above the "
                 f"{MIN_CHART_H} in minimum chart height — dropped: "
                 + ", ".join(s[1] or s[0] for s in dropped))
            for kind, name in dropped:
                if kind == "series" and name in removed:
                    removed.remove(name)
            break
        kept.append(spec)
        total += n
    return kept, total


def _strip_series(p, names, removed):
    """A shallow-copied p with `removed` series taken out of **Series:**
    and out of every data row's value list."""
    keep_idx = [i for i, n in enumerate(names) if n not in removed]
    p2 = dict(p)
    p2["series"] = ", ".join(names[i] for i in keep_idx)
    p2["data"] = [
        (label, [(vals[i] if isinstance(vals, list) and i < len(vals)
                  else vals) for i in keep_idx])
        for label, vals in p["data"]
    ]
    return p2


def route_rows(p, kind, h):
    """Step-1 hook: (p, h, rows) — p with named-series rows removed from the
    plotted set, the chart height shortened per printed row, and the resolved
    [(label, [cell-text])] plan for draw_rows (None = nothing to draw)."""
    tokens = [t.strip() for t in str(p.get("data_rows", "")).split(",")
              if t.strip()]
    if not tokens:
        return p, h, None
    if kind not in ROW_KINDS:
        warn(f"Data-Rows on chart:{kind} — vertical categorical charts only "
             f"({'/'.join(ROW_KINDS)}); ignored")
        return p, h, None
    names, _multi = split_series(p)
    specs, removed = _classify(tokens, names)
    specs, total = _cap_specs(specs, removed, names, h)
    if not specs:
        return p, h, None
    p2 = _strip_series(p, names, removed) if removed else p
    plot_names, _ = split_series(p2)
    rows = []
    for spec_kind, name in specs:
        if spec_kind == "values":
            for n in (plot_names or [None]):
                vals = _values_of(p2, n, plot_names)
                rows.append((n or "Values",
                             [_fmt_value(v, n) for v in vals]))
        elif spec_kind == "growth":
            first = plot_names[0] if plot_names else None
            rows.append(("Growth %",
                         _growth_cells(_values_of(p2, first, plot_names))))
        elif spec_kind == "cagr":
            first = plot_names[0] if plot_names else None
            rows.append(("CAGR p.a.",
                         _cagr_cells(_values_of(p2, first, plot_names))))
        else:  # "series" (from the original p) / "series-keep"
            vals = _values_of(p, name, names)
            rows.append((name, [_fmt_value(v, name) for v in vals]))
    return p2, h - ROW_H * total, rows


def _truncate(text, width_in):
    budget = max(3, int(width_in / LABEL_CHAR_W))
    if len(text) <= budget:
        return text
    return text[:budget - 1].rstrip() + "…"


def draw_rows(slide, p, pal, rows, left, top, w, h_chart):
    """Step-6 hook: draw the routed `rows` under the (already shortened)
    chart box.  Cells align to charts._plot_box category centres; the row
    label is right-aligned in the gutter left of the plot.  11pt; the first
    row prints in `text`, the rest in `text_muted`; a thin `surface` rule
    sits above the first row."""
    import builders as B  # late: builders imports this module's caller
    px, _py, pw, _ph = _plot_box(left, top, w, h_chart)
    n = len(p["data"])
    if not n or not rows:
        return
    label_right = px - 0.06
    label_left = max(LABEL_MIN_X, label_right - LABEL_MAX_W)
    band_top = top + h_chart
    B.add_rect(slide, label_left, band_top, px + pw - label_left,
               RULE_PT_IN, pal["surface"])
    cell_w = pw / n
    for r, (label, cells) in enumerate(rows):
        color = pal["text"] if r == 0 else pal["text_muted"]
        ry = band_top + r * ROW_H + (ROW_H - CELL_H) / 2
        B.add_tb(slide, _truncate(label, label_right - label_left),
                 label_left, ry, label_right - label_left, CELL_H,
                 size=FONT_PT, color=color, align=PP_ALIGN.RIGHT,
                 font=pal["font_label"])
        for i, cell in enumerate(cells[:n]):
            if not cell:
                continue
            cx = px + pw * (i + 0.5) / n
            B.add_tb(slide, cell, cx - cell_w / 2, ry, cell_w, CELL_H,
                     size=FONT_PT, color=color, align=PP_ALIGN.CENTER,
                     font=pal["font_label"])
