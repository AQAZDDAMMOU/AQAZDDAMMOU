# Charts Guide

## Preferred Path: Outline-Driven Charts

You normally don't write chart code. Declare the chart in the outline and `build_deck.py` generates it via `scripts/charts.py`:

```markdown
**Visual:** chart:bar        ← bar | hbar | line | pie | doughnut | area | scatter | bubble | combo | stacked | stacked-100
**Series:** Market Size ($B)
**Data:**
- 2024: 42
- 2025: 67
- 2026: 104
```

The pipeline automatically applies palette colors, axis/tick label styling, and — critically — transparent chart and plot areas.

### Multi-series (outline)

Comma-separate series names and values per row:

```markdown
**Visual:** chart:line
**Series:** Revenue, Costs
**Data:**
- Q1: 4.2, 3.1
- Q2: 5.1, 3.0
- Q3: 6.3, 2.8
```

`--check` validates that every data row has the same number of comma-separated values as series names.

## The White-Box Defect (and the fix)

Default python-pptx charts render with an **opaque white chart area** — the most common visual defect in AI-generated decks on dark slides. python-pptx has no public API for the chart-area fill, so `scripts/charts.py` clears it at the XML level:

```python
from charts import make_chart_transparent, add_native_chart

# add_native_chart() already calls this. If you hand-roll a chart, you MUST:
make_chart_transparent(chart)   # inserts <c:spPr><a:noFill/> on chartSpace + plotArea
```

`scripts/qa_check.py` flags any chart that is missing this.

`chart:scatter` is an XY chart (markers only): with single-value rows the numeric labels are x (else the row index); with `**Series:** X, Y1[, Y2 …]` and `- label: x, y1, …` rows the first column is X and every further column is one Y series. Full contract in `generation-guide.md` ("Data blocks").

### Combo (bar + line) and bubble (`scripts/chart_combo.py`)

`chart:combo` builds one clustered column chart with every `**Series:**` name, then moves the `- Line-Series:` series (default: the last name) into a `c:lineChart` group inside the same `c:plotArea`, sharing the column group's axis ids — one natively editable chart, not two stacked frames. `- Secondary: <names>` adds a right-hand `c:valAx` (delete=0, crosses at max) with a hidden `c:catAx` partner and rewires the named series' chart group to the new pair; a group holding both named and unnamed series is split in two. `- Trend:` appends a native `c:trendline` (dashed accent3, equation/R² off) to the first plotted series.

`chart:bubble` uses `BubbleChartData` (`**Series:** X, Y, Size`, 3-value rows); it is transparent-styled like scatter with series fills at ~60% alpha so overlaps read. On both scatter and bubble a `[Group]` row-label suffix buckets single-Y rows into one series per group, and `- Partition: x=…, y=…` draws quadrant tints (before the chart) plus dashed lines against explicitly-set axis ranges. `- Size-Legend: on` adds min/max reference circles.

QA awareness: `chart_workbook` reads a combo as the union of its bar+line groups and a bubble like a scatter (the size column is not compared); `chart_lint` parses both — its dual-axis honesty warn intentionally still fires on a `- Secondary:` chart; `deck_to_outline` recovers `chart:combo` + `- Line-Series:` and `- Item N [group]: x, y[, size]` rows.

Outline syntax and field-by-field rules: `generation-guide.md` § "Newer directives (Cycle 7 — think-cell parity)".

## Custom Charts

For chart placement the outline can't express, call the helper directly (for `"scatter"`, pass the x values as `categories`):

```python
from charts import add_native_chart

add_native_chart(slide, pal, "line",
                 categories=["Q1", "Q2", "Q3"],
                 series={"Revenue": (4.2, 5.1, 6.3), "Costs": (3.1, 3.0, 2.8)},
                 left=7.0, top=1.5, w=5.5, h=4.5)
```

Multi-series gets a legend automatically; pie/doughnut slices cycle through the palette accents.

---

## Shape-Based Visuals (better than charts for stats)

For stat callouts, progress bars, and KPI displays, custom shapes look more polished than native chart objects. The `stat-callout` layout covers cards; for progress bars:

```python
from helpers import add_rect, add_tb

def add_progress_bar(slide, pal, label, percent, value_label, left, top,
                     width=9.0, height=0.45):
    add_tb(slide, label, left, top, 3.5, height,
           size=14, color=pal["text"], font=pal["font_body"])
    track_w = width - 4.0
    add_rect(slide, left + 3.7, top + 0.05, track_w, height - 0.1, pal["surface"])
    add_rect(slide, left + 3.7, top + 0.05, track_w * percent / 100,
             height - 0.1, pal["accent1"])
    add_tb(slide, value_label, left + width - 0.25, top, 0.9, height,
           size=13, bold=True, color=pal["accent1"], font=pal["font_body"])
```

## Annotations

- **Benchmark line:** `- Benchmark: 120 Industry average` (existing).
- **Value line:** `- Value-Line: <label>, <value>` — see below. Cycle 7 adds
  stat value lines: `- Value-Line: mean|median|wmean ["label"]` computes the
  value from the plotted data (see "Chart options (Cycle 7)").
- **CAGR arrow:** `- CAGR: on` — computed from first/last values of a
  single-series bar/column/line chart. Plot-box placement is approximate:
  always confirm in visual QA. Cycle 7 adds `- CAGR: series` (per-series
  labels) and `- CAGR: FY22-FY24, FY24-FY27` (span arrows) — see below.
- **Difference arrows:** `- Diff-Arrow:` (repeatable) — think-cell-style
  deltas between two bars, see "Difference arrows (Cycle 7)".
- **Target ticks / error bars / category breaks:** routed series and plot
  marks, see "Chart options (Cycle 7)".
- **Same scale:** `- Axis-Max: N` pins the value axis (applied last, so it
  wins over the axis max set by Benchmark/CAGR). Use the same N across
  slides being compared — or let `**Scale-Group:** auto` do it (below).
- **Waterfall bracket:** `- Bracket: <bar A>, <bar B>[, "label"]` — exact
  geometry (shape-drawn), no QA caveat.

## Difference arrows (Cycle 7)

```markdown
**Visual:** chart:bar
- Diff-Arrow: FY24, FY26            ← total difference, default mode pct
- Diff-Arrow: FY24, FY26 both       ← modes: abs | pct | both
- Diff-Arrow: Revenue@FY26 -> Costs@FY26   ← level difference (two series)
**Data:**
- FY24: 50
- FY26: 62
```

- **Total difference** (`CatA, CatB [mode]`): single series compares the two
  bar values; multi series compares per-category totals (stacked reading).
  Drawn think-cell style: a thin leader from the top of each bar to an arrow
  column just right of the plot box, a vertical double-headed connector
  between the leaders, and a rounded bubble label centred on it.
- **Level difference** (`SerA@Cat -> SerB@Cat`, `->` or `→`): a vertical
  double-headed connector directly above the category's cluster between the
  two series' bar tops, bubble label beside it.
- Labels: `abs` → `+12`, `pct` (default) → `+24%` of the *from* value,
  `both` → `+12 (+24%)` — all through the deck number policy
  (`fmt_abs`/`fmt_pct`). A zero/negative base warns and falls back to abs.
- Kinds: bar/column/line/stacked/stacked-100. On stacked-100 the plotted
  column totals are constant (100%), so the arrow compares segment sums and
  the build says so; `hbar` warns and skips (horizontal geometry is not v1).
- Unknown category/series names warn and skip that arrow. A malformed spec
  (neither grammar) is a fail-fast `--check` error. Axis handling mirrors
  the CAGR arrow: an existing axis max is respected, never shrunk; plot-box
  placement is approximate — confirm in visual QA.

## Chart options (Cycle 7)

All of these are per-slide chart fields applied by `chart_visual.place_chart`
in a fixed order (data shaping → native chart → CAGR/benchmark → Axis-Max →
native options → shape-drawn annotations). Bad values warn and fall back —
they never fail a build (only malformed Diff-Arrow / Error-Bars grammar is a
`--check` error).

- **Stat value lines** — `- Value-Line: mean|median|wmean ["label"]`: the
  line's value is computed from the plotted data (first/only series);
  `wmean` needs exactly 2 series and is Σ(v·w)/Σw with series 2 as weights
  (the weights series still plots). Labels default to `Mean` / `Median` /
  `Weighted mean`. The numeric `- Value-Line: <label>, <value>` form is
  unchanged.
- **Sort** — `- Sort: desc|asc` on bar/column/hbar reorders the data rows by
  the first series' value before the chart is built. (Tornado's own
  `- Sort: off` is untouched.)
- **Ordinal** — `- Ordinal: on` prefixes category labels `1. `, `2. ` … in
  plotted order (after any sort); bar/column/hbar only.
- **100% line** — `- Total-Line: on` on `chart:stacked-100` and `mekko`
  draws a thin muted rule + `100%` label at the top of the plot.
- **Target ticks** — `- Target: <series name>` on bar/column removes that
  series from the plotted set and draws it as short thick `accent3` ticks
  (60% of the category slot) at its value per category, with a `— <name>`
  legend chip under the plot. Unknown name warns, plot unchanged.
- **Error bars** — `- Error-Bars: Low, High` on bar/column/line: the two
  named series become native `c:errBars` custom min/max whiskers on the
  first remaining series (deltas from its values) and leave the plotted
  set. Fully editable in PowerPoint. Needs a third series to remain.
- **Per-series CAGR** — `- CAGR: series` on a multi-series line/column puts
  `<name> +8.2% p.a.` at each series' right end in the series color.
- **Multi-span CAGR** — `- CAGR: FY22-FY24, FY24-FY27` on a single series
  draws one arrow + `+12.4% p.a.` label per span between the named
  categories; an unknown category warns and skips that span.
- **Category breaks** — `- Category-Break: FY25 "Actual | Forecast"`
  (repeatable): dashed vertical separator after the category, the optional
  label split on `|` left/right of the line at the top.
- **Display units** — `- Units: $M` on vertical kinds: a small muted label
  above the top-left of the plot box. (On a `table` slide the same field is
  the units row.)
- **Label rotation / skip** — `- Label-Rotation: 45` (clamped ±90; native
  `a:bodyPr@rot`) and `- Label-Skip: 2` (native `c:tickLblSkip`).
- **Gap width** — `- Gap-Width: 50` on bar/column/hbar/stacked (native,
  clamped 0–500).
- **New kind** — `chart:stacked` (`COLUMN_STACKED`), styled like bar/column
  with a legend when multi-series.

## Chart Intelligence (think-cell-parity behaviors)

### Largest-remainder rounding

Displayed percentage sets always sum to the stated total: `charts.round_to_sum(
values, total=100, decimals=0)` implements largest-remainder rounding —
`[33.33, 33.33, 33.33]` renders as `33% / 33% / 34%`, never `33+33+33=99`.
Applied automatically to mekko column shares and stacked-100 `pct` labels
(per column, so each column's labels sum to 100). Any negative input falls
back to plain rounding. Single-value % labels (waterfall bracket, funnel
stage-to-stage conversion) are exempt — there is no group to reconcile.

### Value line

```markdown
**Visual:** chart:bar
- Value-Line: Target, 40
**Data:**
- 2023: 30
- 2024: 44
```

Draws a dashed horizontal reference line with a small right-aligned label
across bar/column/line chart visuals. Unlike `- Benchmark:` (which widens the
axis to fit), a value line never changes an already-set axis: if the value
falls outside the computed axis range, the build prints a warning and skips
the line. Place it after `- Axis-Max:` decisions — it is drawn against the
final axis range.

### Dual labels on 100%-stacked columns

```markdown
**Visual:** chart:stacked-100
**Series:** Subs, Services, Hardware
- Labels: both                ← pct | abs | both
**Data:**
- 2023: 1, 1, 1
- 2024: 8.4, 7.6, 4
```

`pct` renders per-segment percentages (largest-remainder per column, so each
column shows exactly 100), `abs` renders the raw values, `both` renders
`42% (8.4)`. Without `- Labels:` the chart renders unlabeled (default
behavior unchanged).

### Same-scale groups

```markdown
**Scale-Group:** auto         ← deck front-matter; default off
```

When `auto`, slides whose chart visuals share a kind (bar/column/hbar/line/
area), have numeric data, and have no explicit `- Axis-Max:` are grouped per
kind; each group of 2+ slides gets one shared axis maximum (`_nice_ceil` of
the group's global data max), printed at build time:

```
  Scale-group: column charts on slides 4, 7 share axis max 50
```

The reverse check lives in `pptx_lint.py`: same-type native charts whose
explicit value-axis maxima differ by more than 10x are flagged as a
"possible dishonest scale comparison" warning.

### Data rows (values printed under the chart)

```markdown
**Visual:** chart:column
**Series:** Revenue, Margin %
- Data-Rows: values, growth
**Data:**
- FY22: 120, 30
- FY23: 150, 28
```

think-cell-style data rows under the plot, aligned to the category centres.
Row kinds (comma list): `values` (one row per plotted series, label = series
name), `growth` (period-over-period % of the first plotted series, first
cell `–`), `cagr` (one row, the start→end CAGR under the last column, label
`CAGR p.a.`), or a **Series:** name — that series is removed from the chart
and printed as a row instead (`- Data-Rows: values, Margin %`). The chart
box shortens 0.42 in per row (never below 2.2 in — rows that don't fit warn
and drop). Vertical categorical kinds only (bar/column/line/stacked/
stacked-100/combo); others warn and ignore. Numbers follow the deck's
`**Number-Format:**` policy; a series whose name contains `%` prints as
percentages. Row placement uses the plot-box approximation — confirm
alignment in visual QA.
