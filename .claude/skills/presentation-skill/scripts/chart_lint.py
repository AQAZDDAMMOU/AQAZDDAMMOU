#!/usr/bin/env python3
"""
chart_lint.py — chart-honesty lint (Misviz / Zelazny rules) for native charts.

pptx_lint.check_axis_scales compares value-axis maxima *across* charts; this
module inspects each chart on its own terms and flags the classic ways an
exhibit misleads a reader:

  1. Truncated bar baseline (bar/column value axis not at zero) — error.
  2. Truncated line/area axis without an axis-range note on the slide.
  3. Dual value axes (two scales the reader cannot reconcile).
  4. 3-D perspective (``c:view3D``) — error.
  5. Pie/doughnut that is not part-to-whole (bad sum, negatives, >7 slices).
  6. Category axis with non-uniform year gaps (warn) or out-of-order dates
     (error).
  7. Missing legend (2+ series) or a value axis with no discoverable units.
  8. An indexed series with no "indexed to 100" note on the slide.
  9. More than one per-point highlight colour.
 10. A long line series that labels every point.

Facts are read from the chart XML (``chart._chartSpace``) with the python-pptx
API used for categories/values where it works.  ``chart_facts(chart)`` returns
the parsed facts as a plain dict so other lints can reuse them:

    {
      "type":              str,              # XL_CHART_TYPE name, else plot tag
      "categories":        [str],
      "series":            [{"name": str, "values": [float|None],
                             "plot_tag": str, "fill_color": str|None,
                             "show_val": bool}],
      "val_axis_min":      float|None,       # c:valAx/c:scaling/c:min
      "val_axis_max":      float|None,
      "has_legend":        bool,
      "n_val_axes":        int,
      "is_3d":             bool,
      "point_fill_colors": [str],            # c:dPt solid fills, hex, no '#'
      "show_val_labels":   bool,             # any c:dLbls/c:showVal val="1"
      "val_axis_title":    str|None,
      "cat_axis_title":    str|None,
    }

Usage (as a library):
    from chart_lint import check_charts
    check_charts(prs, issues)   # issues = {"error": [...], "warn": [...]}
"""
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from pptx.oxml.ns import qn
from qa_check import CHART_UNREADABLE, iter_shapes, safe_chart

# ── chart-family tags (localnames of c:plotArea children) ───────────────────
BAR_TAGS = ("barChart", "bar3DChart")
LINE_TAGS = ("lineChart", "line3DChart", "stockChart")
AREA_TAGS = ("areaChart", "area3DChart")
PIE_TAGS = ("pieChart", "pie3DChart", "doughnutChart", "ofPieChart")
# XY plots: both axes are value axes by construction, not a dual-scale chart.
XY_TAGS = ("scatterChart", "bubbleChart")
# rule 7(b) — charts that owe the reader a unit on the value axis
UNIT_BEARING_TAGS = BAR_TAGS + LINE_TAGS + AREA_TAGS

# ── thresholds ──────────────────────────────────────────────────────────────
MAX_PIE_SLICES = 7           # more slices than the eye can rank
PIE_WHOLE = 100.0            # a %-labelled pie must sum to the whole
PIE_SUM_TOL = 0.5            # rounding slack on that sum
MIN_SERIES_FOR_LEGEND = 2    # this many series without a legend = unreadable
MAX_LABELLED_LINE_POINTS = 8  # label start/inflection/end beyond this
MIN_CATS_FOR_INTERVAL = 3    # two points are always "uniform"
GAP_TOL = 1e-6               # float slack when comparing category gaps
MAX_HIGHLIGHT_COLORS = 1     # one highlight colour per exhibit

# ── text patterns ───────────────────────────────────────────────────────────
# a series name carrying any of these reads as carrying its unit
UNIT_HINT_RX = re.compile(r"\(|\$|€|£|%|bps|\bM\b|\bB\b|\bK\b|x\b")
# Date-like category labels: a 4-digit calendar year, FYnn/FYnnnn, or those
# forms behind a Qn prefix.  A *bare* two-digit number is deliberately NOT a
# year ("25" on an axis is a value), and a bare 4-digit number only counts
# inside YEAR_MIN..YEAR_MAX ("3000" is a value, not a year).
YEAR_RX = re.compile(r"^(FY)?(\d{2}|\d{4})$", re.IGNORECASE)
QUARTER_RX = re.compile(r"^Q([1-4])\s?(FY)?(\d{2}|\d{4})$", re.IGNORECASE)
TWO_DIGIT_EPOCH = 2000       # "FY24" -> 2024
YEAR_MIN = 1900              # plausible calendar-year window for a bare 4-digit
YEAR_MAX = 2100              # label; outside it the label is a value, not a year
QUARTERS_PER_YEAR = 4
INDEX_WORD = "index"
INDEX_NOTE = "indexed to 100"
AXIS_NOTE_WORD = "axis"
AXIS_NOTE_QUALIFIERS = ("truncated", "starts at", "from")
INDEXED_WORD = "indexed"
_HEX6 = re.compile(r"^[0-9A-Fa-f]{6}$")


# ── XML helpers ─────────────────────────────────────────────────────────────
def _local(tag):
    """Local name of a (possibly namespaced) lxml tag."""
    return tag.rsplit("}", 1)[-1] if isinstance(tag, str) else ""


def _val(el, default=None):
    """Numeric @val of an element, or `default` when absent/unparsable."""
    if el is None:
        return default
    raw = el.get("val")
    if raw is None:
        return default
    try:
        return float(raw)
    except (TypeError, ValueError):
        return default


def _flag(el):
    """Boolean @val of an element ('1'/'true' = True)."""
    if el is None:
        return False
    return (el.get("val") or "").strip().lower() in ("1", "true")


def _solid_fill_hex(spPr):
    """Hex of a <a:solidFill><a:srgbClr> under `spPr`, else None."""
    if spPr is None:
        return None
    fill = spPr.find(qn("a:solidFill"))
    if fill is None:
        return None
    clr = fill.find(qn("a:srgbClr"))
    if clr is None:
        return None
    value = (clr.get("val") or "").upper()
    return value if _HEX6.match(value) else None


def _element_text(el):
    """Concatenated <a:t> text under an element, or None."""
    if el is None:
        return None
    parts = [t.text or "" for t in el.iter(qn("a:t"))]
    text = "".join(parts).strip()
    return text or None


def _cache_points(holder):
    """Ordered <c:pt> values under a <c:cat>/<c:val> holder."""
    if holder is None:
        return []
    points = {}
    for pt in holder.iter(qn("c:pt")):
        try:
            idx = int(pt.get("idx", "0"))
        except (TypeError, ValueError):
            continue
        node = pt.find(qn("c:v"))
        points[idx] = node.text if node is not None else None
    return [points[k] for k in sorted(points)]


def _to_float(text):
    try:
        return float(text)
    except (TypeError, ValueError):
        return None


# ── fact extraction ─────────────────────────────────────────────────────────
def _plot_elements(plot_area):
    if plot_area is None:
        return []
    return [child for child in plot_area
            if _local(child.tag).endswith("Chart")]


def _chart_type_name(chart, plot_elements):
    try:
        ctype = chart.chart_type
    except (ValueError, KeyError, AttributeError, NotImplementedError):
        ctype = None
    if ctype is not None:
        name = getattr(ctype, "name", None)
        if name:
            return str(name)
    return _local(plot_elements[0].tag) if plot_elements else ""


def _api_series(chart):
    try:
        return list(chart.series)
    except (ValueError, KeyError, AttributeError, TypeError):
        return []


def _series_values(api_series, index, ser_el):
    """Values via the python-pptx API, falling back to the numCache XML
    (c:val, or c:yVal on an XY/bubble series)."""
    if index < len(api_series):
        try:
            return [None if v is None else float(v)
                    for v in api_series[index].values]
        except (ValueError, TypeError, AttributeError, KeyError):
            pass
    holder = ser_el.find(qn("c:val"))
    if holder is None:
        holder = ser_el.find(qn("c:yVal"))
    return [_to_float(v) for v in _cache_points(holder)]


def _series_name(api_series, index, ser_el):
    if index < len(api_series):
        try:
            name = api_series[index].name
            if name:
                return str(name)
        except (ValueError, TypeError, AttributeError, KeyError):
            pass
    tx = ser_el.find(qn("c:tx"))
    if tx is None:
        return ""
    values = _cache_points(tx)
    if values:
        return values[0] or ""
    node = tx.find(qn("c:v"))
    return (node.text or "") if node is not None else ""


def _shows_value(el):
    """True when the element carries a <c:dLbls> with showVal="1"."""
    if el is None:
        return False
    dlbls = el.find(qn("c:dLbls"))
    return dlbls is not None and _flag(dlbls.find(qn("c:showVal")))


def _series_facts(chart, plot_elements):
    api = _api_series(chart)
    facts, index = [], 0
    for plot_el in plot_elements:
        tag = _local(plot_el.tag)
        for ser_el in plot_el.findall(qn("c:ser")):
            facts.append({
                "name": _series_name(api, index, ser_el),
                "values": _series_values(api, index, ser_el),
                "plot_tag": tag,
                "fill_color": _solid_fill_hex(ser_el.find(qn("c:spPr"))),
                "show_val": _shows_value(ser_el) or _shows_value(plot_el),
            })
            index += 1
    return facts


def _categories(chart, plot_elements):
    try:
        cats = list(chart.plots[0].categories)
        if cats:
            return ["" if c is None else str(c) for c in cats]
    except (ValueError, TypeError, AttributeError, IndexError, KeyError):
        pass
    for plot_el in plot_elements:
        for ser_el in plot_el.findall(qn("c:ser")):
            cats = _cache_points(ser_el.find(qn("c:cat")))
            if cats:
                return ["" if c is None else str(c) for c in cats]
    return []


def _point_fill_colors(plot_elements):
    colors = []
    for plot_el in plot_elements:
        for dpt in plot_el.iter(qn("c:dPt")):
            hexval = _solid_fill_hex(dpt.find(qn("c:spPr")))
            if hexval:
                colors.append(hexval)
    return colors


def _axis_title(axis_el):
    if axis_el is None:
        return None
    return _element_text(axis_el.find(qn("c:title")))


def chart_facts(chart):
    """Parse one native chart into the plain fact dict documented above."""
    chart_space = chart._chartSpace
    chart_el = chart_space.find(qn("c:chart"))
    plot_area = chart_el.find(qn("c:plotArea")) if chart_el is not None else None
    plot_elements = _plot_elements(plot_area)
    val_axes = plot_area.findall(qn("c:valAx")) if plot_area is not None else []
    cat_axes = plot_area.findall(qn("c:catAx")) if plot_area is not None else []

    vmin = vmax = None
    if val_axes:
        scaling = val_axes[0].find(qn("c:scaling"))
        if scaling is not None:
            vmin = _val(scaling.find(qn("c:min")))
            vmax = _val(scaling.find(qn("c:max")))

    return {
        "type": _chart_type_name(chart, plot_elements),
        "categories": _categories(chart, plot_elements),
        "series": _series_facts(chart, plot_elements),
        "val_axis_min": vmin,
        "val_axis_max": vmax,
        "has_legend": chart_el is not None
        and chart_el.find(qn("c:legend")) is not None,
        "n_val_axes": len(val_axes),
        "is_3d": chart_el is not None
        and chart_el.find(qn("c:view3D")) is not None,
        "point_fill_colors": _point_fill_colors(plot_elements),
        "show_val_labels": any(_flag(el)
                               for el in chart_space.iter(qn("c:showVal"))),
        "val_axis_title": _axis_title(val_axes[0] if val_axes else None),
        "cat_axis_title": _axis_title(cat_axes[0] if cat_axes else None),
    }


# ── shared derivations ──────────────────────────────────────────────────────
def _plot_tags(facts):
    return {s["plot_tag"] for s in facts["series"]}


def _has_tag(facts, tags):
    return bool(_plot_tags(facts) & set(tags))


def _numeric_values(facts):
    return [v for s in facts["series"] for v in s["values"] if v is not None]


def _slide_text(slide):
    """All text-frame text on the slide, casefolded and space-joined."""
    parts = []
    for shape in iter_shapes(slide.shapes):
        if getattr(shape, "has_text_frame", False):
            parts.append(shape.text_frame.text)
    return " ".join(parts).casefold()


def _chart_label(shape, facts):
    name = (getattr(shape, "name", "") or "").strip()
    if name:
        return name
    for series in facts["series"]:
        if series["name"]:
            return series["name"]
    return "untitled"


def _parse_year(text):
    """Year (fractional for quarters) for a date-like category label, else None.

    Only the forms in YEAR_RX/QUARTER_RX count, so plain numeric axis labels
    ("100", "75", "50", "25") are never mistaken for a date sequence.
    """
    label = (text or "").strip()
    match = QUARTER_RX.match(label)
    if match:
        year = _year_value(match.group(2), match.group(3))
        if year is None:
            return None
        return year + int(match.group(1)) / QUARTERS_PER_YEAR
    match = YEAR_RX.match(label)
    if match:
        year = _year_value(match.group(1), match.group(2))
        return None if year is None else float(year)
    return None


def _year_value(fy_prefix, digits):
    """Calendar year for a label's digits, or None when it isn't a year.

    Two digits are a year only behind an "FY" prefix (FY25 -> 2025); four
    digits only inside the plausible YEAR_MIN..YEAR_MAX window.
    """
    if len(digits) == 2:
        return TWO_DIGIT_EPOCH + int(digits) if fy_prefix else None
    year = int(digits)
    return year if YEAR_MIN <= year <= YEAR_MAX else None


# ── rules ───────────────────────────────────────────────────────────────────
def _rule_axis_baseline(prefix, facts, text, issues):
    """1. bar baseline must be zero; 2. truncated line/area needs a note."""
    vmin = facts["val_axis_min"]
    if vmin is None or vmin <= 0:
        return
    if _has_tag(facts, BAR_TAGS):
        issues["error"].append(
            f"{prefix} bar axis not at zero (value-axis minimum {vmin:g}) — "
            "bars must start at zero")
        return
    if not _has_tag(facts, LINE_TAGS + AREA_TAGS):
        return
    noted = (AXIS_NOTE_WORD in text
             and any(q in text for q in AXIS_NOTE_QUALIFIERS))
    if noted or INDEXED_WORD in text:
        return
    issues["warn"].append(
        f"{prefix} truncated line axis (minimum {vmin:g}) — add an "
        "axis-range note")


def _rule_dual_axis(prefix, facts, issues):
    """3. two value axes = two scales the reader cannot reconcile. An XY
    plot (scatter/bubble) has a value axis on each side by construction —
    its two axes are X and Y, not two scales — so it is exempt."""
    if _has_tag(facts, XY_TAGS) and facts["n_val_axes"] == 2:
        return
    if facts["n_val_axes"] >= 2:
        issues["warn"].append(
            f"{prefix} dual axis ({facts['n_val_axes']} value axes) — "
            "readers cannot compare two scales; split into two exhibits")


def _rule_three_d(prefix, facts, issues):
    """4. 3-D perspective distorts every value it draws."""
    if facts["is_3d"]:
        issues["error"].append(
            f"{prefix} 3-D chart (c:view3D) — perspective distorts values; "
            "use a flat chart")


def _rule_pie(prefix, facts, issues):
    """5. pie/doughnut must be part-to-whole with at most 7 slices."""
    if not _has_tag(facts, PIE_TAGS):
        return
    values = _numeric_values(facts)
    labels = " ".join(facts["categories"]
                      + [s["name"] for s in facts["series"]])
    reasons = []
    if "%" in labels and values:
        total = sum(values)
        if abs(total - PIE_WHOLE) > PIE_SUM_TOL:
            reasons.append(f"%-labelled slices sum to {total:g}, not 100")
    if any(v < 0 for v in values):
        reasons.append("negative values")
    if len(facts["categories"]) > MAX_PIE_SLICES:
        reasons.append(f"{len(facts['categories'])} slices")
    if reasons:
        issues["warn"].append(
            f"{prefix} pie must be part-to-whole / ≤{MAX_PIE_SLICES} slices "
            f"({'; '.join(reasons)})")


def _rule_category_intervals(prefix, facts, issues):
    """6. date categories must run in order at a uniform interval.

    "In order" means strictly monotonic in *either* direction: descending
    years are standard on horizontal bars, where the newest period sits at the
    top. Only a sequence that changes direction (or repeats a year) is an
    error; the interval-uniformity warning is computed on absolute gaps so it
    applies to both directions.
    """
    cats = facts["categories"]
    if len(cats) < 2:
        return
    years = [_parse_year(c) for c in cats]
    if any(y is None for y in years):
        return
    shown = ", ".join(cats)
    ascending = all(b > a for a, b in zip(years, years[1:]))
    descending = all(b < a for a, b in zip(years, years[1:]))
    if not (ascending or descending):
        issues["error"].append(f"{prefix} out-of-order dates ({shown})")
        return
    if len(years) < MIN_CATS_FOR_INTERVAL:
        return
    gaps = [abs(b - a) for a, b in zip(years, years[1:])]
    if max(gaps) - min(gaps) > GAP_TOL:
        issues["warn"].append(
            f"{prefix} inconsistent tick intervals ({shown}) — equal gaps or "
            "an explicit break")


def _rule_legend_and_units(prefix, facts, issues):
    """7. label the series (legend) and the value axis (units)."""
    series = facts["series"]
    if len(series) >= MIN_SERIES_FOR_LEGEND and not facts["has_legend"]:
        issues["warn"].append(
            f"{prefix} {len(series)} series and no legend — readers cannot "
            "tell the series apart")
    if not _has_tag(facts, UNIT_BEARING_TAGS):
        return
    if facts["val_axis_title"] or facts["show_val_labels"]:
        return
    if any(UNIT_HINT_RX.search(s["name"] or "") for s in series):
        return
    issues["warn"].append(
        f"{prefix} value axis has no units — add an axis title, a unit in the "
        "series name, or data labels")


def _rule_indexed(prefix, facts, text, issues):
    """8. an indexed series is meaningless without its base."""
    labels = [s["name"] or "" for s in facts["series"]] + facts["categories"]
    if not any(INDEX_WORD in label.casefold() for label in labels):
        return
    if INDEX_NOTE in text:
        return
    issues["warn"].append(
        f"{prefix} indexed series without an '{INDEX_NOTE}' note on the slide")


def _rule_highlight_colors(prefix, facts, issues):
    """9. one highlight colour; pie slices are colour-coded by design."""
    if _has_tag(facts, PIE_TAGS):
        return
    series_fills = {s["fill_color"] for s in facts["series"] if s["fill_color"]}
    highlights = sorted({c for c in facts["point_fill_colors"]
                         if c not in series_fills})
    if len(highlights) > MAX_HIGHLIGHT_COLORS:
        issues["warn"].append(
            f"{prefix} more than one highlight colour "
            f"({', '.join(highlights)}) — highlight a single point")


def _rule_labelled_points(prefix, facts, issues):
    """10. a long line series labelled at every point is unreadable."""
    worst = 0
    for series in facts["series"]:
        if series["plot_tag"] not in LINE_TAGS or not series["show_val"]:
            continue
        worst = max(worst, len(series["values"]))
    if worst > MAX_LABELLED_LINE_POINTS:
        issues["warn"].append(
            f"{prefix} data labels on all {worst} line points — label "
            "start/inflection/end only")


def _unreadable(slide_no, shape, exc, issues):
    """Degrade an unparsable chart part to a warning instead of a traceback."""
    label = (getattr(shape, "name", "") or "").strip() or "untitled"
    issues["warn"].append(
        f"Slide {slide_no}: chart '{label}' could not be analysed ({exc})")


def check_chart(slide_no, shape, chart, slide_text, issues):
    """Run every honesty rule against one chart; append to `issues`.

    Returns the parsed facts, or None when the chart part could not be read —
    one malformed chart must not abort the whole lint run.
    """
    try:
        facts = chart_facts(chart)
    except Exception as exc:  # defensive boundary: any malformed chart part
        _unreadable(slide_no, shape, exc, issues)
        return None
    prefix = f"Slide {slide_no}: chart '{_chart_label(shape, facts)}' —"
    _rule_axis_baseline(prefix, facts, slide_text, issues)
    _rule_dual_axis(prefix, facts, issues)
    _rule_three_d(prefix, facts, issues)
    _rule_pie(prefix, facts, issues)
    _rule_category_intervals(prefix, facts, issues)
    _rule_legend_and_units(prefix, facts, issues)
    _rule_indexed(prefix, facts, slide_text, issues)
    _rule_highlight_colors(prefix, facts, issues)
    _rule_labelled_points(prefix, facts, issues)
    return facts


def check_charts(prs, issues):
    """Chart-honesty lint for every native chart in the presentation."""
    for slide_no, slide in enumerate(prs.slides, 1):
        text = None
        for shape in iter_shapes(slide.shapes):
            if not getattr(shape, "has_chart", False):
                continue
            if text is None:
                text = _slide_text(slide)
            chart = safe_chart(shape)
            if chart is None:      # unreachable part / broken relationship
                issues["warn"].append(f"Slide {slide_no}: {CHART_UNREADABLE}")
                continue
            check_chart(slide_no, shape, chart, text, issues)
