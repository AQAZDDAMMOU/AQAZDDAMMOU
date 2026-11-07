#!/usr/bin/env python3
"""
chart_workbook.py — embedded-workbook round-trip and cache-vs-workbook drift
for native pptx charts.

A native chart carries its data twice: the rendered cache baked into the
chart XML (``c:cat/c:strCache``, ``c:val/c:numCache`` — what PowerPoint
paints without touching the linked file) and an embedded workbook
(``ppt/embeddings/*.xlsx``) linked from the chart part via
``c:externalData r:id`` — the sheet PowerPoint opens on "Edit Data". An
analyst who edits the sheet and forgets to refresh the chart (or any build
step that only ever writes the cache) leaves the two out of sync.

This module reads both sides and reports the drift:

    cache_series(chart_part)          -> {"categories", "series"} from the XML cache
    embedded_workbook(chart_part)     -> the linked .xlsx blob, or None
    read_workbook_series(xlsx_bytes)  -> same shape, read via openpyxl
    drift(cache, workbook)            -> human-readable mismatch messages
    deck_drift(prs)                   -> {slide position: [messages]}

``deck_to_outline.extract_chart`` prefers the workbook's values when reading
a chart back out of a deck — the sheet carries full float precision, while
the cache is whatever was last rendered, which a hand edit can leave stale.
``diff_deck.diff_outline`` reports drift as a warning (an error under
``--strict``) via its ``check_workbook_drift`` step.

openpyxl is an optional dependency: with it unavailable, ``read_workbook_
series`` returns None and every check downstream degrades to "no workbook
to compare against" rather than raising.

python-pptx's own workbook layout (``chart.xlsx.CategoryWorkbookWriter``):
sheet 1, categories in column A from row 2, series in columns B... with
series names in row 1 — the layout ``read_workbook_series`` assumes. An XY
(scatter) chart is the one exception on both sides: its cache carries
``c:xVal``/``c:yVal`` (the x values stand in as the categories here) and
its workbook (``XyWorkbookWriter``) stacks one name-row/x-y block per
series down columns A:B, which ``read_workbook_series`` detects.

The chart-XML cache reader below is a local, minimal re-implementation of
the pattern chart_lint.py uses internally (``_cache_points`` et al.) —
chart_lint's helpers are private to that module and are not imported here.
"""
import sys
import zipfile
from io import BytesIO
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

from lxml import etree  # noqa: E402
from pptx.oxml.ns import qn  # noqa: E402

from qa_check import iter_shapes, safe_chart  # noqa: E402

# What a malformed chart part or a hand-stripped relationship graph raises.
# AttributeError is kept here (not treated as a programming error) because
# python-pptx's own relationship walk raises it on a malformed part — so a
# typo'd attribute in this module would be masked by the same catch. Named
# concretely so an actual programming error (e.g. TypeError) still surfaces
# instead of degrading to a silent None.
PART_ERRORS = (KeyError, ValueError, AttributeError, etree.XMLSyntaxError,
               OSError)
# ...plus what an unreadable .xlsx blob raises on the way through openpyxl.
WORKBOOK_ERRORS = PART_ERRORS + (zipfile.BadZipFile,)


# ── chart-XML cache reading ──────────────────────────────────────────────────

def _cache_points(holder):
    """Ordered <c:pt> values under a <c:cat>/<c:val> holder, indexed by
    their @idx (missing indices leave gaps; python-pptx never writes those,
    but a hand-edited chart part could)."""
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


def _series_name(ser_el, index):
    """A series' display name from <c:tx>, or a positional fallback."""
    tx = ser_el.find(qn("c:tx"))
    if tx is not None:
        pts = _cache_points(tx)
        if pts and pts[0]:
            return pts[0]
        node = tx.find(qn("c:v"))
        if node is not None and node.text:
            return node.text
    return f"Series {index + 1}"


def cache_series(chart_part) -> dict:
    """{"categories": [str], "series": {name: [float|None]}} read from the
    chart XML's rendered cache (c:cat/c:strCache, c:val/c:numCache) — what
    PowerPoint paints without touching the linked workbook.

    `chart_part` is a ChartPart (e.g. `shape.chart.part`); its `_element`
    *is* the <c:chartSpace> root (python-pptx's Chart._chartSpace is the
    same object)."""
    chart_space = chart_part._element
    chart_el = chart_space.find(qn("c:chart"))
    plot_area = chart_el.find(qn("c:plotArea")) if chart_el is not None else None
    categories = []
    series = {}
    if plot_area is not None:
        index = 0
        for plot_el in plot_area:
            tag = plot_el.tag
            local = tag.rsplit("}", 1)[-1] if isinstance(tag, str) else ""
            if not local.endswith("Chart"):
                continue
            for ser_el in plot_el.findall(qn("c:ser")):
                # XY (scatter) series carry c:xVal/c:yVal instead of
                # c:cat/c:val: the x values stand in as the categories.
                if not categories:
                    cat_el = _first(ser_el, "c:cat", "c:xVal")
                    points = _cache_points(cat_el)
                    if cat_el is not None and cat_el.tag == qn("c:xVal"):
                        # numeric x values are spelled like the workbook
                        # reader spells them ('12', not '12.0')
                        categories = [_x_label(c) for c in points]
                    else:
                        # c:cat strings stay verbatim: '01', '1.0' and
                        # '007' are labels, not numbers to canonicalise
                        categories = [c or "" for c in points]
                name = _series_name(ser_el, index)
                values = [_to_float(v) for v in
                          _cache_points(_first(ser_el, "c:val", "c:yVal"))]
                series[name] = values
                index += 1
    return {"categories": categories, "series": series}


def _first(ser_el, *tags):
    """The first of `tags` present under a <c:ser>, or None."""
    for tag in tags:
        el = ser_el.find(qn(tag))
        if el is not None:
            return el
    return None


def _x_label(text):
    """A cached category/x value as a label: numeric x values spell like
    the workbook reader spells them ('12', not '12.0'), so the two sides
    compare by label."""
    if text is None:
        return ""
    number = _to_float(text)
    if number is None:
        return text
    return f"{number:g}"


# ── embedded workbook ────────────────────────────────────────────────────────

def embedded_workbook(chart_part) -> bytes | None:
    """The linked .xlsx blob for chart_part's <c:externalData r:id>, or None
    when the chart carries no linked workbook (or the relationship cannot
    be resolved — a dangling rId, a hand-stripped part).

    Delegates to python-pptx's ChartWorkbook (chart_part.chart_workbook),
    which itself resolves c:externalData's r:id through the part's OPC
    relationships to the EmbeddedXlsxPart — so this *is* that rel-walk, via
    the public API rather than re-deriving it over raw lxml."""
    try:
        xlsx_part = chart_part.chart_workbook.xlsx_part
    except PART_ERRORS:  # a malformed part/relationship graph
        return None
    return xlsx_part.blob if xlsx_part is not None else None


def read_workbook_series(xlsx_bytes) -> dict | None:
    """Same shape as cache_series(), read from an embedded .xlsx blob via
    openpyxl. None when openpyxl is not installed (an optional dependency),
    `xlsx_bytes` is None, or the bytes are not a readable workbook.

    Layout python-pptx writes (CategoryWorkbookWriter): sheet 1, categories
    in column A from row 2, series in columns B... with series names in
    row 1.
    """
    if xlsx_bytes is None:
        return None
    try:
        import openpyxl
    except ImportError:
        return None
    try:
        from openpyxl.utils.exceptions import InvalidFileException
    except ImportError:                       # pragma: no cover - very old
        blob_errors = WORKBOOK_ERRORS
    else:
        blob_errors = WORKBOOK_ERRORS + (InvalidFileException,)
    try:
        wb = openpyxl.load_workbook(BytesIO(xlsx_bytes), read_only=True,
                                    data_only=True)
        rows = list(wb.worksheets[0].iter_rows(values_only=True))
    # ...plus IndexError: a readable workbook with no worksheets at all.
    except blob_errors + (IndexError,):  # corrupt/unreadable workbook part
        return None
    if not rows:
        return {"categories": [], "series": {}}
    if _is_xy_layout(rows):
        return _read_xy_rows(rows)
    header = rows[0]
    names = [str(h) if h is not None else f"Series {i}"
            for i, h in enumerate(header[1:], 1)]
    categories = []
    series = {name: [] for name in names}
    for row in rows[1:]:
        if not row or row[0] is None:
            continue
        categories.append(_cell_label(row[0]))
        for i, name in enumerate(names, 1):
            v = row[i] if i < len(row) else None
            series[name].append(float(v) if isinstance(v, (int, float))
                                else None)
    return {"categories": categories, "series": series}


def _cell_label(value):
    """A workbook cell as a category label; numbers spell like the chart
    cache's x values ('2019', not '2019.0')."""
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, (int, float)):
        return f"{value:g}"
    return str(value)


def _name_row(row):
    """True for a python-pptx XY series header row: column A empty,
    column B the series name, nothing else — or a bubble block's header,
    which additionally carries "Size" in column C (BubbleWorkbookWriter)."""
    if not (bool(row) and row[0] is None and len(row) >= 2
            and isinstance(row[1], str) and row[1] != ""):
        return False
    rest = [c for c in row[2:] if c is not None]
    return rest == [] or rest == ["Size"]


def _is_xy_layout(rows):
    """python-pptx's XyWorkbookWriter stacks one (name row, x/y rows,
    blank row) block per series down columns A:B; a second name row below
    the first data row is what tells it apart from the category layout
    (one header row, then label/values rows). A bubble workbook announces
    itself on row 1 already — its header carries the literal "Size" column
    (x/y/size blocks; a single-series scatter block reads identically in
    either layout, so it needs no row-1 check)."""
    if any(_name_row(row) for row in rows[1:]):
        return True
    first = rows[0]
    if not (_name_row(first) and len(first) >= 3 and first[2] == "Size"):
        return False
    # a category workbook whose second series is literally named "Size"
    # looks identical on row 1 — its data rows carry string labels in
    # column A, a bubble block's x values never do
    nxt = next((r for r in rows[1:] if any(c is not None for c in r)), None)
    return nxt is None or not isinstance(nxt[0], str)


def _read_xy_rows(rows):
    """XY layout -> {"categories": [x labels of the first series],
    "series": {name: [y, ...]}}. Every series is built over the first
    series' x values (the builder writes one shared X column), so only
    the y side of a later block is compared."""
    categories, series, name = [], {}, None
    for row in rows:
        if _name_row(row):
            name = row[1]
            series.setdefault(name, [])
            continue
        if name is None or not row or row[0] is None:
            continue
        if len(series) == 1:
            categories.append(_cell_label(row[0]))
        v = row[1] if len(row) > 1 else None
        series[name].append(float(v) if isinstance(v, (int, float)) else None)
    return {"categories": categories, "series": series}


# ── drift ─────────────────────────────────────────────────────────────────

def _label_index(cats):
    """First-occurrence {label: index} for `cats`, plus the set of labels
    that repeat (a hand-edited sheet can duplicate a category row)."""
    index, dupes = {}, set()
    for i, c in enumerate(cats):
        if c in index:
            dupes.add(c)
        else:
            index[c] = i
    return index, dupes


def _point_message(name, cat, cv, wv, tol):
    """One drift message for a single (series, category) cell, or None
    when the two sides agree. A value missing on either side compares as
    "unpaired" rather than a numeric difference."""
    if cv is None or wv is None:
        if cv is None and wv is None:
            return None
        cv_s = "unpaired" if cv is None else cv
        wv_s = "unpaired" if wv is None else wv
        return f"Series {name!r} @ {cat!r}: cache {cv_s} vs workbook {wv_s}"
    if abs(cv - wv) > tol:
        return f"Series {name!r} @ {cat!r}: cache {cv} vs workbook {wv}"
    return None


def _ordered_labels(primary, secondary):
    """`primary` labels in order, then any `secondary`-only labels
    appended in their own order — first occurrence of each wins."""
    seen, labels = set(), []
    for cats in (primary, secondary):
        for label in cats:
            if label not in seen:
                seen.add(label)
                labels.append(label)
    return labels


def drift(cache, workbook, tol=1e-9) -> list[str]:
    """Human-readable mismatches between the chart's rendered cache and its
    linked workbook. [] when there is nothing to compare (no workbook) or
    every value agrees within `tol`.

    Per-point comparison aligns by category *label* whenever both sides
    carry labels — a category inserted or reordered mid-sheet (not just
    appended at the tail) must not read as drift on every point after it.
    Only when one side has no category labels at all does comparison fall
    back to position. A value missing on either side (a length mismatch,
    a genuinely empty cell, or a label present on only one side) compares
    as "unpaired" rather than a numeric difference. A duplicated category
    label is reported once as "duplicate category 'X'" (first occurrence
    of the label wins the alignment)."""
    if workbook is None:
        return []
    messages = []
    cache_cats = cache.get("categories") or []
    wb_cats = workbook.get("categories") or []
    if cache_cats != wb_cats:
        messages.append(
            f"categories differ: cache {cache_cats!r} vs workbook {wb_cats!r}")

    cache_idx, cache_dupes = _label_index(cache_cats)
    wb_idx, wb_dupes = _label_index(wb_cats)
    use_labels = bool(cache_cats) and bool(wb_cats)
    labels = []
    if use_labels:
        for label in sorted(cache_dupes | wb_dupes):
            messages.append(f"duplicate category {label!r}")
        primary, secondary = (wb_cats, cache_cats) \
            if len(wb_cats) >= len(cache_cats) else (cache_cats, wb_cats)
        labels = _ordered_labels(primary, secondary)

    cache_series_map = cache.get("series") or {}
    wb_series_map = workbook.get("series") or {}
    names = list(cache_series_map) + [n for n in wb_series_map
                                      if n not in cache_series_map]

    for name in names:
        cache_vals = cache_series_map.get(name)
        wb_vals = wb_series_map.get(name)
        if cache_vals is None:
            messages.append(f"Series {name!r}: only in embedded workbook")
            continue
        if wb_vals is None:
            messages.append(f"Series {name!r}: missing from embedded workbook")
            continue
        if len(cache_vals) != len(wb_vals):
            messages.append(
                f"Series {name!r}: length differs (cache {len(cache_vals)} "
                f"points vs workbook {len(wb_vals)})")

        if use_labels:
            for label in labels:
                ci, wi = cache_idx.get(label), wb_idx.get(label)
                cv = cache_vals[ci] if ci is not None and ci < len(cache_vals) else None
                wv = wb_vals[wi] if wi is not None and wi < len(wb_vals) else None
                msg = _point_message(name, label, cv, wv, tol)
                if msg:
                    messages.append(msg)
            continue

        length = max(len(cache_vals), len(wb_vals))
        for i in range(length):
            if i < len(cache_cats):
                cat = cache_cats[i]
            elif i < len(wb_cats):
                cat = wb_cats[i]
            else:
                cat = f"#{i}"
            cv = cache_vals[i] if i < len(cache_vals) else None
            wv = wb_vals[i] if i < len(wb_vals) else None
            msg = _point_message(name, cat, cv, wv, tol)
            if msg:
                messages.append(msg)
    return messages


def deck_drift(prs) -> dict[int, list[str]]:
    """{slide position: [drift messages]} for every native chart in the
    deck that carries a readable embedded workbook. A slide with no drift
    (or with charts that have no workbook to compare against — no linked
    workbook, or openpyxl unavailable) is simply absent from the dict."""
    result = {}
    for slide_no, slide in enumerate(prs.slides, 1):
        msgs = []
        for shape in iter_shapes(slide.shapes):
            chart = safe_chart(shape)
            if chart is None:
                continue
            chart_part = chart.part
            workbook = read_workbook_series(embedded_workbook(chart_part))
            msgs.extend(drift(cache_series(chart_part), workbook))
        if msgs:
            result[slide_no] = msgs
    return result
