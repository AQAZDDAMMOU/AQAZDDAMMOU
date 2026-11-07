#!/usr/bin/env python3
"""
deck_to_outline.py — recover a markdown outline from an existing .pptx.

The restyle/normalize path: extract an existing deck's content (text,
tables, native-chart data, images, speaker notes) into the outline format
build_deck.py consumes, so a messy or off-brand deck can be re-authored
onto this toolchain's layouts and palettes:

    python3 scripts/deck_to_outline.py old.pptx --output outline.md \\
        [--assets-dir assets] [--palette midnight-executive] \\
        [--report fidelity.json]
    # review/edit outline.md, then:
    python3 scripts/build_deck.py outline.md --output restyled.pptx

Extraction is lossy by design — the point is recovering *content*, not
cloning geometry. Everything that cannot be carried over (SmartArt,
unsupported chart types, freeform art) is listed per slide in the fidelity
report instead of being silently dropped; review it before delivering a
restyled deck. Embedded pictures are exported to
<assets-dir>/extracted/ and referenced from the outline by path.
"""
import argparse
import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn

import chart_workbook
from charts import CHART_TYPES
from qa_check import _slide_title, iter_shapes

EMU_IN = 914400
FOOTER_ZONE = 0.86          # top beyond this fraction of height = footer
TITLE_MIN_PT = 34           # a run this big on slide 1 = title-slide signal
SECTION_MAX_WORDS = 6       # heading-only slides this short = divider
FULL_IMAGE_COVER = 0.85     # picture covering this slide fraction = full-image
WORKBOOK_DIFF_TOL = 1e-9    # float slack when counting cache-vs-workbook diffs
PAGENUM_RX = re.compile(r"^(B·)?\d+$")
_DIAGRAM_NS = b"http://schemas.openxmlformats.org/drawingml/2006/diagram"

# XL_CHART_TYPE -> outline chart kind (later CHART_TYPES keys win, so the
# shared COLUMN_CLUSTERED enum maps back to 'column').
KIND_BY_XL = {xl: kind for kind, xl in CHART_TYPES.items()}


def _shape_box_in(shape):
    try:
        if None in (shape.left, shape.top, shape.width, shape.height):
            return None
        return (shape.left / EMU_IN, shape.top / EMU_IN,
                shape.width / EMU_IN, shape.height / EMU_IN)
    except (AttributeError, TypeError):
        return None


def _clean_label(text):
    """Chart/table labels must survive '- label: value' parsing, where the
    first colon splits label from value."""
    return str(text).replace(":", " —").strip() or "Item"


def _fmt_value(v):
    return f"{v:g}"


def extract_chart(shape, dropped):
    """Native chart -> {'kind', 'series' (names or None), 'data',
    'extra' (optional '- Key: value' outline lines)} or None."""
    chart = shape.chart
    tags = set(_plot_tags(chart))
    if "barChart" in tags and "lineChart" in tags:
        # chart_combo output: chart.chart_type only reads the first plot
        # group, so a combo would extract as its columns alone.
        return _extract_combo(chart, dropped)
    kind = KIND_BY_XL.get(chart.chart_type)
    if kind is None:
        name = getattr(chart.chart_type, "name", str(chart.chart_type))
        dropped.append(f"chart type {name} has no outline equivalent — "
                       "data not recovered")
        return None
    try:
        plot = chart.plots[0]
        categories = [str(c) if c is not None else "" for c in plot.categories]
        series_list = list(plot.series)
    except (IndexError, ValueError, AttributeError) as e:
        dropped.append(f"chart data unreadable ({e})")
        return None
    if not series_list:
        dropped.append("chart has no series — data not recovered")
        return None
    if kind == "scatter":
        return _extract_scatter(chart, series_list, dropped)
    if kind == "bubble":
        return _extract_bubble(chart, dropped)

    # The embedded workbook (if any) carries full float precision; the XML
    # cache is whatever was last rendered, which a hand edit in PowerPoint's
    # "Edit Data" sheet can leave stale. Prefer the workbook when its shape
    # (category/series counts) lines up with what the cache gave us —
    # series line up positionally, since both are written in ChartData's
    # add_series() order.
    wb_bytes = chart_workbook.embedded_workbook(chart.part)
    workbook = chart_workbook.read_workbook_series(wb_bytes)
    wb_series_values = None
    if workbook is not None:
        wb_cats = workbook.get("categories") or []
        wb_series = list(workbook.get("series", {}).values())
        if len(wb_cats) == len(categories) and len(wb_series) == len(series_list):
            wb_series_values = wb_series

    rows, none_seen, diff_count = [], False, 0
    for i in range(max(len(s.values) for s in series_list)):
        label = _clean_label(categories[i] if i < len(categories)
                             else f"Item {i + 1}")
        values = []
        for si, s in enumerate(series_list):
            cache_v = s.values[i] if i < len(s.values) else None
            wb_v = (wb_series_values[si][i]
                   if wb_series_values is not None and i < len(wb_series_values[si])
                   else None)
            if wb_v is not None and (cache_v is None
                                     or abs(wb_v - cache_v) > WORKBOOK_DIFF_TOL):
                diff_count += 1
            v = wb_v if wb_v is not None else cache_v
            if v is None:
                none_seen = True
                v = 0.0
            values.append(float(v))
        rows.append((label, values))
    if none_seen:
        dropped.append("chart had empty data points — recovered as 0")
    # Fidelity notes only when they carry information — not on every chart:
    # the workbook was used and actually differed from the cache somewhere,
    # or the workbook exists but couldn't be read/used at all (openpyxl
    # missing, corrupt part, or a shape that didn't line up with the cache).
    # Cache and workbook agreeing, or no embedded workbook to begin with,
    # says nothing worth a [DROPPED] line.
    if wb_series_values is not None:
        if diff_count:
            dropped.append("chart data: workbook (cache differed on "
                           f"{diff_count} point(s))")
    elif wb_bytes is not None:
        dropped.append("chart data: cache (embedded workbook unreadable)")
    names = None
    if len(series_list) > 1:
        names = [_clean_label(s.name or f"Series {i + 1}")
                 for i, s in enumerate(series_list)]
    return {"kind": kind, "series": names, "data": rows}


def _plot_tags(chart):
    """Local tag names of the plot-group children of c:plotArea, in
    document order ('barChart', 'lineChart', ...)."""
    chart_el = chart._chartSpace.find(qn("c:chart"))
    plot_area = chart_el.find(qn("c:plotArea")) if chart_el is not None else None
    if plot_area is None:
        return []
    return [el.tag.rsplit("}", 1)[-1] for el in plot_area
            if isinstance(el.tag, str)
            and el.tag.rsplit("}", 1)[-1].endswith("Chart")]


def _cat_cache(ser_el):
    """Cached category label texts of one <c:ser>, by point index."""
    holder = ser_el.find(qn("c:cat"))
    if holder is None:
        return []
    points = {}
    for pt in holder.iter(qn("c:pt")):
        node = pt.find(qn("c:v"))
        try:
            points[int(pt.get("idx", "0"))] = \
                node.text if node is not None else ""
        except (TypeError, ValueError):
            continue
    return [points[k] or "" for k in sorted(points)]


def _extract_combo(chart, dropped):
    """bar+line combo (chart_combo.py) -> the union of both plot groups'
    series over the column group's categories, plus the '- Line-Series:'
    outline line that reproduces the split."""
    chart_el = chart._chartSpace.find(qn("c:chart"))
    plot_area = chart_el.find(qn("c:plotArea"))
    series, categories, line_names = [], [], []
    for plot_el in plot_area:
        tag = (plot_el.tag.rsplit("}", 1)[-1]
               if isinstance(plot_el.tag, str) else "")
        if not tag.endswith("Chart"):
            continue
        for ser_el in plot_el.findall(qn("c:ser")):
            tx = ser_el.find(qn("c:tx"))
            node = next(tx.iter(qn("c:v")), None) if tx is not None else None
            name = _clean_label((node.text if node is not None else "")
                                or f"Series {len(series) + 1}")
            # keyed by c:pt idx so a sparse cache cannot misalign values
            by_idx = {}
            val = ser_el.find(qn("c:val"))
            if val is not None:
                for pt in val.iter(qn("c:pt")):
                    v = pt.find(qn("c:v"))
                    try:
                        by_idx[int(pt.get("idx", len(by_idx)))] = \
                            float(v.text)
                    except (TypeError, ValueError, AttributeError):
                        by_idx[int(pt.get("idx", len(by_idx)))] = 0.0
            values = [by_idx.get(i, 0.0)
                      for i in range(max(by_idx, default=-1) + 1)]
            if not categories:
                categories = _cat_cache(ser_el)
            series.append((name, values))
            if tag == "lineChart":
                line_names.append(name)
    if not series or not any(vals for _, vals in series):
        dropped.append("combo chart has no cached values — data not "
                       "recovered")
        return None
    n = max(len(vals) for _, vals in series)
    rows = []
    for i in range(n):
        label = _clean_label(categories[i] if i < len(categories)
                             and categories[i] else f"Item {i + 1}")
        rows.append((label, [vals[i] if i < len(vals) else 0.0
                             for _, vals in series]))
    extra = ([f"- Line-Series: {', '.join(line_names)}"]
             if line_names else [])
    return {"kind": "combo", "series": [name for name, _ in series],
            "data": rows, "extra": extra}


def _extract_bubble(chart, dropped):
    """Bubble chart -> '**Series:** X, Y, Size' with '- label: x, y, size'
    rows; 2+ series read as [Group] buckets ('- Item N [group]: …')."""
    ser_els = list(chart._chartSpace.iter(qn("c:ser")))
    multi = len(ser_els) > 1
    rows, idx = [], 1
    for si, el in enumerate(ser_els):
        xs = _xy_cache(el, "c:xVal")
        ys = _xy_cache(el, "c:yVal")
        sizes = _xy_cache(el, "c:bubbleSize")
        if not xs or not ys:
            continue
        if len(ys) < len(xs) or len(sizes) < len(xs):
            dropped.append("bubble series with fewer y/size points than x "
                           "values — missing values recovered as 0")
        tx = el.find(qn("c:tx"))
        node = next(tx.iter(qn("c:v")), None) if tx is not None else None
        group = _clean_label((node.text if node is not None else "")
                             or f"Group {si + 1}")
        for i, x in enumerate(xs):
            label = f"Item {idx}" + (f" [{group}]" if multi else "")
            rows.append((label, [x, ys[i] if i < len(ys) else 0.0,
                                 sizes[i] if i < len(sizes) else 0.0]))
            idx += 1
    if not rows:
        dropped.append("bubble chart has no cached x/y values — data not "
                       "recovered")
        return None
    return {"kind": "bubble", "series": ["X", "Y", "Size"], "data": rows}


def _xy_cache(ser_el, tag):
    """Cached <c:xVal>/<c:yVal> numbers of one XY series, by point index."""
    holder = ser_el.find(qn(tag))
    if holder is None:
        return []
    points = {}
    for pt in holder.iter(qn("c:pt")):
        node = pt.find(qn("c:v"))
        try:
            points[int(pt.get("idx", "0"))] = float(node.text)
        except (TypeError, ValueError, AttributeError):
            continue
    return [points[k] for k in sorted(points)]


def _extract_scatter(chart, series_list, dropped):
    """XY scatter -> the outline's scatter contract (generation-guide):
    one series -> '- x: y' rows; several -> '**Series:** X, Y1, ...' with
    '- Item N: x, y1, ...' rows over the first series' x values (the
    builder writes one shared X column; a series whose x values differ
    is noted and its y values are paired by point index)."""
    ser_els = list(chart._chartSpace.iter(qn("c:ser")))
    xs = _xy_cache(ser_els[0], "c:xVal") if ser_els else []
    ys = [_xy_cache(el, "c:yVal") for el in ser_els]
    if not xs or not any(ys):
        dropped.append("scatter chart has no cached x/y values — data not "
                       "recovered")
        return None
    if len(ser_els) > 1 and any(_xy_cache(el, "c:xVal") != xs
                                for el in ser_els[1:]):
        # per-series x vectors: the [Group] scatter contract (chart_combo)
        # — recover as grouped single-Y rows rather than forcing a shared X.
        return _extract_grouped_scatter(ser_els, series_list)
    n = len(xs)
    short = [i for i, y in enumerate(ys) if len(y) < n]
    if short:
        dropped.append("scatter series with fewer points than x values — "
                       "missing y recovered as 0")
    if len(ser_els) == 1:
        rows = [(_fmt_value(x), [ys[0][i] if i < len(ys[0]) else 0.0])
                for i, x in enumerate(xs)]
        return {"kind": "scatter", "series": None, "data": rows}
    names = ["X"] + [_clean_label(s.name or f"Series {i + 1}")
                     for i, s in enumerate(series_list)]
    x_title = _scatter_x_title(chart)
    if x_title:
        names[0] = _clean_label(x_title)
    rows = [(f"Item {i + 1}", [x] + [y[i] if i < len(y) else 0.0 for y in ys])
            for i, x in enumerate(xs)]
    return {"kind": "scatter", "series": names, "data": rows}


def _extract_grouped_scatter(ser_els, series_list):
    """Per-group scatter series -> '**Series:** X, Y' with
    '- Item N [group]: x, y' rows (the chart_combo [Group] contract)."""
    rows, idx = [], 1
    for si, el in enumerate(ser_els):
        xs = _xy_cache(el, "c:xVal")
        ys = _xy_cache(el, "c:yVal")
        name = (series_list[si].name if si < len(series_list) else "") \
            or f"Group {si + 1}"
        group = _clean_label(name)
        for i, x in enumerate(xs):
            rows.append((f"Item {idx} [{group}]",
                         [x, ys[i] if i < len(ys) else 0.0]))
            idx += 1
    return {"kind": "scatter", "series": ["X", "Y"], "data": rows}


def _scatter_x_title(chart):
    """The X (first value) axis title text of an XY chart, or None."""
    plot_area = chart._chartSpace.find(qn("c:chart")).find(qn("c:plotArea"))
    val_axes = plot_area.findall(qn("c:valAx")) if plot_area is not None else []
    if not val_axes:
        return None
    title = val_axes[0].find(qn("c:title"))
    if title is None:
        return None
    text = " ".join(t.text for t in title.iter(qn("a:t")) if t.text).strip()
    return text or None


def export_picture(shape, out_dir, slide_no, index, dropped):
    """Save an embedded picture to out_dir; returns the file path or None."""
    try:
        image = shape.image
    except (AttributeError, ValueError):
        dropped.append("linked/unreadable picture — not exported")
        return None
    ext = image.ext or "png"
    out_dir.mkdir(parents=True, exist_ok=True)
    path = out_dir / f"slide{slide_no:02d}-{index}.{ext}"
    path.write_bytes(image.blob)
    return path


def _iter_body_paragraphs(shape):
    for para in shape.text_frame.paragraphs:
        text = "".join(r.text for r in para.runs).strip() or para.text.strip()
        if text:
            yield text


def _max_run_pt(shape):
    sizes = [r.font.size.pt for para in shape.text_frame.paragraphs
             for r in para.runs if r.font.size]
    return max(sizes, default=0)


def extract_slide(slide, slide_no, sh_in, assets_dir, report):
    """One slide -> content dict for classify()/render_slide_md()."""
    dropped = report["dropped"]
    title_shape, title = _slide_title(slide)
    content = {"title": title.strip(), "bullets": [], "source": None,
               "notes": None, "tables": [], "charts": [], "pictures": [],
               "picture_cover": False, "title_pt": 0}
    if title_shape is not None:
        content["title_pt"] = _max_run_pt(title_shape)

    if _DIAGRAM_NS in slide.part.blob:
        dropped.append("SmartArt diagram — flattens on extraction, "
                       "re-author as a diagram: visual or layout")

    text_shapes = []
    for shape in iter_shapes(slide.shapes):
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            continue
        if getattr(shape, "has_chart", False):
            chart = extract_chart(shape, dropped)
            if chart:
                content["charts"].append(chart)
            continue
        if getattr(shape, "has_table", False):
            content["tables"].append(
                [[cell.text.strip() for cell in row.cells]
                 for row in shape.table.rows])
            continue
        if shape.shape_type in (MSO_SHAPE_TYPE.PICTURE,
                                MSO_SHAPE_TYPE.LINKED_PICTURE):
            box = _shape_box_in(shape)
            path = export_picture(shape, assets_dir / "extracted",
                                  slide_no, len(content["pictures"]) + 1,
                                  dropped)
            if path:
                content["pictures"].append(str(path))
                if box and box[2] * box[3] >= FULL_IMAGE_COVER * sh_in[0] * sh_in[1]:
                    content["picture_cover"] = True
            continue
        if getattr(shape, "has_text_frame", False) \
                and shape.text_frame.text.strip():
            if title_shape is not None \
                    and shape.shape_id == title_shape.shape_id:
                continue
            text_shapes.append(shape)

    # Reading order: top-to-bottom, then left-to-right.
    def sort_key(s):
        box = _shape_box_in(s)
        return (box[1], box[0]) if box else (99.0, 99.0)

    for shape in sorted(text_shapes, key=sort_key):
        box = _shape_box_in(shape)
        in_footer = box is not None and box[1] > FOOTER_ZONE * sh_in[1]
        for text in _iter_body_paragraphs(shape):
            if PAGENUM_RX.match(text):
                continue
            if text.lower().startswith("source:"):
                content["source"] = text.split(":", 1)[1].strip()
                continue
            if in_footer:
                continue  # footer strings are deck chrome, not content
            content["bullets"].append(text)

    if slide.has_notes_slide:
        notes = slide.notes_slide.notes_text_frame.text.strip()
        if notes:
            content["notes"] = " ".join(notes.split())
    report["recovered"] = {
        "bullets": len(content["bullets"]), "tables": len(content["tables"]),
        "charts": len(content["charts"]),
        "pictures": len(content["pictures"]),
        "notes": bool(content["notes"]),
    }
    return content


def classify(content, slide_no):
    """Pick the outline layout that best re-homes the recovered content."""
    n_words = len(content["title"].split())
    if content["charts"]:
        return "two-column-split"
    if content["tables"]:
        return "table"
    if content["picture_cover"] and not content["bullets"]:
        return "full-image"
    if content["pictures"]:
        return "two-column-split"
    if slide_no == 1 and content["title_pt"] >= TITLE_MIN_PT \
            and len(content["bullets"]) <= 2:
        return "title"
    if not content["bullets"] and not content["pictures"] \
            and content["title"] and n_words <= SECTION_MAX_WORDS:
        return "section-divider"
    return "bullet-list"


def render_slide_md(content, layout, slide_no, report):
    """Content dict + layout -> outline markdown lines for one slide."""
    dropped = report["dropped"]
    heading = content["title"] or f"Slide {slide_no}"
    lines = [f"## Slide {slide_no}: {heading}", f"**Layout:** {layout}"]

    if layout == "title":
        lines.append(f'- Title: "{heading}"')
        if content["bullets"]:
            lines.append(f'- Subtitle: "{content["bullets"][0]}"')
            dropped.extend(f"title-slide text not carried: {b[:50]!r}"
                           for b in content["bullets"][1:])
        content = {**content, "bullets": []}

    if content["charts"]:
        chart = content["charts"][0]
        lines.append(f"**Visual:** chart:{chart['kind']}")
        if chart["series"]:
            lines.append(f"**Series:** {', '.join(chart['series'])}")
        for extra in content["charts"][1:]:
            dropped.append("additional chart on the slide — one chart per "
                           "slide; split it into its own slide")
    if content["pictures"] and not content["charts"]:
        lines.append(f"**Visual:** {content['pictures'][0]}")
        dropped.extend(f"additional picture not placed: {p}"
                       for p in content["pictures"][1:])
    elif content["pictures"] and content["charts"]:
        dropped.extend(f"picture not placed (chart takes the visual): {p}"
                       for p in content["pictures"])

    if layout != "full-image":
        for bullet in content["bullets"]:
            lines.append(f"- {bullet}")

    if layout == "table" and content["tables"]:
        table = content["tables"][0]
        width = max(len(r) for r in table)
        for i, row in enumerate(table):
            cells = [c.replace("|", "/") for c in row + [""] * (width - len(row))]
            lines.append("| " + " | ".join(cells) + " |")
            if i == 0:
                lines.append("|" + "---|" * width)
        dropped.extend("additional table not carried"
                       for _ in content["tables"][1:])
    elif content["tables"]:
        dropped.extend(f"table not carried on layout {layout}"
                       for _ in content["tables"])

    if content["charts"]:
        lines.append("**Data:**")
        for label, values in content["charts"][0]["data"]:
            lines.append(f"- {label}: {', '.join(_fmt_value(v) for v in values)}")
        lines.extend(content["charts"][0].get("extra", []))

    if content["source"]:
        lines.append(f"- Source: {content['source']}")
    if content["notes"]:
        lines.append(f'- Notes: "{content["notes"]}"')
    lines.append("")
    return lines


def extract_outline(pptx_path, assets_dir, palette=None):
    """Whole deck -> (outline markdown, fidelity report dict)."""
    prs = Presentation(str(pptx_path))
    sh_in = (prs.slide_width / EMU_IN, prs.slide_height / EMU_IN)
    lines = [f"# Extracted from {Path(pptx_path).name}", ""]
    if palette:
        lines += [f"**Palette:** {palette}", ""]
    fidelity = {"deck": Path(pptx_path).name, "slides": {}}
    for n, slide in enumerate(prs.slides, 1):
        report = {"recovered": {}, "dropped": [], "layout": None}
        content = extract_slide(slide, n, sh_in, Path(assets_dir), report)
        layout = classify(content, n)
        report["layout"] = layout
        lines.extend(render_slide_md(content, layout, n, report))
        fidelity["slides"][str(n)] = report
    fidelity["total_dropped"] = sum(
        len(s["dropped"]) for s in fidelity["slides"].values())
    return "\n".join(lines), fidelity


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("pptx")
    parser.add_argument("--output", required=True, help="Outline .md to write")
    parser.add_argument("--assets-dir", default="assets",
                        help="Where extracted/ images land (default: assets)")
    parser.add_argument("--palette", default=None,
                        help="Palette key to write into the outline "
                             "front-matter")
    parser.add_argument("--report", default=None,
                        help="Write the fidelity report JSON here")
    args = parser.parse_args()

    if not Path(args.pptx).is_file():
        print(f"ERROR: {args.pptx} not found", file=sys.stderr)
        sys.exit(2)
    outline, fidelity = extract_outline(args.pptx, args.assets_dir,
                                        args.palette)
    Path(args.output).write_text(outline, encoding="utf-8")
    n = len(fidelity["slides"])
    print(f"Extracted {n} slide(s) → {args.output}")
    for sn, rep in fidelity["slides"].items():
        for item in rep["dropped"]:
            print(f"  [DROPPED] slide {sn}: {item}", file=sys.stderr)
    if args.report:
        Path(args.report).write_text(
            json.dumps(fidelity, indent=2, ensure_ascii=False) + "\n",
            encoding="utf-8")
        print(f"Fidelity report: {args.report}")
    if fidelity["total_dropped"]:
        print(f"\n{fidelity['total_dropped']} element(s) could not be "
              "carried over — review before rebuilding "
              "(python3 scripts/build_deck.py "
              f"{args.output} --check)", file=sys.stderr)


if __name__ == "__main__":
    main()
