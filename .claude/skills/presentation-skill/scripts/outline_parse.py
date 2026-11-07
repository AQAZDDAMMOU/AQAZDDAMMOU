#!/usr/bin/env python3
"""
outline_parse.py — the markdown outline parser.

`parse_outline(md_text)` turns `## Slide N:` sections and their `- Key: value`
lines (mapped through `FIELD_KEYS`) into plain slide dicts plus the deck
front-matter (`META_KEYS`).  Layouts left unset are auto-picked here; a
profile may re-pick them later (`smart_layout.repick`).  Outline syntax is
documented in references/generation-guide.md.

Split out of build_deck.py (which re-exports every name here under its old
spelling) so that outline_replay and the revision tools can import the
parser without pulling in the build orchestrator — and so running
`build_deck.py` as `__main__` never imports a second copy of itself.
"""
import re

from helpers import warn
from outline_data import (SLIDE_HEADING_PREFIX, _parse_data_point,
                          _series_count)
from smart_layout import auto_layout
from verbatim import COPY_FIELDS, outline_declares_locked

KV_RE = re.compile(r'(\w+)="([^"]*)"')

# "- Key: value" prefixes mapped to slide-dict fields
FIELD_KEYS = {
    "Title": "title", "Subtitle": "subtitle", "Heading": "heading",
    "Headline": "heading", "Caption": "caption", "Contact": "contact",
    "Notes": "notes", "Series": "series", "Source": "source",
    "Left label": "left_label", "Right label": "right_label",
    # consulting layouts
    "Quote": "quote", "Attribution": "attribution", "Context": "context",
    "Value": "value", "Label": "label", "Callout": "callout",
    "Current": "current", "Situation": "situation",
    "Recommendation": "recommendation",
    "Bracket": "bracket", "CAGR": "cagr", "Axis-Max": "axis_max",
    "Axis-Break": "axis_break",
    "Value-Line": "value_line", "Labels": "labels_mode",
    "X-axis": "x_axis", "Y-axis": "y_axis",
    "Q1": "q1", "Q2": "q2", "Q3": "q3", "Q4": "q4",
    "Sticker": "sticker", "Kicker": "kicker",
    "Scale": "scale", "Sort": "sort", "Marker": "marker",
    "Data-File": "data_file", "Evidence": "evidence",
    "Support": "support",
    # Cycle 7 — chart annotations & options
    "Diff-Arrow": "diff_arrows", "Ordinal": "ordinal",
    "Total-Line": "total_line", "Target": "target",
    "Error-Bars": "error_bars", "Units": "units",
    "Label-Rotation": "label_rotation", "Label-Skip": "label_skip",
    "Gap-Width": "gap_width", "Category-Break": "category_breaks",
    "Data-Rows": "data_rows",
    # Cycle 7 — combo / secondary axis / trend / scatter-bubble
    "Line-Series": "line_series", "Secondary": "secondary",
    "Trend": "trend", "Partition": "partition", "Size-Legend": "size_legend",
    # Cycle 7 — footnotes, agenda columns, pagination
    "Footnote": "footnotes", "Note": "note", "Columns": "columns",
    "Hidden": "hidden",
    # Cycle 7 — table upgrades
    "Align": "align", "Zebra": "zebra", "Total-Row": "total_row",
    "Format": "formats", "Summary": "summary",
}

# `- Key:` lines that may repeat on one slide; parsed into a list, in order.
LIST_FIELDS = {"diff_arrows", "category_breaks", "footnotes", "formats"}

# Cycle 7 keys are guarded so they cannot hijack pre-Cycle-7 outlines
# (mirrors the Q1-Q4 matrix guard): chart-option keys are consumed only
# after the slide has declared a `chart:` visual — declare **Visual:**
# before the option lines; table keys only on an explicit table layout;
# `- Columns:` only on agenda.  Inside an open **Data:** block a line that
# parses as a valid data point stays data even when its label matches a
# guarded key ('- Target: 30' under **Data:** is a data row).
CHART_KEYS = {"diff_arrows", "ordinal", "total_line", "target", "error_bars",
              "label_rotation", "label_skip", "gap_width", "category_breaks",
              "data_rows", "line_series", "secondary", "trend", "partition",
              "size_legend"}
TABLE_KEYS = {"align", "zebra", "total_row", "formats", "summary"}
TABLE_OPTION_LAYOUTS = {"table", "comps-table"}
GUARDED_FIELDS = CHART_KEYS | TABLE_KEYS | {"columns", "units"}


_PLAIN_NUM_RX = re.compile(r"^[$€£]?-?[\d,]*\.?\d+%?x?$")


def _plain_numbers(item, n_values):
    """True when a `label: value` line's value is exactly `n_values` bare
    numbers — i.e. it reads as a data row, not as option syntax (kv pairs
    like 'x=10' or word values like 'desc' fail)."""
    _, sep, value = item.partition(":")
    if not sep:
        return False
    parts = [v.strip() for v in value.split(",")]
    return len(parts) == n_values and \
        all(_PLAIN_NUM_RX.match(v) for v in parts)


def _guard_ok(field, current):
    """True when a guarded Cycle 7 key applies to the slide as declared so
    far (parse-order contract, like Q1-Q4: the guard reads what is above)."""
    if field not in GUARDED_FIELDS:
        return True
    has_chart = any("chart:" in str(current.get(k, ""))
                    for k in ("visual", "visual_left", "visual_right"))
    if field == "units":
        return has_chart or current.get("layout") in TABLE_OPTION_LAYOUTS
    if field == "total_line":   # stacked-100 (chart) or the mekko exhibit
        return has_chart or current.get("layout") == "mekko"
    if field in CHART_KEYS:
        return has_chart
    if field in TABLE_KEYS:
        return current.get("layout") in TABLE_OPTION_LAYOUTS
    return current.get("layout") == "agenda"   # columns


def _parse_kv(line):
    return {k.lower(): v for k, v in KV_RE.findall(line)}



META_KEYS = {"Palette": "palette", "Footer": "footer",
              "Page-Numbers": "page_numbers", "Size": "size",
              "Density": "density", "Purpose": "purpose",
              "Variant": "variant", "Motif": "motif", "Takeaway": "takeaway",
              "Exhibits": "exhibits", "Auto-Agenda": "auto_agenda",
              "Stamp": "stamp", "Scale-Group": "scale_group",
              "Tracker": "tracker", "Sections": "sections",
              "References": "references", "Evidence-File": "evidence_file",
              "Style-File": "style_file", "Author": "author",
              "Number-Format": "number_format", "Profile": "profile",
              "Copy": "copy",
              "Page-Start": "page_start", "File-Stamp": "file_stamp",
              "Date": "date"}


def _clean_bullet(text):
    """Strip symmetric wrapping quotes from a bullet, including quotes that
    wrap only the body after an `icon:<name>` prefix ('icon:x "t"' -> 'icon:x t').
    Asymmetric/inner quotes are left intact."""
    text = text.strip()
    m = re.match(r"^(icon:[a-z0-9-]+\s+)(.*)$", text)
    prefix, body = (m.group(1), m.group(2)) if m else ("", text)
    if len(body) >= 2 and body[0] == body[-1] and body[0] in "\"'":
        body = body[1:-1]
    return prefix + body


def parse_outline(md_text, verbatim=None):
    """Parse a markdown outline into (deck_meta, list of slide dicts).

    Deck-level front-matter: `**Palette:** aurora`, `**Footer:** "Confidential"`,
    `**Page-Numbers:** on`, `**Size:** 16:9|4:3` before the first `## Slide`.

    verbatim: None detects `**Copy:** locked` from the text itself, True/False
    force it (`--verbatim`); verbatim keeps bullets and copy fields exact.
    """
    if verbatim is None:
        verbatim = outline_declares_locked(md_text)
    clean = str.strip if verbatim else _clean_bullet
    meta, slides, current, in_data = {}, [], None, False

    for lineno, raw in enumerate(md_text.splitlines(), 1):
        line = raw.strip()
        if current is None and line.startswith("**"):
            key, _, value = line.lstrip("*").partition(":**")
            if key.strip() in META_KEYS:
                meta[META_KEYS[key.strip()]] = value.strip().strip('"')
                continue
        if line.startswith("## Appendix"):
            if current:
                slides.append(current)
                current = None
            meta["_appendix_from"] = len(slides)  # subsequent slides are backup
            continue
        if line.startswith(SLIDE_HEADING_PREFIX):
            if current:
                slides.append(current)
            current = {"bullets": [], "stats": [], "cards": [], "items": [],
                       "data": [], "table_rows": [], "steps": [], "tiles": [],
                       "matrix_items": [], "bars": [], "milestones": [],
                       "nodes": [], "left_bullets": [], "right_bullets": [],
                       "layers": [], "lanes": []}
            current["_line"] = lineno
            if "_appendix_from" in meta:
                current["_appendix"] = True
            in_data = False
            m = re.match(r"## Slide \d+:\s*(.*)", line)
            if m:
                heading = m.group(1).strip()
                m_attr = re.search(r"\{([^{}]*)\}\s*$", heading)
                if m_attr:
                    pairs = re.findall(r"(\w[\w-]*)=([\w./#-]+)",
                                       m_attr.group(1))
                    if not pairs:
                        warn(f"line {lineno}: heading attributes "
                             f"{m_attr.group(0)!r} did not parse — values "
                             "must be unquoted (layout=waterfall)")
                    for k, v in pairs:
                        if k == "layout":
                            current["layout"] = v.lower()
                        elif k == "palette":
                            current["palette"] = v.lower()
                        else:
                            warn(f"line {lineno}: unknown heading attribute "
                                 f"'{k}' ignored (supported: layout, palette)")
                    heading = heading[:m_attr.start()].strip()
                current["heading"] = heading
            continue
        if current is None:
            continue
        if not line:
            continue  # blank lines inside **Data:** blocks are ignored

        if line.startswith("**"):
            in_data = False
            key, _, value = line.lstrip("*").partition(":**")
            key, value = key.strip(), value.strip()
            if key == "Layout":
                current["layout"] = value.lower()
            elif key == "Visual":
                current["visual"] = value
            elif key == "Visual-Left":
                current["visual_left"] = value
            elif key == "Visual-Right":
                current["visual_right"] = value
            elif key == "Palette":
                current["palette"] = value.lower()
            elif key == "Series":
                current["series"] = value
            elif key == "Benchmark":
                current["benchmark"] = value
            elif key == "Periods":
                current["periods"] = [s.strip() for s in value.split(",")
                                      if s.strip()]
            elif key == "Data":
                in_data = True
            elif key == "Repeat":
                current["repeat"] = value
            elif key == "Copy":
                warn(f"line {lineno}: **Copy:** must be deck front matter "
                     "(before the first '## Slide') — verbatim NOT enabled")
            continue

        if line.startswith("|"):
            cells = [c.strip() for c in line.strip("|").split("|")]
            if not all(set(c) <= {"-", ":", " "} for c in cells):  # skip ---- row
                current["table_rows"].append(cells)
            continue

        if not line.startswith("- "):
            continue
        item = line[2:].strip()
        indent = len(raw) - len(raw.lstrip(" ")) >= 2

        # Known field keys take precedence over (and terminate) a Data block,
        # so '- Heading: "32% YoY"' after **Data:** isn't read as a data point.
        matched = False
        for prefix, field in FIELD_KEYS.items():
            # Q1-Q4 are quadrant labels only on matrix slides; elsewhere they
            # are ordinary labels (e.g. '- Q1: 4.2, 3.1' chart data rows)
            if field in ("q1", "q2", "q3", "q4") and \
                    current.get("layout") != "matrix-2x2":
                continue
            if item.startswith(prefix + ":"):
                if not _guard_ok(field, current):
                    continue
                if field in GUARDED_FIELDS and in_data and \
                        _plain_numbers(item, _series_count(current)):
                    break  # a real data row shadowed by a guarded key name
                val = item[len(prefix) + 1:].strip()
                val = (val if verbatim and field in COPY_FIELDS
                       else val.strip('"'))
                if field in LIST_FIELDS:
                    current.setdefault(field, []).append(val)
                else:
                    current[field] = val
                matched = True
                if field not in GUARDED_FIELDS:
                    in_data = False  # legacy keys terminate a **Data:** block
                break
        if matched:
            continue

        if in_data:
            point = _parse_data_point(item, _series_count(current))
            if point:
                current["data"].append(point)
                # raw row text, for validators that need to see what a
                # single-series parse narrowed away (chart:scatter x, y rows)
                current.setdefault("_data_raw", []).append(item)
                continue
            in_data = False  # fall through: not a data line after all

        if item.startswith("Left:"):
            current["left_bullets"].append(clean(item[5:]))
        elif item.startswith("Right:"):
            current["right_bullets"].append(clean(item[6:]))
        elif re.match(r"Card \d+:", item):
            current["cards"].append(_parse_kv(item))
        elif re.match(r"Step \d+:", item):
            current["steps"].append(_parse_kv(item))
        elif re.match(r"Tile \d+:", item):
            current["tiles"].append(_parse_kv(item))
        elif item.startswith("Item:"):
            current["matrix_items"].append(_parse_kv(item))
        elif item.startswith("Bar:"):
            current["bars"].append(_parse_kv(item))
        elif item.startswith("Milestone:"):
            current["milestones"].append(_parse_kv(item))
        elif item.startswith("Span:"):
            current.setdefault("spans", []).append(_parse_kv(item))
        elif item.startswith("Shade:"):
            current.setdefault("shades", []).append(_parse_kv(item))
        elif item.startswith("Line:"):
            current.setdefault("gantt_lines", []).append(_parse_kv(item))
        elif item.startswith("Tombstone:"):
            current.setdefault("tombstones", []).append(_parse_kv(item))
        elif item.startswith("Logo:"):
            current.setdefault("logos", []).append(item[5:].strip())
        elif item.startswith("Node:"):
            current["nodes"].append(_parse_kv(item))
        elif item.startswith("Layer:"):
            current["layers"].append(_parse_kv(item))
        elif item.startswith("Lane:"):
            current["lanes"].append(_parse_kv(item))
        elif item.startswith("Value=") or item.startswith("Stat"):
            kv = _parse_kv(item)
            if kv:
                current["stats"].append(kv)
        elif "date=" in item.lower() and KV_RE.search(item):
            current["items"].append(_parse_kv(item))
        elif re.match(r"(Point|Finding) \d+:", item):
            current["bullets"].append(clean(item.split(":", 1)[1]))
            current.setdefault("_bullet_indent", []).append(indent)
        else:
            current["bullets"].append(clean(item))
            current.setdefault("_bullet_indent", []).append(indent)

    if current:
        slides.append(current)
    for slide in slides:
        if not slide.get("layout"):
            slide["layout"] = auto_layout(slide)
            slide["_auto_layout"] = True   # a profile may re-pick it later
    return meta, slides

