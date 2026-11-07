#!/usr/bin/env python3
"""
outline_validate.py — fail-fast outline validation for build_deck.

`validate(slides, ctx, meta)` returns `(errors, warnings)` for a parsed
outline: any error means build_deck writes no .pptx. Split out of
build_deck.py; every public name here is re-exported from build_deck, so
`from build_deck import validate` (and ACTION_TITLE_LAYOUTS,
NO_FOOTER_LAYOUTS, ...) keeps working.

Layout-specific checks live inline in `validate()`; the tree/matrix/RACI
row checks that need their own loops are the `_validate_*` helpers above it.
"""
import re
from pathlib import Path

from builders import LAYOUT_MAP
from builders_banking import MAX_LEAGUE_ROWS, MAX_LOGOS, MAX_TOMBSTONES
from builders_grid import MAX_WHW_STEPS
from helpers import parse_visual, parse_visual_opts, resolve_image_path
from outline_data import _series_count
from narrative import (check_house_style, check_unsourced_stats, load_style,
                       resolve_style, validate_narrative)
from palettes import PALETTES
import profiles
from template_helpers import load_template_config, validate_template_mapping
from verbatim import copy_is_locked, is_verbatim_deck, verbatim_cap_errors

NO_FOOTER_LAYOUTS = {"title", "closing", "section-divider", "full-image"}

# Layouts whose headings should pass the consulting "titles test"
ACTION_TITLE_LAYOUTS = {
    "bullet-list", "two-column-split", "exec-summary", "exec-summary-scqa",
    "comparison", "table", "comps-table", "stat-callout", "timeline",
    "waterfall",
    "matrix-2x2", "chart-callout", "dashboard", "funnel", "harvey-scorecard",
    "process-flow", "mekko", "gantt", "bar-mekko",
    "heatmap-table", "tornado", "football-field",
    "driver-tree", "stakeholder-map", "raci",
}


_TREE_NEEDS_ROWS_HINT = {
    "org-chart": 'Id=".." Label=".." Parent=".." [Role=".."]',
}
_DEFAULT_TREE_NEEDS_ROWS_HINT = 'Id=".." Label=".." Value=".." Parent=".."'


def _validate_tree(p, where, errors, layout):
    """Tree layouts (driver-tree, org-chart): one root, parents exist, unique
    ids, depth<=3, <=12 nodes."""
    nodes = p.get("nodes", [])
    if not nodes:
        hint = _TREE_NEEDS_ROWS_HINT.get(layout, _DEFAULT_TREE_NEEDS_ROWS_HINT)
        errors.append(f"{where}: {layout} needs '- Node: {hint}' rows")
        return
    ids = [n.get("id", "").strip() for n in nodes]
    if len(nodes) > 12:
        errors.append(f"{where}: {layout} supports at most 12 nodes "
                      f"(got {len(nodes)})")
    if not all(ids):
        errors.append(f"{where}: every {layout} Node needs an Id")
        return
    if len(set(ids)) != len(ids):
        errors.append(f"{where}: {layout} Node Ids must be unique")
        return
    roots = [n for n in nodes if not n.get("parent", "").strip()]
    if len(roots) != 1:
        errors.append(f"{where}: {layout} needs exactly one root node "
                      f"(empty Parent) — got {len(roots)}")
    known = set(ids)
    bad = sorted({n["parent"] for n in nodes
                  if n.get("parent", "").strip() and n["parent"] not in known})
    if bad:
        errors.append(f"{where}: {layout} Parent ids not found: "
                      f"{', '.join(bad)}")
    if len(roots) != 1 or bad:
        return
    children = {}
    for n in nodes:
        if n.get("parent", "").strip():
            children.setdefault(n["parent"], []).append(n["id"])
    depth, queue = {roots[0]["id"]: 1}, [roots[0]["id"]]
    while queue:
        nid = queue.pop(0)
        for ch in children.get(nid, []):
            depth[ch] = depth[nid] + 1
            queue.append(ch)
    if len(depth) != len(nodes):
        missing = sorted(set(ids) - set(depth))
        errors.append(f"{where}: {layout} nodes not reachable from the "
                      f"root (cycle?): {', '.join(missing)}")
    elif max(depth.values()) > 3:
        errors.append(f"{where}: {layout} depth exceeds 3 levels "
                      "(root + 2 — split the decomposition)")


def _validate_driver_tree(p, where, errors):
    _validate_tree(p, where, errors, "driver-tree")


_TARGET_KEYS = {"targetx": "TargetX", "targety": "TargetY"}


def _validate_stakeholder_map(p, where, errors, warnings):
    """Stakeholder-map: 2+ valid 0-1 items; targets, when given, 0-1 pairs."""
    valid = 0
    for it in p.get("matrix_items", []):
        try:
            x, y = float(it.get("x", "")), float(it.get("y", ""))
            if 0 <= x <= 1 and 0 <= y <= 1:
                valid += 1
        except ValueError:
            pass
        given = [k for k in _TARGET_KEYS if it.get(k)]
        if len(given) == 1:
            warnings.append(f"{where}: Item {it.get('name', '')!r} has only "
                            "one of TargetX/TargetY — arrow will be skipped")
        for k in given:
            try:
                if not 0 <= float(it[k]) <= 1:
                    raise ValueError
            except ValueError:
                errors.append(f"{where}: Item {it.get('name', '')!r} "
                              f"{_TARGET_KEYS[k]} must be a 0-1 float "
                              f"(got {it[k]!r})")
    if valid < 2:
        errors.append(f"{where}: stakeholder-map needs 2+ '- Item: "
                      'Name=".." X="0-1" Y="0-1"\' rows')


def _validate_raci(p, where, errors, warnings):
    """RACI: header + body table; cells R/A/C/I or blank; one 'A' per row."""
    t_rows = p.get("table_rows", [])
    if len(t_rows) < 2 or max((len(r) for r in t_rows), default=0) < 2:
        errors.append(f"{where}: raci needs a markdown table "
                      "(| Activity | Person… | header plus body rows)")
        return
    for row in t_rows[1:]:
        cells = [c.strip().upper() for c in row[1:]]
        for c in cells:
            if c and c not in ("R", "A", "C", "I"):
                errors.append(f"{where}: raci cell {c!r} in row {row[0]!r} "
                              "must be a single R/A/C/I letter or blank")
        n_a = cells.count("A")
        if n_a != 1:
            warnings.append(f"{where}: activity {row[0]!r} has {n_a} 'A' "
                            "cells — exactly one Accountable per activity")


def _raw_value_count(raw_row):
    """How many comma-separated numeric columns a raw `- label: v1, v2` data
    row carries (0 when it is not a data row)."""
    from outline_data import NUM_RE
    _, sep, value = str(raw_row).partition(":")
    if not sep:
        return 0
    # a comma glued to exactly three digits is a thousands separator (1,200),
    # any other comma is a column break (12, 30 / 8,22)
    parts = re.split(r",(?!\d{3}(?!\d))", value)
    return sum(1 for part in parts if NUM_RE.search(part))


from outline_validate_exhibits import (  # noqa: E402  (re-export)
    _validate_chart_options, _validate_combo_bubble, _validate_gantt,
    _validate_table)


def _validate_scatter(p, where, n_series, errors, warnings):
    """chart:scatter data contract (charts.scatter_xy): multi-value rows need
    `**Series:** X, Y1[, ...]` (first name is the X column); single-value
    rows use numeric labels as x, else the 1-based row index."""
    if parse_visual(p.get("visual", "")) != ("chart", "scatter") \
            or not p.get("data"):
        return
    from charts import labels_numeric
    multi_rows = any(isinstance(v, list) and len(v) > 1 for _, v in p["data"])
    if not multi_rows and n_series < 2:
        # A single-series parse keeps only the first number of `- A: 12, 30`;
        # the raw row text still shows the second column the author wrote.
        multi_rows = any(_raw_value_count(raw) > 1
                         for raw in p.get("_data_raw") or [])
    if multi_rows and n_series < 2:
        errors.append(
            f"{where}: chart:scatter with multi-value rows needs "
            "'**Series:** X, Y1[, Y2 ...]' — the first name is the X column, "
            f"each further name one Y series (got **Series:** {p.get('series')!r})")
    elif not multi_rows and not labels_numeric(p["data"]):
        warnings.append(
            f"{where}: chart:scatter labels are not all numeric — x falls "
            "back to the row index (1, 2, ...); use numeric labels or "
            "'**Series:** X, Y' with '- label: x, y' rows")


def validate(slides, ctx, meta=None):
    """Fail-fast outline checks. Returns (errors, warnings)."""
    meta = meta or {}
    errors, warnings = [], []
    if not slides:
        return ["Outline contains no '## Slide N:' sections"], []

    deck_pal = meta.get("palette")
    if deck_pal and deck_pal not in PALETTES:
        warnings.append(f"unknown deck palette '{deck_pal}' — default will be used")

    if meta.get("tracker", "").lower() == "tabs" and not any(
            p.get("layout") == "section-divider" and not p.get("_appendix")
            for p in slides):
        warnings.append("**Tracker:** tabs needs section-divider slides — "
                        "falling back to the plain section label")

    copy_val = str(meta.get("copy", "")).strip()
    if copy_val and not copy_is_locked(copy_val):
        warnings.append(f"**Copy:** value {copy_val!r} is not 'locked' — "
                        "verbatim mode NOT enabled")
    if is_verbatim_deck(ctx, meta):   # locked copy: truncation is an error
        errors.extend(verbatim_cap_errors(slides))

    for n, p in enumerate(slides, 1):
        where = f"Slide {n} (line {p['_line']})" if p.get("_line") else f"Slide {n}"
        layout = p.get("layout", "bullet-list")
        if layout not in LAYOUT_MAP:
            errors.append(f"{where}: unknown layout '{layout}' "
                          f"(valid: {', '.join(sorted(LAYOUT_MAP))})")
            continue
        if "palette" in p and p["palette"] not in PALETTES:
            warnings.append(f"{where}: unknown palette '{p['palette']}', "
                            "default will be used")

        for vis_key in ("visual", "visual_left", "visual_right"):
            spec = p.get(vis_key, "")
            kind, value, vis_opts = parse_visual_opts(spec, warn_opts=False)
            if not kind:
                continue
            if kind == "image":
                given = [s for s in spec.split("|")[1:] if s.strip()]
                if len(vis_opts) < len(given):
                    warnings.append(
                        f"{where}: {vis_key} has malformed/unknown image "
                        f"option(s) in {spec!r} — supported: alpha=0-100, "
                        "duotone (they will be ignored)")
                if vis_opts and ctx.get("template_path"):
                    warnings.append(
                        f"{where}: image options {list(vis_opts)} on {vis_key} "
                        "are ignored in template mode")
            if kind == "image" and not resolve_image_path(value, ctx):
                errors.append(f"{where}: image not found for {vis_key}: '{value}'")
            elif kind == "chart":
                if value not in __import__("charts").CHART_TYPES:
                    errors.append(f"{where}: unknown chart type '{value}'")
                if vis_key == "visual" and not p.get("data"):
                    errors.append(f"{where}: 'chart:{value}' visual but no "
                                  "**Data:** block (or no parsable '- label: value' lines)")
                elif vis_key != "visual":
                    errors.append(f"{where}: charts on {vis_key} are not supported "
                                  "(use **Visual:** on two-column-split)")
            elif kind == "diagram":
                src = resolve_image_path("diagram:" + value, ctx)
                if src is None:
                    errors.append(
                        f"{where}: diagram source not found: {value!r} "
                        "(searched outline dir and assets/diagrams/)")
                elif src.suffix.lower() not in (".mmd", ".dot", ".gv"):
                    errors.append(
                        f"{where}: diagram source must be .mmd/.dot/.gv, "
                        f"got {src.suffix!r}")
                else:
                    import shutil
                    needs = ("dot" if src.suffix.lower() in (".dot", ".gv")
                             else "mmdc")
                    if not shutil.which(needs):
                        warnings.append(
                            f"{where}: renderer `{needs}` not on PATH — build "
                            "will fail until it is installed "
                            "(see scripts/render_diagram.py docstring)")

        if not p.get("notes") and not p.get("_appendix") \
                and layout != "section-divider":
            warnings.append(f"{where}: missing '- Notes:' speaker notes")

        # Titles test: content-slide headings should be action titles —
        # full-sentence takeaways, not topic labels ("Leadership" fails).
        if (layout in ACTION_TITLE_LAYOUTS and not p.get("_appendix")
                and len(p.get("heading", "").split()) < 5):
            warnings.append(
                f"{where}: heading {p.get('heading', '')!r} is a topic label, "
                "not an action title — state the takeaway as a sentence "
                "(run --titles for the titles test)")

        if layout in ACTION_TITLE_LAYOUTS and not p.get("_appendix"):
            # Heading length lives in house_style (it is style-file tunable
            # via max_title_words) — do not duplicate the rule here.
            heading = p.get("heading", "")
            chart_ish = parse_visual(p.get("visual", ""))[0] == "chart" or \
                layout in ("waterfall", "mekko", "bar-mekko", "chart-callout")
            if chart_ish and heading and not any(c.isdigit() for c in heading):
                warnings.append(f"{where}: exhibit slide heading has no "
                                "number — quantify the takeaway")
            if " and " in heading.lower():
                warnings.append(f"{where}: heading joins two messages with "
                                "'and' — consider splitting the slide")

        if p.get("kicker"):
            kick_words = {w for w in re.findall(r"[a-z0-9']+",
                                                p["kicker"].lower())
                          if len(w) > 3}
            head_words = {w for w in re.findall(r"[a-z0-9']+",
                                                p.get("heading", "").lower())
                          if len(w) > 3}
            if kick_words and len(kick_words & head_words) / len(kick_words) > 0.6:
                warnings.append(f"{where}: kicker restates the title — "
                                "make it advance the argument")
        if p.get("kicker") and layout in NO_FOOTER_LAYOUTS:
            warnings.append(f"{where}: Kicker is not rendered on '{layout}' "
                            "layouts — move it to a content slide")
        if p.get("kicker") and layout == "cards-4":
            errors.append(f"{where}: cards-4 and Kicker cannot coexist — the "
                          "2x2 card grid leaves no room for card bodies above "
                          "the kicker band; drop one or switch to bullet-list")

        if layout in ("title", "closing") and not p.get("title") and p.get("heading"):
            warnings.append(f"{where}: no '- Title:' — using '## Slide N:' heading "
                            f"({p['heading']!r}) as title text")

        n_series = _series_count(p)
        if n_series > 1 and parse_visual(p.get("visual", ""))[0] == "chart":
            if not p.get("data"):
                errors.append(
                    f"{where}: multi-series chart needs **Data:** rows with "
                    f"{n_series} comma-separated values per line "
                    f"(**Series:** {p.get('series')})")
            else:
                for label, vals in p["data"]:
                    if not isinstance(vals, list) or len(vals) != n_series:
                        errors.append(
                            f"{where}: multi-series row {label!r} needs {n_series} "
                            f"comma-separated values (got **Series:** {p.get('series')})")
                        break
        _validate_scatter(p, where, n_series, errors, warnings)
        _validate_chart_options(p, where, errors)
        _validate_combo_bubble(p, where, errors, warnings)

        import builders
        cap = builders.density()["bullet_max"]
        if len(p.get("bullets", [])) > cap:
            warnings.append(f"{where}: {len(p['bullets'])} bullets — only the "
                            f"first {cap} render; split the slide")

        # Overflow guard: estimated text height vs the layout's vertical
        # budget (catches long bullets; the bullet-count cap catches many).
        if layout in ("bullet-list", "bullets", "two-column-split"):
            import textfit
            geo = (builders.twocol_geometry() if layout == "two-column-split"
                   else builders.bullet_geometry())
            shown = p.get("bullets", [])[:geo["max"]]
            if shown:
                ratio, _, overflowing = textfit.bullets_fit(
                    shown, geo["font_pt"], geo["col_w"], geo["avail_h"],
                    cols=geo["cols"], step_in=geo["step"])
                if ratio > 1.4:
                    errors.append(
                        f"{where}: text overflows layout capacity "
                        f"(~{round(ratio * 100)}%) — split the slide or cut copy")
                elif ratio > 1.0:
                    warnings.append(
                        f"{where}: slide ~{round(ratio * 100)}% of text "
                        "capacity — autofit will shrink text")
                # Per-slot collision guard: non-terminal bullets that overflow
                # their fixed y-step slot will render on top of the next bullet.
                if overflowing and geo["step"] > 0:
                    line_h = geo["font_pt"] * 1.2 / 72.0
                    for idx in overflowing:
                        plain, has_icon = textfit.strip_markup(shown[idx])
                        w = geo["col_w"] - (textfit.ICON_INDENT if has_icon
                                            else textfit.MARKER_INDENT)
                        lines = max(
                            textfit.estimate_lines(plain, geo["font_pt"], w), 1)
                        sev = lines * line_h / geo["step"]
                        msg = (
                            f"{where}: bullet {idx + 1} wraps to ~{lines} lines "
                            "and will collide with the next bullet — "
                            "shorten it or split the slide")
                        if sev > 1.8:
                            errors.append(msg)
                        else:
                            warnings.append(msg)
            head = textfit.strip_markup(p.get("heading", ""))[0]
            if head:
                h_w, h_pt = ((6.0, 30) if layout == "two-column-split"
                             else (11.9, 32))
                h_lines = textfit.estimate_lines(head, h_pt, h_w, bold=True)
                if h_lines > 2:
                    warnings.append(f"{where}: title wraps to ~{h_lines} "
                                    "lines — shorten the heading")
                elif h_lines == 2:
                    wrapped = textfit.wrap_lines(head, h_pt, h_w, bold=True)
                    if wrapped and len(wrapped[-1].split()) == 1:
                        warnings.append(
                            f"{where}: title wraps with the single word "
                            f"{wrapped[-1]!r} orphaned on line 2 — shorten "
                            "or rephrase the heading")
        if layout == "stat-callout" and not p.get("stats"):
            errors.append(f"{where}: stat-callout requires "
                          '\'- Value="..." Label="..." Sublabel="..."\' lines')
        if layout in ("table", "comps-table"):
            if not p.get("table_rows"):
                errors.append(f"{where}: {layout} layout requires markdown "
                              "table rows")
            else:
                _validate_table(p, where, errors, warnings)
        if layout in ("cards-3", "cards-4"):
            want = 3 if layout == "cards-3" else 4
            if len(p.get("cards", [])) != want:
                warnings.append(f"{where}: {layout} has {len(p.get('cards', []))} "
                                f"cards (expected {want})")
        if layout == "timeline" and not p.get("items") and not p.get("bullets"):
            errors.append(f"{where}: timeline needs '- Date=\"..\" Title=\"..\"' "
                          "items or 'Date: text' bullets")
        if layout == "full-image" and parse_visual(p.get("visual", ""))[0] \
                not in ("image", "diagram"):
            warnings.append(f"{where}: full-image without an image visual "
                            "renders a plain background")

        # consulting layout requirements
        numeric_data = [d for d in p.get("data", [])
                        if isinstance(d[1], (int, float, list))]
        sentinel_rows = [d for d in p.get("data", [])
                         if d[1] in ("total", "subtotal")]
        if parse_visual(p.get("visual", ""))[0] == "chart" and sentinel_rows:
            errors.append(f"{where}: 'total'/'subtotal' rows are only valid on "
                          "waterfall layouts, not charts")
        if layout != "waterfall" and layout in LAYOUT_MAP and sentinel_rows:
            errors.append(f"{where}: 'total'/'subtotal' rows are only valid on "
                          f"waterfall layouts, not {layout}")
        if layout == "waterfall":
            wdata = p.get("data", [])
            if len(wdata) < 2:
                errors.append(f"{where}: waterfall needs a **Data:** block — "
                              "start value, +/- deltas, optional "
                              "'- End label: total' row")
            else:
                if wdata[0][1] in ("total", "subtotal"):
                    errors.append(f"{where}: waterfall's first row must be the "
                                  "start value, not 'total'/'subtotal'")
                if wdata[-1][1] == "subtotal":
                    errors.append(f"{where}: waterfall's last row cannot be "
                                  "'subtotal' — use 'total' to close the bridge")
                for lbl, v in wdata[1:]:
                    if isinstance(v, list) and len(v) > 1 and \
                            not (all(x >= 0 for x in v) or all(x <= 0 for x in v)):
                        errors.append(
                            f"{where}: stacked waterfall row {lbl!r} mixes "
                            "signs — segments in one bar must share the "
                            "delta's sign")
        if layout == "funnel" and len(numeric_data) < 2:
            errors.append(f"{where}: funnel needs a **Data:** block with 2+ "
                          "numeric '- Stage: value' rows")
        if layout == "matrix-2x2" and not p.get("matrix_items"):
            errors.append(f"{where}: matrix-2x2 needs '- Item: Name=\"..\" "
                          "X=\"0-1\" Y=\"0-1\"' rows")
        if layout == "harvey-scorecard" and len(p.get("table_rows", [])) < 2:
            errors.append(f"{where}: harvey-scorecard needs a markdown table "
                          "(header row + criteria rows, cells 0-4)")
        if layout in ("process-flow", "next-steps") and not p.get("steps") \
                and not p.get("bullets"):
            errors.append(f"{where}: {layout} needs '- Step N: ...' rows")
        if layout == "big-number" and not p.get("value"):
            errors.append(f"{where}: big-number needs '- Value: \"$4.2B\"'")
        if layout == "chart-callout":
            if not p.get("callout"):
                errors.append(f"{where}: chart-callout needs '- Callout: \"...\"'")
            if not p.get("data") and parse_visual(p.get("visual", ""))[0] \
                    not in ("image", "diagram"):
                errors.append(f"{where}: chart-callout needs a chart **Data:** "
                              "block or an image visual")
        if layout == "dashboard" and not p.get("tiles"):
            errors.append(f"{where}: dashboard needs '- Tile N: Value=\"..\" "
                          "Label=\"..\" Delta=\"..\" Trend=\"up|down\"' rows")
        if layout == "architecture-layers" and not p.get("layers"):
            errors.append(f"{where}: architecture-layers needs '- Layer: "
                          "Name=\"..\" Items=\"a, b\"' rows")
        if layout == "architecture-layers" and len(p.get("layers", [])) > 6:
            errors.append(f"{where}: architecture-layers supports at most 6 "
                          f"layers (got {len(p['layers'])})")
        if layout == "swimlane-process":
            lanes = p.get("lanes", [])
            if not lanes:
                errors.append(f"{where}: swimlane-process needs '- Lane: "
                              "Name=\"..\" Steps=\"a > b\"' rows")
            elif len(lanes) > 4:
                errors.append(f"{where}: swimlane-process supports at most 4 "
                              f"lanes (got {len(lanes)})")
            elif not all(lane.get("steps", "").strip() for lane in lanes):
                errors.append(f"{where}: every swimlane-process Lane needs "
                              "Steps=\"a > b\" ('>' separated)")
        if layout == "what-how-who":
            steps = p.get("steps", [])
            if len(steps) < 2:
                errors.append(f"{where}: what-how-who needs 2-"
                              f"{MAX_WHW_STEPS} '- Step N: Title=\"..\" "
                              "What=\"..\" How=\"..\" Who=\"..\"' rows")
            elif not any(s.get(k, "").strip() for s in steps
                         for k in ("what", "how", "who")):
                errors.append(f"{where}: what-how-who steps need What=/"
                              "How=/Who= values")
            elif len(steps) > MAX_WHW_STEPS:
                warnings.append(f"{where}: what-how-who renders the first "
                                f"{MAX_WHW_STEPS} steps (got {len(steps)})")
        if layout == "tombstone-grid":
            stones = p.get("tombstones", [])
            if not stones:
                errors.append(f"{where}: tombstone-grid needs '- Tombstone: "
                              "Client=\"..\" Deal=\"..\" Value=\"..\" "
                              "Date=\"..\"' rows")
            else:
                bad = [str(i) for i, t in enumerate(stones, 1)
                       if not (t.get("client", "").strip()
                               and t.get("deal", "").strip())]
                if bad:
                    errors.append(f"{where}: every Tombstone needs Client= "
                                  f"and Deal= (missing on row(s) "
                                  f"{', '.join(bad)})")
                if len(stones) > MAX_TOMBSTONES:
                    warnings.append(f"{where}: tombstone-grid renders the "
                                    f"first {MAX_TOMBSTONES} tombstones "
                                    f"(got {len(stones)})")
        if layout == "league-table":
            t_rows = p.get("table_rows", [])
            if len(t_rows) < 2:
                errors.append(f"{where}: league-table needs a markdown table "
                              "(header row + ranked rows)")
            elif len(t_rows) - 1 > MAX_LEAGUE_ROWS:
                warnings.append(f"{where}: league-table renders the first "
                                f"{MAX_LEAGUE_ROWS} rows "
                                f"(got {len(t_rows) - 1})")
        if layout == "logo-wall":
            logos = p.get("logos", [])
            if not logos:
                errors.append(f"{where}: logo-wall needs '- Logo: file.png "
                              "\"Name\"' rows")
            elif len(logos) > MAX_LOGOS:
                warnings.append(f"{where}: logo-wall renders the first "
                                f"{MAX_LOGOS} logos (got {len(logos)})")
        if layout == "value-chain" and not 2 <= len(p.get("steps", [])) <= 6:
            errors.append(f"{where}: value-chain needs 2-6 '- Step N: "
                          "Title=\"..\" [Sub=\"a, b\"]' rows")
        if layout == "quote-evidence" and not p.get("quote"):
            errors.append(f"{where}: quote-evidence needs '- Quote: \"...\"'")
        if layout == "exec-summary-scqa" and not (
                p.get("situation") and p.get("recommendation")):
            errors.append(f"{where}: exec-summary-scqa needs '- Situation:' and "
                          "'- Recommendation:' (findings via '- Finding N:')")
        if layout == "agenda" and not p.get("bullets"):
            errors.append(f"{where}: agenda needs section names as bullets")
        if layout == "mekko":
            if _series_count(p) < 2 or not any(
                    isinstance(d[1], list) for d in p.get("data", [])):
                errors.append(f"{where}: mekko needs **Series:** with 2+ names "
                              "and multi-value **Data:** rows")
        if layout == "gantt":
            _validate_gantt(p, where, errors, warnings)
        if layout == "bar-mekko":
            ok_bars = [b for b in p.get("bars", [])
                       if "size" in b and "value" in b]
            if len(ok_bars) < 2:
                errors.append(f"{where}: bar-mekko needs 2+ '- Bar: "
                              'Label=".." Size=".." Value=".."\' rows')
            for bar in ok_bars:
                try:
                    size, value = float(bar["size"]), float(bar["value"])
                    if size <= 0 or value < 0:
                        errors.append(f"{where}: Bar {bar.get('label', '')!r} "
                                      "needs Size>0 and Value>=0")
                except ValueError:
                    errors.append(f"{where}: Bar {bar.get('label', '')!r} has "
                                  "non-numeric Size/Value")
        if layout == "heatmap-table":
            from builders_consulting import _cell_value
            t_rows = p.get("table_rows", [])
            n_cols = max((len(r) for r in t_rows), default=0)
            numeric = any(_cell_value(c) is not None
                          for r in t_rows[1:] for c in r)
            if len(t_rows) < 2 or n_cols < 2 or not numeric:
                errors.append(f"{where}: heatmap-table needs a markdown table "
                              "(header + body rows, 2+ columns) with at least "
                              "one numeric body column")
        if layout in ("tornado", "football-field"):
            pairs = [d for d in p.get("data", [])
                     if isinstance(d[1], list) and len(d[1]) == 2]
            if len(pairs) < 2:
                errors.append(f"{where}: {layout} needs '**Series:** Low, "
                              "High' and 2+ '- Name: low, high' **Data:** rows")
            elif layout == "football-field":
                for lbl, (lo, hi) in pairs:
                    if not lo < hi:
                        errors.append(f"{where}: football-field row {lbl!r} "
                                      f"needs low < high (got {lo:g}, {hi:g})")
                if p.get("marker"):
                    from builders_consulting import _parse_marker
                    parsed = _parse_marker(p["marker"])
                    gmin = min(v[0] for _, v in pairs)
                    gmax = max(v[1] for _, v in pairs)
                    if not parsed:
                        warnings.append(
                            f"{where}: Marker needs 'label, value' (got "
                            f"{p['marker']!r}) — it will be skipped")
                    elif not gmin <= parsed[1] <= gmax:
                        warnings.append(
                            f"{where}: Marker value {parsed[1]:g} outside "
                            f"[{gmin:g}, {gmax:g}] — it will be skipped")
        if layout == "driver-tree":
            _validate_driver_tree(p, where, errors)
        if layout == "org-chart":
            _validate_tree(p, where, errors, "org-chart")
        if layout == "stakeholder-map":
            _validate_stakeholder_map(p, where, errors, warnings)
        if layout == "raci":
            _validate_raci(p, where, errors, warnings)

    # Narrative + data integrity
    warnings.extend(validate_narrative(meta, slides))
    warnings.extend(check_unsourced_stats(slides))

    # Client profile: mandatory slides are structural, notes style editorial.
    profile = ctx.get("_profile") or {}
    errors.extend(profiles.check_mandatory(slides, profile))
    warnings.extend(profiles.check_notes_style(slides, profile))

    # House style: editorial word lists + action-title hygiene. A declared or
    # requested style file that cannot be loaded is fail-fast.
    style_path, style_errors = resolve_style(meta, ctx)
    if style_errors:
        errors.extend(style_errors)
    else:
        try:
            style = load_style(style_path)
        except ValueError as e:
            errors.append(f"style file {str(style_path)!r}: {e}")
        else:
            warnings.extend(check_house_style(
                slides, meta, profiles.merge_style(style, profile)))

    template_path = ctx.get("template_path")
    if template_path and Path(template_path).is_file():
        from pptx import Presentation
        config = load_template_config(template_path)
        prs = Presentation(str(template_path))
        warnings.extend(validate_template_mapping(prs, slides, config))

    # Layout variety: 3+ identical consecutive layouts reads as monotonous
    # (appendix/backup slides are deliberately dense and repetitive — exempt)
    layouts = [p.get("layout", "bullet-list") for p in slides
               if not p.get("_appendix")]
    run_start = 0
    for i in range(1, len(layouts) + 1):
        if i == len(layouts) or layouts[i] != layouts[run_start]:
            if i - run_start >= 3 and layouts[run_start] not in ("full-image",):
                warnings.append(
                    f"Slides {run_start + 1}-{i}: '{layouts[run_start]}' used "
                    f"{i - run_start}x consecutively — vary layouts (cards, "
                    "stat-callout, two-column-split...)")
            run_start = i
    return errors, warnings
