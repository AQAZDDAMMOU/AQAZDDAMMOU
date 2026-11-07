#!/usr/bin/env python3
"""
outline_validate_exhibits.py — the Cycle 7 exhibit/option validators.

Table options, gantt furniture, chart annotation grammar and the
combo/bubble contracts, split out of outline_validate.py (which re-exports
them under their old `_validate_*` names) to keep that module under the
size ceiling.  Same conventions: append to `errors` for fail-fast issues
(no .pptx is written), `warnings` for warn-and-skip ones.  Nothing here
imports build_deck.
"""
from helpers import parse_visual


def _validate_table(p, where, errors, warnings):
    """Cycle 7 table options (table / comps-table): a units row that does not
    match the column count and threshold-format grammar failures are errors
    (fail-fast); unknown sort/format columns and unknown option values warn
    and the builder falls back."""
    import builders_table as bt
    rows = p.get("table_rows", [])
    header = rows[0]
    n_cols = max(len(r) for r in rows)
    if p.get("units"):
        units = bt.split_units(p["units"])
        if len(units) != n_cols:
            errors.append(
                f"{where}: '- Units:' has {len(units)} entries for {n_cols} "
                "columns — one per column, blanks allowed "
                "('- Units: , $M, %')")
    rules, fmt_errors = bt.parse_format_rules(p.get("formats"))
    errors.extend(f"{where}: {e}" for e in fmt_errors)
    for rule in rules:
        if bt.find_column(rule["column"], header) is None:
            warnings.append(f"{where}: '- Format: {rule['spec']}' names "
                            f"unknown column {rule['column']!r} — rule "
                            "skipped")
    sort_val = (p.get("sort") or "").strip()
    if sort_val and sort_val.lower() not in ("off", "none"):
        col, direction = bt.parse_sort(sort_val)
        if direction is None:
            warnings.append(f"{where}: '- Sort: {sort_val}' — expected "
                            "'<column header> asc|desc'; not sorted")
        elif bt.find_column(col, header) is None:
            warnings.append(f"{where}: '- Sort:' column {col!r} not in the "
                            "table header — not sorted")
    align = (p.get("align") or "").strip().lower()
    if align and align not in bt.ALIGN_MODES:
        warnings.append(f"{where}: '- Align: {align}' unknown "
                        "(auto|left|decimal) — auto used")
    total = (p.get("total_row") or "").strip().lower()
    if total and total not in ("on", "off", "auto"):
        warnings.append(f"{where}: '- Total-Row: {total}' unknown "
                        "(on|auto|off) — auto used")
    zebra = (p.get("zebra") or "").strip().lower()
    if zebra and zebra not in bt.ON_VALUES + bt.OFF_VALUES:
        warnings.append(f"{where}: '- Zebra: {zebra}' unknown (on|off) — "
                        "zebra kept")
    if p.get("summary"):
        _, unknown = bt.parse_summary(p["summary"])
        for t in unknown:
            warnings.append(f"{where}: '- Summary:' stat {t!r} unknown "
                            "(mean|median|quartiles) — skipped")



def _validate_gantt(p, where, errors, warnings):
    """Gantt checks. An unparsable/reversed `**Periods:**` expansion is an
    error (fail-fast, like a missing **Periods:**); malformed Span/Shade/Line
    rows only warn — the builder skips them (warn + skip, never crash)."""
    from builders_gantt import _parse_period_range, expand_periods
    n_periods, expandable = 0, True
    if not p.get("periods"):
        errors.append(f"{where}: gantt needs '**Periods:** Q1, Q2, ...'")
    else:
        try:
            n_periods = len(expand_periods(p["periods"])[0])
        except ValueError as exc:
            errors.append(f"{where}: {exc}")
            expandable = False
    if not p.get("bars"):
        errors.append(f"{where}: gantt needs '- Bar: Row=\"..\" "
                      "Label=\"..\" Start=\"1\" End=\"2\"' rows")
    for bar in p.get("bars", []):
        try:
            s, e = float(bar.get("start", 1)), float(bar.get("end", 1))
            if expandable and not (1 <= s <= e <= n_periods):
                errors.append(f"{where}: Bar {bar.get('label', '')!r} "
                              f"Start/End outside 1-{n_periods}")
        except ValueError:
            errors.append(f"{where}: Bar {bar.get('label', '')!r} has "
                          "non-numeric Start/End")
    for sp in p.get("spans", []):
        if not sp.get("rows"):
            warnings.append(f"{where}: Span needs Rows=\"a, b\" — "
                            "the bracket will be skipped")
    for sh in p.get("shades", []):
        if _parse_period_range(sh.get("periods"), n_periods) is None:
            warnings.append(
                f"{where}: Shade Periods={sh.get('periods', '')!r} is not a "
                f"valid range within 1-{n_periods} — the band will be skipped")
    for ln in p.get("gantt_lines", []):
        try:
            at = float(ln.get("at", ""))
            if not 0.5 <= at <= n_periods + 0.5:
                warnings.append(
                    f"{where}: Line At={at:g} outside periods "
                    f"1-{n_periods} — the line will be skipped")
        except ValueError:
            warnings.append(f"{where}: Line At={ln.get('at', '')!r} is not "
                            "a number — the line will be skipped")



def _validate_chart_options(p, where, errors):
    """Cycle 7 chart annotation grammar (fail-fast on malformed specs;
    unknown names/values stay warn-and-skip at build time)."""
    from chart_annotations import parse_arrow
    for spec in p.get("diff_arrows") or []:
        if parse_arrow(spec) is None:
            errors.append(
                f"{where}: Diff-Arrow {spec!r} did not parse — expected "
                "'CatA, CatB [abs|pct|both]' or 'SerA@Cat -> SerB@Cat'")
    eb = str(p.get("error_bars", "")).strip()
    if eb and len([s for s in eb.split(",") if s.strip()]) != 2:
        errors.append(f"{where}: Error-Bars needs exactly two series names "
                      "('- Error-Bars: Low, High')")

def _validate_combo_bubble(p, where, errors, warnings):
    """chart:combo / chart:bubble contracts (chart_combo.py): combo needs
    2+ series with at least one column left; `- Line-Series:` and
    `- Secondary:` names must be **Series:** names; bubble rows carry
    exactly (x, y, size)."""
    kind = parse_visual(p.get("visual", ""))
    if kind[0] != "chart":
        return
    from chart_combo import parse_names
    names = parse_names(p.get("series"))
    shown = ", ".join(names) or "none"
    if kind[1] == "combo":
        if len(names) < 2:
            errors.append(
                f"{where}: chart:combo needs 2+ **Series:** names — the line "
                "series come from '- Line-Series:' (default: the last name); "
                f"got **Series:** {p.get('series')!r}")
        line_names = parse_names(p.get("line_series"))
        for nm in line_names:
            if nm not in names:
                errors.append(f"{where}: Line-Series {nm!r} is not a "
                              f"**Series:** name ({shown})")
        if names and line_names and set(names) <= set(line_names):
            errors.append(f"{where}: Line-Series names every series — at "
                          "least one must stay a column (use chart:line)")
    elif p.get("line_series"):
        warnings.append(f"{where}: - Line-Series: only applies to "
                        f"chart:combo — ignored on chart:{kind[1]}")
    if p.get("secondary"):
        if kind[1] in ("combo", "line"):
            for nm in parse_names(p["secondary"]):
                if nm not in names:
                    errors.append(f"{where}: Secondary {nm!r} is not a "
                                  f"**Series:** name ({shown})")
        else:
            warnings.append(
                f"{where}: - Secondary: only applies to chart:combo / "
                f"chart:line — ignored on chart:{kind[1]}")
    if kind[1] == "bubble":
        for label, vals in p.get("data") or []:
            if not (isinstance(vals, list) and len(vals) == 3):
                errors.append(
                    f"{where}: chart:bubble rows need exactly 3 values "
                    f"(x, y, size) — row {label!r}; declare "
                    "'**Series:** X, Y, Size' with '- label: x, y, size' rows")
                break

