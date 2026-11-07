#!/usr/bin/env python3
"""
outline_data.py — chart/table data carried by an outline slide.

The leaf module behind every consumer that reads numbers out of an outline:
`_parse_data_point` turns a `- label: value` row into `(label, value)` (or a
per-series value list), `_series_count` counts the `**Series:**` names that
decide how many values a row must carry, and `load_data_files` resolves
`- Data-File: x.csv|x.xlsx` rows into `slide['data']` with a provenance line
stamped into the speaker notes.

Split out of build_deck.py so that outline_validate.py (which needs
`_series_count`) no longer reaches back into build_deck through a
function-local import; build_deck re-exports every name under its old
spelling.  Nothing here imports build_deck.
"""
import re
from pathlib import Path

NUM_RE = re.compile(r"-?\d[\d,]*\.?\d*")
# parse_outline's front-matter boundary: deck meta is the `**Key:** value`
# lines before the *first slide heading*. verbatim.outline_declares_locked
# scans the same window, so both read this one constant (kept in this leaf
# module so neither has to import build_deck).
SLIDE_HEADING_PREFIX = "## Slide "


def _series_count(slide):
    """Number of chart series declared via **Series:** (comma-separated names)."""
    series = slide.get("series", "")
    if not series or "," not in series:
        return 1
    return len([n for n in series.split(",") if n.strip()])


def _parse_data_point(text, n_series=1):
    """'2024: $42B' -> ('2024', 42.0); multi: '2024: 42, 38' -> ('2024', [42.0, 38.0]).
    Waterfall sentinels ('total'/'end' closing bar, 'subtotal' running-total
    bar) parse in single- and multi-series mode alike."""
    label, sep, value = text.partition(":")
    if not sep:
        return None
    sentinel = value.strip().lower()
    if sentinel in ("total", "end"):        # waterfall closing bar
        return label.strip().strip('"'), "total"
    if sentinel == "subtotal":              # waterfall running-total bar
        return label.strip().strip('"'), "subtotal"
    if n_series > 1:
        nums = []
        for part in value.split(","):
            m = NUM_RE.search(part)
            if not m:
                return None
            nums.append(float(m.group().replace(",", "")))
        if len(nums) != n_series:
            return None
        return label.strip().strip('"'), nums
    m = NUM_RE.search(value)
    if not m:
        return None
    return label.strip().strip('"'), float(m.group().replace(",", ""))



# ── CSV chart data (- Data-File:) ────────────────────────────────────────────
def _resolve_data_file(spec, ctx):
    """Resolve a Data-File path: outline dir, then assets dir, then as-is."""
    candidates = [Path(ctx.get("outline_dir", ".")) / spec,
                  Path(ctx.get("assets_dir", "assets")) / spec, Path(spec)]
    for c in candidates:
        if c.is_file():
            return c
    return None


def _csv_value(text):
    """Numeric cell like _parse_data_point ($/,/% stripped); 'total'/'end'
    pass through for waterfalls; None when unparsable."""
    t = text.strip()
    if t.lower() in ("total", "end"):
        return "total"
    m = NUM_RE.search(t)
    return float(m.group().replace(",", "")) if m else None


def _parse_csv_rows(rows):
    """rows = [(row_no, cells), ...] -> (data, series_names|None, errors).

    2 columns: label,value rows (an unparsable first row is treated as a
    header). 3+ columns: header row gives series names, body rows give
    per-series value lists."""
    errors = []
    if not rows:
        return [], None, ["no data rows"]
    width = len(rows[0][1])
    if width < 2:
        return [], None, ["needs 2+ columns (label,value)"]
    bad_width = [str(no) for no, cells in rows if len(cells) != width]
    if bad_width:
        return [], None, [f"row {', '.join(bad_width)}: expected {width} "
                          "columns"]
    if width == 2:
        data = []
        for idx, (no, cells) in enumerate(rows):
            value = _csv_value(cells[1])
            if value is None:
                if idx == 0 and len(rows) > 1:
                    continue  # header row
                errors.append(f"row {no}: value {cells[1].strip()!r} "
                              "is not numeric")
                continue
            data.append((cells[0].strip(), value))
        if not data and not errors:
            errors.append("no data rows")
        return data, None, errors
    series = [c.strip() for c in rows[0][1][1:]]
    data = []
    for no, cells in rows[1:]:
        values = [_csv_value(c) for c in cells[1:]]
        bad = [cells[1:][i] for i, v in enumerate(values)
               if not isinstance(v, float)]
        if bad:
            errors.append(f"row {no}: non-numeric value(s) "
                          f"{', '.join(repr(b.strip()) for b in bad)}")
            continue
        data.append((cells[0].strip(), values))
    if not data and not errors:
        errors.append("no data rows after the header")
    return data, series, errors



XLSX_SUFFIXES = (".xlsx", ".xlsm")


def _read_xlsx_rows(path):
    """Active-sheet cells as [(row_no, [str, ...]), ...] like csv.reader
    output. Returns (rows, error) — error set when openpyxl is missing or
    the workbook is unreadable."""
    try:
        from openpyxl import load_workbook
    except ImportError:
        return None, ("needs openpyxl for .xlsx data "
                      "(pip install openpyxl), or export the sheet to CSV")
    try:
        wb = load_workbook(str(path), read_only=True, data_only=True)
    except Exception as e:  # openpyxl raises zipfile/xml/os error variants
        return None, f"unreadable ({e})"
    rows = []
    for no, row in enumerate(wb.active.iter_rows(values_only=True), 1):
        cells = ["" if v is None else str(v) for v in row]
        if any(c.strip() for c in cells):
            rows.append((no, cells))
    wb.close()
    return rows, None


def _stamp_provenance(p, spec, n_rows):
    """Append a data-provenance line to speaker notes so exhibit numbers
    trace back to their source file (evidence-ledger companion)."""
    stamp = f"Data source: {spec} ({n_rows} rows)"
    p["notes"] = f"{p['notes']} {stamp}" if p.get("notes") else stamp


def load_data_files(slides, ctx):
    """Resolve '- Data-File: x.csv|x.xlsx' rows into slide['data']
    (multi-column headers -> slide['series'] unless already set) and stamp
    provenance into speaker notes. Returns (errors, warnings); call right
    after parse_outline in build and --check paths."""
    import csv
    errors, warnings = [], []
    for n, p in enumerate(slides, 1):
        spec = p.get("data_file")
        if not spec:
            continue
        where = (f"Slide {n} (line {p['_line']})" if p.get("_line")
                 else f"Slide {n}")
        path = _resolve_data_file(spec, ctx)
        if path is None:
            errors.append(f"{where}: Data-File not found: '{spec}' "
                          "(tried outline dir, assets dir, literal path)")
            continue
        if path.suffix.lower() in XLSX_SUFFIXES:
            rows, err = _read_xlsx_rows(path)
            if err:
                errors.append(f"{where}: Data-File '{spec}' {err}")
                continue
        else:
            try:
                text = path.read_text(encoding="utf-8-sig")
            except (OSError, UnicodeDecodeError) as e:
                errors.append(f"{where}: Data-File '{spec}' unreadable ({e})")
                continue
            rows = [(no, cells) for no, cells
                    in enumerate(csv.reader(text.splitlines()), 1)
                    if any(c.strip() for c in cells)]
        data, series, errs = _parse_csv_rows(rows)
        if errs:
            errors.extend(f"{where}: Data-File '{spec}': {e}" for e in errs)
            continue
        if p.get("data"):
            warnings.append(f"{where}: both **Data:** rows and Data-File "
                            "given — Data-File wins")
        if series and not p.get("series"):
            p["series"] = ", ".join(series)
        p["data"] = data
        _stamp_provenance(p, spec, len(data))
    return errors, warnings


# ── CJK detection ────────────────────────────────────────────────────────────
