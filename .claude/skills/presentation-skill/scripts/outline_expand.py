#!/usr/bin/env python3
"""
outline_expand.py — `**Repeat:** data.csv [rows=a-b] [sort=col]` row-templated
slide expansion.

A slide carrying `- **Repeat:** data/initiatives.csv [rows=1-20] [sort=name]`
clones into one slide per data row, with every `{column}` reference in a
string field replaced by that row's value. Expansion runs inside
outline_replay.replay(), right after parse_outline, so build_deck.build(),
diff_deck.diff_outline, qa_checklist.build_checklist and gen_handout.gen_handout
all see the identical expanded slide list and numbering.

Fail-fast: an unresolvable data file, an unknown `{field}` reference or more
than REPEAT_MAX rows is an error (no clones for that slide, build aborts);
more than REPEAT_WARN rows is a warning only. A successful expansion also
emits an `[info] Slide N expands to K slides (file)` line into .warnings so
`--check` reports it (build_deck.build() already prints every warning). A
selection resolving to zero rows drops the slide with a warning instead
(`Slide N: **Repeat:** <file> selected 0 rows — slide dropped`) — `{field}`
references are still checked against the file's header columns even then.

`{accent}...{/}` is helpers.py's inline rich-text markup (generation-guide.md),
not a Repeat field — RESERVED_TOKENS excludes it from substitution/unknown
detection so accent runs pass through untouched on a Repeat slide. A data
column literally named `accent` therefore can never be referenced by
`{accent}` — no `{{accent}}`-style escaping is needed, the token is simply
left alone. `{row}`/`{rows}` are likewise reserved builtins, so a CSV/xlsx
column named `row` or `rows` is shadowed and unreachable too.
"""
import csv
import re
from pathlib import Path

import outline_data

REPEAT_WARN = 60
REPEAT_MAX = 200

# {field} — not preceded/followed by another brace, so '{{literal}}' (a
# double-brace escape) is left untouched here and unwrapped separately.
FIELD_RX = re.compile(r"(?<!\{)\{([A-Za-z0-9_ .-]+)\}(?!\})")

# Names FIELD_RX must never treat as a Repeat field: {accent} is helpers.py's
# rich-text marker ({accent}...{/}), not a column reference. Left verbatim in
# the output and never reported as an unknown field.
RESERVED_TOKENS = frozenset({"accent"})

_BRACKET_RX = re.compile(r"\[(\w+)=([^\]]+)\]")
_ROWS_RX = re.compile(r"^(\d+)-(\d+)$")


def parse_repeat(spec):
    """'data/x.csv [rows=1-20] [sort=name]' -> {"file", "rows": (a,b)|None,
    "sort": str|None}. Raises ValueError on a malformed spec."""
    spec = (spec or "").strip()
    opts = {m.group(1).lower(): m.group(2).strip()
           for m in _BRACKET_RX.finditer(spec)}
    file_part = _BRACKET_RX.sub("", spec).strip()
    if not file_part:
        raise ValueError(f"missing data file in {spec!r}")
    result = {"file": file_part, "rows": None, "sort": opts.get("sort") or None}
    if "rows" in opts:
        m = _ROWS_RX.match(opts["rows"])
        if not m:
            raise ValueError(f"invalid rows= range {opts['rows']!r} "
                             "(expected 'a-b')")
        a, b = int(m.group(1)), int(m.group(2))
        if a < 1 or b < a:
            raise ValueError(f"invalid rows= range {opts['rows']!r}")
        result["rows"] = (a, b)
    return result


def read_rows(path):
    """Header -> value dicts for a Repeat data file (.csv or .xlsx).

    Returns (rows, errors); errors is non-empty (and rows is []) when the
    file can't be parsed at all — a per-row cell problem does not happen
    here because every value round-trips as a string."""
    path = Path(path)
    if path.suffix.lower() in outline_data.XLSX_SUFFIXES:
        xlsx_rows, err = outline_data._read_xlsx_rows(path)
        if err:
            return [], [err]
        if not xlsx_rows:
            return [], []
        header = [str(c).strip() for c in xlsx_rows[0][1]]
        rows = [{header[i]: (cells[i] if i < len(cells) else "")
                for i in range(len(header))}
               for _, cells in xlsx_rows[1:]]
        return rows, []
    try:
        text = path.read_text(encoding="utf-8-sig")
    except (OSError, UnicodeDecodeError) as e:
        return [], [f"unreadable ({e})"]
    reader = csv.DictReader(text.splitlines(), restval="")
    if reader.fieldnames is None:
        return [], []
    return [dict(row) for row in reader], []


def read_header(path):
    """Column names for a Repeat data file (.csv or .xlsx), independent of
    row count. Used to validate `{field}` references even when a
    **Repeat:** selection ends up with zero rows (empty file, or every row
    filtered out by `rows=`) — read_rows() alone can't answer that since an
    empty result carries no row dict to read keys from. Returns [] if the
    file itself can't be parsed; read_rows() already reported that error by
    the time this is called, so a silent [] here is not a hidden failure."""
    path = Path(path)
    if path.suffix.lower() in outline_data.XLSX_SUFFIXES:
        xlsx_rows, err = outline_data._read_xlsx_rows(path)
        if err or not xlsx_rows:
            return []
        return [str(c).strip() for c in xlsx_rows[0][1]]
    try:
        text = path.read_text(encoding="utf-8-sig")
    except (OSError, UnicodeDecodeError):
        return []
    reader = csv.DictReader(text.splitlines(), restval="")
    return list(reader.fieldnames or [])


def substitute(obj, row, *, n, total, unknown):
    """Deep copy of obj with every `{field}` in a string replaced from
    `row` (or the builtins `{row}`/`{rows}` = this clone's 1-based index /
    total clone count). `{{`/`}}` collapse to literal `{`/`}`. `{accent}`
    (RESERVED_TOKENS) is left untouched — it is helpers.py rich-text markup,
    not a Repeat field, and is never reported as unknown. Names not in `row`
    and not a builtin/reserved token are left as-is and added to `unknown`
    (a set the caller inspects afterwards) rather than raising, so every
    reference in the slide is reported in one pass."""
    if isinstance(obj, str):
        def repl(m):
            key = m.group(1)
            if key in RESERVED_TOKENS:
                return m.group(0)
            if key == "row":
                return str(n)
            if key == "rows":
                return str(total)
            if key in row:
                return str(row[key])
            unknown.add(key)
            return m.group(0)
        out = FIELD_RX.sub(repl, obj)
        return out.replace("{{", "{").replace("}}", "}")
    if isinstance(obj, dict):
        return {k: substitute(v, row, n=n, total=total, unknown=unknown)
                for k, v in obj.items()}
    if isinstance(obj, list):
        return [substitute(v, row, n=n, total=total, unknown=unknown)
                for v in obj]
    if isinstance(obj, tuple):
        return tuple(substitute(v, row, n=n, total=total, unknown=unknown)
                    for v in obj)
    return obj  # numbers, bools, None etc. — left alone


def expand_repeats(slides, ctx):
    """Clone every `repeat`-bearing slide into one slide per data row.

    Returns (slides, errors, warnings): slides with each `- **Repeat:**`
    slide replaced in place by its clones (marked `_repeat={"file","row",
    "key"}`, `repeat` dropped); errors for a missing file, an unparsable
    **Repeat:** spec, an unknown `{field}`, or more than REPEAT_MAX rows
    (that slide contributes no clones and no info line); warnings for more
    than REPEAT_WARN rows plus, on success, an `[info] Slide N expands to K
    slides (file)` line so --check surfaces the expansion. A selection that
    resolves to zero rows (empty file, or every row filtered out by `rows=`)
    contributes no clones either, but is a warning rather than a silent
    info line — `Slide N: **Repeat:** <file> selected 0 rows — slide
    dropped` — since an outline slide vanishing is worth a human's
    attention; an unknown `{field}` in that slide is still checked against
    the file's header columns and reported as the usual error."""
    out, errors, warnings = [], [], []
    for i, p in enumerate(slides, 1):
        spec = p.get("repeat")
        if not spec:
            out.append(p)
            continue
        where = (f"Slide {i} (line {p['_line']})" if p.get("_line")
                 else f"Slide {i}")
        try:
            parsed = parse_repeat(spec)
        except ValueError as exc:
            errors.append(f"{where}: **Repeat:** {exc}")
            continue

        path = outline_data._resolve_data_file(parsed["file"], ctx)
        if path is None:
            errors.append(f"{where}: **Repeat:** data file not found: "
                          f"'{parsed['file']}' (tried outline dir, assets "
                          "dir, literal path)")
            continue

        rows, read_errors = read_rows(path)
        if read_errors:
            errors.extend(f"{where}: **Repeat:** '{parsed['file']}' {e}"
                          for e in read_errors)
            continue

        if parsed["sort"]:
            sort_col = parsed["sort"]
            if rows and sort_col not in rows[0]:
                errors.append(f"{where}: **Repeat:** sort column "
                              f"'{sort_col}' not found (columns: "
                              f"{', '.join(rows[0].keys())})")
                continue
            rows = sorted(rows, key=lambda r: r.get(sort_col, ""))

        if parsed["rows"]:
            a, b = parsed["rows"]
            rows = rows[a - 1:b]

        n_rows = len(rows)
        if n_rows > REPEAT_MAX:
            errors.append(f"{where}: **Repeat:** {n_rows} rows exceeds "
                          f"REPEAT_MAX ({REPEAT_MAX}) — narrow rows= or "
                          "split the data file")
            continue
        if n_rows > REPEAT_WARN:
            warnings.append(f"{where}: **Repeat:** {n_rows} rows exceeds "
                            f"REPEAT_WARN ({REPEAT_WARN}) — the deck will "
                            "be large")

        if n_rows == 0:
            # No clones to produce, but a bad {field} reference is still a
            # real authoring error — check it against the file's header
            # columns so it isn't silently hidden behind an empty selection.
            unknown = set()
            header = read_header(path)
            substitute(p, {col: "" for col in header}, n=0, total=0,
                       unknown=unknown)
            if unknown:
                errors.append(f"{where}: **Repeat:** unknown field(s) "
                              f"{{{', '.join(sorted(unknown))}}} — not "
                              f"columns in '{parsed['file']}' (columns: "
                              f"{', '.join(sorted(header))})")
                continue
            warnings.append(f"{where}: **Repeat:** {parsed['file']} "
                            "selected 0 rows — slide dropped")
            continue

        unknown = set()
        clones = []
        for k, row in enumerate(rows, 1):
            clone = substitute(p, row, n=k, total=n_rows, unknown=unknown)
            clone["_repeat"] = {"file": parsed["file"], "row": k,
                                "key": f"r{k}"}
            clone.pop("repeat", None)
            clones.append(clone)

        if unknown:
            cols = sorted(rows[0].keys()) if rows else []
            errors.append(f"{where}: **Repeat:** unknown field(s) "
                          f"{{{', '.join(sorted(unknown))}}} — not columns "
                          f"in '{parsed['file']}' (columns: "
                          f"{', '.join(cols)})")
            continue

        warnings.append(f"[info] Slide {i} expands to {n_rows} slides "
                        f"({parsed['file']})")
        out.extend(clones)
    return out, errors, warnings
