"""builders_table.py — the table layout family: `table` and `comps-table`.

`build_table_slide` moved here out of builders.py (which re-exports it, so
`builders.build_table_slide` and the `TABLE_*` constants keep resolving).
Cycle 7 (§7, think-cell parity) adds the table options on top of the classic
zebra table:

  - Align: auto|left|decimal    numeric columns right-align (default auto);
                                `decimal` also pads decimals per column to
                                the column max (display only)
  - Units: , $M, %, x           italic muted units row under the header
                                (one entry per column, blanks allowed)
  - Zebra: off                  flat body fills (all bg_deep)
  - Total-Row: on|auto|off      bold last body row + accent1 top rule;
                                auto (default) styles it when the first
                                cell matches /^total/i
  - Format: Margin < 10 red     threshold coloring — matching numeric cells
                                get rag_* colored bold text (repeatable
                                and/or `;`-separated)
  - Summary: mean, median, quartiles   computed rows per numeric column
                                (NM/blank excluded), labels in column 0
  - Sort: <column> asc|desc     body rows sorted by that column's numeric
                                value, NM last; header/units/total/summary
                                rows stay pinned
  - cell_shapes tokens          a body cell that is exactly `harvey:3`,
                                `rag:green`, `bar:64%`, `dots:3/5`, `[x]`
                                draws the shape centred over the cell (the
                                cell text stays empty)
  - NM cells                    NM / n.m. / n/m / N/A / n/a / – / - render
                                muted italic and are excluded from
                                numeric-ness, stats and threshold formats

`comps-table` is `table` preset with `- Align: decimal`,
`- Summary: mean, median` and `- Total-Row: off` (banking comps/precedents).

Two behaviors are on by default (a deliberate upgrade for legacy tables —
`- Align: left` / `- Total-Row: off` restore the old look): numeric columns
right-align, a last row starting "Total" gets total styling, and NM cells
render muted.  A table that trips none of those renders byte-identical to
the classic builder.  Unknown option values warn and fall back — never
crash a build.
Nothing here imports build_deck; `builders` is imported function-locally
(the cell_shapes pattern) because builders.py imports this module at its
bottom to register LAYOUTS.
"""
import re
import statistics

from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

import shape_groups
from cell_shapes import draw_cell_token, parse_cell_token
from helpers import warn
from numfmt import fmt_abs, get_policy, round_half_up
from outline_data import NUM_RE

TABLE_FONT_PT = 14        # design-canvas body size for table cells
TABLE_CELL_PAD_IN = 0.1   # OOXML tcPr marL/marR default, per side
TABLE_MIN_PT = 10.0       # keep cells at or above qa_check.MIN_PT
SUB_FONT_PT = 11          # units and summary rows (design pt)
TABLE_LEFT, TABLE_TOP, TABLE_W = 0.7, 1.9, 11.9   # design-inch frame
ROW_H_IN, MAX_TABLE_H = 0.55, 5.2
TOTAL_RULE_PT = 1.0       # thin accent1 rule above a total row
SUMMARY_DECIMALS_CAP = 2

ALIGN_MODES = ("auto", "left", "decimal")
OFF_VALUES = ("off", "no", "false")     # mirrors builders_exhibits sort-off
ON_VALUES = ("", "on", "yes", "true")
TOTAL_RX = re.compile(r"^total", re.I)
NM_TOKENS = frozenset({"nm", "n.m.", "n/m", "n/a", "–", "-"})
# A numeric-ish cell: NUM_RE core with optional currency prefix / unit suffix.
_CELL_NUM_RX = re.compile(
    r"^([$€£]?)\s*(" + NUM_RE.pattern + r")\s*([%x×]?)$")
FORMAT_RX = re.compile(
    r"^(.+?)\s*(<=|>=|<|>|=)\s*(-?[\d.,]+)\s+(red|amber|green)\s*$", re.I)
STATE_TO_RAG = {"red": "bad", "amber": "mid", "green": "good"}
SUMMARY_STATS = ("mean", "median", "quartiles")


# ── cell classification ──────────────────────────────────────────────────────

def is_nm(text):
    """True for a not-meaningful marker cell (NM, n.m., n/a, en/hyphen dash)."""
    return str(text).strip().lower() in NM_TOKENS


def cell_number(text):
    """The numeric value of a cell ('$4.2', '12%', '1,200', '3.1x') or None
    for NM / blank / plain text."""
    m = _CELL_NUM_RX.match(str(text).strip())
    if not m or is_nm(text):
        return None
    try:
        return float(m.group(2).replace(",", ""))
    except ValueError:
        return None


def cell_decimals(text):
    """Decimals the cell's number displays (0 for non-numeric cells)."""
    m = _CELL_NUM_RX.match(str(text).strip())
    if not m or is_nm(text) or "." not in m.group(2):
        return 0
    return len(m.group(2).split(".", 1)[1])


def is_numericish(text):
    """Numeric-alignment vote: a number, an NM marker, or blank (neutral)."""
    t = str(text).strip()
    return not t or is_nm(t) or cell_number(t) is not None


def numeric_columns(body_rows, n_cols):
    """Per-column flags: every body cell numeric-ish AND at least one real
    number (an all-blank column stays left-aligned)."""
    flags = []
    for c in range(n_cols):
        cells = [(row[c] if c < len(row) else "") for row in body_rows]
        flags.append(all(is_numericish(t) for t in cells)
                     and any(cell_number(t) is not None for t in cells))
    return flags


def pad_decimals(text, dec):
    """Pad a numeric cell's fraction to `dec` places (display only) — the
    `- Align: decimal` mode.  Non-numeric cells pass through untouched."""
    m = _CELL_NUM_RX.match(str(text).strip())
    if not m or is_nm(text) or dec <= 0:
        return text
    prefix, num, suffix = m.groups()
    ip, _, fp = num.partition(".")
    return f"{prefix}{ip}.{fp.ljust(dec, '0')}{suffix}"


# ── option parsing (shared with outline_validate) ────────────────────────────

def split_units(value):
    """'- Units: , $M, %, x' -> ['', '$M', '%', 'x'] (blanks preserved)."""
    return [s.strip() for s in str(value).split(",")]


def parse_format_rules(specs):
    """Threshold formats -> (rules, errors).

    Each entry of `specs` (the parsed `- Format:` list) may hold several
    `;`-separated rules of the form '<column> <op> <value> <state>'.
    A rule dict carries column, op, value, state (rag_* key) and the raw
    spec for messages.  Grammar failures are returned as errors (fail-fast
    in validate); unknown columns are the caller's warning.
    """
    rules, errors = [], []
    for entry in specs or []:
        for spec in str(entry).split(";"):
            spec = spec.strip()
            if not spec:
                continue
            m = FORMAT_RX.match(spec)
            if not m:
                errors.append(
                    f"'- Format: {spec}' does not parse — expected "
                    "'<column> < <= > >= = <value> red|amber|green'")
                continue
            try:
                value = float(m.group(3).replace(",", ""))
            except ValueError:
                errors.append(f"'- Format: {spec}': {m.group(3)!r} is not "
                              "a number")
                continue
            rules.append({"column": m.group(1).strip(), "op": m.group(2),
                          "value": value,
                          "state": STATE_TO_RAG[m.group(4).lower()],
                          "spec": spec})
    return rules, errors


def parse_sort(value):
    """'- Sort: EV/EBITDA desc' -> ('EV/EBITDA', 'desc'); (None, None) when
    the trailing direction is missing or unknown."""
    parts = str(value).strip().rsplit(None, 1)
    if len(parts) == 2 and parts[1].lower() in ("asc", "desc"):
        return parts[0].strip(), parts[1].lower()
    return None, None


def parse_summary(value):
    """'- Summary: mean, quartiles' -> (['mean', 'quartiles'], unknown)."""
    stats, unknown = [], []
    for tok in str(value).split(","):
        t = tok.strip().lower()
        if t:
            (stats if t in SUMMARY_STATS else unknown).append(t)
    return stats, unknown


def find_column(name, header):
    """Casefold-exact header lookup -> column index or None."""
    target = str(name).strip().casefold()
    for i, cell in enumerate(header):
        if str(cell).strip().casefold() == target:
            return i
    return None


def _op_hit(value, op, target):
    return {"<": value < target, "<=": value <= target,
            ">": value > target, ">=": value >= target,
            "=": value == target}[op]


# ── summary rows ─────────────────────────────────────────────────────────────

def _column_values(body_rows, c):
    """Numeric values of a column across body rows (NM/blank excluded)."""
    return [v for v in (cell_number(row[c] if c < len(row) else "")
                        for row in body_rows) if v is not None]


def _column_affixes(body_rows, c):
    """(prefix, suffix) carried onto summary values when every numeric cell
    in the column spells the same ones ('7.5x' columns average to '8.1x')."""
    pairs = set()
    for row in body_rows:
        t = str(row[c] if c < len(row) else "").strip()
        m = _CELL_NUM_RX.match(t)
        if m and not is_nm(t):
            pairs.add((m.group(1), m.group(3)))
    return pairs.pop() if len(pairs) == 1 else ("", "")


def _fmt_stat(value, dec, prefix, suffix):
    """Summary-cell spelling: the deck's numfmt policy when one is installed,
    else the column's own decimals (max seen, capped at 2)."""
    if get_policy() is not None:
        return fmt_abs(value)
    sign = "-" if value < 0 else ""
    return f"{sign}{prefix}{round_half_up(abs(value), dec):,.{dec}f}{suffix}"


_STAT_ROWS = {
    "mean": (("Mean", lambda vals: statistics.fmean(vals)),),
    "median": (("Median", lambda vals: statistics.median(vals)),),
    "quartiles": (
        ("Q1", lambda vals: statistics.quantiles(vals, n=4,
                                                 method="inclusive")[0]),
        ("Q3", lambda vals: statistics.quantiles(vals, n=4,
                                                 method="inclusive")[2]),
    ),
}
_STAT_MIN_N = {"mean": 1, "median": 1, "quartiles": 2}


def summary_rows(stats, body_rows, n_cols, num_cols):
    """The appended `- Summary:` rows, computed per numeric column over the
    body rows handed in (the caller already excludes a total row).  Label in
    column 0; a column without enough values stays blank."""
    per_val = {c: _column_values(body_rows, c)
               for c in range(1, n_cols) if num_cols[c]}
    per_dec = {c: min(max((cell_decimals(row[c] if c < len(row) else "")
                           for row in body_rows), default=0),
                      SUMMARY_DECIMALS_CAP)
               for c in per_val}
    per_affix = {c: _column_affixes(body_rows, c) for c in per_val}
    rows = []
    for stat in stats:
        for label, fn in _STAT_ROWS[stat]:
            cells = [label] + [""] * (n_cols - 1)
            for c, vals in per_val.items():
                if len(vals) >= _STAT_MIN_N[stat]:
                    prefix, suffix = per_affix[c]
                    cells[c] = _fmt_stat(fn(vals), per_dec[c], prefix, suffix)
            rows.append(cells)
    return rows


# ── option resolution ────────────────────────────────────────────────────────

def _resolve_align(p):
    mode = (p.get("align") or "auto").strip().lower()
    if mode not in ALIGN_MODES:
        warn(f"table '- Align: {mode}' unknown (auto|left|decimal) — "
             "auto used")
        return "auto"
    return mode


def _resolve_zebra(p):
    value = (p.get("zebra") or "").strip().lower()
    if value in ON_VALUES:
        return True
    if value in OFF_VALUES:
        return False
    warn(f"table '- Zebra: {value}' unknown (on|off) — zebra kept")
    return True


def _resolve_total(p, body):
    """Index (into `body`) of the styled total row, or None."""
    mode = (p.get("total_row") or "auto").strip().lower()
    if mode not in ("on", "off", "auto"):
        warn(f"table '- Total-Row: {mode}' unknown (on|auto|off) — "
             "auto used")
        mode = "auto"
    if mode == "off" or not body:
        return None
    if mode == "on":
        return len(body) - 1
    last = body[-1]
    return len(body) - 1 if last and TOTAL_RX.match(str(last[0]).strip()) \
        else None


def _resolve_formats(p, header):
    """Threshold rules with resolved column indexes; bad grammar and unknown
    columns warn and drop the rule (validate() already flagged them)."""
    rules, errors = parse_format_rules(p.get("formats"))
    for e in errors:
        warn(f"table {e} — rule skipped")
    resolved = []
    for rule in rules:
        idx = find_column(rule["column"], header)
        if idx is None:
            warn(f"table '- Format: {rule['spec']}' names unknown column "
                 f"{rule['column']!r} — rule skipped")
            continue
        resolved.append((idx, rule["op"], rule["value"], rule["state"]))
    return resolved


def _sorted_body(p, header, body, total_idx):
    """Body rows under `- Sort: <column> asc|desc` — numeric rows ordered by
    the column's value, NM/blank rows after them in author order, a total
    row pinned last.  Returns (rows, total row or None)."""
    total = body[total_idx] if total_idx is not None else None
    rows = [r for i, r in enumerate(body) if i != total_idx]
    spec = (p.get("sort") or "").strip()
    if not spec or spec.lower() in ("off", "none"):
        return rows, total
    col, direction = parse_sort(spec)
    if direction is None:
        warn(f"table '- Sort: {spec}' — expected '<column header> asc|desc'; "
             "not sorted")
        return rows, total
    idx = find_column(col, header)
    if idx is None:
        warn(f"table '- Sort:' column {col!r} not in the header — not sorted")
        return rows, total
    keyed = [(cell_number(r[idx] if idx < len(r) else ""), r) for r in rows]
    numeric = [(v, r) for v, r in keyed if v is not None]
    numeric.sort(key=lambda vr: vr[0], reverse=direction == "desc")
    return [r for _, r in numeric] + [r for v, r in keyed if v is None], total


# ── render plan ──────────────────────────────────────────────────────────────

def _cell(text, kind, c, rules, dec_by_col):
    """Per-cell render record: display text, token, NM flag, rag color key."""
    text = str(text)
    token = rag = None
    nm = False
    if kind in ("body", "total") and text.strip():
        token = parse_cell_token(text)
        if token is None:
            nm = is_nm(text)
            value = cell_number(text)
            if value is not None:
                for idx, op, target, state in rules:
                    if idx == c and _op_hit(value, op, target):
                        rag = state
                if c in dec_by_col:
                    text = pad_decimals(text, dec_by_col[c])
    return {"text": text, "token": token, "nm": nm, "rag": rag}


def _plan_rows(p, rows, n_cols):
    """The rendered row sequence: header, units?, body (sorted, total last),
    summary rows — each a dict of kind + per-cell records."""
    header, body = rows[0], rows[1:]
    align_mode = _resolve_align(p)
    total_idx = _resolve_total(p, body)
    body_rows, total = _sorted_body(p, header, body, total_idx)
    num_cols = numeric_columns(body, n_cols) if align_mode != "left" \
        else [False] * n_cols
    rules = _resolve_formats(p, header)
    dec_by_col = {}
    if align_mode == "decimal":
        dec_by_col = {c: max((cell_decimals(r[c] if c < len(r) else "")
                              for r in body), default=0)
                      for c in range(n_cols) if num_cols[c]}

    def record(row, kind):
        cells = [_cell(row[c] if c < len(row) else "", kind, c,
                       rules, dec_by_col) for c in range(n_cols)]
        return {"kind": kind, "cells": cells}

    plan = [record(header, "header")]
    if p.get("units"):
        units = split_units(p["units"])
        if len(units) != n_cols:
            warn(f"table '- Units:' has {len(units)} entries for {n_cols} "
                 "columns — padded/truncated")
        units = (units + [""] * n_cols)[:n_cols]
        plan.append(record(units, "units"))
    for row in body_rows:
        plan.append(record(row, "body"))
    if total is not None:
        plan.append(record(total, "total"))
    if p.get("summary"):
        stats, unknown = parse_summary(p["summary"])
        for t in unknown:
            warn(f"table '- Summary:' stat {t!r} unknown "
                 "(mean|median|quartiles) — skipped")
        for cells in summary_rows(stats, body_rows, n_cols, num_cols):
            plan.append(record(cells, "summary"))
    return plan, num_cols


_ROW_FILL = {"header": "accent1", "units": "bg_deep", "total": "bg_deep",
             "summary": "bg_deep"}


def _row_fill(rd, pal, zebra_on, zebra_idx):
    if rd["kind"] == "body":
        return pal["surface"] if zebra_on and zebra_idx % 2 == 0 \
            else pal["bg_deep"]
    return pal[_ROW_FILL[rd["kind"]]]


def _cell_style(rd, cd, c, pal):
    """(bold, italic, color, base design-pt) for one cell record."""
    kind = rd["kind"]
    if kind == "header":
        return True, False, pal["bg_deep"], TABLE_FONT_PT
    if kind == "units":
        return False, True, pal["text_muted"], SUB_FONT_PT
    if kind == "summary":
        return c > 0, True, pal["text_muted"], SUB_FONT_PT
    # body / total
    bold = kind == "total"
    if cd["nm"]:
        return bold, True, pal["text_muted"], TABLE_FONT_PT
    if cd["rag"]:
        return True, False, pal[f"rag_{cd['rag']}"], TABLE_FONT_PT
    return bold, False, pal["text"], TABLE_FONT_PT


def _cell_top_rule(cell, hex_color, pt=TOTAL_RULE_PT):
    """Thin top border on a table cell (the total-row rule) — python-pptx has
    no cell-border API, so write a:lnT into tcPr directly."""
    from pptx.oxml import parse_xml
    from pptx.oxml.ns import nsdecls, qn
    tcPr = cell._tc.get_or_add_tcPr()
    existing = tcPr.find(qn("a:lnT"))
    if existing is not None:
        tcPr.remove(existing)
    ln = parse_xml(
        f'<a:lnT {nsdecls("a")} w="{int(Pt(pt))}" cap="flat">'
        f'<a:solidFill><a:srgbClr val="{hex_color}"/></a:solidFill></a:lnT>')
    tcPr.insert(0, ln)   # lnL/lnR are never written, so first slot is valid


# ── builders ─────────────────────────────────────────────────────────────────

def build_table_slide(prs, p, pal, ctx):
    import builders as B
    from palettes import hex_rgb
    from textfit import shrink_to_fit_pt
    slide = B._blank_slide(prs, pal, pal["bg"])
    B._heading(slide, p, pal)
    rows = p.get("table_rows", [])
    if not rows:
        warn("table slide has no rows; rendering as bullet-list")
        return B.build_bullet_slide(prs, p, pal, ctx)
    n_cols = max(len(r) for r in rows)
    plan, num_cols = _plan_rows(p, rows, n_cols)
    n_rows = len(plan)
    table_h = min(ROW_H_IN * n_rows, MAX_TABLE_H)
    tl, tt, tw, th = B._sc(TABLE_LEFT, TABLE_TOP, TABLE_W, table_h)
    frame = slide.shapes.add_table(n_rows, n_cols, Inches(tl), Inches(tt),
                                   Inches(tw), Inches(th))
    table = frame.table
    # Accessibility: mark the header row and row banding on a:tblPr so
    # checkers (and qa_check --accessibility) see them; visual styling
    # below stays custom.
    table.first_row = True
    table.horz_banding = True
    zebra_on = _resolve_zebra(p)
    # PowerPoint ignores <a:normAutofit> inside table cells, so long copy has
    # to be shrunk with a literal run size against the cell's own box.
    col_w_in = [table.columns[c].width / Inches(1) - 2 * TABLE_CELL_PAD_IN
                for c in range(n_cols)]
    tokens, zebra_idx = [], 0
    for r, rd in enumerate(plan):
        row_h_in = table.rows[r].height / Inches(1)
        fill = _row_fill(rd, pal, zebra_on, zebra_idx)
        if rd["kind"] == "body":
            zebra_idx += 1
        for c in range(n_cols):
            cd = rd["cells"][c]
            cell = table.cell(r, c)
            cell.fill.solid()
            cell.fill.fore_color.rgb = hex_rgb(fill)
            tf = cell.text_frame
            tf.word_wrap = True
            para = tf.paragraphs[0]
            if num_cols[c]:
                para.alignment = PP_ALIGN.RIGHT
            if cd["token"] is not None:
                tokens.append((cd["token"], r, c))
                continue          # the shape carries the state; no cell text
            run = para.add_run()
            run.text = cd["text"]
            bold, italic, color, design_pt = _cell_style(rd, cd, c, pal)
            base_pt = B._fs(design_pt)
            fit_pt = shrink_to_fit_pt(run.text, base_pt, col_w_in[c], row_h_in,
                                      bold=bold, family=pal["font_body"],
                                      min_pt=TABLE_MIN_PT)
            run.font.size = Pt(min(fit_pt, base_pt))
            run.font.name = pal["font_body"]
            run.font.bold = bold
            if italic:
                run.font.italic = True
            run.font.color.rgb = hex_rgb(color)
        if rd["kind"] == "total":
            for c in range(n_cols):
                _cell_top_rule(table.cell(r, c), pal["accent1"])
    if tokens:
        _draw_tokens(slide, pal, table, tokens, n_rows, n_cols)
    return slide


def _draw_tokens(slide, pal, table, tokens, n_rows, n_cols):
    """Draw the queued cell_shapes tokens centred over their cell boxes
    (design inches, from the table's own geometry) and group them so a
    token-heavy table cannot trip the L4 ungrouped-exhibit signal — the
    table itself is a graphicFrame and can never join a group."""
    import builders as B
    sx, sy = B._SCALE["x"], B._SCALE["y"]
    col_w = [table.columns[c].width / Inches(1) / sx for c in range(n_cols)]
    row_h = [table.rows[r].height / Inches(1) / sy for r in range(n_rows)]
    col_x, row_y = [TABLE_LEFT], [TABLE_TOP]
    for w in col_w[:-1]:
        col_x.append(col_x[-1] + w)
    for h in row_h[:-1]:
        row_y.append(row_y[-1] + h)
    shapes = []
    for token, r, c in tokens:
        shapes.extend(draw_cell_token(slide, pal, token, col_x[c], row_y[r],
                                      col_w[c], row_h[r]) or [])
    shape_groups.group_shapes(slide, shapes)


COMPS_DEFAULTS = {"align": "decimal", "summary": "mean, median",
                  "total_row": "off"}


def build_comps_table_slide(prs, p, pal, ctx):
    """`comps-table`: the banking comps/precedents preset — decimal-aligned
    numeric columns, Mean/Median summary rows, no total-row styling.  The
    slide's own keys override the preset."""
    q = dict(p)
    for key, value in COMPS_DEFAULTS.items():
        q.setdefault(key, value)
    return build_table_slide(prs, q, pal, ctx)


LAYOUTS = {
    "table": build_table_slide,
    "comps-table": build_comps_table_slide,
}
