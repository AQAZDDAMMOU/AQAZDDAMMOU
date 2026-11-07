#!/usr/bin/env python3
"""
numeric_lint.py — numeric-consistency lint across a built deck.

The checks a partner runs by eye on the last read-through, made
deterministic.  All findings are warnings (a number can be legitimately
unusual; a wrong number is a judgement call the author must make):

  (a) Formatting consistency — one decimal convention per table column,
      one negative convention per slide (parentheses or minus, not both),
      one spelling per unit per slide (M vs MM, K vs k).
  (b) Table totals — a row labelled 'Total' must equal the sum of the rows
      *as displayed*, within the rounding slack those displays allow.
  (c) KPI consistency — the same stat-tile label must carry the same value
      on every slide it appears on ($14.2M on one slide, $14.5M on the next
      is the classic pre-delivery embarrassment).  Values are compared in
      base units, so '$14.2M', '14.2M' and '$14,200,000' are one number.
  (d) Title vs exhibit — a number in an action title must be visible in the
      exhibit below it: a chart value, another number on the slide, or a
      number the reader can derive from the chart (a total, a % change, a
      CAGR, a difference).

Entry point: `check_numbers(prs, issues)`, called from pptx_lint.lint_deck.
"""
from collections import defaultdict

from chart_lint import chart_facts
from numfmt import base_value, find_tokens, parse_token
from qa_check import CHART_UNREADABLE, _slide_title, iter_shapes, safe_chart

EMU_IN = 914400
LABEL_BELOW_IN = 0.6        # caption directly under its big number
LABEL_ABOVE_IN = 0.4        # big-number style: caption above the number
LABEL_SIZE_DROP_PT = 8      # caption must be at least this much smaller
KPI_REL_TOL = 0.005         # 0.5% — two spellings of the same KPI
TITLE_REL_TOL = 0.005       # 0.5% relative …
TITLE_ABS_TOL = 0.05        # … or 0.05 absolute for percentage points
ROUNDING_SLACK = 0.5        # half a unit of the last displayed decimal, per row
MIN_TOTAL_ROWS = 2          # a 'Total' over fewer rows is not worth checking
MINUS_STYLES = ("minus", "dash")


class _TextShape:
    """A text shape reduced to what the numeric checks need."""

    def __init__(self, shape):
        self.text = shape.text_frame.text.strip()
        self.left = (shape.left or 0) / EMU_IN
        self.top = (shape.top or 0) / EMU_IN
        self.right = self.left + (shape.width or 0) / EMU_IN
        self.bottom = self.top + (shape.height or 0) / EMU_IN
        self.size = _max_font_pt(shape)
        self.shape = shape


def _max_font_pt(shape):
    """Largest explicit run size in the shape, or None when unset."""
    best = None
    for para in shape.text_frame.paragraphs:
        for run in para.runs:
            if run.font.size is not None:
                pt = run.font.size.pt
                best = pt if best is None else max(best, pt)
    return best


def _text_shapes(slide):
    return [_TextShape(sh) for sh in iter_shapes(slide.shapes)
            if getattr(sh, "has_text_frame", False)
            and sh.text_frame.text.strip()]


def _tables(slide):
    return [sh.table for sh in iter_shapes(slide.shapes)
            if getattr(sh, "has_table", False)]


def _cell_texts(table):
    for r in range(len(table.rows)):
        for c in range(len(table.columns)):
            yield table.cell(r, c).text


def _slide_tokens(texts, tables):
    """Every number written anywhere on the slide."""
    tokens = []
    for item in texts:
        tokens.extend(find_tokens(item.text))
    for table in tables:
        for text in _cell_texts(table):
            tokens.extend(find_tokens(text))
    return tokens


# ── (a) formatting consistency ──────────────────────────────────────────────
def _check_negative_styles(n, tokens, issues):
    parens = [t for t in tokens if t.neg_style == "parens"]
    minus = [t for t in tokens if t.neg_style in MINUS_STYLES]
    if parens and minus:
        issues["warn"].append(
            f"Slide {n}: mixed negative-number styles — parentheses "
            f"({parens[0].raw}) and minus ({minus[0].raw}); pick one")


def _check_unit_spellings(n, tokens, issues):
    spellings = defaultdict(set)
    for tok in tokens:
        if tok.unit and tok.suffix:
            spellings[tok.unit].add(tok.suffix)
    for unit, seen in sorted(spellings.items()):
        if len(seen) > 1:
            issues["warn"].append(
                f"Slide {n}: mixed unit spellings for {unit} — "
                f"{', '.join(sorted(seen))}; pick one")


def _check_column_decimals(n, table, issues):
    n_rows, n_cols = len(table.rows), len(table.columns)
    for c in range(n_cols):
        header = table.cell(0, c).text.strip() or f"column {c + 1}"
        decimals = {}
        for r in range(1, n_rows):
            tok = parse_token(table.cell(r, c).text)
            if tok is not None and tok.kind != "multiple":
                decimals.setdefault(tok.decimals, tok.raw)
        if len(decimals) > 1:
            shown = ", ".join(decimals[d] for d in sorted(decimals))
            issues["warn"].append(
                f"Slide {n}: table column {header!r} mixes "
                f"{'- and '.join(str(d) for d in sorted(decimals))}-decimal "
                f"numbers ({shown})")


# ── (b) table totals ────────────────────────────────────────────────────────
def _is_total_row(table, r):
    return table.cell(r, 0).text.strip().casefold().startswith("total")


def _check_table_totals(n, table, issues):
    n_rows, n_cols = len(table.rows), len(table.columns)
    if n_rows < MIN_TOTAL_ROWS + 2:  # header + rows + total
        return
    total_rows = [r for r in range(1, n_rows) if _is_total_row(table, r)]
    for tr in total_rows:
        label = table.cell(tr, 0).text.strip()
        for c in range(1, n_cols):
            total = parse_token(table.cell(tr, c).text)
            if total is None or total.kind == "multiple":
                continue
            parts = []
            for r in range(1, n_rows):
                if r in total_rows:
                    continue
                tok = parse_token(table.cell(r, c).text)
                if tok is None or tok.unit != total.unit:
                    parts = []
                    break  # a non-numeric row makes the sum unverifiable
                parts.append(tok)
            if len(parts) < MIN_TOTAL_ROWS:
                continue
            shown_sum = sum(t.value for t in parts)
            dec = total.decimals
            tol = ROUNDING_SLACK * (10 ** -dec) * len(parts)
            if abs(shown_sum - total.value) > tol:
                header = table.cell(0, c).text.strip() or f"column {c + 1}"
                issues["warn"].append(
                    f"Slide {n}: table {label!r} col {header!r} shows "
                    f"{total.value:,.{dec}f} but rows sum to "
                    f"{shown_sum:,.{dec}f}")


# ── (c) KPI consistency across slides ───────────────────────────────────────
def _overlaps_horizontally(a, b):
    return a.left < b.right and b.left < a.right


def _label_for(value_item, texts):
    """The caption belonging to a big number: nearest smaller text below it
    (or just above it, big-number style), overlapping it horizontally."""
    if value_item.size is None:
        return None
    best, best_gap = None, None
    for item in texts:
        if item is value_item or item.size is None:
            continue
        if value_item.size - item.size < LABEL_SIZE_DROP_PT:
            continue
        if not _overlaps_horizontally(value_item, item):
            continue
        if parse_token(item.text) is not None:
            continue  # another number, not a caption
        below = item.top - value_item.bottom
        above = value_item.top - item.bottom
        if -0.05 <= below <= LABEL_BELOW_IN:
            gap = (0, below)
        elif -0.05 <= above <= LABEL_ABOVE_IN:
            gap = (1, above)
        else:
            continue
        if best_gap is None or gap < best_gap:
            best, best_gap = item, gap
    return best.text if best else None


def _collect_kpis(n, texts, kpis):
    for item in texts:
        tok = parse_token(item.text)
        if tok is None:
            continue
        label = _label_for(item, texts)
        if label:
            kpis[label.casefold()].append((n, base_value(tok), item.text, label))


def _check_kpi_consistency(kpis, issues):
    for entries in kpis.values():
        slides = {n for n, _, _, _ in entries}
        if len(slides) < 2:
            continue
        values = [v for _, v, _, _ in entries]
        span = max(values) - min(values)
        if span <= KPI_REL_TOL * max(abs(v) for v in values):
            continue
        label = entries[0][3]
        shown = ", ".join(f"{raw} (slide {n})" for n, _, raw, _ in entries)
        issues["warn"].append(
            f"Slide {min(slides)}: KPI {label!r} has different values across "
            f"slides — {shown}")


# ── (d) title vs exhibit ────────────────────────────────────────────────────
def _series_derivations(values):
    """Numbers a reader can derive from one series, in display units."""
    out = []
    clean = [v for v in values if v is not None]
    if len(clean) < 2:
        return out
    out.append(sum(clean))
    pairs = list(zip(clean, clean[1:])) + [(clean[0], clean[-1])]
    for a, b in pairs:
        out.append(b - a)
        out.append(a - b)
        if a:
            out.append((b / a - 1) * 100)          # % change
            out.append(b / a * 100)                # index vs base
    if clean[0] > 0 and clean[-1] > 0 and len(clean) > 1:
        cagr = (clean[-1] / clean[0]) ** (1 / (len(clean) - 1)) - 1
        out.append(cagr * 100)
    return out


def _same_shape(a, b):
    """python-pptx hands out a fresh proxy per access — compare shape ids."""
    if a is None or b is None:
        return False
    return getattr(a, "shape_id", None) == getattr(b, "shape_id", object())


def _exhibit_numbers(facts, slide, title_shape):
    """Every number the exhibit shows or implies, as plain floats."""
    numbers = []
    for series in facts["series"]:
        values = [v for v in series["values"] if v is not None]
        numbers.extend(values)
        numbers.extend(_series_derivations(values))
    for axis in (facts["val_axis_min"], facts["val_axis_max"]):
        if axis is not None:
            numbers.append(axis)
    for shape in iter_shapes(slide.shapes):
        if _same_shape(shape, title_shape) \
                or not getattr(shape, "has_text_frame", False):
            continue
        for tok in find_tokens(shape.text_frame.text):
            numbers.append(tok.value)
            numbers.append(base_value(tok))
    for table in _tables(slide):
        for text in _cell_texts(table):
            for tok in find_tokens(text):
                numbers.append(tok.value)
                numbers.append(base_value(tok))
    return numbers


def _matches(token, numbers):
    abs_tol = TITLE_ABS_TOL if token.kind == "pct" else 0.0
    for value in (token.value, base_value(token)):
        for candidate in numbers:
            tol = max(TITLE_REL_TOL * abs(candidate), abs_tol)
            if abs(value - candidate) <= tol:
                return True
    return False


def _check_title_vs_exhibit(n, slide, issues):
    frames = [sh for sh in iter_shapes(slide.shapes)
              if getattr(sh, "has_chart", False)]
    charts = [c for c in (safe_chart(sh) for sh in frames) if c is not None]
    if len(charts) < len(frames):
        issues["warn"].append(f"Slide {n}: {CHART_UNREADABLE}")
    if len(charts) != 1:
        return  # a bare title number has no single exhibit to check against
    title_shape, title = _slide_title(slide)
    tokens = find_tokens(title)
    if not tokens:
        return
    numbers = _exhibit_numbers(chart_facts(charts[0]), slide, title_shape)
    for token in tokens:
        if not _matches(token, numbers):
            issues["warn"].append(
                f"Slide {n}: title number {token.raw} not found in exhibit "
                "(no chart value, slide number or derived total/change matches)")


# ── entry point ─────────────────────────────────────────────────────────────
def check_numbers(prs, issues):
    """Append numeric-consistency warnings for the whole deck to `issues`."""
    kpis = defaultdict(list)
    for n, slide in enumerate(prs.slides, 1):
        texts = _text_shapes(slide)
        tables = _tables(slide)
        tokens = _slide_tokens(texts, tables)
        _check_negative_styles(n, tokens, issues)
        _check_unit_spellings(n, tokens, issues)
        for table in tables:
            _check_column_decimals(n, table, issues)
            _check_table_totals(n, table, issues)
        _collect_kpis(n, texts, kpis)
        _check_title_vs_exhibit(n, slide, issues)
    _check_kpi_consistency(kpis, issues)
    return issues
