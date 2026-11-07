#!/usr/bin/env python3
"""
outline_ops.py — split / merge / move / renumber markdown outline slides.

Pure text rewriting: every public function is `md_text -> md_text` (or
returns block-level structure). Non-target slides are reproduced byte for
byte; only the slide(s) an operation targets get their body regenerated.

This intentionally does NOT import build_deck's heavy build machinery for
its own logic — it only reuses `build_deck.FIELD_KEYS` (the `- Key:`
directive table) and, in `main()`, `build_deck.build(..., check_only=True)`
to validate the rewritten outline before exiting.

Outline shape (see references/generation-guide.md):
    optional deck front-matter (`**Key:** value` lines before the first
    `## Slide`), then repeated `## Slide N: heading {attrs}` blocks holding
    `**Key:**` directives, `- Key: value` directives, plain `- ` / `|`
    content lines, and an optional `**Data:**` block. An optional
    `## Appendix` marker line may sit between two slide blocks.
"""
import argparse
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

from build_deck import FIELD_KEYS, build  # noqa: E402

SLIDE_RX = re.compile(r"^## Slide (\d+):(.*)$")
APPENDIX_PREFIX = "## Appendix"

# Trailing heading attributes, e.g. '... {layout=waterfall palette=aurora}'
# (mirrors build_deck.parse_outline's m_attr regex).
ATTR_RX = re.compile(r"\{([^{}]*)\}\s*$")

LAYOUT_LINE_RX = re.compile(r"^\*\*Layout:\*\*\s*(.*)$")
VISUAL_LINE_RX = re.compile(r"^\*\*Visual:\*\*\s*(.*)$")

# Slide-dict list fields that hold enumerable, splittable/mergeable content
# (bullets, cards, items, steps, table rows, tiles) — exported for callers
# (e.g. the comment-driven revision tool's apply_plan) that need to know
# which parsed-slide fields correspond to the content lines this module
# concatenates or divides. outline_ops itself classifies content lines
# directly from raw text (see `_scan_block`), independent of this list.
LIST_KEYS = ("bullets", "cards", "items", "steps", "table_rows", "tiles")
# Output name handed to build_deck.build(check_only=True) after a transform:
# validation returns before anything is written or the name is read, so this
# is a placeholder only — no such file is ever created.
CHECK_ONLY_OUTPUT = "outline-check.pptx"

# Layouts that render exactly one data-driven exhibit; splitting one in half
# would produce two half-charts, so split() refuses these outright (along
# with any slide carrying a **Visual:** chart: or a **Data:** block).
FORBIDDEN_SPLIT_LAYOUTS = frozenset({
    "waterfall", "funnel", "mekko", "bar-mekko", "gantt", "tornado",
    "football-field", "driver-tree", "org-chart", "matrix-2x2",
    "harvey-scorecard", "raci",
})

# '- Key:' directives that merge() combines (n's value + ' ' + m's value)
# instead of dropping m's copy outright.
_MERGE_COMBINE_KEYS = ("Kicker", "Source", "Notes")


def _classify(stripped):
    """Classify a stripped slide-body line: 'meta' ('**Key:**' directive),
    'field' ('- Key: value' directive, a known FIELD_KEYS prefix), 'content'
    (any other '- ' or '|' line), or None (blank/unrecognized — dropped,
    matching build_deck.parse_outline's own silent skip of such lines)."""
    if not stripped:
        return None
    if stripped.startswith("**"):
        return "meta"
    if stripped.startswith("- "):
        item = stripped[2:]
        for key in FIELD_KEYS:
            if item.startswith(key + ":"):
                return "field"
        return "content"
    if stripped.startswith("|"):
        return "content"
    return None


def _scan_block(body_lines):
    """body_lines = a block's raw lines *excluding* the heading line.

    Returns (directive_lines, content_lines, has_data, layout, visual):
    the first two are raw (terminator-included) line lists in original
    order; has_data is True iff a '**Data:**' line is present; layout/visual
    are the lower-cased/raw values of the block's '**Layout:**'/'**Visual:**'
    lines, or None if absent.
    """
    directives, content = [], []
    has_data, layout, visual = False, None, None
    for raw in body_lines:
        stripped = raw.rstrip("\r\n").strip()
        kind = _classify(stripped)
        if kind == "meta":
            directives.append(raw)
            m = LAYOUT_LINE_RX.match(stripped)
            if m:
                layout = m.group(1).strip().lower()
            m = VISUAL_LINE_RX.match(stripped)
            if m:
                visual = m.group(1).strip()
            if stripped.startswith("**Data:**"):
                has_data = True
        elif kind == "field":
            directives.append(raw)
        elif kind == "content":
            content.append(raw)
        # else: blank/unrecognized — dropped
    return directives, content, has_data, layout, visual


def blocks(md_text) -> tuple[str, list[dict]]:
    """Parse md_text into (front_matter_text, [block, ...]).

    Each block is {"n": int|None, "heading": str|None, "lines": [...],
    "start": int}. "lines" are the block's raw lines (terminators included,
    via str.splitlines(keepends=True)) starting with the heading/marker
    line. "heading" is the exact text after 'Slide N:' (None for the
    '## Appendix' pseudo-block, whose "n" is None and whose "lines" are
    reproduced unchanged by render()). "start" is the 1-based line number
    of the heading/marker line.
    """
    lines = md_text.splitlines(keepends=True)
    front, blks, current = [], [], None
    for lineno, raw in enumerate(lines, start=1):
        stripped = raw.rstrip("\r\n")
        m = SLIDE_RX.match(stripped)
        is_appendix = stripped.startswith(APPENDIX_PREFIX)
        if m or is_appendix:
            if current is not None:
                blks.append(current)
            if m:
                current = {"n": int(m.group(1)), "heading": m.group(2),
                           "lines": [raw], "start": lineno}
            else:
                current = {"n": None, "heading": None,
                           "lines": [raw], "start": lineno}
            continue
        if current is not None:
            current["lines"].append(raw)
        else:
            front.append(raw)
    if current is not None:
        blks.append(current)
    return "".join(front), blks


def _eol(raw) -> str:
    """The line terminator of a raw line ('\\r\\n', '\\n', or '' at EOF)."""
    for end in ("\r\n", "\n", "\r"):
        if raw.endswith(end):
            return end
    return ""


def doc_eol(md_text) -> str:
    """The line terminator this outline uses ('\\r\\n' for a CRLF file).

    Taken from the first terminated line, defaulting to '\\n' for a
    single-line file with no terminator at all. Every line an operation
    *builds* (rather than reproduces) must end with this, or a CRLF outline
    comes back with LF lines spliced through it.
    """
    for raw in md_text.splitlines(keepends=True):
        end = _eol(raw)
        if end:
            return end
    return "\n"


def render(front, blocks) -> str:
    """Reassemble front matter + blocks into outline text, renumbering
    '## Slide N:' headings 1..N in the given order. Pseudo-blocks (n=None,
    e.g. '## Appendix') are reproduced verbatim and consume no number.
    Bytes for every line other than a renumbered heading are untouched, and
    a rewritten heading keeps the line terminator it had (so a CRLF outline
    stays CRLF, and a file whose last line has none keeps it that way)."""
    out = [front]
    i = 0
    for b in blocks:
        if b["n"] is None:
            out.append("".join(b["lines"]))
            continue
        i += 1
        out.append(f"## Slide {i}:{b['heading']}{_eol(b['lines'][0])}")
        out.append("".join(b["lines"][1:]))
    return "".join(out)


def renumber(md_text) -> str:
    """Renumber '## Slide N:' headings 1..N in document order. A no-op
    (byte-identical) on an outline whose slides are already sequential."""
    front, blks = blocks(md_text)
    return render(front, blks)


def _find_block(blks, n):
    for i, b in enumerate(blks):
        if b["n"] == n:
            return i
    raise ValueError(f"no '## Slide {n}:' block found")


def move(md_text, n, k) -> str:
    """Move the block numbered n to position k (1-based) among all blocks
    (including any '## Appendix' marker, which occupies a slot like any
    other block); k is clamped to [1, len(blocks)]. Renumbers on output."""
    front, blks = blocks(md_text)
    src = _find_block(blks, n)
    blk = blks.pop(src)
    dest = max(0, min(k - 1, len(blks)))
    blks.insert(dest, blk)
    return render(front, blks)


def _split_heading_attrs(heading):
    """heading = raw text after 'Slide N:' (leading space included).
    Returns (title, attrs) where attrs is '' or the trailing '{...}' text
    verbatim (build_deck.parse_outline only recognizes attrs at line end)."""
    m = ATTR_RX.search(heading)
    if m:
        return heading[:m.start()].strip(), heading[m.start():m.end()]
    return heading.strip(), ""


def _fmt_heading(title, attrs=""):
    text = f" {title.strip()}"
    return f"{text} {attrs}" if attrs else text


def merge(md_text, n, m, *, heading=None, layout=None) -> str:
    """Merge slide m into slide n: keeps n's directives (Kicker/Source/Notes
    combined with m's value via a space join when both are present, or
    m's alone when only m has one), drops m's other directive lines,
    concatenates m's content lines after n's, then removes m and
    renumbers. Raises ValueError when n and m declare different explicit
    **Layout:** values and no `layout=` override is given."""
    front, blks = blocks(md_text)
    eol = doc_eol(md_text)
    i, j = _find_block(blks, n), _find_block(blks, m)
    blk_n, blk_m = blks[i], blks[j]

    dir_n, content_n, _, layout_n, _ = _scan_block(blk_n["lines"][1:])
    dir_m, content_m, _, layout_m, _ = _scan_block(blk_m["lines"][1:])

    if layout is None and layout_n and layout_m and layout_n != layout_m:
        raise ValueError(
            f"slide {n} layout {layout_n!r} != slide {m} layout "
            f"{layout_m!r}; pass layout= to override")

    m_combine = {}
    for raw in dir_m:
        stripped = raw.strip()
        if not stripped.startswith("- "):
            continue
        item = stripped[2:]
        for key in _MERGE_COMBINE_KEYS:
            if item.startswith(key + ":"):
                m_combine[key] = item[len(key) + 1:].strip().strip('"')

    consumed = set()
    merged_directives = []
    layout_written = False
    for raw in dir_n:
        stripped = raw.strip()
        line = raw
        if layout is not None and stripped.startswith("**Layout:**"):
            line = f"**Layout:** {layout}{eol}"
            layout_written = True
        elif stripped.startswith("- "):
            item = stripped[2:]
            for key in _MERGE_COMBINE_KEYS:
                if item.startswith(key + ":") and key in m_combine:
                    n_val = item[len(key) + 1:].strip().strip('"')
                    combined = f"{n_val} {m_combine[key]}".strip()
                    line = f'- {key}: "{combined}"{eol}'
                    consumed.add(key)
        merged_directives.append(line)

    if layout is not None and not layout_written:
        merged_directives.insert(0, f"**Layout:** {layout}{eol}")

    for key in _MERGE_COMBINE_KEYS:
        if key in m_combine and key not in consumed:
            merged_directives.append(f'- {key}: "{m_combine[key]}"{eol}')

    new_heading = _fmt_heading(heading) if heading else blk_n["heading"]
    merged_block = {
        "n": blk_n["n"], "heading": new_heading,
        "lines": ([f"## Slide {blk_n['n']}:{new_heading}{eol}"]
                  + merged_directives + content_n + content_m),
        "start": blk_n["start"],
    }

    blks = [b for b in blks if b is not blk_m]
    idx_n = next(k2 for k2, b in enumerate(blks) if b is blk_n)
    blks[idx_n] = merged_block
    return render(front, blks)


def split(md_text, n, *, at=None, heading=None) -> str:
    """Split slide n into two slides at content-line index `at` (default:
    len(content)//2). Both halves copy every directive line from n; content
    lines (non-directive '- '/'|' lines) before `at` go to the first half,
    the rest to the second. Headings default to '<original title> (1/2)'
    and '(2/2)'; pass heading= to use a different base title. Raises
    ValueError for single-exhibit slides: a **Visual:** chart:, a
    **Data:** block, or one of FORBIDDEN_SPLIT_LAYOUTS."""
    front, blks = blocks(md_text)
    eol = doc_eol(md_text)
    idx = _find_block(blks, n)
    blk = blks[idx]
    directives, content, has_data, layout, visual = _scan_block(
        blk["lines"][1:])

    if has_data or (visual and visual.lower().startswith("chart:")) \
            or layout in FORBIDDEN_SPLIT_LAYOUTS:
        raise ValueError(
            f"slide {n}: layout {layout!r} is a single-exhibit layout "
            "and cannot be split")
    if len(content) < 2:
        raise ValueError(
            f"slide {n}: only {len(content)} content line(s) — nothing to split")

    at = len(content) // 2 if at is None else at
    if not (0 < at < len(content)):
        raise ValueError(
            f"slide {n}: --at {at} out of range for {len(content)} "
            "content line(s)")
    first, second = content[:at], content[at:]

    base_title, attrs = _split_heading_attrs(blk["heading"])
    title = heading.strip() if heading else base_title
    h1 = _fmt_heading(f"{title} (1/2)", attrs)
    h2 = _fmt_heading(f"{title} (2/2)", attrs)

    blk1 = {"n": blk["n"], "heading": h1,
            "lines": ([f"## Slide {blk['n']}:{h1}{eol}"] + directives
                      + first),
            "start": blk["start"]}
    blk2 = {"n": blk["n"], "heading": h2,
            "lines": ([f"## Slide {blk['n']}:{h2}{eol}"] + directives
                      + second),
            "start": blk["start"]}

    new_blks = blks[:idx] + [blk1, blk2] + blks[idx + 1:]
    return render(front, new_blks)


def main():
    parser = argparse.ArgumentParser(
        description="Split/merge/move/renumber a markdown outline.")
    parser.add_argument("outline", help="Path to outline.md")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_split = sub.add_parser("split", help="Split one slide into two")
    p_split.add_argument("n", type=int)
    p_split.add_argument("--at", type=int, default=None,
                         help="Content-line index to split at (default: half)")
    p_split.add_argument("--heading", default=None,
                         help="Base heading for both halves (default: original)")

    p_merge = sub.add_parser("merge", help="Merge slide M into slide N")
    p_merge.add_argument("nm", metavar="N,M")
    p_merge.add_argument("--heading", default=None)
    p_merge.add_argument("--layout", default=None)

    p_move = sub.add_parser("move", help="Move slide N to position K")
    p_move.add_argument("n", type=int)
    p_move.add_argument("k", type=int)

    p_renum = sub.add_parser("renumber", help="Renumber slides 1..N")

    for p in (p_split, p_merge, p_move, p_renum):
        g = p.add_mutually_exclusive_group(required=True)
        g.add_argument("--in-place", action="store_true",
                       help="Overwrite the input outline")
        g.add_argument("--output", default=None, help="Write to this path")

    args = parser.parse_args()
    text = Path(args.outline).read_text(encoding="utf-8")

    try:
        if args.cmd == "split":
            result = split(text, args.n, at=args.at, heading=args.heading)
        elif args.cmd == "merge":
            n_str, _, m_str = args.nm.partition(",")
            if not m_str:
                raise ValueError("merge expects 'N,M'")
            result = merge(text, int(n_str), int(m_str),
                           heading=args.heading, layout=args.layout)
        elif args.cmd == "move":
            result = move(text, args.n, args.k)
        else:
            result = renumber(text)
    except ValueError as exc:
        print(f"  [ERROR] {exc}", file=sys.stderr)
        sys.exit(1)

    out_path = Path(args.outline) if args.in_place else Path(args.output)
    out_path.write_text(result, encoding="utf-8")
    print(f"Wrote: {out_path}")

    ok = build(str(out_path), CHECK_ONLY_OUTPUT, check_only=True)
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
