#!/usr/bin/env python3
"""
refresh_deck.py — rebuild only the slides whose bound data changed.

The night-before scenario: the partner sends revised Q4 numbers, the CSVs
behind two exhibits are overwritten, and nothing else about the deck moved.
A full rebuild would clobber every hand edit the delivered deck picked up
after delivery; picking slides by hand needs someone to know which exhibit
reads which file.

    python3 scripts/refresh_deck.py outline.md delivered.pptx \\
        --output refreshed.pptx --log refresh-log.md

Each slide `build_deck.build()` writes carries a `PK_DATA_SHA` tag (see
slide_tags.py): the sha1 of the `- Data-File:` bytes (or the inline `- Data:`
rows) that produced its exhibit. refresh_deck replays the outline, recomputes
that hash per slide, and rebuilds — through `patch_deck.patch`, so hand edits
elsewhere survive — only the slides where the two disagree. Everything else in
the delivered deck is left byte-for-byte alone.

It also answers the question the partner actually asks ("what changed?"):
every number it can pair between the delivered exhibit (chart cache, table
cells, stat tiles) and the refreshed outline is written to a change log:

    ## Slide 4 — revenue-compounds-through-fy25
    - Revenue@FY25: 12 → 14 (+2, +16.7%)

Numbers that exist on only one side are listed under "unpaired" rather than
guessed at — a renamed series or an added category is a content change, not a
data refresh.

Scope: refresh never *inserts* slides. An outline slide with no tagged match
in the delivered deck is reported and skipped; adding a slide is
`patch_deck.py`/`edit_deck.py` work. A delivered deck with no slide tags at
all cannot be matched safely — the CLI exits 2 and asks for one rebuild with
this version first.
"""
import argparse
import sys
from pathlib import Path
from typing import NamedTuple

sys.path.insert(0, str(Path(__file__).resolve().parent))

from pptx import Presentation
from pptx.exc import PythonPptxError

from chart_workbook import cache_series
from helpers import parse_visual
# numeric_lint owns the value/label pairing heuristic used by its KPI check
# (nearest smaller caption below/above a big number, horizontally
# overlapping); stat tiles are read here with exactly that pairing so the two
# tools never disagree about what a tile is called.
from numeric_lint import _label_for, _text_shapes
from numfmt import base_value, parse_token
from outline_replay import make_ctx, replay
from patch_deck import patch
from qa_check import iter_shapes, safe_chart
from slide_tags import data_sha, find_slide, read_tags, slide_key, title_sha

UNTAGGED_MSG = ("delivered deck carries no slide tags — rebuild once with "
                "this version, then refresh")
# Every outline slide missed: the caller is pointing at the wrong deck (or one
# built from a renamed outline) — never a silent "nothing to do".
UNMATCHED_MSG = ("none of the {n} outline slides matched the delivered deck "
                 "— wrong deck, or rebuild once with this version")
ERROR_PREFIX = "ERROR: "
WARN_PREFIX = "WARN: "
EXIT_UNTAGGED = 2
# Why a refresh failed, for the CLI's exit code (only the untagged deck earns
# its own code; everything else is a plain failure).
REASON_UNTAGGED = "untagged"
REASON_FAILED = "failed"
# Series name python-pptx/charts.py falls back to when a single-series chart
# has neither a **Series:** field nor a heading.
DEFAULT_SERIES = "Series 1"
LABEL_JOIN = "@"
# Repeated stat captions / table row labels on one slide: 'Label', 'Label (2)', ...
REPEAT_LABEL_FMT = " ({n})"
LOG_TITLE = "# Refresh change log"
# Values closer than this are the same number written twice (float noise from
# the chart cache round-trip), not a change.
EPSILON = 1e-9
# What opening a delivered deck raises when the file is not a readable .pptx
# (truncated download, a .key renamed, a stripped package).
DECK_READ_ERRORS = (OSError, ValueError, KeyError, PythonPptxError)


# ── which slides moved ──────────────────────────────────────────────────────
class _Scan(NamedTuple):
    """What one outline-vs-delivered comparison found.

    `changed_slides()` publishes the first two fields; `refresh()` needs the
    counts as well, to tell "nothing moved" apart from "nothing matched", and
    the `untagged` flag so the CLI can pick its exit code without reopening
    the deck. `prs` is the delivered deck this scan opened, so `refresh()`
    can read its old numbers off the same Presentation (None when the scan
    never got that far).
    """
    entries: list
    errors: list
    n_slides: int = 0      # outline slides replayed
    n_matched: int = 0     # of those, found in the delivered deck
    untagged: bool = False
    prs: object = None     # the delivered deck, opened once per run


def _is_untagged(tags) -> bool:
    """True when no slide in a `read_tags()` list carries PK_* tags."""
    return not any(t is not None for t in tags)


def _scan(outline, delivered, *, assets_dir=None, force_all=False,
          profile=None) -> _Scan:
    """The full comparison `changed_slides()` summarises (see `_Scan`)."""
    result = replay(outline, ctx=make_ctx(outline, assets_dir=assets_dir,
                                          profile=profile))
    if result.errors:
        return _Scan([], [ERROR_PREFIX + e for e in result.errors])

    try:
        prs = Presentation(str(delivered))
    except DECK_READ_ERRORS as e:
        return _Scan([], [f"{ERROR_PREFIX}{delivered} could not be opened "
                          f"({e})"])
    tags = read_tags(prs)
    if _is_untagged(tags):
        return _Scan([], [ERROR_PREFIX + UNTAGGED_MSG], untagged=True)

    entries, errors, matched = [], [], 0
    seen = set()
    for pos, p in enumerate(result.slides, 1):
        key = slide_key(p, seen)
        heading = p.get("heading") or p.get("title") or ""
        found = find_slide(tags, key, title_sha(heading))
        if found is None:
            errors.append(
                f"{WARN_PREFIX}outline slide {pos} ('{key}') has no tagged "
                "match in the delivered deck — skipped (refresh never "
                "inserts slides)")
            continue
        matched += 1
        old_sha = (tags[found - 1] or {}).get("PK_DATA_SHA", "")
        new_sha = data_sha(p, result.ctx)
        if not force_all and old_sha == new_sha:
            continue
        entries.append({"outline_pos": pos, "delivered_pos": found,
                        "key": key, "old_sha": old_sha, "new_sha": new_sha})
    return _Scan(entries, errors, n_slides=len(result.slides),
                 n_matched=matched, prs=prs)


def changed_slides(outline, delivered, *, assets_dir=None,
                   force_all=False, profile=None) -> tuple[list[dict], list[str]]:
    """([{outline_pos, delivered_pos, key, old_sha, new_sha}], errors).

    Outline positions come from `replay()` + `slide_key()` with a fresh `seen`
    set — the order `build()` uses. Each is located in the delivered deck by
    PK_KEY (falling back to the title hash) and kept when its recomputed
    PK_DATA_SHA differs from the stamped one, or when `force_all` is set.

    `errors` mixes fatal lines (ERROR_PREFIX: the outline does not build, the
    delivered deck is untagged) with reported-and-skipped ones (WARN_PREFIX:
    an outline slide with no match in the delivered deck). This function
    reports per-slide facts only; the policy call on a deck where *nothing*
    matched belongs to `refresh()`, which sees the counts via `_scan()`.
    """
    scan = _scan(outline, delivered, assets_dir=assets_dir,
                 force_all=force_all, profile=profile)
    return scan.entries, scan.errors


# ── the numbers on a delivered slide ────────────────────────────────────────
def _chart_numbers(shape, out):
    """Chart-cache values as 'series@category' (what PowerPoint paints)."""
    cache = cache_series(shape.chart.part)
    categories = cache["categories"]
    for name, values in cache["series"].items():
        for idx, value in enumerate(values):
            if value is None or idx >= len(categories):
                continue
            out[f"{name}{LABEL_JOIN}{categories[idx]}"] = float(value)


def _unique_label(label, seen):
    """`label` the first time, then 'label (2)', 'label (3)', ... per `seen`.

    `seen` is a {label: count} dict owned by the caller — one per slide and
    per exhibit kind. Both the delivered-deck readers and the outline readers
    go through this, in the same positional order (tile order / table row
    order), so a repeated caption or row label pairs up instead of the last
    occurrence silently overwriting the earlier ones.
    """
    n = seen.get(label, 0) + 1
    seen[label] = n
    return label if n == 1 else f"{label}{REPEAT_LABEL_FMT.format(n=n)}"


def _table_numbers(table, out):
    """Body cells as 'rowlabel@columnheader' (repeated row labels: 'row (2)')."""
    n_rows, n_cols = len(table.rows), len(table.columns)
    seen = {}
    for r in range(1, n_rows):
        row_label = table.cell(r, 0).text.strip()
        if not row_label:
            continue
        row_label = _unique_label(row_label, seen)
        for c in range(1, n_cols):
            header = table.cell(0, c).text.strip()
            value = base_value(parse_token(table.cell(r, c).text))
            if header and value is not None:
                out[f"{row_label}{LABEL_JOIN}{header}"] = value


def _stat_numbers(slide, out):
    """Big-number tiles, keyed by the caption paired with each value.

    Tiles sharing a caption are numbered in z-order ('Label', 'Label (2)'),
    which is draw order — the outline's stat order for our builders.
    """
    texts = _text_shapes(slide)
    seen = {}
    for item in texts:
        token = parse_token(item.text)
        if token is None:
            continue
        label = _label_for(item, texts)
        if label:
            out[_unique_label(label.strip(), seen)] = base_value(token)


def old_numbers(prs, pos) -> dict:
    """{label: value} for the exhibit on 1-based slide `pos` of `prs`.

    `prs` is a Presentation (or a path to one). Charts are read from the XML
    cache, tables cell by cell, stat tiles by value/caption pairing — the
    same labels `new_numbers()` derives from the outline, so the two can be
    diffed directly.
    """
    if not hasattr(prs, "slides"):
        prs = Presentation(str(prs))
    slides = list(prs.slides)
    if not 1 <= pos <= len(slides):
        return {}
    slide = slides[pos - 1]
    out = {}
    for shape in iter_shapes(slide.shapes):
        if safe_chart(shape) is not None:
            _chart_numbers(shape, out)
        elif getattr(shape, "has_table", False):
            _table_numbers(shape.table, out)
    _stat_numbers(slide, out)
    return out


# ── the numbers the outline now carries ─────────────────────────────────────
def _series_names(p):
    """The chart's series names, exactly as builders/charts.py name them."""
    field = str(p.get("series") or "")
    if "," in field:
        return [n.strip() for n in field.split(",") if n.strip()]
    return [field or p.get("heading") or p.get("title") or DEFAULT_SERIES]


def _scatter_numbers(p, names, out, kind="scatter"):
    """chart:scatter / chart:bubble rows keyed 'Y@x' — the chart cache's
    spelling (the x value stands in as the category, `%g`-spelled; see
    chart_workbook), not the row label, which an XY chart never plots.
    [Group] rows key by group (the cached series name); a bubble's size
    column is not cached as a series value and is not keyed."""
    from chart_combo import xy_refresh_points
    from chart_workbook import _x_label
    for name, x, y in xy_refresh_points(p, names, kind, names[0]):
        if x is None or y is None:
            continue
        out[f"{name}{LABEL_JOIN}{_x_label(x)}"] = float(y)


def _data_numbers(p, out):
    names = _series_names(p)
    visual = parse_visual(p.get("visual", ""))
    if visual in (("chart", "scatter"), ("chart", "bubble")):
        _scatter_numbers(p, names, out, kind=visual[1])
        return
    for row in p.get("data") or []:
        try:
            label, value = row
        except (TypeError, ValueError):
            continue
        values = value if isinstance(value, (list, tuple)) else [value]
        for idx, item in enumerate(values):
            if not isinstance(item, (int, float)) or isinstance(item, bool):
                continue  # waterfall 'total'/'subtotal' sentinels
            name = names[idx] if idx < len(names) else f"Series {idx + 1}"
            out[f"{name}{LABEL_JOIN}{label}"] = float(item)


def _table_row_numbers(p, out):
    rows = p.get("table_rows") or []
    if len(rows) < 2:
        return
    header = rows[0]
    seen = {}
    for cells in rows[1:]:
        if not cells or not cells[0].strip():
            continue
        row_label = _unique_label(cells[0].strip(), seen)
        for c in range(1, min(len(cells), len(header))):
            name = header[c].strip()
            value = base_value(parse_token(cells[c]))
            if name and value is not None:
                out[f"{row_label}{LABEL_JOIN}{name}"] = value


def _stat_row_numbers(p, out):
    seen = {}
    for stat in p.get("stats") or []:
        label = str(stat.get("label") or "").strip()
        value = base_value(parse_token(stat.get("value")))
        if label and value is not None:
            out[_unique_label(label, seen)] = value


def new_numbers(slide_dict) -> dict:
    """{label: value} for one replayed outline slide.

    Mirrors `old_numbers()`'s labels: chart rows become 'series@category'
    (the series names `builders._place_visual` passes to the chart), markdown
    table rows 'rowlabel@columnheader', and Value=/Label= stats their label.
    """
    out = {}
    _data_numbers(slide_dict, out)
    _table_row_numbers(slide_dict, out)
    _stat_row_numbers(slide_dict, out)
    return out


# ── change log ──────────────────────────────────────────────────────────────
def change_log(entries, old, new) -> dict:
    """{"slides": [...]} — every number that moved, slide by slide.

    `old` and `new` map an entry's `key` to a {label: value} dict (from
    `old_numbers()` / `new_numbers()`). Labels present on both sides with a
    different value become `changes`; labels on one side only are listed
    under `unpaired` rather than guessed at.
    """
    slides = []
    for entry in entries:
        before = old.get(entry["key"], {})
        after = new.get(entry["key"], {})
        changes = []
        for label in sorted(set(before) & set(after)):
            was, now = before[label], after[label]
            if abs(now - was) <= EPSILON:
                continue
            delta = now - was
            changes.append({
                "label": label, "old": was, "new": now, "delta": delta,
                "pct": (delta / abs(was) * 100.0) if was else None,
            })
        slides.append({
            "key": entry["key"], "pos": entry["delivered_pos"],
            "changes": changes,
            "unpaired": {"old": sorted(set(before) - set(after)),
                         "new": sorted(set(after) - set(before))},
        })
    return {"slides": slides}


def _fmt(value):
    """12.0 -> '12'; 16.666 -> '16.67'; 14200000.0 -> '14,200,000'."""
    if value is None:
        return "n/a"
    if abs(value - round(value)) < EPSILON:
        return f"{round(value):,}"
    return f"{value:,.2f}".rstrip("0").rstrip(".")


def _signed(value):
    return ("+" if value >= 0 else "") + _fmt(value)


def _change_line(change):
    pct = change["pct"]
    tail = _signed(change["delta"])
    if pct is not None:
        tail += f", {pct:+.1f}%"
    return (f"- {change['label']}: {_fmt(change['old'])} → "
            f"{_fmt(change['new'])} ({tail})")


def render_log_md(log) -> str:
    """The change log as markdown, ready to paste into an email."""
    lines = [LOG_TITLE, ""]
    slides = log.get("slides") or []
    if not slides:
        lines.append("No slide data changed.")
        return "\n".join(lines) + "\n"
    for slide in slides:
        lines.append(f"## Slide {slide['pos']} — {slide['key']}")
        for change in slide["changes"]:
            lines.append(_change_line(change))
        if not slide["changes"]:
            lines.append("- (no paired number changed)")
        unpaired = slide.get("unpaired") or {}
        if unpaired.get("old"):
            lines.append("- Only in the delivered deck: "
                         + ", ".join(unpaired["old"]))
        if unpaired.get("new"):
            lines.append("- Only in the refreshed outline: "
                         + ", ".join(unpaired["new"]))
        lines.append("")
    return "\n".join(lines).rstrip("\n") + "\n"


# ── the refresh itself ──────────────────────────────────────────────────────
def _describe(entries):
    return ", ".join(f"{e['outline_pos']}→{e['delivered_pos']} ({e['key']})"
                     for e in entries)


def _skipped_note(scan) -> str:
    """' (K outline slide(s) had no tagged match and were skipped)', or ''."""
    skipped = scan.n_slides - scan.n_matched
    if not skipped:
        return ""
    return (f" ({skipped} outline slide(s) had no tagged match and were "
            "skipped)")


def _refresh(outline, delivered, output, *, assets_dir=None, palette=None,
             template=None, profile=None, verbatim=False, force_all=False,
             dry_run=False, log_path=None) -> tuple[bool, str]:
    """`refresh()` plus the failure reason the CLI needs for its exit code."""
    delivered = Path(delivered)
    if not delivered.is_file():
        print(f"{ERROR_PREFIX}{delivered} not found", file=sys.stderr)
        return False, REASON_FAILED

    scan = _scan(outline, delivered, assets_dir=assets_dir,
                 force_all=force_all, profile=profile)
    for message in scan.errors:
        print(message, file=sys.stderr)
    if any(m.startswith(ERROR_PREFIX) for m in scan.errors):
        return False, (REASON_UNTAGGED if scan.untagged else REASON_FAILED)
    if scan.n_slides and not scan.n_matched:
        # Every slide missed: the outline and the deck are not the same deck.
        print(ERROR_PREFIX + UNMATCHED_MSG.format(n=scan.n_slides),
              file=sys.stderr)
        return False, REASON_FAILED
    entries = scan.entries
    if not entries:
        note = _skipped_note(scan) or (" — nothing to refresh (use --all to "
                                       "rebuild every slide anyway).")
        print("No slide data changed" + note)
        return True, ""

    # `_scan` already opened the delivered deck; read the old numbers off it.
    prs = scan.prs if scan.prs is not None else delivered
    old = {e["key"]: old_numbers(prs, e["delivered_pos"]) for e in entries}
    slides = replay(outline, ctx=make_ctx(outline, assets_dir=assets_dir,
                                          profile=profile)).slides
    new = {e["key"]: new_numbers(slides[e["outline_pos"] - 1])
           for e in entries}
    log = change_log(entries, old, new)
    rendered = render_log_md(log)

    print(f"Data changed on {len(entries)} slide(s): {_describe(entries)}"
          + _skipped_note(scan))
    if dry_run:
        print("\n" + rendered)
        print("--dry-run: no deck and no log written.")
        return True, ""

    selection = ",".join(str(e["outline_pos"]) for e in entries)
    if not patch(outline, delivered, selection, output, palette=palette,
                 assets_dir=assets_dir, template=template, profile=profile,
                 verbatim=verbatim):
        return False, REASON_FAILED

    if log_path:
        try:
            Path(log_path).write_text(rendered, encoding="utf-8")
        except OSError as e:
            # The deck is already written and correct — say so, and still
            # print the log below rather than losing it.
            print(f"{ERROR_PREFIX}change log not written to {log_path} ({e})",
                  file=sys.stderr)
        else:
            print(f"Change log → {log_path}")
    print("\n" + rendered)
    return True, ""


def refresh(outline, delivered, output, *, assets_dir=None, palette=None,
            template=None, profile=None, verbatim=False, force_all=False,
            dry_run=False, log_path=None) -> bool:
    """Rebuild the slides whose data moved; True on success.

    Nothing is written when no slide's data changed (there is nothing to
    refresh), when `dry_run` is set, or when the refresh fails — including
    the "wrong deck" case where no outline slide matched at all.

    `palette`/`template`/`profile`/`verbatim` are handed to `patch_deck.patch`
    and must match how the delivered deck was built, or the refreshed slides
    come back styled differently from their neighbours. A deck built with CLI
    `--profile NAME` records nothing about it, so repeat the flag here (an
    outline `**Profile:**` line needs nothing); verbatim *is* recorded in the
    slide tags, and patch() refuses a verbatim deck that would be rebuilt
    without it.
    """
    return _refresh(outline, delivered, output, assets_dir=assets_dir,
                    palette=palette, template=template, profile=profile,
                    verbatim=verbatim, force_all=force_all,
                    dry_run=dry_run, log_path=log_path)[0]


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("outline", help="Outline .md (source of truth)")
    parser.add_argument("delivered", help="Delivered .pptx to refresh")
    parser.add_argument("--output", required=True,
                        help="Refreshed .pptx (not written when nothing "
                             "changed, or under --dry-run)")
    parser.add_argument("--all", dest="force_all", action="store_true",
                        help="Rebuild every tagged slide, not just the ones "
                             "whose data hash moved")
    parser.add_argument("--dry-run", action="store_true",
                        help="Report what would change; write nothing")
    parser.add_argument("--log", dest="log_path", default=None,
                        help="Write the change log to this markdown file")
    parser.add_argument("--palette", default=None)
    parser.add_argument("--assets-dir", default=None)
    parser.add_argument("--template", default=None,
                        help="Template .pptx the deck was built with")
    parser.add_argument("--profile", default=None,
                        help="Engagement profile the deck was built with "
                             "(repeat the build's --profile; an outline "
                             "'**Profile:**' line needs nothing)")
    parser.add_argument("--verbatim", action="store_true",
                        help="Rebuild with locked copy, as the delivered "
                             "deck was built (see **Copy:** locked)")
    args = parser.parse_args()
    ok, reason = _refresh(args.outline, args.delivered, args.output,
                          assets_dir=args.assets_dir, palette=args.palette,
                          template=args.template, profile=args.profile,
                          verbatim=args.verbatim, force_all=args.force_all,
                          dry_run=args.dry_run, log_path=args.log_path)
    if ok:
        sys.exit(0)
    sys.exit(EXIT_UNTAGGED if reason == REASON_UNTAGGED else 1)


if __name__ == "__main__":
    main()
