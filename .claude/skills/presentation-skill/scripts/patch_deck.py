#!/usr/bin/env python3
"""
patch_deck.py — rebuild selected slides from the outline and swap them into
an already-delivered deck, preserving hand edits everywhere else.

A full rebuild clobbers any manual edits the deck picked up after
delivery; this patches only the outline slides you name:

    python3 scripts/patch_deck.py outline.md delivered.pptx "3,7,9" \\
        --output patched.pptx [--palette X] [--assets-dir assets]
    python3 scripts/patch_deck.py outline.md delivered.pptx "4-6" \\
        --output patched.pptx

The outline is rebuilt once in a scratch dir (so palette, section trackers,
and page numbers are computed with full deck context), the requested
slides are extracted from that rebuild, appended into the delivered deck
via the rId-remapping merge engine, and moved into place; the stale slides
are removed.

Every slide `build_deck.build()` produces is stamped with a stable
`ppt/tags/` identity (see slide_tags.py): an outline slug key plus a title
hash. When the delivered deck carries these tags, each selected outline
slide is located by that identity — so a slide a human reordered, or that
now sits after an inserted slide, is still found and replaced correctly.
Only an untagged delivered deck (or --by-position) falls back to the
older contract: the same 1-based deck position in both files, which
requires the outline and delivered slide counts to match exactly.

After patching, run the QA gate and a visual regression against the
last-delivered baseline (references/qa-guide.md) — only the patched
slide(s) should differ.
"""
import argparse
import difflib
import sys
import tempfile
from collections import namedtuple
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from pptx import Presentation

from build_deck import build
from edit_deck import (SLIDE_RT, _gc_unreferenced, _iter_rels,
                       _parse_selection, _part_rels_path,
                       _presentation_parts, _resolve_target, _slide_entries,
                       _write_xml, _xml, extract, pack,
                       remove, reorder, unpack)
from edit_merge import append_decks
from outline_replay import make_ctx, replay
from qa_check import _slide_title_text
from slide_tags import read_tags, slide_key
from slide_tags import title_sha as _title_sha

CLOSEST_TITLES_LIMIT = 3
# The delivered deck records verbatim mode in its slide tags (PK_FLAGS), so a
# rebuild that is not verbatim can be caught. It records nothing about the
# profile it was built with — `--profile` has to be repeated by the caller.
VERBATIM_REFUSAL = (
    "ERROR: delivered deck was built verbatim (PK_FLAGS) — pass --verbatim "
    "or declare **Copy:** locked in the outline")
NOT_VERBATIM_REFUSAL = (
    "ERROR: delivered deck was not built verbatim but the rebuild would be — "
    "rebuild the whole deck once")

# The delivered deck, read once per run: its Presentation, its per-slide
# `read_tags()` mappings, and its slide count.
_DeliveredDeck = namedtuple("_DeliveredDeck", "prs tags n_slides")


def _open_delivered(delivered) -> _DeliveredDeck:
    """Open the delivered deck and read everything patching needs from it."""
    prs = Presentation(str(delivered))
    return _DeliveredDeck(prs, read_tags(prs), len(prs.slides))


def _verbatim_mismatch(outline, delivered, verbatim):
    """The refusal message when the delivered deck's copy mode and this
    rebuild's disagree, else None. `delivered` is a path or an open
    Presentation.

    Patching a verbatim deck from a plain rebuild would silently unlock the
    copy on exactly the slides being replaced — quotes stripped, `**bold**`
    turned into runs — while the rest of the deck stays locked. The mirror
    case locks the copy on only the patched slides of a deck that was built
    plain. Either way the deck ends up half one thing and half the other,
    which a patch cannot fix — a full rebuild can.
    """
    from diff_deck import tagged_verbatim   # the same PK_FLAGS read
    from verbatim import outline_declares_locked
    deck_locked = tagged_verbatim(delivered)
    if verbatim:
        rebuild_locked = True
    else:
        try:
            md_text = Path(outline).read_text(encoding="utf-8")
        except OSError as e:
            # Not a copy-mode question at all — say what actually failed.
            return f"ERROR: outline {outline} unreadable: {e}"
        rebuild_locked = outline_declares_locked(md_text)
    if deck_locked and not rebuild_locked:
        return VERBATIM_REFUSAL
    if rebuild_locked and not deck_locked:
        return NOT_VERBATIM_REFUSAL
    return None


def _closest_titles(outline_title, delivered_prs, limit=CLOSEST_TITLES_LIMIT):
    """Best-effort fuzzy suggestions from the delivered deck's own titles."""
    titles = [t for t in (_slide_title_text(s) for s in delivered_prs.slides)
              if t]
    if not titles:
        return "(none)"
    matches = difflib.get_close_matches(outline_title or "", titles,
                                        n=limit, cutoff=0.0)
    return ", ".join(f"'{m}'" for m in matches) if matches else "(none)"


def _format_positions(positions):
    """[2] -> '2'; [2, 5] -> '2 and 5'; [2, 5, 9] -> '2, 5 and 9'."""
    labels = [str(p) for p in positions]
    if len(labels) <= 1:
        return labels[0] if labels else ""
    return ", ".join(labels[:-1]) + f" and {labels[-1]}"


def _matches_by(tags, field, value):
    """1-based positions where tags[i][field] == value (value must be truthy)."""
    if not value:
        return []
    return [n for n, t in enumerate(tags, 1) if t and t.get(field) == value]


def resolve_targets(outline, delivered, selection, *, assets_dir=None,
                    by_position=False, deck=None, profile=None):
    """Map an outline slide selection onto delivered-deck positions.

    Returns ([(outline_pos, delivered_pos, key), ...], errors). Outline
    positions come from replay(outline).slides + slide_key with a fresh
    `seen` set — the same order build() uses. A tagged delivered deck
    (unless by_position) matches each selected outline slide by PK_KEY,
    falling back to the title hash, so hand reorders/inserts in the
    delivered deck don't break patching. An untagged deck, or
    --by-position, falls back to today's position+count contract.

    `deck` is an already-open `_DeliveredDeck` (see `_open_delivered`) —
    `patch()` passes the one it read for the copy-mode check rather than
    opening the delivered file a second time.
    """
    result = replay(outline, ctx=make_ctx(outline, assets_dir=assets_dir,
                                          profile=profile))
    if result.errors:
        return [], list(result.errors)
    slides = result.slides
    n_outline = len(slides)

    try:
        positions = _parse_selection(selection, n_outline)
    except ValueError as e:
        return [], [f"ERROR: {e}"]

    seen = set()
    keys = [slide_key(p, seen) for p in slides]

    if deck is None:
        deck = _open_delivered(delivered)
    prs, tags, n_delivered = deck.prs, deck.tags, deck.n_slides
    tagged = any(t is not None for t in tags)

    if by_position or not tagged:
        if n_outline != n_delivered:
            return [], [
                f"ERROR: delivered deck has {n_delivered} slide(s) but the "
                f"outline builds {n_outline} — positions cannot be trusted; "
                "reconcile the outline (or use edit_deck.py) and retry"]
        return [(p, p, keys[p - 1]) for p in positions], []

    targets, errors = [], []
    for p in positions:
        slide = slides[p - 1]
        key = keys[p - 1]
        heading = slide.get("heading") or slide.get("title") or ""
        key_matches = _matches_by(tags, "PK_KEY", key)
        matches = key_matches or _matches_by(
            tags, "PK_TITLE_SHA", _title_sha(heading))
        if len(matches) > 1:
            errors.append(
                f"ERROR: outline slide {p} ('{key}') matches delivered "
                f"positions {_format_positions(matches)} — ambiguous; "
                "duplicate the slide in the outline instead, or use "
                "--by-position")
            continue
        if not matches:
            closest = _closest_titles(heading, prs)
            errors.append(
                f"ERROR: outline slide {p} ('{key}') has no tagged match "
                f"in delivered deck (closest titles: {closest})")
            continue
        targets.append((p, matches[0], key))
    return targets, errors


def patch(outline, delivered, selection, output, palette=None,
          assets_dir=None, template=None, by_position=False, profile=None,
          verbatim=False):
    """Returns True on success; no file is written on failure.

    `palette`, `template`, `profile` and `verbatim` are the build options the
    delivered deck was made with: the scratch rebuild has to be given the same
    ones or the patched slides come back styled differently from their
    neighbours. Only verbatim is recoverable from the deck itself (slide tag
    PK_FLAGS) and is refused when it disagrees; a deck built with CLI
    `--profile NAME` records nothing, so that flag must be repeated here —
    an outline that declares `**Profile:** name` needs nothing.
    """
    delivered = Path(delivered)
    if not delivered.is_file():
        print(f"ERROR: {delivered} not found", file=sys.stderr)
        return False
    if Path(output).resolve() == delivered.resolve():
        print("ERROR: --output must differ from the delivered deck "
              "(keep it as the rollback copy)", file=sys.stderr)
        return False
    deck = _open_delivered(delivered)
    refusal = _verbatim_mismatch(outline, deck.prs, verbatim)
    if refusal:
        print(refusal, file=sys.stderr)
        return False

    selection = str(selection)

    with tempfile.TemporaryDirectory() as td:
        td = Path(td)
        rebuilt = td / "rebuilt.pptx"
        print("Rebuilding the deck from the outline (scratch copy)...")
        if not build(outline, rebuilt, palette_key=palette,
                     template_path=template, assets_dir=assets_dir,
                     profile=profile, verbatim=verbatim):
            print("ERROR: outline rebuild failed — fix the outline first",
                  file=sys.stderr)
            return False

        targets, errors = resolve_targets(
            outline, delivered, selection, assets_dir=assets_dir,
            by_position=by_position, deck=deck, profile=profile)
        for e in errors:
            print(e, file=sys.stderr)
        if errors or not targets:
            return False
        targets.sort(key=lambda t: t[0])  # ascending outline position

        n_delivered = deck.n_slides

        outline_sel = ",".join(str(o) for o, _d, _k in targets)
        single = td / "single.pptx"
        if not extract(rebuilt, outline_sel, single):
            return False
        merged = td / "merged.pptx"
        if not append_decks(delivered, single, None, merged):
            return False

        work = td / "unpacked"
        unpack(merged, work)
        replacement_for = _replacement_map(work, targets, n_delivered)

        # New slides land at n_delivered+1 .. n_delivered+k, in target
        # (ascending outline-position) order. Build the final arrangement:
        # each target's delivered position is replaced by its new slide,
        # and the displaced stale positions are pushed to the tail so
        # `reorder` (which requires a full permutation) can place them —
        # they're removed from the tail afterward, highest position first.
        order = list(range(1, n_delivered + 1))
        stale = [d for _o, d, _k in targets]
        for j, (_o, d, _k) in enumerate(targets, start=1):
            order[d - 1] = n_delivered + j
        order_full = order + stale
        if order_full != list(range(1, len(order_full) + 1)) and not reorder(
                work, ",".join(str(i) for i in order_full)):
            return False
        for pos in range(len(order_full), n_delivered, -1):
            if not remove(work, pos):
                return False
        _retarget_slide_links(work, replacement_for)
        _gc_unreferenced(work)  # drop orphaned tags parts + CT overrides
        if not pack(work, output):
            return False

    for o, d, key in targets:
        print(f"outline slide {o} ({key}) -> position {d}")
    print(f"\nPatched {len(targets)} slide(s) → {output} "
          f"({n_delivered} slides)")
    print("Verify: python3 scripts/qa_check.py "
          f"{output} && visual_regress against the delivered baseline")
    return True


def _replacement_map(work, targets, n_delivered):
    """{stale slide part basename: replacement basename} for the swap.

    Read from the merged package's slide order before it is rearranged:
    delivered slides sit at 1..n_delivered, the rebuilt slides follow in
    target order, so target j's stale part (delivered position d) is
    replaced by part n_delivered + j."""
    pres_p, rels_p, _ = _presentation_parts(Path(work))
    entries = _slide_entries(_xml(pres_p), _xml(rels_p))
    fnames = [fname for _el, _rid, fname in entries]
    return {fnames[d - 1]: fnames[n_delivered + j - 1]
            for j, (_o, d, _k) in enumerate(targets, start=1)
            if d <= n_delivered and n_delivered + j <= len(fnames)}


def _retarget_slide_links(work, replacement_for):
    """Re-point navigation hyperlinks that jumped to a swapped-out slide.

    After the stale slides are removed, a kept slide's slide-jump rel
    (agenda row, tracker chip) would dangle — a package-integrity defect.
    Retarget it at the replacement slide; a dangling link to anything else
    is left for qa --integrity to surface (patch never removes other
    slides). The rebuilt slides' own navigation links were stripped by
    extract(); patch() warns about that once."""
    root = Path(work)
    slides_dir = root / "ppt" / "slides"
    retargeted = 0
    for slide_xml in sorted(slides_dir.glob("slide*.xml")):
        part = f"ppt/slides/{slide_xml.name}"
        rels_p = _part_rels_path(root, part)
        if not rels_p.exists():
            continue
        tree, changed = _xml(rels_p), False
        for rel in _iter_rels(tree):
            if rel.get("Type") != SLIDE_RT \
                    or rel.get("TargetMode") == "External":
                continue
            target = _resolve_target(part, rel.get("Target"))
            if (root / target).exists():
                continue
            new = replacement_for.get(Path(target).name)
            if new:
                rel.set("Target", new)
                changed = True
                retargeted += 1
        if changed:
            _write_xml(tree, rels_p)
    if retargeted:
        print(f"  Navigation: {retargeted} slide link(s) re-pointed at "
              "patched slide(s)")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("outline", help="Outline .md (source of truth)")
    parser.add_argument("delivered", help="Delivered .pptx to patch")
    parser.add_argument("selection",
                        help='1-based outline slide position(s) to rebuild '
                             'and swap: "4", "3,7,9", or "2-4"')
    parser.add_argument("--output", required=True)
    parser.add_argument("--palette", default=None)
    parser.add_argument("--template", default=None,
                        help="Template .pptx the deck was built with")
    parser.add_argument("--assets-dir", default=None)
    parser.add_argument("--by-position", action="store_true",
                        help="Match slides by position+count only, even if "
                             "the delivered deck carries PK_KEY tags")
    parser.add_argument("--profile", default=None,
                        help="Engagement profile the deck was built with "
                             "(repeat the build's --profile; an outline "
                             "'**Profile:**' line needs nothing)")
    parser.add_argument("--verbatim", action="store_true",
                        help="Rebuild with locked copy, as the delivered "
                             "deck was built (see **Copy:** locked)")
    args = parser.parse_args()
    ok = patch(args.outline, args.delivered, args.selection, args.output,
               palette=args.palette, assets_dir=args.assets_dir,
               template=args.template, by_position=args.by_position,
               profile=args.profile, verbatim=args.verbatim)
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
