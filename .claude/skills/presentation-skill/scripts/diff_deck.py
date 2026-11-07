#!/usr/bin/env python3
"""
diff_deck.py — Compare outline content to built .pptx text (content QA).

Two directions, both required for a clean delivery:

  outline -> deck  every author-typed content row (bullets, card/stat/step
                   fields, table cells) must survive into the deck; a missing
                   one is an **error**. Chrome (title/subtitle/heading) stays a
                   warning because builders sometimes restyle or drop it.
  deck -> outline  every deck paragraph must trace back to the outline, deck
                   meta or known builder chrome; anything else is phantom
                   content (a hand edit or a hallucination) and is warned
                   about — an error under --strict.

Classification rules live in content_trace.py.

Usage:
    python3 scripts/diff_deck.py outline.md deck.pptx
    python3 scripts/diff_deck.py outline.md deck.pptx --strict

Uses markitdown when installed for the outline->deck blobs; the deck->outline
direction always uses python-pptx so paragraph boundaries survive.
"""
import argparse
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

from content_trace import (MIN_TRACE_LEN,  # noqa: E402
                           boilerplate_set, classify_deck_paragraph,
                           deck_corpus, normalize, slide_corpus, word_set)
from outline_replay import make_ctx, replay  # noqa: E402
from verbatim import copy_is_locked  # noqa: E402

PLACEHOLDER_RX = re.compile(
    r"lorem|ipsum|xxxx|click to add|\[ add visual \]|this page layout", re.I)

ADDED_SNIPPET = 60
DROPPED_SNIPPET = 50  # also used for the missing-chrome message
ADDED_MSG = "added content not in outline"
DROPPED_MSG = "dropped outline content"


# ── outline side: what the author typed ──────────────────────────────────────

# Builders truncate: only the first N entries of a collection reach the slide,
# and several print only some of an entry's kv fields. This table mirrors that
# so a row the builder was never going to render is not reported as dropped.
# Caps come from the builders' own geometry helpers/constants where they exist;
# bare literals carry the builder line they were read from. Layouts absent from
# the table render none of these keys (they draw matrix_items / bars /
# milestones / tiles / nodes / layers / lanes, which this check does not cover).
ALL = None                  # builder renders every entry
_BULLET_GEO = "bullet_geo"  # builders.bullet_geometry()["max"]
_TWOCOL_GEO = "twocol_geo"  # builders.twocol_geometry()["max"]
_EXEC_MAX = "exec_max"      # builders.density()["exec_max"]
_CHAIN_STEPS = "chain_steps"  # builders_structure.MAX_CHAIN_STEPS

# Slots = the kv fields a builder prints. A slot is either a precedence group
# (the first field present in the group is the one rendered — mirrors the
# builders' `a or b` fallbacks) or, where the choice is conditional, a callable
# that yields an entry's rendered values exactly as the builder derives them.
# SPLIT_COLON marks plain bullets a builder splits at the first colon into two
# rendered fields, so the whole bullet never appears verbatim on the slide.
SPLIT_COLON = "split_colon"
STAT_SLOTS = (("value",), ("label",), ("sublabel",))


def logo_fields(raw):
    """Mirror builders_banking.build_logo_wall_slide: only the display name
    reaches the slide — the image path is a directive, not deck copy."""
    from builders_banking import parse_logo
    yield parse_logo(raw)[1]


def card_fields(card):
    """Mirror _build_cards (builders.py:516-519). `subtitle` is rendered only
    when the card has no `name` — with a name, the second line is the title
    and any subtitle is dropped by the builder."""
    yield card.get("name") or card.get("title", "")
    yield card.get("title", "") if card.get("name") else card.get("subtitle", "")
    yield card.get("bio") or card.get("body") or card.get("desc", "")

LAYOUT_CONTENT = {
    "agenda":            {"bullets": 8},            # builders_agenda.AGENDA_MAX_ROWS
    "bullet-list":       {"bullets": _BULLET_GEO},  # builders.py:447
    "bullets":           {"bullets": _BULLET_GEO},  # builders.py:447
    "cards-3":           {"cards": (3, card_fields)},  # builders.py:490 + :531
    "cards-4":           {"cards": (4, card_fields)},  # builders.py:490 + :535
    "comps-table":       {"table_rows": ALL},       # builders_table.py
    "exec-summary":      {"bullets": _EXEC_MAX},    # builders.py:409
    "exec-summary-scqa": {"bullets": 4},            # builders_consulting.py:348
    "harvey-scorecard":  {"table_rows": ALL},       # builders_consulting.py:105
    "heatmap-table":     {"table_rows": ALL},       # builders_consulting.py:436
    # league-table: header + MAX_LEAGUE_ROWS body rows (builders_banking.py)
    "league-table":      {"table_rows": 11},
    "logo-wall":         {"logos": (24, logo_fields)},  # builders_banking.py
    "next-steps":        {"bullets": 6,             # builders_consulting.py:294
                          "steps": (6,              # builders_consulting.py:291
                                    (("action", "title"), ("owner",),
                                     ("when",)))},
    "process-flow":      {"bullets": 6,             # builders_consulting.py:160
                          "steps": (6,              # builders_consulting.py:157
                                    (("title",), ("desc",)))},
    "quote-evidence":    {"stats": (1,              # builders_consulting.py:274
                                    (("value",), ("label",)))},
    "raci":              {"table_rows": ALL},       # builders_consulting.py:694
    "stat-callout":      {"stats": (3, STAT_SLOTS)},   # builders.py:539
    "table":             {"table_rows": ALL},       # builders_table.py
    # timeline reads bullets only when no '- Date=' items were given.
    "timeline":          {"bullets": (5, SPLIT_COLON, "items"),  # builders.py:569-572
                          "items": (5,              # builders.py:568
                                    (("date",), ("title",), ("desc",)))},
    "tombstone-grid":    {"tombstones": (12,        # builders_banking.py
                                         (("client",), ("deal",), ("value",),
                                          ("date",), ("role",)))},
    "two-column-split":  {"bullets": _TWOCOL_GEO},  # builders.py:472
    "value-chain":       {"steps": (_CHAIN_STEPS,   # builders_structure.py:222
                                    (("title",), ("sub",)))},
    "what-how-who":      {"steps": (5,              # builders_grid.py
                                    (("title",), ("what",), ("how",),
                                     ("who",)))},
}


def resolve_cap(cap):
    """Late-bind the density/geometry-derived caps — density() is a process
    global that tests and --density flip at runtime."""
    import builders
    if cap == _BULLET_GEO:
        return builders.bullet_geometry()["max"]
    if cap == _TWOCOL_GEO:
        return builders.twocol_geometry()["max"]
    if cap == _EXEC_MAX:
        return builders.density()["exec_max"]
    if cap == _CHAIN_STEPS:
        from builders_structure import MAX_CHAIN_STEPS
        return MAX_CHAIN_STEPS
    return cap


def _slot_values(entries, slots):
    for entry in entries:
        if callable(slots):
            yield from (str(v) for v in slots(entry) if v)
            continue
        for slot in slots:
            for key in slot:
                if entry.get(key):
                    yield str(entry[key])
                    break


ICON_PREFIX_RX = re.compile(r"^icon:[a-z0-9-]+\s+")


def raw_form(text):
    """Verbatim compare form: collapse whitespace and drop the `icon:` bullet
    directive (never deck copy) — nothing else. Case, wrapping quotes and
    `**bold**`/`{accent}…{/}` markers all survive, so any deviation between
    the locked outline and the deck shows up as a dropped row."""
    return re.sub(r"\s+", " ", ICON_PREFIX_RX.sub("", str(text or "").strip()))


def _strip_cell_token(text):
    """Drop a leading cell-state token ('rag:green On track' -> 'On track') —
    like an icon: prefix it renders as a marker shape, never as deck copy.
    The grammar stays cell_shapes' own, reached through builders."""
    from builders import split_cell_token
    return split_cell_token(str(text))[1]


def _clean(parts, verbatim=False):
    """Normalize, drop the too-short-to-match, de-duplicate, keep order."""
    form = raw_form if verbatim else normalize
    seen, out = set(), []
    for part in parts:
        norm = form(_strip_cell_token(part)) if part else ""
        if len(norm) >= MIN_TRACE_LEN and norm not in seen:
            seen.add(norm)
            out.append(norm)
    return out


def content_rows(slide, verbatim=False):
    """Author-typed content rows the slide's builder actually renders —
    losing one of these is a delivery defect. Under verbatim the rows keep
    their exact spelling (see raw_form) instead of being normalized."""
    rows = []
    for key, rule in LAYOUT_CONTENT.get(slide.get("layout", ""), {}).items():
        rule = rule if isinstance(rule, tuple) else (rule, None)
        cap, slots = rule[0], rule[1]
        # A third element names the key this one is only a fallback for: the
        # builder ignores it entirely when that key carries entries.
        if len(rule) > 2 and slide.get(rule[2]):
            continue
        entries = slide.get(key, [])
        cap = resolve_cap(cap)
        if cap is not None:
            entries = entries[:cap]
        if key == "table_rows":
            cells = (cell for row in entries for cell in row)
            if slide.get("layout") in ("table", "comps-table"):
                # A body cell that is exactly a cell_shapes token renders as
                # a shape with no cell text — it can never be "dropped".
                from cell_shapes import parse_cell_token
                cells = (c for c in cells if parse_cell_token(c) is None)
            rows.extend(cells)
        elif slots == SPLIT_COLON:
            for bullet in entries:
                head, _, tail = str(bullet).partition(":")
                rows.extend((head, tail))
        elif slots:
            rows.extend(_slot_values(entries, slots))
        else:
            rows.extend(entries)
    return _clean(rows, verbatim)


def chrome_text(slide, verbatim=False):
    """Deck chrome — builders may restyle or drop it (ghost mode), so a miss
    here is a warning, not an error. Under verbatim it is compared raw like
    the content rows (a locked title keeps its quotes); the finding stays a
    warning because dropping chrome is a builder's prerogative."""
    layout = slide.get("layout", "")
    parts = [slide.get(k) for k in ("title", "subtitle", "caption", "contact")]
    # "## Slide N: Title" section labels are not deck copy — skip on title/closing.
    if slide.get("heading") and layout not in ("title", "closing"):
        parts.append(slide["heading"])
    return _clean(parts, verbatim)


# ── deck side: text extraction ───────────────────────────────────────────────

def _as_prs(prs_or_path):
    """An open Presentation from either an already-open one or a path.

    Every deck-side reader below takes either, so one run of `diff_outline`
    opens the .pptx once and hands the same Presentation to all of them.
    """
    if hasattr(prs_or_path, "slides"):
        return prs_or_path
    from pptx import Presentation
    return Presentation(str(prs_or_path))


def _iter_shapes(shapes):
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    for shape in shapes:
        yield shape
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from _iter_shapes(shape.shapes)


def extract_slide_paragraphs(prs_or_path):
    """Per-slide list of visible paragraphs (text frames + table cells).

    Speaker notes and chart-internal text are excluded on purpose: notes are
    not on the slide, and chart labels live in chart XML, not text frames.
    """
    prs = _as_prs(prs_or_path)
    per_slide = []
    for slide in prs.slides:
        texts = []
        for shape in _iter_shapes(slide.shapes):
            if getattr(shape, "has_text_frame", False):
                texts.extend(p.text.strip()
                             for p in shape.text_frame.paragraphs)
            if getattr(shape, "has_table", False):
                texts.extend(c.text.strip() for c in shape.table.iter_cells())
        per_slide.append([t for t in texts if t])
    return per_slide


def _pptx_texts(prs_or_path):
    """Per-slide text blob straight from python-pptx, unnormalized
    (text frames + table cells + speaker notes)."""
    prs = _as_prs(prs_or_path)
    texts = []
    for slide in prs.slides:
        chunks = []
        for shape in _iter_shapes(slide.shapes):
            if getattr(shape, "has_text_frame", False):
                t = shape.text_frame.text.strip()
                if t:
                    chunks.append(t)
            if getattr(shape, "has_table", False):
                chunks.extend(c.text.strip() for c in shape.table.iter_cells()
                              if c.text.strip())
        if slide.has_notes_slide:
            notes = slide.notes_slide.notes_text_frame.text.strip()
            if notes:
                chunks.append(notes)
        texts.append(" ".join(chunks))
    return texts


def extract_pptx_texts(pptx_path, raw=False, prs=None):
    """Return list of per-slide text blobs (speaker notes included).

    raw=True (verbatim mode) keeps the deck's own spelling — no lowercasing,
    no markup stripping — so the compare is byte-for-byte. markitdown is
    skipped there: it rewrites text into markdown.

    `prs` is an already-open Presentation for `pptx_path`, used for the
    python-pptx reads; markitdown still needs the file itself.
    """
    source = pptx_path if prs is None else prs
    if raw:
        return [raw_form(blob) for blob in _pptx_texts(source)]
    try:
        from markitdown import MarkItDown
        md = MarkItDown()
        result = md.convert(str(pptx_path), extract_pages=True)
        if getattr(result, "pages", None):
            return [normalize(p.content) for p in result.pages]
    except ImportError:
        pass
    except Exception as e:
        print(f"  [WARN] markitdown failed: {e}", file=sys.stderr)

    return [normalize(blob) for blob in _pptx_texts(source)]


# ── checks ───────────────────────────────────────────────────────────────────

def check_evidence_values(meta, slides, built, ctx, errors, warnings,
                          verbatim=False):
    """Every ledger value a slide references must appear in the built deck's
    text for that slide, unchanged (a transcription slip on a material
    number is a delivery-blocking defect)."""
    from evidence import load_for_outline, parse_refs
    form = raw_form if verbatim else normalize   # must match `built`
    ledger, load_errs = load_for_outline(meta, ctx)
    if ledger is None:
        return
    if load_errs:
        warnings.extend(load_errs)
        return
    for n, slide in enumerate(slides, 1):
        blob = built[n - 1] if n <= len(built) else ""
        for ref in parse_refs(slide.get("evidence")):
            value = ledger.get(ref, {}).get("value")
            if value and form(value) not in blob:
                errors.append(
                    f"Slide {n}: evidence {ref} value {value!r} missing "
                    "from the built deck — material number did not survive")


def check_added_content(slides, paragraphs, meta, ctx, errors, warnings,
                        strict):
    """Deck -> outline: paragraphs that trace back to nothing the author
    wrote and are not known builder chrome."""
    deck = deck_corpus(meta, slides, ctx)
    deck_words = word_set(deck)
    # what-how-who gutter labels and the cell_shapes checkbox tick are
    # builder chrome — printed by the builders, never author copy.
    boilerplate = boilerplate_set(meta, slides) | {"what", "how", "who", "✓"}
    for n in range(1, len(paragraphs) + 1):
        # Deck slides past the end of the outline (an appended or duplicated
        # slide) have no corpus of their own — everything on them is judged
        # against the deck-wide vocabulary alone.
        slide = slides[n - 1] if n <= len(slides) else {}
        corpus = slide_corpus(slide)
        corpus_words = word_set(corpus)
        seen = set()
        for text in paragraphs[n - 1]:
            norm = normalize(text)
            if norm in seen:
                continue
            if classify_deck_paragraph(text, corpus, deck, boilerplate,
                                       deck_words, corpus_words) != "added":
                continue
            seen.add(norm)
            msg = f"Slide {n}: {ADDED_MSG}: {text[:ADDED_SNIPPET]!r}"
            (errors if strict else warnings).append(msg)


def _has_native_chart(prs) -> bool:
    """True when any slide carries a native pptx chart (graphicFrame)."""
    from qa_check import iter_shapes, safe_chart
    return any(safe_chart(shape) is not None
               for slide in prs.slides
               for shape in iter_shapes(slide.shapes))


def check_workbook_drift(prs_or_path, errors, warnings, strict):
    """Cache-vs-embedded-workbook drift on native charts (chart_workbook.py):
    an analyst edit in PowerPoint's "Edit Data" sheet that never made it
    back into the rendered chart cache, or vice versa. A warning by default
    — an error under --strict, same as the rest of this module's deck-side
    findings.

    Without openpyxl (an optional dependency) the check cannot run. That is
    reported the way `qa_check --integrity` reports a missing optional tool:
    a single `[INFO]` line on stderr, and only for a deck that actually has
    a native chart — never a deck-level warning, which every QA run of every
    deck would otherwise carry and `qa_summary` would count.

    `prs_or_path` is an open Presentation (what `diff_outline` passes) or a
    path to one."""
    prs = _as_prs(prs_or_path)
    try:
        import openpyxl  # noqa: F401
    except ImportError:
        if _has_native_chart(prs):
            print("  [INFO] pip install openpyxl for the embedded workbook "
                  "drift check (skipped)", file=sys.stderr)
        return
    import chart_workbook
    for slide_no, msgs in chart_workbook.deck_drift(prs).items():
        for msg in msgs:
            text = f"Slide {slide_no}: embedded workbook drift — {msg}"
            (errors if strict else warnings).append(text)


def tagged_verbatim(prs_or_path):
    """True when the deck says it was built with locked copy (slide tag
    PK_FLAGS carries `verbatim`) — even if the outline no longer says so.

    `prs_or_path` is an open Presentation or a path (`read_tags` takes
    either)."""
    from slide_tags import read_tags
    return any("verbatim" in (t.get("PK_FLAGS") or "")
               for t in read_tags(prs_or_path) if t)


def diff_outline(outline_path, pptx_path, strict=False, verbatim=None,
                 profile=None, assets_dir=None):
    """verbatim: None auto-detects (the deck's tags, then the outline's own
    `**Copy:** locked`); True/False force the mode. Under verbatim the two
    sides are compared raw — quotes and rich-text markers included — and
    every dropped content row is an error regardless of `strict`.

    profile: the engagement profile the deck was built with when it came
    from a CLI `--profile` flag (an outline `**Profile:**` is picked up by
    replay() on its own) — it can repick auto-chosen layouts and set the
    density, both of which move the truncation caps compared here.
    assets_dir: where `<assets>/profiles/`, data files and palettes live
    (default `assets/`, like build_deck)."""
    outline_path = Path(outline_path)
    # Mirror the transform sequence build() applies so slide positions line
    # up (Auto-Agenda / References insert slides; Data-File fills data;
    # the profile/density decide layouts and caps).
    # Data-file/evidence resolution errors are build's concern, not diff's.
    ctx = make_ctx(outline_path, assets_dir=assets_dir, profile=profile)
    prs = _as_prs(pptx_path)   # every reader below takes it — open once
    if verbatim is None and tagged_verbatim(prs):
        verbatim = True
    ctx["verbatim"] = verbatim   # tri-state: None lets the outline decide
    rr = replay(outline_path, ctx=ctx)
    meta, slides = rr.meta, rr.slides
    if verbatim is None:
        verbatim = copy_is_locked(meta.get("copy"))
    built = extract_pptx_texts(pptx_path, raw=verbatim, prs=prs)
    errors, warnings = [], []
    # A profile that failed to resolve (typo'd --profile, wrong --assets-dir)
    # means this diff ran at default density/layouts — say so rather than
    # letting the resulting "dropped rows" stand unexplained.
    warnings.extend(f"replay: {e} — diff ran without the profile"
                    for e in rr.errors if "profile" in e)
    check_evidence_values(meta, slides, built, ctx, errors, warnings, verbatim)

    if len(built) != len(slides):
        warnings.append(
            f"Slide count mismatch: outline {len(slides)} vs deck {len(built)}")

    for n, slide in enumerate(slides, 1):
        where = f"Slide {n}"
        blob = built[n - 1] if n <= len(built) else ""
        if PLACEHOLDER_RX.search(blob):
            errors.append(f"{where}: leftover placeholder text in deck")
        for phrase in content_rows(slide, verbatim):
            if phrase not in blob:
                errors.append(f"{where}: {DROPPED_MSG}: "
                              f"{phrase[:DROPPED_SNIPPET]!r}")
        for phrase in chrome_text(slide, verbatim):
            if phrase not in blob:
                msg = (f"{where}: expected text not found in deck: "
                       f"{phrase[:DROPPED_SNIPPET]!r}")
                (errors if strict else warnings).append(msg)

    check_added_content(slides, extract_slide_paragraphs(prs), meta, ctx,
                        errors, warnings, strict)
    check_workbook_drift(prs, errors, warnings, strict)
    return errors, warnings


def build_parser():
    parser = argparse.ArgumentParser(description="Diff outline vs built deck text.")
    parser.add_argument("outline")
    parser.add_argument("pptx")
    parser.add_argument("--strict", action="store_true",
                        help="Treat missing/added text as errors")
    parser.add_argument("--verbatim", action=argparse.BooleanOptionalAction,
                        default=None,
                        help="Force the locked-copy (raw) compare on or off; "
                             "default: detect from the deck tags / outline")
    parser.add_argument("--profile", default=None,
                        help="Engagement profile the deck was built with "
                             "(repeat the build's --profile; an outline "
                             "**Profile:** is picked up automatically)")
    parser.add_argument("--assets-dir", default=None,
                        help="Assets root (profiles/, data files); default "
                             "./assets")
    return parser


def main():
    args = build_parser().parse_args()

    errors, warnings = diff_outline(args.outline, args.pptx, args.strict,
                                    verbatim=args.verbatim,
                                    profile=args.profile,
                                    assets_dir=args.assets_dir)
    for w in warnings:
        print(f"  [WARN]  {w}")
    for e in errors:
        print(f"  [ERROR] {e}")
    findings = errors + warnings
    added = sum(1 for f in findings if ADDED_MSG in f)
    dropped = sum(1 for f in findings if DROPPED_MSG in f)
    print(f"\n{Path(args.pptx).name}: {len(errors)} error(s), "
          f"{len(warnings)} warning(s) — {dropped} dropped, {added} added")
    sys.exit(1 if errors else 0)


if __name__ == "__main__":
    main()
