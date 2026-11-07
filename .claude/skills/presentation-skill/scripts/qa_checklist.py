#!/usr/bin/env python3
"""
qa_checklist.py — generate the binary, instance-specific checklist a VLM
answers when it inspects rendered slides.

Open-ended "review this slide" prompts under-report defects; binary
questions naming *this deck's* content catch far more (PresentBench,
SlideAudit).  So the checklist is built from the outline, not from generic
advice: one item per bullet / stat / card / table row that must appear
verbatim, one per evidence-ledger value that must survive unchanged, plus
the handful of generic flaws a renderer can show but geometry cannot
(placeholder stubs, title-vs-content, leaked notes, language drift, bullet
count, chart annotation).  Geometry defects — overlap, misalignment,
crowding, balance — are *not* here: geometry_report.py measures those
deterministically and a VLM only adds noise.

Slide numbers line up with the built deck: the same Data-File /
Auto-Agenda / References transforms build_deck applies are replayed here.

Usage:
    python3 scripts/qa_checklist.py outline.md
    python3 scripts/qa_checklist.py outline.md --evidence evidence.json
    python3 scripts/qa_checklist.py outline.md --output checklist.md
    python3 scripts/qa_checklist.py outline.md --json
"""
import argparse
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

from evidence import load_ledger, parse_refs  # noqa: E402
from outline_replay import make_ctx, replay  # noqa: E402
from textfit import strip_markup  # noqa: E402

MAX_COMPLETENESS_ITEMS = 12   # per slide; the rest fold into one item
QUOTE_CHARS = 90              # content quoted in an item before eliding
MIN_CONTENT_LEN = 3           # shorter strings are labels, not content
MAX_BULLETS = 6               # bullets per slide before the deck reads dense

DECK_ITEM_ID = "QD.1"

# Layouts whose exhibit is drawn from shapes rather than a native chart —
# they need the same units / axis / source annotation a chart does.
EXHIBIT_LAYOUTS = frozenset({
    "waterfall", "mekko", "bar-mekko", "chart-callout", "funnel", "tornado",
    "football-field", "gantt",
})

# Slide-dict keys holding author-typed bullet strings.
BULLET_KEYS = ("bullets", "left_bullets", "right_bullets")
# Slide-dict keys holding lists of key=value content rows.
KV_LIST_KEYS = ("stats", "cards", "items", "steps", "tiles", "matrix_items",
                "milestones", "bars", "nodes", "layers", "lanes")
# kv keys that name an asset or a directive, never rendered copy.
NON_TEXT_KV = frozenset({"icon", "id", "parent", "targetx", "targety"})

# Layouts that are meant to be sparse — asking whether they are a
# "title-only stub" would fail every title slide in every deck.
SPARSE_LAYOUTS = frozenset({"title", "closing", "section-divider",
                            "full-image", "quote"})


def _is_body_slide(slide):
    return slide.get("layout", "") not in SPARSE_LAYOUTS


def _has_bullets(slide):
    return any(slide.get(key) for key in BULLET_KEYS)


def _always(_slide):
    return True


# Generic per-slide flaws: (question, what to look at, when it applies).
# {n} is the slide number.  Order is the order they are asked in.
GENERIC_ITEMS = (
    ("Does slide {n} carry real content — no placeholder or lorem text, and "
     "not a title-only stub", "whole slide", _is_body_slide),
    ("Is the claim in slide {n}'s title actually supported by what the slide "
     "shows", "title vs the body and exhibit", _always),
    ("Is slide {n} free of leaked speaker notes, prompt text or authoring "
     "instructions in the visible body", "all visible text", _always),
    ("Do slide {n}'s spelling, capitalization and terminology match the rest "
     "of the deck", "headings, labels and body copy", _always),
    ("Does slide {n} show at most " + str(MAX_BULLETS) + " bullets",
     "the body list", _has_bullets),
)

CHART_ITEM = ("Does slide {n}'s exhibit carry complete annotation — units, "
              "axis labels and a source line", "the chart and its footnote")

DECK_ITEM = ("Do all exhibits in the deck share one visual style — the same "
             "colors, fonts, label placement and gridline treatment",
             "every chart and exhibit slide, side by side")

HEADER = """\
# Visual QA checklist — {name}

**Protocol.** Answer **one item per call**. Every answer is yes or no with
slide-number evidence — partial credit does not exist, so *partial = no*.
Any item answered no on a delivery slide fails the gate; record the slide
number and fix it in the outline, never in the deck.

Geometry defects (overlap, misalignment, uneven gaps, crowding, visual-mass
balance) are **not** on this list — `geometry_report.py` measures those
deterministically. Judge only what a render can show and geometry cannot.
"""


def _plain(text):
    """Outline copy as it reaches the slide: rich-text markers stripped."""
    stripped, _ = strip_markup(str(text or ""))
    return " ".join(stripped.split())


def _elide(text):
    return text if len(text) <= QUOTE_CHARS else text[:QUOTE_CHARS - 1] + "…"


def content_units(slide):
    """(what-to-look-at, verbatim text) for every author-typed content row.

    Bullets, key=value rows (stats, cards, steps, timeline items, …) and
    table rows — the copy whose loss or truncation is a delivery defect.
    """
    units = []
    for key in BULLET_KEYS:
        for text in slide.get(key, []):
            units.append(("body text", _plain(text)))
    for key in KV_LIST_KEYS:
        for entry in slide.get(key, []):
            text = " · ".join(
                _plain(v) for k, v in entry.items()
                if str(k).lower() not in NON_TEXT_KV and _plain(v))
            if text:
                units.append((key.replace("_", " "), text))
    for row in slide.get("table_rows", []):
        text = " | ".join(_plain(c) for c in row)
        if text:
            units.append(("table row", text))
    return [(where, text) for where, text in units
            if len(text) >= MIN_CONTENT_LEN]


def slide_title(slide):
    return (_plain(slide.get("title")) or _plain(slide.get("heading"))
            or "(untitled)")


def is_exhibit(slide):
    """Slides whose main visual is a chart or a shape-built exhibit."""
    if str(slide.get("visual", "")).strip().lower().startswith("chart"):
        return True
    return slide.get("layout", "") in EXHIBIT_LAYOUTS


def slide_items(slide_no, slide, ledger):
    """Every checklist item for one slide, in ask order."""
    items = []
    title = slide_title(slide)

    def add(kind, question, evidence):
        items.append({"slide": slide_no, "id": f"Q{slide_no}.{len(items) + 1}",
                      "kind": kind, "question": question,
                      "evidence": f"slide {slide_no}, {evidence}",
                      "slide_title": title})

    units = content_units(slide)
    for where, text in units[:MAX_COMPLETENESS_ITEMS]:
        add("completeness",
            f'Does slide {slide_no} show "{_elide(text)}" in full, '
            f"verbatim and untruncated?", where)
    remaining = len(units) - MAX_COMPLETENESS_ITEMS
    if remaining > 0:
        add("completeness",
            f"Are the remaining {remaining} content rows of slide "
            f"{slide_no} all present, verbatim and untruncated?",
            "every remaining bullet, card, stat and table row")

    for ref in parse_refs(slide.get("evidence")):
        entry = ledger.get(ref) if ledger is not None else None
        value = (entry or {}).get("value")
        if not value:
            if ledger is not None:
                reason = ("is not in the ledger" if entry is None
                          else "has no 'value'")
                print(f"qa_checklist: [WARN] slide {slide_no}: evidence "
                      f"{ref} {reason} — no fidelity item generated",
                      file=sys.stderr)
            continue
        source = entry.get("source", "")
        add("fidelity",
            f'Does slide {slide_no} state evidence {ref} as exactly '
            f'"{_plain(value)}", unrounded and unrestated?',
            f"the number and its source footnote ({source})"
            if source else "the number and its source footnote")

    for question, evidence, applies in GENERIC_ITEMS:
        if applies(slide):
            add("generic", question.format(n=slide_no) + "?", evidence)
    if is_exhibit(slide):
        add("generic", CHART_ITEM[0].format(n=slide_no) + "?", CHART_ITEM[1])
    return items


def build_checklist(outline_path, evidence_path=None, profile=None,
                    assets_dir=None):
    """Checklist items for a whole outline, plus the one deck-level item.

    Slide numbering mirrors build_deck's transform order so item ids point
    at the right slide of the built deck.
    """
    outline_path = Path(outline_path)
    rr = replay(outline_path, ctx=make_ctx(outline_path, assets_dir=assets_dir,
                                           profile=profile))
    slides = rr.slides
    is_evidence_error = set(rr.evidence_errors)
    for err in rr.errors:
        if err in is_evidence_error:
            # A broken/unresolvable **Evidence-File:** or ledger reference —
            # distinct from a missing Data-File, so say so rather than
            # pointing the author at their data rows.
            print(f"qa_checklist: [WARN] {err} — evidence ledger problem; "
                  "fidelity items for that reference are missing",
                  file=sys.stderr)
        else:
            # build_deck fails on these; here they only mean the data rows
            # are missing, so say which items the checklist could not
            # generate.
            print(f"qa_checklist: [WARN] {err} — completeness items for "
                  "that slide's data are missing", file=sys.stderr)

    ledger = None
    if evidence_path:
        ledger, errors = load_ledger(evidence_path)
        if errors:
            raise ValueError(
                f"evidence ledger {Path(evidence_path).name} unusable: "
                + "; ".join(errors))

    items = []
    for n, slide in enumerate(slides, 1):
        items.extend(slide_items(n, slide, ledger))
    items.append({"slide": 0, "id": DECK_ITEM_ID, "kind": "generic",
                  "question": DECK_ITEM[0] + "?",
                  "evidence": DECK_ITEM[1], "slide_title": "Deck"})
    return items


def render_markdown(items, name):
    """The checklist as the markdown a reviewing agent works through."""
    lines = [HEADER.format(name=name)]
    current = None
    for item in items:
        if item["slide"] != current:
            current = item["slide"]
            lines.append("")
            lines.append(f"### Slide {current} — {item['slide_title']}"
                         if current else "### Deck")
        lines.append(f"- [ ] {item['id']}: {item['question']} "
                     f"(evidence: {item['evidence']})")
    return "\n".join(lines) + "\n"


JSON_KEYS = ("slide", "id", "question", "kind")


def main():
    parser = argparse.ArgumentParser(description="Build the per-slide binary "
                                                 "QA checklist for an outline.")
    parser.add_argument("outline")
    parser.add_argument("--evidence", default=None,
                        help="Evidence ledger JSON — adds fidelity items")
    parser.add_argument("--output", default=None,
                        help="Write the markdown checklist here (default: "
                             "stdout)")
    parser.add_argument("--json", action="store_true",
                        help="Emit the items as JSON instead of markdown")
    parser.add_argument("--profile", default=None,
                        help="Engagement profile the deck was built with "
                             "(repeat the build's --profile; an outline "
                             "**Profile:** is picked up automatically)")
    parser.add_argument("--assets-dir", default=None,
                        help="Assets root (profiles/, data files); default "
                             "./assets")
    args = parser.parse_args()

    try:
        items = build_checklist(args.outline, args.evidence,
                                profile=args.profile,
                                assets_dir=args.assets_dir)
    except (OSError, ValueError) as exc:
        print(f"qa_checklist: {exc}", file=sys.stderr)
        sys.exit(2)

    if args.json:
        payload = [{k: i[k] for k in JSON_KEYS} for i in items]
        print(json.dumps(payload, indent=2))
        return

    text = render_markdown(items, Path(args.outline).name)
    if args.output:
        Path(args.output).write_text(text, encoding="utf-8")
        slides = len({i["slide"] for i in items if i["slide"]})
        print(f"{args.output}: {len(items)} item(s) across {slides} slide(s)")
    else:
        print(text, end="")


if __name__ == "__main__":
    main()
