"""
content_trace.py — Does a line of deck text trace back to what the author wrote?

The vocabulary behind diff_deck's deck->outline direction:

  slide_corpus(slide)            every string a parsed slide dict carries
  deck_corpus(meta, slides, ctx) the same across the deck, plus deck meta,
                                 section names and evidence-ledger citations
  boilerplate_set(meta, slides)  chrome strings builders print themselves
  classify_deck_paragraph(...)   'traced' | 'boilerplate' | 'added'

Kept separate from diff_deck.py so the classification rules (and the
constants that calibrate them) can be read and tested on their own.
"""
import re

from textfit import strip_markup

# Shortest corpus string usable for substring tracing — shorter entries
# ("on", "no") would match almost any paragraph.
MIN_TRACE_LEN = 4
# Words at least this long must be traceable for the word-set rule to pass.
MIN_WORD_LEN = 4
# A corpus entry may only account for a *longer* deck paragraph if it is a
# substantial phrase — otherwise a 4-char token (a chart year) traces any
# sentence containing it.
MIN_VOUCH_LEN = 12
MIN_VOUCH_WORDS = 3
# Longest alpha token that still counts as a label abbreviation ("q3", "bn").
MAX_ABBREV_LEN = 3

# Deck chrome the builders print themselves, plus page/exhibit numbering.
# (`b·` alone is the appendix page prefix run left beside a slidenum field.)
BOILERPLATE_RX = re.compile(
    r"^(\d+|b·\d*|source:.*|exhibit \d+|data source:.*|backup|agenda"
    r"|notes?|\+?[-–]?\d[\d.,]*%?|[a-z]{1,3})$", re.I)
# 'Source: …', 'Data source: …' and 'Exhibit 3 · Source: …' footnotes —
# with `- Note:` the line leads 'Note: … · Source: …' (or stands alone).
SOURCE_RX = re.compile(
    r"^(exhibit\s+\d+\s*[·|.\-–—]\s*)?(note:.*·\s*)?(data\s+)?"
    r"sources?\s*[:·]|^note\s*[:·]", re.I)
# slide_chrome.add_footnotes lines: superscript number + footnote text.
FOOTNOTE_BLOCK_RX = re.compile(r"^[⁰¹²³⁴⁵⁶⁷⁸⁹]{1,2}\s")
# slide_chrome.add_file_stamp: '<output filename> · <date>'.
FILE_STAMP_RX = re.compile(r"^\S[^·]*\.pptx\s*·", re.I)
# Footnote body markers: the outline's `[1]` renders as a superscript digit
# run, so both spellings normalize to the bare digits.
MARKER_RX = re.compile(r"\[(\d{1,2})\]")
# Chart annotations builders compute: 'CAGR +32.4%', '+18%', '$4.2B', 'Δ 12pp'.
NUMERIC_LABEL_RX = re.compile(
    r"^(?:(?:cagr|δ|delta|total|avg|average|median|value)\s*)?"
    r"[+\-–—]?\s*[$€£¥]?\d[\d.,]*\s*(?:%|x|bps|pp|bn|[kmbt])?$", re.I)
# Fixed labels emitted by builders that never appear in the outline.
BUILDER_LABELS = frozenset({
    "total", "subtotal", "before", "after", "thank you", "agenda", "backup",
    "appendix", "source", "sources", "notes", "note", "situation",
    "complication", "question", "findings", "recommendation", "benchmark",
    "value", "exhibit",
    # data-row placeholder cells (chart_data_rows.py: en dash / NM)
    "\u2013", "-", "nm",
    # data-row fixed labels (chart_data_rows.py)
    "growth %", "cagr p.a.", "values",
    # gantt Responsible/Remark column headers (builders_gantt.py)
    "responsible", "remark",
    # table summary-row labels (builders_table.py; Q1/Q3 pass the
    # numeric-token allowance below)
    "mean", "median",
    # cell_shapes checkbox tick — the [x] token's rendered glyph
    "✓",
    # harvey-scorecard legend (builders_consulting.py:149)
    "empty = not met · quarter fills = partial · solid = fully met",
})

# Slide-dict keys whose values are directives/asset names, never deck copy.
NON_TEXT_KEYS = frozenset({
    "icon", "layout", "palette", "visual", "visual_left", "visual_right",
    "data_file", "evidence", "labels_mode", "sort", "scale", "marker",
    "axis_max", "axis_break", "cagr", "value_line", "columns", "hidden",
})
# Deck-level meta values rendered as chrome on some or every slide.
DECK_META_TEXT_KEYS = ("footer", "stamp", "takeaway", "purpose")

WORD_RX = re.compile(r"[a-z0-9]+")
ALPHA_RX = re.compile(r"[a-z]+")
DIGIT_RX = re.compile(r"\d")


def normalize(text):
    """Lowercase, collapse whitespace and drop icon: prefixes plus rich-text
    markers (**…**, {accent}…{/}) — they never reach the rendered deck.
    `[1]` footnote markers collapse to their digits, matching the
    superscript run they render as."""
    plain, _ = strip_markup(text or "")
    plain = MARKER_RX.sub(r"\1", plain)
    return re.sub(r"\s+", " ", plain.lower()).strip()


def fmt_num(value):
    """Render a parsed data value the way it was typed ('42', not '42.0')."""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def words(norm_text):
    return [w for w in WORD_RX.findall(norm_text) if len(w) >= MIN_WORD_LEN]


def word_set(corpus):
    out = set()
    for entry in corpus:
        out.update(words(entry))
    return out


# ── corpora ──────────────────────────────────────────────────────────────────

def slide_corpus(slide):
    """Every string a slide dict carries (nested lists/dicts, kv items, data
    labels and values, table cells), normalized.

    Directive-only keys (icon names, layout, visual specs) are skipped — they
    render as images or geometry, never as text.
    """
    corpus = set()

    def walk(value):
        if isinstance(value, str):
            norm = normalize(value)
            if norm:
                corpus.add(norm)
        elif isinstance(value, bool):
            return
        elif isinstance(value, (int, float)):
            corpus.add(fmt_num(value))
        elif isinstance(value, dict):
            for key, val in value.items():
                key = str(key)
                if key.startswith("_") or key.lower() in NON_TEXT_KEYS:
                    continue
                walk(val)
        elif isinstance(value, (list, tuple, set)):
            for item in value:
                walk(item)

    walk(slide)
    return corpus


def meta_strings(meta):
    return {n for n in (normalize(str(meta.get(k) or ""))
                        for k in DECK_META_TEXT_KEYS) if n}


def section_names(slides):
    """Section-divider headings — reused as agenda chips and tracker tabs."""
    return {n for n in (normalize(s.get("heading") or s.get("title") or "")
                        for s in slides
                        if s.get("layout") == "section-divider") if n}


def ledger_strings(meta, ctx):
    """Citation/source/value strings from the evidence ledger — they reach the
    deck as source footnotes and speaker-note citations."""
    if not meta.get("evidence_file") or ctx is None:
        return set()
    from evidence import load_for_outline
    ledger, errors = load_for_outline(meta, ctx)
    if not ledger or errors:
        return set()
    return {n for entry in ledger.values() for n in
            (normalize(str(v)) for v in entry.values()) if n}


def deck_corpus(meta, slides, ctx=None):
    """Deck-wide vocabulary: every slide's corpus plus deck meta values,
    section names and evidence-ledger citations."""
    corpus = set()
    for slide in slides:
        corpus |= slide_corpus(slide)
    return corpus | meta_strings(meta) | section_names(slides) \
        | ledger_strings(meta, ctx)


def boilerplate_set(meta, slides):
    """Normalized strings that are deck chrome rather than slide copy."""
    return set(BUILDER_LABELS) | meta_strings(meta) | section_names(slides)


# ── classification ───────────────────────────────────────────────────────────

def is_boilerplate(norm, boilerplate):
    if norm in boilerplate:
        return True
    if BOILERPLATE_RX.match(norm) or SOURCE_RX.match(norm) \
            or NUMERIC_LABEL_RX.match(norm) or FOOTNOTE_BLOCK_RX.match(norm) \
            or FILE_STAMP_RX.match(norm):
        return True
    # Numeric-token allowance: a label whose only words are abbreviations
    # around a number ('q3 2026', '$14.2m', '256 h100s').
    return bool(DIGIT_RX.search(norm)) and all(
        len(a) <= MAX_ABBREV_LEN for a in ALPHA_RX.findall(norm))


def vouches_for(entry, norm):
    """Does an outline string `entry` account for a longer deck paragraph?

    Only a substantial phrase may: a short or purely numeric corpus entry (a
    chart year, an axis label) otherwise vouches for any sentence that happens
    to contain it, which is exactly how fabricated copy slips through. The
    match must also land on word boundaries.
    """
    if len(entry) < MIN_VOUCH_LEN or len(WORD_RX.findall(entry)) < MIN_VOUCH_WORDS:
        return False
    if not ALPHA_RX.search(entry):
        return False
    return re.search(rf"(?<!\w){re.escape(entry)}(?!\w)", norm) is not None


def is_traced(norm, corpus, deck_words, corpus_words=None):
    if norm in corpus:
        return True
    for entry in corpus:
        # Truncation direction (the deck shows part of a longer outline row)
        # stays lenient; the containment direction is gated by vouches_for.
        if len(entry) >= MIN_TRACE_LEN and norm in entry:
            return True
        if vouches_for(entry, norm):
            return True
    para_words = set(words(norm))
    if corpus_words is None:
        corpus_words = word_set(corpus)
    return bool(para_words) and para_words <= (corpus_words | deck_words)


def classify_deck_paragraph(text, corpus, deck, boilerplate, deck_words=None,
                            corpus_words=None):
    """'traced' (comes from this slide's outline content), 'boilerplate'
    (deck chrome / a computed chart label) or 'added' (phantom content).

    `corpus` is the slide's corpus, `deck` the deck-wide one; pass
    `deck_words` / `corpus_words` to reuse the word sets instead of
    recomputing them for every paragraph.
    """
    norm = normalize(text)
    if not norm:
        return "boilerplate"
    if deck_words is None:
        deck_words = word_set(deck)
    if is_traced(norm, corpus, deck_words, corpus_words):
        return "traced"
    if is_boilerplate(norm, boilerplate):
        return "boilerplate"
    return "added"
