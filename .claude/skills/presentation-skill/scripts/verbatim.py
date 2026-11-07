#!/usr/bin/env python3
"""
verbatim.py — the fail-fast check behind verbatim copy mode.

`**Copy:** locked` (or `--verbatim`) promises the deck carries the author's
approved copy unchanged: build_deck.parse_outline stops stripping wrapping
quotes, helpers.parse_rich_segments stops turning `**bold**`/`{accent}…{/}`
into runs, and diff_deck compares raw text instead of normalized text.

The one silent edit left is a builder truncating a list that overruns its
layout's capacity — normally a warning ("only the first N render"). Under
verbatim that is a build error, raised here so build_deck.validate() writes
no .pptx at all.

Every "is this deck verbatim?" answer comes from `copy_is_locked()` here —
the parse-time detector and the post-parse one must never disagree, or a
build ends up half-verbatim (quotes kept, markup still parsed, no cap gate).
"""

from outline_data import SLIDE_HEADING_PREFIX

COPY_KEY = "Copy"          # build_deck.META_KEYS["Copy"] -> meta["copy"]
LOCKED = "locked"

# Slide fields that are deck copy: verbatim keeps their wrapping quotes.
# Everything else a `- Key: value` line can carry (Data-File, Evidence,
# Scale, Sort, CAGR, …) is a directive, where quotes are quoting syntax and
# must still come off — `- Data-File: "q3.csv"` has to resolve to a path.
COPY_FIELDS = frozenset({
    "title", "subtitle", "heading", "caption", "contact", "notes", "source",
    "quote", "attribution", "context", "value", "label", "callout",
    "current", "situation", "recommendation", "sticker", "kicker", "support",
    "left_label", "right_label", "x_axis", "y_axis", "q1", "q2", "q3", "q4",
})

# Caps deliberately kept out of diff_deck.LAYOUT_CONTENT: adding them there
# would change the normal-mode diff for existing decks. Same shape as that
# table (layout -> {key: cap}), with the builder line each was read from.
SUPPLEMENTAL_CAPS = {
    "matrix-2x2":      {"matrix_items": 12},  # builders_consulting.py:49
    "stakeholder-map": {"matrix_items": 12},  # builders_consulting.py:652
    "dashboard":       {"tiles": 6},          # builders_consulting.py:234
}


def copy_is_locked(value):
    """The single test for the `**Copy:**` directive: exactly 'locked'.

    Anything else ('locked (legal approved)') leaves verbatim off — the mode
    is all-or-nothing, so a value that only looks locked must not half-enable
    it. build_deck.validate() warns about such a value.
    """
    return str(value or "").strip().lower() == LOCKED


def outline_declares_locked(md_text):
    """True when the outline's front matter carries `**Copy:** locked`.

    Front matter = the `**Key:** value` lines before the first slide heading,
    read exactly the way build_deck.parse_outline reads META_KEYS: the same
    boundary constant, the same partition, the same quote stripping — so this
    pre-parse detector and the parsed `meta["copy"]` cannot disagree. Any
    other `## ` line before the first slide (a stray `## Note`, the
    `## Appendix` backup marker) leaves the window open in both, and a
    `**Copy:**` line inside a slide body is not front matter, so it does not
    enable the mode.
    """
    for raw in (md_text or "").splitlines():
        line = raw.strip()
        if line.startswith(SLIDE_HEADING_PREFIX):
            break
        if not line.startswith("**"):
            continue
        key, _, value = line.lstrip("*").partition(":**")
        if key.strip() == COPY_KEY:
            return copy_is_locked(value.strip().strip('"'))
    return False


def is_verbatim_deck(ctx, meta):
    """True when this build is verbatim: the `--verbatim` flag (carried in
    ctx) or the outline's own `**Copy:** locked` front matter."""
    return bool((ctx or {}).get("verbatim")
                or copy_is_locked((meta or {}).get("copy")))


def verbatim_cap_errors(slides):
    """Errors for every author list a builder would silently truncate.

    Caps come from diff_deck.LAYOUT_CONTENT — the same table that tells the
    outline->deck diff which rows a builder was never going to render — so
    the two cannot drift apart — plus SUPPLEMENTAL_CAPS for the keys that
    table does not track. Imported inside the function: diff_deck imports
    build_deck at module level, and build_deck calls this.
    """
    from diff_deck import LAYOUT_CONTENT, resolve_cap
    errors = []
    for n, p in enumerate(slides, 1):
        layout = p.get("layout", "bullet-list")
        rules = {**LAYOUT_CONTENT.get(layout, {}),
                 **SUPPLEMENTAL_CAPS.get(layout, {})}
        for key, rule in rules.items():
            is_tuple = isinstance(rule, tuple)
            # A third element names the key this one is only a fallback for:
            # the builder ignores it entirely when that key carries entries.
            if is_tuple and len(rule) > 2 and p.get(rule[2]):
                continue
            cap = resolve_cap(rule[0] if is_tuple else rule)
            if not isinstance(cap, int):
                continue        # ALL — the builder renders every entry
            rows = p.get(key) or []
            if len(rows) > cap:
                errors.append(
                    f"Slide {n}: {len(rows)} {key} exceed {layout} cap {cap}"
                    " — verbatim mode refuses to drop content")
    return errors
