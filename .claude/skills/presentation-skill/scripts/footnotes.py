#!/usr/bin/env python3
"""
footnotes.py — footnote numbering and the source-line normaliser.

`- Footnote:` entries (repeatable → `p["footnotes"]`) are numbered 1..n in
order by `number_footnotes()`; an explicit leading `N ` / `N.` / `[N]` pins
the number and the auto counter continues from it.  `superscript()` spells a
number in Unicode superscript digits for the rendered block and the body
markers (`helpers.parse_rich_segments` turns `[n]` into superscript runs).

`normalize_source()` is the single formatter behind
`slide_chrome.add_source`: it strips any `Source:`/`Sources:` prefix variant,
dedupes `;` / ` | ` separated entries case-insensitively, re-joins with
`'; '`, and prefixes `- Note:` text MBB-style (`Note: … · Source: …`).
Under verbatim (`**Copy:** locked`) the source text is left byte-exact —
only the note prefix still applies.
"""
import re

SUPERSCRIPT_DIGITS = "⁰¹²³⁴⁵⁶⁷⁸⁹"

# Explicit number pin at the head of a `- Footnote:` entry: `[2] text`,
# `2. text` or `2 text` (1–2 digits, so a year never reads as a pin).
# Explicit pins need a delimiter — `[N]` or `N.` — so a footnote whose text
# merely starts with a number ('3 of 5 sites migrated') is never renumbered
# or truncated (review finding; the bare `N ` form was dropped for this).
PIN_RX = re.compile(
    r"^(?:\[(\d{1,2})\]\s*|(\d{1,2})\.\s+)(\S.*)$")

# Leading `Source:` / `Sources:` / `source -` spelling variants.
SOURCE_PREFIX_RX = re.compile(r"^\s*sources?\s*[:\-–—]\s*", re.I)


def number_footnotes(entries):
    """[(number, text)] for a slide's footnote list.

    Auto numbering is positional (1..n); a pinned entry takes its own number
    and the counter continues from it (`a, [5] b, c` → 1, 5, 6).
    """
    out, counter = [], 0
    for raw in entries or []:
        text = str(raw).strip()
        m = PIN_RX.match(text)
        if m:
            num = int(next(g for g in m.groups()[:2] if g is not None))
            text = m.group(3).strip()
        else:
            num = counter + 1
        counter = num
        out.append((num, text))
    return out


def superscript(num):
    """`12` → `¹²` (Unicode superscript digits)."""
    return "".join(SUPERSCRIPT_DIGITS[int(d)] for d in str(int(num)))


def _dedupe_entries(body):
    """Split on `;` and ` | `, trim, drop exact case-insensitive duplicates."""
    parts = [e.strip() for chunk in body.split(";")
             for e in chunk.split(" | ")]
    seen, entries = set(), []
    for part in parts:
        if not part:
            continue
        key = part.casefold()
        if key not in seen:
            seen.add(key)
            entries.append(part)
    return entries


def normalize_source(source, note=None, verbatim=False):
    """The rendered source-line text, or '' when there is nothing to print.

    - source only  → `Source: a; b` (prefix variants normalised, entries
      deduped case-insensitively, rejoined with '; ')
    - note + source → `Note: … · Source: …` (note first — MBB convention)
    - note only    → `Note: …`
    - verbatim     → the source text is not normalised (bytes stay exact,
      today's bare `Source: ` prefixing included); the note still prefixes.
    """
    source = (source or "").strip()
    note = (note or "").strip()
    if source:
        if verbatim:
            line = source if source.lower().startswith("source") \
                else f"Source: {source}"
        else:
            body = SOURCE_PREFIX_RX.sub("", source)
            line = "Source: " + "; ".join(_dedupe_entries(body))
        return f"Note: {note} · {line}" if note else line
    return f"Note: {note}" if note else ""
