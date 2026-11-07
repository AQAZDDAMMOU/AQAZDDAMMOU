#!/usr/bin/env python3
"""
revision_grammar.py — the reviewer-instruction grammar behind revision_plan.

Three pure parsers, split out of `revision_plan.py` to keep both files inside
the module-size budget:

  parse_action(text)   one comment/feedback line -> a mechanical action dict
                       ({"op": "manual"} when nothing in the grammar fits, which
                       is the normal case for prose feedback — Claude handles it)
  locate_field(slide, anchor_text)
                       the field of a replayed outline slide a quoted fragment
                       refers to: exact -> casefold -> fuzzy (difflib)
  parse_feedback(md)   an emailed feedback.md -> [{"slide_ref", "text"}]

The grammar is deliberately small and literal. Anything ambiguous stays
`manual` rather than guessing: a wrong auto-edit costs more than a line of
work for the human.
"""
import difflib
import re

# Tail of a whole-line slide command: nothing but optional trailing
# punctuation may follow the verb (see the slide-level entries below).
_END = r"\s*[.!]?$"

# (op, regex, argument group names) — first match wins, so the more specific
# patterns come first. Matched case-insensitively against the comment text
# once its leading '@claude'/'please' politeness has been stripped.
ACTION_RX = [
    ("delete_bullet", r"^(?:delete|remove|drop|kill)\s+bullet\s+(\d+)\b",
     ("index",)),
    # the lookahead hands 'delete this slide' to kill_slide below — without
    # it the anchor rule would swallow every 'kill this slide'
    ("delete_anchor",
     r"^(?:delete|remove|drop|kill)\s+(?:this|that|it)\b(?!\s+slide\b)", ()),
    ("replace", r"^replace\s+[\"“](.+?)[\"”]\s+with\s+[\"“](.+?)[\"”]$",
     ("old", "new")),
    ("replace", r"^[\"“]?(.+?)[\"”]?\s*(?:->|→)\s*[\"“]?(.+?)[\"”]?$",
     ("old", "new")),
    ("retitle", r"^(?:retitle|title|heading)\s*:\s*(.+)$", ("text",)),
    ("add_bullet", r"^add\s+bullet\s*:\s*(.+)$", ("text",)),
    # Slide-level verbs are whole-line commands, anchored at both ends (bar
    # trailing punctuation): without the end anchor ordinary review prose
    # like 'remove slide numbers from the footer' or "delete slide 4's
    # chart" would silently delete a slide.
    ("move_slide", r"^move\s+to\s+slide\s+(\d+)" + _END, ("to",)),
    ("kill_slide",
     r"^(?:kill|delete|remove|drop)\s+(?:this\s+)?slide" + _END, ()),
    ("split_slide", r"^split\s+(?:this\s+)?slide" + _END, ()),
    ("merge_slide", r"^merge\s+with\s+slide\s+(\d+)" + _END, ("other",)),
]
_COMPILED = [(op, re.compile(rx, re.IGNORECASE), groups)
             for op, rx, groups in ACTION_RX]
_INT_ARGS = frozenset({"index", "to", "other"})
_POLITENESS_RX = re.compile(r"^(?:@claude|claude|please)\b[\s,:;-]*",
                            re.IGNORECASE)
MANUAL = {"op": "manual"}

# Fields of a replayed outline slide a reviewer can point at, in the order
# locate_field() scans them.
FIELD_KINDS = ("heading", "bullet", "card", "stat", "table_cell", "step",
               "item", "notes", "source", "kicker")
_DICT_KINDS = (("card", "cards"), ("stat", "stats"), ("step", "steps"),
               ("item", "items"))
_STRING_KINDS = (("notes", "notes"), ("source", "source"),
                 ("kicker", "kicker"))
MATCH_RATIO = 0.8
# Below this length a quoted fragment is too short to trust as a substring
# match ("the", "42%") — such anchors must match a field outright or fuzzily.
MIN_FRAGMENT_CHARS = 8

# Slide references an emailed feedback.md may use, e.g.
# 'Slide 4:', 'S 4', 'p. 4', '#4' and '"Recovering margin":'.
_NUM_REF_RX = [
    re.compile(r"^slide\s*(\d+)\s*[:.\-–—]?\s*", re.IGNORECASE),
    re.compile(r"^#\s*(\d+)\b\s*[:.\-–—]?\s*"),
    re.compile(r"^p\.?\s*(\d+)\b\s*[:.\-–—]?\s*", re.IGNORECASE),
    re.compile(r"^s\s*(\d+)\b\s*[:.\-–—]?\s*", re.IGNORECASE),
]
_TITLE_REF_RX = re.compile(r"^[\"“](.+?)[\"”]\s*:\s*")
_LIST_MARKER_RX = re.compile(r"^(?:[-*+]|\d+[.)])\s+")


# ── the command grammar ─────────────────────────────────────────────────────
def parse_action(text) -> dict:
    """One reviewer line -> {"op": ..., **args}, or {"op": "manual"}.

    Only single-line instructions are mechanical; anything else (prose, a
    multi-line note, an unrecognised verb) is left for a human/Claude.
    """
    cleaned = _POLITENESS_RX.sub("", str(text or "").strip()).strip()
    if not cleaned or "\n" in cleaned:
        return dict(MANUAL)
    # A trailing full stop is punctuation, not part of the instruction — but
    # only retry without it, so 'Retitle: Growth stalls.' keeps its own.
    for candidate in (cleaned, cleaned.rstrip(".").strip()):
        action = _match_action(candidate)
        if action is not None:
            return action
    return dict(MANUAL)


def _match_action(text):
    for op, rx, groups in _COMPILED:
        m = rx.match(text)
        if not m:
            continue
        action = {"op": op}
        for name, value in zip(groups, m.groups()):
            action[name] = int(value) if name in _INT_ARGS else value.strip()
        return action
    return None


# ── locating the field a comment points at ──────────────────────────────────
def _kv_text(entry):
    if isinstance(entry, dict):
        return " ".join(str(v) for v in entry.values() if v)
    return str(entry)


def _candidates(slide):
    """(kind, index, text) for every addressable field, in FIELD_KINDS order.

    `index` is 0-based within its own list (0 for the single-value fields);
    table cells are numbered row-major across the whole table.
    """
    heading = slide.get("heading") or slide.get("title") or ""
    if heading:
        yield ("heading", 0, str(heading))
    for i, bullet in enumerate(slide.get("bullets") or []):
        yield ("bullet", i, str(bullet))
    for kind, key in _DICT_KINDS:
        for i, entry in enumerate(slide.get(key) or []):
            yield (kind, i, _kv_text(entry))
    index = 0
    for row in slide.get("table_rows") or []:
        for cell in row:
            yield ("table_cell", index, str(cell))
            index += 1
    for kind, key in _STRING_KINDS:
        value = slide.get(key)
        if value:
            yield (kind, 0, str(value))


def _score(anchor, text):
    """Similarity of a casefolded anchor to a casefolded field.

    difflib's ratio for typos, plus a containment leg (longest common block
    over the anchor's length) so a reviewer quoting *part* of a bullet still
    lands on the whole bullet.
    """
    matcher = difflib.SequenceMatcher(None, anchor, text, autojunk=False)
    ratio = matcher.ratio()
    if len(anchor) >= MIN_FRAGMENT_CHARS:
        block = matcher.find_longest_match(0, len(anchor), 0, len(text))
        ratio = max(ratio, block.size / len(anchor))
    return ratio


def locate_field(slide: dict, anchor_text) -> dict | None:
    """{"kind", "index", "text"} for the field `anchor_text` refers to.

    Exact match -> casefolded match -> best difflib score >= MATCH_RATIO.
    None when nothing is close enough *or* when two fields tie for the best
    fuzzy score: an ambiguous anchor is a job for a human, not a guess.
    """
    anchor = str(anchor_text or "").strip()
    if not anchor:
        return None
    candidates = list(_candidates(slide))
    for kind, index, text in candidates:
        if text == anchor:
            return {"kind": kind, "index": index, "text": text}
    folded = anchor.casefold()
    for kind, index, text in candidates:
        if text.casefold() == folded:
            return {"kind": kind, "index": index, "text": text}
    scored = [(_score(folded, text.casefold()), kind, index, text)
              for kind, index, text in candidates]
    if not scored:
        return None
    best = max(score for score, _k, _i, _t in scored)
    if best < MATCH_RATIO:
        return None
    tied = [c for c in scored if c[0] == best]
    if len(tied) > 1:
        return None
    _score_, kind, index, text = tied[0]
    return {"kind": kind, "index": index, "text": text}


# ── emailed feedback ────────────────────────────────────────────────────────
def _slide_ref(line):
    """(slide_ref, remaining text) for one feedback line."""
    for rx in _NUM_REF_RX:
        m = rx.match(line)
        if m:
            return int(m.group(1)), line[m.end():].strip()
    m = _TITLE_REF_RX.match(line)
    if m:
        return m.group(1).strip(), line[m.end():].strip()
    return None, line


def parse_feedback(md_text) -> list[dict]:
    """An emailed feedback.md -> [{"slide_ref": int|str|None, "text": str}].

    A reference may lead the line ('Slide 4: drop bullet 2') or stand on its
    own line, in which case it is the context for the lines under it until the
    next reference or blank line. Markdown headings that carry no reference
    are document furniture and are dropped.
    """
    items, context = [], None
    for raw in str(md_text or "").splitlines():
        line = raw.strip()
        if not line:
            context = None
            continue
        ref, text = _slide_ref(_LIST_MARKER_RX.sub("", line))
        if ref is None and line.startswith("#"):
            # a markdown heading: a reference under the '#'s is context,
            # anything else is a title and not feedback
            ref, text = _slide_ref(line.lstrip("#").strip())
            if ref is None:
                continue
        if ref is not None and not text:
            context = ref
            continue
        if not text:
            continue
        items.append({"slide_ref": ref if ref is not None else context,
                      "text": text})
    return items
