"""House-style lints: word lists, action-title hygiene, typography slips.

A deck may declare `**Style-File:** style.json` in front-matter (or the build
may pass `--style path`); otherwise `<assets>/style.json` is used when present
and the built-in `DEFAULT_STYLE` otherwise.  Every rule is a *warning* — house
style is editorial, not structural — and each finding reads:

    Slide 7: house style — marketing adjectives: leverage, world-class

Findings are aggregated per slide per rule so a wordy slide produces one line
per category rather than one line per hit.  Rules covered:

1. `not X but Y` constructions (a consulting-prose tell).
2. marketing adjectives / banned phrases / hedges / judgemental terms.
3. passive voice in action titles (content layouts only).
4. duplicate and near-duplicate slide titles.
5. title punctuation policy (`title_punctuation`).
6. sentence-vs-title case drift (`title_case`).
7. orphan footnote markers (body marker with no source entry, and vice versa).
8. typography slips (double spaces, `..`, space before punctuation).

Heading length (`max_title_words`) is owned here too — it moved out of
`build_deck.validate()`; only the "and"-compound heading check stays there.
"""
import difflib
import json
import re
from pathlib import Path

DEFAULT_STYLE = {
    "banned_phrases": [],
    "marketing_adjectives": [
        "cutting-edge", "world-class", "seamless", "robust", "leverage",
        "synergy", "synergies", "game-changing", "best-in-class",
        "innovative", "holistic", "next-generation", "unlock", "empower",
        "revolutionary", "transformative", "delve", "state-of-the-art",
        "unparalleled", "groundbreaking", "paradigm", "turnkey",
        "frictionless", "elevate", "supercharge", "turbocharge",
    ],
    "hedges": [
        "may potentially", "could potentially", "arguably", "somewhat",
        "fairly", "it could be argued", "to some extent", "possibly",
    ],
    "judgemental": [
        "failed to", "neglected", "incompetent", "poor management",
        "mismanaged", "careless", "lazy", "ignored the obvious",
    ],
    "title_case": "sentence",
    "title_punctuation": "none",
    "max_title_words": 15,
    "near_duplicate_ratio": 0.85,
}

LIST_KEYS = ("banned_phrases", "marketing_adjectives", "hedges", "judgemental")
TITLE_CASE_MODES = ("sentence", "title", "any")
TITLE_PUNCTUATION_MODES = ("none", "period", "consistent")
STYLE_FILENAME = "style.json"

# Case drift: a heading is "Title Case" when this share of eligible words is
# capitalised, and only when there are at least this many eligible words.
CASE_DRIFT_RATIO = 0.60
CASE_DRIFT_MIN_WORDS = 3
CASE_DRIFT_MIN_LEN = 3          # words this short are ignored (of, the, GPU)
# Headings shorter than this are already flagged as topic labels elsewhere.
MIN_ACTION_TITLE_WORDS = 5
PUNCT_MINORITY_MIN_SLIDES = 3   # 'consistent' policy needs a real majority

NOT_BUT_RX = re.compile(r"\bnot\s+(?:only\s+|just\s+)?[^,.;]{1,40}\bbut\b",
                        re.I)
IRREGULAR_PARTICIPLES = ("made", "done", "seen", "taken", "given", "built",
                         "held")
# Words that end in -ed/-en but are adverbs, adjectives, or numerals after a
# copula ("margins are often thinner", "the gap is even wider") — never
# participles, so they must not read as passive voice.
PASSIVE_STOPWORDS = frozenset({
    "often", "even", "then", "when", "seven", "eleven", "ten", "green",
    "open", "oven", "red", "sudden", "golden", "wooden",
})
PASSIVE_RX = re.compile(
    r"\b(?:is|are|was|were|be|been|being)\s+"
    r"(\w+ed|\w+en|" + "|".join(IRREGULAR_PARTICIPLES) + r")\b", re.I)

SUPERSCRIPT_DIGITS = "⁰¹²³⁴⁵⁶⁷⁸⁹"
UNIT_SUPERSCRIPTS = "²³"        # m², km³ — units, not footnote markers
_SUPER_MAP = {ch: str(i) for i, ch in enumerate(SUPERSCRIPT_DIGITS)}
BRACKET_MARKER_RX = re.compile(r"\[(\d{1,2})\]")
SUPER_MARKER_RX = re.compile(f"[{SUPERSCRIPT_DIGITS}]" + "{1,2}")
FOOTNOTE_LINE_RX = re.compile(
    r"^\[?\(?([0-9" + SUPERSCRIPT_DIGITS + r"]{1,2})\)?\]?[\s.:)]")

DOUBLE_SPACE_RX = re.compile(r"\S {2,}\S")
DOUBLE_DOT_RX = re.compile(r"(?<!\.)\.\.(?!\.)")
# Two lowercase letters before the period keep abbreviations ("U.S.", "e.g.")
# and initials ("A.Sharma") out of the missing-space check.
NO_SPACE_AFTER_PERIOD_RX = re.compile(r"(?<=[a-z]{2})\.[A-Z]")
SPACE_BEFORE_PUNCT_RX = re.compile(r"\s[,.](?:\s|$|\w)")

# Slide fields that are not visible prose (speaker notes and footnotes are
# checked separately; the rest are directives, data, or file paths).
NON_PROSE_FIELDS = {
    "notes", "source", "footnote", "footnotes", "note", "layout", "palette",
    "visual", "visual_left", "visual_right", "data_file", "evidence", "data",
    "labels_mode", "sort", "scale", "cagr", "axis_max", "axis_break",
    "value_line", "bracket", "marker", "columns", "hidden",
}
FOOTNOTE_FIELDS = ("source", "notes", "footnote")

WORD_PUNCT = "—–.,;:!?()[]\"'“”‘’"
BREAK_TOKENS = ("—", "–", "-", "|")


# --------------------------------------------------------------------------
# style loading
# --------------------------------------------------------------------------

def load_style(path):
    """Load a house-style JSON file merged over DEFAULT_STYLE.

    `path=None` returns a deep-enough copy of the defaults.  Any problem —
    missing file, malformed JSON, wrong value type, unknown key — raises
    ValueError so callers can fail fast.
    """
    style = _default_copy()
    if path is None:
        return style
    path = Path(path)
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except OSError as e:
        raise ValueError(f"unreadable ({e})") from e
    except json.JSONDecodeError as e:
        raise ValueError(f"invalid JSON ({e})") from e
    if not isinstance(raw, dict):
        raise ValueError("must be a JSON object mapping style keys to values")

    unknown = sorted(set(raw) - set(DEFAULT_STYLE))
    if unknown:
        raise ValueError(f"unknown key(s) {unknown} (supported: "
                         f"{', '.join(sorted(DEFAULT_STYLE))})")
    for key in LIST_KEYS:
        if key in raw:
            value = raw[key]
            if not isinstance(value, list) or not all(
                    isinstance(v, str) for v in value):
                raise ValueError(f"'{key}' must be a list of strings")
            style[key] = [v.strip() for v in value if v.strip()]
    if "title_case" in raw:
        style["title_case"] = _enum(raw, "title_case", TITLE_CASE_MODES)
    if "title_punctuation" in raw:
        style["title_punctuation"] = _enum(raw, "title_punctuation",
                                           TITLE_PUNCTUATION_MODES)
    if "max_title_words" in raw:
        value = raw["max_title_words"]
        if not isinstance(value, int) or isinstance(value, bool) or value < 1:
            raise ValueError("'max_title_words' must be a positive integer")
        style["max_title_words"] = value
    if "near_duplicate_ratio" in raw:
        value = raw["near_duplicate_ratio"]
        if isinstance(value, bool) or not isinstance(value, (int, float)) \
                or not 0 < value <= 1:
            raise ValueError("'near_duplicate_ratio' must be a number in (0, 1]")
        style["near_duplicate_ratio"] = float(value)
    return style


def _default_copy():
    return {k: (list(v) if isinstance(v, list) else v)
            for k, v in DEFAULT_STYLE.items()}


def merge_lists(style, extra):
    """Return a copy of `style` with its LIST_KEYS extended by `extra`.

    Used to fold a client profile's own word lists (banned phrases and the
    like) into the resolved house style.  Entries are appended in order and
    deduped case-insensitively against what the style already carries; every
    other style key is copied through untouched.
    """
    merged = {k: (list(v) if isinstance(v, list) else v)
              for k, v in (style or {}).items()}
    for key in LIST_KEYS:
        items = (extra or {}).get(key) or []
        current = merged.setdefault(key, [])
        seen = {str(v).lower() for v in current}
        for item in items:
            text = str(item).strip()
            if text and text.lower() not in seen:
                current.append(text)
                seen.add(text.lower())
    return merged


def _enum(raw, key, allowed):
    value = raw[key]
    if not isinstance(value, str) or value not in allowed:
        raise ValueError(f"'{key}' must be one of {', '.join(allowed)}")
    return value


def resolve_style(meta, ctx):
    """Resolve the style file for a build. Returns (path or None, errors).

    Order: ctx['style_path'] (CLI --style, resolved as given — i.e. against
    the working directory, like every other CLI path) > `**Style-File:**`
    (outline dir, then assets dir — it is authored inside the outline) >
    `<assets>/style.json` > built-in defaults (None).  An explicitly
    requested file that does not exist is an error.
    """
    cli_spec = ctx.get("style_path")
    if cli_spec:
        candidate = Path(str(cli_spec))
        if candidate.is_file():
            return candidate, []
        return None, [f"style file {str(cli_spec)!r} not found "
                      "(--style is resolved against the working directory)"]

    spec = (meta or {}).get("style_file")
    if spec:
        for candidate in (Path(ctx.get("outline_dir", ".")) / str(spec),
                          Path(ctx.get("assets_dir", "assets")) / str(spec)):
            if candidate.is_file():
                return candidate, []
        return None, [f"style file {spec!r} not found (looked in the outline "
                      "dir and the assets dir)"]
    default_path = Path(ctx.get("assets_dir", "assets")) / STYLE_FILENAME
    if default_path.is_file():
        return default_path, []
    return None, []


# --------------------------------------------------------------------------
# text extraction
# --------------------------------------------------------------------------

def _flatten(value):
    """Yield every string reachable in a slide field value."""
    if isinstance(value, str):
        if value.strip():
            yield value
    elif isinstance(value, dict):
        for v in value.values():
            yield from _flatten(v)
    elif isinstance(value, (list, tuple)):
        for v in value:
            yield from _flatten(v)


def prose_texts(slide):
    """Visible prose on a slide (excludes speaker notes and footnotes)."""
    out = []
    for key, value in slide.items():
        if key.startswith("_") or key in NON_PROSE_FIELDS:
            continue
        out.extend(_flatten(value))
    return out


def _narrative_texts(slide):
    """Heading, bullets, and kicker — where the argument is made."""
    texts = [slide.get("heading", ""), slide.get("kicker", "")]
    texts.extend(_flatten(slide.get("bullets", [])))
    return [t for t in texts if t]


def _footnote_texts(slide):
    out = []
    for key in FOOTNOTE_FIELDS:
        out.extend(_flatten(slide.get(key, "")))
    return out


def _action_title_layouts():
    """ACTION_TITLE_LAYOUTS from outline_validate (imported lazily:
    outline_validate imports narrative, which re-exports this module)."""
    from outline_validate import ACTION_TITLE_LAYOUTS
    return ACTION_TITLE_LAYOUTS


def _slide_title(slide):
    return (slide.get("heading") or slide.get("title") or "").strip()


def _is_content_slide(slide, action_layouts):
    return (not slide.get("_appendix")
            and slide.get("layout", "bullet-list") in action_layouts)


# --------------------------------------------------------------------------
# individual rules
# --------------------------------------------------------------------------

def _term_hits(texts, terms):
    """Terms from `terms` present in `texts` (word-boundary, case-folded)."""
    terms = [t for t in (terms or []) if t and t.strip()]
    if not terms:
        return []
    pattern = r"(?<!\w)(?:" + "|".join(
        re.escape(t.strip()) for t in terms) + r")(?!\w)"
    seen, hits = set(), []
    for text in texts:
        for m in re.finditer(pattern, text, re.I):
            key = m.group(0).casefold()
            if key not in seen:
                seen.add(key)
                hits.append(m.group(0))
    return hits


def _passive_hits(heading):
    return [" ".join(m.group(0).split())
            for m in PASSIVE_RX.finditer(heading)
            if m.group(1).casefold() not in PASSIVE_STOPWORDS]


def _case_drift(heading, mode):
    """Return a drift description for `heading`, or None."""
    if mode == "any":
        return None
    words, skip_next = [], False
    for i, token in enumerate(heading.split()):
        after_break = skip_next
        skip_next = token.endswith(":") or token in BREAK_TOKENS
        if i == 0 or after_break:
            continue
        word = token.strip(WORD_PUNCT)
        if (len(word) <= CASE_DRIFT_MIN_LEN or word.isupper()
                or any(c.isdigit() for c in word) or not word[0].isalpha()):
            continue
        words.append(word)
    if len(words) < CASE_DRIFT_MIN_WORDS:
        return None
    capitalised = sum(1 for w in words if w[0].isupper())
    ratio = capitalised / len(words)
    if mode == "sentence" and ratio >= CASE_DRIFT_RATIO:
        return ("Title Case heading — house style is sentence case "
                f"({capitalised}/{len(words)} words capitalised)")
    if mode == "title" and ratio <= 1 - CASE_DRIFT_RATIO:
        return ("sentence-case heading — house style is Title Case "
                f"({capitalised}/{len(words)} words capitalised)")
    return None


def _markers(texts):
    """Footnote marker numbers found in body prose."""
    found = set()
    for text in texts:
        found.update(BRACKET_MARKER_RX.findall(text))
        for m in SUPER_MARKER_RX.finditer(text):
            if _is_superscript_unit(text, m):
                continue
            found.add("".join(_SUPER_MAP[c] for c in m.group(0)).lstrip("0")
                      or "0")
    return found


def _is_superscript_unit(text, match):
    """True for `m²` / `km³`: a squared/cubed marker glued to a unit name."""
    if any(c not in UNIT_SUPERSCRIPTS for c in match.group(0)):
        return False
    start = match.start()
    return start > 0 and text[start - 1].isalpha()


def _declared_footnote_numbers(slide):
    """Numbers the `- Footnote:` list (`p["footnotes"]`) declares — resolved
    through footnotes.number_footnotes so auto (positional) and pinned
    entries both count."""
    entries = slide.get("footnotes") or []
    if not entries:
        return set()
    from footnotes import number_footnotes
    return {str(num) for num, _ in number_footnotes(entries)}


def _footnote_numbers(texts):
    """Footnote numbers declared in source/notes/footnote fields."""
    found = set()
    for text in texts:
        for line in re.split(r"[\n;]", text):
            m = FOOTNOTE_LINE_RX.match(line.strip())
            if m:
                digits = "".join(_SUPER_MAP.get(c, c) for c in m.group(1))
                found.add(digits.lstrip("0") or "0")
    return found


def _typography_hits(texts):
    hits = []
    checks = (
        ("double space", DOUBLE_SPACE_RX),
        ("stray '..'", DOUBLE_DOT_RX),
        ("missing space after period", NO_SPACE_AFTER_PERIOD_RX),
        ("space before punctuation", SPACE_BEFORE_PUNCT_RX),
    )
    for label, rx in checks:
        if any(rx.search(t) for t in texts):
            hits.append(label)
    return hits


# --------------------------------------------------------------------------
# entry point
# --------------------------------------------------------------------------

def check_house_style(slides, meta, style, action_layouts=None):
    """Return house-style warnings for a parsed outline.

    `meta` is accepted for symmetry with the other narrative checks (and so
    deck-level policy can be read here later); rules are per slide today.
    """
    style = style or _default_copy()
    layouts = action_layouts if action_layouts is not None \
        else _action_title_layouts()
    findings = {}   # slide number -> list of (rule, detail)

    def add(n, rule, detail):
        findings.setdefault(n, []).append((rule, detail))

    titles = []      # (slide number, title) for duplicate detection
    trailing = []    # (slide number, has trailing period)

    for n, slide in enumerate(slides, 1):
        if slide.get("_auto") or slide.get("_appendix"):
            continue
        prose = prose_texts(slide)
        narrative_text = _narrative_texts(slide)
        heading = _slide_title(slide)
        is_content = _is_content_slide(slide, layouts)

        if any(NOT_BUT_RX.search(t) for t in narrative_text):
            add(n, "not-but construction",
                "rewrite as a positive claim ('X, not Y' or two sentences)")

        for rule, key in (("banned phrases", "banned_phrases"),
                          ("marketing adjectives", "marketing_adjectives"),
                          ("hedges", "hedges"),
                          ("judgemental language", "judgemental")):
            hits = _term_hits(prose, style.get(key, []))
            if hits:
                add(n, rule, ", ".join(hits))

        if is_content and heading and \
                len(heading.split()) >= MIN_ACTION_TITLE_WORDS:
            passive = _passive_hits(heading)
            if passive:
                add(n, "passive voice",
                    f"{', '.join(passive)} — name the actor in the title")
            drift = _case_drift(
                heading, style.get("title_case", DEFAULT_STYLE["title_case"]))
            if drift:
                add(n, "case drift", drift)

        if is_content and heading:
            max_words = style.get("max_title_words",
                                  DEFAULT_STYLE["max_title_words"])
            if len(heading.split()) > max_words:
                add(n, "title length", f"heading exceeds {max_words} words — "
                    "tighten the takeaway")
            titles.append((n, heading))
            trailing.append((n, heading.endswith(".")
                             and not heading.endswith("..")))

        markers = _markers(prose)
        footnotes = _footnote_numbers(_footnote_texts(slide))
        footnotes |= _declared_footnote_numbers(slide)
        orphans = []
        for num in sorted(markers - footnotes, key=_as_int):
            orphans.append(f"marker [{num}] has no source/notes entry")
        for num in sorted(footnotes - markers, key=_as_int):
            orphans.append(f"footnote {num} has no marker in the body")
        if orphans:
            add(n, "orphan footnote marker", "; ".join(orphans))

        typo = _typography_hits(prose)
        if typo:
            add(n, "typography", ", ".join(typo))

    for n, detail in _title_punctuation_findings(
            trailing, style.get("title_punctuation",
                                DEFAULT_STYLE["title_punctuation"])):
        add(n, "title punctuation", detail)
    for n, detail in _duplicate_findings(
            titles, style.get("near_duplicate_ratio",
                              DEFAULT_STYLE["near_duplicate_ratio"])):
        add(n, "near-duplicate title", detail)

    return [f"Slide {n}: house style — {rule}: {detail}"
            for n in sorted(findings)
            for rule, detail in findings[n]]


def _as_int(value):
    try:
        return int(value)
    except ValueError:
        return 0


def _title_punctuation_findings(trailing, mode):
    out = []
    if not trailing:
        return out
    if mode == "none":
        return [(n, "heading ends with a period — house style is no terminal "
                 "punctuation") for n, has_period in trailing if has_period]
    if mode == "period":
        return [(n, "heading has no terminal period — house style ends "
                 "headings with one") for n, has_period in trailing
                if not has_period]
    # 'consistent': report whichever treatment is in the minority
    with_period = [n for n, has in trailing if has]
    without = [n for n, has in trailing if not has]
    if not with_period or not without or len(trailing) < PUNCT_MINORITY_MIN_SLIDES:
        return out
    if len(with_period) < len(without):
        minority = with_period
        note, majority = "ends with a period", "most headings do not"
    else:
        minority = without
        note, majority = "has no terminal period", "most headings do"
    return [(n, f"heading {note} while {majority} — pick one treatment")
            for n in minority]


def _duplicate_findings(titles, ratio):
    out = []
    for i, (n, title) in enumerate(titles):
        for m, other in titles[:i]:
            score = difflib.SequenceMatcher(None, other.casefold(),
                                            title.casefold()).ratio()
            if score >= ratio:
                out.append((n, f"reads like slide {m} ({other!r}) — "
                            f"{score:.0%} similar; sharpen one of them"))
                break
    return out
