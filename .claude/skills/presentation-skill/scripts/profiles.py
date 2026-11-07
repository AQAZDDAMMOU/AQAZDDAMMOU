"""profiles.py — engagement/client profiles.

A profile is a JSON file at `<assets>/profiles/<name>.json` that carries the
defaults one client or engagement always wants, so an outline does not have to
repeat them:

    {
      "palette": "acme", "footer": "Acme — Confidential", "stamp": "DRAFT",
      "author": "Acme Advisory", "number_format": "units=M decimals=1",
      "banned_phrases": ["synergy capture"],
      "mandatory_slides": [{"layout": "exec-summary"},
                           {"title_contains": "Next steps"}],
      "preferred_layouts": ["exec-summary"], "notes_style": "bullets"
    }

A deck selects one with `**Profile:** acme` in front-matter or `--profile acme`
on the command line.  Precedence is CLI flag > outline front-matter > profile >
variant preset > default, so `apply()` only fills deck-meta keys the outline
left empty (the CLI flags are applied later in `build()` and already win).

Beyond defaults a profile carries *checks*: mandatory slides (missing one is a
fail-fast error), banned phrases (merged into the house style, so a hit is an
editorial warning), a speaker-notes style (warning), and preferred layouts that
break auto-layout ties.  `record_lessons()` keeps a per-profile lessons file so
findings that recur across deliveries surface as a checklist for the next deck.

The resolved profile lives in the build's `ctx['_profile']` — no module-level
state.
"""
import json
import re
from datetime import date
from pathlib import Path

PROFILES_DIRNAME = "profiles"
NAME_RX = re.compile(r"[A-Za-z0-9][A-Za-z0-9_-]*")

# Every key a profile file may carry.
PROFILE_KEYS = {"palette", "footer", "stamp", "density", "size", "author",
                "number_format", "style_file", "logo", "banned_phrases",
                "mandatory_slides", "preferred_layouts", "notes_style",
                "page_numbers", "tracker"}
# Keys applied as deck-meta defaults below the outline's own front-matter.
META_FROM_PROFILE = {"palette", "footer", "stamp", "density", "size", "author",
                     "number_format", "style_file", "page_numbers", "tracker"}
STRING_KEYS = META_FROM_PROFILE | {"logo", "notes_style"}
STRING_LIST_KEYS = {"banned_phrases", "preferred_layouts"}
MANDATORY_KEYS = ("layout", "title_contains")
NOTES_STYLES = ("bullets", "prose", "none")

# Lessons: a finding's signature ignores which slide it landed on and any
# figures inside it, so "Slide 3: 12 shapes overlap" and "Slide 9: 481 shapes
# overlap" are recognised as the same recurring problem.
SLIDE_PREFIX_RX = re.compile(
    r"^Slide\s+\d+\s*(?:\(line\s+\d+\))?\s*[:\-—–]?\s*")
DIGITS_RX = re.compile(r"\d+")
DIGIT_TOKEN = "#"
MIN_LESSON_COUNT = 2        # a finding is a "lesson" once it has recurred


# --------------------------------------------------------------------------
# loading
# --------------------------------------------------------------------------

def profile_path(name, assets_dir):
    """<assets>/profiles/<name>.json for a validated profile name."""
    return (Path(assets_dir or "assets") / PROFILES_DIRNAME
            / f"{safe_name(name)}.json")


def safe_name(name):
    """Reject names that could escape an assets subdirectory (path
    traversal, absolute paths, etc). Shared by profiles.py (profile/lessons
    file names) and brand_kit.py (--name)."""
    text = str(name or "").strip()
    if not NAME_RX.fullmatch(text):
        raise ValueError(f"name {text!r} must be alphanumeric "
                         "(plus interior - and _)")
    return text


def load_profile(path):
    """Load and validate a profile JSON file.

    Raises ValueError on any problem — a requested profile is never skipped.
    """
    path = Path(path)
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except OSError as e:
        raise ValueError(f"unreadable ({e})") from e
    except json.JSONDecodeError as e:
        raise ValueError(f"invalid JSON ({e})") from e
    return validate_profile(raw)


def validate_profile(raw):
    """Validate a profile dict; returns a normalised copy."""
    if not isinstance(raw, dict):
        raise ValueError("must be a JSON object mapping profile keys "
                         "to values")
    unknown = sorted(set(raw) - PROFILE_KEYS)
    if unknown:
        raise ValueError(f"unknown key(s) {unknown} (supported: "
                         f"{', '.join(sorted(PROFILE_KEYS))})")
    profile = {}
    for key in sorted(STRING_KEYS & set(raw)):
        value = raw[key]
        if not isinstance(value, str):
            raise ValueError(f"'{key}' must be a string")
        profile[key] = value.strip()
    for key in sorted(STRING_LIST_KEYS & set(raw)):
        value = raw[key]
        if not isinstance(value, list) or not all(
                isinstance(v, str) for v in value):
            raise ValueError(f"'{key}' must be a list of strings")
        profile[key] = [v.strip() for v in value if v.strip()]
    if "notes_style" in profile and profile["notes_style"] not in NOTES_STYLES:
        raise ValueError("'notes_style' must be one of "
                         f"{', '.join(NOTES_STYLES)}")
    if "mandatory_slides" in raw:
        profile["mandatory_slides"] = _validate_mandatory(
            raw["mandatory_slides"])
    return profile


def _validate_mandatory(value):
    if not isinstance(value, list):
        raise ValueError("'mandatory_slides' must be a list of objects like "
                         '{"layout": "exec-summary"}')
    entries = []
    for item in value:
        if not isinstance(item, dict) or not item:
            raise ValueError("each 'mandatory_slides' entry must be a "
                             "non-empty object with "
                             f"{' and/or '.join(MANDATORY_KEYS)}")
        unknown = sorted(set(item) - set(MANDATORY_KEYS))
        if unknown:
            raise ValueError(f"'mandatory_slides' entry has unknown key(s) "
                             f"{unknown} (supported: "
                             f"{', '.join(MANDATORY_KEYS)})")
        if not all(isinstance(v, str) and v.strip() for v in item.values()):
            raise ValueError("'mandatory_slides' values must be non-empty "
                             "strings")
        entries.append({k: item[k].strip() for k in MANDATORY_KEYS
                        if k in item})
    return entries


def write_profile(name, assets_dir, data):
    """Validate `data` and write it to <assets>/profiles/<name>.json."""
    profile = validate_profile(data)
    out = profile_path(name, assets_dir)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(profile, indent=2, ensure_ascii=False) + "\n",
                   encoding="utf-8")
    return out


# --------------------------------------------------------------------------
# resolution into a build
# --------------------------------------------------------------------------

def resolve(meta, ctx):
    """Resolve the build's profile. Returns (profile or None, errors).

    Order: `ctx['profile']` (the --profile flag) > `**Profile:**` front-matter.
    A requested profile that is missing or malformed is an error, never a
    silent fallback.
    """
    ctx = ctx or {}
    name = str(ctx.get("profile") or (meta or {}).get("profile") or "").strip()
    if not name:
        return None, []
    try:
        path = profile_path(name, ctx.get("assets_dir", "assets"))
    except ValueError as e:
        return None, [str(e)]
    if not path.is_file():
        return None, [f"profile {name!r} not found (looked for {path})"]
    try:
        profile = load_profile(path)
    except ValueError as e:
        return None, [f"profile {name!r} ({path}): {e}"]
    profile["_name"] = name
    return profile, []


def apply(meta, ctx):
    """Fill deck-meta defaults from the profile. Returns (new meta, errors).

    Stores the resolved profile in `ctx['_profile']` ({} when there is none) so
    `validate()` and the layout pass can read it without module-level state.
    Only keys the outline left absent or empty are filled — front-matter wins.
    """
    meta = dict(meta or {})
    profile, errors = resolve(meta, ctx)
    if ctx is not None:
        ctx["_profile"] = profile or {}
    if profile:
        for key in sorted(META_FROM_PROFILE):
            if profile.get(key) and not str(meta.get(key, "")).strip():
                meta[key] = profile[key]
    return meta, errors


# --------------------------------------------------------------------------
# checks
# --------------------------------------------------------------------------

def check_mandatory(slides, profile):
    """Errors for mandatory slides the outline never delivers."""
    profile = profile or {}
    name = profile.get("_name", "")
    errors = []
    for want in profile.get("mandatory_slides", []):
        if any(_matches(slide, want) for slide in slides):
            continue
        what = " + ".join(
            f"layout {value}" if key == "layout"
            else f"title containing {value!r}"
            for key, value in want.items())
        errors.append(f"profile {name!r}: mandatory slide missing: {what}")
    return errors


def _matches(slide, want):
    if "layout" in want and slide.get("layout") != want["layout"]:
        return False
    if "title_contains" in want:
        title = f"{slide.get('title', '')} {slide.get('heading', '')}".lower()
        if want["title_contains"].lower() not in title:
            return False
    return True


def check_notes_style(slides, profile):
    """Warnings for speaker notes that break the profile's notes style."""
    profile = profile or {}
    style = profile.get("notes_style", "")
    if style not in NOTES_STYLES:
        return []
    name = profile.get("_name", "")
    warnings = []
    for n, slide in enumerate(slides, 1):
        notes = (slide.get("notes") or "").strip()
        lines = [line.strip() for line in notes.splitlines() if line.strip()]
        detail = ""
        if style == "none" and notes:
            detail = "expects no speaker notes"
        elif style == "bullets" and lines and not all(
                line.startswith("- ") for line in lines):
            detail = "expects every notes line to start with '- '"
        elif style == "prose" and any(line.startswith("- ") for line in lines):
            detail = "expects prose notes, not '- ' bullet lines"
        if detail:
            warnings.append(f"Slide {n}: profile {name!r}: notes style "
                            f"{style!r} {detail}")
    return warnings


def merge_style(style, profile):
    """House style with the profile's banned phrases merged in."""
    extra = (profile or {}).get("banned_phrases") or []
    if not extra:
        return style
    import house_style
    return house_style.merge_lists(style, {"banned_phrases": extra})


# --------------------------------------------------------------------------
# lessons
# --------------------------------------------------------------------------

def lessons_paths(name, assets_dir):
    """(markdown, json) lesson files for a profile."""
    base = profile_path(name, assets_dir).with_suffix("")
    return (base.with_name(f"{base.name}.lessons.md"),
            base.with_name(f"{base.name}.lessons.json"))


def signature(check, level, message):
    """Stable id for a finding: no slide number, no figures."""
    return f"{check}:{level}:{normalise(message)}"


def normalise(message):
    text = SLIDE_PREFIX_RX.sub("", str(message).strip())
    return DIGITS_RX.sub(DIGIT_TOKEN, text).strip()


def record_findings(record):
    """(check, level, message) triples from a qa_summary record."""
    triples = []
    for check, result in (record.get("checks") or {}).items():
        if not isinstance(result, dict):
            continue
        for level, key in (("error", "errors"), ("warning", "warnings")):
            for message in result.get(key) or []:
                triples.append((check, level, message))
        for lines in (result.get("findings") or {}).values():
            for line in lines:
                triples.append((check, "finding", line))
    return triples


def record_lessons(record, name, assets_dir, *, now=None):
    """Fold a QA record into the profile's lessons files.

    Counts every distinct finding signature (once per record) in
    `<name>.lessons.json`. A signature that has now recurred and is not yet
    in `<name>.lessons.md` is appended there under a dated heading; one
    already present has its `(seen N×)` count refreshed in place when the
    count has grown. Returns every line appended or updated (in that order).
    """
    md_path, json_path = lessons_paths(name, assets_dir)
    counts = _load_counts(json_path)
    seen = {}
    for check, level, message in record_findings(record):
        seen.setdefault(signature(check, level, message),
                        (check, normalise(message)))
    for sig, (check, message) in seen.items():
        entry = counts.get(sig) or {"check": check, "message": message,
                                    "count": 0}
        entry["count"] += 1
        counts[sig] = entry

    body = md_path.read_text(encoding="utf-8") if md_path.is_file() else ""
    new_lines = []
    changed_lines = []
    for sig, (check, message) in seen.items():
        entry = counts[sig]
        if entry["count"] < MIN_LESSON_COUNT:
            continue
        line = _lesson_line(check, message, entry["count"])
        m = _stem_rx(check, message).search(body)
        if m:
            if m.group(0) != line:
                body = body[:m.start()] + line + body[m.end():]
                changed_lines.append(line)
            continue
        new_lines.append(line)

    json_path.parent.mkdir(parents=True, exist_ok=True)
    json_path.write_text(json.dumps(counts, indent=2, ensure_ascii=False)
                         + "\n", encoding="utf-8")
    if new_lines:
        heading = f"## {(now or date.today()).isoformat()}"
        if not body:
            body = f"# Lessons — {name}\n\nRecurring QA findings across " \
                   "deliveries for this profile.\n"
        if heading not in body:
            body = body.rstrip("\n") + f"\n\n{heading}\n"
        body = body.rstrip("\n") + "\n" + "\n".join(new_lines) + "\n"
    if new_lines or changed_lines:
        md_path.write_text(body, encoding="utf-8")
    return new_lines + changed_lines


def _lesson_line(check, message, count):
    return f"- [{check}] {message} (seen {count}×)"


def _stem(check, message):
    """The part of a lesson line that stays stable as the count grows."""
    return f"- [{check}] {message} ("


def _stem_rx(check, message):
    """Match the exact existing lesson line for (check, message), whatever
    its current `(seen N×)` count, so it can be refreshed in place."""
    return re.compile(rf"^{re.escape(_stem(check, message))}seen \d+×\)$",
                      re.MULTILINE)


def _load_counts(path):
    if not path.is_file():
        return {}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    return data if isinstance(data, dict) else {}
