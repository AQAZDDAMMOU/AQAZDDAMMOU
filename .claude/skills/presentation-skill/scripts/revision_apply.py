#!/usr/bin/env python3
"""
revision_apply.py — perform a revision plan's mechanical actions on the outline.

Split out of `revision_plan.py` (which builds and renders the plan) so both
files stay inside the module-size budget: everything here edits the outline's
raw markdown through `outline_ops`, so an auto-applied edit is byte-for-byte
the edit a human would have typed — no reformatting of the lines around it.

    from revision_apply import apply_plan
    ledger = apply_plan(plan, "outline.md")

Every action either applies or is skipped with a reason; nothing raises on a
stale plan, which is what makes re-running one a safe no-op. Deletes are
verified against the bullet text the plan recorded, never by index alone, so
a re-run (or an outline edited by hand in between) can never take out the
wrong line.

Known limitation: `_bullet_lines` re-derives build_deck.parse_outline's
"which '- ' lines are bullets" rule (parse_outline is at its size ceiling and
does not expose it). The literal prefixes, the quadrant-key rule and the
'**Data:**' block are mirrored exactly; what is *not* mirrored is
parse_outline's fall-through of a malformed data row back into `bullets`, and
a layout that build_deck auto-picks rather than the outline declaring it. Both
only shift bullet *numbering*, and the recorded-text check turns any such
divergence into a skip or a match-by-text — never a wrong deletion.
"""
import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import outline_ops
# _clean_bullet is build_deck's own bullet-text normaliser: borrowed
# rather than re-implemented, so recorded text and outline line agree
from build_deck import FIELD_KEYS, KV_RE, _clean_bullet

# Copy of the outline as it was before an applied edit, next to the outline
# itself; overwritten each time, so there is exactly one rollback point.
BACKUP_SUFFIX = ".bak"
# Record of what a plan already did, next to the plan itself:
# '<plan stem>.applied.json'
LEDGER_SUFFIX = ".applied.json"
ALREADY_APPLIED = "already applied"
# The ledger's "applied" list holds item *fingerprints*, not the positional
# ids: a plan rebuilt after a second review round renumbers its items, so an
# id-keyed ledger would skip brand-new work under an old item's number. The
# marker is what tells a fingerprint-keyed file from one written before this
# (whose positional ids cannot be matched safely and are therefore ignored).
LEDGER_KEY_SCHEME = "fingerprint"
TEXT_OPS = frozenset({"replace", "retitle", "add_bullet", "delete_bullet",
                      "delete_anchor"})
SLIDE_OPS = frozenset({"kill_slide", "split_slide", "merge_slide",
                       "move_slide"})
# a move/merge renumbers the slides around it, so every slide-level op still
# holding plan-time numbers has to wait for the next round
RENUMBERING_OPS = frozenset({"move_slide", "merge_slide"})
STALE_NUMBERING = "slide numbering changed — re-plan and re-apply"
_HEADING_DIRECTIVES = ("Heading", "Title")
_LAYOUT_RX = re.compile(r"^\*\*Layout:\*\*\s*(.*)$")
# build_deck.parse_outline's non-bullet '- ' rows, with its literal
# (case-sensitive) prefixes: '- Card 1:', '- Left:', '- Value=…', '- Stat…'
_NUMBERED_RX = re.compile(r"^(?:Card|Step|Tile) \d+:")
_LITERAL_PREFIXES = ("Item:", "Bar:", "Milestone:", "Node:", "Layer:",
                     "Lane:", "Left:", "Right:", "Value=", "Stat")
_DATE_KEY = "date="
# '- Point 1: "…"' / '- Finding 2: "…"' (exec-summary, exec-summary-scqa) are
# bullets, but build_deck stores them *without* the prefix:
#   bullets.append(clean(item.split(":", 1)[1]))
# so _bullet_text has to strip it the same way or a recorded delete never
# matches its line. The coupling is per-prefix: any further prefix build_deck
# strips before storing a bullet must be mirrored here too.
_PREFIXED_BULLET_RX = re.compile(r"^(?:Point|Finding) \d+:")
# '- Q1:' … '- Q4:' are quadrant labels on a matrix slide and ordinary
# content anywhere else — the one FIELD_KEYS rule that depends on layout
_QUADRANT_KEYS = ("Q1", "Q2", "Q3", "Q4")
_MATRIX_LAYOUT = "matrix-2x2"


def _block_layout(lines):
    """The block's declared '**Layout:**' value, '' when it has none."""
    for raw in lines:
        m = _LAYOUT_RX.match(raw.strip())
        if m:
            return m.group(1).strip().lower()
    return ""


def _is_directive(item_text, layout):
    for key in FIELD_KEYS:
        if key in _QUADRANT_KEYS and layout != _MATRIX_LAYOUT:
            continue
        if item_text.startswith(key + ":"):
            return True
    return False


def _is_structured(item_text):
    """A '- ' row parse_outline routes to cards/steps/items/… , not bullets."""
    if _NUMBERED_RX.match(item_text) or item_text.startswith(_LITERAL_PREFIXES):
        return True
    return _DATE_KEY in item_text.lower() and bool(KV_RE.search(item_text))


def _bullet_lines(lines):
    """Indices of a block's plain-bullet lines, in outline order.

    Mirrors build_deck.parse_outline: '- ' lines that are neither a known
    '- Key:' directive nor a structured row ('- Card 1:', '- Left:', …), and
    are not inside a '**Data:**' block.
    """
    layout = _block_layout(lines)
    found, in_data = [], False
    for i, raw in enumerate(lines):
        stripped = raw.strip()
        if stripped.startswith("**"):
            in_data = stripped.startswith("**Data:**")
            continue
        if not stripped.startswith("- "):
            continue
        item_text = stripped[2:].strip()
        if _is_directive(item_text, layout):
            in_data = False
            continue
        if in_data or _is_structured(item_text):
            continue
        found.append(i)
    return found


def _bullet_text(raw):
    """The bullet text build_deck would parse out of one '- ' line."""
    item = raw.strip()[2:].strip()
    if _PREFIXED_BULLET_RX.match(item):
        item = item.split(":", 1)[1]
    return _clean_bullet(item)


def _set_heading(block, title, eol):
    _old, attrs = _split_heading(block["heading"])
    heading = f" {title}" + (f" {attrs}" if attrs else "")
    block["heading"] = heading
    block["lines"][0] = f"## Slide {block['n']}:{heading}{eol}"


def _split_heading(heading):
    m = outline_ops.ATTR_RX.search(heading)
    if m:
        return heading[:m.start()].strip(), m.group(0).strip()
    return heading.strip(), ""


def _directive_key(stripped):
    """'Heading'/'Title' when this line is that directive, else None."""
    for key in _HEADING_DIRECTIVES:
        if stripped.startswith(f"- {key}:"):
            return key
    return None


def _write_directive(raw, key, title, eol):
    """`raw` rewritten to carry `title`, keeping its indent and quoting."""
    stripped = raw.strip()
    quote = '"' if stripped.rstrip().endswith('"') else ""
    indent = raw[:len(raw) - len(raw.lstrip())]
    return f"{indent}- {key}: {quote}{title}{quote}{eol}"


def _sync_heading_directives(lines, title, eol, *, mirroring=None):
    """Point every '- Heading:'/'- Title:' line on the slide at `title`.

    `mirroring` limits the rewrite to directives whose value still reads
    exactly that — the way a `replace` on the heading tells an echo of the
    heading (which is what actually renders) from a slide's own separate
    copy, e.g. a title slide's '- Title:'. `retitle` renames the slide
    outright and passes no filter.
    """
    for i, raw in enumerate(lines[1:], 1):
        stripped = raw.strip()
        key = _directive_key(stripped)
        if key is None:
            continue
        if mirroring is not None and \
                stripped[len(key) + 3:].strip().strip('"') != mirroring:
            continue
        lines[i] = _write_directive(raw, key, title, eol)


def _retitle(block, title, eol):
    if _split_heading(block["heading"])[0] == title:
        return None, f"heading already reads {title!r}"
    _set_heading(block, title, eol)
    _sync_heading_directives(block["lines"], title, eol)
    return block, ""


def _replace(block, old, new, eol):
    lines = block["lines"]
    if old in block["heading"]:
        was = _split_heading(block["heading"])[0]
        now = was.replace(old, new, 1)
        _set_heading(block, now, eol)
        # the '- Heading:' twin renders instead of the '## Slide N:' line, so
        # an edit that stopped at the heading would change nothing on screen
        _sync_heading_directives(lines, now, eol, mirroring=was)
        return block, ""
    for i, raw in enumerate(lines[1:], 1):
        if old in raw:
            lines[i] = raw.replace(old, new, 1)
            return block, ""
    return None, f"{old!r} is not on this slide any more"


def _add_bullet(block, text, eol):
    lines = block["lines"]
    wanted = f"- {text}"
    if any(raw.strip() == wanted for raw in lines):
        return None, "bullet is already on the slide"
    bullets = _bullet_lines(lines)
    after = bullets[-1] if bullets else (
        max((i for i, raw in enumerate(lines) if raw.strip()), default=0))
    lines.insert(after + 1, wanted + eol)
    return block, ""


def _matches(raw, expected):
    """True when a raw '- ' line still holds the text the plan recorded.

    Compared against every form build_deck could have stored for that line:
    its cleaned bullet text, the raw item, and the raw item minus a
    'Point N:'/'Finding N:' prefix (**Copy:** locked / verbatim outlines keep
    the quotes `_clean_bullet` would strip).
    """
    stripped = raw.strip()[2:].strip()
    unprefixed = (stripped.split(":", 1)[1].strip()
                  if _PREFIXED_BULLET_RX.match(stripped) else stripped)
    return expected in (_bullet_text(raw), stripped, unprefixed)


def _delete_bullet(block, position, expected=None):
    """Delete the bullet at `position` (the reviewer's 1-based number).

    `expected` is the bullet text the plan recorded when it was built. With
    it the delete is idempotent and edit-safe: the line at `position` is only
    removed when it still reads that way, otherwise the slide's bullets are
    searched for that exact text, and a bullet that is simply gone (a second
    run of the same plan) is skipped rather than taking out its neighbour.
    """
    lines = block["lines"]
    bullets = _bullet_lines(lines)
    in_range = 1 <= position <= len(bullets)
    if expected is None:          # hand-written plan: index is all we have
        if not in_range:
            return None, (f"bullet {position} is out of range "
                          f"({len(bullets)} bullet(s) on the slide)")
        del lines[bullets[position - 1]]
        return block, ""
    if in_range and _matches(lines[bullets[position - 1]], expected):
        del lines[bullets[position - 1]]
        return block, ""
    for i in bullets:
        if _matches(lines[i], expected):
            del lines[i]
            return block, ""
    return None, "bullet text no longer present"


def _delete_anchor(block, field):
    if not field:
        return None, "nothing located to delete — resolve by hand"
    if field["kind"] != "bullet":
        return None, (f"the anchor points at a {field['kind']}, which "
                      "delete-this does not touch — resolve by hand")
    return _delete_bullet(block, field["index"] + 1, field.get("text"))


def _apply_text_op(md_text, item):
    """(new markdown, "") on success, (None, reason) when skipped."""
    action = item["action"]
    number = item["outline_slide"]
    front, blocks = outline_ops.blocks(md_text)
    block = next((b for b in blocks if b["n"] == number), None)
    if block is None:
        return None, f"the outline has no slide {number} any more"
    op = action["op"]
    # every line built here (not merely reproduced) carries the file's own
    # terminator, so a CRLF outline never comes back with LF lines in it
    eol = outline_ops.doc_eol(md_text)
    if op == "replace":
        block, reason = _replace(block, action["old"], action["new"], eol)
    elif op == "retitle":
        block, reason = _retitle(block, action["text"], eol)
    elif op == "add_bullet":
        block, reason = _add_bullet(block, action["text"], eol)
    elif op == "delete_bullet":
        field = item.get("field") or {}
        block, reason = _delete_bullet(block, action["index"],
                                       field.get("text"))
    else:
        block, reason = _delete_anchor(block, item.get("field"))
    if block is None:
        return None, reason
    return outline_ops.render(front, blocks), ""


def _apply_slide_op(md_text, item):
    action = item["action"]
    number = item["outline_slide"]
    op = action["op"]
    try:
        if op == "kill_slide":
            front, blocks = outline_ops.blocks(md_text)
            block = next((b for b in blocks if b["n"] == number), None)
            if block is None:
                return None, f"the outline has no slide {number} any more"
            blocks.remove(block)
            return outline_ops.render(front, blocks), ""
        if op == "move_slide":
            return outline_ops.move(md_text, number, action["to"]), ""
        if op == "split_slide":
            return outline_ops.split(md_text, number), ""
        return outline_ops.merge(md_text, number, action["other"]), ""
    except ValueError as exc:
        return None, str(exc)


def item_key(item) -> str:
    """The ledger key for one plan item: its fingerprint, else its id.

    `build_plan` stamps every item with a `fingerprint` keyed on the
    comment's own identity (its thread id) or, for emailed feedback, on the
    text of the note — both stable across a re-plan. A hand-written plan
    that carries none falls back to the display id, which is positional:
    fine for a plan nobody rebuilds, and no worse than before.
    """
    return str(item.get("fingerprint") or item["id"])


def _read_ledger(ledger_path) -> tuple:
    """`(data, unreadable)`: the ledger file as a dict, plus whether a file
    was there but could not be used.

    An absent, unreadable or malformed ledger yields an empty dict either
    way: a damaged record must never stop a plan from being applied, only
    from being applied twice. `unreadable` is False only when there is
    genuinely nothing on disk to report on (missing path, or no file) —
    it is True when a file exists but couldn't be parsed into a usable
    ledger, so callers can warn that the idempotency guard is not in
    effect for this run rather than silently reapplying already-done work.
    """
    if not ledger_path:
        return {}, False
    path = Path(ledger_path)
    if not path.is_file():
        return {}, False
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (ValueError, OSError):
        return {}, True
    if not isinstance(data, dict):
        return {}, True
    return data, False


def load_applied(ledger_path) -> list:
    """The item keys `<stem>.applied.json` already records as applied.

    Empty for a ledger written before the keys became fingerprints: its
    positional ids belong to whatever plan was current then, and matching a
    re-planned item against them would skip work that was never done.
    """
    data, _ = _read_ledger(ledger_path)
    return _applied_keys(data)


def _applied_keys(data) -> list:
    if data.get("key_scheme") != LEDGER_KEY_SCHEME:
        return []
    applied = data.get("applied")
    return [str(k) for k in applied] if isinstance(applied, list) else []


def _ledger_notes(data, ledger_path, plan, unreadable=False) -> list:
    """Warnings about the ledger found on disk — never reasons to refuse."""
    if unreadable:
        return [f"applied-ledger {Path(ledger_path).name} is unreadable — "
                "applying without idempotency guard"]
    if not data:
        return []
    name = Path(ledger_path).name
    notes = []
    if data.get("key_scheme") != LEDGER_KEY_SCHEME:
        notes.append(f"{name} predates comment-identity keys and was "
                     "ignored — nothing is skipped as already applied")
    was, now = data.get("deck"), Path(plan.get("deck") or "").name
    if was and now and was != now:
        notes.append(f"{name} was written for {was}, not {now} — applying "
                     "anyway; check the two decks are the same review")
    return notes


def merge_applied(ledger, ledger_path) -> dict:
    """`ledger` as it should be written to `ledger_path`.

    Mixed tense by design: "applied" accumulates across every round of the
    plan (that is what makes the skip work), while "skipped", "manual",
    "applied_ids", "warnings" and "backup" describe *this* run only.
    """
    merged = list(load_applied(ledger_path))
    for key in ledger.get("applied") or []:
        if key not in merged:
            merged.append(key)
    return {**ledger, "applied": merged, "key_scheme": LEDGER_KEY_SCHEME}


def _sort_index(item):
    """Highest bullet index first within a slide (deletes stay in step)."""
    action = item["action"]
    if "index" in action:
        return action["index"]
    field = item.get("field") or {}
    return field.get("index", -1)


def apply_plan(plan, outline_path, *, ledger_path=None) -> dict:
    """Perform the plan's mechanical actions on the outline; write it back.

    Text-level edits run first, per slide, highest bullet index first (so an
    earlier delete never shifts a later one); slide-level ops follow in
    descending slide number, for the same reason. Anything the outline no
    longer supports — a replace whose text is gone, a deleted slide — is
    skipped with a reason rather than raising, which makes re-running a plan
    a safe no-op. A move or merge renumbers its neighbours, so it ends the
    slide-level pass: every slide op behind it is skipped for the next round
    rather than acting on numbers that have moved under it.

    Resolved threads are skipped too — the reviewer marked them settled, so
    the edit has usually already been made.

    `ledger_path` (the CLI passes the plan's own `<stem>.applied.json`) makes
    that no-op exact rather than best-effort: every item whose `item_key`
    the ledger already lists under "applied" is skipped as `already
    applied`, before the outline is even looked at. Without it,
    re-application is only as idempotent as the text matching — `replace "X"
    with "X."` would fire a second time, since the replacement still
    contains what it replaced. Callers that pass a ledger path are expected
    to write the merged ledger back to it (`merge_applied`, which
    `revision_plan.write_ledger` does).

    Those keys are comment fingerprints, not the positional item ids, so the
    same ledger can sit next to a plan rebuilt for a second review round:
    round 2's brand-new "c1" is a different comment from round 1's and is
    applied, while the item that *is* the same thread is still skipped.
    Anything odd about the ledger found on disk — an older, id-keyed file, a
    ledger written for a different deck — is reported under the returned
    ledger's "warnings" and applied through, never refused.

    The outline is read and written as bytes, so a CRLF file stays CRLF, and
    it is only written when something actually changed — a plan that applies
    nothing leaves the file byte-for-byte alone. When it is written, the
    original is copied to `<outline>.bak` first (overwriting any earlier
    backup) and that path is reported as the ledger's "backup" (None when
    nothing was written).
    """
    path = Path(outline_path)
    original = path.read_bytes()
    md_text = original.decode("utf-8")
    applied, skipped, manual = [], [], []
    text_items, slide_items = [], []
    prior, prior_unreadable = _read_ledger(ledger_path)
    done = set(_applied_keys(prior))
    for item in plan["items"]:
        op = (item.get("action") or {}).get("op", "manual")
        if item_key(item) in done:
            skipped.append((item["id"], ALREADY_APPLIED))
        elif item.get("resolved"):
            skipped.append((item["id"], "resolved"))
        elif op not in TEXT_OPS and op not in SLIDE_OPS:
            manual.append(item["id"])
        elif item.get("outline_slide") is None:
            skipped.append((item["id"], "no outline slide matched"))
        elif op in TEXT_OPS:
            text_items.append(item)
        else:
            slide_items.append(item)

    ordered = sorted(text_items,
                     key=lambda i: (-i["outline_slide"], -_sort_index(i)))
    ordered += sorted(slide_items, key=lambda i: -i["outline_slide"])
    renumbered = False
    for item in ordered:
        op = item["action"]["op"]
        if renumbered and op in SLIDE_OPS:
            skipped.append((item["id"], STALE_NUMBERING))
            continue
        handler = _apply_text_op if op in TEXT_OPS else _apply_slide_op
        new_text, reason = handler(md_text, item)
        if new_text is None:
            skipped.append((item["id"], reason))
        else:
            md_text = new_text
            applied.append(item)
            renumbered = renumbered or op in RENUMBERING_OPS
    ledger = {"applied": [item_key(i) for i in applied],
              "applied_ids": [i["id"] for i in applied],
              "deck": Path(plan.get("deck") or "").name,
              "skipped": skipped, "manual": manual,
              "warnings": _ledger_notes(prior, ledger_path, plan,
                                        unreadable=prior_unreadable),
              "backup": None}
    new_bytes = md_text.encode("utf-8")
    if new_bytes != original:
        backup = path.with_name(path.name + BACKUP_SUFFIX)
        backup.write_bytes(original)
        path.write_bytes(new_bytes)
        ledger["backup"] = str(backup)
    return ledger
