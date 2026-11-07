#!/usr/bin/env python3
"""
revision_plan.py — reviewer comments (and emailed feedback) -> a revision plan.

The comment-driven revision loop: a reviewer marks up the delivered deck in
PowerPoint (or mails back a `feedback.md`), and this module turns that into
work on the *outline*, which is where the deck is actually authored.

    python3 scripts/edit_deck.py comments plan reviewed.pptx \\
        --outline outline.md [--feedback feedback.md] --output plan
    python3 scripts/edit_deck.py comments apply plan.json outline.md
    python3 scripts/patch_deck.py outline.md reviewed.pptx 2,5 --output v2.pptx
    python3 scripts/edit_deck.py comments strip v2.pptx --output final.pptx

Each comment is mapped to the outline slide it belongs to (by the `PK_KEY`
slide tag, falling back to deck position), to the field it points at (via the
comment's shape/position anchor, or the text it quotes), and — when it reads
as a mechanical instruction in `revision_grammar`'s command grammar — to an
action that `apply_plan` performs on the outline markdown itself. Everything
else stays `manual`: the plan's markdown twin is the work list for Claude.

The mechanical edits themselves are performed by `revision_apply.apply_plan`,
re-exported here. It is idempotent: re-running it on an already-applied plan
changes nothing and reports every item as skipped, so a half-finished round
can always be finished by running it again.
"""
import hashlib
import json
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from lxml import etree
from pptx import Presentation

import outline_ops
from outline_replay import replay
from qa_check import iter_shapes
from revision_apply import (LEDGER_SUFFIX,  # noqa: F401  (re-exported)
                            apply_plan, merge_applied)
from revision_grammar import (ACTION_RX, FIELD_KINDS,  # noqa: F401
                              MATCH_RATIO, locate_field, parse_action,
                              parse_feedback)
from shape_groups import canvas_shapes
from slide_tags import read_tags, slide_key

TAG_KEY = "PK_KEY"
# Ledger keys. An item's `id` is positional ("c3" = the third comment in
# deck order), so it shifts the moment a comment is stripped, resolved or
# added — useless for telling a later round's work from an earlier one's.
# `fingerprint` is keyed on the comment's own identity instead: a thread id
# (legacy 'slide:author:idx', modern the cm GUID) survives a re-plan, and
# emailed feedback is keyed on the text of the note itself.
FP_COMMENT = "comment:"
FP_FEEDBACK = "feedback:"
FP_TEXT_CHARS = 16
ANCHOR_TEXT_CHARS = 80
ANCHOR_NEAR_EMU = 685800          # 0.75in — a balloon further off anchors nothing
# etree.XMLSyntaxError: a damaged XML part anywhere under the package
# (a truncated comment part, a mangled rels file) — reported, not raised
DECK_READ_ERRORS = (ValueError, KeyError, OSError, zipfile.BadZipFile,
                    etree.XMLSyntaxError)


# ── comment anchors ─────────────────────────────────────────────────────────
def _bounds(shape):
    try:
        return (int(shape.left), int(shape.top), int(shape.width),
                int(shape.height))
    except (TypeError, ValueError):
        return None


def _nearest_shape(slide, x, y):
    """The shape containing (x, y), else the closest centre within 0.75in."""
    best, best_distance = None, None
    for shape in canvas_shapes(slide.shapes):
        box = _bounds(shape)
        if box is None:
            continue
        left, top, width, height = box
        if left <= x <= left + width and top <= y <= top + height:
            return shape
        dx = left + width / 2 - x
        dy = top + height / 2 - y
        distance = (dx * dx + dy * dy) ** 0.5
        if best_distance is None or distance < best_distance:
            best, best_distance = shape, distance
    if best is not None and best_distance <= ANCHOR_NEAR_EMU:
        return best
    return None


def anchor_info(slide, anchor) -> dict | None:
    """{"shape": name, "text": first 80 chars} for a comment's anchor.

    A modern comment's shape mark wins; a legacy balloon only has a position,
    so the shape under it (or the nearest one within 0.75in) is reported.
    None when the anchor is missing or points at empty canvas.
    """
    if not anchor:
        return None
    shape = None
    if anchor.get("shape_id") is not None:
        shape = next((s for s in iter_shapes(slide.shapes)
                      if s.shape_id == anchor["shape_id"]), None)
    elif anchor.get("x_emu") is not None and anchor.get("y_emu") is not None:
        shape = _nearest_shape(slide, anchor["x_emu"], anchor["y_emu"])
    if shape is None:
        return None
    text = ""
    if getattr(shape, "has_text_frame", False):
        text = shape.text_frame.text.strip()
    return {"shape": shape.name, "text": text[:ANCHOR_TEXT_CHARS]}


# ── deck slide -> outline slide ─────────────────────────────────────────────
def _authored_numbers(md_text, slides):
    """Per replayed slide: the '## Slide N:' number that authored it.

    None for slides the outline has no block for (auto agenda/sources), which
    are reported and left alone — they are regenerated, not edited.
    """
    _front, blocks = outline_ops.blocks(md_text)
    by_line = {b["start"]: b["n"] for b in blocks if b["n"] is not None}
    return [by_line.get(p.get("_line")) for p in slides]


class _SlideMap:
    """deck position -> (outline slide number, key), warning once per reason."""

    def __init__(self, tags, keys, numbers):
        self.tags = tags
        self.by_key = {}
        for i, key in enumerate(keys):
            self.by_key.setdefault(key, numbers[i])
        self.numbers = numbers
        self.warnings = []

    def _warn(self, message):
        if message not in self.warnings:
            self.warnings.append(message)

    def lookup(self, deck_pos):
        tag = self.tags[deck_pos - 1] if deck_pos <= len(self.tags) else None
        key = (tag or {}).get(TAG_KEY, "")
        if key:
            if key in self.by_key and self.by_key[key] is not None:
                return self.by_key[key], key
            self._warn(f"deck slide {deck_pos} ('{key}') is not in the "
                       "outline any more — its comments need a human")
            return None, key
        if deck_pos <= len(self.numbers) and self.numbers[deck_pos - 1]:
            self._warn("the deck carries no PK_KEY slide tags — comments "
                       "mapped to the outline by position")
            return self.numbers[deck_pos - 1], ""
        self._warn(f"deck slide {deck_pos} has no outline slide "
                   f"(the outline builds {len(self.numbers)} slide(s))")
        return None, ""


def _heading_number(slides, numbers, title):
    """Outline slide number whose heading matches a '"Title":' feedback ref."""
    headings = [{"heading": s.get("heading") or s.get("title") or ""}
                for s in slides]
    for want_exact in (True, False):
        for i, slide in enumerate(headings):
            if numbers[i] is None:
                continue
            found = locate_field(slide, title)
            if found and found["kind"] == "heading" and (
                    not want_exact or found["text"] == title):
                return numbers[i]
    return None


# ── building the plan ───────────────────────────────────────────────────────
def _field_for(action, anchor, slide):
    """The outline field an item is about (the quoted old text wins).

    'delete bullet N' resolves to the bullet that number means *now*, text
    included: `revision_apply` deletes by that text, so a re-run (or an
    outline edited in between) can never take out the neighbouring line.
    """
    if slide is None:
        return None
    if action.get("op") == "delete_bullet":
        bullets = slide.get("bullets") or []
        index = action["index"] - 1
        if not 0 <= index < len(bullets):
            return None
        return {"kind": "bullet", "index": index, "text": str(bullets[index])}
    for query in (action.get("old"), (anchor or {}).get("text")):
        found = locate_field(slide, query) if query else None
        if found:
            return found
    return None


def _text_fingerprint(text) -> str:
    """Ledger key for an emailed-feedback note: a digest of the note itself.

    Feedback lines carry no id of their own, so the text is the identity:
    re-sending the same `feedback.md` re-applies nothing, while an edited or
    new note is new work.
    """
    digest = hashlib.sha1(text.strip().encode("utf-8")).hexdigest()
    return FP_FEEDBACK + digest[:FP_TEXT_CHARS]


def _item(item_id, source, text, action, *, fingerprint="", author="",
          outline_slide=None, key="", anchor=None, field=None, thread=(),
          deck_slide=None, resolved=False):
    """One plan item — the shape both comment and feedback items take.

    `id` is the display handle (`revision-plan.md` refers to items by it);
    `fingerprint` is the stable key the applied-ledger uses.
    """
    return {
        "id": item_id,
        "fingerprint": fingerprint,
        "source": source,
        "author": author,
        "deck_slide": deck_slide,
        "outline_slide": outline_slide,
        "key": key,
        "anchor": anchor,
        "field": field,
        "text": text,
        "thread": list(thread),
        "resolved": resolved,
        "action": action,
    }


def _comment_items(pptx, prs, slide_map, slides_by_number):
    """Plan items for every comment thread in the deck, replies folded in."""
    from pptx_comments import read_comments
    items, by_thread = [], {}
    for comment in read_comments(pptx):
        if comment.parent_id:
            parent = by_thread.get(comment.parent_id)
            if parent is not None:
                parent["thread"].append({"author": comment.author,
                                         "text": comment.text,
                                         "created": comment.created})
                parent["resolved"] = parent["resolved"] or comment.resolved
                continue
        outline_slide, key = slide_map.lookup(comment.slide_pos)
        slide = slides_by_number.get(outline_slide)
        anchor = None
        if comment.slide_pos <= len(prs.slides):
            anchor = anchor_info(prs.slides[comment.slide_pos - 1],
                                 comment.anchor)
        action = parse_action(comment.text)
        item = _item(
            f"c{len(items) + 1}", "comment", comment.text, action,
            fingerprint=FP_COMMENT + comment.thread_id,
            author=comment.author, deck_slide=comment.slide_pos,
            outline_slide=outline_slide, key=key, anchor=anchor,
            field=_field_for(action, anchor, slide),
            thread=[{"author": comment.author, "text": comment.text,
                     "created": comment.created}],
            resolved=bool(comment.resolved))
        items.append(item)
        if not comment.parent_id:
            by_thread[comment.thread_id] = item
    return items


def _feedback_items(feedback_md, slides, numbers, slides_by_number, warnings):
    """Plan items for an emailed feedback.md (no deck anchor to go on)."""
    items = []
    text = Path(feedback_md).read_text(encoding="utf-8")
    for entry in parse_feedback(text):
        ref = entry["slide_ref"]
        if isinstance(ref, str):
            outline_slide = _heading_number(slides, numbers, ref)
            if outline_slide is None:
                warnings.append(f"feedback: no outline slide titled {ref!r}")
        else:
            outline_slide = ref
            if ref is not None and ref not in numbers:
                warnings.append(f"feedback: the outline has no slide {ref}")
                outline_slide = None
        action = parse_action(entry["text"])
        slide = slides_by_number.get(outline_slide)
        items.append(_item(
            f"f{len(items) + 1}", "feedback", entry["text"], action,
            fingerprint=_text_fingerprint(entry["text"]),
            outline_slide=outline_slide,
            field=_field_for(action, None, slide)))
    return items


def build_plan(pptx, outline, *, feedback_md=None, assets_dir=None) -> dict:
    """The revision plan for a reviewed deck: every comment, mapped and parsed.

    Raises ValueError when the outline does not replay or the deck cannot be
    read — a plan built on a broken input would silently mis-target edits.
    """
    result = replay(outline, assets_dir=assets_dir)
    if result.errors:
        raise ValueError(f"{outline} does not replay: "
                         + "; ".join(str(e) for e in result.errors))
    md_text = Path(outline).read_text(encoding="utf-8")
    numbers = _authored_numbers(md_text, result.slides)
    seen = set()
    keys = [slide_key(p, seen) for p in result.slides]
    slides_by_number = {n: p for n, p in zip(numbers, result.slides)
                        if n is not None}

    try:
        prs = Presentation(str(pptx))
        tags = read_tags(prs)
    except DECK_READ_ERRORS as exc:
        raise ValueError(f"{pptx} could not be read as a .pptx ({exc})") \
            from exc
    slide_map = _SlideMap(tags, keys, numbers)
    items = _comment_items(pptx, prs, slide_map, slides_by_number)

    warnings = list(slide_map.warnings)
    if feedback_md:
        items.extend(_feedback_items(feedback_md, result.slides, numbers,
                                     slides_by_number, warnings))
    return {"deck": str(pptx), "outline": str(outline), "items": items,
            "warnings": warnings}


# ── rendering ───────────────────────────────────────────────────────────────
def _action_line(item):
    action = dict(item["action"])
    op = action.pop("op")
    args = ", ".join(f"{k}={v!r}" for k, v in sorted(action.items()))
    line = f"action: {op}" + (f" ({args})" if args else "")
    field = item.get("field")
    if field:
        line += f" — field: {field['kind']}[{field['index']}]"
    return line


def render_plan_md(plan) -> str:
    """The plan as a work list, grouped by outline slide (unmapped last)."""
    out = [f"# Revision plan — {Path(plan['deck']).name}", "",
           f"Outline: `{plan['outline']}`  ·  items: {len(plan['items'])}", ""]
    for warning in plan.get("warnings") or []:
        out.append(f"> WARNING: {warning}")
    if plan.get("warnings"):
        out.append("")
    groups = {}
    for item in plan["items"]:
        groups.setdefault(item["outline_slide"], []).append(item)
    for number in sorted(groups, key=lambda n: (n is None, n or 0)):
        out.append(f"## Slide {number}" if number is not None
                   else "## Unmapped")
        out.append("")
        for item in groups[number]:
            who = item["author"] or item["source"]
            state = " [resolved]" if item.get("resolved") else ""
            out.append(f"- **[{item['id']}]**{state} ({who}) {item['text']}")
            out.append(f"  - {_action_line(item)}")
            if item.get("anchor"):
                out.append(f"  - anchor: {item['anchor']['shape']} — "
                           f"\"{item['anchor']['text']}\"")
            for reply in item["thread"][1:]:
                out.append(f"  - reply ({reply['author']}): {reply['text']}")
        out.append("")
    return "\n".join(out).rstrip() + "\n"


def write_plan(plan, output_stem) -> tuple[Path, Path]:
    """Write `<stem>.json` + `<stem>.md`; returns both paths."""
    stem = str(output_stem)
    json_path, md_path = Path(stem + ".json"), Path(stem + ".md")
    json_path.write_text(json.dumps(plan, indent=2, ensure_ascii=False) + "\n",
                         encoding="utf-8")
    md_path.write_text(render_plan_md(plan), encoding="utf-8")
    return json_path, md_path


def write_ledger(ledger, output_stem) -> Path:
    """Write the `<stem>.applied.json` record of what apply_plan did.

    Merged, not overwritten: the keys of everything an earlier round applied
    stay on the list, so `apply_plan(..., ledger_path=…)` keeps skipping
    them. The fields are mixed-tense on purpose — "applied" is cumulative
    across rounds, everything else ("applied_ids", "skipped", "manual",
    "warnings", "backup", "deck") describes the run that just finished.
    """
    path = Path(str(output_stem) + LEDGER_SUFFIX)
    payload = merge_applied(ledger, path)
    path.write_text(json.dumps(payload, indent=2, ensure_ascii=False) + "\n",
                    encoding="utf-8")
    return path


# ── the `edit_deck comments` CLI ────────────────────────────────────────────
def _cli_list(args):
    from pptx_comments import read_comments
    comments = read_comments(args.pptx)
    for comment in comments:
        state = "resolved" if comment.resolved else "open"
        reply = "  reply: " if comment.parent_id else ""
        print(f"Slide {comment.slide_pos}  [{state}] {comment.author}: "
              f"{reply}{comment.text}")
    print(f"{len(comments)} comment(s) in {Path(args.pptx).name}.")
    return True


def _cli_plan(args):
    plan = build_plan(args.pptx, args.outline, feedback_md=args.feedback,
                      assets_dir=args.assets_dir)
    json_path, md_path = write_plan(plan, args.output)
    for warning in plan["warnings"]:
        print(f"  [WARN] {warning}", file=sys.stderr)
    auto = sum(1 for i in plan["items"] if i["action"]["op"] != "manual")
    print(f"{len(plan['items'])} item(s), {auto} mechanical.")
    print(f"Wrote: {json_path} and {md_path}")
    return True


def _cli_apply(args):
    plan = json.loads(Path(args.plan).read_text(encoding="utf-8"))
    stem = str(args.plan)[:-len(".json")] if str(args.plan).endswith(".json") \
        else str(args.plan)
    # the plan's own ledger: what an earlier run of *this* plan already did
    ledger = apply_plan(plan, args.outline,
                        ledger_path=Path(stem + LEDGER_SUFFIX))
    ledger_path = write_ledger(ledger, stem)
    for warning in ledger.get("warnings") or []:
        print(f"  [WARN] {warning}", file=sys.stderr)
    for item_id, reason in ledger["skipped"]:
        print(f"  [SKIP] {item_id}: {reason}", file=sys.stderr)
    print(f"{len(ledger['applied'])} applied, {len(ledger['skipped'])} "
          f"skipped, {len(ledger['manual'])} left manual.")
    if ledger.get("backup"):
        print(f"Backup of the outline as it was: {ledger['backup']}")
        print(f"Wrote: {args.outline} and {ledger_path}")
    else:
        print(f"Outline unchanged. Wrote: {ledger_path}")
    return True


def _cli_strip(args):
    from pptx_comments import strip_comments
    slides = None
    if args.slides:
        slides = [int(n) for n in str(args.slides).split(",") if n.strip()]
    removed = strip_comments(args.pptx, args.output, slides=slides,
                             resolved_only=args.resolved)
    if args.resolved:
        print("Note: legacy comments carry no resolved state and are kept.")
    print(f"Removed {removed} comment(s) → {args.output}")
    return True


def run_cli(args) -> bool:
    """Dispatch `edit_deck comments <list|plan|apply|strip>`."""
    handlers = {"list": _cli_list, "plan": _cli_plan, "apply": _cli_apply,
                "strip": _cli_strip}
    try:
        return handlers[args.comments_cmd](args)
    except DECK_READ_ERRORS as exc:
        print(f"  [ERROR] {exc}", file=sys.stderr)
        return False
