#!/usr/bin/env python3
"""
pptx_comments.py — write review notes into a .pptx as real PowerPoint comments.

QA findings are far more useful pinned to the slide they belong to than in a
terminal log, so `qa_summary --comments` pipes every error and warning
through here.

This writes the **legacy** comment parts (`ppt/commentAuthors.xml` plus one
`ppt/comments/commentN.xml` per slide, `p:cmLst/p:cm`), not the newer
modern-comments parts: legacy comments are what every PowerPoint version
still opens, and python-pptx round-trips a deck containing them.  Keynote
and Google Slides ignore them (they neither show nor destroy them).

python-pptx has no comment API, so the deck is rewritten as a zip: every
existing part is copied through byte-for-byte, the comment parts are added,
and three files are amended — the slide's `.rels`, `presentation.xml.rels`
and `[Content_Types].xml`.

Reading goes the other way and covers **both** formats: `read_comments`
returns the legacy `p:cm` balloons *and* the modern comments PowerPoint
2019+/365 writes (`ppt/comments/modernComment_*.xml`, threaded replies,
resolved state, shape anchors), so reviewer feedback made in any client can
be turned back into work.  `strip_comments` is the eraser — by slide, or
only the resolved threads — and leaves no orphan part, rel or override.

    from pptx_comments import add_comments, read_comments, strip_comments
    add_comments("deck.pptx", "deck-reviewed.pptx",
                 [(1, "axis starts at 40"), (3, "Total does not add up")])
    for c in read_comments("reviewed.pptx"):
        print(c.slide_pos, c.author, c.text)
    strip_comments("reviewed.pptx", "clean.pptx", resolved_only=True)
"""
import posixpath
import re
import shutil
import zipfile
from dataclasses import dataclass
from typing import Optional
from datetime import datetime, timezone
from pathlib import Path

from lxml import etree

NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
REL_COMMENTS = ("http://schemas.openxmlformats.org/officeDocument/2006/"
                "relationships/comments")
REL_AUTHORS = ("http://schemas.openxmlformats.org/officeDocument/2006/"
               "relationships/commentAuthors")
CT_COMMENTS = ("application/vnd.openxmlformats-officedocument."
               "presentationml.comments+xml")
CT_AUTHORS = ("application/vnd.openxmlformats-officedocument."
              "presentationml.commentAuthors+xml")

# Modern comments (PowerPoint 2019+/365) live alongside the legacy parts.
NS_P188 = "http://schemas.microsoft.com/office/powerpoint/2018/8/main"
NS_PC = "http://schemas.microsoft.com/office/powerpoint/2013/main/command"
RT_MODERN = "http://schemas.microsoft.com/office/2018/10/relationships/comments"
RT_MODERN_AUTHORS = ("http://schemas.microsoft.com/office/2018/10/"
                     "relationships/authors")
CT_MODERN_COMMENTS = "application/vnd.ms-powerpoint.comments+xml"
CT_MODERN_AUTHORS = "application/vnd.ms-powerpoint.authors+xml"
STATUS_RESOLVED = "resolved"
KIND_LEGACY = "legacy"
KIND_MODERN = "modern"

AUTHORS_PART = "ppt/commentAuthors.xml"
PRESENTATION_PART = "ppt/presentation.xml"
MODERN_AUTHORS_PART = "ppt/authors.xml"
COMMENTS_DIR = "ppt/comments"
PRESENTATION_RELS = "ppt/_rels/presentation.xml.rels"
CONTENT_TYPES = "[Content_Types].xml"
POS_X_EMU = 228600          # 0.25in from the slide's left edge
POS_Y_EMU = 228600          # first balloon 0.25in down …
POS_STEP_EMU = 274320       # … each further balloon 0.3in lower
MAX_INITIALS = 3
XML_DECL = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'


def _q(ns, tag):
    return f"{{{ns}}}{tag}"


def _serialize(element):
    return XML_DECL + etree.tostring(element, xml_declaration=False,
                                     encoding="UTF-8")


def _initials(name):
    parts = [p for p in re.split(r"\s+", name.strip()) if p]
    if not parts:
        return "QA"
    return "".join(p[0].upper() for p in parts)[:MAX_INITIALS]


def _read_parts(src):
    """Every zip entry of `src` as bytes, plus the original entry order."""
    with zipfile.ZipFile(src) as zin:
        order = [info.filename for info in zin.infolist()]
        return order, {name: zin.read(name) for name in order}


def _write_parts(dst, order, parts):
    """Repack `parts`, keeping the source order and appending new names."""
    new_names = [name for name in parts if name not in order]
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for name in order + sorted(new_names):
            if name in parts:            # deleted parts simply drop out
                zout.writestr(name, parts[name])


def _slide_partnames(parts):
    """Real zip entry names of the slides, in `p:sldIdLst` order.

    Read from the package itself — `p:sldIdLst` → `presentation.xml.rels` —
    because PowerPoint never renumbers `slideN.xml`: deleting a slide leaves a
    gap and reordering rewrites only `p:sldIdLst`.  python-pptx papers over
    both by renaming the parts to `slide1..slideN` *in memory*, so its
    partnames stop matching the zip exactly when they matter most.
    """
    blob = parts.get(PRESENTATION_PART)
    if blob is None:
        raise ValueError(f"not a presentation: {PRESENTATION_PART} is missing")
    targets = {}
    rels = parts.get(PRESENTATION_RELS)
    if rels is not None:
        for rel in etree.fromstring(rels):
            if rel.get("TargetMode") == "External":
                continue
            targets[rel.get("Id")] = _resolve_target(PRESENTATION_PART,
                                                     rel.get("Target", ""))
    names = []
    for sld in etree.fromstring(blob).findall(
            f'{_q(NS_P, "sldIdLst")}/{_q(NS_P, "sldId")}'):
        target = targets.get(sld.get(_q(NS_R, "id")))
        if target is not None:
            names.append(target)
    return names


def _slide_index(partname):
    """The N in ppt/slides/slideN.xml (comment parts mirror that number)."""
    m = re.search(r"slide(\d+)\.xml$", partname)
    return int(m.group(1)) if m else None


def _next_rid(rels_root):
    used = {r.get("Id") for r in rels_root}
    n = 1
    while f"rId{n}" in used:
        n += 1
    return f"rId{n}"


def _rels_root(parts, rels_name):
    if rels_name in parts:
        return etree.fromstring(parts[rels_name])
    return etree.Element(_q(NS_REL, "Relationships"), nsmap={None: NS_REL})


def _add_rel(parts, rels_name, rel_type, target):
    """Add a relationship (or reuse an existing one of the same type)."""
    root = _rels_root(parts, rels_name)
    for rel in root:
        if rel.get("Type") == rel_type and rel.get("Target") == target:
            parts[rels_name] = _serialize(root)
            return rel.get("Id")
    rel = etree.SubElement(root, _q(NS_REL, "Relationship"))
    rel.set("Id", _next_rid(root))
    rel.set("Type", rel_type)
    rel.set("Target", target)
    parts[rels_name] = _serialize(root)
    return rel.get("Id")


def _add_override(parts, part_name, content_type):
    root = etree.fromstring(parts[CONTENT_TYPES])
    name = "/" + part_name
    for override in root.findall(_q(NS_CT, "Override")):
        if override.get("PartName") == name:
            return
    override = etree.SubElement(root, _q(NS_CT, "Override"))
    override.set("PartName", name)
    override.set("ContentType", content_type)
    parts[CONTENT_TYPES] = _serialize(root)


def _author_id(parts, author):
    """Author id in commentAuthors.xml, creating the part if needed."""
    if AUTHORS_PART in parts:
        root = etree.fromstring(parts[AUTHORS_PART])
    else:
        root = etree.Element(_q(NS_P, "cmAuthorLst"), nsmap={"p": NS_P})
    existing = root.findall(_q(NS_P, "cmAuthor"))
    for element in existing:
        if element.get("name") == author:
            parts[AUTHORS_PART] = _serialize(root)
            return int(element.get("id", "1")), root
    new_id = max((int(e.get("id", "0")) for e in existing), default=0) + 1
    element = etree.SubElement(root, _q(NS_P, "cmAuthor"))
    element.set("id", str(new_id))
    element.set("name", author)
    element.set("initials", _initials(author))
    element.set("lastIdx", "1")
    element.set("clrIdx", str(new_id - 1))
    parts[AUTHORS_PART] = _serialize(root)
    return new_id, root


def _bump_last_idx(parts, authors_root, author_id, last_idx):
    for element in authors_root.findall(_q(NS_P, "cmAuthor")):
        if int(element.get("id", "0")) == author_id:
            element.set("lastIdx", str(last_idx))
    parts[AUTHORS_PART] = _serialize(authors_root)


def _max_existing_idx(parts, author_id):
    """Highest `idx` this author already used across every comment part.

    `p:cm/@idx` is per-author and deck-wide (it is what `cmAuthor/@lastIdx`
    counts), not per-slide: restarting at 1 on each slide gives two comments
    the same identity and leaves `lastIdx` describing only the last slide.
    """
    highest = 0
    for name, blob in parts.items():
        if not name.startswith(f"{COMMENTS_DIR}/"):
            continue
        try:
            root = etree.fromstring(blob)
        except etree.XMLSyntaxError:
            continue
        for cm in root.findall(_q(NS_P, "cm")):
            if int(cm.get("authorId", "0")) != author_id:
                continue
            try:
                highest = max(highest, int(cm.get("idx", "0")))
            except ValueError:
                continue
    return highest


def _comment_root(parts, part_name):
    if part_name in parts:
        return etree.fromstring(parts[part_name])
    return etree.Element(_q(NS_P, "cmLst"), nsmap={"a": NS_A, "p": NS_P})


def _append_comment(root, author_id, idx, text, when, slot):
    """Append one `p:cm`. `idx` is the deck-wide per-author id; `slot` is the
    0-based position of the balloon on *this* slide."""
    comment = etree.SubElement(root, _q(NS_P, "cm"))
    comment.set("authorId", str(author_id))
    comment.set("dt", when)
    comment.set("idx", str(idx))
    pos = etree.SubElement(comment, _q(NS_P, "pos"))
    pos.set("x", str(POS_X_EMU))
    pos.set("y", str(POS_Y_EMU + POS_STEP_EMU * slot))
    body = etree.SubElement(comment, _q(NS_P, "text"))
    body.text = text


def add_comments(src_pptx, dst_pptx, comments, author="QA"):
    """Copy `src_pptx` to `dst_pptx` with `comments` attached to their slides.

    comments: list of (slide_number, text) — 1-based slide numbers; entries
    pointing past the end of the deck are skipped.
    Returns the number of comments written.
    """
    src, dst = Path(src_pptx), Path(dst_pptx)
    order, parts = _read_parts(src)
    slides = _slide_partnames(parts)
    wanted = [(n, text) for n, text in comments if 1 <= n <= len(slides)]
    if not wanted:
        shutil.copyfile(src, dst)
        return 0

    author_id, authors_root = _author_id(parts, author)
    when = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S")
    per_slide = {}
    for n, text in wanted:
        per_slide.setdefault(n, []).append(text)

    written = 0
    idx = _max_existing_idx(parts, author_id)   # deck-wide, per author
    for n, texts in sorted(per_slide.items()):
        slide_part = slides[n - 1]
        number = _slide_index(slide_part) or n
        comment_part = f"{COMMENTS_DIR}/comment{number}.xml"
        root = _comment_root(parts, comment_part)
        slot = len(root.findall(_q(NS_P, "cm")))
        for text in texts:
            idx += 1
            _append_comment(root, author_id, idx, text, when, slot)
            slot += 1
            written += 1
        parts[comment_part] = _serialize(root)
        _add_override(parts, comment_part, CT_COMMENTS)
        slide_rels = (f"{Path(slide_part).parent.as_posix()}/_rels/"
                      f"{Path(slide_part).name}.rels")
        _add_rel(parts, slide_rels, REL_COMMENTS,
                 f"../comments/comment{number}.xml")
    _bump_last_idx(parts, authors_root, author_id, idx)

    _add_override(parts, AUTHORS_PART, CT_AUTHORS)
    _add_rel(parts, PRESENTATION_RELS, REL_AUTHORS, "commentAuthors.xml")

    _write_parts(dst, order, parts)
    return written


# ── reading comments (legacy + modern) ──────────────────────────────────────
@dataclass
class Comment:
    """One reviewer note. Replies are Comments too, with `parent_id` set.

    Equality is field-based (the dataclass default); hashing is by
    `thread_id` alone, and the asymmetry is deliberate. `anchor` is a dict,
    so no field-based hash is possible at all, and the thread id is the
    identity callers actually key on (revision_plan folds replies into their
    root thread that way). Two notes on one thread therefore land in the
    same bucket while staying unequal — which is what a hash is allowed to
    do; equal Comments always share a thread_id, so the hash/eq contract
    holds.
    """
    slide_pos: int              # 1-based position in the deck
    author: str
    text: str
    created: str                # ISO timestamp, or '' when the file omits it
    thread_id: str              # legacy: 'slide:authorId:idx'; modern: cm id
    parent_id: Optional[str]    # reply → the id of the thread it answers
    resolved: bool
    anchor: Optional[dict]      # {'shape_id': int} | {'x_emu':, 'y_emu':} | None
    kind: str                   # KIND_LEGACY | KIND_MODERN

    def __hash__(self):
        return hash(self.thread_id)


def _int(value):
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def _rels_name(part_name: str) -> str:
    """'ppt/slides/slide1.xml' → 'ppt/slides/_rels/slide1.xml.rels'."""
    directory = posixpath.dirname(part_name)
    return f"{directory}/_rels/{posixpath.basename(part_name)}.rels"


def _resolve_target(owner_part: str, target: str) -> str:
    """Package-absolute name of a relationship target."""
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(owner_part),
                                             target))


def _comment_rels(parts, slide_part: str):
    """(rel id, part name, kind) of every comment part a slide links to."""
    blob = parts.get(_rels_name(slide_part))
    if blob is None:
        return []
    found = []
    for rel in etree.fromstring(blob):
        if rel.get("TargetMode") == "External":
            continue
        kind = {REL_COMMENTS: KIND_LEGACY, RT_MODERN: KIND_MODERN}.get(
            rel.get("Type"))
        if kind is None:
            continue
        found.append((rel.get("Id"), _resolve_target(slide_part,
                                                     rel.get("Target", "")),
                      kind))
    return found


def _author_names(parts, part_name: str, list_tag: str, author_tag: str):
    blob = parts.get(part_name)
    if blob is None:
        return {}
    root = etree.fromstring(blob)
    if root.tag != list_tag:
        return {}
    return {element.get("id", ""): element.get("name", "")
            for element in root.findall(author_tag)}


def _dml_text(body) -> str:
    """Flatten a DrawingML `txBody` to text, one line per paragraph."""
    if body is None:
        return ""
    paragraphs = [
        "".join(node.text or "" for node in para.iter(_q(NS_A, "t")))
        for para in body.findall(_q(NS_A, "p"))]
    if not paragraphs:
        paragraphs = ["".join(node.text or ""
                              for node in body.iter(_q(NS_A, "t")))]
    return "\n".join(paragraphs).strip()


def _legacy_anchor(cm):
    pos = cm.find(_q(NS_P, "pos"))
    if pos is None:
        return None
    x, y = _int(pos.get("x")), _int(pos.get("y"))
    return None if x is None or y is None else {"x_emu": x, "y_emu": y}


def _modern_anchor(cm):
    """A shape mark wins over the balloon position — it survives re-layout."""
    mark = cm.find(f'{_q(NS_P188, "deMkLst")}/{_q(NS_PC, "spMk")}')
    if mark is not None:
        shape_id = _int(mark.get("id"))
        if shape_id is not None:
            return {"shape_id": shape_id}
    pos = cm.find(_q(NS_P188, "pos"))
    if pos is None:
        return None
    x, y = _int(pos.get("x")), _int(pos.get("y"))
    return None if x is None or y is None else {"x_emu": x, "y_emu": y}


def _replies(cm):
    return cm.findall(f'{_q(NS_P188, "replyLst")}/{_q(NS_P188, "reply")}')


def _read_legacy(root, slide_pos: int, authors):
    found = []
    for cm in root.findall(_q(NS_P, "cm")):
        author_id = cm.get("authorId", "")
        text_node = cm.find(_q(NS_P, "text"))
        found.append(Comment(
            slide_pos=slide_pos,
            author=authors.get(author_id, ""),
            text=(text_node.text or "").strip() if text_node is not None else "",
            created=cm.get("dt", ""),
            thread_id=f"{slide_pos}:{author_id}:{cm.get('idx', '')}",
            parent_id=None,
            resolved=False,             # legacy comments carry no such state
            anchor=_legacy_anchor(cm),
            kind=KIND_LEGACY))
    return found


def _read_modern(root, slide_pos: int, authors):
    found = []
    for cm in root.findall(_q(NS_P188, "cm")):
        thread_id = cm.get("id", "")
        resolved = cm.get("status", "") == STATUS_RESOLVED
        anchor = _modern_anchor(cm)
        found.append(Comment(
            slide_pos=slide_pos,
            author=authors.get(cm.get("authorId", ""), ""),
            text=_dml_text(cm.find(_q(NS_P188, "txBody"))),
            created=cm.get("created", ""),
            thread_id=thread_id,
            parent_id=None,
            resolved=resolved,
            anchor=anchor,
            kind=KIND_MODERN))
        for reply in _replies(cm):
            found.append(Comment(
                slide_pos=slide_pos,
                author=authors.get(reply.get("authorId", ""), ""),
                text=_dml_text(reply.find(_q(NS_P188, "txBody"))),
                created=reply.get("created", ""),
                thread_id=reply.get("id", ""),
                parent_id=thread_id,
                # a reply is settled once its thread is
                resolved=resolved or reply.get("status", "") == STATUS_RESOLVED,
                anchor=anchor,
                kind=KIND_MODERN))
    return found


def _parse_comment_part(part_name: str, blob: bytes):
    """Parse a comment part, reporting damage as a ValueError naming it.

    Same contract as the missing-presentation-part check: a damaged package
    is a reported error, never an lxml traceback out of `edit_deck comments
    list/plan/strip` or qa_summary.
    """
    try:
        return etree.fromstring(blob)
    except etree.XMLSyntaxError as exc:
        raise ValueError(
            f"{part_name}: malformed comment XML: {exc}") from exc


def read_comments(pptx_path) -> list[Comment]:
    """Every comment in the deck, in deck order, replies after their parent."""
    _order, parts = _read_parts(Path(pptx_path))
    slide_parts = _slide_partnames(parts)
    authors = {
        KIND_LEGACY: _author_names(parts, AUTHORS_PART, _q(NS_P, "cmAuthorLst"),
                                   _q(NS_P, "cmAuthor")),
        KIND_MODERN: _author_names(parts, MODERN_AUTHORS_PART,
                                   _q(NS_P188, "authorLst"),
                                   _q(NS_P188, "author")),
    }
    readers = {KIND_LEGACY: _read_legacy, KIND_MODERN: _read_modern}
    found = []
    for slide_pos, slide_part in enumerate(slide_parts, 1):
        for _rid, part_name, kind in _comment_rels(parts, slide_part):
            blob = parts.get(part_name)
            if blob is None:                    # dangling rel — nothing to read
                continue
            root = _parse_comment_part(part_name, blob)
            found.extend(readers[kind](root, slide_pos, authors[kind]))
    return found


# ── stripping comments ──────────────────────────────────────────────────────
def _remove_override(parts, part_name: str):
    root = etree.fromstring(parts[CONTENT_TYPES])
    name = "/" + part_name
    for override in root.findall(_q(NS_CT, "Override")):
        if override.get("PartName") == name:
            root.remove(override)
    parts[CONTENT_TYPES] = _serialize(root)


def _remove_rel(parts, rels_name: str, *, rel_id=None, rel_type=None):
    blob = parts.get(rels_name)
    if blob is None:
        return
    root = etree.fromstring(blob)
    for rel in list(root):
        if (rel_id is not None and rel.get("Id") == rel_id) or \
                (rel_type is not None and rel.get("Type") == rel_type):
            root.remove(rel)
    parts[rels_name] = _serialize(root)


def _drop_part(parts, part_name: str):
    """Delete a part together with its own rels and its content-type override."""
    parts.pop(part_name, None)
    parts.pop(_rels_name(part_name), None)
    if CONTENT_TYPES in parts:
        _remove_override(parts, part_name)


def _is_resolved(cm, kind: str) -> bool:
    return kind == KIND_MODERN and cm.get("status", "") == STATUS_RESOLVED


def _drop_comments(root, kind: str, resolved_only: bool) -> int:
    """Remove matching comments in place; returns how many notes went away."""
    tag = _q(NS_P, "cm") if kind == KIND_LEGACY else _q(NS_P188, "cm")
    removed = 0
    for cm in root.findall(tag):
        if resolved_only and not _is_resolved(cm, kind):
            continue
        removed += 1 + (len(_replies(cm)) if kind == KIND_MODERN else 0)
        root.remove(cm)
    return removed


def _cm_tag(kind: str) -> str:
    return _q(NS_P, "cm") if kind == KIND_LEGACY else _q(NS_P188, "cm")


def _comment_parts(parts):
    """Every comment part in the package, whether a slide links to it or not."""
    return sorted(name for name in parts
                  if name.startswith(f"{COMMENTS_DIR}/")
                  and not name.startswith(f"{COMMENTS_DIR}/_rels/"))


def _list_kind(root) -> Optional[str]:
    return {_q(NS_P, "cmLst"): KIND_LEGACY,
            _q(NS_P188, "cmLst"): KIND_MODERN}.get(root.tag)


def _has_comments(blob, kind: str) -> bool:
    """Unreadable parts count as referencing an author: dropping the author
    list under them would leave the surviving part pointing at nothing."""
    try:
        root = etree.fromstring(blob)
    except etree.XMLSyntaxError:
        return True
    return _list_kind(root) == kind and bool(root.findall(_cm_tag(kind)))


def _drop_unused_authors(parts):
    """An author list only earns its place while some comment still cites it."""
    for part_name, rel_type, kind in (
            (AUTHORS_PART, REL_AUTHORS, KIND_LEGACY),
            (MODERN_AUTHORS_PART, RT_MODERN_AUTHORS, KIND_MODERN)):
        if part_name not in parts:
            continue
        if any(_has_comments(parts[name], kind)
               for name in _comment_parts(parts)):
            continue
        _drop_part(parts, part_name)
        _remove_rel(parts, PRESENTATION_RELS, rel_type=rel_type)


def _remove_rels_to(parts, part_name: str):
    """Drop every relationship anywhere in the package aimed at `part_name`."""
    for rels_name in [n for n in parts
                      if "/_rels/" in n and n.endswith(".rels")]:
        owner = posixpath.join(
            posixpath.dirname(posixpath.dirname(rels_name)),
            posixpath.basename(rels_name)[:-len(".rels")])
        root = etree.fromstring(parts[rels_name])
        keep = [rel for rel in root
                if rel.get("TargetMode") == "External"
                or _resolve_target(owner, rel.get("Target", "")) != part_name]
        if len(keep) != len(root):
            for rel in list(root):
                if rel not in keep:
                    root.remove(rel)
            parts[rels_name] = _serialize(root)


def _sweep_orphan_comments(parts, seen, resolved_only: bool) -> int:
    """Clear comment parts no slide links to — they ship in the file all the
    same, so a whole-deck strip that left them behind would not be one."""
    removed = 0
    for part_name in _comment_parts(parts):
        if part_name in seen:
            continue
        try:
            root = etree.fromstring(parts[part_name])
        except etree.XMLSyntaxError:
            # Unreadable *and* orphaned: nothing links to it and nothing can
            # read it, so dropping it is the only thing left to do.
            _drop_part(parts, part_name)
            _remove_rels_to(parts, part_name)
            continue
        kind = _list_kind(root)
        if kind is None:
            continue
        removed += _drop_comments(root, kind, resolved_only)
        if root.findall(_cm_tag(kind)):
            parts[part_name] = _serialize(root)
        else:
            _drop_part(parts, part_name)
            _remove_rels_to(parts, part_name)
    return removed


def strip_comments(src_pptx, dst_pptx, *, slides=None,
                   resolved_only: bool = False) -> int:
    """Copy `src_pptx` to `dst_pptx` without its comments.

    slides: 1-based slide numbers to clear (default: the whole deck).
    resolved_only: keep every open thread — legacy comments have no resolved
    state, so they are all treated as open and survive.
    Returns the number of comments removed (replies counted individually).
    """
    src, dst = Path(src_pptx), Path(dst_pptx)
    order, parts = _read_parts(src)
    slide_parts = _slide_partnames(parts)
    if slides is None:
        wanted = set(range(1, len(slide_parts) + 1))
    else:
        wanted = {int(n) for n in slides}
        unknown = sorted(n for n in wanted if not 1 <= n <= len(slide_parts))
        if unknown:
            raise ValueError(
                f"strip_comments: slide(s) {unknown} outside a "
                f"{len(slide_parts)}-slide deck")

    removed, seen = 0, set()
    for slide_pos, slide_part in enumerate(slide_parts, 1):
        for rel_id, part_name, kind in _comment_rels(parts, slide_part):
            seen.add(part_name)
            if slide_pos not in wanted or part_name not in parts:
                continue
            root = _parse_comment_part(part_name, parts[part_name])
            removed += _drop_comments(root, kind, resolved_only)
            if root.findall(_cm_tag(kind)):
                parts[part_name] = _serialize(root)
            else:
                _drop_part(parts, part_name)
                _remove_rel(parts, _rels_name(slide_part), rel_id=rel_id)
    if slides is None:
        removed += _sweep_orphan_comments(parts, seen, resolved_only)
    _drop_unused_authors(parts)
    _write_parts(dst, order, parts)
    return removed
