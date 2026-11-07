#!/usr/bin/env python3
"""
slide_tags.py — per-slide identity written into the .pptx itself.

Every slide built by `build_deck.build()` gets one `ppt/tags/tagN.xml` part
(`p:tagLst` of `p:tag name="PK_*" val="..."`), linked from the slide through
`p:cSld/p:custDataLst/p:tags r:id="..."`. PowerPoint and Keynote round-trip
these tags untouched and never show them, so a delivered deck stays
self-describing even after a human reorders, duplicates or renames slides:

    PK_SLIDE      1-based position at build time (stale after a reorder)
    PK_KEY        stable outline key (slug of the heading, unique per deck)
    PK_LAYOUT     outline layout name
    PK_TITLE_SHA  sha1 of the heading text — matches a renamed-key slide
    PK_DATA_SHA   sha1 of the exhibit's data (file bytes or inline rows)
    PK_EVIDENCE   comma-joined evidence-ledger ids
    PK_FLAGS      comma-joined sorted subset of verbatim/repeat/ghost/template
    PK_VERSION    tag schema version

Consumers (patch_deck, refresh_deck, diff_deck, comment extraction) locate a
slide with `read_tags()` + `find_slide()` instead of trusting its position.

Flag sources: `verbatim` from `ctx["verbatim"]` (set by `build_deck.build()`
for `--verbatim` / `**Copy:** locked`), `repeat` from `p["_repeat"]`, `ghost`
from `p["_ghost"]` and `template` from `p["_templated"]` — the last two are
set by `build_deck.build()` on the slide dict it is about to tag, because
only the build loop knows which branch produced the slide.
"""
import hashlib
import re
from types import SimpleNamespace

from lxml import etree
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.opc.package import Part
from pptx.oxml.ns import qn

from outline_data import _resolve_data_file
import shape_groups

NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
RT_TAGS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags"
CT_TAGS = "application/vnd.openxmlformats-officedocument.presentationml.tags+xml"
TAGS_PARTNAME_TMPL = "/ppt/tags/tag%d.xml"

TAG_KEYS = ("PK_SLIDE", "PK_KEY", "PK_LAYOUT", "PK_TITLE_SHA", "PK_DATA_SHA",
            "PK_EVIDENCE", "PK_FLAGS", "PK_VERSION")
TAG_VERSION = "1"
KEY_MAX = 48
FALLBACK_KEY = "slide"
AUTO_PREFIX = "auto-"

# Top fraction of slide height a text shape must start in to be a title
# candidate (see `name_shapes`).
ROLE_TITLE_BAND = 0.22
# Generic "couldn't tell what this is" kind roles: always carry a -k suffix,
# even the first (and maybe only) one on a slide, since they are the most
# likely to gain a sibling on a later edit.
_ALWAYS_NUMBERED_ROLES = frozenset({"text", "shape"})

_XML_DECL = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
_WORD_RX = re.compile(r"[a-z0-9]+")
_EVIDENCE_SPLIT_RX = re.compile(r"[,\s]+")


# ── keys and hashes ─────────────────────────────────────────────────────────
def slug(text):
    """'Pricing bridge: FY24→FY26!' → 'pricing-bridge-fy24-fy26' (<= KEY_MAX).

    Returns '' when the text holds no ascii alphanumerics at all.
    """
    words = _WORD_RX.findall(str(text or "").lower())
    if not words:
        return ""
    return "-".join(words)[:KEY_MAX].rstrip("-")


def slide_key(p, seen):
    """Stable per-deck key for one outline slide; records it in `seen`.

    Auto-inserted slides (agenda, sources) key off their layout so an added
    authored slide never shifts them; repeat clones carry their row number.
    Collisions get a '-2', '-3', … suffix in build order.
    """
    if p.get("_auto"):
        base = AUTO_PREFIX + (slug(p.get("layout", "")) or FALLBACK_KEY)
    else:
        base = slug(p.get("heading") or p.get("title") or "") or FALLBACK_KEY
    repeat = p.get("_repeat")
    if repeat:
        base = f"{base}-r{repeat.get('row', '')}"
    key, n = base, 2
    while key in seen:
        key = f"{base}-{n}"
        n += 1
    seen.add(key)
    return key


def title_sha(text):
    """sha1 of a heading — the fallback match when a slide's key changed."""
    return hashlib.sha1(str(text or "").strip().encode("utf-8")).hexdigest()


def _inline_data_text(p):
    """Inline chart data as 'label\\tv1\\tv2\\n' lines ('' when there is none)."""
    lines = []
    for row in p.get("data") or []:
        try:
            label, value = row
        except (TypeError, ValueError):
            continue
        values = value if isinstance(value, (list, tuple)) else [value]
        lines.append("\t".join([str(label)] + [str(v) for v in values]))
    return "".join(line + "\n" for line in lines)


def data_sha(p, ctx):
    """sha1 of the slide's exhibit data — '' when the slide carries none.

    A `- Data-File:` slide hashes the resolved file's bytes (so an edited CSV
    is detectable without re-parsing it); otherwise the inline `- Data:` rows
    are serialised. The series names are appended either way, because
    renaming a series changes the exhibit without changing a number.
    """
    series = str(p.get("series") or "")
    spec = p.get("data_file")
    if spec:
        path = _resolve_data_file(spec, ctx)
        if path is None:
            return ""
        try:
            body = path.read_bytes()
        except OSError:
            body = b""
    else:
        text = _inline_data_text(p)
        if not text:
            return ""
        body = text.encode("utf-8")
    return hashlib.sha1(body + b"\n" + series.encode("utf-8")).hexdigest()


def _evidence_ids(p):
    """'E1, E2' → 'E1,E2' (whitespace-tolerant, order preserved)."""
    raw = str(p.get("evidence") or "").strip()
    return ",".join(i for i in _EVIDENCE_SPLIT_RX.split(raw) if i)


def _flags(p, ctx):
    flags = set()
    if (ctx or {}).get("verbatim"):
        flags.add("verbatim")
    if p.get("_repeat"):
        flags.add("repeat")
    if p.get("_ghost"):
        flags.add("ghost")
    if p.get("_templated"):
        flags.add("template")
    return ",".join(sorted(flags))


def tags_for(p, n, ctx, seen):
    """The full PK_* mapping for outline slide `p` at 1-based position `n`."""
    return {
        "PK_SLIDE": str(n),
        "PK_KEY": slide_key(p, seen),
        "PK_LAYOUT": str(p.get("layout") or ""),
        "PK_TITLE_SHA": title_sha(p.get("heading") or p.get("title") or ""),
        "PK_DATA_SHA": data_sha(p, ctx or {}),
        "PK_EVIDENCE": _evidence_ids(p),
        "PK_FLAGS": _flags(p, ctx),
        "PK_VERSION": TAG_VERSION,
    }


# ── write ───────────────────────────────────────────────────────────────────
def _tags_blob(mapping):
    root = etree.Element(f"{{{NS_P}}}tagLst", nsmap={"p": NS_P})
    for name, value in mapping.items():
        tag = etree.SubElement(root, f"{{{NS_P}}}tag")
        tag.set("name", str(name).upper())
        tag.set("val", str(value))
    return _XML_DECL + etree.tostring(root, xml_declaration=False,
                                      encoding="UTF-8")


def _csld(slide):
    return slide.shapes._spTree.getparent()


def _drop_existing(slide):
    """Remove any p:custDataLst/p:tags and orphan the part it pointed at."""
    csld = _csld(slide)
    for el in csld.findall(f"{{{NS_P}}}custDataLst"):
        for tags in el.findall(f"{{{NS_P}}}tags"):
            rid = tags.get(f"{{{NS_R}}}id")
            if rid and rid in slide.part.rels:
                slide.part.rels.pop(rid)
        csld.remove(el)


def write_tags(slide, mapping):
    """Attach `mapping` to `slide` as a fresh ppt/tags/tagN.xml part.

    Replaces any tags already on the slide. The `p:custDataLst` element goes
    immediately after `p:spTree`, which is where the schema requires it —
    PowerPoint refuses to open the file otherwise.
    """
    _drop_existing(slide)
    package = slide.part.package
    part = Part(package.next_partname(TAGS_PARTNAME_TMPL), CT_TAGS, package,
                _tags_blob(mapping))
    rid = slide.part.relate_to(part, RT_TAGS)
    csld = _csld(slide)
    custom = etree.SubElement(csld, f"{{{NS_P}}}custDataLst")
    tags = etree.SubElement(custom, f"{{{NS_P}}}tags")
    tags.set(f"{{{NS_R}}}id", rid)
    csld.insert(csld.index(slide.shapes._spTree) + 1, custom)


# ── read ────────────────────────────────────────────────────────────────────
def _tags_part(slide):
    """The slide's tags part via p:custDataLst, else any RT_TAGS rel."""
    for el in _csld(slide).findall(f"{{{NS_P}}}custDataLst"):
        for tags in el.findall(f"{{{NS_P}}}tags"):
            rid = tags.get(f"{{{NS_R}}}id")
            rel = slide.part.rels.get(rid) if rid else None
            if rel is not None and not rel.is_external:
                return rel.target_part
    for rel in slide.part.rels.values():
        if rel.reltype == RT_TAGS and not rel.is_external:
            return rel.target_part
    return None


def _parse_tags(part):
    try:
        root = etree.fromstring(part.blob)
    except (etree.XMLSyntaxError, AttributeError, ValueError):
        return None
    out = {}
    for tag in root.findall(f"{{{NS_P}}}tag"):
        name = tag.get("name")
        if name:
            out[name] = tag.get("val") or ""
    return out or None


def parse_tags_blob(blob):
    """`{name: val}` for a raw tags-part XML blob, or None when unparseable."""
    return _parse_tags(SimpleNamespace(blob=blob))


def rewrite_tag(blob, name, value):
    """Return `blob` with the `p:tag` named `name` set to `value`.

    Only that one tag changes (it is appended when absent); every other tag
    and the document's byte layout stay as the writer produced them. Raises
    ValueError on an unparseable blob — callers decide whether to skip.
    """
    try:
        root = etree.fromstring(blob)
    except (etree.XMLSyntaxError, ValueError) as e:
        raise ValueError(f"not a tags part: {e}") from e
    name = str(name).upper()
    for tag in root.findall(f"{{{NS_P}}}tag"):
        if tag.get("name") == name:
            tag.set("val", str(value))
            break
    else:
        tag = etree.SubElement(root, f"{{{NS_P}}}tag")
        tag.set("name", name)
        tag.set("val", str(value))
    return _XML_DECL + etree.tostring(root, xml_declaration=False,
                                      encoding="UTF-8")


def read_tags(pptx_path_or_prs):
    """[mapping | None] per slide, in deck order (index 0 = slide 1)."""
    prs = (pptx_path_or_prs if hasattr(pptx_path_or_prs, "slides")
           else Presentation(str(pptx_path_or_prs)))
    out = []
    for slide in prs.slides:
        part = _tags_part(slide)
        out.append(_parse_tags(part) if part is not None else None)
    return out


def find_slide(tags, key, title_sha=""):
    """1-based position of `key` in a `read_tags()` list, else None.

    Falls back to the title hash so a slide whose heading was rewritten in the
    outline (new key, same slide) is still found in the delivered deck.
    """
    for n, t in enumerate(tags, 1):
        if t and t.get("PK_KEY") == key:
            return n
    if title_sha:
        for n, t in enumerate(tags, 1):
            if t and t.get("PK_TITLE_SHA") == title_sha:
                return n
    return None


# ── semantic shape names ────────────────────────────────────────────────────
def _is_connector(shape):
    if shape.shape_type == MSO_SHAPE_TYPE.LINE:
        return True
    try:
        return shape._element.tag == qn("p:cxnSp")
    except AttributeError:
        return False


def _walk_shapes(shapes):
    """Every shape in z-order, a group followed by its own children."""
    for shape in shapes:
        yield shape
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from _walk_shapes(shape.shapes)


def _shape_kind(shape):
    """The role a shape earns on its own merits (no chrome, isn't the title)."""
    if getattr(shape, "has_chart", False):
        return "chart"
    if getattr(shape, "has_table", False):
        return "table"
    if shape.shape_type in (MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.LINKED_PICTURE):
        return "picture"
    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        return "exhibit" if shape_groups.is_exhibit_group(shape) else "group"
    if _is_connector(shape):
        return "connector"
    if getattr(shape, "has_text_frame", False) and shape.text_frame.text.strip():
        return "text"
    return "shape"


def _max_run_size(shape):
    """Largest explicit run font size (Pt) on the shape, 0 if none is set."""
    best = 0
    for para in shape.text_frame.paragraphs:
        for run in para.runs:
            if run.font.size is not None:
                best = max(best, run.font.size)
    return best


def _is_decorative_glyph(text):
    """True for a lone punctuation/quote-mark glyph (“ ” • —) that isn't a
    real headline, even when it's the largest, top-most run on the slide
    (e.g. quote-evidence's giant curly-quote mark)."""
    return not any(ch.isalnum() for ch in text) or len(text) <= 2


def _pick_title(candidates, slide_height):
    """The slide's title shape: largest top-band run size, ties -> topmost.

    A slide with no real headline in the title band (only decorative glyphs)
    correctly gets no title at all — never promote a decorative shape."""
    band = ROLE_TITLE_BAND * slide_height
    in_band = []
    for shape in candidates:
        if not getattr(shape, "has_text_frame", False):
            continue
        text = shape.text_frame.text.strip()
        if not text or _is_decorative_glyph(text):
            continue
        top = shape.top if shape.top is not None else 0
        if top >= band:
            continue
        in_band.append((_max_run_size(shape), top, shape))
    if not in_band:
        return None
    in_band.sort(key=lambda t: (-t[0], t[1]))
    return in_band[0][2]


def name_shapes(slide, n, *, chrome=None, templated=False):
    """Assign every built shape on `slide` an `s<NN>-<role>[-<k>]` name.

    `chrome` maps `shape_id -> role` for shapes the build-time chrome helpers
    added (footer, pagenum, tracker, kicker, sticker, source, stamp,
    exhibit-badge) — those roles win outright. Otherwise the single largest
    top-band text shape becomes `title` (ties broken by topmost); everything
    else is classified by kind (chart/table/picture/group/connector, else
    text/shape). Template-mode placeholders keep their own names and are
    skipped entirely when `templated=True`. `-k` numbering (from 1, in
    z-order) is added whenever a role has more than one shape, and always for
    the generic 'text'/'shape' kinds even when there's only one.

    A shape-built exhibit is wrapped in a `p:grpSp` (see `shape_groups`), so
    the walk descends into groups: the group itself becomes `exhibit` (or the
    generic `group` for a foreign one) and its children are named exactly as
    they would have been standing loose on the slide — grouping must not
    rename anything, or every QA message that quotes a shape name would move.

    Returns `{shape_id: name}` for every shape it renamed.
    """
    chrome = chrome or {}
    prefix = f"s{n:02d}"
    # document order == z-order, groups included, children after their group
    shapes = list(_walk_shapes(slide.shapes))
    renamable = [sh for sh in shapes
                 if not (templated and sh.is_placeholder)]

    slide_height = slide.part.package.presentation_part.presentation.slide_height
    title_candidates = [sh for sh in renamable if sh.shape_id not in chrome]
    title_shape = _pick_title(title_candidates, slide_height)
    title_id = title_shape.shape_id if title_shape is not None else None

    roles = {}
    for sh in renamable:
        if sh.shape_id in chrome:
            roles[sh.shape_id] = chrome[sh.shape_id]
        elif sh.shape_id == title_id:
            roles[sh.shape_id] = "title"
        else:
            roles[sh.shape_id] = _shape_kind(sh)

    role_counts = {}
    for role in roles.values():
        role_counts[role] = role_counts.get(role, 0) + 1

    result, seen = {}, {}
    for sh in renamable:
        role = roles[sh.shape_id]
        if role != "title" and (role_counts[role] > 1
                                or role in _ALWAYS_NUMBERED_ROLES):
            seen[role] = seen.get(role, 0) + 1
            name = f"{prefix}-{role}-{seen[role]}"
        else:
            name = f"{prefix}-{role}"
        sh.name = name
        result[sh.shape_id] = name
    return result
