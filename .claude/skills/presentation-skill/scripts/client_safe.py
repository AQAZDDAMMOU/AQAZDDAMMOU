#!/usr/bin/env python3
"""
client_safe.py — pre-send client-safety scrub for delivered decks.

A deck that is visually finished can still leak: reviewer comments, hidden
"backup" slides, speaker notes written for the internal team, the author's
name in core properties, shapes parked off-canvas, and the embedded workbook
behind every native chart. This module has three faces:

1. `scan()` / `scan_presentation()` — read-only audit of a .pptx. Errors for
   the things that must never ship (comment parts, hidden slides, denylisted
   terms); warnings for the things that usually should not (identity in core
   properties, off-slide/hidden shapes, unused layouts, embedded workbooks,
   the package thumbnail, speaker notes). Wired into `pptx_lint --client-safe`.
2. `clean_unpacked()` — destructive clean-up of an `edit_deck unpack`ed tree:
   drops comment parts, the thumbnail, unused layouts/masters, and optionally
   speaker notes and hidden slides; scrubs the core-properties identity while
   keeping dc:title. Wired into `edit_deck clean --client-safe`.
3. `sanitize_numbers()` — replaces every digit run of 3+ digits in slide text
   and chart caches with random digits (years 1900-2100 and 1-2 digit tokens
   survive), for turning a live client deck into a shareable sample. Wired
   into `edit_deck sanitize-numbers`.

`save_hygiene.py` scrubs the *stock python-pptx* core properties at build
time; this module complements it by reporting (and clearing) real identity
left in any deck, including ones this toolchain did not build.

Usage:
    python3 scripts/client_safe.py scan deck.pptx [--denylist terms.txt]
"""
import argparse
import random
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from lxml import etree
from pptx import Presentation
from pptx.oxml.ns import qn

import edit_deck
import save_hygiene
from edit_deck import NS, _iter_rels, _part_rels_path, _resolve_target, _xml, _write_xml
from qa_check import iter_shapes, safe_chart
from shape_groups import canvas_shapes

EMU_IN = 914400
PREFIX = "client-safety: "
MAX_LOCATIONS = 5           # denylist locations reported per term
MAX_NAMES = 3               # layout/master names listed in an aggregate warn
MIN_DIGIT_RUN = 3           # shorter digit runs are page numbers/quarters
YEAR_MIN, YEAR_MAX = 1900, 2100

COMMENT_DIR = "ppt/comments"
COMMENT_AUTHOR_PARTS = ("ppt/commentAuthors.xml", "ppt/authors.xml")
THUMBNAIL_DIR = "docProps"
THUMBNAIL_STEM = "thumbnail"
EMBEDDINGS_DIR = "/ppt/embeddings/"
WORKBOOK_EXTS = (".xlsx", ".xlsm", ".xlsb", ".xls")
DENYLIST_EXIT_CODE = 2
APP_PART = "/docProps/app.xml"
MASTER_DIR = "ppt/slideMasters"

EP_NS = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
CP_NS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
DC_NS = "http://purl.org/dc/elements/1.1/"
C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"

DIGIT_RUN_RX = re.compile(r"\d+")
HIDDEN_TRUE = ("1", "true")


# ── shared text helpers ─────────────────────────────────────────────────────
def _element_text(element):
    """Every <a:t> under an element, newline-joined."""
    return "\n".join(t.text or "" for t in element.iter(qn("a:t")))


def _notes_text(slide):
    if not slide.has_notes_slide:
        return ""
    frame = slide.notes_slide.notes_text_frame
    return frame.text if frame is not None else ""


def load_denylist(path):
    """Read a denylist file: one term per line, `#` comments and blanks out.

    Exits with status DENYLIST_EXIT_CODE on an unreadable file: both callers
    are CLIs, and a traceback is a poor answer to a mistyped path."""
    try:
        raw = Path(path).read_text(encoding="utf-8")
    except OSError as exc:
        print(f"[ERROR] cannot read denylist {path}: {exc}", file=sys.stderr)
        raise SystemExit(DENYLIST_EXIT_CODE) from None
    terms = []
    for line in raw.splitlines():
        term = line.split("#", 1)[0].strip()
        if term:
            terms.append(term)
    return terms


# ── scan ────────────────────────────────────────────────────────────────────
def _part_names(prs):
    return [str(p.partname) for p in prs.part.package.iter_parts()]


def _company(prs):
    """docProps/app.xml <Company> (python-pptx does not expose it)."""
    for part in prs.part.package.iter_parts():
        if str(part.partname) != APP_PART:
            continue
        try:
            root = etree.fromstring(part.blob, save_hygiene.safe_parser())
        except etree.XMLSyntaxError:
            return ""
        el = root.find(f"{{{EP_NS}}}Company")
        return (el.text or "").strip() if el is not None else ""
    return ""


def _check_comment_parts(names, issues):
    parts = [n for n in names
             if n.startswith(f"/{COMMENT_DIR}/")
             or n.lstrip("/") in COMMENT_AUTHOR_PARTS]
    if parts:
        issues["error"].append(
            f"{PREFIX}{len(parts)} reviewer-comment part(s) present "
            f"({', '.join(sorted(parts)[:MAX_NAMES])}) — comments travel with "
            "the file and open in the client's PowerPoint")


def _check_hidden_slides(prs, issues):
    for n, slide in enumerate(prs.slides, 1):
        if slide.element.get("show") == "0":
            issues["error"].append(
                f"Slide {n}: {PREFIX}hidden slide — hidden slides still ship "
                "in the file and un-hide with two clicks")


def _check_shapes(prs, issues):
    """Hidden shapes (any nesting depth) and off-slide shapes.

    Bounds are checked on `shape_groups.canvas_shapes`: a group child's
    left/top are group-child coordinates, mapped onto the canvas by the
    group's chOff/chExt transform, so comparing them against the slide box
    reports false hits — except through the identity transform a built
    exhibit group carries, where child coordinates *are* canvas coordinates
    and the children must be checked exactly as they were before grouping."""
    sw, sh = prs.slide_width / EMU_IN, prs.slide_height / EMU_IN
    for n, slide in enumerate(prs.slides, 1):
        for shape in iter_shapes(slide.shapes):
            nv = next(shape._element.iter(qn("p:cNvPr")), None)
            if nv is not None and nv.get("hidden") in HIDDEN_TRUE:
                issues["warn"].append(
                    f"Slide {n}: {PREFIX}hidden shape '{shape.name}' — "
                    "invisible in the deck but present in the XML")
        for shape in canvas_shapes(slide.shapes):
            if None in (shape.left, shape.top, shape.width, shape.height):
                continue
            left, top = shape.left / EMU_IN, shape.top / EMU_IN
            right, bottom = left + shape.width / EMU_IN, top + shape.height / EMU_IN
            if right <= 0 or bottom <= 0 or left >= sw or top >= sh:
                issues["warn"].append(
                    f"Slide {n}: {PREFIX}off-slide shape '{shape.name}' at "
                    f"({left:.2f},{top:.2f})in — parked outside the canvas "
                    "but still in the file")


def _check_core_props(prs, issues):
    cp = prs.core_properties
    # dc:title is deliberately absent: `save_hygiene` *requires* a non-empty
    # dc:title (an empty one is a package-hygiene error) and the clean-up keeps
    # it, so warning about it would be a warning nothing can ever clear.
    fields = [("dc:creator", cp.author),
              ("cp:lastModifiedBy", cp.last_modified_by),
              ("Company", _company(prs))]
    present = [f"{k}={v.strip()!r}" for k, v in fields if (v or "").strip()]
    if present:
        issues["warn"].append(
            f"{PREFIX}core properties carry identity: {'; '.join(present)} — "
            "visible in Finder/Explorer and SharePoint; "
            "`edit_deck.py clean <dir> --client-safe` clears them "
            "(dc:title is intentionally kept as the deck name)")


def _check_unused_layouts(prs, issues):
    used = {slide.slide_layout.part.partname for slide in prs.slides}
    if not used:
        return
    layouts = [lay for m in prs.slide_masters for lay in m.slide_layouts]
    unused = [lay for lay in layouts if lay.part.partname not in used]
    if unused:
        names = ", ".join(repr(lay.name) for lay in unused[:MAX_NAMES])
        more = f" +{len(unused) - MAX_NAMES} more" if len(unused) > MAX_NAMES else ""
        issues["warn"].append(
            f"{PREFIX}{len(unused)} of {len(layouts)} slide layouts are unused "
            f"({names}{more}) — template leftovers ship with the deck; "
            "`edit_deck.py clean <dir> --client-safe` drops them")
    dead_masters = [m for m in prs.slide_masters
                    if all(lay.part.partname not in used for lay in m.slide_layouts)]
    if dead_masters:
        issues["warn"].append(
            f"{PREFIX}{len(dead_masters)} slide master(s) have no used layout "
            "— remove them before sending")


def _check_package_extras(names, issues):
    embeds = [n for n in names if n.startswith(EMBEDDINGS_DIR)]
    books = [n for n in embeds if n.lower().endswith(WORKBOOK_EXTS)]
    others = [n for n in embeds if n not in books]
    if books:
        issues["warn"].append(
            f"{PREFIX}{len(books)} embedded workbook(s) under "
            "ppt/embeddings/ — native charts need them, but embedded chart "
            "data is shareable: the client can open every underlying number")
    if others:
        issues["warn"].append(
            f"{PREFIX}{len(others)} embedded object(s) under ppt/embeddings/ "
            f"({', '.join(sorted(others)[:MAX_NAMES])}) — OLE payloads travel "
            "with the deck; open them before sending to see what they hold")
    thumbs = [n for n in names
              if n.startswith(f"/{THUMBNAIL_DIR}/{THUMBNAIL_STEM}")]
    if thumbs:
        issues["warn"].append(
            f"{PREFIX}package thumbnail present ({thumbs[0]}) — it previews "
            "slide 1 in file browsers and can lag the current content")


def _check_notes(prs, issues):
    with_notes = [n for n, slide in enumerate(prs.slides, 1)
                  if _notes_text(slide).strip()]
    if with_notes:
        shown = with_notes[:MAX_LOCATIONS]
        where = ", ".join(str(n) for n in shown)
        more = (f" (+{len(with_notes) - len(shown)} more)"
                if len(with_notes) > len(shown) else "")
        issues["warn"].append(
            f"{PREFIX}{len(with_notes)} slide(s) have speaker notes "
            f"(slide(s) {where}{more}) — strip before sending if they are "
            "internal")


def _chart_locations(slide_no, slide):
    """[(label, text)] for the chart parts on one slide.

    Chart titles, axis titles and data labels live in `ppt/charts/chartN.xml`,
    not in the slide XML, so a denylisted client name in a chart title is
    invisible to `_element_text(slide.element)`.
    """
    out = []
    for k, shape in enumerate(
            (sh for sh in iter_shapes(slide.shapes)
             if getattr(sh, "has_chart", False)), 1):
        chart = safe_chart(shape)
        if chart is None:
            continue
        try:
            root = etree.fromstring(chart.part.blob,
                                    save_hygiene.safe_parser())
        except (etree.XMLSyntaxError, AttributeError, ValueError):
            continue
        text = "\n".join(t.text or "" for t in root.iter(qn("a:t")))
        if text.strip():
            out.append((f"slide {slide_no} chart {k}", text))
    return out


def _text_locations(prs):
    """[(location label, searchable text)] over every place text can hide."""
    out = []
    for n, slide in enumerate(prs.slides, 1):
        out.append((f"slide {n} text", _element_text(slide.element)))
        notes = _notes_text(slide)
        if notes:
            out.append((f"slide {n} notes", notes))
        out.extend(_chart_locations(n, slide))
    for master_no, master in enumerate(prs.slide_masters, 1):
        out.append((f"master {master_no}", _element_text(master.element)))
        for layout in master.slide_layouts:
            out.append((f'layout "{layout.name}"',
                        _element_text(layout.element)))
    cp = prs.core_properties
    for field, value in (("creator", cp.author),
                         ("lastModifiedBy", cp.last_modified_by),
                         ("title", cp.title),
                         ("company", _company(prs))):
        if value:
            out.append((f"core-props {field}", value))
    return out


def _check_denylist(prs, denylist, issues):
    terms = [t.strip() for t in (denylist or []) if t and t.strip()]
    if not terms:
        return
    locations = _text_locations(prs)
    for term in terms:
        needle = term.casefold()
        hits = [label for label, text in locations if needle in text.casefold()]
        if not hits:
            continue
        shown = hits[:MAX_LOCATIONS]
        more = f" (+{len(hits) - len(shown)} more)" if len(hits) > len(shown) else ""
        issues["error"].append(
            f"{PREFIX}denylisted term {term!r} appears in "
            f"{', '.join(shown)}{more}")


def scan_presentation(prs, denylist=None, issues=None):
    """Append client-safety findings for an open `Presentation` to `issues`."""
    issues = {"error": [], "warn": []} if issues is None else issues
    names = _part_names(prs)
    _check_comment_parts(names, issues)
    _check_hidden_slides(prs, issues)
    _check_denylist(prs, denylist, issues)
    _check_core_props(prs, issues)
    _check_shapes(prs, issues)
    _check_unused_layouts(prs, issues)
    _check_package_extras(names, issues)
    _check_notes(prs, issues)
    return issues


def scan(pptx, denylist=None, issues=None):
    """Client-safety audit of a .pptx path (or an already-open Presentation)."""
    prs = pptx if hasattr(pptx, "slides") else Presentation(str(pptx))
    return scan_presentation(prs, denylist, issues)


# ── clean_unpacked ──────────────────────────────────────────────────────────
def _owner_part(root, rels_path):
    """<root>/ppt/slides/_rels/slide1.xml.rels → 'ppt/slides/slide1.xml'."""
    rel = rels_path.relative_to(root)
    owner_dir = rel.parent.parent
    name = rel.name[:-len(".rels")]
    return (owner_dir / name).as_posix().lstrip(".").lstrip("/") if name else ""


def _remove_part_refs(root, parts):
    """Drop every Relationship and content-type Override for `parts`."""
    gone = set(parts)
    if not gone:
        return
    for rels_path in sorted(root.rglob("*.rels")):
        if not rels_path.is_file():
            continue
        owner = _owner_part(root, rels_path)
        tree = _xml(rels_path)
        changed = False
        for rel in list(_iter_rels(tree)):
            if rel.get("TargetMode") == "External":
                continue
            if _resolve_target(owner, rel.get("Target")) in gone:
                rel.getparent().remove(rel)
                changed = True
        if changed:
            _write_xml(tree, rels_path)
    ct_path = root / "[Content_Types].xml"
    if ct_path.exists():
        ct = _xml(ct_path)
        doomed = {f"/{p}" for p in gone}
        removed = False
        for override in ct.getroot().findall(f"{{{NS['ct']}}}Override"):
            if override.get("PartName") in doomed:
                ct.getroot().remove(override)
                removed = True
        if removed:
            _write_xml(ct, ct_path)


def _delete_parts(root, parts):
    for part in parts:
        (root / part).unlink(missing_ok=True)
        _part_rels_path(root, part).unlink(missing_ok=True)
    _remove_part_refs(root, parts)


def _comment_parts(root):
    parts = []
    cdir = root / COMMENT_DIR
    if cdir.is_dir():
        parts += [f"{COMMENT_DIR}/{f.name}" for f in sorted(cdir.iterdir())
                  if f.is_file()]
    parts += [p for p in COMMENT_AUTHOR_PARTS if (root / p).exists()]
    return parts


def _thumbnail_parts(root):
    tdir = root / THUMBNAIL_DIR
    if not tdir.is_dir():
        return []
    return [f"{THUMBNAIL_DIR}/{f.name}" for f in sorted(tdir.iterdir())
            if f.is_file() and f.stem.lower().startswith(THUMBNAIL_STEM)]


def _notes_parts(root):
    ndir = root / "ppt" / "notesSlides"
    if not ndir.is_dir():
        return []
    return [f"ppt/notesSlides/{f.name}" for f in sorted(ndir.glob("*.xml"))]


def _drop_hidden_slides(root):
    pres_p, rels_p, _ = edit_deck._presentation_parts(root)
    entries = edit_deck._slide_entries(_xml(pres_p), _xml(rels_p))
    hidden = [pos for pos, (_, _, fname) in enumerate(entries, 1)
              if _xml(root / "ppt" / "slides" / fname).getroot().get("show") == "0"]
    for pos in sorted(hidden, reverse=True):
        edit_deck.remove(root, pos)
    if not hidden:
        return []
    return [f"removed {len(hidden)} hidden slide(s) at position(s) "
            f"{', '.join(str(p) for p in hidden)}"]


def _used_layouts(root):
    """Package paths of the layouts referenced by at least one slide."""
    used = set()
    slides_dir = root / "ppt" / "slides"
    if not slides_dir.is_dir():
        return used
    for slide in sorted(slides_dir.glob("slide*.xml")):
        part = f"ppt/slides/{slide.name}"
        rels_path = _part_rels_path(root, part)
        if not rels_path.exists():
            continue
        for rel in _iter_rels(_xml(rels_path)):
            if rel.get("Type") == edit_deck.LAYOUT_RT:
                used.add(_resolve_target(part, rel.get("Target")))
    return used


def _strip_id_entries(xml_path, list_tag, item_tag, rids):
    """Remove <p:sldLayoutId>/<p:sldMasterId> entries whose r:id is in rids."""
    if not rids:
        return
    tree = _xml(xml_path)
    lst = tree.getroot().find(f"p:{list_tag}", NS)
    if lst is None:
        return
    for item in lst.findall(f"p:{item_tag}", NS):
        if item.get(f"{{{NS['r']}}}id") in rids:
            lst.remove(item)
    _write_xml(tree, xml_path)


def _drop_unused_layouts(root):
    """Delete layouts no slide uses and masters left with no used layout."""
    used = _used_layouts(root)
    if not used:
        return []
    doomed_layouts, doomed_masters = [], []
    master_dir = root / MASTER_DIR
    for master_xml in sorted(master_dir.glob("slideMaster*.xml")
                             if master_dir.is_dir() else []):
        mpart = f"{MASTER_DIR}/{master_xml.name}"
        rels_path = _part_rels_path(root, mpart)
        if not rels_path.exists():
            continue
        rid_to_layout = {
            rel.get("Id"): _resolve_target(mpart, rel.get("Target"))
            for rel in _iter_rels(_xml(rels_path))
            if rel.get("Type") == edit_deck.LAYOUT_RT
        }
        drop = {rid: part for rid, part in rid_to_layout.items()
                if part not in used}
        if not drop:
            continue
        doomed_layouts.extend(drop.values())
        if len(drop) == len(rid_to_layout):
            doomed_masters.append(mpart)
        else:
            _strip_id_entries(master_xml, "sldLayoutIdLst", "sldLayoutId",
                              set(drop))
    if doomed_masters:
        pres_p, pres_rels_p, _ = edit_deck._presentation_parts(root)
        rids = {rel.get("Id") for rel in _iter_rels(_xml(pres_rels_p))
                if rel.get("Type") == edit_deck.MASTER_RT
                and _resolve_target("ppt/presentation.xml",
                                    rel.get("Target")) in doomed_masters}
        _strip_id_entries(pres_p, "sldMasterIdLst", "sldMasterId", rids)
    _delete_parts(root, doomed_layouts + doomed_masters)
    actions = []
    if doomed_layouts:
        actions.append(f"removed {len(doomed_layouts)} unused slide layout(s)")
    if doomed_masters:
        actions.append(f"removed {len(doomed_masters)} unused slide master(s)")
    return actions


def _scrub_identity(xml_path, ns_uri, tags):
    """Blank the named child elements; returns the values cleared."""
    if not xml_path.exists():
        return []
    tree = _xml(xml_path)
    cleared = []
    for tag in tags:
        el = tree.getroot().find(f"{{{ns_uri}}}{tag}")
        if el is not None and (el.text or "").strip():
            cleared.append(f"{tag}={el.text.strip()!r}")
            el.text = None
    if cleared:
        _write_xml(tree, xml_path)
    return cleared


def _scrub_core_props(root):
    """Clear creator/lastModifiedBy/company; dc:title is deliberately kept."""
    core = root / "docProps" / "core.xml"
    cleared = _scrub_identity(core, DC_NS, ("creator",))
    cleared += _scrub_identity(core, CP_NS, ("lastModifiedBy",))
    cleared += _scrub_identity(root / "docProps" / "app.xml", EP_NS,
                               ("Company", "Manager"))
    if not cleared:
        return []
    return [f"scrubbed core properties: {', '.join(cleared)} (dc:title kept)"]


def clean_unpacked(src_dir, drop_notes=False, drop_hidden=False):
    """Strip client-unsafe parts from an `edit_deck unpack`ed directory.

    Always: reviewer comments, the package thumbnail, unused layouts/masters,
    the core-properties identity. Optionally speaker notes and hidden slides.
    Returns the list of actions taken (empty = nothing to clean)."""
    root = Path(src_dir)
    if not (root / "[Content_Types].xml").exists():
        raise ValueError(f"{root} is not an unpacked .pptx "
                         "(no [Content_Types].xml) — run `edit_deck.py unpack` first")
    actions = []
    if drop_hidden:
        actions += _drop_hidden_slides(root)
    if drop_notes:
        notes = _notes_parts(root)
        if notes:
            _delete_parts(root, notes)
            actions.append(f"removed {len(notes)} speaker-notes part(s)")
    for parts, label in ((_comment_parts(root), "reviewer-comment part"),
                         (_thumbnail_parts(root), "package thumbnail part")):
        if parts:
            _delete_parts(root, parts)
            actions.append(f"removed {len(parts)} {label}(s)")
    actions += _drop_unused_layouts(root)
    actions += _scrub_core_props(root)
    edit_deck._gc_unreferenced(root)
    return actions


# ── sanitize_numbers ────────────────────────────────────────────────────────
def _scramble(rng, length):
    return str(rng.randint(1, 9)) + "".join(
        str(rng.randint(0, 9)) for _ in range(length - 1))


def _sanitize_text(text, rng, counter):
    """Randomise digit runs of MIN_DIGIT_RUN+ digits; years survive."""
    def repl(match):
        run = match.group()
        if len(run) < MIN_DIGIT_RUN:
            return run
        if len(run) == 4 and YEAR_MIN <= int(run) <= YEAR_MAX:
            return run
        counter[0] += 1
        return _scramble(rng, len(run))
    return DIGIT_RUN_RX.sub(repl, text)


def _sanitize_nodes(nodes, rng, counter):
    """Rewrite the text of every node in place; True when anything changed."""
    changed = False
    for el in nodes:
        if not el.text:
            continue
        new = _sanitize_text(el.text, rng, counter)
        if new != el.text:
            el.text = new
            changed = True
    return changed


def _chart_value_nodes(tree):
    """<c:v> inside <c:numCache>/<c:numLit> — the cached chart numbers.

    Restricted to the numeric caches on purpose: <c:v> under a strCache holds
    category and series *names*, which the <a:t> pass does not cover either."""
    out = []
    for cache_tag in ("numCache", "numLit"):
        for cache in tree.getroot().iter(f"{{{C_NS}}}{cache_tag}"):
            out.extend(cache.iter(f"{{{C_NS}}}v"))
    return out


def _sanitize_part(xml_path, rng, counter, extra_nodes=None):
    """Sanitize every <a:t> in a part, plus whatever `extra_nodes` selects.

    Chart parts carry numbers in both places: the caches (`extra_nodes`) and
    <a:t> runs inside chart titles, axis titles and overridden data labels
    (c:tx/c:rich, c:dLbl/c:tx/c:rich) — miss the latter and real figures stay
    printed on the exhibit."""
    tree = _xml(xml_path)
    nodes = list(tree.getroot().iter(qn("a:t")))
    if extra_nodes is not None:
        nodes.extend(extra_nodes(tree))
    if _sanitize_nodes(nodes, rng, counter):
        _write_xml(tree, xml_path)


def sanitize_numbers(src_dir, seed=None):
    """Randomise the numbers in an unpacked deck; returns the count replaced.

    Slide text, chart text (titles, axis titles, data labels) and cached chart
    values (<c:v> under numCache/numLit) are rewritten so a real client deck
    can be shared as a sample. Embedded workbooks under ppt/embeddings/ are NOT
    touched — drop them (or the charts) if the underlying data must not
    travel."""
    root = Path(src_dir)
    rng = random.Random(seed)
    counter = [0]
    for slide in edit_deck._slide_xml_files(root):
        _sanitize_part(slide, rng, counter)
    charts_dir = root / "ppt" / "charts"
    for chart in sorted(charts_dir.glob("chart*.xml")):
        _sanitize_part(chart, rng, counter, _chart_value_nodes)
    return counter[0]


# ── CLI ─────────────────────────────────────────────────────────────────────
def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[1])
    sub = parser.add_subparsers(dest="cmd", required=True)
    p_scan = sub.add_parser("scan", help="Audit a .pptx for client-unsafe content")
    p_scan.add_argument("pptx")
    p_scan.add_argument("--denylist", default=None,
                        help="File of terms to flag: one per line, # comments")
    args = parser.parse_args(argv)

    denylist = load_denylist(args.denylist) if args.denylist else None
    issues = scan(args.pptx, denylist)
    for err in issues["error"]:
        print(f"  [ERROR] {err}")
    for warn in issues["warn"]:
        print(f"  [WARN]  {warn}")
    print(f"\n{Path(args.pptx).name}: {len(issues['error'])} error(s), "
          f"{len(issues['warn'])} warning(s)")
    return 1 if issues["error"] else 0


if __name__ == "__main__":
    sys.exit(main())
