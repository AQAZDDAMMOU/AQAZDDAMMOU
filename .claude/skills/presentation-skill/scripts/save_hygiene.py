#!/usr/bin/env python3
"""
save_hygiene.py — Keynote/macOS-safe package hygiene for generated .pptx files.

python-pptx saves leave four artefacts that PowerPoint tolerates but Keynote,
Google Slides and OOXML validators do not:

1. `<p:sldSz cx="12188952" cy="6858000" type="screen4x3"/>` — a 16:9 deck whose
   declared `type` says 4:3 and whose `cx` is 3048 EMU off the canonical grid
   (13.33in rounded instead of exactly 12192000).
2. A `notesMaster` relationship in `ppt/_rels/presentation.xml.rels` with no
   matching `<p:notesMasterIdLst>` in `ppt/presentation.xml` — Keynote drops
   speaker notes when the list is missing.
3. A vestigial `ppt/printerSettings/printerSettings1.bin` part plus its
   `<Default Extension="bin" .../>` content-type entry.
4. Stock core properties: `lastModifiedBy=Steve Canny`,
   `description=generated using python-pptx`, empty creator/title, 2013 dates.

`sanitize_presentation()` fixes all four on an in-memory `Presentation` right
before it is saved; `check_package_hygiene()` re-detects them in a saved file
(used by `qa_check.py --integrity`).

Usage:
    python3 scripts/save_hygiene.py deck.pptx [--output clean.pptx]
        [--author "Analyst Name"] [--title "Deck Title"]

With no --output the file is sanitized in place; --title defaults to the
filename stem. Omitting --author preserves a real existing dc:creator — only an
empty or stock ("Steve Canny") identity is replaced. Exit status 0 when the
saved package is clean, 1 when hygiene errors remain.
"""
import argparse
import datetime as dt
import sys
import zipfile
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn

# Canonical OOXML slide sizes (EMU) → the `type` attribute PowerPoint writes.
SIZE_TYPES = {
    (12192000, 6858000): "screen16x9",
    (9144000, 6858000): "screen4x3",
}
# 13.33in rounds to 12188952 EMU; snap anything this close to the canonical grid.
SNAP_TOL_EMU = 10000
SCREEN_TYPES = frozenset(SIZE_TYPES.values())

PPTX_DEFAULT_AUTHOR = "Steve Canny"
PPTX_DEFAULT_DESCRIPTION = "python-pptx"
# The stock created/modified stamps in python-pptx's bundled default.pptx.
PPTX_STOCK_CREATED = dt.datetime(2013, 1, 27, 9, 14, 16)
FALLBACK_AUTHOR = "presentation-skill"
INITIAL_REVISION = 1

HYGIENE_PREFIX = "hygiene: "
PRESENTATION_PART = "ppt/presentation.xml"
PRESENTATION_RELS = "ppt/_rels/presentation.xml.rels"
CORE_PART = "docProps/core.xml"
CONTENT_TYPES_PART = "[Content_Types].xml"
PRINTER_SETTINGS_DIR = "ppt/printerSettings/"

_NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    "dc": "http://purl.org/dc/elements/1.1/",
}
# A .pptx handed to the checker is untrusted input: no entity expansion, no
# network access, no unbounded tree.
_SAFE_PARSER = etree.XMLParser(resolve_entities=False, no_network=True,
                               huge_tree=False)


def safe_parser():
    """The hardened lxml parser used for untrusted OOXML.

    Shared so every module that parses parts out of a foreign .pptx
    (`client_safe.py`, …) gets the same entity/network/tree guarantees
    instead of lxml's permissive defaults.
    """
    return _SAFE_PARSER


# ── In-memory sanitizer ──────────────────────────────────────────────────────
def _nearest_known_size(cx, cy):
    """Return the canonical (cx, cy) key within SNAP_TOL_EMU, else None."""
    for (kcx, kcy) in SIZE_TYPES:
        if abs(cx - kcx) <= SNAP_TOL_EMU and abs(cy - kcy) <= SNAP_TOL_EMU:
            return (kcx, kcy)
    return None


def _snap_slide_size(prs):
    """Align `p:sldSz` cx/cy to the canonical grid and fix its `type`.

    Unknown sizes (custom templates) keep their dimensions; only a `type` that
    falsely claims a screen aspect is removed — a legitimate `letter`/`A4`
    type on a non-screen canvas is left alone."""
    sld_sz = prs.part._element.find(qn("p:sldSz"))
    if sld_sz is None:
        return None
    try:
        cx, cy = int(sld_sz.get("cx")), int(sld_sz.get("cy"))
    except (TypeError, ValueError):
        return None  # malformed template: leave it for check_package_hygiene
    declared = sld_sz.get("type")
    known = _nearest_known_size(cx, cy)
    if known is None:
        if declared in SCREEN_TYPES:
            del sld_sz.attrib["type"]
            return (f"removed p:sldSz type={declared!r} (canvas "
                    f"{cx}x{cy} EMU is not a known screen size)")
        return None
    want_type = SIZE_TYPES[known]
    if (cx, cy) == known and declared == want_type:
        return None
    sld_sz.set("cx", str(known[0]))
    sld_sz.set("cy", str(known[1]))
    sld_sz.set("type", want_type)
    return (f"snapped p:sldSz {cx}x{cy} type={declared!r} → "
            f"{known[0]}x{known[1]} type={want_type!r}")


def _inject_notes_master_id_list(prs):
    """Add `<p:notesMasterIdLst>` when notesMaster rels exist without one."""
    pres = prs.part._element
    if pres.find(qn("p:notesMasterIdLst")) is not None:
        return None
    r_ids = [rId for rId, rel in prs.part.rels.items()
             if rel.reltype == RT.NOTES_MASTER]
    if not r_ids:
        return None
    ids = "".join(f'<p:notesMasterId r:id="{rId}"/>' for rId in r_ids)
    lst = parse_xml(
        f'<p:notesMasterIdLst {nsdecls("p", "r")}>{ids}</p:notesMasterIdLst>')
    # Schema order: after p:sldMasterIdLst, before p:handoutMasterIdLst /
    # p:sldIdLst / p:sldSz.
    anchor = pres.find(qn("p:sldMasterIdLst"))
    if anchor is not None:
        anchor.addnext(lst)
    else:
        pres.insert(0, lst)
    return f"added p:notesMasterIdLst ({len(r_ids)} notes master(s))"


def _drop_printer_settings(prs):
    """Unlink `ppt/printerSettings/*.bin`; the `bin` content-type Default
    disappears with it because python-pptx rebuilds [Content_Types].xml from
    the reachable part graph."""
    part = prs.part
    dropped = []
    for rId, rel in list(part.rels.items()):
        if rel.reltype != RT.PRINTER_SETTINGS:
            continue
        part.drop_rel(rId)
        if rId in part.rels:  # drop_rel refuses when referenced elsewhere
            part.rels._rels.pop(rId, None)
        dropped.append(rId)
    if not dropped:
        return None
    return f"dropped printerSettings relationship(s): {', '.join(dropped)}"


def _resolve_identity(current, author, fallback):
    """Identity to write into a core-properties name field.

    An explicit `author` always wins. With no author, a real existing value is
    preserved — only an empty field or python-pptx's stock `Steve Canny` is
    replaced, so sanitizing someone else's deck never erases its provenance."""
    if author:
        return author
    if current and current != PPTX_DEFAULT_AUTHOR:
        return current
    return fallback


def _scrub_core_properties(prs, author, title):
    """Replace python-pptx's stock docProps/core.xml identity fields."""
    cp = prs.core_properties
    actions = []
    creator = _resolve_identity(cp.author, author, FALLBACK_AUTHOR)
    if cp.author != creator:
        actions.append(f"set dc:creator={creator!r}")
        cp.author = creator
    editor = _resolve_identity(cp.last_modified_by, author, creator)
    if cp.last_modified_by != editor:
        actions.append(f"set cp:lastModifiedBy={editor!r}")
        cp.last_modified_by = editor
    if title and cp.title != title:
        actions.append(f"set dc:title={title!r}")
        cp.title = title
    if cp.comments:
        actions.append("cleared dc:description "
                       f"({cp.comments!r})")
        cp.comments = ""
    if cp.revision != INITIAL_REVISION:
        actions.append(f"set cp:revision={INITIAL_REVISION}")
        cp.revision = INITIAL_REVISION
    now = dt.datetime.now(dt.timezone.utc)
    if cp.created is None or cp.created == PPTX_STOCK_CREATED:
        actions.append("refreshed dcterms:created")
        cp.created = now
    cp.modified = now  # a save is a modification; not reported as a fix
    return actions


def sanitize_presentation(prs, *, author="", title=""):
    """Fix Keynote/macOS compatibility defects on `prs` in place.

    `author` and `title` populate docProps/core.xml; an empty `author` keeps a
    real existing creator (see `_resolve_identity`) and an empty `title` leaves
    dc:title untouched. Returns the list of actions taken (empty = the package
    was already clean), suitable for a one-line build log."""
    actions = []
    for fix in (_snap_slide_size, _inject_notes_master_id_list,
                _drop_printer_settings):
        action = fix(prs)
        if action:
            actions.append(action)
    actions.extend(_scrub_core_properties(prs, author, title))
    return actions


# ── Saved-package checker ────────────────────────────────────────────────────
def _read(zf, name):
    try:
        return zf.read(name)
    except KeyError:
        return None


def _check_slide_size(pres_xml, errors):
    sld_sz = pres_xml.find("p:sldSz", _NS)
    if sld_sz is None:
        errors.append("p:sldSz missing from ppt/presentation.xml")
        return
    try:
        cx, cy = int(sld_sz.get("cx")), int(sld_sz.get("cy"))
    except (TypeError, ValueError):
        errors.append("p:sldSz has a non-integer cx/cy")
        return
    declared = sld_sz.get("type")
    known = _nearest_known_size(cx, cy)
    if known is None:
        if declared in SCREEN_TYPES:
            errors.append(
                f"p:sldSz type={declared!r} contradicts canvas {cx}x{cy} EMU")
        return
    if (cx, cy) != known:
        errors.append(
            f"p:sldSz cx={cx} cy={cy} is off the canonical grid "
            f"(expected {known[0]}x{known[1]} for {SIZE_TYPES[known]})")
    if declared != SIZE_TYPES[known]:
        errors.append(
            f"p:sldSz type={declared!r} should be "
            f"{SIZE_TYPES[known]!r} for a {known[0]}x{known[1]} canvas")


def _check_notes_master(pres_xml, rels_xml, errors):
    has_list = pres_xml.find("p:notesMasterIdLst", _NS) is not None
    has_rel = False
    if rels_xml is not None:
        has_rel = any(r.get("Type") == RT.NOTES_MASTER
                      for r in rels_xml.findall("pr:Relationship", _NS))
    if has_rel and not has_list:
        errors.append("notesMaster relationship without a "
                      "<p:notesMasterIdLst> in ppt/presentation.xml "
                      "(Keynote drops speaker notes)")
    elif has_list and not has_rel:
        errors.append("<p:notesMasterIdLst> without a notesMaster "
                      "relationship (dangling r:id)")


def _check_printer_settings(names, ct_xml, errors):
    parts = [n for n in names if n.startswith(PRINTER_SETTINGS_DIR)]
    if parts:
        errors.append(f"vestigial printerSettings part(s): {', '.join(parts)}")
    if ct_xml is not None:
        for default in ct_xml.findall("ct:Default", _NS):
            if default.get("Extension") == "bin":
                errors.append('[Content_Types].xml keeps a printerSettings '
                              '<Default Extension="bin"> entry')


def _text(core_xml, path):
    el = core_xml.find(path, _NS)
    return (el.text or "").strip() if el is not None else ""


def _check_core_properties(core_xml, errors):
    if core_xml is None:
        errors.append(f"{CORE_PART} is missing")
        return
    if _text(core_xml, "cp:lastModifiedBy") == PPTX_DEFAULT_AUTHOR:
        errors.append(f"cp:lastModifiedBy is the python-pptx default "
                      f"{PPTX_DEFAULT_AUTHOR!r}")
    if PPTX_DEFAULT_DESCRIPTION in _text(core_xml, "dc:description"):
        errors.append("dc:description still says it was "
                      f"'generated using {PPTX_DEFAULT_DESCRIPTION}'")
    if not _text(core_xml, "dc:title"):
        errors.append("empty dc:title (the deck name shown in Finder, "
                      "Keynote and SharePoint)")


def check_package_hygiene(pptx_path):
    """Return `hygiene: …` error strings for a saved .pptx (empty = clean)."""
    path = Path(pptx_path)
    if not path.is_file():
        return [f"{HYGIENE_PREFIX}{path} is not a file"]
    errors = []
    try:
        with zipfile.ZipFile(path) as zf:
            names = zf.namelist()
            pres_raw = _read(zf, PRESENTATION_PART)
            rels_raw = _read(zf, PRESENTATION_RELS)
            core_raw = _read(zf, CORE_PART)
            ct_raw = _read(zf, CONTENT_TYPES_PART)
    except (zipfile.BadZipFile, OSError) as e:
        return [f"{HYGIENE_PREFIX}cannot read {path.name}: {e}"]
    if pres_raw is None:
        return [f"{HYGIENE_PREFIX}{PRESENTATION_PART} is missing"]
    try:
        pres_xml = etree.fromstring(pres_raw, _SAFE_PARSER)
        rels_xml = etree.fromstring(rels_raw, _SAFE_PARSER) if rels_raw else None
        core_xml = etree.fromstring(core_raw, _SAFE_PARSER) if core_raw else None
        ct_xml = etree.fromstring(ct_raw, _SAFE_PARSER) if ct_raw else None
    except etree.XMLSyntaxError as e:
        return [f"{HYGIENE_PREFIX}malformed OOXML in {path.name}: {e}"]
    _check_slide_size(pres_xml, errors)
    _check_notes_master(pres_xml, rels_xml, errors)
    _check_printer_settings(names, ct_xml, errors)
    _check_core_properties(core_xml, errors)
    return [f"{HYGIENE_PREFIX}{e}" for e in errors]


# ── CLI ──────────────────────────────────────────────────────────────────────
def main(argv=None):
    parser = argparse.ArgumentParser(
        description="Sanitize a .pptx for Keynote/macOS and scrub the stock "
                    "python-pptx core properties.")
    parser.add_argument("pptx", help="Path to the .pptx to sanitize")
    parser.add_argument("--output", default=None,
                        help="Write here instead of sanitizing in place")
    parser.add_argument("--author", default="",
                        help="dc:creator / cp:lastModifiedBy value")
    parser.add_argument("--title", default="",
                        help="dc:title value (default: the filename stem)")
    args = parser.parse_args(argv)

    src = Path(args.pptx)
    if not src.is_file():
        print(f"[ERROR] no such file: {src}", file=sys.stderr)
        return 2
    dest = Path(args.output) if args.output else src
    prs = Presentation(str(src))
    # Mirror build_deck's fallback so an untitled deck can be made clean.
    actions = sanitize_presentation(prs, author=args.author,
                                    title=args.title or src.stem)
    prs.save(str(dest))
    for action in actions:
        print(f"  {action}")
    print(f"Saved: {dest} ({len(actions)} fix(es) applied)")
    remaining = check_package_hygiene(dest)
    for err in remaining:
        print(f"  [ERROR] {err}", file=sys.stderr)
    return 1 if remaining else 0


if __name__ == "__main__":
    sys.exit(main())
