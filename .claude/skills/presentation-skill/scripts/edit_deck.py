#!/usr/bin/env python3
"""
edit_deck.py — Surgical editing of existing .pptx files via OOXML unpack/pack.

Subcommands:
    unpack <deck.pptx> <dir>          Extract + pretty-print XML for editing
    pack <dir> <out.pptx>             Validate XML and re-zip into a .pptx
    list <dir>                        Show slide order, files, and titles
    duplicate <dir> <N>               Duplicate slide at deck position N (see list)
    remove <dir> <N>                  Remove slide at deck position N
    reorder <dir> <order>             Reorder slides: e.g. 3,1,2,4 (positions from list)
    clean <dir> [--client-safe]       Remove unreferenced media from ppt/media/;
                                      --client-safe also strips comments, the
                                      thumbnail, unused layouts/masters and the
                                      core-props identity (+ --drop-notes,
                                      --drop-hidden)
    sanitize-numbers <dir> [--seed N] Randomise 3+ digit numbers in slide text
                                      and chart caches (years survive)
    extract <deck.pptx> <sel> --output sub.pptx
                                      Split: new deck from positions "1,3-5"
    append <dst.pptx> <src.pptx> [--slides 2-4] --output merged.pptx
                                      Merge: copy src slides after dst's last

Workflow: unpack → (structural ops: duplicate/remove) → edit slide XML text →
pack → QA with qa_check.py + render_slides.py. See references/editing.md.
"""
import argparse
import posixpath
import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

import slide_tags

NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}
SLIDE_RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
NOTES_RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"
LAYOUT_RT = ("http://schemas.openxmlformats.org/officeDocument/2006/"
             "relationships/slideLayout")
MASTER_RT = ("http://schemas.openxmlformats.org/officeDocument/2006/"
             "relationships/slideMaster")
# `edit_deck duplicate` re-keys the clone's PK_KEY as <key>-copy, -copy2, ...
COPY_KEY_SUFFIX = "-copy"


def _xml(path):
    return etree.parse(str(path))


def _write_xml(tree, path):
    tree.write(str(path), xml_declaration=True, encoding="UTF-8", standalone=True)


# ── unpack / pack ────────────────────────────────────────────────────────────
def unpack(pptx, out_dir):
    out = Path(out_dir)
    if out.exists():
        shutil.rmtree(out)
    with zipfile.ZipFile(pptx) as zf:
        zf.extractall(out)
    # Pretty-print XML so the Edit tool can target lines. lxml only indents
    # element-only content, so runs/text nodes are untouched.
    count = 0
    for f in out.rglob("*"):
        if f.suffix.lower() in (".xml", ".rels") and f.is_file():
            try:
                tree = etree.parse(str(f))
                f.write_bytes(etree.tostring(tree, pretty_print=True,
                                             xml_declaration=True,
                                             encoding="UTF-8", standalone=True))
                count += 1
            except etree.XMLSyntaxError:
                pass
    print(f"Unpacked {pptx} → {out}/ ({count} XML files formatted)")
    print("Slides:", ", ".join(s.name for s in sorted(
        (out / "ppt" / "slides").glob("slide*.xml"),
        key=lambda x: int(re.search(r"\d+", x.name).group()))))


def pack(src_dir, out_pptx):
    src = Path(src_dir)
    bad = []
    for f in src.rglob("*"):
        if f.suffix.lower() in (".xml", ".rels") and f.is_file():
            try:
                etree.parse(str(f))
            except etree.XMLSyntaxError as e:
                bad.append(f"{f.relative_to(src)}: {e}")
    if bad:
        print("XML validation failed — fix before packing:", file=sys.stderr)
        for b in bad:
            print(f"  [ERROR] {b}", file=sys.stderr)
        return False

    files = sorted(p for p in src.rglob("*") if p.is_file())
    # [Content_Types].xml must be present; put it first in the archive
    files.sort(key=lambda p: p.name != "[Content_Types].xml")
    with zipfile.ZipFile(out_pptx, "w", zipfile.ZIP_DEFLATED) as zf:
        for f in files:
            zf.write(f, f.relative_to(src).as_posix())
    print(f"Packed {src}/ → {out_pptx} ({len(files)} files)")
    return True


# ── slide structure helpers ──────────────────────────────────────────────────
def _presentation_parts(src):
    src = Path(src)
    return (src / "ppt" / "presentation.xml",
            src / "ppt" / "_rels" / "presentation.xml.rels",
            src / "[Content_Types].xml")


def _slide_entries(pres_tree, rels_tree):
    """Ordered [(sldId el, rId, slide filename)] from presentation.xml."""
    rid_to_target = {
        r.get("Id"): r.get("Target")
        for r in rels_tree.getroot().findall(f"{{{NS['rel']}}}Relationship")
        if r.get("Type") == SLIDE_RT
    }
    entries = []
    for sld in pres_tree.getroot().findall("p:sldIdLst/p:sldId", NS):
        rid = sld.get(f"{{{NS['r']}}}id")
        entries.append((sld, rid, Path(rid_to_target[rid]).name))
    return entries


def _slide_file_num(fname):
    """'slide3.xml' -> 3."""
    m = re.search(r"\d+", fname)
    return int(m.group()) if m else None


def _resolve_position(src_dir, position):
    """Map 1-based deck position (as shown by `list`) to slide file number."""
    pres_p, rels_p, _ = _presentation_parts(src_dir)
    entries = _slide_entries(_xml(pres_p), _xml(rels_p))
    if position < 1 or position > len(entries):
        print(f"ERROR: position {position} out of range (1-{len(entries)})",
              file=sys.stderr)
        return None
    fname = entries[position - 1][2]
    num = _slide_file_num(fname)
    if num is None:
        print(f"ERROR: cannot parse slide number from {fname}", file=sys.stderr)
    return num


def list_slides(src_dir):
    pres_p, rels_p, _ = _presentation_parts(src_dir)
    pres, rels = _xml(pres_p), _xml(rels_p)
    for pos, (_, rid, fname) in enumerate(_slide_entries(pres, rels), 1):
        slide_path = Path(src_dir) / "ppt" / "slides" / fname
        texts = etree.parse(str(slide_path)).getroot().findall(".//a:t", NS)
        first = next((t.text for t in texts if t.text and t.text.strip()), "")
        print(f"  {pos:2d}. {fname:<14} ({rid})  {first[:60]}")


def _all_tag_keys(src):
    """Every PK_KEY across the deck's ppt/tags parts (unparseable ones skipped)."""
    keys = set()
    for part in (Path(src) / "ppt" / "tags").glob("tag*.xml"):
        tags = slide_tags.parse_tags_blob(part.read_bytes()) or {}
        if tags.get("PK_KEY"):
            keys.add(tags["PK_KEY"])
    return keys


def _unique_copy_key(key, taken):
    """`<key>-copy`, then `-copy2`, `-copy3`, ... — first one not in `taken`."""
    n = 1
    while True:
        candidate = f"{key}{COPY_KEY_SUFFIX}" + (str(n) if n > 1 else "")
        if candidate not in taken:
            return candidate
        n += 1


def _retag_clone(src, clone_part, rels_root):
    """Give a duplicated slide its own tags part with a unique PK_KEY.

    `rels_root` is the clone's (not yet written) .rels root. When it carries
    a tags relationship the target part is copied to ppt/tags/tagM.xml
    (M = max existing + 1), the rel is re-pointed at the copy and only the
    copy's PK_KEY is rewritten (`<key>-copy`, `-copy2`, ...). Every other
    PK_* value is left as-is. Returns (source_partname, new_partname) for
    the [Content_Types].xml registration, or None when there is nothing to
    re-tag (foreign deck) — the clone is then byte-identical to before.
    """
    rel = next((r for r in rels_root.findall(f"{{{NS['rel']}}}Relationship")
                if r.get("Type") == slide_tags.RT_TAGS
                and r.get("TargetMode") != "External"), None)
    if rel is None:
        return None
    src = Path(src)
    src_partname = _resolve_target(clone_part, rel.get("Target"))
    src_path = src / src_partname
    if not src_path.exists():
        return None
    tags_dir = src / "ppt" / "tags"
    nums = [int(m.group(1)) for f in tags_dir.glob("tag*.xml")
            if (m := re.fullmatch(r"tag(\d+)\.xml", f.name))]
    new_partname = f"ppt/tags/tag{max(nums, default=0) + 1}.xml"
    blob = src_path.read_bytes()
    tags = slide_tags.parse_tags_blob(blob) or {}
    if tags.get("PK_KEY"):
        new_key = _unique_copy_key(tags["PK_KEY"], _all_tag_keys(src))
        blob = slide_tags.rewrite_tag(blob, "PK_KEY", new_key)
    (src / new_partname).write_bytes(blob)
    rel.set("Target", _relative_target(clone_part, new_partname))
    return src_partname, new_partname


def _register_override_like(ct_root, src_partname, new_partname):
    """Add a [Content_Types].xml Override for `new_partname` mirroring the
    one `src_partname` has (no-op when the source is covered by a Default)."""
    src_override = next(
        (o for o in ct_root.findall(f"{{{NS['ct']}}}Override")
         if o.get("PartName") == "/" + src_partname), None)
    if src_override is None:
        return
    override = etree.SubElement(ct_root, f"{{{NS['ct']}}}Override")
    override.set("PartName", "/" + new_partname)
    override.set("ContentType", src_override.get("ContentType"))


def duplicate(src_dir, position):
    slide_num = _resolve_position(src_dir, position)
    if slide_num is None:
        return False
    src = Path(src_dir)
    slides_dir = src / "ppt" / "slides"
    src_xml = slides_dir / f"slide{slide_num}.xml"
    if not src_xml.exists():
        print(f"ERROR: {src_xml} not found", file=sys.stderr)
        return False

    nums = [int(re.search(r"\d+", f.name).group())
            for f in slides_dir.glob("slide*.xml")]
    new_num = max(nums) + 1
    new_xml = slides_dir / f"slide{new_num}.xml"
    shutil.copy(src_xml, new_xml)

    # slide rels: copy but drop the notesSlide relationship (notes aren't cloned)
    src_rels = slides_dir / "_rels" / f"slide{slide_num}.xml.rels"
    new_tags_part = None
    if src_rels.exists():
        rels = _xml(src_rels)
        root = rels.getroot()
        for r in root.findall(f"{{{NS['rel']}}}Relationship"):
            if r.get("Type") == NOTES_RT:
                root.remove(r)
        new_tags_part = _retag_clone(src, f"ppt/slides/slide{new_num}.xml", root)
        _write_xml(rels, slides_dir / "_rels" / f"slide{new_num}.xml.rels")

    # [Content_Types].xml override(s)
    pres_p, rels_p, ct_p = _presentation_parts(src)
    ct = _xml(ct_p)
    override = etree.SubElement(ct.getroot(), f"{{{NS['ct']}}}Override")
    override.set("PartName", f"/ppt/slides/slide{new_num}.xml")
    override.set("ContentType",
                 "application/vnd.openxmlformats-officedocument"
                 ".presentationml.slide+xml")
    if new_tags_part is not None:
        _register_override_like(ct.getroot(), *new_tags_part)
    _write_xml(ct, ct_p)

    # presentation rels + sldIdLst (insert right after the source slide)
    pres, rels = _xml(pres_p), _xml(rels_p)
    rel_root = rels.getroot()
    max_rid = max(int(r.get("Id")[3:]) for r in
                  rel_root.findall(f"{{{NS['rel']}}}Relationship"))
    new_rid = f"rId{max_rid + 1}"
    rel = etree.SubElement(rel_root, f"{{{NS['rel']}}}Relationship")
    rel.set("Id", new_rid)
    rel.set("Type", SLIDE_RT)
    rel.set("Target", f"slides/slide{new_num}.xml")
    _write_xml(rels, rels_p)

    entries = _slide_entries(pres, _xml(rels_p))
    max_sldid = max(int(e[0].get("id")) for e in entries)
    source_el = next(e[0] for e in entries if e[2] == f"slide{slide_num}.xml")
    new_sld = etree.SubElement(source_el.getparent(), f"{{{NS['p']}}}sldId")
    new_sld.set("id", str(max_sldid + 1))
    new_sld.set(f"{{{NS['r']}}}id", new_rid)
    source_el.addnext(new_sld)
    _write_xml(pres, pres_p)

    print(f"Duplicated position {position} (slide{slide_num}.xml) → slide{new_num}.xml "
          f"(inserted after position {position})")
    return True


def remove(src_dir, position):
    slide_num = _resolve_position(src_dir, position)
    if slide_num is None:
        return False
    src = Path(src_dir)
    pres_p, rels_p, ct_p = _presentation_parts(src)
    pres, rels, ct = _xml(pres_p), _xml(rels_p), _xml(ct_p)

    target_fname = f"slide{slide_num}.xml"
    entries = _slide_entries(pres, rels)
    entry = next((e for e in entries if e[2] == target_fname), None)
    if entry is None:
        print(f"ERROR: {target_fname} is not in the slide list", file=sys.stderr)
        return False
    sld_el, rid, _ = entry

    sld_el.getparent().remove(sld_el)
    _write_xml(pres, pres_p)
    rel_root = rels.getroot()
    for r in rel_root.findall(f"{{{NS['rel']}}}Relationship"):
        if r.get("Id") == rid:
            rel_root.remove(r)
    _write_xml(rels, rels_p)
    ct_root = ct.getroot()
    for o in ct_root.findall(f"{{{NS['ct']}}}Override"):
        if o.get("PartName") == f"/ppt/slides/{target_fname}":
            ct_root.remove(o)
    _write_xml(ct, ct_p)

    for f in (src / "ppt" / "slides" / target_fname,
              src / "ppt" / "slides" / "_rels" / f"{target_fname}.rels"):
        f.unlink(missing_ok=True)
    print(f"Removed position {position} ({target_fname}; media orphans left in ppt/media/)")
    return True


def reorder(src_dir, order_str):
    """Reorder slides by 1-based positions from `list` output."""
    src = Path(src_dir)
    pres_p, rels_p, _ = _presentation_parts(src)
    pres, rels = _xml(pres_p), _xml(rels_p)
    entries = _slide_entries(pres, rels)
    n = len(entries)
    try:
        order = [int(x.strip()) for x in order_str.split(",")]
    except ValueError:
        print("ERROR: order must be comma-separated positions, e.g. 3,1,2,4",
              file=sys.stderr)
        return False
    if sorted(order) != list(range(1, n + 1)):
        print(f"ERROR: order must be a permutation of 1-{n}, got {order}",
              file=sys.stderr)
        return False

    lst = pres.getroot().find("p:sldIdLst", NS)
    children = [entries[i - 1][0] for i in order]
    for child in list(lst):
        lst.remove(child)
    for child in children:
        lst.append(child)
    _write_xml(pres, pres_p)
    print(f"Reordered slides to positions: {order_str}")
    return True


def clean_orphans(src_dir):
    """Delete ppt/media files not referenced by any .rels."""
    src = Path(src_dir)
    media = src / "ppt" / "media"
    if not media.is_dir():
        print("No ppt/media/ directory — nothing to clean.")
        return True
    referenced = set()
    for rels in src.rglob("*.rels"):
        text = rels.read_text(encoding="utf-8", errors="ignore")
        for m in re.finditer(r'(?:Target="[^"]*/)?media/([^"]+)"', text):
            referenced.add(m.group(1))
    removed = 0
    for f in media.iterdir():
        if f.is_file() and f.name not in referenced:
            f.unlink()
            removed += 1
            print(f"  Removed orphan: {f.name}")
    print(f"Clean complete: {removed} orphan file(s) removed.")
    return True


# ── extract / append (split & merge) ────────────────────────────────────────
def _parse_selection(spec, n):
    """'1,3-5' → [1, 3, 4, 5]. Ordered, deduped, validated against 1..n.
    Raises ValueError on empty/malformed/out-of-range/reversed input."""
    out, seen = [], set()
    for tok in (spec or "").split(","):
        tok = tok.strip()
        if not tok:
            continue
        m = re.fullmatch(r"(\d+)(?:-(\d+))?", tok)
        if not m:
            raise ValueError(f"bad selection token {tok!r} "
                             "(use positions like 2, 5-7)")
        a, b = int(m.group(1)), int(m.group(2) or m.group(1))
        if a > b:
            raise ValueError(f"reversed range {tok!r}")
        for p in range(a, b + 1):
            if p < 1 or p > n:
                raise ValueError(f"position {p} out of range (1-{n})")
            if p not in seen:
                seen.add(p)
                out.append(p)
    if not out:
        raise ValueError("empty selection")
    return out


def _part_rels_path(root, part):
    """ppt/slides/slide3.xml → <root>/ppt/slides/_rels/slide3.xml.rels."""
    p = Path(part)
    return Path(root) / p.parent / "_rels" / (p.name + ".rels")


def _resolve_target(base_part, target):
    """Relationship Target (relative to base part's dir) → package path."""
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(
        posixpath.join(posixpath.dirname(base_part), target))


def _relative_target(from_part, to_part):
    return posixpath.relpath(to_part, posixpath.dirname(from_part))


def _iter_rels(tree):
    return tree.getroot().findall(f"{{{NS['rel']}}}Relationship")


def _gc_unreferenced(root):
    """Delete parts unreachable from the package roots; prune CT overrides.

    Used by extract: removing slides orphans their notesSlides/charts/
    embeddings/media — dangling parts invite PowerPoint repair prompts.
    """
    root = Path(root)
    reachable = set()

    def walk(part):
        if part in reachable:
            return
        reachable.add(part)
        rels = _part_rels_path(root, part)
        if rels.exists():
            for rel in _iter_rels(_xml(rels)):
                if rel.get("TargetMode") == "External":
                    continue
                walk(_resolve_target(part, rel.get("Target")))

    for rel in _iter_rels(_xml(root / "_rels" / ".rels")):
        if rel.get("TargetMode") != "External":
            walk(_resolve_target("", rel.get("Target")))

    removed = []
    for f in sorted(root.rglob("*")):
        if not f.is_file():
            continue
        part = f.relative_to(root).as_posix()
        if part == "[Content_Types].xml" or part == "_rels/.rels":
            continue
        if f.parent.name == "_rels":
            owner = (Path(part).parent.parent / f.name[:-len(".rels")]).as_posix()
            keep = owner in reachable
        else:
            keep = part in reachable
        if not keep:
            f.unlink()
            removed.append(part)

    if removed:
        ct_p = root / "[Content_Types].xml"
        ct = _xml(ct_p)
        gone = {f"/{p}" for p in removed}
        for o in ct.getroot().findall(f"{{{NS['ct']}}}Override"):
            if o.get("PartName") in gone:
                ct.getroot().remove(o)
        _write_xml(ct, ct_p)
        print(f"  GC: removed {len(removed)} unreferenced part(s)")


def extract(pptx, selection, output):
    """Split: copy slides at the selected 1-based positions into a new deck."""
    pptx, output = Path(pptx), Path(output)
    if output.resolve() == pptx.resolve():
        print("ERROR: --output must differ from the input deck",
              file=sys.stderr)
        return False
    with tempfile.TemporaryDirectory() as td:
        work = Path(td) / "unpacked"
        unpack(pptx, work)
        pres_p, rels_p, _ = _presentation_parts(work)
        n = len(_slide_entries(_xml(pres_p), _xml(rels_p)))
        try:
            keep = set(_parse_selection(selection, n))
        except ValueError as e:
            print(f"ERROR: {e}", file=sys.stderr)
            return False
        for pos in sorted(set(range(1, n + 1)) - keep, reverse=True):
            if not remove(work, pos):
                return False
        _strip_dangling_slide_links(work)
        _gc_unreferenced(work)
        if not pack(work, output):
            return False
    print(f"Extracted {len(keep)} slide(s) from {pptx} → {output}")
    return True




def _remove_hlink_elements(xml_path, rids):
    """Strip <a:hlinkClick>/<a:hlinkHover> elements whose r:id is in rids.

    Removes just the hlink element — its parent run/shape is untouched.
    Used after a slide-jump Relationship is dropped, so the slide XML
    doesn't reference an rId missing from its .rels. Returns count removed.
    """
    if not rids:
        return 0
    tree = _xml(xml_path)
    doomed = [
        el
        for tag in ("hlinkClick", "hlinkHover")
        for el in tree.getroot().iter(f"{{{NS['a']}}}{tag}")
        if el.get(f"{{{NS['r']}}}id") in rids
    ]
    for el in doomed:
        el.getparent().remove(el)
    if doomed:
        _write_xml(tree, xml_path)
    return len(doomed)


def _strip_dangling_slide_links(src_dir):
    """Drop slide-jump rels whose Target slide no longer exists.

    Used by extract after slide removals: a kept slide may carry a
    hyperlink jump to a removed slide. Removes the Relationship and the
    matching hlink elements from the slide XML; warns per stripped link.
    """
    root = Path(src_dir)
    slides_dir = root / "ppt" / "slides"
    if not slides_dir.is_dir():
        return
    for slide_xml in slides_dir.glob("slide*.xml"):
        part = f"ppt/slides/{slide_xml.name}"
        rels_p = _part_rels_path(root, part)
        if not rels_p.exists():
            continue
        tree = _xml(rels_p)
        dangling = set()
        for rel in list(_iter_rels(tree)):
            if rel.get("Type") != SLIDE_RT:
                continue
            if rel.get("TargetMode") == "External":
                continue
            target = _resolve_target(part, rel.get("Target"))
            if not (root / target).exists():
                print(f"  WARN: stripping dangling slide link "
                      f"{rel.get('Id')} in {part} (target {target} "
                      f"was removed)", file=sys.stderr)
                dangling.add(rel.get("Id"))
                rel.getparent().remove(rel)
        if dangling:
            _write_xml(tree, rels_p)
            _remove_hlink_elements(slide_xml, dangling)



# Merge engine (append subcommand) lives in edit_merge.py — re-exported for
# callers that import append_decks from edit_deck.
try:
    from edit_merge import append_decks  # noqa: E402,F401
except ImportError:  # pragma: no cover - only under `python3 edit_deck.py`
    # Running this file as a script makes it '__main__', so edit_merge's
    # `from edit_deck import ...` re-imports this module while edit_merge is
    # still initialising. Bind a shim; the script's own import resolves once
    # edit_merge finishes.
    def append_decks(*args, **kwargs):
        from edit_merge import append_decks as _append_decks
        return _append_decks(*args, **kwargs)

A_T = "{http://schemas.openxmlformats.org/drawingml/2006/main}t"


def _slide_xml_files(src_dir):
    """slideN.xml paths sorted by N."""
    slides_dir = Path(src_dir) / "ppt" / "slides"
    files = [p for p in slides_dir.glob("slide*.xml") if p.stem[5:].isdigit()]
    return sorted(files, key=lambda p: int(p.stem[5:]))


def inventory(src_dir):
    """Every text run in every slide as [{slide, run, text}] (document order).

    'run' is the index of the <a:t> within its slide — the stable address
    replace_runs() uses. Returns the list (the CLI prints it as JSON).

    NOTE: 'slide' is the slideN.xml file number, not the deck position —
    after a reorder these differ; run the 'list' subcommand to correlate.
    """
    out = []
    for f in _slide_xml_files(src_dir):
        n = int(f.stem[5:])
        tree = _xml(f)
        for i, t in enumerate(tree.iter(A_T)):
            out.append({"slide": n, "run": i, "text": t.text or ""})
    return out


def replace_runs(src_dir, edits_json):
    """Apply [{slide, run, text}] edits; formatting (rPr) is untouched.

    NOTE: 'slide' is the slideN.xml file number, not the deck position —
    after a reorder these differ; run the 'list' subcommand to correlate.

    On error, slides processed before the failure are already written;
    re-unpack from the original .pptx to reset.
    """
    import json
    edits = json.loads(Path(edits_json).read_text(encoding="utf-8"))
    by_slide = {}
    for e in edits:
        by_slide.setdefault(int(e["slide"]), {})[int(e["run"])] = e["text"]
    for f in _slide_xml_files(src_dir):
        n = int(f.stem[5:])
        if n not in by_slide:
            continue
        tree = _xml(f)
        runs = list(tree.iter(A_T))
        for idx, new_text in by_slide[n].items():
            if idx >= len(runs):
                raise SystemExit(
                    f"slide {n}: run {idx} out of range (has {len(runs)})")
            runs[idx].text = new_text
        _write_xml(tree, f)
        print(f"slide {n}: {len(by_slide[n])} run(s) replaced")


def _add_comments_parser(sub):
    """`comments list|plan|apply|strip` — the reviewer revision loop.

    Everything here is done by revision_plan.py / pptx_comments.py; this is
    only the argument surface.
    """
    p_cm = sub.add_parser("comments",
                          help="Reviewer comments → revision plan → outline")
    cm = p_cm.add_subparsers(dest="comments_cmd", required=True)
    c_ls = cm.add_parser("list", help="Print every comment in a deck")
    c_ls.add_argument("pptx")
    c_pl = cm.add_parser("plan", help="Comments (+ emailed feedback) → plan")
    c_pl.add_argument("pptx")
    c_pl.add_argument("--outline", required=True)
    c_pl.add_argument("--feedback", default=None,
                      help="Emailed feedback.md to fold in")
    c_pl.add_argument("--assets-dir", default=None)
    c_pl.add_argument("--output", required=True,
                      help="Path stem: writes <stem>.json and <stem>.md")
    c_ap = cm.add_parser("apply", help="Apply a plan's mechanical actions")
    c_ap.add_argument("plan", help="The <stem>.json written by 'plan'")
    c_ap.add_argument("outline", help="Outline to rewrite in place")
    c_st = cm.add_parser("strip", help="Remove comments before delivery")
    c_st.add_argument("pptx")
    c_st.add_argument("--output", required=True)
    scope = c_st.add_mutually_exclusive_group(required=True)
    scope.add_argument("--slides", default=None,
                       help="Only these 1-based slides, e.g. 3,7")
    scope.add_argument("--resolved", action="store_true",
                       help="Only resolved threads (legacy comments carry no "
                            "resolved state and always survive)")
    scope.add_argument("--all", action="store_true",
                       help="Every comment in the deck")


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[1])
    sub = parser.add_subparsers(dest="cmd", required=True)
    p_un = sub.add_parser("unpack")
    p_un.add_argument("pptx")
    p_un.add_argument("dir")
    p_pk = sub.add_parser("pack")
    p_pk.add_argument("dir")
    p_pk.add_argument("pptx")
    p_ls = sub.add_parser("list")
    p_ls.add_argument("dir")
    p_du = sub.add_parser("duplicate")
    p_du.add_argument("dir")
    p_du.add_argument(
        "n", type=int, help="1-based deck position (matches list output)")
    p_rm = sub.add_parser("remove")
    p_rm.add_argument("dir")
    p_rm.add_argument(
        "n", type=int, help="1-based deck position (matches list output)")
    p_ro = sub.add_parser("reorder")
    p_ro.add_argument("dir")
    p_ro.add_argument(
        "order", help="Comma-separated positions, e.g. 3,1,2,4")
    p_cl = sub.add_parser("clean")
    p_cl.add_argument("dir")
    p_cl.add_argument("--client-safe", action="store_true",
                      help="Also strip comments, the package thumbnail, unused "
                           "layouts/masters and the core-props identity")
    p_cl.add_argument("--drop-notes", action="store_true",
                      help="Remove every speaker-notes part (implies "
                           "--client-safe)")
    p_cl.add_argument("--drop-hidden", action="store_true",
                      help="Remove hidden slides (implies --client-safe)")
    p_sn = sub.add_parser("sanitize-numbers",
                          help="Randomise 3+ digit numbers in slide text and "
                               "chart caches (years 1900-2100 survive)")
    p_sn.add_argument("dir")
    p_sn.add_argument("--seed", type=int, default=None,
                      help="Seed for reproducible sanitizing")
    p_inv = sub.add_parser("inventory", help="JSON dump of every text run (addresses are slideN.xml file numbers)")
    p_inv.add_argument("dir")
    p_rep = sub.add_parser("replace")
    p_rep.add_argument("dir")
    p_rep.add_argument("edits", help="JSON: [{slide, run, text}, ...]")
    p_ex = sub.add_parser("extract", help="Split: new deck from selected slides")
    p_ex.add_argument("pptx")
    p_ex.add_argument("selection", help='1-based positions: "3-7", "2,5,9", "1,3-5"')
    p_ex.add_argument("--output", required=True)
    p_ap = sub.add_parser("append", help="Merge: copy src slides after dst's last slide")
    p_ap.add_argument("dst")
    p_ap.add_argument("src")
    p_ap.add_argument("--slides", default=None,
                      help='Positions in src to copy (default: all), e.g. "2-4"')
    p_ap.add_argument("--output", required=True)
    _add_comments_parser(sub)
    args = parser.parse_args()

    ok = True
    if args.cmd == "unpack":
        unpack(args.pptx, args.dir)
    elif args.cmd == "pack":
        ok = pack(args.dir, args.pptx)
    elif args.cmd == "list":
        list_slides(args.dir)
    elif args.cmd == "duplicate":
        ok = duplicate(args.dir, args.n)
    elif args.cmd == "remove":
        ok = remove(args.dir, args.n)
    elif args.cmd == "reorder":
        ok = reorder(args.dir, args.order)
    elif args.cmd == "clean":
        ok = clean_orphans(args.dir)
        if ok and (args.client_safe or args.drop_notes or args.drop_hidden):
            from client_safe import clean_unpacked
            actions = clean_unpacked(args.dir, drop_notes=args.drop_notes,
                                     drop_hidden=args.drop_hidden)
            for action in actions:
                print(f"  {action}")
            print(f"Client-safety clean: {len(actions)} action(s).")
    elif args.cmd == "sanitize-numbers":
        from client_safe import sanitize_numbers
        replaced = sanitize_numbers(args.dir, args.seed)
        print(f"Sanitized {replaced} number(s) in {args.dir}")
    elif args.cmd == "inventory":
        import json
        print(json.dumps(inventory(args.dir), indent=2, ensure_ascii=False))
    elif args.cmd == "replace":
        replace_runs(args.dir, args.edits)
    elif args.cmd == "extract":
        ok = extract(args.pptx, args.selection, args.output)
    elif args.cmd == "append":
        ok = append_decks(args.dst, args.src, args.slides, args.output)
    elif args.cmd == "comments":
        import revision_plan
        ok = revision_plan.run_cli(args)
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
