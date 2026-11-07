#!/usr/bin/env python3
"""
edit_merge.py — the deck-merge engine behind `edit_deck.py append`.

Copies slides between .pptx packages with their full dependency graph
(layout, master, theme, media, charts), renaming parts to avoid collisions
and remapping every rId. Split out of edit_deck.py (file-size budget);
edit_deck re-exports append_decks so the CLI and imports are unchanged.
"""
import posixpath
import re
import shutil
import sys
import tempfile
from pathlib import Path

from lxml import etree

from edit_deck import (
    LAYOUT_RT, MASTER_RT, NOTES_RT, NS, SLIDE_RT,
    _iter_rels, _parse_selection, _part_rels_path, _presentation_parts,
    _relative_target, _remove_hlink_elements, _resolve_target,
    _slide_entries, _write_xml, _xml, pack, unpack)


def _ct_index(root):
    """[Content_Types].xml → ({'/part': ct}, {'ext': ct})."""
    tree = _xml(Path(root) / "[Content_Types].xml")
    overrides = {o.get("PartName"): o.get("ContentType")
                 for o in tree.getroot().findall(f"{{{NS['ct']}}}Override")}
    defaults = {d.get("Extension").lower(): d.get("ContentType")
                for d in tree.getroot().findall(f"{{{NS['ct']}}}Default")}
    return overrides, defaults


def _ct_register(dst_ct, src_ct, src_part, new_part):
    """Mirror src's content-type declaration for new_part into dst's tree."""
    overrides, defaults = src_ct
    dst_root = dst_ct.getroot()
    ct = overrides.get(f"/{src_part}")
    if ct is not None:
        if not any(o.get("PartName") == f"/{new_part}" for o in
                   dst_root.findall(f"{{{NS['ct']}}}Override")):
            o = etree.SubElement(dst_root, f"{{{NS['ct']}}}Override")
            o.set("PartName", f"/{new_part}")
            o.set("ContentType", ct)
        return
    ext = src_part.rsplit(".", 1)[-1].lower()
    ct = defaults.get(ext)
    if ct is None:
        raise ValueError(f"no content type for {src_part} in source deck")
    existing = {d.get("Extension").lower(): d.get("ContentType")
                for d in dst_root.findall(f"{{{NS['ct']}}}Default")}
    if ext not in existing:
        d = etree.SubElement(dst_root, f"{{{NS['ct']}}}Default")
        d.set("Extension", ext)
        d.set("ContentType", ct)
    elif existing[ext] != ct:  # rare: same extension, different type
        o = etree.SubElement(dst_root, f"{{{NS['ct']}}}Override")
        o.set("PartName", f"/{new_part}")
        o.set("ContentType", ct)


def _new_part_name(dst_root, src_part, counters):
    """Collision-free dst name continuing the family numbering.

    'ppt/slideLayouts/slideLayout3.xml' → 'ppt/slideLayouts/slideLayout12.xml'
    (12 = max existing index in dst + copies already made this run + 1).
    Extensionless filenames keep the whole name as the stem (no mangling).
    """
    d, fname = posixpath.dirname(src_part), posixpath.basename(src_part)
    if "." in fname:
        stem, _, ext = fname.rpartition(".")
    else:
        stem, ext = fname, ""
    prefix = stem.rstrip("0123456789") or "part"
    key = (d, prefix.lower(), ext.lower())
    if key not in counters:
        suffix = (r"\." + re.escape(ext) if ext else "") + r"$"
        rx = re.compile(re.escape(prefix) + r"(\d+)" + suffix, re.IGNORECASE)
        mx = 0
        dst_dir = Path(dst_root) / d
        if dst_dir.is_dir():
            for f in dst_dir.iterdir():
                m = rx.fullmatch(f.name)
                if m:
                    mx = max(mx, int(m.group(1)))
        counters[key] = mx
    counters[key] += 1
    new_fname = (f"{prefix}{counters[key]}.{ext}" if ext
                 else f"{prefix}{counters[key]}")
    return f"{d}/{new_fname}"


def _copy_part_tree(state, src_part):
    """Copy src_part and its internal dependency graph into dst.

    Returns the new package path. rIds inside each copied part stay valid
    because its .rels is copied wholesale — only Targets are rewritten.
    Masters are special-cased: their slideLayout relationships are deferred
    and pruned in _finalize_masters to just the layouts copied this run.
    """
    if src_part in state["done"]:
        return state["map"][src_part]
    new_part = state["map"].get(src_part) or _new_part_name(
        state["dst"], src_part, state["counters"])
    state["map"][src_part] = new_part
    state["done"].add(src_part)

    src_file = state["src"] / src_part
    dst_file = state["dst"] / new_part
    dst_file.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy(src_file, dst_file)
    _ct_register(state["dst_ct"], state["src_ct"], src_part, new_part)

    src_rels = _part_rels_path(state["src"], src_part)
    if not src_rels.exists():
        return new_part
    tree = _xml(src_rels)
    root = tree.getroot()
    is_master = src_part.startswith("ppt/slideMasters/")
    dropped_rids = set()
    for rel in list(_iter_rels(tree)):
        if rel.get("TargetMode") == "External":
            continue
        rtype = rel.get("Type")
        target = _resolve_target(src_part, rel.get("Target"))
        if rtype == NOTES_RT:
            root.remove(rel)  # notes are dropped on import (documented)
            continue
        if is_master and rtype == LAYOUT_RT:
            state["master_layouts"].setdefault(new_part, []).append(
                (rel, target))
            continue
        if rtype == SLIDE_RT and target not in state["selected"]:
            print(f"  WARN: dropping link from {src_part} to unselected "
                  f"{target}", file=sys.stderr)
            dropped_rids.add(rel.get("Id"))
            root.remove(rel)
            continue
        new_target = _copy_part_tree(state, target)
        rel.set("Target", _relative_target(new_part, new_target))
    if dropped_rids:
        # The copied XML still carries <a:hlinkClick r:id="..."> pointing
        # at the rIds just dropped — strip them so the part stays valid.
        _remove_hlink_elements(dst_file, dropped_rids)
    if is_master:
        state["master_rels"][new_part] = tree  # written in _finalize_masters
    else:
        rels_out = _part_rels_path(state["dst"], new_part)
        rels_out.parent.mkdir(parents=True, exist_ok=True)
        _write_xml(tree, rels_out)
    return new_part


def _max_visual_id(dst_root):
    """Max id across sldMasterIdLst + every master's sldLayoutIdLst.
    Master/layout ids are presentation-unique with floor 2147483648."""
    mx = 2147483647
    pres = _xml(Path(dst_root) / "ppt" / "presentation.xml")
    for el in pres.getroot().findall("p:sldMasterIdLst/p:sldMasterId", NS):
        mx = max(mx, int(el.get("id")))
    for m in (Path(dst_root) / "ppt" / "slideMasters").glob("slideMaster*.xml"):
        for el in _xml(m).getroot().findall("p:sldLayoutIdLst/p:sldLayoutId", NS):
            mx = max(mx, int(el.get("id")))
    return mx


def _finalize_masters(state, idctr):
    """Write deferred master rels; prune+renumber each sldLayoutIdLst.

    A copied master's rels reference ALL its source layouts — keep only the
    layouts actually copied this run, and give surviving entries fresh
    presentation-unique ids. Returns the updated id counter.
    """
    for new_master, tree in state["master_rels"].items():
        kept_rids = set()
        for rel, src_target in state["master_layouts"].get(new_master, []):
            if src_target in state["map"]:
                rel.set("Target", _relative_target(
                    new_master, state["map"][src_target]))
                kept_rids.add(rel.get("Id"))
            else:
                rel.getparent().remove(rel)
        rels_out = _part_rels_path(state["dst"], new_master)
        rels_out.parent.mkdir(parents=True, exist_ok=True)
        _write_xml(tree, rels_out)

        master_p = state["dst"] / new_master
        mt = _xml(master_p)
        lst = mt.getroot().find("p:sldLayoutIdLst", NS)
        if lst is not None:
            for el in list(lst):
                if el.get(f"{{{NS['r']}}}id") in kept_rids:
                    idctr += 1
                    el.set("id", str(idctr))
                else:
                    lst.remove(el)
        _write_xml(mt, master_p)
    return idctr


def append_decks(dst_pptx, src_pptx, slides_spec, output):
    """Merge: copy selected src slides (default all) after dst's last slide.

    Copies each slide with its full dependency graph — layout, master, theme,
    media, charts (+ colors/style/embedded xlsx) — renamed to avoid
    collisions. Source slides keep their own master/theme (faithful import);
    speaker notes are dropped.
    """
    dst_pptx, src_pptx, output = Path(dst_pptx), Path(src_pptx), Path(output)
    with tempfile.TemporaryDirectory() as td:
        dst_root, src_root = Path(td) / "dst", Path(td) / "src"
        unpack(dst_pptx, dst_root)
        unpack(src_pptx, src_root)
        s_pres, s_rels, _ = _presentation_parts(src_root)
        entries = _slide_entries(_xml(s_pres), _xml(s_rels))
        n = len(entries)
        try:
            sel = _parse_selection(slides_spec or f"1-{n}", n)
        except ValueError as e:
            print(f"ERROR: {e}", file=sys.stderr)
            return False

        selected_parts = [f"ppt/slides/{entries[p - 1][2]}" for p in sel]
        state = {
            "src": src_root, "dst": dst_root,
            "map": {}, "done": set(), "counters": {},
            "selected": set(selected_parts),
            "master_rels": {}, "master_layouts": {},
            "src_ct": _ct_index(src_root),
            "dst_ct": _xml(dst_root / "[Content_Types].xml"),
        }
        # Pre-assign slide names so slide→slide links within the selection
        # resolve regardless of copy order.
        for part in selected_parts:
            state["map"][part] = _new_part_name(
                dst_root, part, state["counters"])
        try:
            new_slides = [_copy_part_tree(state, p) for p in selected_parts]
        except (ValueError, OSError, etree.XMLSyntaxError) as e:
            print(f"ERROR: copy failed: {e}", file=sys.stderr)
            return False
        idctr = _finalize_masters(state, _max_visual_id(dst_root))
        _write_xml(state["dst_ct"], dst_root / "[Content_Types].xml")

        # Register new slides + masters in dst presentation.xml(.rels).
        d_pres_p, d_rels_p, _ = _presentation_parts(dst_root)
        pres, rels = _xml(d_pres_p), _xml(d_rels_p)
        rel_root = rels.getroot()
        max_rid = max((int(r.get("Id")[3:]) for r in _iter_rels(rels)
                       if re.fullmatch(r"rId\d+", r.get("Id") or "")),
                      default=0)
        sld_lst = pres.getroot().find("p:sldIdLst", NS)
        max_sldid = max((int(e.get("id")) for e in
                         sld_lst.findall("p:sldId", NS)), default=255)
        for new_slide in new_slides:
            max_rid += 1
            rel = etree.SubElement(rel_root, f"{{{NS['rel']}}}Relationship")
            rel.set("Id", f"rId{max_rid}")
            rel.set("Type", SLIDE_RT)
            rel.set("Target", new_slide[len("ppt/"):])
            max_sldid += 1
            sld = etree.SubElement(sld_lst, f"{{{NS['p']}}}sldId")
            sld.set("id", str(max_sldid))
            sld.set(f"{{{NS['r']}}}id", f"rId{max_rid}")
        master_lst = pres.getroot().find("p:sldMasterIdLst", NS)
        for new_master in state["master_rels"]:
            max_rid += 1
            rel = etree.SubElement(rel_root, f"{{{NS['rel']}}}Relationship")
            rel.set("Id", f"rId{max_rid}")
            rel.set("Type", MASTER_RT)
            rel.set("Target", new_master[len("ppt/"):])
            idctr += 1
            mid = etree.SubElement(master_lst, f"{{{NS['p']}}}sldMasterId")
            mid.set("id", str(idctr))
            mid.set(f"{{{NS['r']}}}id", f"rId{max_rid}")
        _write_xml(rels, d_rels_p)
        _write_xml(pres, d_pres_p)

        if not pack(dst_root, output):
            return False
    print(f"Appended {len(new_slides)} slide(s) from {src_pptx} → {output} "
          f"({len(state['master_rels'])} master(s) imported; notes dropped)")
    return True

