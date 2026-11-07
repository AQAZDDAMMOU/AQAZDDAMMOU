#!/usr/bin/env python3
"""
deck_sections.py — PowerPoint sections (p14:sectionLst) for a built deck.

Section dividers in the outline become real PowerPoint sections (the
navigation pane / slide sorter groups), default on when the deck has 2+
non-appendix `section-divider` slides; `**Sections:** off` disables.  Split
out of build_deck.py; build_deck re-exports `_section_plan` and
`inject_sections` under their old names.
"""
from helpers import warn

P14_NS = "http://schemas.microsoft.com/office/powerpoint/2010/main"
SECTIONLST_EXT_URI = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"


def _section_plan(slides_data):
    """Group slide indices into ordered, named sections.

    Pre-divider slides -> 'Opening'; each non-appendix section-divider starts
    a section named by its heading; appendix slides -> 'Backup'.
    Returns [(name, [slide_index, ...]), ...] with empty groups dropped.
    """
    plan = [["Opening", []]]
    for i, p in enumerate(slides_data):
        if p.get("_appendix"):
            if plan[-1][0] != "Backup":
                plan.append(["Backup", []])
        elif p.get("layout") == "section-divider":
            plan.append([p.get("heading") or p.get("title") or "Section", []])
        plan[-1][1].append(i)
    return [(name, idxs) for name, idxs in plan if idxs]


def inject_sections(prs, slides_data, meta):
    """Inject PowerPoint sections into presentation.xml (p14:sectionLst).

    Default on when the deck has 2+ non-appendix section dividers;
    `**Sections:** off` disables. Non-template builds only: built slides
    must correspond 1:1 to slides_data for the slide-id grouping to hold.
    Deterministic section GUIDs (uuid5 of index+name) keep rebuilds stable.
    Returns the section names injected, or None when skipped.
    """
    import uuid

    from lxml import etree
    from pptx.oxml.ns import qn

    mode = meta.get("sections", "").lower()
    if mode == "off":
        return None
    if mode not in ("", "on"):
        warn(f"unknown **Sections:** value {meta['sections']!r} — "
             "expected on|off; using default")
    dividers = [p for p in slides_data
                if p.get("layout") == "section-divider"
                and not p.get("_appendix")]
    if len(dividers) < 2:
        return None
    slide_ids = [s.get("id") for s in prs.slides._sldIdLst]
    if len(slide_ids) != len(slides_data):
        warn("sections skipped: built slides do not match the outline 1:1")
        return None

    p14 = "{%s}" % P14_NS
    section_lst = etree.Element(p14 + "sectionLst", nsmap={"p14": P14_NS})
    plan = _section_plan(slides_data)
    for index, (name, idxs) in enumerate(plan):
        guid = uuid.uuid5(uuid.NAMESPACE_URL, f"deck-section-{index}-{name}")
        section = etree.SubElement(section_lst, p14 + "section")
        section.set("name", name)
        section.set("id", "{%s}" % str(guid).upper())
        sld_id_lst = etree.SubElement(section, p14 + "sldIdLst")
        for i in idxs:
            etree.SubElement(sld_id_lst, p14 + "sldId").set(
                "id", str(slide_ids[i]))

    ext = etree.Element(qn("p:ext"))
    ext.set("uri", SECTIONLST_EXT_URI)
    ext.append(section_lst)
    ext_lst = prs.element.find(qn("p:extLst"))
    if ext_lst is None:
        # p:extLst is the last child in the CT_Presentation sequence
        ext_lst = etree.SubElement(prs.element, qn("p:extLst"))
    ext_lst.append(ext)
    return [name for name, _ in plan]

