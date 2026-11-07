#!/usr/bin/env python3
"""
shape_groups.py — wrap a shape-built exhibit in a real PowerPoint group.

The shape-built exhibits (waterfall, mekko, gantt, harvey scorecard, the
structure layouts …) are drawn as dozens of loose rectangles, connectors and
text boxes.  Visually that is exactly right; as a *deliverable* it is not: the
client drags one bar and the exhibit falls apart, and `editability.deck_level`
caps the deck at L4 ("native, ungrouped exhibits") for precisely that reason.

`group_shapes()` moves the shapes an exhibit builder just drew into one
`p:grpSp` so the exhibit moves, scales and copies as a single object — L5 —
while every child stays individually selectable and editable in PowerPoint.
The card-style core layouts (exec-summary, cards, stat-callout, timeline,
comparison, dashboard, next-steps, agenda, SCQA) are built the same way —
a frame, an accent bar and a text box per card — and hand their body over
too; the heading always stays top-level.

**The child transform is deliberately the identity.**  A `p:grpSp` maps child
coordinates onto the canvas through `chOff/chExt -> off/ext`; setting
`chOff == off` and `chExt == ext` makes that map a no-op, so every child keeps
the absolute EMU geometry the builder computed and every geometry consumer
(qa_check bounds, geometry_report overlaps) reads the same numbers it read
before grouping.  `canvas_shapes()` is the walker that relies on this: it
descends into identity groups (child coords are canvas coords) and stops at a
scaled one (child coords are not).

Grouping is *gated* on size: a builder that drew fewer than
`editability.UNGROUPED_EXHIBIT_SHAPES` shapes has not drawn an exhibit worth
wrapping (a two-bar sketch, a fallback), and a container nobody asked for is
its own small tax.  The gate counts *every* shape the builder hands over, not
just the textless autoshapes that trip the L4 signal: the slide chrome
(background bar, tracker, kicker) is added after the builder returns and
contributes plain shapes of its own, so a body-only count of the L4 signal
would leave slides sitting one shape above the threshold with nothing left to
group.  The gate is evaluated per build, so the same layout groups or not
depending on how much data the slide carries.
"""
from lxml import etree
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn

from editability import UNGROUPED_EXHIBIT_SHAPES

# Group name and marker.  `slide_tags.name_shapes` renames the group to
# `s<NN>-exhibit`; the marker survives that rename so later passes can still
# tell a builder-made exhibit group from a group in a foreign deck.
EXHIBIT_NAME = "exhibit"
EXHIBIT_MARK = "pk:exhibit"

# Element tags this module will move into a p:grpSp, in z-order.  A
# p:graphicFrame (table, native chart, SmartArt) is deliberately absent:
# PowerPoint itself refuses to group one, so a body holding one is left flat
# rather than written into a file PowerPoint would not have produced.
_CHILD_TAGS = frozenset((qn("p:sp"), qn("p:cxnSp"), qn("p:pic"),
                         qn("p:grpSp")))


def _shape_type(shape):
    try:
        return shape.shape_type
    except (AttributeError, TypeError, ValueError, NotImplementedError):
        return None


def _bounds(shape):
    """(left, top, right, bottom) in EMU, or None without resolved geometry."""
    try:
        left, top, width, height = (shape.left, shape.top,
                                    shape.width, shape.height)
    except (AttributeError, TypeError):
        return None
    if None in (left, top, width, height):
        return None
    return (int(left), int(top), int(left) + int(width), int(top) + int(height))


def _xfrm(group_element):
    pr = group_element.find(qn("p:grpSpPr"))
    return None if pr is None else pr.find(qn("a:xfrm"))


def is_group(shape):
    return _shape_type(shape) == MSO_SHAPE_TYPE.GROUP


def is_exhibit_group(shape):
    """True for a group this module created (marker survives renaming)."""
    if not is_group(shape):
        return False
    nv = shape._element.find(qn("p:nvGrpSpPr"))
    pr = None if nv is None else nv.find(qn("p:cNvPr"))
    return pr is not None and pr.get("descr") == EXHIBIT_MARK


def is_identity_group(shape):
    """True when the group's child transform is a no-op, i.e. child left/top
    are already canvas coordinates."""
    if not is_group(shape):
        return False
    xfrm = _xfrm(shape._element)
    if xfrm is None:
        return False
    off, ext = xfrm.find(qn("a:off")), xfrm.find(qn("a:ext"))
    ch_off, ch_ext = xfrm.find(qn("a:chOff")), xfrm.find(qn("a:chExt"))
    if None in (off, ext, ch_off, ch_ext):
        return False
    return (off.get("x"), off.get("y"), ext.get("cx"), ext.get("cy")) == \
        (ch_off.get("x"), ch_off.get("y"),
         ch_ext.get("cx"), ch_ext.get("cy"))


def canvas_shapes(shapes):
    """Every shape whose left/top are canvas coordinates.

    Descends into identity groups — whose children are positioned in canvas
    coordinates — and yields a scaled/offset group as one opaque box, because
    its children's raw coordinates are not comparable with the slide.  Use
    this for bounds work; use `qa_check.iter_shapes` when you want every
    shape, container included.
    """
    for shape in shapes:
        if is_group(shape) and is_identity_group(shape):
            yield from canvas_shapes(shape.shapes)
        else:
            yield shape


def _new_group_element(spTree, shape_id, name, box):
    """A `p:grpSp` with an identity child transform over `box`."""
    left, top, right, bottom = box
    group = etree.SubElement(spTree, qn("p:grpSp"))
    nv = etree.SubElement(group, qn("p:nvGrpSpPr"))
    pr = etree.SubElement(nv, qn("p:cNvPr"))
    pr.set("id", str(shape_id))
    pr.set("name", name)
    pr.set("descr", EXHIBIT_MARK)
    etree.SubElement(nv, qn("p:cNvGrpSpPr"))
    etree.SubElement(nv, qn("p:nvPr"))
    grp_pr = etree.SubElement(group, qn("p:grpSpPr"))
    xfrm = etree.SubElement(grp_pr, qn("a:xfrm"))
    for tag, attrs in ((qn("a:off"), {"x": str(left), "y": str(top)}),
                       (qn("a:ext"), {"cx": str(right - left),
                                      "cy": str(bottom - top)}),
                       (qn("a:chOff"), {"x": str(left), "y": str(top)}),
                       (qn("a:chExt"), {"cx": str(right - left),
                                        "cy": str(bottom - top)})):
        element = etree.SubElement(xfrm, tag)
        for key, value in attrs.items():
            element.set(key, value)
    return group


def group_shapes(slide, shapes, name=EXHIBIT_NAME):
    """Move `shapes` into one `p:grpSp` on `slide`; return the group shape.

    Returns None — leaving the slide untouched — when the shapes are not a
    real exhibit (fewer than `UNGROUPED_EXHIBIT_SHAPES` of them, none with
    resolvable geometry, or one that is not a top-level member of this
    slide's shape tree) or when any of them is a graphic frame — a table or
    a native chart, which PowerPoint will not group.  Z-order is preserved:
    children keep their relative order and the group takes the position of
    the lowest of them.
    """
    shapes = [shape for shape in shapes if shape is not None]
    if len(shapes) < UNGROUPED_EXHIBIT_SHAPES:
        return None

    spTree = slide.shapes._spTree
    members = []
    for shape in shapes:
        element = shape._element
        if element.tag not in _CHILD_TAGS or element.getparent() is not spTree:
            return None
        members.append(element)
    boxes = [box for box in (_bounds(shape) for shape in shapes)
             if box is not None]
    if len(members) < 2 or not boxes:
        return None

    box = (min(b[0] for b in boxes), min(b[1] for b in boxes),
           max(b[2] for b in boxes), max(b[3] for b in boxes))
    # spTree document order is z-order; group the members in that order and
    # drop the group in at the lowest one's slot.
    order = {id(element): index for index, element in enumerate(spTree)}
    members.sort(key=lambda element: order[id(element)])
    insert_at = order[id(members[0])]

    group = _new_group_element(spTree, slide.shapes._next_shape_id, name, box)
    spTree.remove(group)
    spTree.insert(insert_at, group)
    for element in members:
        group.append(element)  # append() moves the element
    return next(shape for shape in slide.shapes if shape._element is group)
