#!/usr/bin/env python3
"""
editability.py — how *editable* a delivered deck actually is.

A deck can look right and still be dead on arrival: a "chart" that is a
flat PNG, a label that is a floating text box parked on top of a card, a
connector anchored to nothing.  The client edits one number and the slide
falls apart.  Two things live here:

* `check_slide(slide, sw, sh, issues, where)` — three deterministic
  per-slide warnings, called once per slide from `qa_check.check_deck()`:
    (a) text-over-shape: a text box whose bounds match (within
        `BOUNDS_TOL_IN`) a textless solid-filled autoshape drawn earlier
        in z-order — the text belongs *in* that shape;
    (b) free connector: a top-level line/connector whose two endpoints
        are both farther than `CONNECTOR_GAP_IN` from every other
        top-level shape — it will not follow anything that is moved;
    (c) tiny fonts + dead space: sub-`TINY_PT` *body* text on a slide
        whose whitespace ratio exceeds `DEAD_SPACE_RATIO` — there is room
        to set it at a readable size.  The fixed source-footnote/footer
        band is chrome, deliberately below body size on every slide, and
        is excluded.

* `deck_level(prs)` — the whole-deck PEI (presentation editability
  index) level L0..L5 with the raw signal counts behind it, reported in
  the `qa_summary.py` delivery record.

`ungrouped_exhibits` is deliberately a *signal only*: it caps the level at
L4 but never raises a warning.  Everything this toolchain builds out of
loose shapes — the shape-built exhibits *and* the card-style core layouts
(exec-summary, cards, stat-callout, timeline, comparison, dashboard,
next-steps, agenda, SCQA) — is wrapped in a real group by
`shape_groups.py` (children of a group are exempt from the count), so a
generated deck reaches L5 and the signal fires on hand-edited or foreign
decks — and on a body too small for `shape_groups`' size gate.
"""
import math

from pptx.enum.dml import MSO_FILL
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn

EMU_IN = 914400

# (a) same-bounds tolerance: builders inset text boxes by >= 0.1in, so
# 0.05in separates "the text box IS the shape" from a deliberate inset.
BOUNDS_TOL_IN = 0.05
# (b) how close a connector endpoint must come to a shape to count as
# anchored to it.
CONNECTOR_GAP_IN = 0.1
# (c) projection legibility floor, and the whitespace ratio above which a
# slide has room to spare.
TINY_PT = 12
DEAD_SPACE_RATIO = 0.55
# Fixed bottom chrome: the source footnote sits at design-y 6.80 of 7.5
# (0.907) and the footer at 7.08.  Both are deliberately set below body
# size on every slide alike, so they are not a "this slide has room"
# defect and are excluded from the tiny-text scan — the same way
# geometry_report excludes chrome from its balance metric.
FOOTNOTE_ZONE_FRAC = 0.90

# deck_level thresholds.
FULL_BLEED_AREA_FRAC = 0.60   # picture this big => the slide IS an image
EXHIBIT_MIN_AREA_SQIN = 2.0   # smaller pictures are icons, not exhibits
UNGROUPED_EXHIBIT_SHAPES = 6  # textless autoshapes that should be a group
# Marker-sized shapes (bullet-state tokens, target ticks, rating dots) are
# not exhibit parts — a 0.35in dot needs no group to be editable.  Mirrors
# the icon-sized-picture exemption in qa_check's contrast gate.
MARKER_MAX_IN = 0.40
PICTURE_MAJORITY_FRAC = 0.50

_TRUE_ATTRS = frozenset(("1", "true", "on"))
_PICTURE_TYPES = (MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.LINKED_PICTURE)
_TEXT_HOST_TYPES = (MSO_SHAPE_TYPE.TEXT_BOX, MSO_SHAPE_TYPE.AUTO_SHAPE)

LEVEL_LABELS = {
    "L0": "rendered image slides",
    "L1": "mostly image slides",
    "L2": "native text, picture exhibits",
    "L3": "native exhibits with editability defects",
    "L4": "native, ungrouped exhibits",
    "L5": "fully native and grouped",
}


# --- shape primitives ----------------------------------------------------

def _is_true(value):
    """OOXML boolean attribute: "1"/"true"/"on", any case; None = absent."""
    return value is not None and str(value).lower() in _TRUE_ATTRS


def _shape_type(shape):
    try:
        return shape.shape_type
    except (AttributeError, TypeError, ValueError, NotImplementedError):
        return None


def leaf_shapes(shapes, in_group=False):
    """(shape, in_group) for every non-group shape, in document order —
    which is z-order, lowest first.  `in_group` is True for anything
    nested inside a GroupShape at any depth."""
    for shape in shapes:
        if _shape_type(shape) == MSO_SHAPE_TYPE.GROUP:
            yield from leaf_shapes(shape.shapes, True)
        else:
            yield shape, in_group


def _bounds_in(shape, allow_flat=False):
    """(left, top, width, height) in inches, or None when the shape has no
    resolved geometry (inheriting placeholders, malformed xfrm).

    `allow_flat` keeps zero-extent geometry, which is real and load-bearing
    for the connector check: an axis-aligned connector has h == 0 (a
    horizontal rule, a dashed benchmark line) or w == 0, and so do the
    hairline rules it may terminate on.  Dropping those would both skip
    the connector and make anything it lands on invisible to the distance
    test.  Area-based callers leave it off — a zero-area shape is never a
    card behind a text box nor an exhibit."""
    try:
        left, top, w, h = shape.left, shape.top, shape.width, shape.height
    except (AttributeError, TypeError):
        return None
    if None in (left, top, w, h):
        return None
    if w < 0 or h < 0 or (not allow_flat and (w == 0 or h == 0)):
        return None
    return (left / EMU_IN, top / EMU_IN, w / EMU_IN, h / EMU_IN)


def _text_of(shape):
    if not getattr(shape, "has_text_frame", False):
        return ""
    try:
        return shape.text_frame.text.strip()
    except (AttributeError, ValueError):
        return ""


def _is_solid(shape):
    try:
        return shape.fill.type == MSO_FILL.SOLID
    except (AttributeError, TypeError, ValueError, NotImplementedError):
        return False


def _is_connector(shape):
    if _shape_type(shape) == MSO_SHAPE_TYPE.LINE:
        return True
    try:
        return shape._element.tag == qn("p:cxnSp")
    except AttributeError:
        return False


def _flips(shape):
    """(flipH, flipV) from the shape's a:xfrm — a connector drawn right to
    left or bottom to top carries the flag instead of a negative extent."""
    try:
        sp_pr = shape._element.find(qn("p:spPr"))
        xfrm = None if sp_pr is None else sp_pr.find(qn("a:xfrm"))
    except AttributeError:
        return False, False
    if xfrm is None:
        return False, False
    return _is_true(xfrm.get("flipH")), _is_true(xfrm.get("flipV"))


def _endpoints(shape, box):
    """The two ends of a line, honouring flipH/flipV."""
    left, top, w, h = box
    flip_h, flip_v = _flips(shape)
    x1, x2 = (left + w, left) if flip_h else (left, left + w)
    y1, y2 = (top + h, top) if flip_v else (top, top + h)
    return ((x1, y1), (x2, y2))


def _point_box_distance(point, box):
    """Shortest distance in inches from a point to a box edge; 0 inside."""
    x, y = point
    left, top, w, h = box
    dx = max(left - x, 0.0, x - (left + w))
    dy = max(top - y, 0.0, y - (top + h))
    return math.hypot(dx, dy)


def _same_bounds(a, b):
    return all(abs(p - q) <= BOUNDS_TOL_IN for p, q in zip(a, b))


# --- (a) text-over-shape -------------------------------------------------

def text_over_shape_hits(slide):
    """Text-bearing shapes parked exactly on top of a textless solid
    autoshape drawn earlier.  Returns the offending shapes."""
    filled_below = []
    hits = []
    for shape, _ in leaf_shapes(slide.shapes):
        box = _bounds_in(shape)
        if box is None:
            continue
        if _text_of(shape):
            if _shape_type(shape) not in _TEXT_HOST_TYPES:
                continue
            if any(_same_bounds(box, prior) for prior in filled_below):
                hits.append(shape)
        elif _shape_type(shape) == MSO_SHAPE_TYPE.AUTO_SHAPE \
                and _is_solid(shape):
            filled_below.append(box)
    return hits


# --- (b) free connectors -------------------------------------------------

def free_connector_hits(slide):
    """Top-level connectors whose *both* endpoints float clear of every
    other top-level shape.  Connectors inside a group move with the group,
    so they are exempt — and since shape-built exhibits are grouped at build
    time (`shape_groups`), this signal effectively fires only on foreign or
    hand-edited decks."""
    entries = [(shape, _bounds_in(shape, allow_flat=True))
               for shape in slide.shapes]
    hits = []
    for shape, box in entries:
        if box is None or not _is_connector(shape):
            continue
        others = [b for other, b in entries
                  if b is not None and other is not shape]
        ends = _endpoints(shape, box)
        if all(all(_point_box_distance(end, b) > CONNECTOR_GAP_IN
                   for b in others) for end in ends):
            hits.append(shape)
    return hits


# --- (c) tiny fonts on a mostly empty slide ------------------------------

def has_tiny_run(slide, sh_in):
    """True when any *body* run carries an explicit size below TINY_PT.

    Three things are deliberately out of scope: runs with no explicit size
    (they inherit from the layout, so the size is not this deck's choice),
    empty runs (unreadable at any size), and the fixed source/footer chrome
    below FOOTNOTE_ZONE_FRAC (set small on every slide alike by design).
    Table cell text and chart data labels are not scanned either — they
    live outside `has_text_frame`, and their sizing is pptx_lint's
    type-scale job."""
    for shape, _ in leaf_shapes(slide.shapes):
        if not getattr(shape, "has_text_frame", False):
            continue
        box = _bounds_in(shape)
        if box is not None and box[1] >= FOOTNOTE_ZONE_FRAC * sh_in:
            continue
        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                size = run.font.size
                if size is not None and size.pt < TINY_PT \
                        and run.text.strip():
                    return True
    return False


def whitespace_ratio(slide, sw_in, sh_in):
    """Uncovered fraction of the slide, from geometry_report's raster."""
    # Imported lazily: geometry_report imports qa_check, which imports
    # this module — a module-level import would close the cycle.
    from geometry_report import _boxes, coverage_metrics
    return coverage_metrics(_boxes(slide), sw_in, sh_in)["whitespace_ratio"]


# --- per-slide entry point ----------------------------------------------

def check_slide(slide, sw, sh, issues, where):
    """Append this slide's editability warnings to `issues["warn"]`.

    Mutates `issues` in place and returns nothing, matching the other
    per-slide `check_*` helpers in qa_check.  `sw`/`sh` are the slide
    dimensions in EMU."""
    for shape in text_over_shape_hits(slide):
        issues["warn"].append(
            f"{where}: {shape.name!r} text box floats over a same-bounds "
            "shape — put the text in the shape")
    for shape in free_connector_hits(slide):
        issues["warn"].append(
            f"{where}: connector {shape.name!r} is anchored to nothing "
            f"(both ends > {CONNECTOR_GAP_IN}in from every shape) — "
            "attach it or group it with what it points at")
    sw_in, sh_in = sw / EMU_IN, sh / EMU_IN
    if has_tiny_run(slide, sh_in) and whitespace_ratio(
            slide, sw_in, sh_in) > DEAD_SPACE_RATIO:
        issues["warn"].append(
            f"{where}: sub-{TINY_PT}pt text on a mostly empty slide — "
            "enlarge")


# --- deck-level PEI ------------------------------------------------------

def _slide_signals(slide, slide_area):
    """Per-slide contribution to the deck signal counts."""
    picture_areas = []
    native_exhibits = 0
    ungrouped_plain_shapes = 0
    has_native_text = False
    for shape, in_group in leaf_shapes(slide.shapes):
        stype = _shape_type(shape)
        box = _bounds_in(shape)
        if getattr(shape, "has_chart", False) \
                or getattr(shape, "has_table", False):
            native_exhibits += 1
        elif stype in _PICTURE_TYPES and box is not None:
            picture_areas.append(box[2] * box[3])
        if _text_of(shape):
            has_native_text = True
        elif stype == MSO_SHAPE_TYPE.AUTO_SHAPE and not in_group:
            marker_sized = box is not None and \
                min(box[2], box[3]) <= MARKER_MAX_IN
            if not marker_sized:
                ungrouped_plain_shapes += 1
    largest = max(picture_areas, default=0.0)
    return {
        "picture_slide": slide_area > 0
        and largest >= FULL_BLEED_AREA_FRAC * slide_area,
        "native_text": has_native_text,
        "native_exhibits": native_exhibits,
        "picture_exhibits": sum(
            1 for a in picture_areas
            if EXHIBIT_MIN_AREA_SQIN <= a < FULL_BLEED_AREA_FRAC * slide_area),
        "text_over_shape": len(text_over_shape_hits(slide)),
        "free_connectors": len(free_connector_hits(slide)),
        "ungrouped_exhibit":
            ungrouped_plain_shapes >= UNGROUPED_EXHIBIT_SHAPES,
    }


def _level_for(signals):
    n = signals["slides"]
    if n == 0:
        return "L5"
    if signals["picture_slides"] == n and signals["native_text_slides"] == 0:
        return "L0"
    if signals["picture_slides"] > PICTURE_MAJORITY_FRAC * n:
        return "L1"
    if signals["native_exhibits"] == 0 and signals["picture_exhibits"] > 0:
        return "L2"
    if signals["text_over_shape"] + signals["free_connectors"] > 0:
        return "L3"
    if signals["ungrouped_exhibits"] > 0:
        return "L4"
    return "L5"


def deck_level(prs):
    """PEI editability level for a whole deck:
    `{"level": "L0".."L5", "label": str, "signals": {...counts...}}`."""
    slide_area = (prs.slide_width / EMU_IN) * (prs.slide_height / EMU_IN)
    signals = {
        "slides": len(prs.slides),
        "picture_slides": 0,
        "native_text_slides": 0,
        "native_exhibits": 0,
        "picture_exhibits": 0,
        "text_over_shape": 0,
        "free_connectors": 0,
        "ungrouped_exhibits": 0,
    }
    for slide in prs.slides:
        s = _slide_signals(slide, slide_area)
        signals["picture_slides"] += int(s["picture_slide"])
        signals["native_text_slides"] += int(s["native_text"])
        signals["native_exhibits"] += s["native_exhibits"]
        signals["picture_exhibits"] += s["picture_exhibits"]
        signals["text_over_shape"] += s["text_over_shape"]
        signals["free_connectors"] += s["free_connectors"]
        signals["ungrouped_exhibits"] += int(s["ungrouped_exhibit"])
    level = _level_for(signals)
    return {"level": level, "label": LEVEL_LABELS[level], "signals": signals}
