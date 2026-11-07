#!/usr/bin/env python3
"""
deck_nav.py — post-build hyperlink pass (agenda rows and tracker chips jump
to their section divider).

`link_navigation(prs, built_slides, slides_data)` runs after
build_deck.build()'s loop, when built Slide objects exist to point
`click_action.target_slide` at.  Matching is by heading text (casefold):

  - on agenda-layout slides, any text shape whose text equals a
    section-divider heading (auto and manual agendas alike; sub-chapter rows
    and heading text simply do not match);
  - on every other slide, only tracker chrome (`sNN-tracker*` shape names
    from slide_tags.name_shapes) — the tab-strip chip labels and the plain
    section label — matched against the full or chip-truncated heading.

Unmatched rows are skipped silently: navigation is a convenience and QA
never fails on a missing link.
"""
from pptx.enum.shapes import MSO_SHAPE_TYPE

from slide_chrome import _tab_label


def _walk(shapes):
    for shape in shapes:
        yield shape
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from _walk(shape.shapes)


def _divider_targets(built_slides, slides_data):
    """heading.casefold() (full and chip-truncated) → built divider slide.
    First divider with a given heading wins."""
    targets = {}
    for p, slide in zip(slides_data, built_slides):
        if p.get("layout") != "section-divider":
            continue
        name = (p.get("heading") or p.get("title", "")).strip()
        if not name:
            continue
        targets.setdefault(name.casefold(), slide)
        targets.setdefault(_tab_label(name).casefold(), slide)
    return targets


def link_navigation(prs, built_slides, slides_data):
    """Wire agenda rows and tracker chips to their section dividers.

    Returns the number of links added."""
    targets = _divider_targets(built_slides, slides_data)
    if not targets:
        return 0
    links = 0
    for p, slide in zip(slides_data, built_slides):
        is_agenda = p.get("layout") == "agenda"
        for shape in _walk(slide.shapes):
            if not getattr(shape, "has_text_frame", False):
                continue
            if not is_agenda and not shape.name.split("-", 1)[-1] \
                    .startswith("tracker"):
                continue
            target = targets.get(shape.text_frame.text.strip().casefold())
            if target is None or target is slide:
                continue
            try:
                shape.click_action.target_slide = target
                links += 1
            except (AttributeError, ValueError):
                continue  # shape type without click action — skip silently
    return links
