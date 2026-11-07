#!/usr/bin/env python3
"""
template_build.py — placeholder-based slides from a user's .pptx template.

`build_template_slide()` maps one outline slide onto the best-matching layout
of the user's template (by placeholder signature — see template_helpers) and
fills its title/body/subtitle/picture placeholders; build_deck falls back to
the styled builders when no layout scores.  Split out of build_deck.py;
build_deck re-exports `TEMPLATE_WANT`, `_fill_placeholders` and
`build_template_slide` under their old names.
"""
from helpers import parse_visual, resolve_image_path
from template_helpers import (LAYOUT_CATEGORIES, find_best_layout,
                              label_placeholder_names)

TEMPLATE_WANT = LAYOUT_CATEGORIES


def _fill_placeholders(slide, p):
    """Fill title/body/subtitle placeholders from the slide dict."""
    from pptx.enum.shapes import PP_PLACEHOLDER
    title_text = p.get("title") or p.get("heading", "")
    body_lines = p.get("bullets", [])
    for ph in slide.placeholders:
        idx = ph.placeholder_format.idx
        ph_type = ph.placeholder_format.type
        if idx == 0 or ph_type in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE):
            ph.text = title_text
        elif ph_type == PP_PLACEHOLDER.SUBTITLE or (idx == 1 and not body_lines):
            ph.text = p.get("subtitle", "")
        elif ph.has_text_frame and body_lines:
            tf = ph.text_frame
            tf.clear()
            for i, line in enumerate(body_lines):
                para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                para.text = line
            body_lines = []  # only fill the first body placeholder


def build_template_slide(prs, p, layout_name, ctx, config=None):
    """Try a placeholder-based slide from the user's template."""
    from pptx.enum.shapes import PP_PLACEHOLDER as PH

    kind, value = parse_visual(p.get("visual", ""))
    if layout_name == "two-column-split" and kind == "image":
        layout, _ = find_best_layout(prs, "picture", config)
        if layout is not None:
            path = resolve_image_path(value, ctx)
            if path:
                slide = prs.slides.add_slide(layout)
                label_placeholder_names(slide, layout)
                _fill_placeholders(slide, p)
                for ph in slide.placeholders:
                    if ph.placeholder_format.type == PH.PICTURE:
                        ph.insert_picture(str(path))
                        break
                return slide

    want = TEMPLATE_WANT.get(layout_name)
    if not want:
        return None
    layout, score = find_best_layout(prs, want, config)
    if layout is None or score < 1:
        return None
    slide = prs.slides.add_slide(layout)
    label_placeholder_names(slide, layout)
    _fill_placeholders(slide, p)
    return slide

