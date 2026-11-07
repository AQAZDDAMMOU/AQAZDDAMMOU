# Template Mode Guide

## When to Use

Use Template Mode when the user provides an existing .pptx file they want to use as the visual template. The goal is to:
- Preserve the slide master, color palette, fonts, and layout structure exactly
- Generate new content slides that look like they were made by the same designer
- Never overlay text boxes on top of background slides (the classic "AI mistake")

---

## Step 1 — Profile the Template

```bash
# Extract thumbnail images of each slide for visual analysis
python scripts/render_slides.py template.pptx --out assets/template-thumbs/

# Extract all slide layouts and placeholders
python scripts/profile_template.py template.pptx
```

`profile_template.py` outputs a JSON summary (units in inches):
```json
{
  "slide_width_in": 13.333,
  "slide_height_in": 7.5,
  "theme_colors": {"dk1": "0A0F1E", "lt1": "F1F5F9", "accent1": "C9A84C"},
  "slide_layouts": [
    {
      "index": 0,
      "name": "Title Slide",
      "placeholder_count": 2,
      "placeholders": [
        {"idx": 0, "type": "CENTER_TITLE", "left_in": 0.5, "top_in": 1.75, "width_in": 9.0, "height_in": 1.25},
        {"idx": 1, "type": "SUBTITLE", "left_in": 0.5, "top_in": 3.0, "width_in": 9.0, "height_in": 0.75}
      ]
    }
  ]
}
```

## Automatic mapping in build_deck.py

`build_deck.py --template T.pptx` maps layouts by **placeholder-type signature** (TITLE+BODY, CENTER_TITLE, PICTURE, etc.), with layout **name hints as a tiebreaker only** — so renamed or non-English templates still work when their placeholders match. Template layouts are first **clustered by signature** (PPTAgent-style: title / section / content / two-content / picture / caption / blank); `comparison` slides map to a two-body "Two Content"-style cluster when the template has one, degrading to the plain content cluster otherwise (`template_helpers.cluster_layouts`; the advisory `.config.json` from profiling lists the clusters). Covered layouts: `title`, `closing`, `section-divider`, `bullet-list`, and `exec-summary`. Layouts with no template equivalent (charts, cards, timelines, stats, comparison, tables) fall back to styled builders on the template's most-blank layout — they inherit slide dimensions but not master styling. `--ghost` does not apply to slides matched from a user template — template placeholder slides render real content; ghost only replaces the styled builders.

`profile_template.py --generate-config` writes advisory JSON (layout inventory + theme colors) for the agent to read when planning slides; it is not auto-loaded by `build_deck.py`.

---

## Step 2 — Map Content to Layouts

After profiling, map each planned slide to the closest template layout:

```python
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor

def add_content_slide(prs, layout_index, title, content_lines):
    """Add a slide using a specific template layout and fill proper placeholders."""
    slide_layout = prs.slide_layouts[layout_index]
    slide = prs.slides.add_slide(slide_layout)

    for ph in slide.placeholders:
        if ph.placeholder_format.idx == 0:
            ph.text = title
        elif ph.placeholder_format.idx == 1:
            tf = ph.text_frame
            tf.clear()
            for line in content_lines:
                p = tf.add_paragraph()
                p.text = line
                p.level = 0
    return slide
```

**Critical rule:** ALWAYS use `prs.slides.add_slide(layout)` to inherit the template's design. Never add_shape a plain rectangle to simulate a slide background.

---

## Step 3 — Add Images to Template Picture Placeholders

```python
from pptx.util import Inches

def add_image_to_picture_placeholder(slide, placeholder_idx, image_path):
    """Insert image into a picture placeholder, preserving aspect ratio."""
    for ph in slide.placeholders:
        if ph.placeholder_format.idx == placeholder_idx:
            ph.insert_picture(image_path)
            return
    # Fallback: if no picture placeholder, add as free-floating image
    slide.shapes.add_picture(image_path, Inches(6.5), Inches(1.5), Inches(5.8), Inches(4.5))
```

---

## Step 4 — Preserve Template Font Colors

When writing text, use the template's theme colors rather than hardcoded hex:

```python
from pptx.dml.color import RGBColor
from pptx.util import Pt

# Match font to template's heading style
run.font.color.theme_color = 1  # Accent 1 from template theme
# OR use explicit hex from profile JSON
run.font.color.rgb = RGBColor(0xC9, 0xA8, 0x4C)
```

---

## Step 5 — Copy Slides from Template

For slides that should be an exact copy of a template layout (e.g., section dividers):

```python
import copy
from lxml import etree

def duplicate_slide(prs, slide_index):
    """Clone an existing slide to use as a structural base."""
    template_slide = prs.slides[slide_index]
    slide_layout = template_slide.slide_layout
    new_slide = prs.slides.add_slide(slide_layout)

    # Deep copy the XML tree from the original
    new_slide._element._p[:] = copy.deepcopy(template_slide._element._p[:])
    return new_slide
```
