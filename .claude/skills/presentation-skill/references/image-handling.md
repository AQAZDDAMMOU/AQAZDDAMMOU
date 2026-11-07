# Image Handling Guide

## Overview

Three image sources, one notation in the outline:

| Source | How to provide | Outline notation |
|--------|---------------|------------------|
| **User-uploaded file** | User shares file in chat → save to `assets/user-images/` | `user-image:filename.png` |
| **URL** | Direct image URL → `scripts/fetch_image.py` | `auto-image:filename.png` (after download) |
| **Auto-sourced** | Found via image search, then downloaded | `auto-image:filename.png` |
| **Hero (title slide)** | Same as user-uploaded; optional alias | `hero-image:filename.png` or `user-image:filename.png` |

Paths resolve relative to `--assets-dir` (default `./assets`), the outline's directory, or as literal paths — `build_deck.py --check` errors on any image it cannot find, so validate before building.

---

## Step 1 — Collect User Images

When the user wants their own images included, ask:

> "Please share the image files in our chat, and for each one let me know:
> 1. Which slide number (or topic area) you want it on
> 2. How it should be displayed: full-bleed background, right-column visual, or comparison side
> 3. Any crop/size preferences"

Save each uploaded image to `assets/user-images/` with a descriptive filename.

---

## Step 2 — Normalize (ALWAYS, before building)

```bash
python3 scripts/prep_images.py assets/user-images/
```

This applies EXIF orientation (iPhone portraits would otherwise embed sideways), converts HEIC/WEBP/BMP/TIFF → PNG, composites transparency onto white, and resizes to max 2400px. HEIC inputs need `pip install pillow-heif`.

---

## Step 3 — Placement via the Outline

All placement preserves aspect ratio automatically (`helpers.add_picture_contain` / `add_picture_cover` — never raw `add_picture` with a fixed box, which distorts).

**Full-bleed background** — `full-image` layout. Image cover-crops to fill the slide (no stretching); a 50%-alpha overlay is added under the caption:

```markdown
## Slide 4: Product Demo
**Layout:** full-image
**Visual:** user-image:product-dashboard.png
- Caption: "Real-time analytics — 47ms p99 latency"
```

**Right-column visual** — `two-column-split` layout. Image fits within the right column, centered:

```markdown
## Slide 7: Architecture
**Layout:** two-column-split
**Visual:** user-image:arch-diagram.png
- Heading: "Scalable Multi-Region Deployment"
- Active-active across 3 AWS regions
- 99.99% SLA since Jan 2026
```

**Side-by-side comparison** — `comparison` layout:

```markdown
## Slide 9: Before / After
**Layout:** comparison
**Visual-Left:** user-image:before.png
**Visual-Right:** user-image:after.png
- Left label: "Legacy Stack (2023)"
- Right label: "New Architecture (2026)"
```

**Hero on the title slide** — `**Visual:** user-image:hero.png` on a `title` layout places the image right of the title.

For custom placements in hand-written code, use the helpers directly:

```python
from helpers import add_picture_contain, add_picture_cover, add_overlay

add_picture_contain(slide, path, left=6.9, top=0.9, w=6.0, h=5.9, alt="diagram")
add_picture_cover(slide, prs, path, alt="hero")     # full-bleed, crops overflow
add_overlay(slide, 0, 5.2, 13.33, 2.3, alpha_pct=50)  # caption readability
```

Every placed image gets alt text (from Caption/Heading) for accessibility.

---

## Step 4 — Auto-Sourcing Logos & Brand Assets

When no user image is provided but a logo/brand visual is needed:

1. Find an official **PNG** URL via image search (SVG cannot be embedded — Wikipedia "thumb" URLs provide PNG renditions of SVG logos).
2. Download + validate + normalize in one step:
   ```bash
   python3 scripts/fetch_image.py "https://.../logo.png" assets/auto/company-logo.png
   ```
   The script rejects HTML error pages and SVG payloads, and runs the same normalization as `prep_images.py`.
3. For dark palettes, prefer a white/light logo variant.
