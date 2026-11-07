# User Images Directory

Place your image files here before running build_deck.py.

Supported formats: PNG, JPG, JPEG, WEBP, HEIC, BMP, TIFF

## How to use images in your outline

Reference files by name in the outline:

```markdown
## Slide 4: Product Demo
**Layout:** full-image
**Visual:** user-image:dashboard-screenshot.png
- Caption: "Live platform — 47ms p99 latency"

## Slide 7: Architecture
**Layout:** two-column-split
**Visual:** user-image:arch-diagram.png
- Heading: "Multi-Region Architecture"
```

## Prep before building

Run this to normalize all images:
```bash
python scripts/prep_images.py assets/user-images/
```
