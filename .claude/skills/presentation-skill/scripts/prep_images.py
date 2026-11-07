#!/usr/bin/env python3
"""
prep_images.py — Normalize user-supplied images before embedding in .pptx.

Usage:
    python scripts/prep_images.py assets/user-images/
    python scripts/prep_images.py my-photo.heic --output assets/user-images/my-photo.png

Actions:
  - Applies EXIF orientation (so iPhone portraits don't embed sideways)
  - Converts HEIC, WEBP, BMP, TIFF → PNG (HEIC needs `pip install pillow-heif`)
  - Composites transparency onto white (python-pptx-safe)
  - Resizes oversized images (max 2400px on longest edge)
  - Drops EXIF metadata (PNG re-encode carries none)
"""
import argparse
import sys
from pathlib import Path

try:
    from PIL import Image, ImageOps
except ImportError:
    print("ERROR: Pillow not installed. Run: pip install Pillow", file=sys.stderr)
    sys.exit(1)

HEIC_SUPPORTED = False
try:
    import pillow_heif
    pillow_heif.register_heif_opener()
    HEIC_SUPPORTED = True
except ImportError:
    pass

SUPPORTED_INPUT = {".heic", ".heif", ".webp", ".bmp", ".tiff", ".tif",
                   ".jpg", ".jpeg", ".png", ".gif"}
MAX_LONG_EDGE = 2400


def process_image(input_path, output_path=None):
    input_path = Path(input_path)
    if not input_path.exists():
        print(f"  [SKIP] Not found: {input_path}", file=sys.stderr)
        return False

    suffix = input_path.suffix.lower()
    if suffix not in SUPPORTED_INPUT:
        print(f"  [SKIP] Unsupported format: {input_path.name}")
        return False
    if suffix in (".heic", ".heif") and not HEIC_SUPPORTED:
        print(f"  [ERROR] {input_path.name}: HEIC requires 'pip install pillow-heif'",
              file=sys.stderr)
        return False

    out = Path(output_path) if output_path else input_path.with_suffix(".png")

    try:
        img = Image.open(input_path)
        # Apply EXIF orientation BEFORE any processing, then forget the EXIF.
        img = ImageOps.exif_transpose(img)
    except Exception as e:
        print(f"  [ERROR] Cannot open {input_path.name}: {e}", file=sys.stderr)
        return False

    # Composite transparency onto white for .pptx compatibility
    if img.mode in ("RGBA", "LA", "P"):
        img = img.convert("RGBA")
        bg = Image.new("RGB", img.size, (255, 255, 255))
        bg.paste(img, mask=img.split()[-1])
        img = bg
    elif img.mode != "RGB":
        img = img.convert("RGB")

    w, h = img.size
    long_edge = max(w, h)
    if long_edge > MAX_LONG_EDGE:
        scale = MAX_LONG_EDGE / long_edge
        new_w, new_h = int(w * scale), int(h * scale)
        img = img.resize((new_w, new_h), Image.LANCZOS)
        print(f"  [RESIZE] {input_path.name}: {w}x{h} → {new_w}x{new_h}")

    out.parent.mkdir(parents=True, exist_ok=True)
    img.save(str(out), "PNG", optimize=True)  # PNG re-encode = no EXIF
    size_kb = out.stat().st_size // 1024
    print(f"  [OK] {input_path.name} → {out.name} ({size_kb}KB)")
    return True


def process_directory(dir_path):
    dir_path = Path(dir_path)
    results = {"ok": 0, "skip": 0, "error": 0}
    for f in sorted(dir_path.iterdir()):
        if not (f.is_file() and f.suffix.lower() in SUPPORTED_INPUT):
            continue
        if f.suffix.lower() == ".png":
            # Still normalize: orientation, transparency, size
            if process_image(f, f):
                results["ok"] += 1
            else:
                results["error"] += 1
        elif process_image(f):
            results["ok"] += 1
        else:
            results["error"] += 1
    print(f"\nDone. OK: {results['ok']}, Errors: {results['error']}")
    return results["error"] == 0


def main():
    parser = argparse.ArgumentParser(description="Prepare images for .pptx embedding.")
    parser.add_argument("input", help="Image file or directory of images")
    parser.add_argument("--output", default=None, help="Output path (single file mode only)")
    args = parser.parse_args()

    p = Path(args.input)
    ok = process_directory(p) if p.is_dir() else process_image(p, args.output)
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
