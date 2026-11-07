#!/usr/bin/env python3
"""
fetch_image.py — Download an image from a URL, validate it, and normalize it.

Usage:
    python scripts/fetch_image.py <url> <output_path>
    python scripts/fetch_image.py "https://example.com/logo.png" assets/auto/logo.png

Validates magic bytes (rejects HTML error pages saved as .png), then runs the
same normalization as prep_images.py (orientation, RGB, resize). SVG sources
are rejected — download a PNG rendition instead (most logo sources offer one).
"""
import sys
import tempfile
import urllib.request
from pathlib import Path

MAGIC = {
    b"\x89PNG": "png",
    b"\xff\xd8\xff": "jpeg",
    b"GIF8": "gif",
    b"BM": "bmp",
}


def sniff(data):
    for magic, kind in MAGIC.items():
        if data.startswith(magic):
            return kind
    if data[:4] == b"RIFF" and data[8:12] == b"WEBP":
        return "webp"
    if data[4:12] in (b"ftypheic", b"ftypheix", b"ftypmif1"):
        return "heic"
    head = data[:512].lstrip().lower()
    if head.startswith(b"<?xml") or head.startswith(b"<svg"):
        return "svg"
    if head.startswith(b"<!doctype") or head.startswith(b"<html"):
        return "html"
    return None


def fetch(url, out_path):
    out = Path(out_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    headers = {"User-Agent": "Mozilla/5.0 (compatible; presentation-deck-skill/1.0)"}
    req = urllib.request.Request(url, headers=headers)
    with urllib.request.urlopen(req, timeout=15) as resp:
        ctype = resp.headers.get("Content-Type", "")
        data = resp.read()

    kind = sniff(data)
    if kind == "html":
        raise ValueError(f"URL returned an HTML page, not an image "
                         f"(Content-Type: {ctype}). Check the URL.")
    if kind == "svg":
        raise ValueError("URL returned an SVG. python-pptx cannot embed SVG — "
                         "find a PNG rendition (e.g. Wikipedia thumb URLs).")
    if kind is None:
        raise ValueError(f"Downloaded data is not a recognized image "
                         f"(Content-Type: {ctype}).")

    # Normalize via prep_images (orientation, RGB, max size) into out path
    from prep_images import process_image
    with tempfile.NamedTemporaryFile(suffix=f".{kind}", delete=False) as tmp:
        tmp.write(data)
        tmp_path = Path(tmp.name)
    try:
        if not process_image(tmp_path, out):
            raise ValueError("Image failed normalization (see errors above).")
    finally:
        tmp_path.unlink(missing_ok=True)
    print(f"Saved {len(data) // 1024}KB ({kind}) → {out}")


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: fetch_image.py <url> <output_path>")
        sys.exit(1)
    try:
        fetch(sys.argv[1], sys.argv[2])
    except Exception as e:
        print(f"ERROR: {e}", file=sys.stderr)
        sys.exit(1)
