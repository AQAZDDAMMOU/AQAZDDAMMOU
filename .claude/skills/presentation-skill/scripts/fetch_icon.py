#!/usr/bin/env python3
"""
fetch_icon.py — Fetch an open-source icon, tint it to a palette color, cache it.

Usage:
    python scripts/fetch_icon.py rocket --color C9A84C [--out assets/icons/]

Sources (tried in order):
  1. Tabler Icons PNG  (https://tabler.io/icons, ~5,900 icons, MIT)
  2. Lucide SVG via cairosvg if installed (https://lucide.dev, MIT)

Icon names are kebab-case (rocket, check, shield-check, trending-up...).
Cached as assets/icons/<name>-<color>.png; build_deck reuses the cache, so a
deck builds offline once its icons have been fetched.
"""
import argparse
import io
import sys
import urllib.request
from pathlib import Path

from PIL import Image

TABLER_PNG = "https://cdn.jsdelivr.net/npm/@tabler/icons-png@latest/icons/outline/{name}.png"
LUCIDE_SVG = "https://unpkg.com/lucide-static@latest/icons/{name}.svg"
UA = {"User-Agent": "Mozilla/5.0 (compatible; presentation-deck-skill/1.0)"}


def _get(url, timeout=15):
    req = urllib.request.Request(url, headers=UA)
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return resp.read()


def _tint(img, color_hex):
    """Recolor a monochrome icon using its alpha channel."""
    img = img.convert("RGBA")
    r, g, b = (int(color_hex[i:i + 2], 16) for i in (0, 2, 4))
    tinted = Image.new("RGBA", img.size, (r, g, b, 0))
    tinted.putalpha(img.getchannel("A"))
    return tinted


def _from_tabler(name, color_hex):
    data = _get(TABLER_PNG.format(name=name))
    return _tint(Image.open(io.BytesIO(data)), color_hex)


def _from_lucide(name, color_hex):
    import cairosvg  # optional dependency
    svg = _get(LUCIDE_SVG.format(name=name)).decode("utf-8")
    svg = svg.replace('stroke="currentColor"', f'stroke="#{color_hex}"')
    png = cairosvg.svg2png(bytestring=svg.encode(), output_width=256,
                           output_height=256)
    return Image.open(io.BytesIO(png)).convert("RGBA")


def fetch_icon(name, color_hex, out_dir="assets/icons"):
    """Return the cached path for icon `name` tinted `color_hex`, fetching if needed.
    Returns None (with a stderr warning) if the icon cannot be obtained."""
    color_hex = color_hex.lstrip("#").upper()
    out = Path(out_dir) / f"{name}-{color_hex}.png"
    if out.is_file():
        return out
    out.parent.mkdir(parents=True, exist_ok=True)

    errors = []
    for source in (_from_tabler, _from_lucide):
        try:
            img = source(name, color_hex)
            img.save(out, "PNG")
            print(f"  [ICON] {name} ({source.__name__[6:]}) → {out}")
            return out
        except ImportError:
            errors.append("lucide: cairosvg not installed")
        except Exception as e:
            errors.append(f"{source.__name__[6:]}: {e}")
    print(f"  [WARN] icon '{name}' unavailable ({'; '.join(errors)})",
          file=sys.stderr)
    return None


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Fetch and tint an icon.")
    parser.add_argument("name", help="Icon name, kebab-case (e.g. shield-check)")
    parser.add_argument("--color", default="C9A84C", help="Hex tint color")
    parser.add_argument("--out", default="assets/icons", help="Cache directory")
    args = parser.parse_args()
    path = fetch_icon(args.name, args.color, args.out)
    sys.exit(0 if path else 1)
