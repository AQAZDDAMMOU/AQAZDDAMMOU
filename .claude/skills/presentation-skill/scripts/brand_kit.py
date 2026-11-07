#!/usr/bin/env python3
"""
brand_kit.py — Build a custom deck palette from a company's web presence.

Usage:
    python3 scripts/brand_kit.py <domain-or-url> --name acme [--assets-dir assets]
    BRANDFETCH_API_KEY=... python3 scripts/brand_kit.py acme.com --name acme

Source 1: Brandfetch API when BRANDFETCH_API_KEY is set (colors, fonts, logo).
Source 2 (fallback): homepage HTML + linked same-origin CSS scrape — hex/rgb()
colors ranked by frequency. Either way the colors are classified into the nine
required palette keys (contrast-checked: text 4.5:1, accents 3.0:1 on bg —
warns when a target can't be met), validated through palettes.load_custom_palettes,
and written to <assets>/palettes/<name>.json. Logo (Brandfetch only) lands at
<assets>/brand/<name>-logo.png. Build with `--palette <name>` afterwards.
"""
import argparse
import colorsys
import json
import os
import re
import sys
import tempfile
import urllib.error
import urllib.parse
import urllib.request
from collections import Counter
from pathlib import Path

from palettes import load_custom_palettes
import profiles
from qa_check import _luminance, contrast_ratio

UA = "Mozilla/5.0 (compatible; presentation-deck-skill/1.0)"
BRANDFETCH_URL = "https://api.brandfetch.io/v2/brands/{domain}"
HEX_RX = re.compile(r"#(?:([0-9a-fA-F]{6})|([0-9a-fA-F]{3}))(?![0-9a-fA-F])")
RGB_RX = re.compile(r"rgba?\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})")
LINK_RX = re.compile(r"<link\b[^>]*>", re.I)
HREF_RX = re.compile(r"""href\s*=\s*["']?([^"'\s>]+)""", re.I)

DEDUPE_DIST = 30        # Euclidean RGB distance for near-identical merge
ACCENT_MIN_DIST = 40    # pairwise distinctness between picked accents
ACCENT_SAT_MIN = 0.25   # HSV saturation floor for accent candidates
ACCENT_LUM_MIN, ACCENT_LUM_MAX = 0.04, 0.96   # near-black/white excluded
DARK_BG_LUM, LIGHT_BG_LUM = 0.15, 0.80        # "very dark" / "very light"
MAX_CSS_FILES = 5
ROLES = ("bg", "bg_deep", "surface", "accent1", "accent2", "accent3",
         "text", "text_muted")


# ── Small color helpers (pure) ───────────────────────────────────────────────

def _rgb(h):
    h = h.lstrip("#")
    return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def _hex6(rgb):
    return "%02X%02X%02X" % tuple(rgb)


def _dist(a, b):
    return sum((x - y) ** 2 for x, y in zip(a, b)) ** 0.5


def _sat(rgb):
    return colorsys.rgb_to_hsv(*(c / 255 for c in rgb))[1]


def _lighten_f(rgb_f, f):
    """Float-internal lighten: each channel moves fraction f toward 255."""
    return tuple(c + (255 - c) * f for c in rgb_f)


def _darken_f(rgb_f, f):
    """Float-internal darken: each channel scaled by (1-f)."""
    return tuple(c * (1 - f) for c in rgb_f)


def _f2rgb(rgb_f):
    """Round float triple to int RGB, clamping to [0, 255]."""
    return tuple(max(0, min(255, round(c))) for c in rgb_f)


def lighten(rgb, f):
    return _f2rgb(_lighten_f(rgb, f))


def darken(rgb, f):
    return _f2rgb(_darken_f(rgb, f))


def blend(a, b, f):
    """`a` moved fraction f toward `b`."""
    return tuple(round(x + (y - x) * f) for x, y in zip(a, b))


def rotate_hue(rgb, deg):
    h, s, v = colorsys.rgb_to_hsv(*(c / 255 for c in rgb))
    r, g, b = colorsys.hsv_to_rgb((h + deg / 360) % 1.0, s, v)
    return (round(r * 255), round(g * 255), round(b * 255))


def _nudge(rgb, targets, toward_light, role="", bg_label=""):
    """Step lightness ±5% in float space (max 50 steps) until rgb clears every
    (other_color, min_ratio) contrast target.  If the loop exhausts its steps
    without meeting every target a [WARN] line is printed to stderr for each
    miss, but the best-effort colour is still returned (exit 0)."""
    rgb_f = tuple(float(c) for c in rgb)
    step_fn = _lighten_f if toward_light else _darken_f
    for _ in range(50):
        candidate = _f2rgb(rgb_f)
        if all(contrast_ratio(candidate, other) >= ratio
               for other, ratio in targets):
            return candidate
        rgb_f = step_fn(rgb_f, 0.05)
    # Loop exhausted — report each miss individually
    final = _f2rgb(rgb_f)
    for other, ratio in targets:
        actual = contrast_ratio(final, other)
        if actual < ratio:
            label = bg_label or "bg"
            print(
                f"[WARN] contrast target missed: {role} on {label}"
                f" = {actual:.2f}:1 (target {ratio})",
                file=sys.stderr,
            )
    return final


# ── Color extraction + dedupe (pure) ─────────────────────────────────────────

def extract_colors(text):
    """Counter of (r, g, b) → frequency from HTML/CSS text.

    Handles #RRGGBB, #RGB (expanded), and rgb()/rgba() integer forms."""
    found = Counter()
    for m in HEX_RX.finditer(text):
        h = m.group(1) or "".join(c * 2 for c in m.group(2))
        found[_rgb(h)] += 1
    for m in RGB_RX.finditer(text):
        vals = tuple(int(v) for v in m.groups())
        if all(v <= 255 for v in vals):
            found[vals] += 1
    return found


def dedupe(pairs, threshold=DEDUPE_DIST):
    """Merge near-identical colors (Euclidean distance < threshold), keeping
    the highest-frequency representative and summing frequencies.
    Returns [(rgb, freq)] sorted by descending frequency."""
    kept = []
    for rgb, freq in sorted(pairs, key=lambda p: -p[1]):
        for i, (rep, total) in enumerate(kept):
            if _dist(rgb, rep) < threshold:
                kept[i] = (rep, total + freq)
                break
        else:
            kept.append((tuple(rgb), freq))
    return sorted(kept, key=lambda p: -p[1])


# ── Palette assembly (pure) ──────────────────────────────────────────────────

def assemble_palette(candidates):
    """Classify [(rgb, freq)] into the nine required palette keys.

    Most-frequent very-dark vs very-light color decides the theme (tie →
    dark); accents are the top saturated, pairwise-distinct candidates with
    hue-rotation fallback; text colors are contrast-nudged to clear WCAG."""
    candidates = sorted(candidates, key=lambda p: -p[1])
    dark_pool = [p for p in candidates if _luminance(p[0]) < DARK_BG_LUM]
    light_pool = [p for p in candidates if _luminance(p[0]) > LIGHT_BG_LUM]
    best_dark = max(dark_pool, key=lambda p: p[1], default=None)
    best_light = max(light_pool, key=lambda p: p[1], default=None)
    if best_dark and (not best_light or best_dark[1] >= best_light[1]):
        dark, bg = True, best_dark[0]
    elif best_light:
        dark, bg = False, best_light[0]
    else:
        dark, bg = True, (16, 24, 32)  # no clear bg candidate → dark slate
    bg_deep = darken(bg, 0.08)
    surface = lighten(bg, 0.06) if dark else darken(bg, 0.04)

    pool = [rgb for rgb, _f in candidates
            if _sat(rgb) > ACCENT_SAT_MIN
            and ACCENT_LUM_MIN <= _luminance(rgb) <= ACCENT_LUM_MAX]
    accents = []
    for rgb in pool:
        if all(_dist(rgb, a) >= ACCENT_MIN_DIST for a in accents):
            accents.append(rgb)
        if len(accents) == 3:
            break
    if not accents:  # nothing saturated on the page → synthesize from bg hue
        h = colorsys.rgb_to_hsv(*(c / 255 for c in bg))[0]
        r, g, b = colorsys.hsv_to_rgb((h + 0.5) % 1.0, 0.55,
                                      0.78 if dark else 0.55)
        accents.append((round(r * 255), round(g * 255), round(b * 255)))
    while len(accents) < 3:
        accents.append(rotate_hue(accents[0], 30 if len(accents) == 1 else -30))

    text = (244, 246, 250) if dark else (17, 24, 39)
    text = _nudge(text, [(bg, 4.5), (surface, 4.5)], toward_light=dark,
                  role="text", bg_label="bg/surface")
    muted = _nudge(blend(text, bg, 0.35), [(bg, 4.5)], toward_light=dark,
                   role="text_muted", bg_label="bg")
    accents = [
        _nudge(a, [(bg, 3.0)], toward_light=dark,
               role=f"accent{i + 1}", bg_label="bg")
        for i, a in enumerate(accents)
    ]
    return {
        "bg": _hex6(bg), "bg_deep": _hex6(bg_deep), "surface": _hex6(surface),
        "accent1": _hex6(accents[0]), "accent2": _hex6(accents[1]),
        "accent3": _hex6(accents[2]),
        "text": _hex6(text), "text_muted": _hex6(muted), "dark": dark,
    }


# ── Network (module-level so tests can monkeypatch _http_get) ────────────────

def _http_get(url, headers=None, timeout=10):
    """GET url → (bytes, final_url). All network flows through here."""
    req = urllib.request.Request(url, headers={"User-Agent": UA,
                                               **(headers or {})})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return resp.read(), resp.geturl()


def fetch_brandfetch(domain, api_key):
    """Brandfetch v2 lookup → (candidates, fonts, logo_url).

    Color types are weighted loosely (brand > accent > dark > light) so the
    shared assembly ranks them like scrape frequencies."""
    raw, _ = _http_get(BRANDFETCH_URL.format(domain=domain),
                       headers={"Authorization": f"Bearer {api_key}"})
    data = json.loads(raw)
    weights = {"brand": 90, "accent": 80, "dark": 60, "light": 50}
    cands = []
    for i, c in enumerate(data.get("colors") or []):
        h = str(c.get("hex") or "").lstrip("#")
        if not re.fullmatch(r"[0-9a-fA-F]{6}", h):
            continue
        weight = weights.get(str(c.get("type") or "").lower(), 40)
        cands.append((_rgb(h), weight - i))
    fonts = [f["name"] for f in (data.get("fonts") or []) if f.get("name")]
    return dedupe(cands), fonts, _pick_logo(data.get("logos"))


def _pick_logo(logos):
    """Best PNG src from Brandfetch logos (type 'logo' preferred)."""
    best = None
    for logo in logos or []:
        score = 2 if str(logo.get("type") or "").lower() == "logo" else 1
        for fmt in logo.get("formats") or []:
            if str(fmt.get("format") or "").lower() == "png" and fmt.get("src"):
                if best is None or score > best[0]:
                    best = (score, fmt["src"])
    return best[1] if best else None


def _css_links(html, base_url):
    """Same-origin stylesheet URLs from <link rel="stylesheet"> tags."""
    origin = urllib.parse.urlparse(base_url).netloc
    links = []
    for tag in LINK_RX.findall(html):
        if "stylesheet" not in tag.lower():
            continue
        m = HREF_RX.search(tag)
        if not m:
            continue
        url = urllib.parse.urljoin(base_url, m.group(1))
        if urllib.parse.urlparse(url).netloc == origin and url not in links:
            links.append(url)
    return links


def scrape_colors(url):
    """(candidates, final_url) from homepage HTML plus up to 5 same-origin
    CSS files. Individual stylesheet failures are non-fatal."""
    raw, final = _http_get(url)
    final = final or url
    html = raw.decode("utf-8", "replace")
    counts = extract_colors(html)
    for css_url in _css_links(html, final)[:MAX_CSS_FILES]:
        try:
            css_raw, _ = _http_get(css_url)
        except (OSError, ValueError) as e:
            print(f"  [WARN] stylesheet {css_url}: {e}", file=sys.stderr)
            continue
        counts.update(extract_colors(css_raw.decode("utf-8", "replace")))
    return dedupe(counts.items()), final


# ── Orchestration ────────────────────────────────────────────────────────────

def _normalize(target):
    """domain-or-url → (bare domain, homepage url)."""
    t = target.strip()
    if "://" not in t:
        t = "https://" + t
    parsed = urllib.parse.urlparse(t)
    domain = parsed.netloc or parsed.path.split("/")[0]
    return domain, f"{parsed.scheme}://{domain}"


def _print_swatches(pal):
    bg = _rgb(pal["bg"])
    print(f"  {'role':<11} {'hex':<8} vs bg")
    for role in ROLES:
        ratio = contrast_ratio(_rgb(pal[role]), bg)
        print(f"  {role:<11} #{pal[role]}  {ratio:.2f}:1")


def _download_logo(url, dest):
    """Fetch logo PNG to dest — failure is non-fatal by design."""
    try:
        data, _ = _http_get(url)
        if not data.startswith(b"\x89PNG"):
            raise ValueError("response is not a PNG")
        dest.parent.mkdir(parents=True, exist_ok=True)
        dest.write_bytes(data)
        print(f"  logo → {dest}")
    except (OSError, ValueError) as e:
        print(f"  [WARN] logo download failed ({e}) — continuing without it",
              file=sys.stderr)


def run(target, name, assets_dir, profile=False):
    """Full pipeline; returns process exit code (0 ok, 1 no colors).

    With `profile=True` an engagement profile pointing at the new palette (and
    the downloaded logo, when there is one) is written alongside it.
    """
    domain, home_url = _normalize(target)
    if not domain:
        print(f"ERROR: cannot parse a domain from {target!r}", file=sys.stderr)
        return 1
    candidates, fonts, logo_url, source = None, [], None, None
    api_key = os.environ.get("BRANDFETCH_API_KEY")
    if api_key:
        try:
            candidates, fonts, logo_url = fetch_brandfetch(domain, api_key)
            source = "brandfetch"
        except (OSError, ValueError, KeyError, TypeError, AttributeError) as e:
            print(f"  [WARN] Brandfetch lookup failed ({e}) — falling back "
                  "to CSS scrape", file=sys.stderr)
            candidates = None
    if not candidates:
        try:
            candidates, source = scrape_colors(home_url)
        except (OSError, ValueError) as e:
            print(f"ERROR: could not fetch brand colors for {domain}: {e}",
                  file=sys.stderr)
            return 1
    if not candidates:
        print(f"ERROR: no colors found at {source}", file=sys.stderr)
        return 1
    if fonts:
        print(f"  brand fonts (informational — palette keeps web-safe "
              f"stacks): {', '.join(fonts)}")
    pal = assemble_palette(candidates)
    pal["_source"] = source

    # Validate through the real loader before touching <assets>/palettes/.
    with tempfile.TemporaryDirectory() as td:
        (Path(td) / f"{name}.json").write_text(
            json.dumps(pal, indent=2) + "\n", encoding="utf-8")
        if name not in load_custom_palettes(td):
            print("ERROR: generated palette failed validation — nothing "
                  "written", file=sys.stderr)
            return 1
    out = Path(assets_dir) / "palettes" / f"{name}.json"
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(pal, indent=2) + "\n", encoding="utf-8")
    print(f"  palette '{name}' ({'dark' if pal['dark'] else 'light'}, "
          f"source: {source}) → {out}")
    _print_swatches(pal)
    logo_path = ""
    if logo_url:
        dest = Path(assets_dir) / "brand" / f"{name}-logo.png"
        _download_logo(logo_url, dest)
        if dest.is_file():
            logo_path = str(dest)
    if profile:
        prof = profiles.write_profile(name, assets_dir, {
            "palette": name, "logo": logo_path, "author": "",
            "number_format": "units=M decimals=1"})
        print(f"  profile '{name}' → {prof}")
    return 0


def main():
    parser = argparse.ArgumentParser(
        description="Generate a custom deck palette from a brand's web "
                    "presence (Brandfetch API or CSS scrape).")
    parser.add_argument("domain", help="Company domain or homepage URL")
    parser.add_argument("--name", required=True,
                        help="Palette name → <assets>/palettes/<name>.json")
    parser.add_argument("--assets-dir", default="assets",
                        help="Assets directory (default: assets)")
    parser.add_argument("--profile", action="store_true",
                        help="Also write <assets>/profiles/<name>.json "
                             "wiring the new palette (and logo) into an "
                             "engagement profile")
    args = parser.parse_args()
    try:
        args.name = profiles.safe_name(args.name)
    except ValueError:
        parser.error("--name must be alphanumeric (plus interior - and _)")
    sys.exit(run(args.domain, args.name, args.assets_dir, args.profile))


if __name__ == "__main__":
    main()
