"""Color palettes and per-palette font stacks for presentation-deck skill.

Design rules encoded here (per the 60-30-10 dominance rule): bg/bg_deep/surface
carry ~60-70% of visual weight, accent1 is THE accent (~10%), accent2/3 support.
Muted text colors are chosen to clear WCAG 4.5:1 against `bg`.
"""
import re

from pptx.dml.color import RGBColor

PALETTES = {
    # ── Dark themes ──────────────────────────────────────────────────────────
    "midnight-executive": {  # finance, strategy, luxury
        "bg": "0A0F1E", "bg_deep": "060912", "surface": "1A2035",
        "accent1": "C9A84C", "accent2": "E8D5B5", "accent3": "8B6914",
        "text": "F1F5F9", "text_muted": "94A3B8",
        "dark": True,
        "font_title": "Gill Sans MT", "font_body": "Calibri", "font_label": "Calibri Light",
    },
    "aurora": {  # AI, tech, innovation
        "bg": "0F0A2A", "bg_deep": "080517", "surface": "1E1040",
        "accent1": "7C3AED", "accent2": "06B6D4", "accent3": "A78BFA",
        "text": "F8FAFC", "text_muted": "CBD5E1",
        "dark": True,
        "font_title": "Gill Sans MT", "font_body": "Calibri", "font_label": "Calibri Light",
    },
    "venture-pitch": {  # startups, product launches
        "bg": "18181B", "bg_deep": "0A0A0A", "surface": "27272A",
        "accent1": "F97316", "accent2": "14B8A6", "accent3": "FEF3C7",
        "text": "FAFAFA", "text_muted": "A1A1AA",
        "dark": True,
        "font_title": "Trebuchet MS", "font_body": "Calibri", "font_label": "Calibri Light",
    },
    "forest": {  # sustainability, ESG
        "bg": "0D1B0E", "bg_deep": "080F08", "surface": "1A2E1C",
        "accent1": "22C55E", "accent2": "A3B18A", "accent3": "8B6914",
        "text": "F0FDF4", "text_muted": "86EFAC",
        "dark": True,
        "font_title": "Gill Sans MT", "font_body": "Calibri", "font_label": "Calibri Light",
    },
    "teal-trust": {  # healthcare, fintech trust, services
        "bg": "0B2B29", "bg_deep": "06201E", "surface": "14403C",
        "accent1": "14B8A6", "accent2": "99F6E4", "accent3": "5EEAD4",
        "text": "F0FDFA", "text_muted": "8CCFC6",
        "dark": True,
        "font_title": "Gill Sans MT", "font_body": "Calibri", "font_label": "Calibri Light",
        # rag_bad override: default C0504D gives only 2.46:1 on this surface;
        # D95F5B clears 3.0:1 (3.14:1) for 12pt text on driver-tree boxes.
        "rag_bad": "D95F5B",
    },
    "charcoal-minimal": {  # design, architecture, monochrome minimal
        "bg": "1C1C1E", "bg_deep": "111113", "surface": "2A2A2E",
        "accent1": "FAFAF9", "accent2": "A8A29E", "accent3": "78716C",
        "text": "F5F5F4", "text_muted": "A1A1AA",
        "dark": True,
        "font_title": "Arial Black", "font_body": "Arial", "font_label": "Arial",
    },
    "ocean-gradient": {  # data, cloud, maritime, deep tech
        "bg": "082F49", "bg_deep": "041C30", "surface": "0D4060",
        "accent1": "06B6D4", "accent2": "38BDF8", "accent3": "7DD3FC",
        "text": "F0F9FF", "text_muted": "94C6E0",
        "dark": True,
        "font_title": "Gill Sans MT", "font_body": "Calibri", "font_label": "Calibri Light",
        # rag_bad override: default C0504D gives only 2.34:1 on this surface;
        # D86A68 clears 3.0:1 (3.22:1) for 12pt text on driver-tree boxes.
        "rag_bad": "D86A68",
    },
    # ── Light themes (backgrounds intentionally never pure #FFFFFF) ─────────
    "swiss-light": {  # corporate, education, tutorials
        "bg": "FAF8F5", "bg_deep": "F0EDE8", "surface": "FFFFFF",
        "accent1": "2563EB", "accent2": "F97066", "accent3": "D4A574",
        "text": "0F172A", "text_muted": "57657B",
        "dark": False,
        "font_title": "Palatino Linotype", "font_body": "Calibri", "font_label": "Calibri Light",
    },
    "warm-terracotta": {  # consumer, hospitality, editorial warmth
        "bg": "FAF3EC", "bg_deep": "F1E6DA", "surface": "FFFFFF",
        "accent1": "C2552C", "accent2": "8A9B6E", "accent3": "D9A87E",
        "text": "2D2A26", "text_muted": "6E6258",
        "dark": False,
        "font_title": "Georgia", "font_body": "Calibri", "font_label": "Calibri Light",
    },
    "berry-cream": {  # lifestyle, brand, creative
        "bg": "FBF6F0", "bg_deep": "F3E9E0", "surface": "FFFFFF",
        "accent1": "8E2A4F", "accent2": "C98A9E", "accent3": "B5838D",
        "text": "33222B", "text_muted": "75606B",
        "dark": False,
        "font_title": "Georgia", "font_body": "Calibri", "font_label": "Calibri Light",
    },
}

DEFAULT_PALETTE = "midnight-executive"

FONT_DEFAULTS = {"font_title": "Gill Sans MT", "font_body": "Calibri",
                 "font_label": "Calibri Light"}
REQUIRED_KEYS = {"bg", "bg_deep", "surface", "accent1", "accent2", "accent3",
                 "text", "text_muted", "dark"}

# Bounded type scale: every font size (pt) the toolchain's builders emit on
# the 16:9 design canvas. Palettes may override with their own ladder;
# pptx_lint warns on sizes outside it (like the color whitelist, but
# advisory — foreign decks and 4:3 canvas rescales legitimately differ).
DEFAULT_TYPE_SCALE = (10, 11, 12, 13, 14, 15, 16, 17, 18, 20, 22, 24,
                      30, 32, 34, 44, 46, 48, 52, 54, 110)


def _fill_defaults(pal):
    pal.setdefault("motif", "icon-circle")
    pal.setdefault("type_scale", list(DEFAULT_TYPE_SCALE))
    for k, v in FONT_DEFAULTS.items():
        pal.setdefault(k, v)
    pal.setdefault("chart_series",
                   [pal["accent1"], pal["accent2"], pal["accent3"]])
    # RAG status fills (heatmap-table '- Scale: rag'); deeper tones on light
    # themes so the chips read as filled against white surfaces. Cell text
    # color is picked per-fill by luminance, so both sets stay readable.
    dark = bool(pal.get("dark"))
    pal.setdefault("rag_bad", "C0504D" if dark else "B3453F")
    pal.setdefault("rag_mid", "E8A33D" if dark else "C97E1F")
    pal.setdefault("rag_good", "4F9D69" if dark else "3D8456")
    return pal


for _key, _pal in PALETTES.items():
    _fill_defaults(_pal)


def load_custom_palettes(palettes_dir):
    """Merge user palette JSONs (one palette per <name>.json) into PALETTES.

    Validates the required token schema; invalid files are skipped with a
    warning (never trust external data). Returns the list of loaded names.
    """
    import json
    import sys
    from pathlib import Path
    palettes_dir = Path(palettes_dir)
    loaded = []
    if not palettes_dir.is_dir():
        return loaded
    for f in sorted(palettes_dir.glob("*.json")):
        try:
            pal = json.loads(f.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as e:
            print(f"  [WARN] palette {f.name}: unreadable ({e})", file=sys.stderr)
            continue
        missing = REQUIRED_KEYS - set(pal)
        if missing or not isinstance(pal.get("dark"), bool):
            print(f"  [WARN] palette {f.name}: missing keys "
                  f"{sorted(missing)} — skipped", file=sys.stderr)
            continue
        hex_keys = REQUIRED_KEYS - {"dark"}
        bad = [k for k in sorted(hex_keys)
               if not re.fullmatch(r"[0-9A-Fa-f]{6}", str(pal.get(k, "")))]
        series = pal.get("chart_series")
        if series is not None:
            if len(series) < 3:
                bad.append("chart_series: too few colors")
            bad += [f"chart_series[{i}]" for i, c in enumerate(series)
                    if not re.fullmatch(r"[0-9A-Fa-f]{6}", str(c))]
        bad += [k for k in ("rag_bad", "rag_mid", "rag_good")
                if k in pal and not re.fullmatch(r"[0-9A-Fa-f]{6}",
                                                 str(pal[k]))]
        if bad:
            print(f"  [WARN] palette {f.name}: non-hex color values "
                  f"{bad} — skipped", file=sys.stderr)
            continue
        scale = pal.get("type_scale")
        if scale is not None and (
                not isinstance(scale, list) or not scale
                or not all(isinstance(s, (int, float)) and s > 0
                           for s in scale)):
            print(f"  [WARN] palette {f.name}: type_scale must be a "
                  "non-empty list of positive numbers — skipped",
                  file=sys.stderr)
            continue
        PALETTES[f.stem] = _fill_defaults(pal)
        loaded.append(f.stem)
    return loaded


VARIANT_PRESETS = {
    "a": {},
    "b": {"palette": "aurora", "density": "comfortable"},
    "c": {"palette": "swiss-light", "density": "comfortable"},
    "consulting": {"palette": "midnight-executive", "density": "compact"},
}


def apply_variant(variant_key, palette_key=None, density=None):
    preset = VARIANT_PRESETS.get((variant_key or "").lower(), {})
    if palette_key is None and preset.get("palette"):
        palette_key = preset["palette"]
    if density is None and preset.get("density"):
        density = preset["density"]
    return palette_key, density


def hex_rgb(h):
    h = h.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


# ── CJK font overlay ─────────────────────────────────────────────────────────
# Deck-wide swap when the outline contains CJK text (build() decides via
# set_cjk); the chosen font replaces font_title/font_body/font_label so
# headings and body render without tofu boxes.
CJK_FONT_STACK = ("PingFang SC", "Hiragino Sans GB", "Noto Sans CJK SC",
                  "Microsoft YaHei")
_CJK = {"on": False, "font": None, "probed": False}


def _pick_cjk_font():
    """First available font from CJK_FONT_STACK (matplotlib font index when
    importable, else a platform guess). Returns None when none is found."""
    try:
        from matplotlib import font_manager
    except ImportError:
        import sys
        return "PingFang SC" if sys.platform == "darwin" else "Noto Sans CJK SC"
    names = {f.name for f in font_manager.fontManager.ttflist}
    for name in CJK_FONT_STACK:
        if name in names:
            return name
    return None


def set_cjk(on):
    """Toggle the deck-wide CJK font overlay (reset to False per build).

    Passing False also resets probe state so warn-once and font detection
    cannot leak across back-to-back builds in the same process.
    """
    import sys
    _CJK["on"] = bool(on)
    if not on:
        _CJK["probed"] = False
        _CJK["font"] = None
        return
    if not _CJK["probed"]:
        _CJK["font"] = _pick_cjk_font()
        _CJK["probed"] = True
    if _CJK["font"]:
        print(f"  [WARN] CJK text detected — using '{_CJK['font']}' for "
              "titles, body, and labels", file=sys.stderr)
    else:
        print("  [WARN] CJK text detected but no CJK-safe font found "
              f"(looked for: {', '.join(CJK_FONT_STACK)}) — keeping palette "
              "fonts", file=sys.stderr)


def get_palette(key):
    pal = PALETTES.get(key, PALETTES[DEFAULT_PALETTE])
    if _CJK["on"] and _CJK["font"]:
        pal = {**pal, "font_title": _CJK["font"], "font_body": _CJK["font"],
               "font_label": _CJK["font"]}
    return pal
