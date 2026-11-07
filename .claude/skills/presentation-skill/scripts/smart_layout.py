"""Auto-select layout when outline omits **Layout:**."""
from helpers import parse_visual

# Pairs the content signals cannot separate on their own: both members are
# defensible for the same slide, so a client profile's `preferred_layouts`
# breaks the tie (see profiles.py).
AMBIGUITY_PAIRS = (("bullet-list", "exec-summary"),
                   ("cards-3", "cards-4"),
                   ("table", "heatmap-table"))
CARDS_3_MAX = 3         # cards-3 holds up to three cards
CARDS_4_MIN = 4         # cards-4 wants a full row of four


def auto_layout(slide, preferred=None):
    """Pick a layout from slide content keys. Returns layout name string.

    `preferred` (a profile's preferred layouts) only breaks ties: it is used
    when the content-driven pick sits in an ambiguity pair and the profile
    prefers the pair's other member.
    """
    if slide.get("layout"):
        return slide["layout"]
    return _prefer(slide, _auto_pick(slide), preferred)[0]


def _prefer(slide, pick, preferred):
    """(layout, reason) — apply a profile preference the content can hold.

    `bullet-list` is also `_auto_pick`'s catch-all, so a blanket preference for
    `exec-summary` would promote content-heavy slides into a layout that
    truncates them (`exec_max` bullets vs `bullet_max`).  The capacity guard
    below keeps the content-driven pick in that case and returns a reason the
    build reports as a warning.
    """
    for pair in AMBIGUITY_PAIRS:
        if pick not in pair:
            continue
        for want in preferred or ():
            if want not in pair or want == pick:
                continue
            fits, detail = _fits(slide, want)
            if fits:
                return want, None
            return pick, f"profile prefers {want} but {detail} — kept {pick}"
    return pick, None


def _fits(slide, layout):
    """(fits, detail) — can this slide's content live in `layout`?"""
    if layout == "exec-summary":
        import builders
        n = len(slide.get("bullets") or [])
        cap = builders.density()["exec_max"]
        return n <= cap, f"{n} bullets exceed its cap ({cap})"
    if layout in ("cards-3", "cards-4"):
        n = len(slide.get("cards") or [])
        if layout == "cards-3":
            return n <= CARDS_3_MAX, (f"{n} cards exceed its cap "
                                      f"({CARDS_3_MAX})")
        return n >= CARDS_4_MIN, (f"the slide has {n} cards "
                                  f"({CARDS_4_MIN} expected)")
    return True, ""       # table <-> heatmap-table hold the same rows


def repick(slides, preferred):
    """Re-pick auto-selected layouts from a profile. Returns warnings.

    The profile is only known after the outline is parsed, so slides marked
    `_auto_layout` get a second pass once `preferred_layouts` is available.
    Slides with an explicit **Layout:** are never touched, and a preference the
    slide's content cannot hold is skipped with a warning rather than applied.
    """
    warnings = []
    if not preferred:
        return warnings
    for n, slide in enumerate(slides, 1):
        if not slide.get("_auto_layout"):
            continue
        layout, reason = _prefer(slide, slide.get("layout"), preferred)
        slide["layout"] = layout
        if reason:
            warnings.append(f"Slide {n}: {reason}")
    return warnings


def _auto_pick(slide):
    """The content-driven pick, before any profile preference applies."""
    # Consulting/structural signatures — unambiguous content keys first.
    if slide.get("layers"):
        return "architecture-layers"
    if slide.get("lanes"):
        return "swimlane-process"
    if slide.get("matrix_items"):
        return "matrix-2x2"
    if slide.get("tiles"):
        return "dashboard"
    bars = slide.get("bars", [])
    if slide.get("milestones") or any(
            "start" in b or "end" in b for b in bars):
        return "gantt"
    if bars and all("size" in b and "value" in b for b in bars):
        # Size/Value bars (no Start/End) are bar-mekko-shaped, not
        # gantt-shaped -- see references/generation-guide.md's
        # auto-selection notes.
        return "bar-mekko"
    if slide.get("nodes"):
        return "driver-tree"   # org-chart must be requested explicitly
    if slide.get("quote"):
        return "quote-evidence"
    if slide.get("value") and not slide.get("data"):
        return "big-number"
    if any(isinstance(v, str) and v in ("total", "subtotal")
           for _, v in slide.get("data", [])):
        return "waterfall"
    if slide.get("steps") and not slide.get("visual"):
        return "process-flow"

    if slide.get("stats"):
        return "stat-callout"
    if slide.get("cards"):
        n = len(slide["cards"])
        if n >= 4:
            return "cards-4"
        if n >= 2:
            return "cards-3"
    if slide.get("table_rows"):
        return "table"
    if slide.get("items") or any(":" in b for b in slide.get("bullets", [])[:3]):
        if slide.get("items") or sum(1 for b in slide.get("bullets", [])
                                     if re_date_bullet(b)) >= 2:
            return "timeline"
    if slide.get("visual_left") or slide.get("visual_right"):
        return "comparison"
    if slide.get("left_bullets") or slide.get("right_bullets"):
        return "comparison"

    kind, _ = parse_visual(slide.get("visual", ""))
    if kind == "image" and slide.get("caption"):
        return "full-image"
    if kind in ("chart", "image"):
        return "two-column-split"

    bullets = slide.get("bullets", [])
    if any(b.lower().startswith("point ") for b in bullets) or len(bullets) <= 4:
        if len(bullets) >= 3 and all(len(b) > 40 for b in bullets[:3]):
            return "exec-summary"

    heading = (slide.get("heading") or slide.get("title") or "").lower()
    if slide.get("contact") or "thank" in heading or "questions" in heading:
        return "closing"
    if slide.get("title") and slide.get("subtitle") and not bullets:
        return "title"
    if not bullets and slide.get("subtitle"):
        return "section-divider"

    return "bullet-list"


def re_date_bullet(text):
    import re
    return bool(re.match(r"^Q[1-4]\s+\d{4}|^\d{4}|^[A-Z][a-z]{2}\s+\d{4}", text.strip()))
