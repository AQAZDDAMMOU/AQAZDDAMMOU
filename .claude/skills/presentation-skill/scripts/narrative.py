"""Purpose-aware narrative templates, validation, and appendix generation.

House-style lints (word lists, action-title hygiene, typography slips) live in
`house_style.py` and are re-exported here so `build_deck.validate()` has one
import site for every editorial check.
"""
import re

from house_style import check_house_style, load_style, resolve_style

__all__ = [
    "PURPOSE_ARCS", "STAT_RX", "SOURCE_RX", "normalize_purpose",
    "validate_narrative", "check_unsourced_stats", "generate_appendix_outline",
    "check_house_style", "load_style", "resolve_style",
]

# Expected layout sequences per presentation purpose (golden-thread guides)
PURPOSE_ARCS = {
    "pitch": {
        "label": "Pitch deck",
        "max_slides": 15,
        "recommended": [
            "title", "bullet-list", "bullet-list", "two-column-split",
            "stat-callout", "two-column-split", "stat-callout", "comparison",
            "cards-3", "timeline", "table", "stat-callout", "closing",
        ],
        "must_include": ("stat-callout", "closing"),
    },
    "strategy": {
        "label": "Internal strategy",
        "max_slides": 20,
        "recommended": [
            "title", "exec-summary", "section-divider", "bullet-list",
            "two-column-split", "comparison", "timeline", "stat-callout", "closing",
        ],
        "must_include": ("exec-summary", "closing"),
    },
    "client": {
        "label": "Client presentation",
        "max_slides": 18,
        "recommended": [
            "title", "exec-summary", "two-column-split", "stat-callout",
            "full-image", "timeline", "closing",
        ],
        "must_include": ("closing",),
    },
    "talk": {
        "label": "Conference talk",
        "max_slides": 25,
        "recommended": [
            "title", "section-divider", "bullet-list", "full-image",
            "stat-callout", "timeline", "closing",
        ],
        "must_include": ("closing",),
    },
    "tutorial": {
        "label": "Tutorial",
        "max_slides": 30,
        "recommended": [
            "title", "bullet-list", "two-column-split", "full-image",
            "bullet-list", "comparison", "timeline", "closing",
        ],
        "must_include": ("closing",),
    },
}

STAT_RX = re.compile(r"\b\d+(?:\.\d+)?%|\$[\d,.]+[BMK]?|\b\d{2,}\+?\b")
SOURCE_RX = re.compile(r"source:|citation:|https?://", re.I)


def normalize_purpose(raw):
    if not raw:
        return None
    key = raw.strip().lower().replace(" ", "-")
    aliases = {
        "pitch-deck": "pitch", "venture": "pitch", "fundraising": "pitch",
        "internal": "strategy", "internal-strategy": "strategy",
        "client-presentation": "client", "conference": "talk",
        "conference-talk": "talk", "workshop": "tutorial",
    }
    return aliases.get(key, key if key in PURPOSE_ARCS else None)


def validate_narrative(meta, slides):
    """Return warnings for purpose-aware narrative checks."""
    warnings = []
    purpose = normalize_purpose(meta.get("purpose", ""))
    if not purpose:
        return warnings

    arc = PURPOSE_ARCS[purpose]
    layouts = [p.get("layout", "bullet-list") for p in slides]

    if len(slides) > arc["max_slides"]:
        warnings.append(
            f"Purpose '{purpose}': {len(slides)} slides exceeds recommended "
            f"max {arc['max_slides']} for {arc['label']}")

    # consulting variants satisfy their classic counterparts
    equivalents = {"exec-summary": {"exec-summary-scqa"},
                   "closing": {"next-steps"}}
    for required in arc["must_include"]:
        accepted = {required} | equivalents.get(required, set())
        if not accepted & set(layouts):
            warnings.append(
                f"Purpose '{purpose}': missing recommended '{required}' layout")

    # Therefore test: flag 3+ consecutive identical non-structural layouts
    structural = {"title", "closing", "section-divider", "full-image"}
    for i in range(len(layouts) - 2):
        a, b, c = layouts[i], layouts[i + 1], layouts[i + 2]
        if a == b == c and a not in structural:
            warnings.append(
                f"Narrative: slides {i + 1}-{i + 3} all use '{a}' — "
                "vary layout to maintain golden-thread momentum")

    if not meta.get("takeaway"):
        warnings.append(
            f"Purpose '{purpose}': add deck **Takeaway:** one-sentence "
            "audience memory hook in front-matter")

    return warnings


def check_unsourced_stats(slides):
    """Warn on slides with big numbers but no Source: field or URL in notes."""
    warnings = []
    for n, p in enumerate(slides, 1):
        if p.get("_auto"):  # synthetic agenda/references slides cite nothing new
            continue
        blob = " ".join([
            p.get("heading", ""), p.get("title", ""), p.get("notes", ""),
            " ".join(p.get("bullets", [])),
            " ".join(str(s) for s in p.get("stats", [])),
        ])
        if STAT_RX.search(blob) and not SOURCE_RX.search(blob):
            if p.get("source"):
                continue
            warnings.append(
                f"Slide {n}: contains statistics but no '- Source:' or "
                "citation in notes — verify numbers before presenting")
    return warnings


def generate_appendix_outline(meta, slides):
    """Return markdown for a suggested appendix (not built automatically)."""
    purpose = normalize_purpose(meta.get("purpose", ""))
    title = meta.get("takeaway") or "Appendix"
    lines = [
        f"# Appendix: {title}",
        "",
        f"**Purpose:** {purpose or 'general'} — detail slides kept out of main deck.",
        "",
    ]
    if purpose == "pitch":
        sections = [
            ("Financial Detail", "table", "Full P&L / unit economics"),
            ("Team Bios", "cards-3", "Extended team backgrounds"),
            ("Technical Architecture", "two-column-split", "Deep-dive diagram"),
            ("Competitive Matrix", "table", "Feature comparison"),
        ]
    elif purpose == "strategy":
        sections = [
            ("Option Analysis", "comparison", "Alternatives considered"),
            ("Risk Register", "table", "Risks and mitigations"),
            ("Detailed Roadmap", "timeline", "Full timeline with owners"),
        ]
    else:
        sections = [
            ("Supporting Data", "table", "Backup metrics"),
            ("References", "bullet-list", "Sources and citations"),
        ]

    for i, (heading, layout, note) in enumerate(sections, 1):
        lines.extend([
            f"## Slide A{i}: {heading}",
            f"**Layout:** {layout}",
            f"- Notes: \"{note}\"",
            "",
        ])
    return "\n".join(lines)
