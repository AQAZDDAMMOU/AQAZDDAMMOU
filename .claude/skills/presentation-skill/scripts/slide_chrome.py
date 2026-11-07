#!/usr/bin/env python3
"""
slide_chrome.py — per-slide chrome drawn on top of a built layout.

Footer + page number (a live `slidenum` field with cached text), section
tracker (plain label or tab strip), sticker, kicker band, source/exhibit
footnote (normalised via footnotes.normalize_source), footnote block, file
stamp and status stamp, plus `deck_title()` for docProps/core.xml and the
`**Page-Start:**` pagination helpers. Each adder returns a list of
`(shape, role)` pairs; build_deck.build() collects them so QA can tell
chrome from content.

Geometry is design-space inches on the 13.33 x 7.5 canvas (builders rescale
for other slide sizes). Split out of build_deck.py; build_deck re-exports
these under their old `_add_*` / `_deck_title` names.
"""
from copy import deepcopy
from pathlib import Path

from lxml import etree
from pptx.oxml.ns import qn

from helpers import warn

# Fixed GUID for the slide-number field: PowerPoint only needs it unique
# within a paragraph, and a constant keeps builds deterministic.
SLIDENUM_FLD_ID = "{53DEA387-2E3B-4C6F-9F3A-4B7C1D2E5A01}"


def add_slidenum_field(tb, cached):
    """Rewrite the numeric tail of a page-number textbox as a live
    `<a:fld type="slidenum">` carrying the same cached text.

    A literal prefix before the number (the appendix `B·`) stays its own
    run; the field inherits the run's formatting.  PowerPoint renumbers the
    field on reorder/`firstSlideNum`; every lint keeps reading the cached
    string until then."""
    p = tb.text_frame.paragraphs[0]._p
    run = p.findall(qn("a:r"))[-1]
    t = run.find(qn("a:t"))
    prefix = t.text[:-len(cached)] if (t.text or "").endswith(cached) else ""
    fld = etree.SubElement(p, qn("a:fld"))
    fld.set("id", SLIDENUM_FLD_ID)
    fld.set("type", "slidenum")
    rPr = run.find(qn("a:rPr"))
    if rPr is not None:
        fld.append(deepcopy(rPr))
    fld_t = etree.SubElement(fld, qn("a:t"))
    fld_t.text = cached
    p.insert(list(p).index(run) + 1, fld)
    if prefix:
        t.text = prefix
    else:
        p.remove(run)
    return tb


def add_footer(slide, i, pal, footer, page_numbers, appendix=False,
               page_start=1):
    import builders
    shapes = []
    if footer:
        shapes.append((builders.add_tb(slide, footer, 0.7, 7.08, 6.0, 0.35,
                        size=11, color=pal["text_muted"],
                        font=pal["font_label"]), "footer"))
    if page_numbers:
        from pptx.enum.text import PP_ALIGN
        num = str(i + page_start)
        label = f"B·{num}" if appendix else num
        tb = builders.add_tb(slide, label, 11.9, 7.08, 0.8, 0.35,
                        size=11, color=pal["text_muted"],
                        font=pal["font_label"], align=PP_ALIGN.RIGHT)
        add_slidenum_field(tb, num)
        shapes.append((tb, "pagenum"))
    return shapes


def parse_page_start(meta):
    """`**Page-Start:** N` → int (default 1; malformed → warn, 1)."""
    raw = str(meta.get("page_start", "") or "").strip()
    if not raw:
        return 1
    try:
        value = int(raw)
        if value < 1:
            raise ValueError
        return value
    except ValueError:
        warn(f"**Page-Start:** {raw!r} is not a positive number — using 1")
        return 1


def set_first_slide_num(prs, page_start):
    """Native pagination offset: `p:presentation@firstSlideNum`, so slide
    number fields renumber to match the cached texts."""
    prs.element.set("firstSlideNum", str(page_start))


def add_section_tracker(slide, pal, section):
    """Section tracker: small current-section label top-right (consulting
    navigation convention for long decks)."""
    if not section:
        return []
    import builders
    from pptx.enum.text import PP_ALIGN
    shape = builders.add_tb(slide, section.upper(), 7.6, 0.12, 5.0, 0.32,
                            size=11, bold=True, color=pal["text_muted"],
                            font=pal["font_label"], align=PP_ALIGN.RIGHT)
    return [(shape, "tracker")]


# Tab strip geometry (design-space inches; builders rescale for 4:3 etc.)
TAB_RIGHT_EDGE = 12.63
TAB_Y, TAB_H, TAB_GAP = 0.12, 0.30, 0.08
TAB_STRIP_MAX_W = 5.5   # over this (or 6+ sections) -> compact text form
TAB_LABEL_CHARS = 14


def _tab_label(name):
    return name[:TAB_LABEL_CHARS] + "…" if len(name) > TAB_LABEL_CHARS else name


def add_tracker_tabs(slide, pal, sections, current, state):
    """**Tracker:** tabs — chip-per-section strip top-right; current section
    chip filled accent1, others outlined muted. Too many sections -> compact
    'n/N · Section' text at the plain-label position (warn once per build)."""
    import builders
    from pptx.enum.text import PP_ALIGN
    labels = [_tab_label(s) for s in sections]
    widths = [max(0.7, 0.09 * len(lb) + 0.25) for lb in labels]
    total = sum(widths) + TAB_GAP * (len(widths) - 1)
    if len(sections) >= 6 or total > TAB_STRIP_MAX_W:
        if not state.get("warned"):
            warn(f"Tracker tabs: {len(sections)} sections exceed tab budget "
                 "— using compact text form")
            state["warned"] = True
        n = sections.index(current) + 1 if current in sections else 0
        shape = builders.add_tb(slide, f"{n}/{len(sections)} · {current}",
                        7.6, 0.12, 5.0, 0.32, size=11, bold=True,
                        color=pal["text_muted"], font=pal["font_label"],
                        align=PP_ALIGN.RIGHT)
        return [(shape, "tracker")]
    shapes, x = [], TAB_RIGHT_EDGE
    for label, w, sec in reversed(list(zip(labels, widths, sections))):
        x -= w
        if sec == current:
            shapes.append((builders.add_round_rect(
                slide, x, TAB_Y, w, TAB_H, pal["accent1"]), "tracker"))
            shapes.append((builders.add_tb(slide, label, x, TAB_Y + 0.02, w,
                            TAB_H - 0.02, size=10, bold=True, color=pal["bg"],
                            font=pal["font_label"], align=PP_ALIGN.CENTER),
                           "tracker"))
        else:
            chip = builders.add_round_rect(slide, x, TAB_Y, w, TAB_H,
                                           pal["bg"],
                                           line_hex=pal["text_muted"],
                                           line_pt=0.75)
            chip.fill.background()  # outline only — works on any background
            shapes.append((chip, "tracker"))
            shapes.append((builders.add_tb(slide, label, x, TAB_Y + 0.02, w,
                            TAB_H - 0.02, size=10, color=pal["text_muted"],
                            font=pal["font_label"], align=PP_ALIGN.CENTER),
                           "tracker"))
        x -= TAB_GAP
    return shapes


def add_sticker(slide, pal, text, below_section=False):
    """Per-slide bordered status tag (ILLUSTRATIVE, PRELIMINARY...) top-right;
    drops below the section tracker label when both are present."""
    import builders
    from pptx.enum.text import PP_ALIGN
    label = text.upper()
    w = min(3.0, max(1.2, 0.12 * len(label)))
    left = 12.63 - w
    top = 0.50 if below_section else 0.14
    box = builders.add_rect(slide, left, top, w, 0.32, pal["bg"],
                            line_hex=pal["accent3"], line_pt=1.0)
    box.fill.background()  # outline only — works on photos and gradients
    txt = builders.add_tb(slide, label, left, top + 0.02, w, 0.28, size=11,
                    bold=True, color=pal["text_muted"], font=pal["font_label"],
                    align=PP_ALIGN.CENTER)
    return [(box, "sticker"), (txt, "sticker")]


def add_kicker_box(slide, pal, text):
    """Takeaway box: full-width accent-edged band above the source/footer."""
    import builders
    bg = builders.add_rect(slide, 0.7, 6.18, 11.93, 0.52, pal["surface"])
    bar = builders.add_rect(slide, 0.7, 6.18, 0.06, 0.52, pal["accent1"])
    txt = builders.add_tb(slide, text, 0.95, 6.25, 11.4, 0.4, size=14,
                          bold=True, color=pal["text"], font=pal["font_body"])
    return [(bg, "kicker"), (bar, "kicker"), (txt, "kicker")]


def add_source(slide, pal, source, exhibit_no=None, note=None):
    """Exhibit attribution footnote (every consulting exhibit cites a source).
    With **Exhibits:** on, sourced slides are numbered 'Exhibit N · Source: ...'
    and the shape carries role `exhibit-badge` instead of plain `source`.
    The text goes through footnotes.normalize_source (prefix spelling, entry
    dedupe, `- Note:` prefixed MBB-style; verbatim decks keep source bytes
    exact) — the line is drawn at design-y 6.80."""
    import builders
    from footnotes import normalize_source
    from helpers import is_verbatim
    text = normalize_source(source, note=note, verbatim=is_verbatim())
    if not text:
        return []
    role = "source"
    if exhibit_no:
        text, role = f"Exhibit {exhibit_no} · {text}", "exhibit-badge"
    shape = builders.add_tb(slide, text, 0.7, 6.80, 9.5, 0.3, size=11,
                            color=pal["text_muted"], font=pal["font_label"])
    return [(shape, role)]


# Footnote block: one text box bottom-anchored just above the source line at
# design-y 6.80, growing upward, one numbered entry per paragraph line.
FOOTNOTE_BOTTOM = 6.80
FOOTNOTE_LINE_H = 0.20
FOOTNOTE_PAD = 0.10   # top+bottom default text insets
FOOTNOTE_MAX_LINES = 3
KICKER_TOP = 6.18     # builders.KICKER_TOP (kept equal; asserted in tests)


def footnote_anchor(p):
    """Bottom edge of the footnote block: just above the kicker band when
    the slide declares one, else just above the source line at 6.80."""
    return KICKER_TOP - 0.06 if p.get("kicker") else FOOTNOTE_BOTTOM


def footnote_top(p):
    """Top y of the slide's footnote block (design inches), for the content
    budget — builders.content_bottom() keeps body content above it."""
    n = min(len(p.get("footnotes") or []), FOOTNOTE_MAX_LINES)
    if not n:
        return FOOTNOTE_BOTTOM
    return footnote_anchor(p) - (FOOTNOTE_PAD + FOOTNOTE_LINE_H * n)


def add_footnotes(slide, pal, entries, bottom=FOOTNOTE_BOTTOM):
    """`- Footnote:` block — superscript-numbered 10pt muted lines above the
    source line, bottom-anchored at `bottom` (default design-y 6.80; the
    build passes footnote_anchor(p) so the block clears a kicker band).
    Max 3 lines; overflow warns and truncates the list."""
    from footnotes import number_footnotes, superscript
    import builders
    from palettes import hex_rgb
    from pptx.util import Pt
    numbered = number_footnotes(entries)
    if len(numbered) > FOOTNOTE_MAX_LINES:
        warn(f"{len(numbered)} footnotes exceed the {FOOTNOTE_MAX_LINES}-line "
             "block — extra entries dropped")
        numbered = numbered[:FOOTNOTE_MAX_LINES]
    if not numbered:
        return []
    h = FOOTNOTE_PAD + FOOTNOTE_LINE_H * len(numbered)
    tb = builders.add_tb(
        slide, f"{superscript(numbered[0][0])} {numbered[0][1]}",
        0.7, bottom - h, 9.5, h, size=10,
        color=pal["text_muted"], font=pal["font_label"])
    for num, text in numbered[1:]:
        run = tb.text_frame.add_paragraph().add_run()
        run.text = f"{superscript(num)} {text}"
        run.font.size = Pt(builders._fs(10))
        run.font.color.rgb = hex_rgb(pal["text_muted"])
        run.font.name = pal["font_label"]
    return [(tb, "footnote")]


def add_file_stamp(slide, pal, text, with_footer=False):
    """**File-Stamp:** on — `<output filename> · <date>` in the footer band,
    muted 10pt; drawn with footer chrome on every footered slide.  Centred
    in the band, unless a footer text is present — then it sits in the free
    stretch between the footer text (ends 6.7) and the logo slot (10.2)."""
    import builders
    from pptx.enum.text import PP_ALIGN
    left, w = (6.9, 3.2) if with_footer else (4.0, 5.33)
    tb = builders.add_tb(slide, text, left, 7.08, w, 0.35, size=10,
                         color=pal["text_muted"], font=pal["font_label"],
                         align=PP_ALIGN.CENTER)
    return [(tb, "file-stamp")]


def add_stamp(slide, pal, text):
    """Bordered status tag (DRAFT, CONFIDENTIAL) top-left, above the heading."""
    import builders
    from pptx.enum.text import PP_ALIGN
    box = builders.add_rect(slide, 0.7, 0.14, 1.5, 0.32, pal["bg"],
                            line_hex=pal["accent3"], line_pt=1.0)
    box.fill.background()  # outline only — works on photos and gradients
    txt = builders.add_tb(slide, text.upper(), 0.7, 0.16, 1.5, 0.28, size=11,
                    bold=True, color=pal["text_muted"], font=pal["font_label"],
                    align=PP_ALIGN.CENTER)
    return [(box, "stamp"), (txt, "stamp")]


TITLE_LAYOUT = "title"

# Profile logo (profiles.py `logo`): design-space boxes, contain-fitted.
LOGO_LAYOUTS = {"title", "closing"}     # corner badge on these…
LOGO_CORNER = (10.9, 6.55, 1.7, 0.55)   # …bottom-right, above the footer zone
LOGO_FOOTER = (10.2, 7.08, 1.5, 0.34)   # footer band, left of the page number


def resolve_logo(spec, ctx):
    """The profile's `logo` as a file path, or None (warned once per build).

    Resolved like any outline image (literal, outline dir, assets dir) plus
    `<assets>/brand/` — where brand_kit.py writes `<name>-logo.png`."""
    if not spec:
        return None
    from helpers import resolve_image_path
    path = resolve_image_path(spec, ctx)
    if path is None:
        candidate = Path(ctx.get("assets_dir", "assets")) / "brand" / spec
        path = candidate if candidate.is_file() else None
    if path is None:
        warn(f"profile logo not found: {spec!r} (tried literal path, outline "
             "dir, assets dir, <assets>/brand/) — building without it")
    return path


def add_logo(slide, layout_name, logo_path):
    """Place the profile logo: a corner badge on title/closing slides, a small
    mark in the footer band (clear of the page number) everywhere a footer is
    drawn. Returns [(picture, "logo")] or [] for layouts that carry neither."""
    import builders
    from outline_validate import NO_FOOTER_LAYOUTS
    if layout_name in LOGO_LAYOUTS:
        box = LOGO_CORNER
    elif layout_name in NO_FOOTER_LAYOUTS:
        return []
    else:
        box = LOGO_FOOTER
    pic = builders.add_picture_contain(slide, logo_path, *box, alt="logo")
    return [(pic, "logo")]


def deck_title(slides, output_path, outline_path=None):
    """Deck name for docProps/core.xml `dc:title`.

    Precedence: the first `title`-layout slide's `- Title:`, else the outline's
    first `# ` heading, else the output filename stem."""
    for slide in slides:
        if slide.get("layout") == TITLE_LAYOUT and slide.get("title"):
            return slide["title"].strip()
    if outline_path:
        try:
            for line in Path(outline_path).read_text(
                    encoding="utf-8").splitlines():
                if line.startswith("# "):
                    return line[2:].strip()
        except OSError:
            pass
    return Path(output_path).stem
