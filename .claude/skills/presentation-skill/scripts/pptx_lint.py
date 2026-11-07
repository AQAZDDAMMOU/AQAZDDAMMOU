#!/usr/bin/env python3
"""
pptx_lint.py — deterministic cross-slide consistency lint for a .pptx.

Complements qa_check.py (per-slide defects) with deck-wide checks in the
Macabacus "Deck Check" / UpSlide "Slide Check" class:

  1. Page-number sequence: consecutive, no gaps or duplicates.
  2. Anti-jiggle: recurring elements (page numbers, footers, section labels)
     must sit at identical coordinates on every slide.
  3. Axis honesty: same-type charts whose explicit value-axis maxima differ
     by more than 10x read as a dishonest scale comparison.
  4. Chart honesty (chart_lint.check_charts): per-chart Misviz/Zelazny rules —
     truncated bar baselines and 3-D perspective are errors; dual axes, pies
     that are not part-to-whole, uneven date ticks, missing legends/units,
     unexplained indexing, extra highlight colours and over-labelled lines
     are warned.
  5. Numeric consistency (numeric_lint.check_numbers): mixed decimal/
     negative/unit conventions, a table 'Total' that does not add up, one
     KPI label carrying two different values, and title numbers that the
     exhibit below never shows — all warned.
  6. Color and font inventory: more than MAX_FONTS distinct fonts reads as
     inconsistent; with --palette, any explicit run or fill color outside
     the palette (plus known extras) is an error.
  7. Typography: with --palette, fonts outside the palette's font pair set
     are errors (brand enforcement, like the color whitelist) and sizes off
     the palette's type_scale ladder are warned; without a palette, type
     ladder sprawl (too many distinct sizes) is warned.
  8. Fonts installed: fonts the local system cannot supply are warned —
     visual QA would silently substitute them.
  9. AI tells: hallmarks of AI-generated decks (a thin accent line under the
     title, long centered body copy) are warned.
 10. Composition: 3+ consecutive slides with the same structural signature
     (shape mix) read as monotonous; 4+ consecutive dense slides with no
     breathing room break reading rhythm — both warned.
 11. Google Slides compatibility (--gslides): fonts outside the Google Slides
     set, SmartArt, non-fade transitions, and embedded media all degrade on
     import — each is warned.
 12. Client safety (--client-safe): reviewer comments, hidden slides and
     `--denylist` terms are errors; core-props identity, off-slide/hidden
     shapes, unused layouts, embedded workbooks, the package thumbnail and
     speaker notes are warned (see client_safe.py).

Issues are emitted in fix order (fix upstream classes first — repairing
structure or color often clears downstream findings): page sequence →
jiggle → axis honesty → chart honesty → numeric consistency → color/palette →
typography → fonts installed → AI tells → composition → Google Slides →
client safety.

Usage:
    python3 scripts/pptx_lint.py deck.pptx [--palette midnight-executive] [--gslides]
"""
import argparse
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn
from chart_lint import check_charts
from numeric_lint import check_numbers
from qa_check import (CHART_UNREADABLE, iter_shapes, safe_chart,
                      _solid_fill_rgb)

EMU_IN = 914400
POS_TOL_IN = 0.03           # jiggle tolerance
FOOTER_ZONE = 0.80          # top > 80% of slide height = footer zone
HEADER_ZONE = 0.12          # top < 12% = header zone (kickers, stamps)
MAX_FONTS = 4
MAX_COLORS = 14
MAX_TYPE_SIZES = 14         # distinct font sizes above this = ladder sprawl
MONOTONY_RUN = 3            # same structural signature this many times in a row
DENSE_WORDS = 130           # words per slide above this = dense
DENSE_RUN = 4               # consecutive dense slides before rhythm warning
TEXTBOX_BUCKETS = (2, 5, 9)  # textbox-count buckets for structure signatures
PAGENUM_RX = re.compile(r"^(B·)?\d+$")
# waterfall negative red, pure white/black (cards, shadows) are always allowed
EXTRA_ALLOWED = {"D9655B", "FFFFFF", "000000"}
_HEX6 = re.compile(r"[0-9A-Fa-f]{6}")

# fonts that survive a Google Slides import without substitution
GSLIDES_SAFE_FONTS = {
    "Arial", "Georgia", "Times New Roman", "Verdana", "Trebuchet MS",
    "Courier New", "Roboto", "Open Sans", "Lato", "Montserrat",
    "Merriweather", "Source Sans Pro", "Playfair Display", "Inter",
    "Oswald", "Raleway", "PT Sans", "Nunito", "Work Sans", "Poppins",
}
_GSLIDES_SAFE_LOWER = {f.lower() for f in GSLIDES_SAFE_FONTS}
DIAGRAM_NS = b"http://schemas.openxmlformats.org/drawingml/2006/diagram"
MEDIA_EXTS = (".mp4", ".mov", ".mp3", ".wav", ".m4a")


def _chart_series_allowed(pal, issues):
    """Return uppercased hex colors from chart_series; warn on bad entries."""
    allowed = set()
    series = pal.get("chart_series")
    if series is None:
        return allowed
    if not isinstance(series, list):
        issues["warn"].append(
            f"palette chart_series is not a list ({type(series).__name__}) — "
            "skipped for whitelist")
        return allowed
    for i, item in enumerate(series):
        if not isinstance(item, str):
            issues["warn"].append(
                f"palette chart_series[{i}] is not a string "
                f"({type(item).__name__}) — skipped")
            continue
        if not _HEX6.fullmatch(item):
            issues["warn"].append(
                f"palette chart_series[{i}] is not a 6-digit hex color — skipped")
            continue
        allowed.add(item.upper())
    return allowed


def _shape_box_in(shape):
    try:
        if None in (shape.left, shape.top):
            return None
        return (shape.left / EMU_IN, shape.top / EMU_IN)
    except (AttributeError, TypeError):
        return None


def _frame_text(shape):
    """Frame text including cached field text — the page number is a
    `<a:fld type="slidenum">` whose cached string python-pptx's `.text`
    skips, so the footer-zone checks read every `a:t` instead."""
    body = shape.text_frame._txBody
    return "\n".join(
        "".join(t.text or "" for t in para.iter(qn("a:t")))
        for para in body.findall(qn("a:p")))


def _iter_text_shapes(prs):
    for n, slide in enumerate(prs.slides, 1):
        for shape in iter_shapes(slide.shapes):
            if getattr(shape, "has_text_frame", False) \
                    and _frame_text(shape).strip():
                yield n, slide, shape


def check_jiggle(prs, issues):
    """Recurring header/footer-zone elements must not move between slides."""
    sh_in = prs.slide_height / EMU_IN
    groups = defaultdict(list)  # role -> [(slide_no, left, top)]
    for n, _slide, shape in _iter_text_shapes(prs):
        box = _shape_box_in(shape)
        if box is None:
            continue
        left, top = box
        in_footer = top > FOOTER_ZONE * sh_in
        in_header = top < HEADER_ZONE * sh_in
        if not (in_footer or in_header):
            continue
        text = _frame_text(shape).strip()
        role = "page-number" if PAGENUM_RX.match(text) else text[:60]
        groups[role].append((n, left, top))
    for role, occ in groups.items():
        if len(occ) < 3:
            continue
        ref_l, ref_t = Counter(
            (round(lf, 2), round(tp, 2)) for _, lf, tp in occ).most_common(1)[0][0]
        for n, left, top in occ:
            if abs(left - ref_l) > POS_TOL_IN or abs(top - ref_t) > POS_TOL_IN:
                issues["error"].append(
                    f"Slide {n}: '{role}' jiggle — at ({left:.2f},{top:.2f})in "
                    f"but majority position is ({ref_l:.2f},{ref_t:.2f})in")


def check_page_sequence(prs, issues):
    """Plain-digit footer-zone page numbers must track slide positions consistently.

    The invariant is: for any two consecutive numbered slides (n1, n2) with
    page numbers (v1, v2), the number delta must equal the slide-position delta:
        v2 - v1  ==  n2 - n1

    This allows gaps that span unnumbered layouts (title, section-divider, etc.)
    while still catching duplicates (v2 == v1 on different slides) and true
    misnumbering (number delta differs from slide delta).

    The page number is a `slidenum` field; its cached text (read via
    _frame_text) is what build_deck wrote, stays monotonic under a
    `**Page-Start:**` offset, and is what this check compares — the delta
    invariant is unaffected by a constant offset.
    """
    sh_in = prs.slide_height / EMU_IN
    pages = []  # (slide_no, value)
    for n, _slide, shape in _iter_text_shapes(prs):
        box = _shape_box_in(shape)
        if box is None or box[1] <= FOOTER_ZONE * sh_in:
            continue
        text = _frame_text(shape).strip()
        if text.isdigit():
            pages.append((n, int(text)))
    for (n1, v1), (n2, v2) in zip(pages, pages[1:]):
        if v2 - v1 != n2 - n1:
            issues["error"].append(
                f"Slide {n2}: page-number sequence broken "
                f"({v1} on slide {n1}, then {v2}, {n2 - n1} slide(s) later)")


def check_axis_scales(prs, issues):
    """Same-type native charts whose explicit value-axis maxima differ by
    more than 10x read as a dishonest scale comparison."""
    groups = defaultdict(list)  # chart_type -> [(slide_no, maximum_scale)]
    for n, slide in enumerate(prs.slides, 1):
        for shape in iter_shapes(slide.shapes):
            if not getattr(shape, "has_chart", False):
                continue
            chart = safe_chart(shape)
            if chart is None:
                issues["warn"].append(f"Slide {n}: {CHART_UNREADABLE}")
                continue
            try:
                max_scale = chart.value_axis.maximum_scale
            except ValueError:
                continue  # pie/doughnut have no value axis
            if max_scale is not None:
                groups[chart.chart_type].append((n, max_scale))
    for ctype, occ in groups.items():
        if len(occ) < 2:
            continue
        (n_lo, lo) = min(occ, key=lambda o: o[1])
        (n_hi, hi) = max(occ, key=lambda o: o[1])
        if lo > 0 and hi > 10 * lo:
            name = getattr(ctype, "name", str(ctype))
            issues["warn"].append(
                f"{name} charts on slides {n_lo} and {n_hi} have value-axis "
                f"maxima {lo:g} vs {hi:g} (>10x apart) — possible dishonest "
                "scale comparison")


LINE_H_MAX_EMU = 76200          # 6pt — thicker than this isn't a "line"
MIN_BODY_SHAPES = 4             # fewer shapes = title/quote/big-number slide


def check_ai_tells(prs, issues):
    """Hallmarks of AI-generated decks: a thin accent line under the title,
    and long centered body copy."""
    from pptx.enum.text import PP_ALIGN
    from pptx.util import Pt
    for n, slide in enumerate(prs.slides, 1):
        shapes = list(iter_shapes(slide.shapes))
        # (a) accent line under the title
        title_bottoms = []
        seen_text = False
        for shape in shapes:
            if not getattr(shape, "has_text_frame", False) \
                    or not shape.text_frame.text.strip():
                continue
            if None in (shape.top, shape.height):
                continue
            first = not seen_text
            seen_text = True
            if shape.top >= 1.2 * EMU_IN:
                continue
            sizes = [r.font.size for para in shape.text_frame.paragraphs
                     for r in para.runs if r.font.size is not None]
            if first or (sizes and max(sizes) >= Pt(20)):
                title_bottoms.append(shape.top + shape.height)
        for shape in shapes:
            if getattr(shape, "has_text_frame", False) \
                    and shape.text_frame.text.strip():
                continue  # text shapes are not accent lines
            try:
                if None in (shape.top, shape.width, shape.height):
                    continue
            except AttributeError:
                continue
            if shape.height > LINE_H_MAX_EMU or shape.width < EMU_IN:
                continue
            if any(0 <= shape.top - b <= 0.25 * EMU_IN for b in title_bottoms):
                issues["warn"].append(
                    f"slide {n}: accent line under title (AI-generated look) "
                    "— remove")
        # (b) long centered body text (sparse slides are intentional centers)
        if len(shapes) < MIN_BODY_SHAPES:
            continue
        for shape in shapes:
            if not getattr(shape, "has_text_frame", False):
                continue
            tf = shape.text_frame
            if len(tf.text.strip()) <= 80:
                continue
            if None in (shape.top, shape.height) \
                    or shape.top <= 1.5 * EMU_IN or shape.height <= EMU_IN:
                continue
            if any(para.alignment == PP_ALIGN.CENTER
                   for para in tf.paragraphs if para.runs):
                issues["warn"].append(
                    f"slide {n}: long centered body text — left-align "
                    "body copy")


def collect_inventory(prs):
    """(fonts, colors, sizes): name/hex/pt -> set of slide numbers using it."""
    fonts, colors, sizes = defaultdict(set), defaultdict(set), defaultdict(set)
    for n, slide, shape in _iter_text_shapes(prs):
        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                if not run.text.strip():
                    continue
                if run.font.name:
                    fonts[run.font.name].add(n)
                if run.font.size is not None:
                    sizes[run.font.size.pt].add(n)
                try:
                    rgb = run.font.color.rgb
                except (TypeError, AttributeError):
                    rgb = None
                if rgb is not None:
                    colors[str(rgb)].add(n)
    for n, slide in enumerate(prs.slides, 1):
        for shape in iter_shapes(slide.shapes):
            fill = _solid_fill_rgb(shape)
            if fill is not None:
                colors["%02X%02X%02X" % fill].add(n)
    return fonts, colors, sizes


def check_inventory(inv, issues, palette_key=None):
    fonts, colors, _ = inv
    if len(fonts) > MAX_FONTS:
        issues["warn"].append(
            f"deck uses {len(fonts)} fonts ({', '.join(sorted(fonts))}) — "
            f"more than {MAX_FONTS} reads as inconsistent")
    if palette_key:
        from palettes import PALETTES
        # Comma-separated keys support per-slide {palette=...} overrides:
        # a color passes when any listed palette allows it.
        keys = [k.strip() for k in palette_key.split(",") if k.strip()]
        allowed, known = set(EXTRA_ALLOWED), []
        for key in keys:
            pal = PALETTES.get(key)
            if pal is None:
                issues["warn"].append(f"unknown palette '{key}' — "
                                      "excluded from whitelist check")
                continue
            known.append(key)
            allowed |= {v.upper() for v in pal.values()
                        if isinstance(v, str) and _HEX6.fullmatch(v)}
            allowed |= _chart_series_allowed(pal, issues)
        if not known:
            issues["warn"].append("no known palettes given — "
                                  "skipping whitelist check")
            return
        for hexv, slides in sorted(colors.items()):
            if hexv.upper() not in allowed:
                where = ", ".join(str(s) for s in sorted(slides)[:5])
                issues["error"].append(
                    f"off-palette color {hexv} on slide(s) {where} "
                    f"(palette: {', '.join(known)})")
    elif len(colors) > MAX_COLORS:
        issues["warn"].append(
            f"deck uses {len(colors)} distinct colors — more than "
            f"{MAX_COLORS}; pass --palette to enforce a whitelist")


# 4:3 decks rescale design-canvas fonts by 10.0/13.33 — that scaled ladder
# is still a consistent type system, so both ladders are accepted.
CANVAS_SCALES = (1.0, 10.0 / 13.33)


def check_typography(inv, issues, palette_key=None):
    """Brand type enforcement, symmetric with the color whitelist.

    With --palette: fonts outside the palettes' title/body/label pairs
    (plus the CJK stack) are errors; sizes off the palettes' type_scale
    ladder (at either canvas scale) are one aggregated warning. Without a
    palette: distinct-size sprawl is warned.
    """
    fonts, _, sizes = inv
    if not palette_key:
        if len(sizes) > MAX_TYPE_SIZES:
            seq = ", ".join(f"{s:g}" for s in sorted(sizes))
            issues["warn"].append(
                f"deck uses {len(sizes)} distinct font sizes ({seq}) — "
                f"more than {MAX_TYPE_SIZES} reads as an undisciplined "
                "type ladder; pass --palette to check against a scale")
        return
    from palettes import CJK_FONT_STACK, PALETTES
    keys = [k.strip() for k in palette_key.split(",") if k.strip()]
    allowed_fonts = {f.lower() for f in CJK_FONT_STACK}
    ladder = set()
    known = False
    for key in keys:
        pal = PALETTES.get(key)
        if pal is None:
            continue  # check_inventory already warned about unknown keys
        known = True
        allowed_fonts |= {str(pal.get(k, "")).lower()
                          for k in ("font_title", "font_body", "font_label")}
        for step in pal.get("type_scale", []):
            for scale in CANVAS_SCALES:
                ladder.add(round(step * scale * 2) / 2)  # nearest 0.5pt
    if not known:
        return
    for name, slide_nos in sorted(fonts.items()):
        if name.lower() not in allowed_fonts:
            where = ", ".join(str(s) for s in sorted(slide_nos)[:5])
            issues["error"].append(
                f"off-brand font '{name}' on slide(s) {where} "
                f"(palette: {', '.join(keys)})")
    off_ladder = sorted(pt for pt in sizes
                        if round(pt * 2) / 2 not in ladder)
    if off_ladder:
        seq = ", ".join(f"{s:g}pt" for s in off_ladder)
        slide_nos = sorted(set().union(*(sizes[pt] for pt in off_ladder)))
        where = ", ".join(str(s) for s in slide_nos[:5])
        issues["warn"].append(
            f"font size(s) {seq} (slide(s) {where}) off the palette type "
            "scale — sizes outside the ladder read as inconsistent")


def _structure_signature(slide):
    """Coarse structural fingerprint: what kinds of content the slide holds.
    Textbox counts are bucketed so minor copy differences don't hide true
    layout repetition."""
    pictures = charts = tables = textboxes = shapes = 0
    for shape in iter_shapes(slide.shapes):
        if shape.shape_type in (MSO_SHAPE_TYPE.PICTURE,
                                MSO_SHAPE_TYPE.LINKED_PICTURE):
            pictures += 1
        elif getattr(shape, "has_chart", False):
            charts += 1
        elif getattr(shape, "has_table", False):
            tables += 1
        elif getattr(shape, "has_text_frame", False) \
                and shape.text_frame.text.strip():
            textboxes += 1
        else:
            shapes += 1
    bucket = sum(1 for b in TEXTBOX_BUCKETS if textboxes > b)
    return (pictures, charts, tables, bucket, min(shapes, 6))


def check_structure_monotony(prs, issues):
    """3+ consecutive slides with the same structural signature read as a
    template on repeat — the composition equivalent of bullet-list fatigue."""
    sigs = [_structure_signature(slide) for slide in prs.slides]
    run_start = 0
    for i in range(1, len(sigs) + 1):
        if i == len(sigs) or sigs[i] != sigs[run_start]:
            if i - run_start >= MONOTONY_RUN:
                issues["warn"].append(
                    f"slides {run_start + 1}-{i} share one structural "
                    f"signature ({i - run_start} in a row) — vary the "
                    "layout (chart, cards, comparison) to break monotony")
            run_start = i


def check_density_rhythm(prs, issues):
    """4+ consecutive dense slides (no light slide in between) exhaust the
    reader — alternate dense exhibits with breathing room."""
    words = []
    for slide in prs.slides:
        count = 0
        for shape in iter_shapes(slide.shapes):
            if getattr(shape, "has_text_frame", False):
                count += len(shape.text_frame.text.split())
        words.append(count)
    run_start = None
    for i, count in enumerate(words + [0]):  # sentinel flushes the last run
        if count > DENSE_WORDS:
            if run_start is None:
                run_start = i
            continue
        if run_start is not None and i - run_start >= DENSE_RUN:
            issues["warn"].append(
                f"slides {run_start + 1}-{i} are all dense "
                f"(>{DENSE_WORDS} words, {i - run_start} in a row) — "
                "insert a light slide (big-number, quote, section divider) "
                "for reading rhythm")
        run_start = None


def installed_fonts():
    """Lowercased family names visible to the renderer, or None if unknowable.

    fc-list is the most truthful source for what LibreOffice will see;
    matplotlib's font_manager is the cross-platform fallback.
    """
    import shutil
    import subprocess
    if shutil.which("fc-list"):
        try:
            out = subprocess.run(["fc-list", ":", "family"],
                                 capture_output=True, text=True, timeout=20)
            if out.returncode == 0:
                names = set()
                for line in out.stdout.splitlines():
                    for fam in line.split(","):
                        names.add(fam.strip().lower())
                if names:
                    return names
        except (OSError, subprocess.TimeoutExpired):
            pass
    try:
        from matplotlib import font_manager
        names = {n.lower() for n in font_manager.get_font_names()}
        return names or None
    except (ImportError, AttributeError):
        return None


def check_fonts_installed(inv, issues):
    installed = installed_fonts()
    if installed is None:
        issues["warn"].append(
            "cannot enumerate installed fonts (no fc-list or matplotlib) — "
            "visual QA may silently substitute fonts")
        return
    fonts = inv[0]
    for name, slides in sorted(fonts.items()):
        if name.lower() not in installed:
            where = ", ".join(str(s) for s in sorted(slides)[:5])
            issues["warn"].append(
                f"font '{name}' (slide(s) {where}) not installed — renderer "
                "will substitute; QA thumbnails won't match PowerPoint")


def check_gslides(prs, issues, inv=None):
    """Google Slides import hazards: font substitution, SmartArt flattening,
    dropped transitions, embedded media."""
    fonts = (inv if inv is not None else collect_inventory(prs))[0]
    for name, slides in sorted(fonts.items()):
        if name.lower() not in _GSLIDES_SAFE_LOWER:
            where = ", ".join(str(s) for s in sorted(slides)[:5])
            issues["warn"].append(
                f"gslides: font '{name}' (slide(s) {where}) not in the Google "
                "Slides font set — will be substituted on import")
    for n, slide in enumerate(prs.slides, 1):
        if DIAGRAM_NS in slide.part.blob:
            issues["warn"].append(
                f"gslides: slide {n} contains SmartArt — flattens to a "
                "static image on Google Slides import")
        trans = slide.element.find(qn("p:transition"))
        if trans is None:
            continue
        child = next(iter(trans), None)
        tag = getattr(child, "tag", None)
        if isinstance(tag, str) and tag != qn("p:fade"):
            issues["warn"].append(
                f"gslides: slide {n} transition '{tag.split('}')[-1]}' "
                "dropped/changed by Google Slides (only fade survives)")
    for part in prs.part.package.iter_parts():
        name = str(part.partname)
        if name.startswith("/ppt/media/") and name.lower().endswith(MEDIA_EXTS):
            issues["warn"].append(
                f"gslides: embedded media {name} — audio/video does not "
                "import into Google Slides")


_CHART_UNREADABLE_RX = re.compile(
    r"^Slide (\d+): " + re.escape(CHART_UNREADABLE) + r"$")


def _dedupe_chart_unreadable(issues):
    """The axis check, chart_lint and numeric_lint sub-checks each reach
    for `shape.chart` independently, so a single dangling chart
    relationship produces one identical warning per sub-check. Collapse
    those to the first occurrence per slide."""
    seen = set()
    deduped = []
    for msg in issues["warn"]:
        m = _CHART_UNREADABLE_RX.match(msg)
        if m:
            slide_no = m.group(1)
            if slide_no in seen:
                continue
            seen.add(slide_no)
        deduped.append(msg)
    issues["warn"] = deduped


def lint_deck(prs, palette_key=None, gslides=False, client_safe=False,
              denylist=None):
    """Checks run — and issues are appended — in fix order: repairing an
    upstream class (structure, color) often clears downstream findings, so
    the printed order is the recommended repair order."""
    issues = {"error": [], "warn": []}
    check_page_sequence(prs, issues)      # 1. structure
    check_jiggle(prs, issues)             # 2. geometry drift
    check_axis_scales(prs, issues)        # 3. data honesty
    check_charts(prs, issues)             # 4. per-chart honesty
    check_numbers(prs, issues)            # 5. numeric consistency
    inv = collect_inventory(prs)
    check_inventory(inv, issues, palette_key)   # 6. color/palette
    check_typography(inv, issues, palette_key)  # 7. type system
    check_fonts_installed(inv, issues)    # 8. fonts installed
    check_ai_tells(prs, issues)           # 9. AI tells
    check_structure_monotony(prs, issues)  # 10. composition
    check_density_rhythm(prs, issues)
    if gslides:
        check_gslides(prs, issues, inv)   # 11. export compat
    if client_safe:
        from client_safe import scan_presentation
        scan_presentation(prs, denylist, issues)  # 12. client safety
    _dedupe_chart_unreadable(issues)
    return issues


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("pptx")
    parser.add_argument("--palette", default=None,
                        help="Enforce palette colors as a whitelist; comma-separate "
                             "keys when slides use {palette=...} overrides")
    parser.add_argument("--assets-dir", default="assets",
                        help="Root assets dir; custom palettes load from "
                             "<assets-dir>/palettes/")
    parser.add_argument("--gslides", action="store_true",
                        help="Warn on features that degrade when the deck is "
                             "imported into Google Slides")
    parser.add_argument("--client-safe", action="store_true",
                        help="Audit the deck for content that must not reach a "
                             "client: comments, hidden slides, internal notes, "
                             "core-props identity (see client_safe.py)")
    parser.add_argument("--denylist", default=None, metavar="FILE",
                        help="File of terms that must not appear anywhere in "
                             "the deck: one per line, # comments; implies "
                             "--client-safe")
    args = parser.parse_args()
    from palettes import load_custom_palettes
    load_custom_palettes(Path(args.assets_dir) / "palettes")
    prs = Presentation(args.pptx)
    denylist = None
    if args.denylist:
        from client_safe import load_denylist
        denylist = load_denylist(args.denylist)
    issues = lint_deck(prs, args.palette, gslides=args.gslides,
                       client_safe=args.client_safe or bool(denylist),
                       denylist=denylist)
    for e in issues["error"]:
        print(f"  [ERROR] {e}")
    for w in issues["warn"]:
        print(f"  [WARN]  {w}")
    print(f"\n{Path(args.pptx).name}: {len(issues['error'])} error(s), "
          f"{len(issues['warn'])} warning(s)")
    sys.exit(1 if issues["error"] else 0)


if __name__ == "__main__":
    main()
