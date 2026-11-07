#!/usr/bin/env python3
"""
build_deck.py — Build a .pptx from a markdown outline.

Usage:
    python scripts/build_deck.py outline.md --output deck.pptx
        [--palette midnight-executive] [--template template.pptx]
        [--assets-dir assets] [--check]

--check validates the outline (layouts, images, chart data) without building.
Outline format: references/generation-guide.md.
"""
import argparse
import sys
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches

from builders import LAYOUT_MAP
# PowerPoint sections and the template-placeholder slide path live in their
# own modules; the old names stay bound here for back-compat.
from deck_sections import (  # noqa: F401  (re-export)
    P14_NS, SECTIONLST_EXT_URI, _section_plan, inject_sections)
from helpers import add_speaker_notes, set_verbatim, warn
# Outline validation lives in outline_validate.py; re-exported here so
# `from build_deck import validate` (and the layout sets) keeps working.
from outline_validate import (  # noqa: F401  (re-export)
    validate, ACTION_TITLE_LAYOUTS, NO_FOOTER_LAYOUTS, _validate_driver_tree,
    _validate_raci, _validate_stakeholder_map, _validate_tree)
# Outline data rows (- label: value, **Series:**, - Data-File:) live in
# outline_data.py (a leaf module shared with outline_validate); re-exported.
from outline_data import (  # noqa: F401  (re-export)
    NUM_RE, SLIDE_HEADING_PREFIX, XLSX_SUFFIXES, _csv_value, _parse_csv_rows, _parse_data_point,
    _read_xlsx_rows, _resolve_data_file, _series_count, _stamp_provenance,
    load_data_files)
# The parser and the deck-level transforms are leaf modules now; every name
# stays importable from here (`from build_deck import parse_outline`, …).
from outline_parse import (  # noqa: F401  (re-export)
    FIELD_KEYS, KV_RE, META_KEYS, _clean_bullet, _parse_kv, parse_outline)
from outline_replay import replay
from outline_transforms import (  # noqa: F401  (re-export)
    CJK_RE, SCALE_GROUP_KINDS, _agenda_slide, _references_slide,
    _text_values, apply_auto_agenda, apply_references, apply_scale_groups,
    deck_has_cjk, divider_page_map, resolve_diagram_visuals)
from palettes import (PALETTES, get_palette, load_custom_palettes,
                      set_cjk)
from render_diagram import RendererMissing
import slide_tags
from template_build import (  # noqa: F401  (re-export)
    TEMPLATE_WANT, _fill_placeholders, build_template_slide)
from template_helpers import load_template_config, palette_from_theme
from verbatim import is_verbatim_deck

# Per-slide chrome (footer, tracker, sticker, kicker, source, stamp) lives in
# slide_chrome.py; the old private names stay bound here for back-compat.
from slide_chrome import (  # noqa: F401  (re-export)
    TAB_LABEL_CHARS, TAB_RIGHT_EDGE, TAB_STRIP_MAX_W, TITLE_LAYOUT, _tab_label,
    add_file_stamp, add_footer, add_footnotes, add_kicker_box, add_logo,
    add_section_tracker, add_source, add_stamp, add_sticker, add_tracker_tabs,
    deck_title, footnote_anchor, parse_page_start, resolve_logo,
    set_first_slide_num)

_add_footer = add_footer
_add_section_tracker = add_section_tracker
_add_tracker_tabs = add_tracker_tabs
_add_sticker = add_sticker
_add_kicker_box = add_kicker_box
_add_source = add_source
_add_stamp = add_stamp
_deck_title = deck_title

# ── Main ─────────────────────────────────────────────────────────────────────
SIZES = {"16:9": (13.33, 7.5), "4:3": (10.0, 7.5)}


def create_presentation(template_path=None, size="16:9"):
    if template_path:
        return Presentation(template_path)
    prs = Presentation()
    w, h = SIZES.get(size, SIZES["16:9"])
    prs.slide_width = Inches(w)
    prs.slide_height = Inches(h)
    return prs


def build(outline_path, output_path, palette_key=None,
          template_path=None, assets_dir=None, check_only=False, size=None,
          density=None, variant=None, ghost=False, style_path=None,
          profile=None, verbatim=False):
    outline_path = Path(outline_path)
    ctx = {
        "outline_dir": outline_path.parent,
        "assets_dir": Path(assets_dir) if assets_dir else Path("assets"),
        "template_path": template_path,
        "style_path": style_path,
        # CLI knobs replay() resolves against the front-matter and profile
        "profile": profile,
        "palette": palette_key,
        "density": density,
        "variant": variant,
        # tri-state: True forces verbatim, None lets `**Copy:** locked` decide
        "verbatim": verbatim or None,
    }
    custom = load_custom_palettes(ctx["assets_dir"] / "palettes")
    if custom:
        print(f"  Loaded custom palettes: {', '.join(custom)}")
    if palette_key and palette_key not in PALETTES:
        print(f"  [WARN] unknown --palette '{palette_key}' — using default",
              file=sys.stderr)
    set_cjk(False)  # deck-wide CJK overlay is decided per build, below
    set_verbatim(False)          # both are process-global: reset, then decide
    # replay() applies the profile, installs the density and repicks layouts;
    # the palette key it hands back already has CLI > variant > front-matter
    # precedence applied.
    rr = replay(outline_path, ctx=ctx)
    meta, slides_data = rr.meta, rr.slides
    palette_key = rr.palette_key
    # Verbatim copy mode: ctx carries it into validate() and the slide tags.
    ctx["verbatim"] = is_verbatim_deck(ctx, meta)
    set_verbatim(ctx["verbatim"])
    # Deck-wide number-format policy: process-global, so always reset it first
    # (a malformed spec must not leave the previous build's policy installed).
    import numfmt
    numfmt.set_policy(None)
    policy_errors = []
    try:
        numfmt.set_policy(numfmt.parse_policy(meta.get("number_format", "")))
    except ValueError as exc:
        policy_errors.append(f"**Number-Format:** {exc}")
    data_errors, data_warnings = rr.errors, rr.warnings

    # Template mode: the template's theme (colours + major/minor faces)
    # overrides the palette; read it now so the measure font below is the face
    # the deck will actually render in.
    template_config, theme_pal = None, None
    if template_path:
        template_config = load_template_config(template_path)
        from profile_template import extract_theme_colors, extract_theme_fonts
        from pptx import Presentation as _Prs
        _tmpl = _Prs(str(template_path))
        theme_pal = palette_from_theme(extract_theme_colors(_tmpl),
                                       extract_theme_fonts(_tmpl))

    import textfit
    # Measure wrapping with the deck's own body face so validate()'s fail-fast
    # capacity gate and the renderer agree on line counts (in template mode:
    # the theme's minor/body face).
    textfit.set_measure_font(
        (theme_pal or {}).get("font_body")
        or get_palette(palette_key or meta.get("palette", "")).get("font_body"))
    errors, warnings = validate(slides_data, ctx, meta)
    errors = policy_errors + data_errors + errors
    warnings = data_warnings + warnings
    for w in warnings:
        print(f"  [WARN] {w}", file=sys.stderr)
    if errors:
        for e in errors:
            print(f"  [ERROR] {e}", file=sys.stderr)
        print(f"\nOutline validation failed: {len(errors)} error(s).", file=sys.stderr)
        return False
    if check_only:
        print(f"Outline OK: {len(slides_data)} slides, "
              f"{len(warnings)} warning(s).")
        return True

    try:
        resolve_diagram_visuals(slides_data, ctx)
    except (RendererMissing, FileNotFoundError, RuntimeError) as e:
        print(f"[ERROR] {e}", file=sys.stderr)
        return False

    # Template mode + References + Exhibits: exhibit badges are skipped on
    # templated slides, so the "Exhibit N" numbers on the Sources slide may
    # not match the actual slide badges.  Warn once so the author is aware.
    if (template_path
            and meta.get("references", "").lower() in ("on", "true", "yes")
            and meta.get("exhibits", "").lower() in ("on", "true", "yes")):
        warn("References + Exhibits in template mode: exhibit numbers on the "
             "Sources slide may not match (templated slides skip badges)")

    if deck_has_cjk(slides_data):
        set_cjk(True)  # overlay CJK-safe fonts on every palette this build

    # precedence: CLI flag > outline front-matter > variant preset > default
    pal_default = get_palette(palette_key or meta.get("palette", ""))
    if theme_pal:
        # copy first: get_palette() hands back the registry's own dict, and
        # the theme override must not leak into the next build
        pal_default = {**pal_default,
                       **{k: v for k, v in theme_pal.items()
                          if not k.startswith("_")}}
    if meta.get("motif"):
        pal_default = {**pal_default, "motif": meta["motif"]}
    size = size or meta.get("size", "16:9")
    footer = meta.get("footer", "")
    page_numbers = meta.get("page_numbers", "").lower() in ("on", "true", "yes")
    exhibits_on = meta.get("exhibits", "").lower() in ("on", "true", "yes")
    exhibit_no = 0

    prs = create_presentation(template_path, size)
    import builders
    builders.set_canvas(prs)  # rescale styled builders to the actual slide size

    # **Page-Start:** N — native firstSlideNum plus offset cached page texts;
    # the divider page map feeds the agenda's page column (builders_agenda).
    page_start = parse_page_start(meta)
    if page_start != 1:
        set_first_slide_num(prs, page_start)
    ctx["_divider_pages"] = divider_page_map(slides_data, page_start)
    # **File-Stamp:** on — `<filename> · <date>` on every footered slide.
    file_stamp_text = None
    if meta.get("file_stamp", "").lower() in ("on", "true", "yes"):
        from datetime import date as _date
        file_stamp_text = (f"{Path(output_path).name} · "
                           f"{meta.get('date') or _date.today().isoformat()}")

    failures = 0
    built_slides = []
    current_section = None
    sections = [s.get("heading") or s.get("title", "") for s in slides_data
                if s.get("layout") == "section-divider"
                and not s.get("_appendix")]
    tracker_tabs = meta.get("tracker", "").lower() == "tabs" and bool(sections)
    tab_state = {}  # warn-once flag for the tab-budget fallback
    seen_keys = set()  # slide_tags.slide_key uniqueness, deck-wide
    logo_path = resolve_logo(ctx["_profile"].get("logo"), ctx)  # profile logo
    for i, p in enumerate(slides_data):
        pal = get_palette(p["palette"]) if "palette" in p else pal_default
        layout_name = p.get("layout", "bullet-list")
        if layout_name == "section-divider":
            current_section = p.get("heading") or p.get("title", "")
        print(f"  Building slide {i + 1}: "
              f"{p.get('heading', p.get('title', '?'))} [{layout_name}]")
        try:
            from helpers import set_footnote_markers
            set_footnote_markers(bool(p.get("footnotes")))
            slide = None
            templated = ghosted = False
            if template_path:
                slide = build_template_slide(prs, p, layout_name, ctx, template_config)
                templated = slide is not None
            if slide is None:
                if ghost and layout_name not in builders.GHOST_KEEP_REAL:
                    ghosted = True
                    slide = builders.build_ghost_slide(prs, p, pal, ctx)
                else:
                    slide = LAYOUT_MAP[layout_name](prs, p, pal, ctx)
            # Stamp the slide's identity (outline key, layout, hashes) into a
            # ppt/tags part so later tools can find it after a hand reorder.
            slide_tags.write_tags(slide, slide_tags.tags_for(
                {**p, "_ghost": ghosted, "_templated": templated},
                i + 1, ctx, seen_keys))
            if p.get("notes"):
                add_speaker_notes(slide, p["notes"])
            if str(p.get("hidden", "")).lower() in ("on", "true", "yes"):
                slide._element.set("show", "0")  # skipped in show and print
            section_label = "BACKUP" if p.get("_appendix") else current_section
            chrome_shapes = []
            if not templated and layout_name not in NO_FOOTER_LAYOUTS:
                chrome_shapes += add_footer(slide, i, pal, footer,
                                            page_numbers,
                                            appendix=p.get("_appendix", False),
                                            page_start=page_start)
                if file_stamp_text:
                    chrome_shapes += add_file_stamp(slide, pal,
                                                    file_stamp_text,
                                                    with_footer=bool(footer))
                if tracker_tabs and section_label \
                        and not p.get("_appendix") and not p.get("_auto"):
                    chrome_shapes += add_tracker_tabs(
                        slide, pal, sections, section_label, tab_state)
                else:
                    chrome_shapes += add_section_tracker(slide, pal,
                                                         section_label)
                if p.get("kicker"):
                    chrome_shapes += add_kicker_box(slide, pal, p["kicker"])
            if not templated and p.get("sticker"):
                chrome_shapes += add_sticker(
                    slide, pal, p["sticker"],
                    below_section=bool(section_label)
                    and layout_name not in NO_FOOTER_LAYOUTS)
            if not templated and p.get("footnotes"):
                chrome_shapes += add_footnotes(
                    slide, pal, p["footnotes"],
                    bottom=footnote_anchor(p))
            if not templated and (p.get("source") or p.get("note")):
                if p.get("source"):
                    exhibit_no += 1
                chrome_shapes += add_source(
                    slide, pal, p.get("source", ""),
                    exhibit_no if (exhibits_on and p.get("source")) else None,
                    note=p.get("note"))
            if not templated and meta.get("stamp"):
                chrome_shapes += add_stamp(slide, pal, meta["stamp"])
            if not templated and logo_path:
                chrome_shapes += add_logo(slide, layout_name, logo_path)
            chrome = {shape.shape_id: role for shape, role in chrome_shapes}
            slide_tags.name_shapes(slide, i + 1, chrome=chrome,
                                   templated=templated)
            built_slides.append(slide)
        except Exception as e:
            failures += 1
            print(f"  [ERROR] Slide {i + 1} failed: {e}", file=sys.stderr)
            import traceback
            traceback.print_exc()

    if failures:
        print(f"\nBuild aborted: {failures} slide(s) failed — no file written.",
              file=sys.stderr)
        return False

    # Post-build navigation pass: agenda rows + tracker chips jump to their
    # section divider (needs the built Slide objects, hence after the loop).
    import deck_nav
    nav_links = deck_nav.link_navigation(prs, built_slides, slides_data)
    if nav_links:
        print(f"  Navigation: {nav_links} hyperlink(s) wired")

    if not template_path:
        # template mode skipped: sections rely on the 1:1 slides_data ->
        # built-slide correspondence the styled pipeline guarantees
        section_names = inject_sections(prs, slides_data, meta)
        if section_names:
            print(f"  Sections: {', '.join(section_names)}")

    from save_hygiene import sanitize_presentation
    fixes = sanitize_presentation(
        prs, author=meta.get("author", ""),
        title=deck_title(slides_data, output_path, outline_path))
    print(f"  Hygiene: {len(fixes)} fix(es) applied")

    prs.save(output_path)
    print(f"\nSaved: {output_path} ({len(built_slides)} slides)")
    return True


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Build a .pptx from a markdown outline.")
    parser.add_argument("outline", help="Path to outline.md")
    parser.add_argument("--output", default="deck.pptx", help="Output .pptx path")
    parser.add_argument("--palette", default=None,
                        help="Color palette: built-ins "
                             f"({', '.join(sorted(PALETTES))}) or a custom "
                             "palette JSON name from <assets>/palettes/")
    parser.add_argument("--template", default=None, help="Optional .pptx template")
    parser.add_argument("--assets-dir", default=None,
                        help="Root dir for user-images/ and auto/ (default: ./assets)")
    parser.add_argument("--size", default=None, choices=sorted(SIZES),
                        help="Slide aspect (default 16:9)")
    parser.add_argument("--density", default=None,
                        choices=("compact", "comfortable"),
                        help="Content density (default compact)")
    parser.add_argument("--variant", default=None,
                        choices=("a", "b", "c", "consulting"),
                        help="Palette+density preset (a=default, b=aurora/comfortable)")
    parser.add_argument("--ghost", action="store_true",
                        help="Build a skeleton deck: real titles, placeholder exhibits (storyline alignment)")
    parser.add_argument("--style", default=None,
                        help="House-style JSON (word lists, title case/"
                             "punctuation); default <assets>/style.json")
    parser.add_argument("--profile", default=None,
                        help="Engagement profile name → defaults and "
                             "checks from <assets>/profiles/<name>.json")
    parser.add_argument("--verbatim", action="store_true",
                        help="Locked copy: render outline text byte-identical "
                             "(same as **Copy:** locked)")
    parser.add_argument("--check", action="store_true",
                        help="Validate the outline only; do not build")
    parser.add_argument("--titles", action="store_true",
                        help="Print slide titles in order (the consulting "
                             "'titles test' — they should read as an argument)")
    args = parser.parse_args()

    if args.titles:
        _, _slides = parse_outline(Path(args.outline).read_text(encoding="utf-8"))
        print("Titles test — read top to bottom; it should work as an argument:\n")
        for n, s in enumerate(_slides, 1):
            tag = " [BACKUP]" if s.get("_appendix") else ""
            print(f"  {n:2d}. {s.get('title') or s.get('heading', '?')}{tag}")
        sys.exit(0)

    ok = build(args.outline, args.output, args.palette,
               args.template, args.assets_dir, args.check, args.size,
               args.density, args.variant, ghost=args.ghost,
               style_path=args.style, profile=args.profile,
               verbatim=args.verbatim)
    sys.exit(0 if ok else 1)
