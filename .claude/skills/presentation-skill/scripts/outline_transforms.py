#!/usr/bin/env python3
"""
outline_transforms.py — deck-level transforms applied to the parsed slide list.

`apply_auto_agenda` (`**Auto-Agenda:** on`), `apply_references`
(`**References:** on` Sources slide), `apply_scale_groups`
(`**Scale-Group:** auto`), `resolve_diagram_visuals` (`diagram:` → PNG) and
`deck_has_cjk` (font-stack swap).  The agenda/references/scale-group trio is
run by `outline_replay.replay()` so every consumer sees the same slide list;
put any new outline-level transform there, not in build().

Split out of build_deck.py, which re-exports every name here.
"""
import re
from pathlib import Path

from helpers import parse_visual, resolve_image_path, warn

def _agenda_slide(sections, current=None):
    slide = {"bullets": list(sections), "stats": [], "cards": [], "items": [],
             "data": [], "table_rows": [], "steps": [], "tiles": [],
             "matrix_items": [], "bars": [], "milestones": [],
             "nodes": [], "left_bullets": [], "right_bullets": [],
             "layers": [], "lanes": [],
             "layout": "agenda", "heading": "Agenda",
             "notes": "Auto-generated agenda tracker.", "_auto": True}
    if current:
        slide["current"] = current
        # Include the section so repeated trackers keep unique titles
        # (qa_check --accessibility flags duplicate slide titles).
        slide["heading"] = f"Where we are — {current}"
    return slide


def _live_sections(slides):
    """Non-appendix section-divider headings whose chapter has content.

    A divider followed by another divider, an appendix slide or deck end has
    zero content slides — its agenda row is dropped with a warning (hand-
    written agenda bullets are never touched by this)."""
    sections = []
    for i, s in enumerate(slides):
        if s.get("layout") != "section-divider" or s.get("_appendix"):
            continue
        name = s.get("heading") or s.get("title", "")
        nxt = slides[i + 1] if i + 1 < len(slides) else None
        if nxt is None or nxt.get("layout") == "section-divider" \
                or nxt.get("_appendix"):
            warn(f"Auto-Agenda: empty chapter {name!r} dropped from the "
                 "agenda (no content slides after its divider)")
            continue
        sections.append(name)
    return sections


def _appendix_agenda(sections):
    slide = _agenda_slide(sections)
    slide["heading"] = "Appendix"
    slide["notes"] = "Auto-generated appendix splitter agenda."
    slide["_appendix"] = True
    return slide


def apply_auto_agenda(meta, slides):
    """**Auto-Agenda:** on  -> overview agenda after the title slide.
                       track -> + current-highlighted agenda after each divider
                                + an 'Appendix' splitter agenda before the
                                  first appendix slide (listing appendix
                                  section-dividers; none present -> skipped).
    Sections come from section-divider headings (appendix dividers excluded);
    empty chapters (divider followed by another divider or deck end) are
    dropped from the auto agendas with a warning, and no tracker agenda is
    inserted after their divider.  No dividers -> no-op.  The overview is
    only inserted when slide 1 is a title slide."""
    mode = meta.get("auto_agenda", "").lower()
    if mode not in ("on", "track"):
        return slides
    sections = _live_sections(slides)
    if not sections:
        return slides
    appendix_secs = [s.get("heading") or s.get("title", "")
                     for s in slides
                     if s.get("layout") == "section-divider"
                     and s.get("_appendix")] if mode == "track" else []
    out, splitter_done = [], False
    for i, s in enumerate(slides):
        if appendix_secs and not splitter_done and s.get("_appendix"):
            out.append(_appendix_agenda(appendix_secs))
            splitter_done = True
        out.append(s)
        if i == 0 and s.get("layout") == "title" and not s.get("_appendix"):
            out.append(_agenda_slide(sections))
        if mode == "track" and s.get("layout") == "section-divider" \
                and not s.get("_appendix") \
                and (s.get("heading") or s.get("title", "")) in sections:
            out.append(_agenda_slide(
                sections, current=s.get("heading") or s.get("title", "")))
    return out


def apply_agenda_columns(meta, slides, ctx=None):
    """Split the ` | ` time suffix off agenda bullets when the slide declares
    `- Columns:` — run inside outline_replay.replay() so every consumer
    (build, diff_deck, qa_checklist, refresh) sees the same row text the
    builder prints.  Times land in `p["times"]`, aligned with the bullets
    (no suffix -> "").  Verbatim decks are left untouched: locked copy
    renders byte-exact and builders_agenda ignores the directive there."""
    from verbatim import is_verbatim_deck
    if is_verbatim_deck(ctx, meta):
        return slides
    for p in slides:
        if p.get("layout") != "agenda" or not (p.get("columns") or "").strip():
            continue
        names, times = [], []
        for raw in p.get("bullets", []):
            name, time_slot = raw, ""
            if " | " in raw:
                name, _, time_slot = raw.rpartition(" | ")
            names.append(name.strip())
            times.append(time_slot.strip())
        p["bullets"], p["times"] = names, times
    return slides


def divider_page_map(slides, page_start=1):
    """Casefolded section-divider heading -> printed page label ('7', 'B·7')
    for the agenda page column (builders_agenda) — first divider with a given
    heading wins.  Call on the final replayed slide list so positions match
    what build() numbers."""
    pages = {}
    for i, s in enumerate(slides):
        if s.get("layout") != "section-divider":
            continue
        name = (s.get("heading") or s.get("title", "")).strip()
        if not name:
            continue
        num = str(i + page_start)
        pages.setdefault(name.casefold(),
                         f"B·{num}" if s.get("_appendix") else num)
    return pages


def _references_slide(bullets):
    return {"bullets": bullets, "stats": [], "cards": [], "items": [],
            "data": [], "table_rows": [], "steps": [], "tiles": [],
            "matrix_items": [], "bars": [], "milestones": [],
            "nodes": [], "left_bullets": [], "right_bullets": [],
            "layers": [], "lanes": [],
            "layout": "bullet-list", "heading": "Sources",
            "notes": "Auto-generated source register.",
            "_appendix": True, "_auto": True}


def apply_references(meta, slides):
    """**References:** on -> append an appendix 'Sources' slide aggregating
    unique '- Source:' values with their slide (or, with **Exhibits:** on,
    exhibit) numbers. Call after apply_auto_agenda so numbers match the deck.
    Default: off; no sources -> warn and no-op.

    Note: in template mode, build() skips exhibit badges on templated slides,
    so the exhibit numbers assigned here (counting every sourced slide) may
    not match the badges visible in the deck.  build() emits a warning when
    both template_path and Exhibits are active."""
    if meta.get("references", "").lower() not in ("on", "true", "yes"):
        return slides
    exhibits_on = meta.get("exhibits", "").lower() in ("on", "true", "yes")
    refs, exhibit_no = {}, 0  # source -> [numbers], encounter order
    for i, p in enumerate(slides, 1):
        if not p.get("source"):
            continue
        exhibit_no += 1
        refs.setdefault(p["source"], []).append(
            exhibit_no if exhibits_on else i)
    if not refs:
        warn("**References:** on but no '- Source:' lines found — "
             "skipping the Sources slide")
        return slides
    noun = "Exhibit" if exhibits_on else "Slide"
    bullets = [f"{noun}{'s' if len(nums) > 1 else ''} "
               f"{', '.join(str(n) for n in nums)} — {src}"
               for src, nums in refs.items()]
    return slides + [_references_slide(bullets)]


def resolve_diagram_visuals(slides, ctx):
    """Render `diagram:` visuals to PNG and rewrite the spec to the image.

    Runs after validate() in build(); raises on renderer failure so the
    fail-fast contract holds (no .pptx written on error)."""
    from render_diagram import render_diagram
    out_dir = Path(ctx.get("assets_dir", "assets")) / "auto" / "diagrams"
    for p in slides:
        for vis_key in ("visual", "visual_left", "visual_right"):
            spec = p.get(vis_key, "")
            kind, value = parse_visual(spec)
            if kind != "diagram":
                continue
            src = resolve_image_path("diagram:" + value, ctx)
            if src is None:
                raise FileNotFoundError(f"diagram source not found: {value!r}")
            png = render_diagram(src, out_dir=out_dir)
            p[vis_key] = str(png)


# Han, Hiragana+Katakana, Hangul syllables
CJK_RE = re.compile(r"[一-鿿぀-ヿ가-힯]")


def _text_values(obj):
    if isinstance(obj, str):
        yield obj
    elif isinstance(obj, dict):
        for k, v in obj.items():
            if not (isinstance(k, str) and k.startswith("_")):
                yield from _text_values(v)
    elif isinstance(obj, (list, tuple)):
        for v in obj:
            yield from _text_values(v)


def deck_has_cjk(slides):
    """Any CJK codepoint in any slide text field -> swap to a CJK font stack."""
    return any(CJK_RE.search(t) for s in slides for t in _text_values(s))


# Chart kinds where sharing one value-axis maximum is meaningful (a value
# axis exists and isn't a fixed 0-100% scale).
SCALE_GROUP_KINDS = ("bar", "column", "hbar", "line", "area")


def apply_scale_groups(meta, slides):
    """**Scale-Group:** auto — slides whose chart visuals share a kind get one
    common axis maximum (honest visual comparison). Slides with an explicit
    Axis-Max are left alone; groups need 2+ slides. Default: off."""
    if meta.get("scale_group", "").lower() != "auto":
        return slides
    groups = {}  # chart kind -> [(slide_no, slide_dict, data_max)]
    for i, p in enumerate(slides, 1):
        kind, value = parse_visual(p.get("visual", ""))
        if kind != "chart" or value not in SCALE_GROUP_KINDS \
                or p.get("axis_max"):
            continue
        nums = []
        for _, v in p.get("data", []):
            if isinstance(v, list):
                nums.extend(x for x in v if isinstance(x, (int, float)))
            elif isinstance(v, (int, float)):
                nums.append(v)
        if not nums:
            continue
        groups.setdefault(value, []).append((i, p, max(nums)))
    from charts import _nice_ceil
    for chart_kind, members in groups.items():
        if len(members) < 2:
            continue
        axis_max = _nice_ceil(max(m for _, _, m in members))
        for _, p, _ in members:
            p["axis_max"] = axis_max
        where = ", ".join(str(i) for i, _, _ in members)
        print(f"  Scale-group: {chart_kind} charts on slides {where} "
              f"share axis max {axis_max:g}")
    return slides

