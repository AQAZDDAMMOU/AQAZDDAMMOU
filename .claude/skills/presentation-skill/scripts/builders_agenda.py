"""builders_agenda.py — the agenda layout (chapter rows with optional
time/page columns, sub-chapter rows, current-section highlight).

Moved out of builders_consulting.py (which re-exports `build_agenda_slide`;
its LAYOUTS entry for "agenda" points here, so builders.LAYOUT_MAP picks the
move up unchanged).

Cycle 7 additions (all opt-in — a plain agenda renders exactly as before):
  - `- Columns: time, page` — right-aligned columns per row, listed order =
    display order.  Time comes from a ` | ` suffix on the section bullet
    (`- Strategic context | 09:00–09:20`); page is the printed page number of
    the section's divider slide, resolved case-insensitively through the
    `ctx["_divider_pages"]` map build_deck.build() fills from
    outline_transforms.divider_page_map() (unresolvable → blank cell).
  - Sub-chapters: indented outline bullets (2+ spaces; the parser records
    `p["_bullet_indent"]`) render as smaller un-numbered muted rows under
    their chapter, with no highlight logic of their own.

All geometry is in design inches on the 13.33 x 7.5 canvas; builders'
wrappers rescale to the actual slide size.
"""
from pptx.enum.text import PP_ALIGN

import builders as B
import shape_groups
from helpers import is_verbatim, warn

# Max rendered rows (chapters + sub-chapters) — mirrored by
# diff_deck.LAYOUT_CONTENT["agenda"].
AGENDA_MAX_ROWS = 8
# Width of one time/page column (design inches); rows end at x 11.5.
AGENDA_COL_W = 1.5
AGENDA_ROW_RIGHT = 11.5
COLUMN_KEYS = ("time", "page")


def parse_agenda_columns(p):
    """`- Columns: time, page` → ordered subset of COLUMN_KEYS.

    Unknown names warn and are ignored; under verbatim the directive is
    ignored entirely (splitting the ` | ` time suffix off a bullet would
    change locked copy)."""
    value = (p.get("columns") or "").strip()
    if not value:
        return []
    if is_verbatim():
        warn("agenda Columns ignored under verbatim copy mode — rows render "
             "the outline bullets byte-exact")
        return []
    cols = []
    for token in value.split(","):
        token = token.strip().lower()
        if not token:
            continue
        if token in COLUMN_KEYS:
            if token not in cols:
                cols.append(token)
        else:
            warn(f"agenda Columns: unknown column {token!r} ignored "
                 f"(supported: {', '.join(COLUMN_KEYS)})")
    return cols


def agenda_rows(p):
    """[(name, time_slot, is_sub)] per agenda bullet, capped at
    AGENDA_MAX_ROWS.  The ` | ` time suffixes were already split into
    `p["times"]` by outline_transforms.apply_agenda_columns during replay
    (so diff/QA see the same row text this builder prints)."""
    indents = p.get("_bullet_indent", [])
    times = p.get("times", [])
    rows = []
    for i, name in enumerate(p.get("bullets", [])):
        rows.append((name,
                     times[i] if i < len(times) else "",
                     bool(indents[i]) if i < len(indents) else False))
    return rows[:AGENDA_MAX_ROWS]


def build_agenda_slide(prs, p, pal, ctx):
    cols = parse_agenda_columns(p)
    rows = agenda_rows(p)
    current = p.get("current", "").strip().lower()
    slide = B._blank_slide(prs, pal, pal["bg_deep"], gradient=True)
    B.add_tb(slide, p.get("heading", "Agenda"), 0.8, 0.7, 11.7, 1.0, size=34,
             bold=True, color=pal["accent1"], font=pal["font_title"])
    body_start = len(slide.shapes)
    row_h = min(0.75, 5.0 / max(len(rows), 1))
    pages = (ctx or {}).get("_divider_pages", {})
    name_w = 9.5 - AGENDA_COL_W * len(cols)
    num = 0
    for i, (name, time_slot, is_sub) in enumerate(rows):
        y = 2.0 + i * (row_h + 0.12)
        if is_sub:
            B.add_tb(slide, name, 2.4, y, name_w - 0.4, 0.5, size=14,
                     color=pal["text_muted"], font=pal["font_body"])
        else:
            num += 1
            is_cur = current and current in name.lower()
            if is_cur:
                B.add_rect(slide, 0.8, y - 0.06, 11.0, row_h, pal["surface"])
                B.add_rect(slide, 0.8, y - 0.06, 0.08, row_h, pal["accent1"])
            B.add_tb(slide, f"{num:02d}", 1.1, y, 0.8, 0.5, size=16, bold=True,
                     color=pal["accent1"] if is_cur else pal["text_muted"],
                     font=pal["font_title"])
            B.add_tb(slide, name, 2.0, y, name_w, 0.5, size=17, bold=is_cur,
                     color=pal["text"] if is_cur else pal["text_muted"],
                     font=pal["font_body"])
        for j, col in enumerate(cols):
            value = time_slot if col == "time" \
                else pages.get(name.casefold(), "")
            if value:
                x = AGENDA_ROW_RIGHT - AGENDA_COL_W * (len(cols) - j)
                B.add_tb(slide, value, x, y, AGENDA_COL_W - 0.1, 0.5, size=13,
                         color=pal["text_muted"], font=pal["font_label"],
                         align=PP_ALIGN.RIGHT)
    shape_groups.group_shapes(slide, list(slide.shapes)[body_start:])
    return slide


LAYOUTS = {
    "agenda": build_agenda_slide,
}
