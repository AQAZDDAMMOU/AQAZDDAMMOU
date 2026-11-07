#!/usr/bin/env python3
"""
gen_handout.py — Render a markdown outline as a readable pre-read handout.

Usage:
    python scripts/gen_handout.py outline.md [--output handout.md]
        [--assets-dir assets]

Default output: <outline-stem>-handout.md next to the outline. Pure markdown
transformation (slides -> numbered sections, **Data:** -> tables, notes ->
'Talk track' blockquotes, sources -> italics); no .pptx is touched.
"""
import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from build_deck import apply_auto_agenda, apply_references, parse_outline
from outline_replay import replay
from textfit import strip_markup


def _clean(text):
    """Plain-text form: drop icon: prefix and rich markers (**…**,
    {accent}…{/}) — deck-only markup that would render literally here."""
    return strip_markup(text)[0]


def _fmt_value(v):
    """42.0 -> '42', 4.2 -> '4.2', 'total' -> 'total'."""
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return str(v)


def _series_names(slide):
    return [s.strip() for s in slide.get("series", "").split(",") if s.strip()]


def _md_table(header, rows):
    lines = ["| " + " | ".join(header) + " |",
             "|" + " --- |" * len(header)]
    lines += ["| " + " | ".join(r) + " |" for r in rows]
    return lines


def _data_table(slide):
    """**Data:** rows -> markdown table (multi-series uses Series headers)."""
    data = slide.get("data", [])
    if not data:
        return []
    names = _series_names(slide)
    n_vals = max((len(v) for _, v in data if isinstance(v, list)), default=1)
    if n_vals > 1:
        headers = (names if len(names) == n_vals
                   else [f"Series {i + 1}" for i in range(n_vals)])
        header = ["Label"] + headers
    else:
        header = ["Label", names[0] if names else "Value"]
    rows = []
    for label, value in data:
        cells = ([_fmt_value(v) for v in value] if isinstance(value, list)
                 else [_fmt_value(value)])
        rows.append([str(label)] + cells)
    return _md_table(header, rows)


def _slide_section(p):
    """Body lines for one slide: bullets, kicker, data table, table, notes,
    source."""
    blocks = []
    bullets = [_clean(b) for b in p.get("bullets", [])]
    if bullets:
        blocks.append([f"- {b}" for b in bullets])
    if p.get("kicker"):
        blocks.append([f"**Takeaway:** {_clean(p['kicker'])}"])
    data_table = _data_table(p)
    if data_table:
        blocks.append(data_table)
    t_rows = p.get("table_rows", [])
    if t_rows:
        blocks.append(_md_table(t_rows[0], t_rows[1:]))
    if p.get("notes"):
        blocks.append([f"> Talk track: {p['notes']}"])
    if p.get("source"):
        src = p["source"]
        text = src if src.lower().startswith("source") else f"Source: {src}"
        blocks.append([f"*{text}*"])
    lines = []
    for block in blocks:
        lines += [""] + block
    return lines


def gen_handout(md_text, ctx=None):
    """Outline markdown -> handout markdown (same slide numbering as the deck)."""
    if ctx:
        # Data-File/evidence resolution needs a real outline directory to
        # resolve against; "outline.md" is a stand-in filename — only
        # ctx["outline_dir"] and ctx["assets_dir"] are actually used.
        outline_path = ctx["outline_dir"] / "outline.md"
        rr = replay(outline_path, ctx=ctx, md_text=md_text)
        meta, slides = rr.meta, rr.slides
    else:
        meta, slides = parse_outline(md_text)
        slides = apply_auto_agenda(meta, slides)
        slides = apply_references(meta, slides)
    out, in_appendix = [], False
    for n, p in enumerate(slides, 1):
        if p.get("_auto") and p.get("layout") == "agenda":
            continue
        if n == 1:
            out.append(
                f"# {_clean(p.get('title') or p.get('heading', 'Presentation'))}")
            if p.get("subtitle"):
                out += ["", _clean(p["subtitle"])]
            if p.get("layout") == "title":
                continue  # consumed as the document title
        if p.get("_appendix") and not in_appendix:
            out += ["", "---", "", "# Appendix"]
            in_appendix = True
        heading = _clean(p.get("heading") or p.get("title", ""))
        out += ["", f"## {n}. {heading}".rstrip()]
        out += _slide_section(p)
    return "\n".join(out).rstrip() + "\n"


def main():
    parser = argparse.ArgumentParser(
        description="Generate a pre-read markdown handout from an outline.")
    parser.add_argument("outline", help="Path to outline.md")
    parser.add_argument("--output", default=None,
                        help="Output path (default: <outline>-handout.md)")
    parser.add_argument("--assets-dir", default=None,
                        help="Root dir for Data-File resolution (default: ./assets)")
    args = parser.parse_args()

    outline = Path(args.outline)
    if not outline.is_file():
        print(f"  [ERROR] outline not found: {outline}", file=sys.stderr)
        return 1
    out_path = (Path(args.output) if args.output
                else outline.with_name(outline.stem + "-handout.md"))
    ctx = {"outline_dir": outline.parent,
           "assets_dir": Path(args.assets_dir) if args.assets_dir
           else Path("assets")}
    rr = replay(outline, ctx=ctx)
    for msg in rr.errors + rr.warnings:
        print(f"[WARN] {msg}", file=sys.stderr)
    text = gen_handout(outline.read_text(encoding="utf-8"), ctx)
    out_path.write_text(text, encoding="utf-8")
    print(f"Saved: {out_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
