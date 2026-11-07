#!/usr/bin/env python3
"""
qa_summary.py — one machine-readable QA record per built deck.

Runs the deterministic QA toolchain (qa_check, pptx_lint, save_hygiene's
package-hygiene check, geometry_report, and — when an outline is given —
diff_deck) and unifies the results with
deck statistics and the deck's editability level (`editability.py`) into a
single JSON delivery record:

    python3 scripts/qa_summary.py deck.pptx \\
        [--outline outline.md] [--palette midnight-executive] \\
        [--assets-dir assets] [--accessibility] [--gslides] \\
        [--output qa-summary.json] [--comments reviewed.pptx]

Without --output the record prints to stdout. With --comments every error and
warning is pinned to its slide as a PowerPoint comment in a copy of the deck.
Exit 1 when any tool reports an error (the delivery gate); warnings never fail
the gate. Rendering-based QA (render_slides + visual inspection) is
intentionally out of scope — this record captures everything checkable
without LibreOffice.
"""
import argparse
import json
import re
import sys
from datetime import datetime, timezone
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

SCHEMA_VERSION = 1
SLIDE_RX = re.compile(r"^Slide (\d+)\b")
DECK_WIDE_SLIDE = 1        # findings with no slide number land on slide 1
COMMENT_AUTHOR = "QA"


def deck_stats(prs):
    """Delivery-note statistics: what the deck is made of."""
    charts = tables = pictures = words = with_notes = 0
    from qa_check import iter_shapes
    for slide in prs.slides:
        if slide.has_notes_slide \
                and slide.notes_slide.notes_text_frame.text.strip():
            with_notes += 1
        for shape in iter_shapes(slide.shapes):
            if getattr(shape, "has_chart", False):
                charts += 1
            elif getattr(shape, "has_table", False):
                tables += 1
            elif shape.shape_type in (MSO_SHAPE_TYPE.PICTURE,
                                      MSO_SHAPE_TYPE.LINKED_PICTURE):
                pictures += 1
            if getattr(shape, "has_text_frame", False):
                words += len(shape.text_frame.text.split())
    from pptx_lint import collect_inventory
    fonts, colors, sizes = collect_inventory(prs)
    return {
        "slides": len(prs.slides),
        "native_charts": charts,
        "tables": tables,
        "pictures": pictures,
        "slides_with_notes": with_notes,
        "words": words,
        "fonts": sorted(fonts),
        "distinct_colors": len(colors),
        "distinct_font_sizes": len(sizes),
    }


def comment_counts(pptx_path):
    """Reviewer comments still on the deck (replies counted individually).

    A delivered deck should carry none; open threads are unfinished revision
    work, resolved ones are just litter (`edit_deck comments strip`).

    A comment part the deck carries but nothing can parse is reported, not
    raised: `{"open": None, "resolved": None, "error": …}`, which
    `build_record` turns into a warning. Losing the comment count must never
    cost the delivery its whole QA record.
    """
    from pptx_comments import read_comments
    try:
        comments = read_comments(pptx_path)
    except ValueError as exc:
        return {"open": None, "resolved": None, "error": str(exc)}
    resolved = sum(1 for c in comments if c.resolved)
    return {"open": len(comments) - resolved, "resolved": resolved}


def run_geometry(prs):
    """geometry_report findings keyed by slide number (findings-only)."""
    from geometry_report import analyze_deck, findings
    per_slide = {}
    for n, metrics in analyze_deck(prs).items():
        lines = findings(metrics)
        if lines:
            per_slide[str(n)] = lines
    return {"findings": per_slide,
            "total": sum(len(v) for v in per_slide.values())}


def build_record(pptx_path, outline=None, palette=None, assets_dir="assets",
                 accessibility=False, gslides=False, profile=None):
    """Assemble the full QA record dict for a deck.

    `profile` is the engagement profile the deck was built with (the same
    name `--profile` files lessons under); diff_deck replays the outline
    with it so profile-driven layouts/density do not read as diff noise."""
    import editability
    from palettes import load_custom_palettes
    load_custom_palettes(Path(assets_dir) / "palettes")
    prs = Presentation(str(pptx_path))

    from qa_check import CHART_UNREADABLE, check_deck
    from pptx_lint import lint_deck
    from save_hygiene import check_package_hygiene
    qa = check_deck(pptx_path, accessibility=accessibility)
    lint = lint_deck(prs, palette, gslides=gslides)
    # qa_check and pptx_lint each independently detect an unreadable chart
    # (both go through qa_check.safe_chart); drop pptx_lint's copy when
    # qa_check already reports the identical message for that slide.
    qa_unreadable = {w for w in qa["warn"] if CHART_UNREADABLE in w}
    lint_warnings = [w for w in lint["warn"]
                     if not (CHART_UNREADABLE in w and w in qa_unreadable)]
    checks = {
        "qa_check": {"errors": qa["error"], "warnings": qa["warn"]},
        "pptx_lint": {"errors": lint["error"], "warnings": lint_warnings},
        # Package hygiene is dependency-free and always on: a Keynote-hostile
        # slide size or leaked python-pptx core properties must fail the gate,
        # not just `qa_check --integrity`.
        "hygiene": {"errors": check_package_hygiene(pptx_path),
                    "warnings": []},
        "geometry": run_geometry(prs),
        "diff_deck": None,
    }
    if outline:
        from diff_deck import diff_outline
        d_errors, d_warnings = diff_outline(outline, pptx_path,
                                            profile=profile,
                                            assets_dir=assets_dir)
        checks["diff_deck"] = {"errors": d_errors, "warnings": d_warnings}

    stats = deck_stats(prs)
    stats["comments"] = comment_counts(pptx_path)
    # additive check (schema stays 1): the record has to be able to say that
    # the comment count is missing, and why, without a traceback
    checks["comments"] = {"errors": [], "warnings": []}
    if stats["comments"].get("error"):
        checks["comments"]["warnings"].append(
            "reviewer comments could not be read: "
            f"{stats['comments']['error']}")

    n_errors = sum(len(c["errors"]) for c in checks.values()
                   if isinstance(c, dict) and "errors" in c)
    n_warnings = sum(len(c["warnings"]) for c in checks.values()
                     if isinstance(c, dict) and "warnings" in c)
    return {
        "schema": SCHEMA_VERSION,
        "deck": Path(pptx_path).name,
        "generated_at": datetime.now(timezone.utc).isoformat(
            timespec="seconds"),
        "palette": palette,
        "outline": str(outline) if outline else None,
        "stats": stats,
        "editability": editability.deck_level(prs),
        "checks": checks,
        "errors": n_errors,
        "warnings": n_warnings,
        "geometry_findings": checks["geometry"]["total"],
        "pass": n_errors == 0,
    }


def _comment_for(check, level, message):
    """One (slide_number, text) comment; deck-wide findings go to slide 1."""
    m = SLIDE_RX.match(message)
    if m:
        return int(m.group(1)), f"[{check}] {level}: {message}"
    return DECK_WIDE_SLIDE, f"[{check}] {level}: deck: {message}"


def collect_comments(record):
    """Every error, warning and geometry finding as (slide_number, text)."""
    comments = []
    for check, result in (record.get("checks") or {}).items():
        if not isinstance(result, dict):
            continue
        for level, key in (("ERROR", "errors"), ("WARNING", "warnings")):
            for message in result.get(key) or []:
                comments.append(_comment_for(check, level, message))
        for slide_no, lines in (result.get("findings") or {}).items():
            for line in lines:
                comments.append((int(slide_no), f"[{check}] {line}"))
    return comments


def _profile_exists(name, assets_dir):
    """True when <assets>/profiles/<name>.json exists and the name is safe."""
    import profiles
    try:
        return profiles.profile_path(name, assets_dir).is_file()
    except ValueError:
        return False        # unsafe name — reported as 'not found'


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("pptx")
    parser.add_argument("--outline", default=None,
                        help="Outline .md to diff deck content against")
    parser.add_argument("--palette", default=None,
                        help="Palette key(s) for the lint whitelist "
                             "(comma-separate for per-slide overrides)")
    parser.add_argument("--assets-dir", default="assets",
                        help="Root assets dir (custom palettes load from "
                             "<assets-dir>/palettes/)")
    parser.add_argument("--accessibility", action="store_true",
                        help="Run qa_check in WCAG AA strict mode")
    parser.add_argument("--gslides", action="store_true",
                        help="Include Google Slides compatibility warnings")
    parser.add_argument("--output", default=None,
                        help="Write the JSON record here (default: stdout)")
    parser.add_argument("--profile", default=None,
                        help="Engagement profile name — recurring findings "
                             "are appended to <assets>/profiles/"
                             "<name>.lessons.md")
    parser.add_argument("--comments", default=None,
                        help="Write a copy of the deck with every finding "
                             "attached to its slide as a PowerPoint comment")
    args = parser.parse_args()

    if not Path(args.pptx).is_file():
        print(f"ERROR: {args.pptx} not found", file=sys.stderr)
        sys.exit(2)
    if args.profile and not _profile_exists(args.profile, args.assets_dir):
        print(f"ERROR: profile {args.profile!r} not found under "
              f"{Path(args.assets_dir) / 'profiles'}/", file=sys.stderr)
        sys.exit(2)
    record = build_record(args.pptx, args.outline, args.palette,
                          args.assets_dir, args.accessibility, args.gslides,
                          profile=args.profile)
    payload = json.dumps(record, indent=2, ensure_ascii=False)
    if args.output:
        Path(args.output).write_text(payload + "\n", encoding="utf-8")
        print(f"Wrote {args.output}")
    else:
        print(payload)
    if args.profile:
        import profiles
        lines = profiles.record_lessons(record, args.profile,
                                        args.assets_dir)
        md_path, _ = profiles.lessons_paths(args.profile, args.assets_dir)
        print(f"Appended {len(lines)} lesson(s) to {md_path}", file=sys.stderr)
    if args.comments:
        from pptx_comments import add_comments
        written = add_comments(args.pptx, args.comments,
                               collect_comments(record), author=COMMENT_AUTHOR)
        print(f"Wrote {written} comment(s) to {args.comments}", file=sys.stderr)
    verdict = "PASS" if record["pass"] else "FAIL"
    print(f"\n{record['deck']}: {verdict} — {record['errors']} error(s), "
          f"{record['warnings']} warning(s), "
          f"{record['geometry_findings']} geometry finding(s)",
          file=sys.stderr)
    sys.exit(0 if record["pass"] else 1)


if __name__ == "__main__":
    main()
