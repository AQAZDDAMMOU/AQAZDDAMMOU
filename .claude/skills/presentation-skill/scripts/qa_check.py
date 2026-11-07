#!/usr/bin/env python3
"""
qa_check.py — Programmatic QA for a generated .pptx.

Usage:
    python3 scripts/qa_check.py deck.pptx
    python3 scripts/qa_check.py deck.pptx --text
    python3 scripts/qa_check.py deck.pptx --accessibility   # WCAG AA strict mode
    python3 scripts/qa_check.py deck.pptx --integrity       # package audit

Checks include bounds, placeholders, contrast, chart transparency, unique slide
titles, projection font sizes, image alt text, text density, and editability
(text boxes parked on same-bounds shapes, unanchored connectors, tiny text on
mostly empty slides — see editability.py). --integrity
adds package-hygiene checks (Keynote-safe slide size, notesMasterIdLst,
printerSettings leftovers, python-pptx stock core properties — see
save_hygiene.py) plus OOXML schema validation via the optional openxml-audit
package (module or CLI); when it is not installed that step prints an info
line and is skipped.
"""
import re
import shutil
import subprocess
import sys
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.dml import MSO_FILL
from pptx.exc import PythonPptxError
from pptx.oxml.ns import qn
from pptx.util import Emu

import editability
import textfit

PLACEHOLDER_RX = re.compile(
    r"lorem|ipsum|xxxx|click to add|\[ add visual \]|placeholder"
    r"|this.*(page|slide).*layout", re.I)
ALT_FILENAME_RX = re.compile(
    r".*\.(png|jpe?g|gif|bmp|svg)$|^image\d*$", re.I)
MIN_PT = 10
PROJECTION_BODY_PT = 18  # federal / Section 508 presentation guidance
DEFAULT_BODY_PT = 18     # assumed size for a run that carries no explicit size
LR_INSET_IN = 0.1        # OOXML bodyPr lIns/rIns default
TB_INSET_IN = 0.05       # OOXML bodyPr tIns/bIns default
OVERFLOW_BAND = 1.30     # warn once the estimate exceeds the box by >30%
EMU_IN = 914400
MAX_WORDS_PER_SLIDE = 220
# Raised when a graphic frame's chart part is unreachable (dangling r:id,
# stripped relationship, truncated package).
CHART_READ_ERRORS = (PythonPptxError, KeyError, ValueError, AttributeError)
CHART_UNREADABLE = "chart could not be read"


def _luminance(rgb):
    def chan(c):
        c = c / 255
        return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
    r, g, b = rgb
    return 0.2126 * chan(r) + 0.7152 * chan(g) + 0.0722 * chan(b)


def contrast_ratio(rgb1, rgb2):
    l1, l2 = sorted((_luminance(rgb1), _luminance(rgb2)), reverse=True)
    return (l1 + 0.05) / (l2 + 0.05)


def slide_bg_rgb(slide):
    try:
        fill = slide.background.fill
        if fill.type == MSO_FILL.SOLID:
            c = fill.fore_color.rgb
            return (c[0], c[1], c[2])
        if fill.type == MSO_FILL.GRADIENT:
            stops = [s.color.rgb for s in fill.gradient_stops]
            return tuple(sum(c[i] for c in stops) // len(stops) for i in range(3))
    except (TypeError, ValueError, AttributeError):
        pass
    return None


def iter_shapes(shapes):
    for shape in shapes:
        yield shape
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from iter_shapes(shape.shapes)


def safe_chart(shape):
    """`shape.chart`, or None when the chart part cannot be reached.

    A graphic frame can outlive its chart part — a dangling `r:id`, a rel
    stripped by a hand edit, a truncated package. python-pptx raises out of the
    part lookup in that case, which used to abort a whole lint run over one
    broken relationship. Callers degrade to a warning instead.
    """
    if not getattr(shape, "has_chart", False):
        return None
    try:
        return shape.chart
    except CHART_READ_ERRORS:
        return None


def _slide_title(slide):
    """Best-effort (shape, text) title: shapes.title, else largest run."""
    try:
        if slide.shapes.title and slide.shapes.title.text.strip():
            return slide.shapes.title, slide.shapes.title.text.strip()
    except (AttributeError, ValueError):
        pass
    best_shape, best, best_size = None, "", 0
    for shape in iter_shapes(slide.shapes):
        if not getattr(shape, "has_text_frame", False):
            continue
        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                if not run.text.strip():
                    continue
                size = run.font.size.pt if run.font.size else 0
                if size >= best_size:
                    best_size = size
                    best = run.text.strip()
                    best_shape = shape
    return best_shape, best


def _slide_title_text(slide):
    """Best-effort title: largest bold text or shapes.title."""
    return _slide_title(slide)[1]


def _solid_fill_rgb(shape):
    try:
        if shape.fill.type == MSO_FILL.SOLID:
            c = shape.fill.fore_color.rgb
            return (c[0], c[1], c[2])
    except (TypeError, ValueError, AttributeError, NotImplementedError):
        pass
    return None


# Pictures at least this big can lie under text (then the background is
# unknowable and contrast is skipped); smaller ones are icons/logos.
PICTURE_BG_MIN_SQIN = 2.0


def _picture_area_sqin(shape):
    try:
        return (shape.width / EMU_IN) * (shape.height / EMU_IN)
    except (TypeError, AttributeError):
        return float("inf")      # unknown size: stay conservative


def effective_bg(text_shape, filled_below, slide_bg):
    """Background a text box actually sits on: the topmost solid-filled shape
    under its center (cards, chevrons, number circles), else the slide bg."""
    try:
        cx = text_shape.left + text_shape.width // 2
        cy = text_shape.top + text_shape.height // 2
    except (TypeError, AttributeError):
        return slide_bg
    bg = slide_bg
    for left, top, w, h, rgb in filled_below:  # document order == z-order
        if left <= cx <= left + w and top <= cy <= top + h:
            bg = rgb
    return bg


def check_contrast(shape, bg_rgb, issues, where, accessibility=False):
    if bg_rgb is None or not getattr(shape, "has_text_frame", False):
        return
    for para in shape.text_frame.paragraphs:
        for run in para.runs:
            if not run.text.strip():
                continue
            try:
                c = run.font.color.rgb
            except (TypeError, AttributeError):
                continue
            if c is None:
                continue
            ratio = contrast_ratio((c[0], c[1], c[2]), bg_rgb)
            size = run.font.size.pt if run.font.size else 16
            bold = bool(run.font.bold)
            large = size >= 18 or (size >= 14 and bold)
            min_ratio = 4.5 if accessibility and not large else 3.0
            if ratio < min_ratio:
                issues["error"].append(
                    f"{where}: contrast {ratio:.1f}:1 (<{min_ratio}:1) for "
                    f"{size:.0f}pt text: {run.text[:40]!r}")
            elif ratio < 4.5 and not large:
                issues["warn"].append(
                    f"{where}: contrast {ratio:.1f}:1 (<4.5:1) for small "
                    f"{size:.0f}pt text: {run.text[:40]!r}")


def check_bounds(shape, sw, sh, issues, where):
    try:
        left, top, w, h = shape.left, shape.top, shape.width, shape.height
    except (AttributeError, TypeError):
        return
    if None in (left, top, w, h):
        return
    tol = Emu(int(0.02 * EMU_IN))
    if left < -tol or top < -tol or left + w > sw + tol \
            or top + h > sh + tol:
        issues["error"].append(
            f"{where}: shape outside slide bounds "
            f"({left / EMU_IN:.2f},{top / EMU_IN:.2f} "
            f"{w / EMU_IN:.2f}x{h / EMU_IN:.2f})")


def _inset_in(value, default_in):
    """A bodyPr inset in inches, falling back to the OOXML default."""
    return default_in if value is None else value / EMU_IN


def check_text(shape, issues, where, accessibility=False):
    """Placeholder text, tiny/projection font sizes, and overflow.

    The overflow estimate measures each paragraph with textfit — the palette's
    own font when build_deck registered one — instead of a character-width
    heuristic, so the warning tracks what PowerPoint will actually render.
    """
    if not getattr(shape, "has_text_frame", False):
        return False
    tf = shape.text_frame
    has_text = False
    min_size = None
    for para in tf.paragraphs:
        for run in para.runs:
            if not run.text.strip():
                continue
            has_text = True
            if PLACEHOLDER_RX.search(run.text):
                issues["error"].append(
                    f"{where}: leftover placeholder text: {run.text[:60]!r}")
            if run.font.size is not None:
                pt = run.font.size.pt
                min_size = pt if min_size is None else min(min_size, pt)
                if pt < MIN_PT:
                    issues["error"].append(
                        f"{where}: {pt:.0f}pt text (<{MIN_PT}pt): {run.text[:40]!r}")
                elif accessibility and pt < PROJECTION_BODY_PT and not run.font.bold:
                    # projection-size guidance is opt-in: consulting decks are
                    # read documents, and compact density is intentional
                    issues["error"].append(
                        f"{where}: body text {pt:.0f}pt (<{PROJECTION_BODY_PT}pt "
                        f"for projection): {run.text[:40]!r}")

    if has_text and min_size and shape.width and shape.height \
            and len(tf.text.strip()) > 2:  # skip decorative glyphs (e.g. big quote marks)
        usable_w = (shape.width / EMU_IN
                    - _inset_in(tf.margin_left, LR_INSET_IN)
                    - _inset_in(tf.margin_right, LR_INSET_IN))
        est_h = (_inset_in(tf.margin_top, TB_INSET_IN)
                 + _inset_in(tf.margin_bottom, TB_INSET_IN))
        est_lines = 0
        # insets wider than the box leave nothing to wrap into: no copy fits
        no_room = usable_w <= 0
        for p in tf.paragraphs:
            if not p.text.strip():
                continue
            if no_room:
                est_lines += 1
                continue
            sizes = [r.font.size.pt for r in p.runs if r.font.size is not None]
            pt = min(sizes) if sizes else DEFAULT_BODY_PT
            family = p.runs[0].font.name if p.runs else None
            bold = any(bool(r.font.bold) for r in p.runs)
            lines = max(textfit.estimate_lines(p.text, pt, usable_w, bold,
                                               family=family), 1)
            est_lines += lines
            est_h += lines * pt * textfit.LINE_PITCH / 72.0
        if no_room or est_h > (shape.height / EMU_IN) * OVERFLOW_BAND:
            issues["warn"].append(
                f"{where}: text likely overflows box (~{est_lines} lines): "
                f"{tf.text[:50]!r}")
    return has_text


def check_chart_bg(shape, issues, where):
    if not getattr(shape, "has_chart", False):
        return False
    chart = safe_chart(shape)
    if chart is None:
        issues["warn"].append(f"{where}: {CHART_UNREADABLE}")
        return True
    cs = chart._chartSpace
    spPr = cs.find(qn("c:spPr"))
    transparent = spPr is not None and spPr.find(qn("a:noFill")) is not None
    if not transparent:
        issues["error"].append(
            f"{where}: chart has opaque background (white-box-on-dark defect)")
    return True


def check_image_alt(shape, issues, where, accessibility=False):
    if shape.shape_type != MSO_SHAPE_TYPE.PICTURE:
        return False
    descr = shape._element._nvXxPr.cNvPr.get("descr", "")
    if not descr or not descr.strip():
        msg = f"{where}: image missing alt text (descr attribute)"
        (issues["error"] if accessibility else issues["warn"]).append(msg)
    elif accessibility and ALT_FILENAME_RX.match(descr.strip()):
        issues["error"].append(
            f"{where}: alt text is a filename ({descr.strip()!r}) — "
            "describe the image instead")
    return True


def _tc_merged(tc):
    def _span(name):
        v = tc.get(name)
        try:
            return int(v) if v else 1
        except ValueError:
            return 1
    if _span("gridSpan") > 1 or _span("rowSpan") > 1:
        return True
    return tc.get("hMerge") in ("1", "true") or tc.get("vMerge") in ("1", "true")


def check_table_a11y(shape, issues, where):
    """MS Accessibility Checker table rules: marked header row, no merges."""
    if not getattr(shape, "has_table", False):
        return
    tbl = shape.table._tbl
    tblPr = tbl.tblPr
    if tblPr is None or tblPr.get("firstRow") not in ("1", "true"):
        issues["error"].append(
            f"{where}: table not marked with a header row (firstRow) — "
            "screen readers cannot announce column headers")
    if any(_tc_merged(tc) for tc in tbl.iter(qn("a:tc"))):
        issues["warn"].append(
            f"{where}: table has merged cells — screen readers may misread; "
            "prefer simple structure")


def check_reading_order(slide, issues, where):
    """Title shape should be the first text-bearing shape in spTree order."""
    title_shape, _ = _slide_title(slide)
    if title_shape is None:
        return
    first_text = next(
        (s for s in iter_shapes(slide.shapes)
         if getattr(s, "has_text_frame", False)
         and s.text_frame.text.strip()), None)
    if first_text is not None and first_text.shape_id != title_shape.shape_id:
        issues["warn"].append(
            f"{where}: title is not first in reading order — screen readers "
            "announce it after other content")


def check_deck(pptx_path, accessibility=False):
    prs = Presentation(str(pptx_path))
    sw, sh = prs.slide_width, prs.slide_height
    issues = {"error": [], "warn": []}
    titles_seen = {}

    for n, slide in enumerate(prs.slides, 1):
        where = f"Slide {n}"
        has_text = has_visual = False
        shape_count = word_count = 0
        # A picture that could sit under text makes the background
        # unknowable; an icon-sized one (the profile logo, bullet markers)
        # cannot, so it must not switch the contrast check off.
        has_picture = any(s.shape_type == MSO_SHAPE_TYPE.PICTURE
                          and _picture_area_sqin(s) >= PICTURE_BG_MIN_SQIN
                          for s in iter_shapes(slide.shapes))
        bg_rgb = None if has_picture else slide_bg_rgb(slide)

        title = _slide_title_text(slide)
        if not title:
            msg = f"{where}: no detectable slide title"
            (issues["error"] if accessibility else issues["warn"]).append(msg)
        else:
            titles_seen.setdefault(title.strip().casefold(), []).append(n)
        if accessibility:
            check_reading_order(slide, issues, where)

        filled_below = []  # accumulating: shapes earlier in doc order are below
        for shape in iter_shapes(slide.shapes):
            shape_count += 1
            check_bounds(shape, sw, sh, issues, where)
            if check_text(shape, issues, where, accessibility):
                has_text = True
                word_count += len(shape.text_frame.text.split())
            # an autoshape carrying its own text (number badge, chevron) is
            # read against its own fill, not the shape behind it
            fill_rgb = _solid_fill_rgb(shape)
            if bg_rgb is None:
                text_bg = None
            elif fill_rgb is not None:
                text_bg = fill_rgb
            else:
                text_bg = effective_bg(shape, filled_below, bg_rgb)
            check_contrast(shape, text_bg, issues, where, accessibility)
            if fill_rgb is not None and None not in (
                    getattr(shape, "left", None), getattr(shape, "top", None),
                    getattr(shape, "width", None), getattr(shape, "height", None)):
                filled_below.append((shape.left, shape.top, shape.width,
                                     shape.height, fill_rgb))
            if check_chart_bg(shape, issues, where):
                has_visual = True
            if check_image_alt(shape, issues, where, accessibility):
                has_visual = True
            if accessibility:
                check_table_a11y(shape, issues, where)
            if shape.shape_type in (MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.TABLE):
                has_visual = True
            if shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE and not getattr(
                    shape, "has_text_frame", False):
                has_visual = True

        editability.check_slide(slide, sw, sh, issues, where)

        if shape_count == 0:
            issues["warn"].append(f"{where}: slide is empty")
        elif has_text and not has_visual and shape_count <= 2:
            issues["warn"].append(
                f"{where}: text-only slide — add icon, chart, or image")
        if word_count > MAX_WORDS_PER_SLIDE:
            issues["warn"].append(
                f"{where}: {word_count} words — dense slide, consider splitting")

    for key, ns in titles_seen.items():
        if len(ns) > 1:
            where_list = ", ".join(str(s) for s in ns)
            msg = (f"duplicate slide title {key!r} on slides {where_list} — "
                   "titles must be unique for navigation")
            (issues["error"] if accessibility else issues["warn"]).append(msg)

    return issues


def dump_text(pptx_path):
    prs = Presentation(str(pptx_path))
    for n, slide in enumerate(prs.slides, 1):
        print(f"\n=== Slide {n} ===")
        for shape in iter_shapes(slide.shapes):
            if getattr(shape, "has_text_frame", False) and shape.text_frame.text.strip():
                print(shape.text_frame.text)
        if slide.has_notes_slide and slide.notes_slide.notes_text_frame.text.strip():
            print(f"[notes] {slide.notes_slide.notes_text_frame.text}")


NUM_TOKEN_RX = re.compile(
    r"[$€£]?\d[\d,.]*(?:%|bps|bn|B|M|k|x|pp)?|FY\d{2,4}|Q[1-4]\s?\d{2,4}",
    re.I)


def dump_numbers(pptx_path):
    """All numeric/period tokens per slide — input for the LLM
    cross-slide consistency check (see references/qa-guide.md)."""
    prs = Presentation(str(pptx_path))
    for n, slide in enumerate(prs.slides, 1):
        tokens = []
        for shape in iter_shapes(slide.shapes):
            if getattr(shape, "has_text_frame", False):
                tokens += [t.strip() for t in
                           NUM_TOKEN_RX.findall(shape.text_frame.text)
                           if t.strip()]
        title = _slide_title_text(slide)
        print(f"Slide {n} [{title[:60]}]: {', '.join(tokens) if tokens else '—'}")


def _integrity_via_module(mod, pptx_path):
    """Validate via an importable openxml_audit module. The package's API is
    probed defensively (validate/audit/check, called with the file path and
    expected to return an iterable of error strings/objects). Returns a list
    of error strings, or None when no recognized entry point exists."""
    fn = next((getattr(mod, name) for name in ("validate", "audit", "check")
               if callable(getattr(mod, name, None))), None)
    if fn is None:
        print("  [WARN] openxml-audit installed but exposes no "
              "validate/audit/check entry point — integrity check skipped",
              file=sys.stderr)
        return None
    try:
        result = fn(str(pptx_path))
    except Exception as exc:
        print(f"  [WARN] openxml-audit module call failed ({exc}); "
              "falling back to CLI/skip", file=sys.stderr)
        return None
    return [str(item) for item in result] if result else []


def _integrity_via_cli(cli, pptx_path):
    """Validate via an `openxml-audit <file>` CLI on PATH. Non-zero exit →
    each non-empty output line becomes an integrity error."""
    try:
        r = subprocess.run([cli, str(pptx_path)],
                           capture_output=True, text=True, timeout=120)
    except (subprocess.TimeoutExpired, OSError) as exc:
        return [f"openxml-audit CLI failed ({exc.__class__.__name__})"]
    if r.returncode == 0:
        return []
    lines = [ln.strip() for ln in (r.stdout + "\n" + r.stderr).splitlines()
             if ln.strip()]
    return lines or [f"openxml-audit exited {r.returncode}"]


def run_integrity(pptx_path):
    """Package hygiene + OOXML schema validation.

    Always runs `save_hygiene.check_package_hygiene()` (deterministic, no
    optional dependency) and prepends its `hygiene: …` errors, then appends
    schema errors from the optional openxml-audit package. Returns a list of
    error strings (empty = valid). When openxml-audit is not installed
    (neither importable module nor CLI on PATH) prints a single info line and
    contributes nothing — never a failure."""
    from save_hygiene import check_package_hygiene
    return check_package_hygiene(pptx_path) + _run_schema_audit(pptx_path)


def _run_schema_audit(pptx_path):
    """OOXML schema validation via the optional openxml-audit package."""
    try:
        import openxml_audit  # pip name: openxml-audit
    except ImportError:
        openxml_audit = None
    if openxml_audit is not None:
        errors = _integrity_via_module(openxml_audit, pptx_path)
        if errors is not None:
            return errors
        return []  # module present but unusable — warned above, skip
    cli = shutil.which("openxml-audit")
    if cli:
        return _integrity_via_cli(cli, pptx_path)
    print("  [INFO] pip install openxml-audit for OOXML schema validation "
          "(skipped)")
    return []


def main():
    if len(sys.argv) < 2:
        print("Usage: qa_check.py deck.pptx [--text] [--numbers] "
              "[--accessibility] [--integrity]")
        sys.exit(2)
    path = Path(sys.argv[1])
    if "--numbers" in sys.argv:
        dump_numbers(path)
        sys.exit(0)
    if "--text" in sys.argv:
        dump_text(path)
        sys.exit(0)
    accessibility = "--accessibility" in sys.argv
    issues = check_deck(path, accessibility=accessibility)
    if "--integrity" in sys.argv:
        from save_hygiene import HYGIENE_PREFIX
        issues["error"].extend(  # hygiene errors already carry their own tag
            e if e.startswith(HYGIENE_PREFIX) else f"integrity: {e}"
            for e in run_integrity(path))
    for e in issues["error"]:
        print(f"  [ERROR] {e}")
    for w in issues["warn"]:
        print(f"  [WARN]  {w}")
    print(f"\n{path.name}: {len(issues['error'])} error(s), "
          f"{len(issues['warn'])} warning(s)")
    sys.exit(1 if issues["error"] else 0)


if __name__ == "__main__":
    main()
