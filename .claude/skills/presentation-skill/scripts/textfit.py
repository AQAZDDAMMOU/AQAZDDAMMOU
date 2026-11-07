"""Build-time text measurement — predict wrapping before the deck renders.

estimate_lines() greedily word-wraps the way PowerPoint will, using PIL glyph
metrics for the *actual* palette font when that face can be found on disk
(resolve_font_file: matplotlib -> fc-match -> well-known font directories),
falling back to Arial/DejaVu and then to an average-character-width
approximation (within ~10% on typical copy).  set_measure_font() lets the
build set the deck's body family once so every caller measures the same way;
measure_font_status() reports whether that family actually resolved.
bullets_fit() scores a bullet block against a layout's vertical budget so
build_deck.validate() and the builders share one capacity model; mismatched
copies of that arithmetic were the #1 source of render-time overflow.
apply_autofit() writes the <a:normAutofit> shrink hint for the 100-140%
overflow band; fit_scale() computes the percentage it should carry, and
shrink_to_fit_pt() returns a literal point size for table cells (PowerPoint
ignores autofit inside table cells).
"""
import logging
import math
import os
import re
import shutil
import subprocess
from pathlib import Path

LINE_PITCH = 1.2117        # PowerPoint single line spacing (1.2 * 1.00975 gap)
LINE_SPACING = LINE_PITCH  # backwards-compatible alias
WIDTH_FACTOR = 0.50        # fallback: avg glyph width as a fraction of font size
WIDTH_FACTOR_BOLD = 0.53
MARKER_INDENT = 0.30       # accent-square marker offset (builders._draw_marker)
ICON_INDENT = 0.66         # icon-in-circle marker offset

FIT_MARGIN = 0.035         # safety headroom subtracted before snapping fontScale
FIT_GRID = 0.025           # fontScale grid: PowerPoint's own 2.5% steps
FIT_GRID_PCT = 2.5         # the same grid expressed in percent
FIT_FLOOR_PCT = 25.0       # never emit a fontScale below this
GRID_EPS = 1e-9            # absorbs binary float error when snapping to the grid
SHRINK_STEP_PT = 0.5       # literal-shrink search step
SHRINK_MIN_PT = 9.0        # literal-shrink floor

FC_MATCH_TIMEOUT_S = 5
_FONT_EXTS = (".ttf", ".otf", ".ttc")
_FONT_DIRS = (
    "~/Library/Fonts",
    "/Library/Fonts",
    "/System/Library/Fonts",
    "/System/Library/Fonts/Supplemental",
    "/usr/share/fonts",
    "C:/Windows/Fonts",
)

# keep in sync with builders.ICON_RX and helpers._RICH_RX
_ICON_RX = re.compile(r"^icon:([a-z0-9-]+)\s+(.*)$")
_RICH_RX = re.compile(r"\*\*(.+?)\*\*|\{accent\}(.+?)\{/\}")

_FONT_CANDIDATES = {
    False: (
        "/System/Library/Fonts/Supplemental/Arial.ttf",
        "/Library/Fonts/Arial.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
        "/System/Library/Fonts/Helvetica.ttc",
    ),
    True: (
        "/System/Library/Fonts/Supplemental/Arial Bold.ttf",
        "/Library/Fonts/Arial Bold.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
        "/System/Library/Fonts/Helvetica.ttc",  # regular face — close enough
    ),
}
_FONT_CACHE = {}
_FILE_CACHE = {}
_MEASURE_FAMILY = None


def set_measure_font(family):
    """Set the module-default family used when a caller passes no family=.

    Process-global (like palettes.set_cjk); tests/conftest.py resets it.
    """
    global _MEASURE_FAMILY
    _MEASURE_FAMILY = (family or "").strip() or None
    return _MEASURE_FAMILY


def get_measure_font():
    """The module-default measurement family, or None."""
    return _MEASURE_FAMILY


def measure_font_status():
    """(requested_family, resolved_path_or_None) for the module default.

    QA tooling reports "measured with fallback font" when the family was
    requested but the face could not be found on this machine.
    """
    if not _MEASURE_FAMILY:
        return (None, None)
    return (_MEASURE_FAMILY, resolve_font_file(_MEASURE_FAMILY, False))


def _font_names(path):
    """(family, style) reported by the font file itself; ("", "") if unreadable."""
    try:
        from PIL import ImageFont
        family, style = ImageFont.truetype(path, 12).getname()
    except (ImportError, OSError, ValueError, IndexError, TypeError):
        return ("", "")
    return (family or "", style or "")


def _family_matches(path, family, bold=False):
    """True when path really is `family` (not a fontconfig fallback).

    With bold=True the face must also be a bold cut: matching only on family
    would happily hand back the regular weight and measure it as bold.
    """
    name, style = _font_names(path)
    if name:
        if name.casefold() != family.casefold():
            return False
        return not bold or "bold" in style.casefold()
    # Pillow unavailable or face unreadable: trust the file name instead.
    stem = Path(path).stem.casefold().replace("-", " ")
    if not stem.startswith(family.casefold()):
        return False
    return not bold or ("bold" in stem or stem.endswith("bd"))


def _resolve_matplotlib(family, bold):
    try:
        from matplotlib.font_manager import FontProperties, findfont
    except ImportError:
        return None
    # findfont logs "Font family ... not found" on every miss; a miss is an
    # expected outcome here, so keep it out of the build's stderr.
    log = logging.getLogger("matplotlib.font_manager")
    prev = log.level
    log.setLevel(logging.ERROR)
    try:
        path = findfont(FontProperties(
            family=family, weight="bold" if bold else "normal"),
            fallback_to_default=False)
    except (ValueError, RuntimeError, OSError):
        return None
    finally:
        log.setLevel(prev)
    if path and os.path.exists(path) and _family_matches(path, family, bold):
        return path
    return None


def _resolve_fc_match(family, bold):
    exe = shutil.which("fc-match")
    if not exe:
        return None
    spec = f"{family}:bold" if bold else family
    try:
        proc = subprocess.run([exe, "-f", "%{file}", spec],
                              capture_output=True, text=True,
                              timeout=FC_MATCH_TIMEOUT_S, check=False)
    except (OSError, subprocess.SubprocessError):
        return None
    path = (proc.stdout or "").strip()
    # fc-match always answers, substituting a fallback face when the family is
    # missing — the family check below is what makes the answer trustworthy.
    if path and os.path.exists(path) and _family_matches(path, family, bold):
        return path
    return None


def _resolve_scan(family, bold):
    # spaces are optional so the Windows "ArialBlack.ttf" spelling also hits;
    # bold suffixes cover both the "Arial Bold" and the "arialbd" conventions
    fam_rx = r"\s?".join(re.escape(part) for part in family.split())
    rx = re.compile(r"^" + fam_rx
                    + (r"(?:[ -]?bold|bd)" if bold else "")
                    + r"\.(ttf|otf|ttc)$", re.I)
    for raw in _FONT_DIRS:
        base = Path(raw).expanduser()
        try:
            if not base.is_dir():
                continue
            entries = base.rglob("*")
        except OSError:
            continue
        try:
            for path in entries:
                if path.suffix.lower() in _FONT_EXTS and rx.match(path.name) \
                        and _family_matches(str(path), family, bold):
                    return str(path)
        except OSError:
            continue
    return None


def resolve_font_file(family, bold=False):
    """Path to the TTF/OTF/TTC for `family` (bold face when bold), else None.

    Search order: matplotlib's font_manager, `fc-match`, then a scan of the
    platform font directories.  A hit only counts when the file's own family
    name matches the request, so fontconfig's silent substitution does not
    make us measure the wrong typeface.  Cached per (family, bold).
    """
    if not family or not str(family).strip():
        return None
    family = str(family).strip()
    key = (family.casefold(), bool(bold))
    if key not in _FILE_CACHE:
        _FILE_CACHE[key] = (_resolve_matplotlib(family, bold)
                            or _resolve_fc_match(family, bold)
                            or _resolve_scan(family, bold))
    return _FILE_CACHE[key]


def _load_font(font_pt, bold, family=None):
    """PIL font for (size, bold, family), cached.

    Tries the requested family (or the module default) first, then the
    Arial/DejaVu candidates; None when PIL or every face is missing.
    """
    family = family or _MEASURE_FAMILY
    key = (font_pt, bool(bold), (family or "").casefold())
    if key not in _FONT_CACHE:
        font = None
        try:
            from PIL import ImageFont
            paths = []
            if family:
                found = resolve_font_file(family, bold)
                if found:
                    paths.append(found)
            paths.extend(_FONT_CANDIDATES[bool(bold)])
            size = max(float(font_pt), 1.0)
            for path in paths:
                try:
                    # Pillow >= 10.1 accepts a float size; without it every
                    # half-point step in `shrink_to_fit_pt` would measure at
                    # the same integer size (12.5 -> 12) and never converge.
                    font = ImageFont.truetype(path, size)
                except TypeError:
                    try:
                        font = ImageFont.truetype(path, max(int(round(size)), 1))
                    except OSError:
                        continue
                except OSError:
                    continue
                break
        except ImportError:
            pass
        _FONT_CACHE[key] = font
    return _FONT_CACHE[key]


def _measurer(font_pt, bold, family=None):
    """Width-in-inches function for (size, bold, family)."""
    font = _load_font(font_pt, bold, family)
    if font is not None:
        return lambda s: font.getlength(s) / 72.0
    per_char = font_pt * (WIDTH_FACTOR_BOLD if bold else WIDTH_FACTOR) / 72.0
    return lambda s: len(s) * per_char


def strip_markup(text):
    """Measurement form of a bullet: drop icon: prefix and rich-text markers.
    Returns (plain_text, has_icon)."""
    text = (text or "").strip()
    m = _ICON_RX.match(text)
    if m:
        text = m.group(2)
    text = _RICH_RX.sub(lambda mm: mm.group(1) or mm.group(2), text)
    return text, bool(m)


def estimate_lines(text, font_pt, box_w_in, bold=False, family=None):
    """Estimated wrapped line count of text in a box_w_in-wide frame (greedy
    word wrap; tokens wider than the box count ceil(width/box) lines)."""
    words = (text or "").split()
    if not words or box_w_in <= 0:
        return 0
    measure = _measurer(font_pt, bold, family)
    space_w = measure(" ")
    lines, cur = 0, 0.0
    for word in words:
        w = measure(word)
        if w > box_w_in:                       # unbreakable long token
            if cur:
                lines += 1
            lines += math.ceil(w / box_w_in)
            cur = 0.0
            continue
        need = w if not cur else cur + space_w + w
        if need <= box_w_in:
            cur = need
        else:
            lines += 1
            cur = w
    if cur:
        lines += 1
    return lines


def wrap_lines(text, font_pt, box_w_in, bold=False, family=None):
    """The wrapped lines themselves (same greedy wrap as estimate_lines).
    Overlong unbreakable tokens land on their own line unsplit."""
    words = (text or "").split()
    if not words or box_w_in <= 0:
        return []
    measure = _measurer(font_pt, bold, family)
    space_w = measure(" ")
    lines, cur, cur_w = [], [], 0.0
    for word in words:
        w = measure(word)
        need = w if not cur else cur_w + space_w + w
        if cur and (w > box_w_in or need > box_w_in):
            lines.append(" ".join(cur))
            cur, cur_w = [word], w
        else:
            cur.append(word)
            cur_w = need
    if cur:
        lines.append(" ".join(cur))
    return lines


def estimate_height_in(text, font_pt, box_w_in, bold=False,
                       line_spacing=LINE_PITCH, family=None):
    """Estimated rendered height (inches) of wrapped text."""
    lines = estimate_lines(text, font_pt, box_w_in, bold, family)
    return lines * font_pt * line_spacing / 72.0


def snap_to_grid(pct, floor_pct=FIT_FLOOR_PCT, grid=FIT_GRID):
    """Snap a raw fontScale percent DOWN to PowerPoint's normAutofit grid
    (2.5% by default), then floor at floor_pct and cap at 100.

    floor_pct defaults to FIT_FLOOR_PCT (25) — 0.0 is not a legal fontScale,
    so callers that want no floor still need to say so explicitly.  Shared
    floor/round logic: fit_scale calls this with floor_pct=FIT_FLOOR_PCT
    (25); builders._autofit_overflow calls it with floor_pct=
    BULLET_FLOOR_PCT (80) for the documented bullet-shrink floor.  The
    return stays a float because half of the 2.5% grid's values are
    half-percents (87.5, 82.5, ...); apply_autofit writes them losslessly.
    """
    grid_pct = round(grid * 100.0, 9)   # 0.025 * 100 is not exactly 2.5
    if grid_pct > 0:
        steps = math.floor(pct / grid_pct + GRID_EPS)
        snapped = round(steps * grid_pct, 6)
    else:
        snapped = pct
    return max(floor_pct, min(100.0, snapped))


def fit_scale(text, font_pt, box_w_in, box_h_in, bold=False, family=None,
              margin=FIT_MARGIN, grid=FIT_GRID):
    """The <a:normAutofit> fontScale percent this copy needs, as a float.

    100.0 when the text already fits.  Otherwise the height ratio less a safety
    margin, snapped DOWN to PowerPoint's 2.5% grid and floored at 25 via
    snap_to_grid.
    """
    if box_h_in <= 0 or box_w_in <= 0 or not (text or "").strip():
        return 100.0
    est_h = estimate_height_in(text, font_pt, box_w_in, bold, family=family)
    if est_h <= 0 or est_h <= box_h_in:
        return 100.0
    raw_pct = (box_h_in / est_h - margin) * 100.0
    return snap_to_grid(raw_pct, floor_pct=FIT_FLOOR_PCT, grid=grid)


def shrink_to_fit_pt(text, font_pt, box_w_in, box_h_in, bold=False,
                     family=None, min_pt=SHRINK_MIN_PT):
    """Largest size <= font_pt (0.5pt steps) whose wrapped height fits box_h_in.

    Table cells ignore <a:normAutofit>, so their copy has to be shrunk with a
    literal run size.  Returns min_pt when even that overflows; never grows.
    """
    if not (text or "").strip() or box_w_in <= 0 or box_h_in <= 0:
        return float(font_pt)
    pt = float(font_pt)
    min_pt = float(min_pt)
    while pt > min_pt:
        if estimate_height_in(text, pt, box_w_in, bold,
                              family=family) <= box_h_in:
            return pt
        pt = round(pt - SHRINK_STEP_PT, 2)
    return min(float(font_pt), min_pt)


def bullets_fit(bullets, font_pt, col_w_in, avail_h_in, cols=1, step_in=0.0,
                bold=False, family=None):
    """Score a bullet block against its vertical budget.

    Splits bullets across cols columns column-major (like build_bullet_slide);
    each bullet consumes at least its fixed step_in slot. Returns
    (ratio, total_lines, overflowing_indices) where:
      ratio               = worst column's estimated height / avail_h_in —
                            >1.0 means autofit territory, >1.4 means split slide.
      total_lines         = sum of wrapped lines across all bullets.
      overflowing_indices = list of bullet indices (0-based) whose wrapped height
                            exceeds step_in AND which are not the last bullet in
                            their column (i.e. they would collide with the next
                            bullet's fixed-y slot).  Empty when step_in == 0.
    """
    bullets = list(bullets or [])
    if not bullets or avail_h_in <= 0:
        return 0.0, 0, []
    cols = max(int(cols), 1)
    per_col = -(-len(bullets) // cols)
    line_h = font_pt * LINE_PITCH / 72.0
    heights, line_counts, total_lines = [], [], 0
    for b in bullets:
        plain, has_icon = strip_markup(b)
        w = col_w_in - (ICON_INDENT if has_icon else MARKER_INDENT)
        lines = max(estimate_lines(plain, font_pt, w, bold, family), 1)
        total_lines += lines
        heights.append(max(lines * line_h, step_in))
        line_counts.append(lines)
    worst = max(sum(heights[c * per_col:(c + 1) * per_col])
                for c in range(cols))
    # Per-slot collision detection: a non-terminal bullet overflows its slot
    # when its wrapped height exceeds the fixed step_in spacing.
    overflowing_indices = []
    if step_in > 0:
        for c in range(cols):
            col_start = c * per_col
            col_end = min(col_start + per_col, len(bullets))
            for pos in range(col_start, col_end - 1):  # skip last in each col
                if line_counts[pos] * line_h > step_in:
                    overflowing_indices.append(pos)
    return worst / avail_h_in, total_lines, overflowing_indices


def apply_autofit(text_frame, scale_pct, ln_reduction_pct=10):
    """Write <a:normAutofit fontScale=".." lnSpcReduction=".."/> into the
    frame's bodyPr (values are 1000ths of a percent: 85% -> "85000").

    scale_pct may be fractional — fit_scale's 2.5% grid produces half-percents
    (87.5 -> "87500"); integer inputs are unaffected."""
    from lxml import etree
    from pptx.oxml.ns import qn
    bodyPr = text_frame._txBody.find(qn("a:bodyPr"))
    for tag in ("a:noAutofit", "a:normAutofit", "a:spAutoFit"):
        for el in bodyPr.findall(qn(tag)):
            bodyPr.remove(el)
    fit = etree.SubElement(bodyPr, qn("a:normAutofit"))
    fit.set("fontScale", str(int(round(float(scale_pct) * 1000))))
    fit.set("lnSpcReduction", str(int(ln_reduction_pct) * 1000))
    return fit
