#!/usr/bin/env python3
"""
visual_regress.py — Compare rendered slide PNGs against a blessed baseline.

Usage:
    python3 scripts/visual_regress.py baseline_dir/ current_dir/ [--update] [--threshold N]

Workflow: render the revised deck with render_slides.py, then diff it against
the thumbnails of the last-delivered deck before sending revisions. Slides are
matched by name/number (the slide-01.png pattern render_slides.py produces)
and compared by perceptual hash: pHash via the optional `imagehash` package
when installed, else a built-in Pillow-only 8x8 average hash. Hamming distance
above the threshold flags the slide and adds a coarse pixel-diff percentage.

Verdicts:
    changed slide   → failure (exit 1)
    slide missing from current (deleted) → failure (exit 1)
    new slide (no baseline)              → warning only
    --update        → bless current into baseline, exit 0
    missing/empty baseline without --update → exit 2 with init instructions

Requires: Pillow only (imagehash strictly optional).
"""
import argparse
import re
import shutil
import sys
from pathlib import Path

DEFAULT_THRESHOLD = 5
PIXEL_TOL = 24          # per-channel delta counted as "different"
PIXEL_DIFF_WIDTH = 256  # coarse-diff working width
SLIDE_RX = re.compile(r"slide-?(\d+)\.png$", re.I)

STATUS_LABEL = {
    "ok": "ok",
    "changed": "CHANGED",
    "new": "new (no baseline)",
    "missing": "MISSING (deleted vs baseline)",
}


def _ahash(path):
    """Pillow-only 8x8 average hash: grayscale, resize, threshold by mean.

    Note: uniform images (solid colour, blank slide) hash to 0 because every
    pixel equals the mean, so none exceeds it — this is a known average-hash
    limitation and not a bug.
    """
    from PIL import Image
    img = Image.open(path).convert("L").resize((8, 8), Image.LANCZOS)
    pixels = list(img.getdata())
    mean = sum(pixels) / len(pixels)
    bits = 0
    for i, value in enumerate(pixels):
        if value > mean:
            bits |= 1 << i
    return bits


def _hamming(a, b):
    """Bit-level hamming distance between two 64-bit integer hashes."""
    return bin(a ^ b).count("1")


def _hashers():
    """Return (hash_fn, distance_fn, label). Prefers imagehash pHash."""
    try:
        import imagehash
        from PIL import Image
        return (lambda p: imagehash.phash(Image.open(p)),
                lambda a, b: int(a - b), "phash")
    except ImportError:
        return _ahash, _hamming, "ahash"


def pixel_diff_pct(path_a, path_b):
    """Coarse percentage of pixels differing by >PIXEL_TOL in any channel.

    Both images are resized to a common ~256px-wide canvas first, so the
    number is a rough magnitude indicator, not a precise pixel count."""
    from PIL import Image
    a = Image.open(path_a).convert("RGB")
    b = Image.open(path_b).convert("RGB")
    w = PIXEL_DIFF_WIDTH
    h = max(1, round(a.height * w / max(a.width, 1)))
    a = a.resize((w, h))
    b = b.resize((w, h))
    differing = sum(
        1 for pa, pb in zip(a.getdata(), b.getdata())
        if any(abs(x - y) > PIXEL_TOL for x, y in zip(pa, pb)))
    return 100.0 * differing / (w * h)


def slide_key(path):
    """Match key for a slide PNG — slide number when parseable, so pdftoppm
    padding variants (slide-1.png vs slide-01.png) pair up; else filename."""
    m = SLIDE_RX.search(path.name)
    return f"slide-{int(m.group(1)):04d}" if m else path.name


def collect(directory):
    """Map slide_key → path for every slide PNG in a directory (grid.png and
    other non-slide files are ignored)."""
    directory = Path(directory)
    if not directory.is_dir():
        return {}
    return {slide_key(p): p for p in sorted(directory.glob("slide*.png"))}


def compare_maps(base, cur, threshold=DEFAULT_THRESHOLD):
    """Compare baseline/current slide maps. Returns (rows, hash_label) where
    each row is {name, distance, status, detail}."""
    hash_fn, dist_fn, label = _hashers()
    rows = []
    for key in sorted(set(base) | set(cur)):
        name = (cur.get(key) or base.get(key)).name
        if key not in cur:
            rows.append({"name": name, "distance": None,
                         "status": "missing", "detail": ""})
        elif key not in base:
            rows.append({"name": name, "distance": None,
                         "status": "new", "detail": ""})
        else:
            distance = int(dist_fn(hash_fn(base[key]), hash_fn(cur[key])))
            if distance > threshold:
                pct = pixel_diff_pct(base[key], cur[key])
                rows.append({"name": name, "distance": distance,
                             "status": "changed",
                             "detail": f"{pct:.1f}% pixels differ"})
            else:
                rows.append({"name": name, "distance": distance,
                             "status": "ok", "detail": ""})
    return rows, label


def bless(current_dir, baseline_dir):
    """Copy current slide PNGs into baseline (creating it), removing stale
    baseline slides that no longer exist in current. Returns count copied."""
    baseline_dir = Path(baseline_dir)
    baseline_dir.mkdir(parents=True, exist_ok=True)
    cur = collect(current_dir)
    for key, stale in collect(baseline_dir).items():
        if key not in cur or stale.name != cur[key].name:
            stale.unlink()
    for path in cur.values():
        shutil.copy2(path, baseline_dir / path.name)
    return len(cur)


def _print_table(rows, label, threshold):
    width = max([len(r["name"]) for r in rows] + [len("slide")])
    print(f"{'slide'.ljust(width)}  {'dist':>4}  status")
    for r in rows:
        dist = "-" if r["distance"] is None else str(r["distance"])
        line = f"{r['name'].ljust(width)}  {dist:>4}  {STATUS_LABEL[r['status']]}"
        if r["detail"]:
            line += f" ({r['detail']})"
        print(line)
    counts = {s: sum(1 for r in rows if r["status"] == s)
              for s in ("ok", "changed", "new", "missing")}
    print(f"\n{len(rows)} slide(s) compared ({label}, threshold {threshold}): "
          f"{counts['ok']} ok, {counts['changed']} changed, "
          f"{counts['new']} new, {counts['missing']} missing")
    return counts


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="Visual regression between two slide-thumbnail directories.")
    parser.add_argument("baseline", help="Baseline directory (last-delivered render)")
    parser.add_argument("current", help="Current directory (revised render)")
    parser.add_argument("--update", action="store_true",
                        help="Bless current into baseline (initializes a missing baseline)")
    parser.add_argument("--threshold", type=int, default=DEFAULT_THRESHOLD,
                        help=f"Max hamming distance before a slide is flagged "
                             f"(default {DEFAULT_THRESHOLD})")
    args = parser.parse_args(argv)

    baseline, current = Path(args.baseline), Path(args.current)
    cur = collect(current)
    if not cur:
        print(f"[ERROR] no slide PNGs in {current}/ — render the deck first: "
              f"python3 scripts/render_slides.py deck.pptx --out {current}/",
              file=sys.stderr)
        return 2

    base = collect(baseline)
    if not base:
        if args.update:
            n = bless(current, baseline)
            print(f"Initialized baseline: {n} slide(s) -> {baseline}/")
            return 0
        print(f"[ERROR] baseline {baseline}/ is missing or empty.\n"
              f"Initialize it from a delivered render with:\n"
              f"  python3 scripts/visual_regress.py {baseline} {current} --update",
              file=sys.stderr)
        return 2

    rows, label = compare_maps(base, cur, args.threshold)
    counts = _print_table(rows, label, args.threshold)

    if args.update:
        n = bless(current, baseline)
        print(f"Baseline updated: {n} slide(s) blessed into {baseline}/")
        return 0
    return 1 if counts["changed"] or counts["missing"] else 0


if __name__ == "__main__":
    sys.exit(main())
