#!/usr/bin/env python3
"""Generate a suggested appendix outline markdown from a main deck outline."""
import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

from build_deck import parse_outline  # noqa: E402
from narrative import generate_appendix_outline  # noqa: E402


def main():
    parser = argparse.ArgumentParser(description="Generate appendix outline markdown.")
    parser.add_argument("outline", help="Main deck outline.md")
    parser.add_argument("--output", "-o", default=None,
                        help="Output path (default: outline-appendix.md)")
    args = parser.parse_args()

    outline_path = Path(args.outline)
    meta, slides = parse_outline(outline_path.read_text(encoding="utf-8"))
    out_path = Path(args.output) if args.output else outline_path.with_name(
        outline_path.stem + "-appendix.md")
    text = generate_appendix_outline(meta, slides)
    out_path.write_text(text, encoding="utf-8")
    print(f"Wrote {out_path}")


if __name__ == "__main__":
    main()
