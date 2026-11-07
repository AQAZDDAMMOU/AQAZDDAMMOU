#!/usr/bin/env python3
"""Build a Claude Desktop–ready skill zip (≤200 files, no .git/.claude).

Usage:
    python3 scripts/package_skill.py
    python3 scripts/package_skill.py --output ~/Desktop/presentation-skill.zip

Creates:
    presentation-skill.zip
      └── presentation-skill/
            ├── SKILL.md
            ├── scripts/
            ├── references/
            └── ...

Claude Desktop rejects zips with >200 files. Do NOT zip the repo root from
Finder if .git or .claude are present — use this script instead.
"""
import argparse
import fnmatch
import os
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SKILL_NAME = "presentation-skill"
MAX_FILES = 200

# Paths relative to repo root to skip entirely.
EXCLUDE_DIRS = {
    ".git",
    ".claude",
    ".superpowers",
    ".pytest_cache",
    ".ruff_cache",
    ".mypy_cache",
    ".venv",
    "venv",
    "__pycache__",
    "assets/icons",
    "assets/auto",
    "assets/qa-thumbs",
    "dist",
    "docs/superpowers",
}

# Top-level dot-prefixed directories are excluded by default (SDD/agent
# scratch like .superpowers, .claude, .git); list any that should still
# ship in the skill zip here. Empty today.
ALLOWED_DOT_DIRS = set()

EXCLUDE_FILES = {
    ".DS_Store",
}

EXCLUDE_GLOBS = (
    "*.pptx",
    "*.pyc",
    "*.pyo",
)


def _excluded(rel: Path) -> bool:
    parts = rel.parts
    if parts and parts[0].startswith(".") and parts[0] not in ALLOWED_DOT_DIRS:
        return True
    for i in range(len(parts)):
        prefix = Path(*parts[: i + 1])
        if str(prefix) in EXCLUDE_DIRS or parts[i] in EXCLUDE_DIRS:
            return True
    if rel.name in EXCLUDE_FILES:
        return True
    name = rel.name
    for pat in EXCLUDE_GLOBS:
        if fnmatch.fnmatch(name, pat):
            return True
    return False


def collect_files(root: Path):
    files = []
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = sorted(
            d for d in dirnames
            if not _excluded((Path(dirpath).relative_to(root) / d))
        )
        for fname in sorted(filenames):
            rel = Path(dirpath).relative_to(root) / fname
            if _excluded(rel):
                continue
            files.append(rel)
    return sorted(files)


def build_zip(output: Path, include_tests: bool = True) -> int:
    files = collect_files(ROOT)
    if not include_tests:
        files = [f for f in files if f.parts[0] != "tests"]
    if not any(f.name == "SKILL.md" for f in files):
        print("ERROR: SKILL.md missing from package", file=sys.stderr)
        sys.exit(1)
    if len(files) > MAX_FILES:
        print(f"ERROR: {len(files)} files exceeds Claude Desktop limit "
              f"({MAX_FILES}). Tighten excludes or pass --no-tests.",
              file=sys.stderr)
        sys.exit(1)

    output.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for rel in files:
            arc = f"{SKILL_NAME}/{rel.as_posix()}"
            zf.write(ROOT / rel, arc)
    return len(files)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--output", "-o",
        default=str(ROOT / "dist" / f"{SKILL_NAME}.zip"),
        help="Output zip path (default: dist/presentation-skill.zip)",
    )
    parser.add_argument(
        "--no-tests", action="store_true",
        help="Omit tests/ to reduce file count",
    )
    args = parser.parse_args()
    out = Path(args.output).expanduser().resolve()
    n = build_zip(out, include_tests=not args.no_tests)
    print(f"Wrote {out} ({n} files, limit {MAX_FILES})")
    print("Upload this zip in Claude Desktop → Settings → Capabilities → Skills")


if __name__ == "__main__":
    main()
