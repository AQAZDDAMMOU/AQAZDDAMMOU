#!/usr/bin/env python3
"""CI smoke test: validate, build, QA, and content-diff the example outline."""
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUTLINE = ROOT / "assets" / "example-outline.md"
OUTPUT = ROOT / "assets" / "smoke-deck.pptx"


def run(cmd):
    print("+", " ".join(cmd))
    r = subprocess.run(cmd, cwd=ROOT)
    if r.returncode != 0:
        sys.exit(r.returncode)


def run_advisory(cmd):
    """Run an optional step — report issues but never fail the smoke."""
    print("+", " ".join(cmd), "(advisory)")
    r = subprocess.run(cmd, cwd=ROOT)
    if r.returncode != 0:
        print(f"  [WARN] advisory step exited {r.returncode} (non-blocking)")


def main():
    py = sys.executable
    run([py, "scripts/build_deck.py", str(OUTLINE), "--check"])
    run([py, "scripts/build_deck.py", str(OUTLINE), "--output", str(OUTPUT)])
    run([py, "scripts/qa_check.py", str(OUTPUT)])
    run_advisory([py, "scripts/qa_check.py", str(OUTPUT), "--integrity"])
    run([py, "scripts/diff_deck.py", str(OUTLINE), str(OUTPUT)])
    run([py, "-m", "pytest", "tests/", "-q"])
    print("\nSmoke test passed.")


if __name__ == "__main__":
    main()
