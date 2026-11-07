"""Render Mermaid (.mmd) / Graphviz (.dot, .gv) sources to PNG for slides.

Routing rule (references/diagrams.md): native builders FIRST for standard
business exhibits; this renderer exists only for high-node-count or
auto-laid-out graphs where hand placement is impractical.  Output is a
raster image — NOT object-editable in PowerPoint.

Renderers are external, optional binaries:
  .dot/.gv -> graphviz `dot`   (brew install graphviz)
  .mmd     -> mermaid `mmdc`   (npm install -g @mermaid-js/mermaid-cli)

Usage:
  python3 scripts/render_diagram.py assets/diagrams/flow.mmd
  python3 scripts/render_diagram.py flow.dot --out assets/auto/diagrams/
"""
import argparse
import shutil
import subprocess
import sys
from pathlib import Path

DOT_DPI = 200        # print-quality raster on a 13.33in canvas
MMDC_SCALE = 3       # mermaid-cli pixel multiplier, ~equivalent sharpness
_STDERR_SNIP = 400   # cap renderer stderr passed through in errors

INSTALL_HINTS = {
    "dot": "graphviz is not installed — `brew install graphviz` "
           "(or apt-get install graphviz)",
    "mmdc": "mermaid-cli is not installed — "
            "`npm install -g @mermaid-js/mermaid-cli`",
}


class RendererMissing(RuntimeError):
    """The required external renderer binary is not on PATH."""


def _command_for(src, out):
    ext = src.suffix.lower()
    if ext in (".dot", ".gv"):
        exe = shutil.which("dot")
        if not exe:
            raise RendererMissing(INSTALL_HINTS["dot"])
        return [exe, "-Tpng", f"-Gdpi={DOT_DPI}", str(src), "-o", str(out)]
    if ext == ".mmd":
        exe = shutil.which("mmdc")
        if not exe:
            raise RendererMissing(INSTALL_HINTS["mmdc"])
        return [exe, "-i", str(src), "-o", str(out),
                "-s", str(MMDC_SCALE), "-b", "transparent"]
    raise ValueError(
        f"unsupported diagram source {src.suffix!r} — use .mmd (Mermaid) "
        "or .dot/.gv (Graphviz)")


def render_diagram(src, out_dir="assets/auto/diagrams"):
    """Render src to <out_dir>/<stem>.png; skip if the PNG is up to date."""
    src = Path(src)
    if not src.is_file():
        raise FileNotFoundError(f"diagram source not found: {src}")
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    out = out_dir / (src.stem + ".png")
    if out.is_file() and out.stat().st_mtime >= src.stat().st_mtime:
        return out
    cmd = _command_for(src, out)
    proc = subprocess.run(cmd, capture_output=True, text=True)
    if proc.returncode != 0 or not out.is_file():
        raise RuntimeError(
            f"{Path(cmd[0]).name} failed on {src.name}: "
            f"{proc.stderr.strip()[:_STDERR_SNIP]}")
    return out


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("source", help=".mmd / .dot / .gv file")
    ap.add_argument("--out", default="assets/auto/diagrams",
                    help="output directory (default: assets/auto/diagrams)")
    args = ap.parse_args(argv)
    try:
        png = render_diagram(args.source, args.out)
    except (RendererMissing, FileNotFoundError, ValueError,
            RuntimeError) as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 1
    print(png)
    return 0


if __name__ == "__main__":
    sys.exit(main())
