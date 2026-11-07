# Diagrams: Native First, Rendered Second

## Routing rule (mandatory order)

1. **Native builders first** for any standard business exhibit. These stay
   fully editable in PowerPoint: `process-flow`, `swimlane-process`,
   `value-chain`, `org-chart`, `driver-tree`, `architecture-layers`,
   `gantt`, `matrix-2x2`, `stakeholder-map`, `raci`, `funnel`, plus all
   chart layouts. If the content fits one of these, DO NOT render an image.
2. **Native charts** (`**Visual:** chart:bar` etc.) for any quantitative
   exhibit. Never screenshot a chart.
3. **Rendered diagrams** (`**Visual:** diagram:file.mmd|.dot|.gv`) ONLY when
   automatic graph layout genuinely outweighs editability: dependency
   graphs, systems architecture with many cross-links, state machines —
   typically >10 nodes or non-planar edge sets. Output is a raster PNG the
   client cannot edit object-by-object; say so when delivering.
4. **Raster images** only for photos, screenshots, and source-research
   figures (with citation). Never use a raster where an editable standard
   business exhibit is feasible.

## Rendered diagram workflow

1. Write the source file under `assets/diagrams/` (e.g.
   `assets/diagrams/dependency-graph.mmd`). Keep node labels short; the
   PNG is scaled to fit the visual box.
2. Reference it: `**Visual:** diagram:dependency-graph.mmd` on a
   `two-column-split`, `full-image`, or `chart-callout` slide.
3. `build_deck.py` renders it automatically (cached by mtime into
   `assets/auto/diagrams/`). Pre-render manually to check output:
   `python3 scripts/render_diagram.py assets/diagrams/dependency-graph.mmd`

## Renderer install

| Source | Renderer | Install |
|--------|----------|---------|
| `.dot`, `.gv` (Graphviz) | `dot` | `brew install graphviz` |
| `.mmd` (Mermaid) | `mmdc` | `npm install -g @mermaid-js/mermaid-cli` |

`--check` warns (does not error) when a renderer is missing; the build
itself fails fast. Prefer Graphviz for plain node-edge graphs (lighter
dependency); Mermaid for sequence/state diagrams.

## Style guidance

- Graphviz: set `rankdir=LR`, `fontname="Helvetica"`, and palette-adjacent
  hex fills so the render doesn't clash with the deck palette.
- Mermaid: default theme on transparent background (`-b transparent` is
  applied automatically) reads acceptably on light palettes; for dark
  palettes prefer Graphviz with explicit fills.
- QA renders (`render_slides.py`) and `geometry_report.py` treat the
  result as a normal picture — no special handling needed.
