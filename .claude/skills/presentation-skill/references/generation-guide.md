# Generation Guide

## Core Build Stack

| Tool | Purpose |
|------|---------|
| `python-pptx` | Native .pptx creation — real shapes, placeholders, charts |
| `Pillow` | Image normalization and aspect-ratio math |
| `scripts/build_deck.py` | Builds the full deck from a markdown outline (also `--check` validation) |
| `scripts/qa_check.py` | Programmatic QA of the built .pptx |
| `scripts/render_slides.py` | Converts .pptx → thumbnail images for visual QA |

`python-pptx` produces native PowerPoint XML with proper placeholders and fully editable elements. Always use the scripts rather than hand-writing pptx code — the layout builders live in `scripts/builders.py`, palettes in `scripts/palettes.py`, chart helpers in `scripts/charts.py`.

---

## Outline Format (exact syntax)

One `## Slide N:` section per slide. The text after `Slide N:` becomes the heading unless overridden with `- Heading:` or `- Title:` (recommended on `title` / `closing` layouts; if omitted, the `## Slide N:` text is used as the title). Always validate with `build_deck.py --check` before building.

**Deck front-matter** (optional, before the first `## Slide`):

```markdown
**Palette:** midnight-executive        ← default for the whole deck
**Purpose:** pitch                     ← pitch|strategy|client|talk|tutorial (narrative hints)
**Takeaway:** "One sentence investors should remember"
**Variant:** b                         ← preset: a=default, b=aurora+comfortable, c=swiss-light+compact
**Footer:** "Confidential · Q3 2026"   ← bottom-left on content slides
**Page-Numbers:** on                   ← bottom-right on content slides
**Size:** 16:9                         ← or 4:3
**Density:** compact                   ← compact (default) or comfortable
**Author:** "Alex Kim"                 ← dc:creator / lastModifiedBy in the saved file
**Number-Format:** units=M decimals=1 negatives=parens   ← deck-wide number policy
**Style-File:** style.json             ← house-style word lists and title rules
```

Omit `**Layout:**` on a slide and the builder auto-selects from content (`smart_layout.py`). Explicit layout always wins. Consulting layouts are auto-selected from their signature rows: `- Layer:` →
architecture-layers, `- Lane:` → swimlane-process, `- Item:` → matrix-2x2,
`- Tile N:` → dashboard, `- Milestone:`/`- Bar:` with Start/End → gantt,
`- Bar:` with Size/Value (no Start/End) → bar-mekko, `- Node:` →
driver-tree (org-chart must be explicit), `- Quote:` → quote-evidence, a
`total` data row → waterfall, `- Step N:` with no visual → process-flow,
`- Value:` with no data → big-number. Explicit `**Layout:**` always wins. See `references/narratives/*.md` for purpose-specific slide arcs.

**Inline rich text** works in bullets, points, captions, and card bodies:
`**bold**` for bold, `{accent}text{/}` for accent-colored emphasis.

**Icons**: prefix any bullet with `icon:<name> ` (kebab-case [Tabler](https://tabler.io/icons)/[Lucide](https://lucide.dev) names) to render an icon-in-circle marker. Cards take `Icon="name"`. Icons are fetched once and cached in `assets/icons/`.

```markdown
## Slide 1: Title
**Layout:** title
**Visual:** user-image:hero.png        ← optional hero image, right side
**Palette:** swiss-light               ← optional per-slide override
- Title: "Accelerating AI Infrastructure"
- Subtitle: "Q4 Strategy Review · December 2026"
- Notes: "Open with the cost-savings story."   ← speaker notes (write these for EVERY slide)

## Slide 2: Executive Summary
**Layout:** exec-summary
- Headline: "Three strategic priorities for H1 2027"
- Point 1: "icon:server Deploy GPU cluster Phase 2 by March (saving **$2.1M/yr**)"
- Point 2: "icon:cloud Migrate {accent}80% of workloads{/} to hybrid cloud by April"
- Point 3: "icon:rocket Launch developer platform beta with 500 design partners"

## Slide 3: Market Opportunity
**Layout:** two-column-split
**Visual:** chart:bar                  ← bar|hbar|line|pie|doughnut|area|scatter|bubble|combo|stacked-100 (scatter data contract: see "Data blocks" below)
**Series:** Market Size ($B)
**Data:**
- 2024: 42
- 2025: 67
- 2026: 104
- 2027: 158
- Heading: "AI infrastructure market growing 55.5% a year"
- Enterprise AI spend reached $104B globally in 2026
- Infrastructure layer captures 41% of total AI spend

## Slide 3b: Revenue vs Costs (multi-series)
**Layout:** two-column-split
**Visual:** chart:line
**Series:** Revenue, Costs
**Data:**
- Q1: 4.2, 3.1
- Q2: 5.1, 3.0
- Q3: 6.3, 2.8
- Heading: "Margin expansion through cloud migration"

## Slide 4: Product Screenshot
**Layout:** full-image
**Visual:** user-image:product-screenshot.png
- Caption: "Live dashboard — 47ms p99 latency"

## Slide 5: Team
**Layout:** cards-3
- Card 1: Name="Alex Kim" Title="CEO & Co-founder" Bio="Ex-Google Brain, 2x founder"
- Card 2: Name="Jordan Lee" Title="CTO" Bio="Built infra at Scale AI"
- Card 3: Name="Sam Patel" Title="Head of Product" Bio="Former PM at Stripe"

## Slide 6: Key Metrics
**Layout:** stat-callout
- Value="$14.2M" Label="Cost Savings" Sublabel="vs. legacy stack"
- Value="99.97%" Label="Uptime SLA" Sublabel="Jan–Jun 2026"

## Slide 7: Roadmap
**Layout:** timeline
- Date="Q3 2026" Title="Cloud migration" Desc="Complete hybrid migration"
- Date="Q4 2026" Title="Phase 2 GPUs" Desc="256 H100s operational"
- Date="Q1 2027" Title="Platform GA" Desc="Billing + RBAC"

## Slide 8: Before / After
**Layout:** comparison
**Visual-Left:** user-image:before.png      ← image OR bullets per side
**Visual-Right:** user-image:after.png
- Left label: "Legacy Stack (2023)"
- Right label: "New Architecture (2026)"
- Left: "Monolith on-prem"                  ← used when no Visual-Left image
- Right: "Active-active multi-region"

## Slide 9: Pricing
**Layout:** table
- Heading: "Plan comparison"
| Plan | Price | Seats |
|------|-------|-------|
| Starter | $99 | 5 |
| Scale | $499 | 50 |
```

Generic bullets (`- text`) work on `bullet-list`, `two-column-split`, `exec-summary` (or use `- Point N:`). Timeline also accepts plain `- Q3 2026: description` bullets.

---

## Layout Types (all implemented in `scripts/builders.py`)

Full per-layout gallery with use/don't-use guidance: `references/gallery.md`.

| Layout Name | When to Use | Content keys |
|-------------|-------------|--------------|
| `title` | Opening slide | Title, Subtitle, optional Visual |
| `section-divider` | Chapter breaks | heading (from `## Slide N:`), Subtitle |
| `exec-summary` | Key takeaways | Headline + up to 4 Point N cards (exactly 3 ⇒ one full-width column) |
| `bullet-list` | Main points | Heading + up to 14 bullets (2 columns; 7 in comfortable) |
| `two-column-split` | Text + visual | Bullets left; chart/image right |
| `cards-3` / `cards-4` | Team, features | Card N with Name/Title/Bio (or Title/Body) |
| `stat-callout` | Big numbers | 2–3 Value/Label/Sublabel stats |
| `timeline` | Chronological steps | Date/Title/Desc items (max 5) |
| `comparison` | Before/after, A vs B | Left/Right labels + images or bullets |
| `table` | Structured data | Markdown table rows; optional Align/Units/Zebra/Total-Row/Format/Summary/Sort (see Cycle 7 directives) |
| `comps-table` | Banking comps / precedents | `table` preset: `- Align: decimal`, `- Summary: mean, median`, `- Total-Row: off` |
| `full-image` | Hero visual | Visual (cover-cropped) + Caption |
| `closing` | Thank you / CTA | Title, Subtitle, Contact |
| `exec-summary-scqa` | Consulting exec summary | Situation, Finding N, Recommendation |
| `agenda` | Section navigation | Section names as bullets, optional Current |
| `waterfall` | Value bridges, cost walks | **Data:** start, ± deltas, `total` row |
| `matrix-2x2` | Prioritization quadrants | X-axis/Y-axis, Q1–Q4, Item rows |
| `harvey-scorecard` | Vendor/capability comparison | Markdown table, cells 0–4 |
| `process-flow` | Operating model steps | Step N: Title/Desc (chevrons, max 6) |
| `big-number` | One hero metric | Value, Label, Context |
| `chart-callout` | Chart + key insight box | chart Data + Callout |
| `dashboard` | KPI tiles with trends | Tile N: Value/Label/Delta/Trend |
| `quote-evidence` | Voice + proof | Quote, Attribution, one Value/Label stat |
| `funnel` | Pipeline drop-off | **Data:** stage rows (auto conversion %) |
| `next-steps` | Conclusions for Q&A | Step N: Action/Owner/When |
| `mekko` | Market maps, share-of-wallet | **Series:** segments + multi-value **Data:** |
| `gantt` | Phased roadmaps, swimlanes | **Periods:** (plain list or `quarters FYa-FYb` / `months Mon YYYY - Mon YYYY`) + Bar/Milestone rows; optional Span/Shade/Line rows, `Responsible=`/`Remark=` columns, `Row="Parent > Child"` |
| `architecture-layers` | Target operating model, tech stack, layered frameworks | `- Layer: Name=".." Items="a, b" [Accent="2"]` rows (max 6 layers, 5 items each) |
| `swimlane-process` | Cross-functional processes, intake/approval flows | `- Lane: Name=".." Steps="a > b > c"` rows (max 4 lanes, 5 steps) |
| `org-chart` | Governance models, org structures, three-lines-of-defense | `- Node: Id=".." Label=".." Parent=".." [Role=".."]` rows (max 12, depth 3) |
| `value-chain` | Value chains, lifecycle stages with sub-activities | `- Step N: Title=".." Sub="a, b"` rows (2-6) + optional `- Support: ".."` band |
| `heatmap-table` | Benchmarks, scorecards by intensity | Markdown table, numeric cells heat-filled; optional `- Scale: rag` |
| `tornado` | Sensitivity analysis | **Series:** Low, High + `- Driver: low, high` rows; optional `- Sort: off` |
| `football-field` | Valuation ranges | **Series:** Low, High + `- Method: low, high` rows; optional `- Marker:` |
| `driver-tree` | Value/KPI decomposition | `- Node: Id Label Value Delta Parent` rows; one root, depth ≤ 3, ≤ 12 nodes |
| `stakeholder-map` | Stakeholder engagement plans | matrix-2x2 Item rows + optional `TargetX`/`TargetY` move arrows |
| `raci` | Roles & responsibilities | Markdown table: activities × people, cells R/A/C/I or blank |
| `what-how-who` | Process grids with per-step accountability | `- Step N: Title=".." What=".." How=".." Who=".."` rows (2-5) |
| `tombstone-grid` | Deal credentials, transaction experience | `- Tombstone: Client=".." Deal=".." Value=".." Date=".." [Role=".."]` rows (max 12) |
| `league-table` | Ranked competitive standings | Markdown table (Rank \| Firm \| Value \| …); in-row bars behind the first numeric column (max 10 rows) |
| `logo-wall` | Client / partner / member grids | `- Logo: file.png "Name"` rows (path optional, max 24) |

### Consulting layout syntax

```markdown
## Slide 2: We recommend consolidating onto two strategic cloud platforms
**Layout:** exec-summary-scqa
- Situation: "Cloud spend doubled to $46M while delivery velocity stalled"
- Finding 1: "61% of spend sits on workloads with <30% utilization"
- Finding 2: "Tooling fragmentation adds 11 days to median release cycle"
- Recommendation: "Consolidate to two platforms; tier workloads — $18M savings by FY27"

## Slide 5: Three levers bridge EBITDA from $42M to $58M
**Layout:** waterfall
**Data:**
- FY24: 42
- Pricing: +9
- Volume: +6
- Cost inflation: -7
- Efficiency: +8
- FY25: total
- Source: "Company financials; team analysis"

## Slide 5b: H1 momentum carries the bridge (subtotals + stacked segments)
**Layout:** waterfall
**Series:** EMEA, NA          ← optional: names the stacked segments
**Data:**
- FY24: 40, 0
- Pricing: +4, +2             ← multi-value delta rows stack segments in one
- Volume: -1, -1                bar (all segments must share the delta's sign)
- H1 run-rate: subtotal       ← auto-computed running-total bar
- Efficiency: +3, +1
- FY25: total

## Slide 6: Workload tiering is the fastest path to value
**Layout:** matrix-2x2
- X-axis: "Implementation effort →"
- Y-axis: "Value at stake →"
- Q1: "Quick wins"
- Item: Name="Workload tiering" X="0.2" Y="0.85"
- Item: Name="Platform exit" X="0.85" Y="0.7"

## Slide 8: Vendor B leads on the criteria that matter most
**Layout:** harvey-scorecard
| Criterion | Vendor A | Vendor B |
|---|---|---|
| Scalability | 2 | 4 |
| Ecosystem | 3 | 3 |

## Slide 9: Spend passes budget in 2026 without intervention
**Layout:** chart-callout
**Visual:** chart:line
**Benchmark:** 60 "FY budget ceiling"
**Data:**
- 2024: 38
- 2025: 52
- 2026: 71
- Callout: "Crossover hits in Q2 2026 — decision needed this quarter"

## Slide 12: Decision required: approve Phase 1 by month end
**Layout:** next-steps
- Step 1: Action="Approve Phase 1 budget ($2.4M)" Owner="CFO" When="This week"
- Step 2: Action="Stand up tiger team" Owner="CTO" When="Within 2 weeks"

## Appendix
## Slide 13: Detailed cost model assumptions
**Layout:** table
...backup slides: marked BACKUP, numbered B·n, exempt from variety/notes warnings
```

```markdown
## Slide 10: Two platforms capture 73% of segment value
**Layout:** mekko
**Series:** Platform A, Platform B, Other
**Data:**
- Compute: 18, 22, 6        ← column width ∝ total, heights = share
- Storage: 9, 11, 4
- Networking: 5, 4, 3

## Slide 11: Three workstreams deliver over six quarters
**Layout:** gantt
**Periods:** Q3 26, Q4 26, Q1 27, Q2 27
- Bar: Row="Infrastructure" Label="Tiering" Start="1" End="2"
- Bar: Row="Infrastructure" Label="Exit" Start="3" End="4"
- Bar: Row="Platform" Label="Pilot" Start="1" End="1"
- Milestone: Row="Platform" Label="GA" At="3"
```

```markdown
## Slide 12: EMEA unit costs run 3x the NA baseline on every metric
**Layout:** heatmap-table
| Region | Cost | Churn % | NPS |
|---|---|---|---|
| NA | $10 | 5% | 62 |
| EMEA | $30 | 7% | 41 |
| APAC | $20 | 6% | 55 |
- Scale: rag                ← optional: red/amber/green terciles instead of
                              the default bg→accent1 heat per column
- Notes: "Heat is normalized per column; text/$/% cells parse as numbers."

## Slide 13: Margin and WACC swings drive a 35-point valuation range
**Layout:** tornado
**Series:** Downside, Upside     ← REQUIRED: exactly 2 names so the rows
**Data:**                          parse as [low, high] pairs
- WACC +/-1pt: -12, +18
- Volume growth: -8, +10
- Gross margin: -15, +20
- Sort: off                 ← optional: keep input order (default sorts by
                              |low|+|high| descending)
- Notes: "Left bars = values[0] (accent2), right = values[1] (accent1)."

## Slide 14: Three valuation methods converge on the $45-52 band
**Layout:** football-field
**Series:** Low, High            ← REQUIRED: 2 names → [low, high] pairs
**Data:**
- DCF: 42, 58
- Trading comps: 45, 52
- Precedent deals: 40, 50
- Marker: Current price, 47 ← optional dashed vertical line + label; must
                              fall inside the data range or it is skipped
- Notes: "Floating bars low→high on a shared nice-interval axis."
```

```markdown
## Slide 15: Pricing drives 60% of the planned revenue uplift
**Layout:** driver-tree
- Node: Id="rev" Label="Revenue" Value="$120M" Delta="+8%" Parent=""
- Node: Id="price" Label="Price" Value="$70M" Delta="+12%" Parent="rev"
- Node: Id="vol" Label="Volume" Value="$50M" Delta="-2%" Parent="rev"
- Node: Id="mix" Label="Product mix" Value="$30M" Delta="+5%" Parent="price"
- Notes: "Root at left, children decompose to the right; Delta colors by sign."

## Slide 16: Move the CFO from skeptic to sponsor before gate review
**Layout:** stakeholder-map
- Item: Name="CFO" X="0.2" Y="0.9" TargetX="0.8" TargetY="0.9"
- Item: Name="COO" X="0.7" Y="0.6"
- Item: Name="CISO" X="0.4" Y="0.3" TargetX="0.55" TargetY="0.35"
- Notes: "X = support, Y = influence (0-1). Targets draw a move arrow."

## Slide 17: Every workstream has one accountable owner from day one
**Layout:** raci
| Activity | Alice | Bob | Cara |
|---|---|---|---|
| Design | A | R | C |
| Build | I | A | R |
| Test | C |  | A |
- Notes: "Cells: single R/A/C/I letter or blank — exactly one A per row."
```

`driver-tree` rows are `- Node:` key-value lines: exactly one root (empty
`Parent=""`), every other `Parent` must name an existing `Id`, Ids unique,
depth ≤ 3 (root + 2), ≤ 12 nodes (`--check` errors otherwise). `Delta`
starting with `+` renders green-ish (`rag_good`), `-` red-ish (`rag_bad`).
`stakeholder-map` reuses the matrix-2x2 engine — same `- Item:` rows, axes
default to "Support →" (x) and "Influence →" (y), overridable with
`- X-axis:` / `- Y-axis:`; `TargetX`/`TargetY` (both required, 0-1) draw a
thin arrow to a hollow circle at the target position. `raci` errors on any
non-blank cell that isn't a single R/A/C/I letter ("A,R" is invalid) and
warns when an activity row has zero or 2+ `A` cells.

`heatmap-table` needs a header row plus body rows with 2+ columns and at
least one numeric body column (`--check` errors otherwise); non-numeric
cells get a plain surface fill and cell text auto-switches black/white by
fill luminance. `tornado` and `football-field` reuse the multi-series data
parser — declare `**Series:**` with exactly two names or the rows won't
parse as pairs (`--check` errors). `football-field` additionally requires
low < high per row. Put `- Scale:` / `- Sort:` / `- Marker:` lines *after*
the data rows — like any `- Key:` field they end the **Data:** block.

`**Benchmark:** <value> "label"` adds a dashed reference line to bar/column/line/area charts. `**Exhibits:** on` (front-matter) numbers every sourced slide's footnote "Exhibit N · Source: ...". `agenda` slides highlight the row matching `- Current: "Section name"`. Section dividers automatically set the top-right section tracker label on subsequent content slides (`- Kicker:` is the bottom takeaway band — a different element).

### architecture-layers

Horizontal stacked bands, listed top layer first. Accent="2" highlights a
band in the secondary accent (use for the layer the slide's action title is
about).

```
## Slide 7: Target Operating Model
**Layout:** architecture-layers
- Layer: Name="Experience" Items="Portal, Mobile, Partner API"
- Layer: Name="AI Services" Items="GenAI copilots, ML scoring" Accent="2"
- Layer: Name="Data Platform" Items="Lakehouse, Feature store"
- Layer: Name="Governance" Items="Policy, Model risk, Audit"
- Notes: Walk bottom-up — governance underpins every layer.
```

### swimlane-process

One lane per actor; steps flow left-to-right within the lane, separated by
`>`. Use `process-flow` instead when a single actor owns the whole flow.

```
## Slide 9: Use-Case Intake Process
**Layout:** swimlane-process
- Lane: Name="Business" Steps="Submit idea > Refine value case"
- Lane: Name="AI CoE" Steps="Screen > Score feasibility > Prioritize"
- Lane: Name="Risk" Steps="Classify risk tier > Approve controls"
- Notes: Handoffs (lane changes) are where cycle time is lost today.
```

### org-chart

Top-down hierarchy (driver-tree's vertical sibling — use `driver-tree` for
value decomposition with numbers, `org-chart` for people/forums/roles).
Root box uses accent1; `Role=` renders as a muted sub-label.

```
## Slide 5: AI Governance Model
**Layout:** org-chart
- Node: Id="steer" Label="AI Steering Committee" Role="Direction & funding"
- Node: Id="coe" Label="AI CoE" Parent="steer" Role="Delivery"
- Node: Id="risk" Label="Model Risk" Parent="steer" Role="2nd line"
- Node: Id="owners" Label="Model Owners" Parent="coe" Role="1st line"
- Notes: Escalation path runs owners -> CoE -> steering committee.
```

### value-chain

Chevron sequence with sub-activities under each stage; `- Support:` adds an
enabling-functions band across the bottom. Use `process-flow` for simple
sequences without sub-activities.

```
## Slide 6: AI Value Chain
**Layout:** value-chain
- Step 1: Title="Data" Sub="Acquisition, Quality, Lineage"
- Step 2: Title="Models" Sub="Build, Evaluate, Register"
- Step 3: Title="Deployment" Sub="Serve, Monitor, Retrain"
- Step 4: Title="Adoption" Sub="Training, Change mgmt"
- Support: "Governance, Security, FinOps"
- Notes: Value leaks at the deployment/adoption seam today.
```

---

## Palette Selection Guide

Palettes (colors + per-theme font stacks, 60-30-10 dominance encoded) are defined once in `scripts/palettes.py`.

| Palette Key | Best For | Theme | Title font |
|-------------|----------|-------|------------|
| `midnight-executive` | Finance, strategy, luxury | Dark | Gill Sans MT |
| `aurora` | AI, tech, innovation | Dark | Gill Sans MT |
| `venture-pitch` | Startups, product launches | Dark | Trebuchet MS |
| `forest` | Sustainability, ESG | Dark | Gill Sans MT |
| `teal-trust` | Healthcare, fintech, services | Dark | Gill Sans MT |
| `charcoal-minimal` | Design, architecture, monochrome | Dark | Arial Black |
| `ocean-gradient` | Data, cloud, deep tech | Dark | Gill Sans MT |
| `swiss-light` | Corporate, education, tutorials | Light | Palatino Linotype |
| `warm-terracotta` | Consumer, hospitality, editorial | Light | Georgia |
| `berry-cream` | Lifestyle, brand, creative | Light | Georgia |

Pick by topic domain, not habit — defaulting to blue is an AI tell. If the user does not specify, use `midnight-executive` for business decks and `aurora` for tech/AI. Light palettes intentionally avoid pure `#FFFFFF` backgrounds.

---

## Design Rules (enforced or checked)

- **Layout variety**: never use the same layout 3+ times in a row (`--check` warns). Alternate text layouts with cards, stats, charts, images.
- **No accent lines under titles** — a hallmark of AI-generated decks. Hierarchy comes from size contrast and whitespace (builders already comply).
- **Every content slide gets a visual**: icon bullets, a chart, an image, or stat cards. `qa_check.py` warns on text-only slides.
- **Sandwich structure**: title/section/closing slides render with a dark gradient for a premium feel; for a light-content deck, set a light deck palette and per-slide override dark on slide 1 and the closer.
- **Left-align body text** (builders comply); center only titles and stat values.
- **Density**: default is `compact` (double-density) — up to 14 bullets flowing into two columns, 8 exec-summary cards in a 2-column grid, 8 points beside a chart/image, 9 comparison rows (~220-word qa warning). `**Density:** comfortable` gives single-column spacing (7 bullets, 5 exec/two-column points). Prefer filling a slide over splitting content across sparse ones. Exception: an `exec-summary` carrying **exactly 3 points** always renders as one column of three full-width cards, in either density — a 2x2 grid with an empty fourth quadrant reads as a mistake and trips the geometry balance check.
- **Speaker notes on every slide** — write `- Notes:` lines in the outline; `--check` warns when missing.

**Data blocks:** blank lines inside `**Data:**` are ignored. Each data row must be `- label: value` (multi-series: `- label: v1, v2`).

**`chart:scatter`** is a native XY chart — markers only, no connecting line, palette accents per series, legend when 2+ series. Its rows map onto x/y as follows:
- **Single-value rows** (`- 2019: 42`): x = the label read as a number when *every* label is numeric (years count; `1,200`, `$42`, `12%` parse too), otherwise x = the 1-based row index (`--check` warns); y = the value. One series, named by `**Series:**` (or the heading).
- **Multi-value rows** with `**Series:** X, Y1[, Y2 …]` (`- Acme: 12, 30, 41`): the **first column is X**, every further column is one Y series named by the matching `**Series:**` name — one scatter series per Y column, all sharing the X column. The row label is not plotted (it is a row identifier only). Multi-value rows with fewer than 2 series names are a `--check` error.
- `- Axis-Max:` sets the Y axis; `- Benchmark:`, `- CAGR:` and `- Value-Line:` do not apply to scatter and are ignored. `deck_to_outline.py` recovers a scatter back to this contract (`- x: y` rows, or `**Series:** X, Y1, …` + `- Item N: x, y1, …`).

---

## Slide Dimensions & Typography

Widescreen **13.33" × 7.5"** by default; `**Size:** 4:3` or `--size 4:3` gives 10" × 7.5" (builders rescale automatically).

Body/bullets render at 16–17pt, captions/labels 10–13pt, titles 30–54pt in the palette's title font. Never go below 10pt (enforced by `qa_check.py`). Fonts are web-safe PowerPoint defaults; PowerPoint substitutes automatically when a font is missing (python-pptx cannot embed fonts).

---

## Charts

Declare `**Visual:** chart:<type>` plus a `**Data:**` block (see Slide 3 above) and the build pipeline generates a native, palette-styled chart with a transparent background — fully editable in PowerPoint. For custom chart code or shape-based visuals (progress bars, stat cards), see `references/charts-guide.md`.

## Diagrams

- `diagram:file.mmd` / `diagram:file.dot` — auto-laid-out graph rendered to
  PNG at build time (source under `assets/diagrams/`). Native layouts first;
  see `references/diagrams.md` for the routing rule.

## Newer directives (June 2026)

**Heading attributes** — set layout/palette on the heading line, keeping
directives off content bullets:

    ## Slide 7: Margin bridge tells the story {layout=waterfall palette=aurora}

Values must be unquoted (`layout=two-column-split`, not `layout="..."`);
unknown keys emit a warning. An explicit `**Layout:**` line still wins if
both are present.

**Deck meta:**
- `**Auto-Agenda:** on` — auto-insert an agenda slide (sections = your
  section-divider headings, appendix excluded) after the title slide.
- `**Auto-Agenda:** track` — additionally insert a "where we are" agenda with
  the current section highlighted after every section divider.
- `**Stamp:** DRAFT` — bordered status tag on every slide (also:
  CONFIDENTIAL, FOR DISCUSSION).
- `**Tracker:** tabs` — replaces the plain-text section label on content
  slides with a compact tab strip top-right: one small chip per section,
  current section filled, the rest outlined (names truncated to 14 chars).
  Requires section dividers (`--check` warns and the build falls back to
  the plain label without them). With 6+ sections, or when the strip would
  outgrow its budget, slides get a compact "n/N · Section" text instead
  (one build warning). Auto-inserted agenda slides and appendix slides
  (BACKUP) keep their existing labels. Compatible with `**Auto-Agenda:**`.
- `**Sections:** off` — disable native PowerPoint sections. With 2+
  section dividers the build injects real sections (collapsible in
  PowerPoint's slide sorter / thumbnail pane): pre-divider slides →
  "Opening", each divider starts a section named by its heading, appendix
  slides → "Backup". Default on at 2+ dividers, off below; styled builds
  only (skipped in template mode). Rebuilds produce identical section ids.
  `**Sections:** on` has no effect when fewer than 2 section dividers are present.
- `**References:** on` — appends an auto-generated "Sources" backup slide
  (appendix-flagged bullet list) aggregating every unique `- Source:` value
  with its slide numbers — or exhibit numbers when `**Exhibits:** on` —
  e.g. "Slides 4, 7 — Gartner 2026". Default off; with no sources the
  slide is skipped with a warning.

**CSV / Excel chart data** — `- Data-File: data/q3.csv` (or `.xlsx` /
`.xlsm`; Excel needs the optional `openpyxl` package and reads the active
sheet) loads a chart's `**Data:**` block from a file (resolved against the
outline dir, then the assets dir, then as given). Two columns =
`label,value` rows (optional header row tolerated; `total`/`end` sentinel
values for waterfall closing bars work only in this 2-column form, not in
3+ column multi-series files); 3+ columns require a header row — first
column is labels, remaining headers become `**Series:**` names (unless
`**Series:**` is set explicitly). `$`, `,`, `%` are stripped like inline
data. Missing files and malformed rows are `--check` errors; when both
`**Data:**` rows and `- Data-File:` are present, the file wins (with a
warning). Every loaded file stamps a provenance line ("Data source:
q3.csv (8 rows)") into the slide's speaker notes so exhibit numbers stay
reproducible from source data.

**Evidence ledger** — `**Evidence-File:** evidence.json` (front-matter;
resolved like Data-File paths) activates source traceability for material
numbers. The ledger maps ids to citations:

```json
{
  "E1": {"value": "42%", "source": "IDC MarketScape 2026", "page": "12"},
  "E2": {"value": "$4.2B", "source": "Company 10-K FY25", "url": "https://..."}
}
```

`source` is required; `value` (the number exactly as it appears on the
slide), `page`, `url`, and `claim` are optional. Slides cite entries with
`- Evidence: E1, E2`. Effects:

- Referenced sources synthesize the slide's `- Source:` footnote (an
  explicit `- Source:` wins) and full citations are appended to speaker
  notes ("Evidence: E1 — IDC MarketScape 2026, p.12").
- **The traceability gate:** with a ledger declared, any non-appendix
  slide carrying material statistics but neither `- Source:` nor
  `- Evidence:` is a fail-fast `--check`/build **error** (without a
  ledger it stays the usual unsourced-stat warning).
- Unknown ids and malformed/missing ledger files are errors; a
  referenced `value` that doesn't appear on its slide, and ledger entries
  nothing references, are warnings.
- `diff_deck.py` verifies every referenced `value` survives into the
  built deck's text unchanged (an error when a material number was
  dropped or transcribed differently).

**Handout / pre-read** — `python3 scripts/gen_handout.py outline.md
[--output handout.md]` (default `<outline>-handout.md`) renders the
outline as a readable markdown pre-read: deck title as H1, one `## N.
<action title>` section per slide (auto-agenda slides skipped, appendix
slides under a `# Appendix` divider), bullets, data/table blocks as
markdown tables, `> Talk track:` blockquotes from speaker notes, and
`*Source: ...*` attributions.

**CJK font stacks** — when any slide text contains CJK codepoints
(Chinese/Japanese/Korean), the build swaps all palette fonts to the first
available of PingFang SC / Hiragino Sans GB / Noto Sans CJK SC /
Microsoft YaHei (deck-wide, warned once); if none is found the palette
fonts are kept and PowerPoint's own substitution applies.

**Image visual options** — pipe-separated options after any image visual
path, combinable:

    **Visual:** user-image:hero.png|alpha=85|duotone

- `alpha=N` — picture opacity N% (0–100), written as native
  `a:alphaModFix` (stays editable in PowerPoint).
- `duotone` — brand-tint the photo with a palette duotone
  (bg → accent1), the standard trick for unifying mixed stock photos.

Applies wherever image visuals are placed: title hero, two-column-split,
full-image, chart-callout image fallback, and comparison
`**Visual-Left:** / **Visual-Right:**`. Malformed options warn and are
ignored. Options are ignored in template mode (a `--check` warning is
emitted so the outline can be cleaned up before distribution).

**Per-slide tags:**
- `- Sticker: ILLUSTRATIVE` — small bordered tag top-right (free text,
  uppercased; e.g. NON-EXHAUSTIVE, PRELIMINARY, BACKUP, FOR DISCUSSION).
  Visual mirror of the deck-level `**Stamp:**`, which stays top-left.
  Coexists with the section label: the sticker drops below it when both
  are present.
- `- Kicker: "So-what sentence"` — bold takeaway band across the bottom of
  the content area, just above the source/footer. Make it advance the
  argument — `--check` warns when it merely restates the slide title.
  Skipped on title/closing/section-divider/full-image layouts.

**Action-title lint (`--check`, warnings):** on action-title layouts,
headings over 15 words, exhibit headings (chart visuals, waterfall, mekko,
bar-mekko, chart-callout) with no number, and headings joining two messages
with " and " are flagged.

**Overflow guard (`--check` + build):** bullet-list and two-column-split
text is measured (PIL glyph metrics when available, else a calibrated
char-width estimate) against the layout's vertical budget. 100–140% of
capacity warns and the build writes a `normAutofit` shrink hint (floor 80%
font scale) so text scales down instead of spilling; over 140% is a
`--check` **error** — split the slide or cut copy. Headings predicted to
wrap past 2 lines also warn.

**AI-tell lint (`pptx_lint.py`):** thin accent lines under titles and long
centered body copy — both hallmarks of AI-generated decks — are warned on
any .pptx, including imported ones.

**Chart/exhibit directives (per slide):**
- `- Axis-Max: 50` — pin the chart value axis; use the same value on sibling
  slides for honest same-scale comparison.
- `- CAGR: on` — growth arrow from first to last column (bar/column/line,
  single series).
- `- Bracket: FY25, FY27` — waterfall difference bracket between two named
  bars; auto-labels the % change. Optional third part = custom label:
  `- Bracket: FY25, FY27, "Run-rate reset"`.
- `- Axis-Break: on` — waterfall/bar-mekko: when the bottom 75%+ of the axis
  carries no detail (a tall base bar dwarfing the deltas), compress the empty
  zone and draw break marks so the differences become readable. No-op when
  the bars are already legible.

**New layouts:**
- `bar-mekko` — profit pool: `- Bar: Label="EMEA" Size="30" Value="6"`
  (width ∝ Size, height = Value; Size>0, Value>=0). 2+ bars required.
- `matrix-2x2` bubbles — add `Size="40"` to `- Item:` rows; bubble area
  scales with Size (BCG growth–share style).
- `heatmap-table` — markdown table with per-column heat fills (bg→accent1);
  `- Scale: rag` switches to red/amber/green terciles (palette keys
  `rag_bad`/`rag_mid`/`rag_good`, overridable in custom palettes).
- `tornado` — sensitivity bars off a central spine; needs `**Series:**` with
  2 names + `- Driver: low, high` rows; `- Sort: off` keeps input order.
- `football-field` — valuation range bars on a shared axis; needs
  `**Series:**` with 2 names + `- Method: low, high` rows (low < high);
  `- Marker: label, value` adds a dashed reference line.
- `driver-tree` — left-to-right value decomposition from `- Node:` rows
  (one root, depth ≤ 3, ≤ 12 nodes); Delta colors green/red by sign.
- `stakeholder-map` — support × influence grid (matrix-2x2 engine);
  `TargetX`/`TargetY` on an item draws a move arrow to a hollow circle.
- `raci` — activities × people markdown table; R/A/C/I letter chips
  (R=accent1, A=accent2, C=accent3, I=muted); one Accountable per row.

**Custom palettes:** drop `<name>.json` into `<assets>/palettes/`:

    {"bg": "101820", "bg_deep": "0A0F14", "surface": "1E2A33",
     "accent1": "FEE715", "accent2": "8DA9C4", "accent3": "5C946E",
     "text": "F4F4F4", "text_muted": "9DB2BF", "dark": true}

All nine keys required; colors must be 6-char hex. Optional keys:
`font_title`, `font_body`, `font_label`, `chart_series` (list of hex —
chart series colors), `motif`. Then `--palette <name>`.

**Brand kit ingestion:** generate a custom palette straight from a
company's web presence instead of hand-writing the JSON:

    python3 scripts/brand_kit.py acme.com --name acme --assets-dir assets

Uses the Brandfetch API when `BRANDFETCH_API_KEY` is set (brand colors,
font names, logo → `<assets>/brand/acme-logo.png`); otherwise scrapes the
homepage plus linked same-origin CSS for hex/rgb() colors ranked by
frequency. Colors are classified into the nine palette keys (most-frequent
very-dark vs very-light color decides the theme; saturated candidates
become accents) with contrast-checked nudging — text ≥ 4.5:1 on bg/surface,
accents ≥ 3.0:1 on bg (warns when a target can't be met). The result is validated through the custom-palette
loader before `<assets>/palettes/acme.json` is written, and a swatch table
(role, hex, contrast vs bg) is printed. Then build with `--palette acme`.


## Newer directives (August 2026)

**Deck author** — `**Author:** "Alex Kim"` sets `dc:creator` and
`cp:lastModifiedBy` on the saved file. Every build sanitizes the package on
save (`save_hygiene.py`): Keynote-safe slide size, the missing
`notesMasterIdLst`, the vestigial `printerSettings` part, and the stock
python-pptx identity (`Steve Canny`, "generated using python-pptx") are
fixed, and `dc:title` defaults to the deck's title-slide heading. Without
`**Author:**` a *real* existing creator is preserved — template-mode decks
keep the template's author — and only the stock/empty identity is replaced.

**Number-Format** — one number policy for the whole deck:

    **Number-Format:** units=M decimals=1 negatives=parens multiples=1x pct=1

| Key | Values | Default | Effect |
|---|---|---|---|
| `units` | `none` `K` `M` `B` | `none` | Appends a suffix — it never rescales. Pass values already in the unit you name. |
| `decimals` | `auto` or `0`–`3` | `auto` | `auto` = 0 for integers, 1 otherwise. |
| `negatives` | `minus` `parens` | `minus` | `minus` draws the typographic en dash (`–3.2`); `parens` draws `(3.2)`, with the unit inside: `(8M)`. |
| `multiples` | `1x` `1.0x` | `1x` | Multiple spelling — parsed and validated, but no builder calls it (see the boundary note). |
| `pct` | `0`–`3` | `0` | Decimals on percentages. |

Omit the directive entirely and every label renders byte-for-byte as
before (plain ASCII `-3.2`) — declaring it is what switches the deck onto
the policy. The one exception is percentage rounding, described in the
boundary note below. Unknown keys and out-of-range values are `--check`
errors.

**Boundary (as of the Cycle 5 follow-ups):** `units`, `decimals` and `negatives` reach
every builder-drawn value label, and are also written onto *native chart*
data labels as a number format (`#,##0.0"M";(#,##0.0"M")`) — the labels
stay switched off, so this only decides how they read if you turn them on
in PowerPoint. `pct` reaches every builder-drawn percentage: funnel step
conversions, waterfall brackets, the CAGR arrow and stacked-100/mekko
shares — and those percentages are now rounded half-up and have their
thousands grouped (`12.5% -> 13%`, `+1,204%`) even with no directive
declared, the one place where omitting it is not byte-for-byte the old
output. `multiples` is honoured by `numfmt.fmt_multiple` only and nothing
calls it — an `x` suffix in a **Data:** row is dropped when the value is
parsed, so no builder can tell `3.5x` from `3.5`. Rounded parts are made to add up
(largest-remainder) in waterfall deltas, stacked segments and stacked-100
absolute labels, so a column's labels sum to the label on the stack.

**Style-File** — house-style lints (`house_style.py`), all warnings,
surfaced in `--check`:

    **Style-File:** style.json

Resolution order: `--style <path>` (CLI, resolved against the working
directory) > `**Style-File:**` (outline dir, then assets dir) >
`<assets>/style.json` > built-in defaults. An explicitly named file that
does not exist is an error, as is malformed JSON or an unknown key.

```json
{
  "banned_phrases": [],
  "marketing_adjectives": ["leverage", "world-class", "seamless", "…"],
  "hedges": ["arguably", "somewhat", "possibly", "…"],
  "judgemental": ["failed to", "mismanaged", "…"],
  "title_case": "sentence",
  "title_punctuation": "none",
  "max_title_words": 15,
  "near_duplicate_ratio": 0.85
}
```

Every key is optional and merges over the defaults shown (the four word
lists ship with sensible contents; `banned_phrases` starts empty).
`title_case` is `sentence` (default) | `title` | `any` — use **`any`** as
the escape hatch on product-name-heavy decks, where capitalised names
otherwise read as case drift. `title_punctuation` is `none` (default) |
`period` | `consistent`. `max_title_words` (15) moved here out of
`build_deck`, so a style file can raise or lower the heading-length cap.
`near_duplicate_ratio` (0.85) is the similarity above which two headings
are reported as near-duplicates.

The rules: "not X but Y" constructions, marketing adjectives / banned
phrases / hedges / judgemental phrasing about client teams, passive voice
in action titles, duplicate and near-duplicate titles, title punctuation
drift, sentence-vs-title case drift, orphan footnote markers (a body
marker with no source entry and vice versa), and typography slips (double
spaces, `..`, space before punctuation). Findings aggregate per slide per
rule: `Slide 7: house style — marketing adjectives: leverage, world-class`.

## Newer directives (Cycle 6 — revision loop, data binding & memory)

**Repeat** — one slide per data row. Put it on a *slide*, not in the deck
front-matter:

```markdown
## Slide 7: {name} ships in {due} with {owner} accountable
**Layout:** cards-3
**Repeat:** data/initiatives.csv [rows=1-20] [sort=owner]
- Card 1: Name="Scope" Title="{scope}" Bio="Initiative {row} of {rows}"
- Card 2: Name="Value" Title="{value}" Bio="Run-rate impact"
- Card 3: Name="Risk" Title="{risk}" Bio="Delivery confidence"
- Source: "Initiative charters, {owner}"
- Notes: "Walk the {name} card set left to right."
```

The file is a `.csv` or `.xlsx` (header row = field names, resolved like
`- Data-File:`). The slide is cloned once per row and `{column}` is
substituted into **every string the slide carries** — heading, bullets,
cards, stat tiles, table rows, steps, kicker, source, notes. Optional
keys go in **literal square brackets** after the path: `[rows=a-b]`
(1-based, inclusive) selects a range, `[sort=col]` orders by a column.
Without the brackets they are read as part of the filename.

| Rule | Behaviour |
|---|---|
| Unknown `{field}` | Build error naming the slide, the field and the file's actual columns |
| `{row}` / `{rows}` | Builtins: 1-based row number and total row count. They **shadow** same-named columns — a column called `row` or `rows` is unreachable |
| `{accent}` | **Reserved** — it is the rich-text markup (`{accent}…{/}`), never a Repeat field. A column literally named `accent` is not substitutable |
| `{{` / `}}` | Literal `{` / `}` in the output |
| A `{field}` at the **end of a heading** | Eaten by the heading-attribute syntax (`{layout=… palette=…}`) — it warns and the field vanishes. Never end a `## Slide N:` line with a substitution; move it earlier in the sentence |
| Missing / unreadable file, unknown `sort=` column | Build error |
| More than 60 expanded slides | Warning |
| More than 200 expanded slides | Error |

`--check` prints `[WARN] [info] Slide N expands to K slides (file)` so you can see
the deck's real length before building. Layout is decided **once**, from the
un-substituted template slide, so every clone shares it (that is the point).
Slide numbering everywhere downstream — `diff_deck`, `qa_checklist`,
`geometry_report`, the QA record — is post-expansion, so `Slide 9` means the
ninth slide of the built deck. A zero-row selection produces zero slides and
cannot detect an unknown-field typo, so sanity-check the row range.

**Copy: locked** — verbatim copy, for legal, regulatory and
approved-messaging decks:

```markdown
**Copy:** locked
```

or `python3 scripts/build_deck.py outline.md --verbatim --output deck.pptx`.
The directive value must be exactly `locked` (`**Copy:** locked (approved)`
is *not* recognised and warns), and it must sit in the deck front-matter —
a `**Copy:**` line inside a slide body is not the deck lock and warns.

What it guards — the four ways the normal pipeline touches your words:

1. **Quote stripping.** Copy fields (`title`, `subtitle`, `heading`,
   `caption`, `quote`, `attribution`, `notes`, `source`, `kicker`,
   `sticker`, `value`, `label`, the SCQA and matrix labels…) keep their
   wrapping quotes instead of having them removed. Directive fields still
   unquote — `**Palette:** "aurora"` is a directive, not copy.
2. **Markup parsing.** `**bold**` and `{accent}…{/}` render literally, as
   the characters you typed. Locked copy cannot also be styled copy.
3. **List truncation.** Any builder cap that would drop a row becomes a
   build **error** — `Slide 1: 5 cards exceed cards-3 cap 3 — verbatim
   mode refuses to drop content` — instead of a silent trim. The
   caps come from `diff_deck.LAYOUT_CONTENT` plus a supplemental table for
   the keys it does not track (matrix-2x2 / stakeholder-map items 12,
   dashboard tiles 6). Fix it by splitting the slide, not by trusting the
   trim.
4. **The diff.** `diff_deck.py` switches to a raw compare — case, quotes and
   markup characters all significant — and any deviation is an error
   regardless of `--strict`. It detects the mode from the deck's own tags
   (`PK_FLAGS` carries `verbatim`) or the outline, and `--verbatim` /
   `--no-verbatim` force it either way.

Not guarded, deliberately: font shrink (`normAutofit`) still runs — it
changes the size, not the words — and catastrophic overflow is still an
error as always. Chrome (title / subtitle / section labels) is still
compared normalized, because builders legitimately restyle it (a section
label rendered in caps is not a copy deviation); a chrome miss stays a
warning. The pipeline has never smart-quoted or re-wrapped text.

**Profile** — the client's standards, saved and reused:

```markdown
**Profile:** acme
```

or `--profile acme` (which wins over the directive). Reads
`<assets>/profiles/acme.json`:

| Key | Type | Effect |
|---|---|---|
| `palette` `footer` `stamp` `density` `size` `author` `number_format` `style_file` `page_numbers` `tracker` | string | Deck-meta defaults — used only where the outline does not set them itself |
| `banned_phrases` | list | Merged into the house-style word lists (`--check` findings) |
| `mandatory_slides` | list of `{"layout": …}` and/or `{"title_contains": …}` | A missing requirement is a `--check` **error** |
| `preferred_layouts` | list | Tie-breaks *auto-chosen* layouts only |
| `notes_style` | `bullets` \| `prose` \| `none` | Speaker-notes lint (warnings) |
| `logo` | string | Image path (literal, outline dir, assets dir, or `<assets>/brand/` — where `brand_kit.py` writes `<name>-logo.png`): a corner badge bottom-right on `title`/`closing` slides and a small mark in the footer band (left of the page number) on every footer-bearing slide, named `sNN-logo`. Template-mode slides are left alone. A missing file **warns** and the deck builds without it |

Unknown keys, wrong types and an out-of-range `notes_style` are errors, so a
malformed profile fails at `--check` rather than half-applying. Precedence,
strictest first: **CLI flag > `**Variant:**` preset (palette and density
only) > outline front-matter > profile > built-in default.** The variant sits
above the outline and the profile for those two keys because `apply_variant()`
runs after `profiles.apply()` and fills a palette/density the CLI left unset;
for every other key the profile fills in beneath the front matter. All of
this is resolved inside `outline_replay.replay()`, so every tool that
replays the outline — `diff_deck`, `qa_checklist`, `qa_summary`,
`patch_deck`, `refresh_deck`, `gen_handout` — sees the same layouts,
density and truncation caps the build used; an outline `**Profile:**` is
picked up automatically, a CLI-only `--profile` has to be repeated on those
tools (`diff_deck.py --profile`, `qa_checklist.py --profile`,
`qa_summary.py --profile`). `preferred_layouts` only re-picks layouts the
builder chose itself — an explicit `**Layout:**` is never overridden — and
only within the ambiguity pairs (bullet-list ↔ exec-summary, cards-3 ↔
cards-4, table ↔ heatmap-table). It is capacity-guarded: when the preferred
layout would truncate the slide's content it is skipped with a warning
rather than dropping rows.

Write one from a company's web presence with
`python3 scripts/brand_kit.py acme.com --name acme --profile` (palette,
logo and sensible defaults), or hand-author the JSON.

**Profile memory (`lessons.md`).** `python3 scripts/qa_summary.py deck.pptx
--outline outline.md --profile acme` appends QA findings that have now
recurred across **two or more runs** for that profile to
`<assets>/profiles/acme.lessons.md`, dated, de-duplicated, with the counts
tracked in a sibling `acme.lessons.json`. Findings are matched by signature
(the check, the level and the message with slide numbers and figures
normalised away), so "Slide 3: 12 shapes overlap" and "Slide 9: 481 shapes
overlap" count as one recurring problem. The `.md` is plain markdown —
prune it freely. **Read it before writing the outline for that client**;
it is the skill's memory across engagements. (Known wrinkle: the "seen N×"
count is written when the line is first appended and is not updated later,
so treat it as "at least N".)

## Newer directives (Cycle 7 — think-cell parity)

### Data rows under charts (`- Data-Rows:`)

```markdown
## Slide 4: Revenue compounds while margin holds
**Layout:** chart-callout
**Visual:** chart:column
**Series:** Revenue, Margin %
**Data:**
- FY22: 120, 30
- FY23: 150, 28
- FY24: 190, 33
- Data-Rows: values, Margin %
```

Prints think-cell-style data rows under the chart, one text cell per
category, aligned to the category centres, with the row label right-aligned
in the gutter left of the plot (long labels truncate with `…`). Comma list
of row kinds:

- `values` — every plotted series' values, one row per series, label = the
  series name.
- `growth` — period-over-period % of the first plotted series; the first
  cell is `–`.
- `cagr` — a single row with the start→end CAGR printed under the last
  column, label `CAGR p.a.`.
- a **Series:** name — that series is *removed from the plotted set* and
  printed as a row instead (the classic "margin % as a row under the
  revenue bars" move). Unknown names warn and drop the row; the last
  remaining series is never removed (it prints as a row but stays plotted).

The chart box is shortened 0.42 in per printed row, never below a 2.2 in
chart height — rows beyond that warn and are dropped. Applies to vertical
categorical kinds only (`bar`, `column`, `line`, `stacked`, `stacked-100`,
`combo`); on any other kind the field warns and is ignored. Values format
through the deck number policy (`fmt_abs`/`fmt_pct`); a series whose name
contains `%` prints as percentages. The first row prints in the text color,
further rows muted, with a thin rule above the block.
### Gantt parity

The `gantt` layout grows think-cell-style roadmap furniture. A gantt using
none of the new keys renders exactly as before.

```markdown
## Slide N: Two programs deliver across eight quarters
**Layout:** gantt
**Periods:** quarters FY26-FY27
- Bar: Row="Platform > Pilot" Label="Pilot" Start="1" End="3" Responsible="A. Chen"
- Bar: Row="Platform > Scale" Label="Scale" Start="4" End="8" Responsible="B. Ruiz" Remark="At risk"
- Bar: Row="Infra" Label="Tiering" Start="2" End="6"
- Milestone: Row="Infra" Label="Exit" At="6"
- Span: Rows="Platform, Infra" Label="Wave 1"
- Shade: Periods="4-5" Label="Change freeze"
- Line: At="6.5" Label="Go-live"
```

- **Period expansion** — `**Periods:** quarters FYa[-FYb]` expands to
  `Q1..Q4` per fiscal year with a spanning year-header row above the period
  labels; `**Periods:** months Mon YYYY - Mon YYYY` expands month names (the
  year header appears only when the range crosses a year boundary). Plain
  comma lists are unchanged. Bar `Start`/`End` and `At=` positions count
  **expanded** columns (FY26-FY27 quarters → 1-8). A malformed or reversed
  range is a build error; the raw string stays in the outline — expansion
  happens in the builder.
- **Label columns** — Bar attrs `Responsible=".."` / `Remark=".."` add up to
  two 1.3in right-hand columns (headers `RESPONSIBLE` / `REMARK`); the grid
  shrinks to fit. The first bar's value per row wins; conflicting values on
  one row warn.
- **Row brackets** — `- Span: Rows="A, B" Label=".."` draws a left square
  bracket across those rows with a rotated label; row labels shift right
  0.35in. One nesting level: overlapping spans warn and are skipped.
- **Shaded bands** — `- Shade: Periods="3-4" Label=".."` (or `Periods="3"`)
  tints those period columns behind the bars (translucent surface, label
  centred at the band top). Malformed/out-of-range → warn + skip.
- **Milestone lines** — `- Line: At="2.5" Label=".."` draws a dashed
  vertical line across the grid at that period position (`At=N` = middle of
  period N), label above the grid. Out-of-range → warn + skip.
- **Hierarchy** — `Row="Parent > Child"` renders the parent as a bold group
  header row with children indented beneath it (bars on the header only when
  a Bar names the parent exactly). Two levels max; deeper warns and
  flattens.
### Difference arrows

`- Diff-Arrow:` (repeatable) draws the delta between two bars on
bar/column/line/stacked/stacked-100 chart visuals:

```markdown
**Visual:** chart:bar
**Series:** Revenue
- Diff-Arrow: FY24, FY26            ← total difference, default mode pct
- Diff-Arrow: FY24, FY26 both       ← modes: abs | pct | both
**Data:**
- FY24: 50
- FY26: 62
```

Total form (`CatA, CatB [mode]`): single series compares the two bar values,
multi series the per-category totals (stacked reading); drawn as leaders to a
double-headed connector just right of the plot with a bubble label. Level
form (`SerA@Cat -> SerB@Cat`, `->` or `→`) compares two series at one
category, directly above its cluster. Labels are `+12` / `+24%` (of the
*from* value) / `+12 (+24%)` via the deck number policy; a zero/negative
base falls back to abs with a warning. Unknown names warn and skip the
arrow; `hbar` warns and skips; a spec matching neither grammar is a
fail-fast `--check` error. Placement uses the plot-box approximation —
confirm in visual QA.

### Chart option set

Per-slide chart fields (all optional; bad values warn and fall back):

```markdown
**Visual:** chart:bar               ← chart:stacked is also new (COLUMN_STACKED)
**Series:** Actual, Target
- Target: Target                    ← series → short accent ticks per category
- Sort: desc                        ← desc | asc, bar/column/hbar
- Ordinal: on                       ← "1. " prefixes in plotted order
- Units: $M                         ← muted label above the plot's top-left
- Gap-Width: 50                     ← native, clamped 0–500
- Label-Rotation: 45                ← native catAx bodyPr rot, clamped ±90
- Label-Skip: 2                     ← native c:tickLblSkip
- Category-Break: FY25 "Actual | Forecast"   ← dashed separator after FY25
- Value-Line: mean "Ø benchmark"    ← mean | median | wmean (computed)
```

- `- Value-Line: mean|median|wmean ["label"]` computes the reference value
  from the plotted data (`wmean` needs exactly 2 series: Σ(v·w)/Σw with
  series 2 as weights). The numeric `- Value-Line: Label, 40` form is
  unchanged.
- `- Error-Bars: Low, High` (bar/column/line) turns the two named series
  into native custom min/max whiskers on the first remaining series and
  removes them from the plot (naming ≠ 2 series is a `--check` error).
- `- Target: <series>` (bar/column) plots the named series as think-cell
  target ticks instead of bars, with a `— <name>` legend chip.
- `- Total-Line: on` (`chart:stacked-100`, `mekko`) draws the muted 100%
  rule + label at the top of the plot.
- `- CAGR: series` (multi-series line/column) labels each series
  `<name> +8.2% p.a.` at its right end; `- CAGR: FY22-FY24, FY24-FY27`
  (single series) draws one arrow + `+12.4% p.a.` label per span.
  `- CAGR: on` is unchanged.

Full details in `references/charts-guide.md` ("Difference arrows" and
"Chart options").

### New layouts: what-how-who, tombstone-grid, league-table, logo-wall

Explicit `**Layout:**` only — the smart-layout picker never auto-selects
these. Every one is drawn from shapes (no native tables) and grouped into a
single `sNN-exhibit` group (editability L5).

**`what-how-who`** — process grid: a chevron step header (first step
accent-filled, like `process-flow`) over a striped three-row matrix with
muted WHAT / HOW / WHO labels in the left gutter. 2–5 steps (fewer than 2
is a `--check` error; extra steps are dropped with a warning):

```markdown
## Slide N: Delivery model spans three steps
**Layout:** what-how-who
- Step 1: Title="Diagnose" What="Baseline the costs" How="Bottom-up model" Who="Deal team"
- Step 2: Title="Design" What="Target operating model" How="Workshops" Who="CoE"
- Step 3: Title="Deliver" What="Migration waves" How="Agile pods" Who="IT + vendor"
```

**`tombstone-grid`** — bordered deal-credential cards in an auto grid (3
per row up to 6 cards, then 4 per row; max 12, extras warn + truncate).
Client renders bold, Value in the accent color, Deal/Date/Role muted.
`Client=` and `Deal=` are required per row (`--check` error):

```markdown
## Slide N: Selected transaction experience
**Layout:** tombstone-grid
- Tombstone: Client="Acme Corp" Deal="Sale to Beta Holdings" Value="$1.2B" Date="Mar 2026" Role="Exclusive advisor"
- Tombstone: Client="Gamma Inc" Deal="IPO on NYSE" Value="$800M" Date="Jan 2026"
```

**`league-table`** — a ranked standings table drawn from shapes (not a
native PowerPoint table). The rank column (column 0 when every body cell
is an integer) renders as bold accent numbers with no fill; the first
all-numeric non-rank column gets an in-row percent bar (a strip along the
row bottom, under its right-aligned value — kept clear of the glyphs so
contrast QA passes), scaled to the column max. Numeric columns are
right-aligned. Max 10 body rows (extras warn + truncate); a missing table
is a `--check` error:

```markdown
## Slide N: We rank second on announced volume
**Layout:** league-table
| Rank | Firm | Volume $B | Share |
|------|------|-----------|-------|
| 1 | First Bank | 42.5 | 18.2% |
| 2 | Our Firm | 38.1 | 16.3% |
```

**`logo-wall`** — a uniform tile grid of client/partner logos, images
contain-fitted onto surface tiles with a 10pt muted name caption below.
The path is optional — a name-only row renders a centred name tile; a
missing image file warns and falls back to the name tile. Max 24 logos
(extras warn + truncate):

```markdown
## Slide N: Clients we serve
**Layout:** logo-wall
- Logo: logos/acme.png "Acme Corp"
- Logo: "Beta Holdings"
```

### Multi-state bullets (cell-state tokens as bullet markers)

On `bullet-list` and `two-column-split`, a bullet that starts with a
cell-state token followed by a space draws that token in the marker slot
instead of the accent square (the same mechanics as `icon:` prefixes).
The token vocabulary is shared with table cell states: `harvey:3[/8]`,
`rag:green|amber|red`, `bar:64%`, `dots:3[/5]`, `[x]` / `[ ]`:

```markdown
- rag:green Workstream A on track
- rag:red Data migration slipping
- harvey:3 Integration 3/4 complete
- [x] Contracts signed
```

A malformed token (`rag:blue`, `bar:120`) warns and renders as literal
text; a bare token with no trailing text also stays literal.

### Table upgrades (`table`, `comps-table`)

Options on a `table` slide (all optional — a table using none of them
renders exactly as before):

```markdown
## Slide 9: Comparable multiples cluster at eight to nine times
**Layout:** table
- Align: auto                 ← auto (default) | left | decimal
- Units: , $M, x, %           ← italic muted units row (one entry per column,
                                blanks allowed; count must match the columns)
- Zebra: off                  ← disable the alternate row fill
- Total-Row: auto             ← on | auto (default) | off — bold last body
                                row + thin accent top rule; auto styles it
                                when the first cell matches /^total/i
- Format: Margin < 10 red; Share >= 20 green
                              ← threshold coloring (repeatable and/or
                                `;`-separated): `<column> < <= > >= = <value>
                                red|amber|green` — matching numeric cells get
                                rag-colored bold text
- Summary: mean, median, quartiles
                              ← appended rows computed per numeric column
                                (quartiles = Q1 and Q3 rows); NM/blank cells
                                excluded; decimals = column max, capped at 2
- Sort: EV/EBITDA desc        ← `<column header> asc|desc` — body rows sorted
                                by that column's numeric value, NM last;
                                header/units/total/summary rows stay pinned
| Company | Revenue | EV/EBITDA | Margin |
|---|---|---|---|
| Acme | 1,200 | 7.5 | 22% |
| Beta | 800 | NM | 8% |
```

- **Numeric alignment** (`- Align: auto`, the default): a column whose body
  cells are all numeric-ish — a number with optional `$ € £` prefix /
  `% x ×` suffix, an NM marker, or blank — is right-aligned, header
  included. `decimal` additionally pads trailing decimals per column to the
  column max (display only: `7.5` next to `9` renders `9.0`). `left` is the
  pre-Cycle-7 behavior.
- **NM cells**: `NM`, `n.m.`, `n/m`, `N/A`, `n/a`, `–`, `-` render muted
  italic and are excluded from numeric-ness, summary stats, sorting order
  (they sort last) and threshold formats.
- **Cell states**: a body cell that is exactly a cell-shapes token —
  `harvey:3`, `harvey:5/8`, `rag:green|amber|red`, `bar:64%`, `dots:3/5`,
  `[x]` / `[ ]` — draws the micro-shape centred over the cell (the cell
  text stays empty). A malformed token (`rag:blue`) warns and renders as
  literal text. `harvey-scorecard` keeps its own 0–4 cell grammar.
- **Validation**: a `- Units:` count that does not match the columns and a
  `- Format:` rule that does not parse are `--check` errors; an unknown
  `- Sort:`/`- Format:` column or option value warns and the builder falls
  back (never crashes the build).

`comps-table` (banking comps / precedent transactions) is `table` with
`- Align: decimal`, `- Summary: mean, median` and `- Total-Row: off` preset
— the slide's own keys override the preset:

```markdown
## Slide 12: Precedent transactions support 8-9x EV/EBITDA
**Layout:** comps-table
- Units: , $M, x, x
- Sort: EV/EBITDA desc
| Target | Revenue | EV/EBITDA | EV/Sales |
|---|---|---|---|
| DealCo | 1,200 | 9.1 | 2.4 |
| MergeCo | 800 | 7.5 | 1.9 |
| NicheCo | 150 | NM | 3.2 |
```

### Combo bar+line, secondary axis, trendline, scatter groups & bubbles

**`chart:combo`** — clustered columns plus one or more line series in a
single native chart (think-cell's classic revenue-bars / margin-line
exhibit). Multi-series is required; by default the **last** `**Series:**`
name renders as the line, or name the line series explicitly:

```markdown
## Slide 4: Revenue compounded 18% while margin held at 23%
**Layout:** chart-callout
**Visual:** chart:combo
**Series:** Revenue $M, Costs $M, Margin %
**Data:**
- FY23: 100, 80, 20
- FY24: 120, 92, 23
- FY25: 140, 108, 23
- Line-Series: Margin %          ← comma list; default = last series
- Secondary: Margin %            ← right-hand value axis for these series
- Trend: linear "Fit"            ← native trendline on the first series
```

| Field | Behaviour |
|---|---|
| `- Line-Series: A, B` | Each name must be a `**Series:**` name (`--check` error otherwise); naming *every* series is an error — at least one stays a column (use `chart:line` for all-line). The lines keep the palette accent their series position already had, 2.5pt, markers off. |
| `- Secondary: A, B` | `chart:combo` and `chart:line` only (elsewhere: warn + ignore). Adds a native right-hand value axis (hidden category-axis pair, crosses at max) and rewires the named series' chart group to it — a group holding both named and unnamed series is split in two. Names must be `**Series:**` names (`--check` error). Secondary axis labels are muted like the primary. chart_lint's dual-axis honesty warn still fires — a secondary axis *is* two scales; keep it deliberate. |
| `- Trend: linear\|exp\|log\|power\|poly2..poly6 ["label"]` | Native `c:trendline` on the first plotted series — dashed accent3, 1.75pt, equation/R² off; the optional trailing quoted label becomes the trendline's name. Kinds: column/bar/line/scatter/combo (elsewhere: warn + skip). Unknown type: warn + linear. |

**Scatter groups** — a `[Group]` suffix on a data-row label buckets
single-Y `chart:scatter` / `chart:bubble` rows into **one series per
group** (legend on, palette accents; the suffix is stripped from the
plotted point):

```markdown
**Visual:** chart:scatter
**Series:** Growth %, Margin %
**Data:**
- Acme [Tier 1]: 12, 30
- Beta [Tier 1]: 8, 22
- Gamma [Tier 2]: 15, 18
```

Multi-Y rows (`**Series:** X, Y1, Y2 …`) already map columns to series —
a `[Group]` suffix there warns and is ignored. Suffix-less rows on a
grouped chart land in an "Other" series.

**`- Partition: x=50, y=30`** — on scatter/bubble: quadrant tint rects
(4 × `surface` at ~35% opacity, drawn *before* the chart so markers stay
on top) plus dashed lines at the given values. The axis ranges are set
explicitly (min 0, max from the data via the nice-ceiling rule) so the
line placement holds — same visual-QA caveat as CAGR/benchmark. Omit
either value to skip that line (a single value draws just its line, no
quadrant tints).

**`chart:bubble`** — native bubble chart: `**Series:** X, Y, Size` with
`- label: x, y, size` rows (**exactly 3 values** per row — `--check`
error otherwise). Transparent-styled like scatter, fills at ~60% opacity
so overlapping bubbles read; `[Group]` labels give one bubble series per
group. `- Size-Legend: on` draws two reference circles (min/max size,
`fmt_abs` labels) bottom-right inside the plot box.

`deck_to_outline.py` recovers all of the above: a combo comes back as
`chart:combo` + `- Line-Series:`, a bubble/grouped scatter as
`- Item N [group]: …` rows, so a rebuild regroups the same series.

### Agenda columns, sub-chapters & navigation

**Agenda columns** — `- Columns: time, page` on an `agenda` slide adds
right-aligned columns to each row (either or both; the listed order is the
display order). Time comes from a ` | ` suffix on the section bullet; page
is the printed page number of that section's divider slide, matched
case-insensitively against section-divider headings (unresolvable → blank
cell). Under `**Copy:** locked` the directive is ignored with a warning
(rows render byte-exact).

```markdown
## Slide 2: Agenda
**Layout:** agenda
- Columns: time, page
- Strategic context | 09:00–09:20
  - Market shifts
- Execution | 09:20–10:00
```

**Sub-chapters** — agenda bullets indented by 2+ spaces render as smaller
un-numbered muted rows under their chapter (no highlight logic of their
own, no page/time resolution).

**Appendix splitter** — with `**Auto-Agenda:** track`, one extra agenda
headed "Appendix" is inserted before the first appendix slide, listing the
appendix section-dividers (none present → skipped). `**Auto-Agenda:** on`
is unchanged.

**Empty chapters** — a section divider followed by zero content slides
(next slide is another divider, an appendix slide, or deck end) is dropped
from the auto-agenda rows with a warning, and no tracker agenda is
inserted after it. Hand-written agenda bullets are never dropped.

**Navigation** — after the build, agenda rows (auto and manual) and
tracker tab chips get `click_action` hyperlinks to their section divider
slide (matched by heading text, casefold; unmatched rows are skipped
silently).

**Live page numbers** — the footer page number is now a native
`slidenum` field carrying the same cached text it printed before, so
PowerPoint renumbers it on reorder while every lint keeps reading the
same string. Appendix slides keep the literal `B·` prefix run before the
field.

**Pagination** — `**Page-Start:** 3` (deck front-matter) sets the native
`firstSlideNum` and offsets the cached page texts. `- Hidden: on` marks
the built slide hidden (`show="0"` — skipped in the slide show and default
print); diff/QA tools still see it like appendix content.

### Footnotes, stamps & source normalisation

**Footnotes** — `- Footnote:` (repeatable) renders a numbered block just
above the source line (10pt muted, one entry per line, max 3 — overflow
warns and truncates). Numbering is 1..n in order; a leading `N `, `N.` or
`[N]` pins the number and the counter continues from it. Body markers
`[1]` render as superscript digit runs (`**Copy:** locked` leaves them
literal); the house-style orphan lint pairs markers with the footnote
list.

```markdown
- Momentum built through the year [1]
- Footnote: Excludes one-time items
- Footnote: [2] Constant currency
```

**Note** — `- Note: Preliminary figures` renders note-first on the source
line: `Note: … · Source: …` (a note with no source stands alone).

**Source normalisation** — `- Source:` values are piped through a
normaliser: any `Source:`/`Sources:` prefix variant is re-emitted as
`Source: `, entries split on `;` / ` | ` are trimmed, deduped
case-insensitively and rejoined with `; `. Under `**Copy:** locked` the
source bytes stay exact.

**File stamp** — `**File-Stamp:** on` (deck front-matter) centres
`<output filename> · <date>` in the footer band of every footered slide;
the date is `**Date:**` when set, else the build date.
