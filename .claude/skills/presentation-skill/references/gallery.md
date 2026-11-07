# Layout Gallery

One entry per layout in `LAYOUT_MAP` (`scripts/builders.py`), grouped by
purpose. Each entry gives a decision-relevant "use when" / "don't use when"
and a minimal valid outline snippet — copy-paste it into an outline and swap
in real content. Syntax details (all fields, defaults, validation rules)
live in `references/generation-guide.md`; this page is for picking the right
layout fast, not for the full field reference. `## Slide N:` in every
snippet means "any slide position" — renumber sequentially in your outline.

---

## Core narrative

### title
**Use when:** Opening slide of the deck — sets the title and subtitle before any content.
**Don't use when:** Marking a mid-deck chapter break — use `section-divider` instead.

```
## Slide N: Accelerating AI Infrastructure
**Layout:** title
- Title: "Accelerating AI Infrastructure"
- Subtitle: "Q4 Strategy Review · December 2026"
```

### section-divider
**Use when:** Marking a new chapter so the narrative visibly resets between sections.
**Don't use when:** The deck has only one or two sections — a single `agenda` slide covers it without extra dividers.

```
## Slide N: Market Opportunity
**Layout:** section-divider
- Subtitle: "Sizing the addressable market and growth drivers"
```

### agenda
**Use when:** Orienting the audience on deck structure, optionally highlighting the current section via `- Current:`.
**Don't use when:** You want that same nav repeated after every section divider — set `**Auto-Agenda:** track` in front-matter instead of hand-authoring repeats.

```
## Slide N: Agenda
**Layout:** agenda
- Market opportunity
- Product roadmap
- Financial plan
- Current: "Product roadmap"
```

### exec-summary
**Use when:** Opening a deck or section with 2-4 headline takeaways the audience must remember.
**Don't use when:** The takeaways follow a Situation → Findings → Recommendation argument — use `exec-summary-scqa`.

```
## Slide N: Three priorities for H1 2027
**Layout:** exec-summary
- Headline: "Three strategic priorities for H1 2027"
- Point 1: "Deploy GPU cluster Phase 2 by March"
- Point 2: "Migrate 80% of workloads to hybrid cloud"
- Point 3: "Launch developer platform beta"
```

### bullet-list
**Use when:** A straightforward list of points with no natural chart, card, or timeline shape.
**Don't use when:** The points map onto a stronger visual (`comparison`, `cards-3`, `process-flow`) — plain bullets read as filler when a shape is available.

**Alias:** `bullets` renders identically — either name works.

```
## Slide N: Three forces are reshaping the market
**Layout:** bullet-list
- Regulatory tailwinds accelerate enterprise adoption
- Compute costs fall 40% year over year
- Open-source models close the capability gap
```

### two-column-split
**Use when:** Pairing a short narrative with a supporting chart or image side-by-side.
**Don't use when:** The chart deserves the full slide with a called-out insight — use `chart-callout`.

```
## Slide N: Enterprise AI spend more than doubles to $104B by 2026
**Layout:** two-column-split
**Visual:** chart:bar
**Series:** Market Size ($B)
**Data:**
- 2024: 42
- 2025: 67
- 2026: 104
- Heading: "Enterprise AI spend reached $104B in 2026"
```

### cards-3
**Use when:** Introducing exactly three parallel items — team, pillars, workstreams.
**Don't use when:** The item count is 4 or varies — a mismatched count triggers a "cards (expected 3)" warning; use `cards-4` or `bullet-list`.

```
## Slide N: Leadership team
**Layout:** cards-3
- Card 1: Name="Alex Kim" Title="CEO" Bio="Ex-Google Brain"
- Card 2: Name="Jordan Lee" Title="CTO" Bio="Built infra at Scale AI"
- Card 3: Name="Sam Patel" Title="Head of Product" Bio="Former PM at Stripe"
```

### cards-4
**Use when:** Introducing exactly four parallel items — capabilities, features, pillars.
**Don't use when:** The item count is 3 or varies — a mismatched count triggers a "cards (expected 4)" warning; use `cards-3` or `bullet-list`.

```
## Slide N: Platform capabilities
**Layout:** cards-4
- Card 1: Title="Ingest" Body="Streaming and batch pipelines"
- Card 2: Title="Train" Body="Distributed GPU training"
- Card 3: Title="Serve" Body="Sub-50ms inference"
- Card 4: Title="Govern" Body="Model risk and audit trails"
```

### comparison
**Use when:** Contrasting two states side-by-side — before/after, us/them, legacy/new.
**Don't use when:** Comparing 3+ options against multiple criteria — use `harvey-scorecard`.

```
## Slide N: Legacy stack vs. new architecture
**Layout:** comparison
- Left label: "Legacy Stack (2023)"
- Right label: "New Architecture (2026)"
- Left: "Monolith on-prem"
- Right: "Active-active multi-region"
```

### full-image
**Use when:** A single hero image should dominate the slide — screenshot, photo, diagram.
**Don't use when:** No supporting image exists — an image-less `full-image` renders a plain background (warned by `--check`); use `bullet-list` or `stat-callout` instead.

```
## Slide N: Live production dashboard
**Layout:** full-image
- Caption: "Live dashboard — 47ms p99 latency"
```

### big-number
**Use when:** One number IS the message and deserves the whole slide.
**Don't use when:** You have 3+ KPIs — use `dashboard`.

```
## Slide N: The prize
**Layout:** big-number
- Value: "$4.2B"
- Label: "Annual value at stake"
- Context: "Bottom-up across 11 validated use cases"
```

### quote-evidence
**Use when:** Pairing a customer or stakeholder quote with one supporting proof stat.
**Don't use when:** You need 2-3 standalone stats without a quote — use `stat-callout`.

```
## Slide N: Customers trust the platform because it just works
**Layout:** quote-evidence
- Quote: "We cut incident response time by 70% in the first quarter."
- Attribution: "VP Engineering, Acme Corp"
- Value="70%" Label="Faster incident response"
```

### closing
**Use when:** Final thank-you / call-to-action slide with contact details.
**Don't use when:** It's a mid-deck chapter break, not the deck's end — use `section-divider`.

```
## Slide N: Questions & Discussion
**Layout:** closing
- Title: "Questions & Discussion"
- Subtitle: "Appendix available upon request"
- Contact: "alex.kim@company.com"
```

---

## Data exhibits

### table
**Use when:** Structured data compares cleanly in rows and columns with no need for heat-coding.
**Don't use when:** The numbers should read at a glance by intensity — use `heatmap-table`.

```
## Slide N: The Scale plan costs 5x more for 10x the seats
**Layout:** table
| Plan | Price | Seats |
|------|-------|-------|
| Starter | $99 | 5 |
| Scale | $499 | 50 |
```

### comps-table
**Use when:** Banking comps or precedent transactions — decimal-aligned numeric columns with Mean/Median summary rows and NM cells.
**Don't use when:** There is a total row to emphasise or no numeric columns — use `table`.

```
## Slide N: Precedent transactions support 8-9x EV/EBITDA
**Layout:** comps-table
- Units: , $M, x
- Sort: EV/EBITDA desc
| Target | Revenue | EV/EBITDA |
|---|---|---|
| DealCo | 1,200 | 9.1 |
| MergeCo | 800 | 7.5 |
| NicheCo | 150 | NM |
```

### heatmap-table
**Use when:** A table's numeric columns should read at a glance by intensity — benchmarks, scorecards.
**Don't use when:** Values within a column aren't comparable (mixed units) — use a plain `table`.

```
## Slide N: EMEA unit costs run 3x the NA baseline
**Layout:** heatmap-table
| Region | Cost | Churn % |
|---|---|---|
| NA | $10 | 5% |
| EMEA | $30 | 7% |
```

### stat-callout
**Use when:** Highlighting 2-3 standalone KPIs with labels and sublabels.
**Don't use when:** You have a single hero number — use `big-number`.

```
## Slide N: Platform performance at a glance
**Layout:** stat-callout
- Value="$14.2M" Label="Cost Savings" Sublabel="vs. legacy stack"
- Value="99.97%" Label="Uptime SLA" Sublabel="Jan–Jun 2026"
```

### timeline
**Use when:** Chronological milestones or a roadmap anchored to dates.
**Don't use when:** Steps aren't date-anchored — use `process-flow`.

```
## Slide N: Three milestones complete the H2 2026 roadmap
**Layout:** timeline
- Date="Q3 2026" Title="Cloud migration" Desc="Cut data center costs 30%"
- Date="Q4 2026" Title="Phase 2 GPUs" Desc="256 H100s operational"
- Date="Q1 2027" Title="Platform GA" Desc="Billing + RBAC"
```

### chart-callout
**Use when:** A chart needs an adjacent "key insight" box calling out the so-what.
**Don't use when:** The chart needs to sit beside prose bullets instead of a callout box — use `two-column-split`.

```
## Slide N: Spend crosses the ceiling in Q2
**Layout:** chart-callout
**Visual:** chart:line
**Data:**
- 2024: 38
- 2025: 52
- 2026: 71
- Callout: "Crossover hits in Q2 2026 — decision needed this quarter"
```

### dashboard
**Use when:** 3+ KPI tiles with trend deltas need to be scanned together.
**Don't use when:** You only have 1-2 numbers — use `stat-callout` or `big-number`.

```
## Slide N: Platform health at a glance
**Layout:** dashboard
- Tile 1: Value="99.97%" Label="Uptime" Delta="+0.02pt" Trend="up"
- Tile 2: Value="$14.2M" Label="Savings" Delta="+18%" Trend="up"
- Tile 3: Value="312" Label="Design partners" Delta="-4%" Trend="down"
```

### waterfall
**Use when:** Explaining how a value moved from A to B — bridges, cost walks, EBITDA drivers.
**Don't use when:** Categories aren't additive — use `chart:bar` instead.

```
## Slide N: Margin fell 4pts, driven by input costs
**Layout:** waterfall
**Data:**
- FY24 margin: 18
- Volume: 2
- Input costs: -5
- Pricing: -1
- FY25 margin: total
```

### funnel
**Use when:** Showing drop-off through sequential pipeline or conversion stages.
**Don't use when:** Stages aren't a strict narrowing sequence — use `process-flow`.

```
## Slide N: Pipeline conversion drops sharply after demo
**Layout:** funnel
**Data:**
- Leads: 5000
- Qualified: 1200
- Demo: 400
- Closed: 90
```

### mekko
**Use when:** Market share across segments where both width (segment size) and height (share) carry meaning.
**Don't use when:** You only need share, not relative size — use a stacked `chart:bar`.

```
## Slide N: Two platforms capture 73% of segment value
**Layout:** mekko
**Series:** Platform A, Platform B, Other
**Data:**
- Compute: 18, 22, 6
- Storage: 9, 11, 4
```

### bar-mekko
**Use when:** Profit-pool style view — bar width = segment size, height = a second metric.
**Don't use when:** You don't have two independent metrics per bar — use `chart:bar`.

```
## Slide N: EMEA drives 60% of profit on 30% of revenue
**Layout:** bar-mekko
- Bar: Label="EMEA" Size="30" Value="6"
- Bar: Label="NA" Size="45" Value="9"
- Bar: Label="APAC" Size="25" Value="3"
```

### tornado
**Use when:** Ranking sensitivity drivers by their swing on an outcome.
**Don't use when:** You're presenting a valuation range, not a sensitivity spread — use `football-field`.

```
## Slide N: WACC swings drive most of the 35-point valuation range
**Layout:** tornado
**Series:** Downside, Upside
**Data:**
- WACC +/-1pt: -12, +18
- Volume growth: -8, +10
```

### football-field
**Use when:** Showing valuation (or estimate) ranges from multiple methods on one shared axis.
**Don't use when:** You're ranking sensitivity drivers, not ranges — use `tornado`.

```
## Slide N: Three valuation methods converge on the $45-52 band
**Layout:** football-field
**Series:** Low, High
**Data:**
- DCF: 42, 58
- Trading comps: 45, 52
```

---

## Consulting frameworks

### exec-summary-scqa
**Use when:** A consulting-style exec summary needs a Situation → Findings → Recommendation argument.
**Don't use when:** The takeaways don't need that argument structure — use plain `exec-summary`.

```
## Slide N: We recommend consolidating onto two strategic cloud platforms
**Layout:** exec-summary-scqa
- Situation: "Cloud spend doubled to $46M while delivery velocity stalled"
- Finding 1: "61% of spend sits on workloads with <30% utilization"
- Recommendation: "Consolidate to two platforms; tier workloads — $18M savings by FY27"
```

### matrix-2x2
**Use when:** Prioritizing items across two dimensions — effort vs. value, impact vs. feasibility.
**Don't use when:** You're plotting stakeholders by support/influence — use `stakeholder-map` (same engine, purpose-built labels).

```
## Slide N: Workload tiering is the fastest path to value
**Layout:** matrix-2x2
- X-axis: "Implementation effort →"
- Y-axis: "Value at stake →"
- Item: Name="Workload tiering" X="0.2" Y="0.85"
- Item: Name="Platform exit" X="0.85" Y="0.7"
```

### harvey-scorecard
**Use when:** Comparing vendors or options against multiple criteria with at-a-glance 0-4 scores.
**Don't use when:** Only two options are being compared — use `comparison`.

```
## Slide N: Vendor B leads on the criteria that matter most
**Layout:** harvey-scorecard
| Criterion | Vendor A | Vendor B |
|---|---|---|
| Scalability | 2 | 4 |
| Ecosystem | 3 | 3 |
```

### process-flow
**Use when:** A single owner runs a short, sequential operating-model flow (chevrons, max 6 steps).
**Don't use when:** Steps split across multiple functions or lanes — use `swimlane-process`.

```
## Slide N: Three steps take us from pilot to scale
**Layout:** process-flow
- Step 1: Title="Pilot" Desc="Validate with 3 business units"
- Step 2: Title="Scale" Desc="Roll out to all regions"
- Step 3: Title="Sustain" Desc="Embed in BAU governance"
```

### next-steps
**Use when:** Closing with concrete actions, owners, and deadlines that stay visible through Q&A.
**Don't use when:** The "steps" are really historical milestones — use `timeline`.

```
## Slide N: Decision required: approve Phase 1 by month end
**Layout:** next-steps
- Step 1: Action="Approve Phase 1 budget ($2.4M)" Owner="CFO" When="This week"
- Step 2: Action="Stand up tiger team" Owner="CTO" When="Within 2 weeks"
```

### driver-tree
**Use when:** Decomposing a KPI into the drivers that build up to it, with deltas.
**Don't use when:** You're modeling a governance hierarchy of people, not numbers — use `org-chart`.

```
## Slide N: Pricing drives 60% of the planned revenue uplift
**Layout:** driver-tree
- Node: Id="rev" Label="Revenue" Value="$120M" Delta="+8%" Parent=""
- Node: Id="price" Label="Price" Value="$70M" Delta="+12%" Parent="rev"
- Node: Id="vol" Label="Volume" Value="$50M" Delta="-2%" Parent="rev"
```

### stakeholder-map
**Use when:** Plotting stakeholders by support and influence, with optional move targets.
**Don't use when:** Your two axes are effort/value for prioritizing initiatives, not people — use `matrix-2x2`.

```
## Slide N: Move the CFO from skeptic to sponsor before gate review
**Layout:** stakeholder-map
- Item: Name="CFO" X="0.2" Y="0.9" TargetX="0.8" TargetY="0.9"
- Item: Name="COO" X="0.7" Y="0.6"
```

### raci
**Use when:** Assigning R/A/C/I responsibility per activity across people or roles.
**Don't use when:** You just need static reporting lines, not per-activity accountability — use `org-chart`.

```
## Slide N: Every workstream has one accountable owner from day one
**Layout:** raci
| Activity | Alice | Bob | Cara |
|---|---|---|---|
| Design | A | R | C |
| Build | I | A | R |
```

### gantt
**Use when:** Phased roadmaps with parallel workstreams and milestones across time periods.
**Don't use when:** It's a simple linear timeline of dated events, not parallel rows — use `timeline`.

```
## Slide N: Three workstreams deliver over six quarters
**Layout:** gantt
**Periods:** Q3 26, Q4 26, Q1 27, Q2 27
- Bar: Row="Infrastructure" Label="Tiering" Start="1" End="2"
- Bar: Row="Platform" Label="Pilot" Start="1" End="1"
- Milestone: Row="Platform" Label="GA" At="3"
```

Cycle 7 extras (all optional; see the generation guide's "Newer directives —
Cycle 7" section): `**Periods:** quarters FY26-FY27` / `months Jan 2027 -
Jun 2027` expand with a spanning year-header row; Bar `Responsible=".."` /
`Remark=".."` add right-hand label columns; `- Span: Rows="A, B"
Label=".."` draws a left row bracket; `- Shade: Periods="3-4" Label=".."`
tints period columns behind the bars; `- Line: At="2.5" Label=".."` draws a
dashed milestone line; `Row="Parent > Child"` nests rows under a bold group
header.

---

## Structural diagrams

### architecture-layers
**Use when:** Layered systems where vertical position means dependency — TOMs, tech stacks, control frameworks.
**Don't use when:** Elements are sequential, not stacked — use `value-chain` or `process-flow`.

```
## Slide N: Target Operating Model
**Layout:** architecture-layers
- Layer: Name="Experience" Items="Portal, Mobile"
- Layer: Name="Services" Items="APIs, Copilots" Accent="2"
- Layer: Name="Data" Items="Lakehouse, Governance"
```

### swimlane-process
**Use when:** Cross-functional processes where handoffs between actors are the point — intake, approval flows.
**Don't use when:** A single actor owns the whole flow — use `process-flow`.

```
## Slide N: Use-case intake runs through three functions
**Layout:** swimlane-process
- Lane: Name="Business" Steps="Submit idea > Refine value case"
- Lane: Name="AI CoE" Steps="Screen > Score feasibility > Prioritize"
- Lane: Name="Risk" Steps="Classify risk tier > Approve controls"
```

### org-chart
**Use when:** Governance models, reporting lines, org structures — three-lines-of-defense.
**Don't use when:** You're decomposing a metric's numeric drivers, not people or forums — use `driver-tree`.

```
## Slide N: AI Governance Model
**Layout:** org-chart
- Node: Id="steer" Label="AI Steering Committee" Role="Direction & funding"
- Node: Id="coe" Label="AI CoE" Parent="steer" Role="Delivery"
- Node: Id="risk" Label="Model Risk" Parent="steer" Role="2nd line"
```

### value-chain
**Use when:** Value chains or lifecycle stages where each stage has sub-activities worth naming.
**Don't use when:** Stages have no sub-activities to show — use `process-flow`.

```
## Slide N: AI Value Chain
**Layout:** value-chain
- Step 1: Title="Data" Sub="Acquisition, Quality, Lineage"
- Step 2: Title="Models" Sub="Build, Evaluate, Register"
- Step 3: Title="Deployment" Sub="Serve, Monitor, Retrain"
```

## Process grids & banking exhibits

### what-how-who
**Use when:** Each step of a process needs its what / how / who spelled out — delivery models, mobilization plans.
**Don't use when:** Steps only need a title and description — use `process-flow`.

```
## Slide N: Delivery model spans three steps
**Layout:** what-how-who
- Step 1: Title="Diagnose" What="Baseline the costs" How="Bottom-up model" Who="Deal team"
- Step 2: Title="Design" What="Target operating model" How="Workshops" Who="CoE"
- Step 3: Title="Deliver" What="Migration waves" How="Agile pods" Who="IT + vendor"
```

### tombstone-grid
**Use when:** Deal credentials or transaction experience — the classic banking tombstone page.
**Don't use when:** You're comparing metrics across deals, not showcasing them — use `table` or `league-table`.

```
## Slide N: Selected transaction experience
**Layout:** tombstone-grid
- Tombstone: Client="Acme Corp" Deal="Sale to Beta Holdings" Value="$1.2B" Date="Mar 2026" Role="Exclusive advisor"
- Tombstone: Client="Gamma Inc" Deal="IPO on NYSE" Value="$800M" Date="Jan 2026"
```

### league-table
**Use when:** Ranked competitive standings where relative size should read at a glance — underwriting league tables, market-share ranks.
**Don't use when:** The data isn't ranked or has no dominant numeric column — use `table`.

```
## Slide N: We rank second on announced volume
**Layout:** league-table
| Rank | Firm | Volume $B | Share |
|------|------|-----------|-------|
| 1 | First Bank | 42.5 | 18.2% |
| 2 | Our Firm | 38.1 | 16.3% |
```

### logo-wall
**Use when:** A client, partner or membership grid where the logos are the message.
**Don't use when:** You need to say something about each entry — use `cards-3` / `cards-4`.

```
## Slide N: Clients we serve
**Layout:** logo-wall
- Logo: logos/acme.png "Acme Corp"
- Logo: "Beta Holdings"
```
