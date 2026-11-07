# QA Guide

## Philosophy

**Assume there are problems. Your job is to find them.**

Your first render is almost never correct. Treat QA as a bug hunt, not a confirmation step. If you find zero issues on first inspection, look harder.

---

## Step 0 — Programmatic Checks (run first, takes seconds)

```bash
python3 scripts/qa_check.py output.pptx
python3 scripts/diff_deck.py outline.md output.pptx   # content vs outline
```

`qa_check.py` catches mechanically-detectable defects: shapes outside slide bounds, text under 10pt, leftover placeholder text, opaque chart backgrounds, WCAG contrast failures (<3:1 error, <4.5:1 small-text warning), word-dense slides, empty slides, duplicate slide titles, and likely text overflow. Exit code 1 = errors must be fixed.

The overflow estimate is **font-true**: `textfit.py` measures with the palette's actual TTF (resolved via matplotlib → `fc-match` → the usual font directories) at PowerPoint's measured 1.2117 line pitch, falling back to Arial/DejaVu and then to an average-width approximation. `build_deck.py` installs the deck's body face before validating, so `--check` and the build agree on line counts — in template mode that is the template theme's *minor* (body) face, read from `theme1.xml` (`profile_template.extract_theme_fonts`) and mapped onto the palette's `font_title`/`font_body`/`font_label` (`palette_from_theme`), so the capacity gate measures with the face the deck really renders in. Overflow in the 100–140% band is written as an `<a:normAutofit>` shrink: bullet autofit measures at `textfit.LINE_PITCH` and snaps the `fontScale` through `textfit.snap_to_grid` onto PowerPoint's own 2.5% grid, floored at `builders.BULLET_FLOOR_PCT` (80), so the value PowerPoint stores is one it would have chosen itself; `textfit.fit_scale` shares that helper. Table cells get a literal point-size shrink (floor 10pt) instead, because PowerPoint ignores autofit inside cells.

**Editability warnings** (`editability.py`, always on) flag decks that look right but break when the client edits them: a text box parked on a same-bounds textless shape (the text belongs *in* that shape), a connector whose both endpoints float free of every other shape (a connector *inside* an exhibit group is exempt — it already moves with what it points at), and sub-12pt body text on a slide with plenty of empty space (the fixed source/footer band is chrome and is exempt). All three are warnings.

**Accessibility / projection mode** (mirrors the MS Accessibility Checker):
```bash
python3 scripts/qa_check.py output.pptx --accessibility
```

Beyond the stricter contrast (4.5:1) and body-font floor (18pt), accessibility mode adds:
- **Slide titles:** every slide must have a detectable title (error), and duplicate titles (case-insensitive) across the deck are reported once with all slide numbers — screen-reader navigation depends on unique titles.
- **Tables:** must be marked with a header row (`firstRow`) — error; merged cells get a warning ("screen readers may misread — prefer simple structure"). Decks built by this skill's table builder pass both.
- **Alt text:** missing alt text escalates to error, and alt text that is just a filename (`hero.png`, `image3`) is an error too — describe the image instead.
- **Reading order:** warns when the title shape is not the first text-bearing shape in document order (screen readers announce content in spTree order, not visual order).

**OOXML integrity** (optional schema validation):
```bash
python3 scripts/qa_check.py output.pptx --integrity
```

`--integrity` runs two things. First, **package hygiene** (`save_hygiene.py`, no dependencies, always runs): the Keynote/Quick-Look trio python-pptx leaves behind — a `<p:sldSz>` whose `type` says 4:3 on a 16:9 deck (and a `cx` 3048 EMU off 12192000), a `notesMaster` relationship with no `<p:notesMasterIdLst>`, a vestigial `ppt/printerSettings/printerSettings1.bin` — plus stock core properties (`lastModifiedBy=Steve Canny`, "generated using python-pptx", empty creator/title). Every build already sanitizes these at save, so a deck this toolchain produced comes back clean; the check exists so it cannot regress, and it is the one to run on a **foreign** deck before sending it to a Mac.

Second, with the optional `openxml-audit` package installed (`pip install openxml-audit`), it validates the .pptx OOXML structure and reports schema problems as errors prefixed `integrity:` (they fail the gate like any other error). Without the package it prints a single `[INFO] pip install openxml-audit …` line and is skipped — never a failure. The CI smoke test runs this as an advisory step.

For a proofreading pass (typos, wrong numbers, leftover draft phrasing):
```bash
python3 scripts/qa_check.py output.pptx --text
```

**Phantom and dropped content** — `diff_deck.py` runs both directions:

- *outline → deck*: every author-typed content row (bullets, card/stat/step fields, table cells) must survive into the deck. A dropped row is an **error** — silent truncation is the defect practitioners complain about most. Builder truncation caps (a `cards-3` slide takes 3 cards, a timeline 5 items) are mirrored in `diff_deck.LAYOUT_CONTENT`, so deliberately over-long input is not reported; the density-dependent caps and any profile-driven layout repick are read from the same `outline_replay.replay()` the build used (repeat a CLI-only `--profile` here — an outline `**Profile:**`/`**Density:**` is picked up on its own). Chrome (title/subtitle/heading) stays a warning: builders sometimes restyle or drop it.
- *deck → outline*: every deck paragraph must trace back to the outline, deck meta, the evidence ledger or known builder chrome. Anything else is **added content** — a hand edit or a hallucination — warned by default, an error under `--strict`. Deck slides beyond the end of the outline are scanned too.
- *chart cache vs embedded workbook*: a native chart carries both a cached copy of its numbers (what renders) and the `.xlsx` workbook behind it (what an analyst edits inside PowerPoint). When they disagree — someone typed a new number into the chart's datasheet — `diff_deck` warns `Slide N: embedded workbook drift — Series 'Revenue' @ 'FY24': cache 12.0 vs workbook 12.4`; `--strict` makes it an error. Reading the workbook needs the optional `openpyxl`; without it the check prints one `[INFO] pip install openpyxl for the embedded workbook drift check (skipped)` line on stderr — only for a deck that actually has a native chart, and never as a finding, so it is not counted in the QA record. `pptx_lint` deliberately does not duplicate this check.
- *locked copy*: when the deck was built with `**Copy:** locked` / `--verbatim`, the compare switches to **raw** — case, wrapping quotes and `**bold**`/`{accent}` markers are all significant — and any deviation is an error regardless of `--strict`. The mode is auto-detected from the deck's own slide tags (`PK_FLAGS` carries `verbatim`) first, then the outline's front matter; `--verbatim` / `--no-verbatim` force it either way. Chrome (title/subtitle/heading) stays on the normalized compare even under verbatim, because builders legitimately restyle it.

Caveat: a `--ghost` deck is placeholder exhibits by construction, so it reports mass dropped content. Diff the real build, not the skeleton.

Outline → deck blobs use MarkItDown when installed, else python-pptx; the deck → outline direction always uses python-pptx so paragraph boundaries survive.

House-style lints (marketing adjectives, hedges, passive titles, case and punctuation drift, orphan footnote markers — see `**Style-File:**` in `references/generation-guide.md`) run one step earlier, in `build_deck.py --check`. Fix them in the outline before building.

The binary VLM checklist for Step 2 is generated from the same outline — `python3 scripts/qa_checklist.py outline.md` — so it can be produced now and reviewed later.

**Slide identity (what makes the revision tools trustworthy).** Every slide `build_deck.py` writes carries a `ppt/tags/tagN.xml` part (`slide_tags.py`) recording `PK_SLIDE` (position), `PK_KEY` (a slug of the heading, made unique within the deck), `PK_LAYOUT`, `PK_TITLE_SHA`, `PK_DATA_SHA` (sha1 of the slide's `Data-File` bytes or inline `Data:` rows), `PK_EVIDENCE`, `PK_FLAGS` (`verbatim`, `repeat`) and `PK_VERSION`. Non-placeholder shapes are renamed `sNN-role[-k]` — `s04-title`, `s04-chart-1`, `s04-kicker` — so the PowerPoint selection pane reads as a map of the slide instead of "TextBox 17". Three consequences for QA:

- `patch_deck`, `refresh_deck`, `diff_deck` and the comment mapper locate a slide by identity, so a reviewer's hand reorder does not silently patch the wrong slide.
- Tags are **slide-level only** (no per-shape tags, none on masters/layouts) and they travel with `edit_deck duplicate/remove/extract/append`. `edit_deck duplicate` gives the clone its own tags part with `PK_KEY` re-keyed `<key>-copy` (`-copy2`, … until unique across the deck; every other `PK_*` value is copied as-is), so the source slide still resolves; two slides that genuinely share one key (a foreign tool cloning tags verbatim) are refused by `patch_deck` as ambiguous rather than guessed.
- `save_hygiene`, `client_safe` and `editability` treat `ppt/tags/*` and `p:custDataLst` as legitimate, so tags never show up as a hygiene or client-safety finding. A deck built before Cycle 6 (or a foreign deck) simply has no tags: the tools say so and fall back to position.

---

## Step 0b — Deck Lint (cross-slide consistency)

```bash
python3 scripts/pptx_lint.py output.pptx [--palette <palette>] [--client-safe [--denylist terms.txt]]
```

Complements `qa_check.py` with deck-wide checks that per-slide inspection misses. Findings are emitted in **fix order** — repair upstream classes first (structure → geometry → data honesty → color → type → composition), since fixing an upstream class often clears downstream findings:
- **Page-number sequence:** consecutive, no gaps or duplicates. The page number is a native `slidenum` field since Cycle 7; the check reads its cached text (what build_deck printed), so a `**Page-Start:**` offset passes — the delta invariant is unaffected by a constant offset — while tampered or duplicated cached numbers still flag.
- **Anti-jiggle:** recurring elements (page numbers, footers, kickers) must sit at identical coordinates on every slide.
- **Axis honesty:** same-type native charts with value-axis maxima >10x apart read as a dishonest scale comparison.
- **Chart honesty** (`chart_lint.py`, Misviz / Zelazny rules, one chart at a time). Errors: a bar/column value axis that does not start at zero, 3-D perspective (`c:view3D`), a category axis whose dates run out of order. Warnings: a truncated line/area axis with no axis-range note on the slide, dual value axes, a pie/doughnut that is not part-to-whole (bad sum, negatives, or >7 slices), non-uniform year/quarter/FY tick gaps, a missing legend with 2+ series or a value axis with no discoverable units, an indexed series with no "Indexed to 100" note, more than one per-point highlight colour (pies exempt), and a long line series that labels every point. A chart the parser cannot read is warned about, never a crash — and exactly **once per slide per run**: the message carries no chart identity, so `pptx_lint` collapses the copies its own axis / chart-honesty / numeric passes each produce for the same slide, and `qa_summary` drops `pptx_lint`'s copy when `qa_check` has already reported the same message, so the warning is never double-counted in the record's totals.
- **Numeric consistency** (`numeric_lint.py`, all warnings — a wrong number is the author's call): mixed decimal conventions within a table column, two negative conventions on one slide (parentheses *and* minus), two spellings of one unit (M vs MM, K vs k), a `Total` row that does not equal the rows as displayed, one stat-tile label carrying different values on different slides (compared in base units, so `$14.2M`, `14.2M` and `$14,200,000` are one number), and a number in an action title that the exhibit below never shows — where "shows" includes what a reader can derive from the chart: a total, a % change, a CAGR, a difference.
- **Font inventory:** flags decks with more distinct fonts than the threshold (inconsistent branding).
- **Palette whitelist:** with `--palette`, any run or fill color outside the palette (plus known extras) is reported as an error.
- **Typography:** with `--palette`, fonts outside the palette's title/body/label pair (plus the CJK stack) are errors — brand type enforcement, symmetric with the color whitelist — and font sizes off the palette's `type_scale` ladder are one aggregated warning (both the 16:9 ladder and its 4:3 canvas rescale are accepted). Without a palette, type-ladder sprawl (>14 distinct sizes) is warned. Custom palette JSONs may override `type_scale` with their own ladder.
- **Composition:** 3+ consecutive slides sharing one structural signature (same mix of charts/tables/pictures/textboxes) read as a template on repeat, and 4+ consecutive dense slides (>130 words each, no light slide between) break reading rhythm — both warned. Fix by varying layouts and inserting breathing room (big-number, quote, section divider).
- **Google Slides compatibility** (`--gslides`): when the deck will be imported into Google Slides, warns on fonts outside the Google Slides set (they get substituted), SmartArt (flattens to a static image), non-fade transitions (dropped or changed), and embedded audio/video (does not import). Run this only when the user mentions Google Slides — the default palettes intentionally use PowerPoint-native fonts.
- **Client safety** (`--client-safe`, `client_safe.py`): run this before any deck leaves the building. **Errors** — reviewer comments, hidden slides, and any hit from `--denylist terms.txt` (one term per line, `#` comments; searched across slides, notes, layouts and masters — past client names hide in slide masters). **Warnings** — identity left in core properties, off-slide (top-level) or hidden shapes, unused layouts and masters (aggregated into one line), embedded chart workbooks vs plain OLE objects, the stale package thumbnail, and the presence of speaker notes at all. Clean the findings with `edit_deck clean --client-safe` (see `references/editing.md`).

For custom-brand decks pass `--assets-dir` so the palette whitelist resolves: `python3 scripts/pptx_lint.py deck.pptx --palette <name> --assets-dir <assets>`. Note: slides using per-slide `{palette=...}` overrides will flag as off-palette against the deck palette — lint against the dominant palette and review those slides by eye.

Run this after `qa_check.py`; fix jiggle and off-palette colors in the outline before visual inspection.

---

## Step 1 — Render All Slides as Images

```bash
python3 scripts/render_slides.py output.pptx --grid --out assets/qa-thumbs/
```

This creates `slide-01.png`, `slide-02.png`, … plus a stitched `grid.png`.

If LibreOffice is not available the script falls back to a Pillow renderer — **the fallback dumps text only and cannot be used to judge layout, images, or charts**. Do not treat visual QA as complete without LibreOffice — rely on `qa_check.py` and opening the `.pptx` directly until LibreOffice is installed.

**Optional higher-fidelity renderers.** LibreOffice remains the default, but two alternates render closer to real PowerPoint when available: ONLYOFFICE's `x2t` converter (ships with ONLYOFFICE Document Server / Desktop Editors; converts .pptx → PDF with a rendering engine built for OOXML fidelity), and Microsoft Graph PDF export (`GET /drive/items/{id}/content?format=pdf` — uploads the deck to OneDrive/SharePoint and uses Microsoft's own rendering service; requires an M365 tenant and auth). Reach for either when LibreOffice output looks suspect — font metrics, chart label placement, effects — then `pdftoppm` the PDF into thumbnails as usual.

---

## Step 1.5 — Geometry Self-Check (instant — actually run it *before* Step 1's render)

```bash
python3 scripts/geometry_report.py output.pptx [--json] [--slides 2,5]
python3 scripts/geometry_report.py output.pptx --render-dir assets/qa-thumbs/
```

Deterministic per-slide layout metrics — no LibreOffice, no images, pure
geometry. It is cheap enough to run before spending a render cycle, and it
pinpoints exact shapes and offsets that thumbnails only show fuzzily,
feeding the fix loop precise targets.

**These classes are measured here. Do not ask the VLM about them in Step 2** —
deterministic geometry beats a vision model on all of them, and asking
anyway just adds confident noise:

- **Overlaps:** pairs of shapes that intersect by more than 0.02 sq in,
  excluding containment (a shape fully inside another is the intentional
  card/background pattern).
- **Escapes container:** content that mostly sits inside a solid card,
  band or picture but spills past its edge by more than 0.05in
  (`escapes container: 'Label' exceeds 'Card' by 0.12in (right)`). Text is
  judged on its glyph box. Two deliberate narrowings, worth knowing before
  you trust a silent result: a shape carrying **its own text** is never
  treated as a container, so on a foreign or template deck whose cards hold
  their copy directly rather than in a separate text box, this class stays
  quiet; and the child must fill at least 90% of the container **on the
  axis it spills along**, so a value label positioned past a bar's end or a
  caption over an image reads as placement, not overflow. Each edge is
  tested independently, so a tall narrow child can be reported as escaping
  the bottom even when its widest overflow is a sideways one.
- **Uneven spacing:** rows/columns of similar-sized shapes whose gap
  sequence is inconsistent ("uneven row spacing: gaps 0.31/0.29/0.55in").
  Runs are split at gaps over 1.5in — a large jump means separate visual
  groups, each judged on its own.
- **Near-misses:** shape edges that are *almost* aligned — off by
  0.04–0.08in. Smaller offsets are invisible at render size; larger ones
  are presumed intentional.
- **Whitespace ratio** (1 − covered/slide area, rasterized so overlaps
  aren't double-counted); below 0.20 the slide is reported as crowded.
- **Visual-mass balance** (AeSlides): the centroid of the covered area,
  normalized to −1..1 from the slide centre.
  `balance_norm = hypot(cx/0.05, cy/0.15)` — above 1 reports
  `visual mass imbalance: centroid off by (x,y) (norm N > 1)`. Three rules
  keep it measuring ink rather than frames: text rasters by its **glyph
  box**, not its text frame; fixed chrome (the title text shape, the footer
  zone, hairline divider rules) is excluded so the template cannot mask a
  lopsided body; and the finding is withheld — the numbers still print in
  `--json` — on slides a bounding-box centroid cannot judge, namely those
  where one chart or picture frame covers more than 25% of the canvas, and
  sparse ones with fewer than 4 composition shapes (title, closing,
  divider, big-number slides are centred by design). The known composition
  this catches is an exec-summary carrying 3 points in a 2×2 card grid —
  the empty fourth quadrant is a real imbalance, so fix it by rebalancing
  the grid rather than by ignoring the finding. The older half-vs-half
  `lr/tb_imbalance_pp` percentages remain in `--json` for continuity but no
  longer raise a finding.
- **Render whitespace map** (`--render-dir`, needs Pillow and an existing
  render): a 16×9 variance grid over each `slide-NN.png`; more than 55%
  ink-free cells reports `render whitespace: 62% empty cells, largest empty
  block 6x4 cells`. This is the mirror of *crowded* — a slide with too
  little on it — and it is the one balance signal that sees real ink rather
  than bounding boxes, so trust it over the pre-render centroid on
  chart-heavy slides. Title, closing and section-divider slides are
  deliberately sparse and will report here; that is expected, not a defect.
- **Word count** — more than 200 words is flagged as text overload.

The report deliberately stays quiet about patterns dense consulting layouts
produce on purpose, so the findings it *does* print deserve attention.
Overlap classification is z-order aware: value labels drawn over bars,
kicker bands layered over cards, and solid-on-solid stacking are suppressed,
while text-on-text collisions and text trapped under a solid shape are
reported — judged on the estimated glyph area (text insets + font sizes),
not the oversized text box, so breathing room inside a frame doesn't read as
a collision. Near-misses skip pairs co-aligned on the opposite edge of the
same axis (tornado/waterfall/football-field bars off a shared spine — the
differing edge is data, not drift) and pairs involving text frames (their
visible edge is the inset glyph edge, not the box edge); at most 4 near-miss
lines print per slide, closest first, with the rest summarized in one line.

Human output lists only slides with findings and ends with a summary line;
`--json` emits full metrics for every slide (useful even when nothing is
flagged — e.g. to compare whitespace ratios across sibling slides).

**Caveat on the balance metric.** Even with glyph rastering it works on
boxes, not pixels, so it declines to judge exhibit-dominated slides rather
than guessing at them — on those, read the `--render-dir` whitespace map
instead. Every other class above is exact.

---

## Step 2 — Visual Inspection (use a fresh-eyes subagent)

**Dispatch a subagent to inspect the thumbnails — even for a 3-slide deck.**
The agent that generated the slides systematically under-reports its own
defects; a fresh context judging only the rendered images catches what the
generator rationalizes away.

**Build the checklist from the outline, not from memory:**

```bash
python3 scripts/qa_checklist.py outline.md \
    [--evidence evidence.json] --output assets/qa-checklist.md
```

`qa_checklist.py` turns the outline into **binary, instance-specific** items
— one per bullet, stat, card and table row that must appear verbatim, one
per evidence-ledger value that must survive unchanged (with `--evidence`),
plus the generic flaws a render can show and geometry cannot. Instance
questions naming *this deck's* copy catch far more real defects than
open-ended "review this slide" prompts (PresentBench, SlideAudit). Slide
numbering mirrors the built deck (Data-File / Auto-Agenda / References
transforms are replayed), so `Q5.3` points at deck slide 5. `--json` emits
`{"slide", "id", "question", "kind"}` records for programmatic tracking.

**Protocol — give the subagent these rules verbatim:**

1. Work the checklist **one item per call**, in order. Do not batch, and do
   not summarize several items into one answer.
2. Answer **yes or no**, with slide-number evidence for every answer.
   **Partial = no.**
3. Alongside the generated items, judge these visual flaws per slide — the
   ones geometry cannot see:
   - **Weak title hierarchy** — is the title the most visually dominant text
     on the slide?
   - **Off-palette color** — does every element's color belong to the deck
     palette?
   - **Text overload / readability** — no paragraph walls, and body text
     large enough to read at projection size?
   - **Image quality** — nothing stretched, pixelated, washed-out, oddly
     cropped or missing (incl. dark-on-dark logos)?
   - **Sibling inconsistency** — do fonts, margins and title positions match
     the neighboring slides?
4. Do **not** answer questions about overlap, misalignment, uneven gaps,
   crowding or visual balance — Step 1.5 already measured those exactly.

**The gate is the binary items.** Any item answered no on a delivery slide
routes back to the fix loop (Step 3). Fix it in the outline, never in the
deck.

**Exemplars.** For "what good looks like", point the subagent at the
rendered exhibits in `references/gallery.md` — they are the reference
rendering for every built-in layout, so a slide that looks unlike its
gallery twin is worth a second look. For a per-exhibit zoom, re-render just
the slides in question at full size:

```bash
python3 scripts/render_slides.py output.pptx --slides 3,7 --out assets/qa-zoom/
```

**Deck-level rubric (one line, not a gate).** After the checklist pass, have
the subagent score the deck 1–5 on PPTEval's three dimensions — **Content**
(text quality, titles supported, images relevant), **Design** (color
harmony, hierarchy, layout craft) and **Coherence** (the storyline flows) —
and record it as a single summary line in the QA record. It is a trend
signal across revisions, not a pass/fail: the gate is the binary items.

**Re-review only what changed.** After a fix loop, re-render and run
`visual_regress.py` against the previous render; re-review only the slides
it flags as changed (plus any slide whose outline you edited). Slides whose
perceptual hash is unchanged were already reviewed.

---

## Step 3 — Fix & Re-render Loop

1. Fix all flagged issues in the **outline** (or edited slide XML for Mode D)
2. Re-run `python3 scripts/build_deck.py` (or `edit_deck.py pack` for Mode D)
3. Re-render affected slides only:
   ```bash
   python3 scripts/render_slides.py output.pptx --slides 3,5,7 --out assets/qa-thumbs/
   ```
4. Re-inspect until a full pass finds zero issues

**Do not deliver until at least one full fix-and-verify cycle is complete.**

---

## Step 4 — Final Gate

Re-run `python3 scripts/qa_check.py output.pptx` after the last fix — it must exit 0 (no errors) before delivery.

For a single machine-readable delivery record covering the whole deterministic toolchain:

```bash
python3 scripts/qa_summary.py output.pptx --outline outline.md \
    --palette <palette> --output qa-summary.json [--comments reviewed.pptx]
```

This runs `qa_check` + `pptx_lint` + `save_hygiene.check_package_hygiene` (always on, dependency-free — a Keynote-hostile slide size or leaked python-pptx core properties fails the gate here, not only under `qa_check --integrity`) + `geometry_report` + `diff_deck` and unifies their results with deck statistics (slides, native charts, tables, notes coverage, fonts, colors) into one JSON record with a top-level `pass` verdict. Exit 1 on any error — usable directly as a CI gate — and the record doubles as the delivery note ("6 slides, 3 native charts, 6/6 with speaker notes, all checks passed"). Rendering-based QA (Steps 1-3) is deliberately out of scope; the record captures everything checkable without LibreOffice.

`stats` also carries **`comments`** — `{"open": n, "resolved": n}` — counting the reviewer comment threads still in the deck (both the legacy and the modern PowerPoint formats, replies included). Any open comment on a deck about to be delivered is a delivery defect: either it is unaddressed feedback, or it is handled and should have been stripped (`edit_deck comments strip`, `references/editing.md`). `--client-safe` already treats leftover comments as an error; this stat is the number you quote in the delivery note. When a comment part in the deck cannot be parsed at all, the count comes back as `{"open": null, "resolved": null, "error": "…"}` and the record's additive `checks.comments` carries the same message as a warning — a damaged comment part costs you the count, never the whole QA record.

Add **`--profile <name>`** on a profiled deck: findings that recur across runs (same message signature seen in ≥ 2 runs) are appended to `<assets>/profiles/<name>.lessons.md`, dated and de-duplicated, with counts tracked in a sibling `.lessons.json`. The trailing `(seen N×)` on a lesson line is a live count, not a first-sighting stamp: a later run rewrites that number in place, and the run reports the refreshed line alongside any newly appended one. Refreshing matches the generated line exactly, so treat those lines as machine-owned — put your own note on a *new* line below a lesson rather than editing it, or the next run stops recognising it and appends a duplicate. Signatures ignore slide numbers and figures, so the same defect landing on different slides still counts as one recurring problem. That file is what the skill reads *before* building the next deck for that client — the QA loop's memory. An unknown profile name is an error, not a silent no-op.

The record also carries a top-level **`editability`** key — `{"level": "L0".."L5", "label": ..., "signals": {...}}` from `editability.py`, the SlidesGen-Bench PEI level for the deck (`schema` stays **1**: the key is purely additive). A generated deck reaches **L5**: each of the 18 shape-built exhibits (waterfall, mekko, gantt, the structural diagrams …) *and* each of the card-style core layouts (`exec-summary`, `cards-3`/`cards-4`, `stat-callout`, `timeline`, `comparison`, `dashboard`, `next-steps`, `agenda`, `exec-summary-scqa`) is wrapped in a real PowerPoint group by `shape_groups.py` — heading top-level, body (card frames, accent bars, markers, text) one `sNN-exhibit` group — so it moves and scales as one object while every bar, card and label stays individually editable; the example deck (`assets/example-outline.md`) reports L5. A deck drops to **L4** when a slide still carries `UNGROUPED_EXHIBIT_SHAPES` (6) or more loose textless autoshapes — a hand-edited or foreign deck, or an exhibit too small to be worth grouping (`closing` and `quote-evidence` draw fewer than six shapes and stay flat by design; they cannot trip the signal). `ungrouped_exhibits` is a signal that caps the level, never a warning.

**Comments for the reviewer** (`--comments out.pptx`) writes a copy of the deck with every error and warning pinned to its slide as a PowerPoint comment (author "QA") — the findings arrive where they apply instead of in a terminal log. These are *legacy* comment parts (`p:cmLst`), which every PowerPoint version opens; Keynote and Google Slides ignore them (they neither show nor destroy them). Ship the clean deck, not the commented copy.

---

## Revision regression (before sending a revised deck)

A revision should change exactly the slides you touched — anything else is collateral damage (reflowed text, a shifted chart, a renumbered footer). Keep the thumbnails of the last-delivered deck as a baseline and diff every revision against it:

```bash
python3 scripts/render_slides.py revised.pptx --out assets/qa-current/
python3 scripts/visual_regress.py assets/qa-baseline/ assets/qa-current/ [--threshold N]
```

- Same-named slide PNGs are compared by perceptual hash — pHash via the optional `imagehash` package when installed, else a built-in Pillow-only average hash. Hamming distance > 5 (tune with `--threshold`) flags the slide and reports a coarse pixel-diff percentage.
- **New slides** (present only in current) are reported as `new (no baseline)` — a warning, not a failure.
- **Deleted slides** (present only in baseline) are failures — removing content must be deliberate.
- Exit 1 on any changed or missing slide. Verify every flagged slide is an intended edit before sending.
- After delivering, bless the new render as the next baseline:
  ```bash
  python3 scripts/visual_regress.py assets/qa-baseline/ assets/qa-current/ --update
  ```
  `--update` also initializes a missing baseline on first delivery (and a run against a missing baseline without `--update` exits with instructions to do exactly that).

**On a data refresh, read the change log first.** `refresh_deck.py --log refresh-log.md` (see `references/editing.md`) lists every number that moved — `Revenue@FY25: 12 → 14 (+2, +16.7%)` — plus the values it could not pair, under **unpaired**. Two QA uses: check the moved numbers against what the client actually sent (a refresh that silently changed a number nobody mentioned is the defect), and treat every **unpaired** entry as a content change needing an outline edit, not a data refresh — a renamed series or an added category. The set of slides in the log should also be exactly the set `visual_regress.py` flags as changed; anything else moved for a reason you have not accounted for. Run `--dry-run` first when you want the log without touching the deck.

---

## Slide Count Per Delivery

Always deliver in one message:
1. The `.pptx` file path (attach or share via your environment's file delivery)
2. A thumbnail grid image showing all slides
3. A brief summary: "Here's your [N]-slide deck on [topic]. Covers: [list key slides]. Ready to edit in PowerPoint."

---

## Cross-slide consistency check (LLM stage)

Deterministic checks can't catch a title that contradicts its chart or an
exec-summary KPI that doesn't match the body. After the programmatic checks:

1. Run `python3 scripts/qa_check.py deck.pptx --numbers` and
   `python3 scripts/build_deck.py outline.md --titles`.
2. Review the two outputs together and check:
   - **Internal arithmetic** — components sum to stated totals (waterfall
     deltas vs endpoints, funnel stages, "3 levers" vs 3 items).
   - **Repeated KPIs** — the same metric quoted on multiple slides has the
     same value everywhere (exec summary vs body vs appendix).
   - **Title vs exhibit** — each action title's claim ("doubled", "-39%",
     "largest segment") is actually supported by that slide's numbers.
   - **Periods** — FY/quarter labels are consistent (no FY25 vs 2025 drift).
3. Fix discrepancies in the outline (never hand-edit the deck), rebuild,
   re-run. This stage is mandatory for consulting/board decks.

The deck-level item `QD.1` from `qa_checklist.py` belongs to this stage
rather than to any single slide: **do all exhibits share one visual style —
the same colors, fonts, label placement and gridline treatment?** Answer it
with the thumbnail grid open, comparing exhibit slides side by side; a
single off-style chart is the most common cross-slide defect a per-slide
pass misses.
