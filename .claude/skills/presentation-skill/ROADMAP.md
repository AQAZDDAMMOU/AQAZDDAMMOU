# Enhancement Roadmap

Cycles 5–8 (research sweep of 22 August 2026; Cycles 5 and 6 delivered, their
open items closed 23 August 2026, 7–8 planned). Sources: academic work
2025–26 (PresentBench, SlidesGen-Bench, AeSlides, SlideAudit, PPTArena,
DeepPresenter/PPTAgent v2, Misviz, ArcDeck), consulting/banking tooling (think-cell
incl. Assist, UpSlide AI Consistency Check, Macabacus Deck Check, Mekko Graphics,
Efficient Elements, Slidewise, MBB/IB convention guides), the agent ecosystem
(anthropics/skills issues, Claude for PowerPoint, Copilot 2026, Gemini in Slides,
NotebookLM, Manus, Genspark, Gamma, Pitch Agent, MemSlides, DECKBench,
power-pptx, ts-pptx, pptx-glimpse), practitioner threads (r/consulting, r/powerpoint,
r/ClaudeAI, HN), and python-pptx 1.0.x issues. Headline findings: (1) commercial
tools converged in 2026 on *review/consistency* features rather than generation —
our deterministic, CI-gateable QA is the right moat, but numeric consistency is
still LLM-only; (2) benchmarks locate the code-generation gap in layout polish
(overflow, occlusion, balance), and deterministic geometry metrics beat VLMs at
judging it — keep geometry out of the VLM checklist; (3) practitioners' top
complaints are trust and "AI fingerprint" (phantom/dropped content, marketing-speak,
Frankenstein formatting, template non-compliance, no memory across runs) and the
revision loop; (4) four delivery-hygiene defects were verified in our own example
build (Keynote-incompatible OOXML, leaked core properties). Re-validated bets:
native editable shapes (SlidesGen-Bench PEI L4/L5 vs commercial L0–L2; Manus moved
to native .pptx), fresh-context render-and-critique (DeepPresenter), binary per-flaw
checklists (SlideAudit, PresentBench). Demoted: the PPTEval-style 1–5 rubric
(weakest human alignment) — keep as a summary number only.

Cycle 4 (August 2026). Sources: online research sweep across commercial tools
(Gamma, Beautiful.ai, Tome, Presenton), competing agent skills (Akxan/ppt-agent-skill,
siril9/presentation-skill, zairuilab/consulting-deck), academic work (PPTAgent/PPTEval,
AutoPresent/SlidesBench CVPR 2025), and practitioner discussion (r/powerpoint).
Headline finding: the field's biggest unmet need is cleanup of existing decks
("make it not ugly" tax — brand-template compliance, formatting unification),
plus evidence traceability; generation itself is comparatively well served.
AutoPresent/SlidesBench empirically re-validates two standing bets: program
generation beats image generation, and iterative render-and-refine improves quality.

Cycle 3 (August 2026). Sources: defects observed while building and QA-ing a real
9-slide exec deck (items 1–7 — each was empirically hit in a single build), plus the
larger bets deferred from the June 2026 research cycle (items 8–12). All June-cycle
tasks (T2–T14) and both implementation plan documents are complete and have been
removed; see git history for their specs.

Validated bets (no change needed): native pptx over image-gen, markdown outline as
source of truth, layout DSL, render-based QA, checklist-driven VLM critique.

Status: `[ ]` open, `[x]` done. Effort: S < 1 day, M 1–3 days, L > 3 days.

---

## Cycles 1–6 — open items round (23 August 2026)

Every "still open", "known issue" and actionable scope-note gap left by
Cycles 3–6 was actioned on `fix/cycles-1-6-open-items` (suite 1271 → 1330
passed / 4 skipped, ruff 55 → 0, `smoke_test.py` green, `package_skill.py`
134 files). What shipped:

- **`build_deck.py` is an orchestrator** (1039 → ~350 lines): new leaf
  modules `outline_parse.py` (parser, `FIELD_KEYS`/`META_KEYS`),
  `outline_data.py` (`_parse_data_point`, `_series_count`,
  `SLIDE_HEADING_PREFIX`, `- Data-File:` loading), `outline_transforms.py`
  (auto-agenda, references, scale groups, diagram visuals, CJK),
  `deck_sections.py` (p14 sections) and `template_build.py` (placeholder
  slides); every old name is re-exported from `build_deck`. No module imports
  `build_deck` back any more (`outline_validate`, `verbatim`, `house_style`,
  `slide_tags`, `outline_replay`, `outline_expand` all read the leaf
  modules), so `build_deck.py` run as `__main__` is imported exactly once.
- **`replay()` is profile- and density-aware.** `profiles.apply`, the
  CLI/variant/front-matter density resolution (`builders.set_density`) and
  the `preferred_layouts` repick moved from `build()` into
  `outline_replay.replay()` (`ReplayResult` gained `palette_key`/`density`),
  so `diff_deck`, `qa_checklist`, `qa_summary`, `patch_deck`, `refresh_deck`
  and `gen_handout` see the layouts and truncation caps the build used — a
  `**Density:** comfortable` exec-summary no longer reads as dropped rows in
  `diff_deck`. `diff_deck.py`/`qa_checklist.py` gained `--profile`;
  `qa_summary --profile` and `patch_deck`/`refresh_deck --profile` pass it
  through.
- **Editability L5 for generated decks, including card layouts.** The
  card-style core layouts (`exec-summary`, `cards-3/4`, `stat-callout`,
  `timeline`, `comparison`, `dashboard`, `next-steps`, `agenda`,
  `exec-summary-scqa`) hand their body to `shape_groups.group_shapes()`; the
  `next-steps` step number is now the circle's own text
  (`helpers.set_shape_text`) instead of a text box parked on it, which had
  capped any deck carrying that layout at L3. The example deck reports
  **L5**; `tests/test_editability.py` asserts it, and the golden
  flat-vs-grouped fixture covers the card layouts too.
- **`chart:scatter` builds** (`XyChartData`; `charts.scatter_xy`): numeric
  labels (else the row index) as x for single-value rows, `**Series:** X,
  Y1[, Y2 …]` first-column-is-X for multi-value rows; markers only, legend on
  2+ series. `--check` errors on x,y rows under a single series name (raw
  rows kept in `_data_raw`, thousands separators stay one number) and warns on
  the index fallback; `deck_to_outline` recovers scatter, `chart_workbook`
  reads `c:xVal`/`c:yVal`, `chart_lint` no longer calls a scatter dual-axis,
  `refresh_deck` pairs scatter points by x on both sides.
- **Template mode measures with the theme's body face**:
  `profile_template.extract_theme_fonts` + `palette_from_theme(colors,
  fonts)` run before `validate()`; the theme override no longer leaks into
  the shared palette dict between builds.
- **Profile `logo` is consumed** (`slide_chrome.resolve_logo`/`add_logo`:
  corner badge on title/closing, footer-band mark elsewhere, role `logo`,
  missing file warns).
- **`edit_deck duplicate` re-keys the clone** (own tags part, `<key>-copy[N]`)
  so `patch_deck` no longer refuses; **`refresh_deck`** numbers repeated stat
  captions / table row labels positionally on both sides.
- The 55 baseline ruff findings were cleared.

Left as documented behaviour (by design, not defects): `multiples=` has no
builder caller (the parser drops the `x` suffix); `--comments` writes legacy
`p:cmLst` comment parts; `pptx_comments.add_comments` writes no reply chains;
`strip --resolved` cannot touch legacy comments; refresh never inserts slides;
tags are slide-level only; `{accent}`/`{row}`/`{rows}` are reserved in
`**Repeat:**`; chrome is compared normalised under verbatim.

## Cycle 7 — think-cell parity: chart annotations, tables, agenda (August 2026)

Theme: close the remaining exhibit gaps against the consulting gold standard.
Delivered 23 August 2026 on `feat/cycle-7-think-cell-parity` (spec:
`docs/superpowers/specs/2026-08-23-cycle-7-think-cell-parity-design.md`; seven
parallel worktree agents; suite 1330 -> 1605). New seams: `chart_visual.py`
(the one chart-feature hook point, with a fixed apply-order contract) and
`cell_shapes.py` (shared multi-state token vocabulary: `harvey:3/8`,
`rag:green`, `bar:64%`, `dots:3/5`, `[x]` — table cells and bullet markers).
`table`, `gantt` and `agenda` builders moved to `builders_table.py` /
`builders_gantt.py` / `builders_agenda.py` (old names re-exported). Scope
notes: hbar difference arrows, decimal *tab* alignment, print-skip distinct
from hidden, agenda time-slot arithmetic and template-mode versions of the
new chrome are deliberate non-goals (spec "Non-goals").  Pre-existing (not
a Cycle 7 regression, surfaced by its QA pass): a `**Tracker:** tabs` deck
caps at editability L3 — the chip text box sits on the chip at same
bounds, tripping the text-over-shape signal.

- [x] **1. Difference arrows (S–M).** `- Diff-Arrow: FY24, FY26 [abs|pct|both]`
  and `Series@Cat → Series@Cat` on column/bar/line/stacked charts — level and
  total differences, one- or two-way, shape-drawn with auto bubble label when
  short. think-cell's most-used decoration after CAGR; we have CAGR and the
  waterfall-only `- Bracket:`.
- [x] **2. Chart option set (S each).** `- Value-Line: mean|median|wmean "label"`
  computed from data; `- Sort: desc|asc` + `- Ordinal: on` on bar/column; 100%
  indicator tick on stacked-100/mekko; net lines / `- Target: series` ticks; error
  bars / min–max whiskers; per-series CAGR labels and multi-span `- CAGR:
  FY22-FY24, FY24-FY27` with "p.a."; category gaps / breaks (Actual | Forecast);
  display units ($M), label rotation/skip, gap width.
- [x] **3. Combo bar+line, secondary axis, trendline (M).** `chart:combo`,
  `- Secondary: series`, `- Trend: linear|log|power` — the most-requested
  python-pptx chart gaps (#338/#607), lxml on `plotArea`/`ser`; PptxGenJS's option
  list is the checklist. Scatter/bubble: groups, partition lines with quadrant
  fills, bubble-size legend.
- [x] **4. Data rows under charts (M).** `- Data-Rows: Growth %, Margin` prints
  value / CAGR / growth rows aligned to columns beneath the x-axis (think-cell 14
  "charts with data tables", Mekko Graphics data rows).
- [x] **5. Agenda + navigation (S).** Time-slot and page-number columns, sub-
  chapters, appendix "splitter" agenda, auto-hide empty chapters; internal
  hyperlinks agenda→section and clickable tracker chips
  (`click_action.target_slide`, native); live `<a:fld type="slidenum">` field
  instead of static page-number text so numbers survive hand inserts; pagination
  controls (start numbering at slide N, print-skip flags).
- [x] **6. Footnotes, stamps, source normaliser (S–M).** Numbered footnote
  markers `[1]` in body text rendering an auto-numbered footnote block above the
  source line; file-path/date stamp; source-line normaliser ("Source: A; B; Firm
  analysis", `Note:` before `Source:`, de-duplicated repeats).
- [x] **7. Table upgrades (M).** Numeric right/decimal alignment, column units
  row, zebra striping, total-row styling, per-cell harvey / checkbox / traffic-
  light / percent-bar, threshold conditional formats, auto mean/median/quartile
  summary rows, NM handling, column sort.
- [x] **8. Gantt parity (M).** Responsible/Remark label columns, row-spanning
  brackets, shaded periods, milestone lines, hierarchical (indented) activities,
  fiscal-year calendar.
- [x] **9. New layouts (S–L).** What/How/Who process grid (pentagon header row +
  striped matrix); multi-state bullet/cell shapes (traffic light, percent bar,
  rating dots, N-segment harvey); banking set — comps/precedents table with NM
  and auto statistics, tombstone grid, league table / logo wall.

## Cycle 8 — Storyline, delivery aids & template fidelity (planned)

Theme: the before (storyline) and after (delivery) of the build, plus the two
larger fidelity bets.

- [ ] **1. Storyline aids (S).** Dot-dash `storyline.md` format (dots = action
  titles / sections, dashes = supporting facts + intended exhibit) consumed by
  `build_deck --ghost` and checked by `narrative.py` for horizontal logic;
  `titles.md` read-through export for partner sign-off; cross-slide contradiction
  check (same KPI, different values — shares Cycle 5 item 2); ArcDeck-style
  commitments doc (thesis, spine, slide budget) + narrative-critic pass;
  inverse-spec check ("reconstruct the brief from the titles alone").
- [ ] **2. Delivery aids (S each).** `**Duration:** 20m` → per-slide time budget
  from notes word count (~130 wpm) with lint on overrun / missing notes;
  Prepare-for-Questions scaffold (anticipated questions per slide, unanswered
  objections, weakest-evidence slides from the ledger, `## Appendix` backup-slide
  stubs); verification trail in notes ("Verified: E1 (IDC 2026 p.12) …; unsourced:
  0") + `- Sticker: VERIFIED` + counts in `qa_summary.json`; `gen_handout
  --recap-email | --talking-points | --one-paragraph`; `condense_outline.py
  --slides 6` exec-version mode (detail slides → `## Appendix`, storyline titles
  kept, moved evidence listed). Evidence: Copilot Prepare-for-Questions, Genspark
  verification trail, Pitch Agent collateral, "6-slide proposal turning into a
  22-slide endurance test".
- [ ] **3. Template mode v2 (M–L).** Semantic shape inventory of non-placeholder
  template shapes (title / subtitle / body / card_N header / source) by heuristic,
  fill in place preserving run formatting, `locked:` shapes (logos, legal) never
  touched, post-build template-compliance report (off-template fonts/sizes/colours
  per slide), "Frankenstein template" warning (multiple masters, orphan layouts)
  with `--clean`. Evidence: template non-compliance is Copilot's #1 enterprise
  complaint (Octigen), anthropics/skills #632/#633.
- [ ] **4. Theme-linked build mode (M–L).** Write the palette into `theme1.xml`
  and emit `schemeClr` + `+mj-lt`/`+mn-lt` references so a client template
  recolours and refonts the deck in PowerPoint; lint resolves scheme refs; sRGB
  stays the default because WCAG checks only hold for our own theme.
- [ ] **5. Translation / localization mode (M).** `**Lang:** de` sets font stack
  and number/date/currency locale; per-language text-expansion factors in textfit
  (DE/FR/ES +20–35%, CJK −30%); lint on mixed currencies/units; documented
  translate path for delivered decks via `edit_deck inventory → translate →
  replace`; RTL declared unsupported. Evidence: Copilot Agent-mode translate,
  Plus AI translate-all, Kimi USD/CNY mixing failure.
- [ ] **6. Layout alternatives sheet (M).** `build_deck --variants 3 --slide K`
  renders one outline slide in three candidate layouts onto a single comparison
  image ("show me 3 ways to visualize this for an exec").
- [ ] **7. QuizBank fidelity gate (M).** For source-grounded decks: generate ~10
  MCQs from the source document, answer them from deck text only; report
  retention in `qa_summary` (SlidesGen-Bench: Spearman 0.71 with humans vs 0.57
  for LLM-judge). Fills source→deck retention (diff_deck is outline→deck only).
- [ ] **8. Render robustness (S).** Retry `soffice` on lock-file / profile
  failures with a fresh `-env:UserInstallation` per call; qa-guide states "never
  drive the user's PowerPoint via COM; skip visual QA if LibreOffice is absent"
  (anthropics/skills #886/#1523).

## Cycle 6 — Revision loop, data binding & memory (August 2026)

Theme: "generation is one hour; revision is twenty" — make the second-to-Nth
iteration cheap, and give the skill memory across runs. Delivered 22 August 2026
on `feat/cycle-6-revision-loop` (31 commits off `main` at `5d877f5`, tip
`1b31487`; suite 809 → 1144 tests, 3 skipped; spec at
`docs/superpowers/specs/2026-08-22-cycle-6-revision-loop-design.md`, plan at
`docs/superpowers/plans/2026-08-22-cycle-6-revision-loop.md`). New modules:
`outline_replay`, `outline_expand`, `outline_ops`, `slide_tags`, `verbatim`,
`profiles`, `chart_workbook`, `refresh_deck`, `revision_plan` (+
`revision_grammar`, `revision_apply`). New directives: `**Repeat:**`,
`**Copy:** locked`, `**Profile:**`. New CLIs: `refresh_deck.py`,
`outline_ops.py`, `edit_deck comments list|plan|apply|strip`. The follow-ups
listed below were actioned on 23 August 2026 on `fix/cycle-6-follow-ups`, which
also split `outline_validate.py` + `slide_chrome.py` out of `build_deck.py` and
`geometry_balance.py` out of `geometry_report.py`.

- [x] **1. Comment-driven revision.** `pptx_comments.py` reads both comment
  generations — legacy `p:cmLst/p:cm` (authors from `commentAuthors.xml`) and
  modern `p188:cm` (replies, `status="resolved"`, shape/position anchors) —
  mapping parts to slides through `p:sldIdLst` so gap-numbered and reordered
  decks are read correctly, and strips by `--slides` / `--resolved` / `--all`.
  `revision_plan.py` (+ `revision_grammar.py` parsers, `revision_apply.py`
  outline editor) turns those comments plus an emailed `feedback.md`
  (`Slide N:`, `S N`, `p. N`, `#N`, `"Title":`) into `revision-plan.json`/`.md`,
  resolving each item to an outline slide (by `PK_KEY`) and field (exact →
  casefold → difflib ≥ 0.8, containment-scored). The command grammar
  (`delete bullet N`, `delete this`, `replace "X" with "Y"` / `X -> Y`,
  `retitle: T`, `add bullet: T`, `move to slide N`, `kill slide`, `split slide`,
  `merge with slide N`) auto-applies to the outline; everything else is a
  `manual` item for Claude. `apply` is idempotent (a delete verifies the text it
  recorded, then skips rather than deleting the wrong line), skips resolved
  items, and stops slide-level ops once a move/merge has renumbered the outline
  ("re-plan"). `edit_deck comments list|plan|apply|strip` is the CLI;
  `qa_summary` stats gain `comments.open/resolved`.
- [x] **2. Batch patch + outline split/merge ops.** `patch_deck.py` takes
  `4`, `3,7,9` or `2-4` (one rebuild, N extracts, one append, one reorder) and
  matches each selected outline slide to the delivered deck by its `PK_KEY`
  tag (fallback `PK_TITLE_SHA`), so hand reorders and inserted slides no longer
  break the swap; ambiguous duplicate keys are refused rather than
  first-matched, orphan tags parts are GC'd, and `--by-position` (or an
  untagged deck) keeps the old position+count contract. `outline_ops.py`
  `split N [--at K] [--heading]` / `merge N,M [--heading] [--layout]` /
  `move N K` / `renumber` rewrites the markdown with every non-target byte
  preserved (`renumber` on a clean outline is a byte-level no-op), refuses to
  split single-exhibit layouts and `**Data:**` slides, and runs `--check` on the
  result (exit 1, file still written).
- [x] **3. Data-refresh mode.** `refresh_deck.py outline.md delivered.pptx
  --output out.pptx [--all] [--dry-run] [--log refresh-log.md]` replays the
  outline, recomputes each slide's `PK_DATA_SHA`, and rebuilds only the movers
  through `patch_deck.patch` — hand edits elsewhere survive byte-for-byte. The
  change log pairs every number it can between the delivered exhibit (chart
  cache, table cells, stat tiles) and the refreshed outline
  (`Revenue@FY25: 12 → 14 (+2, +16.7%)`), comparing in base units so `M`/`K`
  spellings stay honest, and lists what it cannot pair under "unpaired" instead
  of guessing. Exit 2 on an untagged deck; an outline that matched nothing is an
  error, not a silent success.
- [x] **4. Row-templated slides.** `**Repeat:** data/initiatives.csv
  [rows=a-b] [sort=col]` on one slide (`outline_expand.py`, hooked into
  `replay` right after the parse, so every consumer sees the expanded list)
  clones it per data row with `{column}` substitution into every string the
  slide carries, plus `{row}`/`{rows}` builtins and `{{`/`}}` literals. Unknown
  fields error with the file's actual columns; `> 60` slides warns, `> 200`
  errors; `--check` prints `[info] Slide N expands to K slides (file)`. Tag key
  `<slug>-rN`, `PK_FLAGS` flag `repeat`.
- [x] **5. Embedded-workbook round-trip.** `chart_workbook.py`
  (`cache_series`/`embedded_workbook`/`read_workbook_series`/`drift`/
  `deck_drift`) reads the `.xlsx` behind a native chart via python-pptx's
  `ChartWorkbook` + optional openpyxl. `deck_to_outline` prefers workbook values
  over the XML cache (fidelity note only when the two disagreed or the workbook
  was unreadable); `diff_deck.check_workbook_drift` warns
  `Slide N: embedded workbook drift — Series 'X' @ 'Cat': cache 12.0 vs
  workbook 12.4` (error under `--strict`), aligned by category label so a
  mid-sheet insert does not misreport. `pptx_lint` deliberately does not
  duplicate it.
- [x] **6. Shape tags carrying outline ids.** `slide_tags.py` stamps every built
  slide (styled, ghost and template mode) with `ppt/tags/tagN.xml` —
  `PK_SLIDE, PK_KEY, PK_LAYOUT, PK_TITLE_SHA, PK_DATA_SHA, PK_EVIDENCE,
  PK_FLAGS, PK_VERSION` — linked via `p:cSld/p:custDataLst/p:tags`, and renames
  non-placeholder shapes `sNN-role[-k]` (title/text/shape/chart/table/picture/
  group/connector plus the chrome roles footer, pagenum, tracker, kicker,
  sticker, source, stamp, exhibit-badge) so the selection pane reads as a map of
  the slide. Tags travel with `edit_deck duplicate/remove/extract/append` and
  are ignored by `deck_to_outline`; `save_hygiene`, `client_safe` and
  `editability` treat them as legitimate.
- [x] **7. Verbatim-copy mode.** `**Copy:** locked` (front matter only) or
  `--verbatim` (`verbatim.py` + the process-global `helpers.set_verbatim`):
  copy fields keep their wrapping quotes (`COPY_FIELDS`; directive fields still
  unquote), `**bold**`/`{accent}…{/}` render literally, and any list cap that
  would drop a row is a **build error** (`LAYOUT_CONTENT` caps plus
  `SUPPLEMENTAL_CAPS` for matrix-2x2/stakeholder-map items and dashboard tiles)
  rather than a silent trim. `diff_deck` compares raw and fails on any deviation
  regardless of `--strict`, auto-detecting the mode from the deck's `PK_FLAGS`
  or the outline and forceable with `--verbatim`/`--no-verbatim`.
- [x] **8. Engagement/client profile memory.** `profiles.py` +
  `<assets>/profiles/<name>.json` (`palette, footer, stamp, density, size,
  author, number_format, style_file, logo, banned_phrases, mandatory_slides,
  preferred_layouts, notes_style, page_numbers, tracker`; unknown keys and wrong
  types are errors) applied via `**Profile:** acme` / `--profile acme`, with
  precedence CLI flag > variant preset (palette/density only) > outline
  front-matter > profile > default. Missing mandatory
  slides are `--check` errors, banned phrases merge into house style,
  `notes_style` lints the notes, and `preferred_layouts` re-picks only
  auto-chosen layouts within the ambiguity pairs, capacity-guarded so it can
  never truncate content. `brand_kit.py <domain> --name n --profile` writes a
  profile; `qa_summary --profile n` appends findings seen in ≥ 2 runs to
  `<assets>/profiles/n.lessons.md` (+ `.lessons.json`), and SKILL.md instructs
  Claude to read that file before building for the profile.

**Cycle 6 scope notes** (where delivery narrowed the spec):

- Item 1: auto-apply is limited to the command grammar above — every other
  comment is a `manual` item, by design. `strip --resolved` can only strip
  *modern* comments; the legacy format carries no resolved state, so those
  survive (use `--slides`/`--all`). Reply-thread folding is covered end-to-end by
  a modern-comment fixture grafted into a built deck
  (`pptx_comments.add_comments` itself still writes only legacy roots, so no
  *builder* produces a real reply chain).
- Item 2: `outline_ops.move` counts the `## Appendix` marker as a position slot.
  A CRLF outline now survives `split`/`merge`/`move` and the revision-apply text
  ops — every synthesized line takes the file's own terminator (`doc_eol`).
- Item 3: refresh **never inserts slides** — an outline slide with no tagged
  match is reported and skipped. Adding a slide stays `patch_deck`/`edit_deck`
  work. ~~Stat-tile labels share one flat namespace, so the same caption twice
  on a slide collapses into one entry~~ — fixed in the open-items round below
  (repeats are numbered `Label (2)` positionally on both sides; table row
  labels too).
- Item 4: `{accent}` is reserved (it is rich-text markup), so a data column
  named `accent` is unreachable; `{row}`/`{rows}` likewise shadow same-named
  columns. Layout is auto-picked once from the un-substituted template slide,
  before expansion. A zero-row selection drops the slide with a warning and still
  checks its `{field}` references against the data file's header row.
- Item 6: tags are **slide-level only** — no per-shape tags parts, no sidecar
  manifest, nothing on masters or layouts. Headline-less layouts
  (`quote-evidence`) legitimately have no `title` shape, and a lone two-character
  in-band title ("Q3") is excluded from title candidacy with the decorative
  glyphs. ~~`edit_deck duplicate` produces duplicate `PK_KEY`s; `patch_deck`
  refuses that deck rather than guessing~~ — fixed in the open-items round
  below (the clone gets its own tags part and a `<key>-copy[N]` key).
- Item 7: chrome (title/subtitle/section labels) is still compared normalized
  under verbatim, because builders legitimately restyle it — a chrome miss stays
  a warning. `textfit` measures markup-stripped text under verbatim (cosmetic).
  `**Copy:** locked (anything else)` and an in-body `**Copy:**` line are both
  *not* verbatim, and each warns.
- Item 8: ~~the `logo` key is recorded but **inert**~~ — consumed since the
  open-items round below (title/closing corner badge + footer-band mark).
  ~~`diff_deck` and `qa_checklist` replay the outline profile-unaware~~ —
  `replay()` now applies the profile, the density and the layout repick
  itself, so every consumer sees the build's layouts and caps (a CLI-only
  `--profile` is repeated on `diff_deck`/`qa_checklist`/`qa_summary`). The
  `lessons.md`
  "seen N×" count is refreshed in place on every later run, but only on a line
  still in its generated shape — annotate a lesson on a *new* line below it, not
  by editing the generated one, or the next run appends a duplicate.

**Cycle 6 follow-ups** — actioned on 23 August 2026 on
`fix/cycle-6-follow-ups` (suite 1173 → 1210 passed / 3 skipped, ruff 55,
`smoke_test.py` green). What shipped: `build_deck.py` (1736) split into
`outline_validate.py` (579) + `slide_chrome.py` (178), leaving 1039 with every
moved name re-exported under its old spelling; `geometry_report.py` (714 → 600)
split into `geometry_balance.py` (142); `pptx_comments.Comment` hashes by
`thread_id`; one shared `profiles.safe_name` (`brand_kit`'s duplicate `NAME_RX`
deleted, and its `main()` now uses the *stripped* return value); the `lessons.md`
"seen N×" count is refreshed in place; the delivered deck is opened once in
`patch_deck`/`refresh_deck`/`diff_deck`; a heading `replace`/`retitle` syncs the
same slide's `- Heading:`/`- Title:`; reply-thread folding has an end-to-end
test; `qa_summary.comment_counts` degrades a malformed comment part into
`stats.comments.error` plus an additive `checks["comments"]` warning (schema
still 1) and `strip_comments` reports it as a named `ValueError`;
`outline_ops.split`/`merge`/`move` build every synthesized line with the
outline's own terminator; `apply` records **comment-identity fingerprints** in
`<plan-stem>.applied.json`, so `replace "X" with "X."` fires exactly once *and* a
round-2 re-plan is not skipped by round 1's ledger (an older id-keyed ledger is
ignored with a notice, a different-deck one warned about); `patch_deck`'s
verbatim guard is symmetric (a locked rebuild against a plain deck is refused —
"rebuild the whole deck once") and names an unreadable outline instead of
blaming verbatim; a zero-row `**Repeat:**` warns and drops the slide while still
checking its `{field}` references; `chart_workbook` narrowed its bare excepts.
~~**Still open from that round:** `outline_validate.validate()` reaches back
into `build_deck` for `_series_count` through a function-local import … and
`build_deck.py` is still 1039 lines~~ — both closed in the open-items round
below: `build_deck.py` is a ~350-line orchestrator, the parser, data rows,
transforms, sections and template path are leaf modules, and no module
imports `build_deck` back (it is imported exactly once under `__main__`).

## Cycle 5 — Delivery hygiene & deterministic trust lints (August 2026)

Theme: kill the "AI fingerprint" and the trust failures with cheap, deterministic,
CI-gateable checks; fix the verified output defects first. Delivered 22 August 2026
on `feat/cycle-5-delivery-hygiene` (merged to main as `037cd90`; suite 477 → 809
tests; plan at `docs/superpowers/plans/2026-08-22-cycle-5-delivery-hygiene.md`).
New modules: `save_hygiene`, `numfmt`, `numeric_lint`, `pptx_comments`,
`chart_lint`, `client_safe`, `content_trace`, `house_style`, `editability`,
`qa_checklist`. New directives: `**Author:**`, `**Number-Format:**`,
`**Style-File:**`.

- [x] **1. Keynote/macOS compatibility + core-properties scrub.** The example
  build shipped `<p:sldSz … type="screen4x3"/>` on a 16:9 deck, no
  `<p:notesMasterIdLst>` despite a notesMaster rel, a leftover
  `ppt/printerSettings/printerSettings1.bin` (the trio that makes Keynote /
  Quick Look refuse python-pptx files — anthropics/skills #1167/#1231) and
  `lastModifiedBy="Steve Canny"` with empty author/title. `save_hygiene.py`
  sanitizes at save (`screen16x9` + cx 12,192,000, id list injected, printer
  part + content-type default dropped, author/title from `**Author:**` / the
  title slide; real identities on template decks are preserved) and doubles as
  a CLI for foreign decks; `qa_check --integrity` and the `qa_summary` gate now
  run `check_package_hygiene` so none of the four can regress.
- [x] **2. Number-format policy + numeric-consistency lint.**
  `**Number-Format:** units=M decimals=1 negatives=parens|minus multiples=1x pct=1`
  (`numfmt.py`) drives every builder-drawn value label; omitted, output is
  byte-for-byte unchanged. Largest-remainder rounding now covers absolute values
  (waterfall deltas vs total, stacked segments vs stack label, stacked-100 abs
  labels) — think-cell round parity, at the policy's precision. `numeric_lint`
  (in `pptx_lint`, all warnings): mixed decimals per table column / negative
  styles / unit spellings per slide; `Total` row ≠ sum; same-label KPI drift
  across slides (unit-normalised); title number ≠ exhibit number (derived set:
  sum, % change, CAGR, differences). `qa_summary --comments out.pptx` emits every
  finding as a PowerPoint comment. Evidence: UpSlide AI Consistency Check,
  Macabacus Deck Check, r/consulting QA checklist. The lint's first catch was our
  own example outline (a "32% YoY" title over 55% CAGR data — fixed).
- [x] **3. Chart-honesty lint v2.** `chart_lint.py` ports the Misviz/Zelazny
  rules onto our chart XML: bar value-axis not at zero, 3-D, and non-monotonic
  dates are errors; truncated line axis without an axis-range note, dual axis,
  inconsistent tick intervals (year/quarter/FY categories only), non-part-to-whole
  or >7-slice pies, missing legend (≥2 series) / axis units, indexed charts
  without "Indexed to 100", >1 highlight colour, too many labelled points are
  warnings. Descending years (newest-at-top hbars) are accepted; unreadable
  chart parts degrade to a warning.
- [x] **4. Client-safety scrub + clean-up command.** `pptx_lint --client-safe
  [--denylist FILE]` (`client_safe.py`): leftover comments, hidden slides and
  denylist hits in slide/notes/master/layout/chart/core-props text are errors;
  identity props, off-slide and hidden shapes, unused masters/layouts, embedded
  workbooks vs OLE objects, stale `thumbnail.jpeg`, speaker notes present are
  warnings. `edit_deck clean --client-safe [--drop-notes] [--drop-hidden]`
  removes them (keeps `dc:title`); `edit_deck sanitize-numbers` randomises
  figures (slides + chart caches + chart text; years and 1–2-digit numbers kept)
  for sharing a deck as a template. Evidence: r/consulting pre-send checklist,
  think-cell Clean-up, MS Document Inspector.
- [x] **5. Phantom / dropped-content check.** `diff_deck` gains the
  deck→outline direction (`content_trace.py`): built text not traceable to the
  outline, data files, the evidence ledger or builder boilerplate is "added
  content" (warn; error with `--strict`), deck slides beyond the outline are
  scanned, and dropped outline content rows are **errors** — with every
  builder's truncation cap mirrored in `diff_deck.LAYOUT_CONTENT` so honest decks
  stay clean. Evidence: r/powerpoint "adds phantom things … randomly takes things
  off", Every.to "80% right is worse than no AI".
- [x] **6. House-style + action-title lints.** `house_style.py` (re-exported via
  `narrative.py`) reads `style.json` (`**Style-File:**` / `--style` /
  `<assets>/style.json` / built-in defaults): "not X but Y", marketing
  adjectives, hedges, judgemental phrasing, passive action titles, title length
  (`max_title_words`, now owned here), duplicate/near-duplicate titles,
  punctuation and sentence/title-case drift (`title_case: any` escape hatch),
  orphan footnote markers, double spaces / missing space after period — all
  surfaced in `--check`. Evidence: ex-MBB partner thread, slideworks action-title
  rules, Efficient Elements 25-rule checker.
- [x] **7. Editability lint + editability level.** `qa_check` warns on text
  boxes floating over a same-bounds shape, free-floating connectors (incl.
  axis-aligned ones) and "tiny fonts + dead space" (sub-12pt body text on a
  mostly empty slide); `qa_summary` reports a SlidesGen-Bench PEI editability
  level L0–L5 with its signals. Evidence: r/ClaudeAI "text doesn't move with the
  shape", SlidesGen-Bench.
- [x] **8. Font-true overflow estimator shared by textfit and qa_check.**
  `textfit` resolves the palette's actual TTFs (+bold) via matplotlib → fc-match
  → font dirs (Arial/DejaVu fallback), measures at PowerPoint's 1.2117 line
  pitch, sets the measure font before `validate()`, offers `fit_scale` (2.5%
  grid, 3.5% margin) and `shrink_to_fit_pt`; `qa_check` dropped its 0.52-em
  heuristic for the same estimator; table cells get literal font shrink (floor
  10pt) because PowerPoint ignores autofit in cells. Evidence: ts-pptx
  measured-text-fit, SlideAudit, grapeot/pptx.skill `check-overflow`.
- [x] **9. Deterministic geometry metrics + VLM checklist protocol.**
  `geometry_report` adds AeSlides' centroid ellipse-normalised balance (x_tol
  0.05 / y_tol 0.15, glyph-box raster, skipped on frame-dominant and sparse
  slides), the "escapes container" collision class and a `--render-dir`
  local-variance whitespace map; `qa_checklist.py` generates PresentBench/
  SlideAudit-style binary, instance-specific items from outline + ledger
  (completeness, fidelity, placeholder/title-only, title–content alignment,
  leaked notes, language, ≤6 bullets, chart-annotation completeness,
  cross-exhibit consistency; *partial = no*, one item per call); the qa-guide
  removes geometry asks from the VLM checklist, adds phash-gated re-review, and
  demotes the 1–5 rubric to a summary number.

**Cycle 5 scope notes** (where delivery narrowed the spec):

- Item 2: `multiples=` and `pct=` were implemented (`numfmt.fmt_multiple` /
  `fmt_pct`) with no builder callers, and native chart data labels used
  PowerPoint's own number format. Both fixed in the follow-up round below —
  `pct=` now reaches every builder-drawn percentage and the policy is written
  onto native chart data labels as a format code; `multiples=` is still
  caller-less by design (the outline parser drops an `x` suffix when it reads a
  data point), now documented as such. `--comments` writes the **legacy** `p:cmLst`
  comment parts (every PowerPoint opens them; Keynote / Google Slides ignore
  them); not opened in real PowerPoint during the cycle.
- Item 7: our shape-built exhibits were intentionally flat, so
  `ungrouped_exhibits` capped a generated deck at **L4**. Grouping shipped in the
  follow-up round below (`shape_groups.py`), so **L5 is reachable**; a deck only
  drops to L4 now when a slide still carries 6+ loose textless autoshapes.
- Item 8: `fit_scale` had no caller — bullet autofit floored a whole-percent
  scale at 80 with a 1.2 literal. Done in the follow-up round below: bullet
  autofit snaps on `textfit.snap_to_grid`'s 2.5% grid (floor
  `builders.BULLET_FLOOR_PCT` 80) and measures at `textfit.LINE_PITCH`.
  ~~Template-mode theme fonts are still not measured~~ — since the
  open-items round below the template theme's minor face is the measure font
  before `validate()`.
- Item 9: exemplar thumbnails / zoom crops are documented via
  `references/gallery.md` + `render_slides --slides N` — no new cropper. The
  balance finding on an exec-summary carrying 3 points in a 2×2 grid was a
  **true positive** (empty fourth quadrant), not calibration noise — fixed in the
  follow-up round below (3 points render as one column of three full-width
  cards), and the example deck now has no balance offenders at all.
- Behaviour changes on existing outlines (documented in qa-guide/README): dense
  table cells now shrink to 10–13.5pt instead of growing the row; label rounding
  is policy-consistent; `diff_deck` dropped rows fail `qa_summary`; `qa_check
  --integrity` / `qa_summary` error on foreign decks with an empty `dc:title`
  (run `save_hygiene.py` first).

**Cycle 5 follow-ups** — **all actioned** on 23 August 2026 on
`fix/cycle-5-follow-ups` (suite 1213 → 1271 passed / 3 skipped, ruff 55
unchanged, `smoke_test.py` green, `package_skill.py` 128 files). What shipped:

- **Group shape-built exhibits (L5).** New `shape_groups.py` wraps every
  shape-built exhibit that drew 6+ shapes (18 layouts across
  `builders_exhibits`/`builders_consulting`/`builders_structure`) in a `p:grpSp`
  named `sNN-exhibit` with an *identity* child transform, so child coordinates
  stay canvas coordinates; every consumer walker was made group-aware
  (`canvas_shapes` for bounds, `iter_shapes` otherwise), a golden test proves
  grouped and flat builds produce identical QA/diff/geometry output, and the two
  LibreOffice renders are pixel-identical.
- **`exec-summary` with exactly 3 points** renders one column of three
  full-width cards instead of a 2×2 grid with an empty quadrant — the pinned
  example-deck balance true positive is gone.
- **Bullet autofit on the real grid.** `builders._autofit_overflow` measures at
  `textfit.LINE_PITCH` and snaps through the new `textfit.snap_to_grid`
  (PowerPoint's 2.5% `fontScale` grid, floor `BULLET_FLOOR_PCT` 80); `fit_scale`
  now shares that helper.
- **Number policy reaches the last labels.** `pct=` routes through `fmt_pct` for
  stacked-100/mekko shares, the waterfall bracket, the funnel `↳` conversion and
  the CAGR arrow (shares rounded at `pct_decimals()` so a column still sums to
  100), and `numfmt.chart_number_format(policy)` is applied to native chart data
  labels (`c:numFmt`, `sourceLinked=0`, labels stay hidden). `multiples=` still
  has no builder caller and is documented as such. With no policy declared the
  output is byte-identical except for two disclosed divergences on those
  percentages: half-up `.5` ties (`12.5% -> 13%`) and thousands grouping
  (`+1,204%`).
- **One "chart could not be read" warning** per chart per run — `pptx_lint`
  dedupes the copies its own passes produce, and `qa_summary` drops
  `pptx_lint`'s copy when `qa_check` already reported the same slide.
- Extract the duplicated outline-replay block shared by `diff_deck.py` and
  `qa_checklist.py` — **done earlier in Cycle 6** (`outline_replay.py`); the
  `build_deck.py` / `geometry_report.py` splits also landed in Cycle 6.

~~**Still open:** group card-style core layouts … so the example deck reports
**L5** instead of **L4**~~ — done in the open-items round below (card layouts
grouped, example deck L5, test tightened to `== "L5"`).

~~**Known issue found along the way (pre-existing, not fixed):**
`charts.CHART_TYPES["scatter"]` cannot be built~~ — fixed in the open-items
round below (`XyChartData` on that code path, with an outline x/y contract).

## Cycle 4 — Existing-deck operations & traceability (research-driven)

- [x] **1. Restyle/normalize mode for existing decks.** The top practitioner pain
  point, and `references/editing.md` previously punted on it ("re-author
  instead"). `deck_to_outline.py`: existing `.pptx` → recovered markdown outline
  (titles, bullets, sources, notes, tables, native-chart data mapped back to
  `chart:` kinds, pictures exported to `<assets>/extracted/`) with heuristic
  layout classification, plus a per-slide fidelity report (`--report`) listing
  everything that could not carry over (SmartArt, unsupported chart types,
  extra exhibits). Round-trips the skill's own example deck in tests.
- [x] **2. Evidence ledger with source traceability.** `evidence.py` +
  `**Evidence-File:** evidence.json` + `- Evidence: E1` refs. Referenced sources
  synthesize the slide's `- Source:` footnote, full citations append to speaker
  notes, and a declared ledger escalates the unsourced-stat check to a fail-fast
  build/`--check` error. `diff_deck` verifies referenced ledger values survive
  into the built deck unchanged.
- [x] **3. Deck-level composition lint.** `pptx_lint.py` gains structure-monotony
  (3+ consecutive slides sharing one structural signature) and density-rhythm
  (4+ consecutive >130-word slides) checks; all checks now run — and emit — in a
  documented fix order (structure → geometry → honesty → color → type →
  AI tells → composition → export compat).
- [x] **4. Machine-readable delivery QA record.** `qa_summary.py` unifies
  qa_check / pptx_lint / geometry_report / diff_deck (narrative checks run
  inside build/--check, not on built decks) plus deck stats into one
  `qa-summary.json` with a `pass` verdict; exit 1 on any error (CI gate).
- [x] **5. Single-slide patch mode.** `patch_deck.py`: rebuild the outline in a
  scratch dir, extract slide N, merge into the delivered deck via
  `edit_merge.append_decks`, reorder into place, remove the stale slide. Refuses
  when outline and delivered slide counts drift (positions untrustworthy).
- [x] **6. Typography system in palettes.** Palettes already carried the font
  pair (font_title/body/label); they now gain a `type_scale` ladder (default =
  every size the builders emit; custom palettes may override). Lint enforces
  both: off-brand fonts are errors (symmetric with the color whitelist, CJK
  stack allowed), off-ladder sizes an aggregated warning (16:9 and 4:3-rescaled
  ladders accepted), size sprawl warned without a palette. Scope note:
  tabular-numeral policy was dropped — DrawingML has no per-run OpenType
  figure-style switch, so it is unenforceable in .pptx output.
- [x] **7. Data ingestion helpers.** `- Data-File:` accepts `.xlsx`/`.xlsm`
  (via optional openpyxl, active sheet, same 2-column / header+multi-series
  contract as CSV), and every loaded data file stamps a provenance line
  ("Data source: q3.csv (8 rows)") into the slide's speaker notes.

Explicitly not chasing (researched and rejected): broad style-family expansion and
AI-generated imagery (general-audience market; consulting restraint is the brand),
HTML→SVG→PPTX pipelines (native shapes are the moat — editability is what users
ask for), and web-app/API/collaboration features (different product).

## Cycle 3 — QA feedback loop (observed defects)

- [x] **1. Kicker-aware content budgeting.** The kicker band overlays the bottom edge
  of cards on `comparison`/`cards-3`, and `cards-4` + kicker produces genuine
  text-on-text overlap with the footer. Builders shrink the content area when
  `- Kicker:` is present; `--check` errors on layouts where the two cannot coexist.
- [x] **2. Vertical balance on sparse layouts.** `process-flow`, `exec-summary`,
  `cards-*`, and `timeline` content ends at ~50–55% slide height, leaving a dead band
  geometry_report flags as visual-mass imbalance. Vertically center the content block
  (or scale card/row heights to content count) when no kicker/source fills the bottom.
- [x] **3. Quoted-bullet handling.** `- icon:target "text"` rendered literal quote
  marks. Strip symmetric wrapping quotes in bullet parsing (after icon-prefix
  extraction) and keep intentional inner quotes.
- [x] **4. Per-slide palette awareness in pptx_lint.** Intentional `{palette=...}`
  overrides flag hard off-palette errors against the deck palette. Lint accepts
  extra palettes (`--palette a --palette b` or comma list) and reports per-slide
  matches instead of false errors.
- [x] **5. diff_deck attribute false positives.** `Icon="robot"` values are reported
  as missing text (icons render as images). Exclude icon/visual attribute values from
  expected-text extraction.
- [x] **6. Title orphan-word check.** `--check` predicts >2-line heading wraps but not
  single-word orphan second lines. Extend the heading-wrap estimator with an
  orphan-word warning.
- [x] **7. Card body top-alignment.** exec-summary-scqa (and sibling card layouts)
  vertically float body text at inconsistent heights across cards. Top-align card
  bodies to a shared first-line position.

## Cycle 3 — Larger bets (deferred from June research)

- [x] **8. File-size splits.** `builders_consulting.py` (1,237 lines) exceeds the
  800-line cap — extract chart-shape layouts into a sibling module.
  `edit_deck.py` (~870 lines) — extract the `append` part-graph/rId-remap merge
  engine into its own module. Pure refactors, no behavior change.
- [x] **9. Label collision avoidance (shape-built charts).** Track placed value-label
  rects on shape-built exhibits and nudge/suppress on overlap (waterfall and
  bar-mekko — the exhibits where labels can actually collide; tornado and
  football-field label positions are structurally collision-free). Native-chart
  labels stay PowerPoint-managed.
- [x] **10. Stacked-segment waterfalls with auto subtotals.** Multi-value delta rows
  stack segments within one bar; a `subtotal` sentinel (like `total`) auto-inserts a
  running-subtotal bar.
- [x] **11. Template-mode slide clustering (PPTAgent Stage I).** Cluster a user
  template's layouts by placeholder signature and map outline slides onto the
  best-matching cluster (text-heavy vs exhibit vs section), improving on pure
  signature matching in `template_helpers.py`.
- [x] **12. Axis breaks ("breaking bars").** For shape-built bar exhibits where one
  bar dwarfs the rest, clamp the tall bar and draw a break mark; opt-in via
  `- Axis-Break: on`.

## Deferred (not scheduled)

- **ONLYOFFICE second renderer integration** — x2t documented as an optional
  higher-fidelity renderer in qa-guide; full integration deferred until LibreOffice
  fidelity becomes a limiting factor.
- **pptx-glimpse second renderer** (`render_slides.py --engine glimpse`, npm) —
  LibreOffice-independent PNG/SVG with unsupported-element diagnostics; LibreOffice's
  Carlito substitution is not metric-exact for Calibri. Revisit with ONLYOFFICE.
- **Open XML SDK validation gate** (`qa_summary --strict-ooxml` via the dotnet
  OOXML-Validator CLI) — catches the "PowerPoint will repair this file" class;
  external dependency, so optional.
- **Map chart** (choropleth from Data-File) — L effort, low demand in our deck mix.
- **Chart scanner** (VLM-assisted extraction of picture-only exhibits in
  `deck_to_outline`) — L effort; fidelity report already names the gap.
- **Custom brand-font embedding** (`embeddedFontLst` + obfuscated fntdata) —
  licence-gated; document "install the font" instead.

## Explicitly not chasing (2026 re-check)

Unchanged from Cycle 4: broad style-family expansion, AI-generated imagery,
HTML/image→editable-pptx pipelines (a 2026 wave of such skills re-confirmed
"export is where everything breaks"), web-app/API/collaboration/analytics.
Added: COM automation of the user's PowerPoint (documented as wiping open decks —
anthropics/skills #886/#1523); Morph transitions / animations (niche for read-decks,
and LibreOffice cannot render them for QA).
