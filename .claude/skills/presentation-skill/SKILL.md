---
name: presentation-skill
description: |
  Use this skill whenever the user wants to create, generate, build, or export a presentation, slide deck, pitch deck, PowerPoint file, or .pptx file. Also trigger for consulting decks, strategy decks, client decks, steering-committee or board decks, exec readouts, and consulting visuals (waterfall/bridge charts, 2x2 matrix, harvey balls, mekko, gantt roadmaps). Also trigger when the user says "make slides", "build a deck", "create a presentation", "generate a PowerPoint", or uploads images and asks to use them in slides, or provides a .pptx template to match.
  Do NOT trigger for: HTML slide decks only (handled by slides skill), simple one-off diagram generation, or image editing without a slide output.
---

# Presentation Deck Skill

Generate professional, visually rich native PowerPoint (.pptx) presentations. See `references/decision-tree.md` for workflow routing.

## Quick Reference

| Goal | Path |
|------|------|
| Pick workflow | Read `references/decision-tree.md` |
| Create from scratch | Read `references/generation-guide.md` |
| Pick the right layout (gallery, all layouts) | Read `references/gallery.md` |
| Domain recipes (AI governance, TOM, control framework, use-case portfolio) | Read `references/recipes/*.md` |
| Pitch / strategy / talk arcs | Read `references/narratives/*.md` |
| Consulting storyline (SCQA, action titles, titles test) | Read `references/storyline.md` |
| Edit an existing .pptx | Read `references/editing.md` |
| Use a user's .pptx template | Read `references/template-mode.md` |
| Embed user-supplied images | Read `references/image-handling.md` |
| Generate charts natively | Read `references/charts-guide.md` |
| QA and render check | Read `references/qa-guide.md` |
| Match a company's brand | `scripts/brand_kit.py <domain> --name <n>` → `--palette <n>` |
| Auto-laid-out diagrams (Mermaid/Graphviz) | Read `references/diagrams.md` — native builders first |
| Pre-read handout | `scripts/gen_handout.py outline.md` |
| Split / merge decks | `edit_deck.py extract` / `append` — see `references/editing.md` |
| Restyle / normalize a foreign deck | `scripts/deck_to_outline.py old.pptx --output outline.md` → edit → rebuild — see `references/editing.md` |
| Patch one slide (or a batch) of a delivered deck | `scripts/patch_deck.py outline.md delivered.pptx "3,7,9" --output patched.pptx` — see `references/editing.md` |
| Reviewer comments / emailed feedback → outline edits | `edit_deck.py comments plan deck.pptx --outline outline.md --output revision-plan` — see `references/editing.md` |
| Rebuild only the slides whose data changed | `scripts/refresh_deck.py outline.md delivered.pptx --output refreshed.pptx --log refresh-log.md` — see `references/editing.md` |
| Split / merge / move slides in the outline | `scripts/outline_ops.py outline.md split 4 --in-place` — see `references/editing.md` |
| One slide per data row | `**Repeat:** data/initiatives.csv` + `{column}` fields — see `references/generation-guide.md` |
| Copy must land byte-identical (legal / approved messaging) | `**Copy:** locked` or `--verbatim` — see `references/generation-guide.md` |
| Reuse a client's standards across engagements | `**Profile:** acme` / `--profile acme`; **read `<assets>/profiles/acme.lessons.md` before building** — see `references/generation-guide.md` |
| Trace material numbers to sources | `**Evidence-File:** evidence.json` + `- Evidence: E1` — see `references/generation-guide.md` |
| One number format deck-wide | `**Number-Format:** units=M decimals=1 negatives=parens` — see `references/generation-guide.md` |
| House style (word lists, title rules) | `**Style-File:** style.json` or `--style path` — see `references/generation-guide.md` |
| Deck author in the file properties | `**Author:** "Alex Kim"` (sets `dc:creator`/`lastModifiedBy`) — see `references/generation-guide.md` |
| Scrub a deck before sending it out | `scripts/pptx_lint.py deck.pptx --client-safe [--denylist terms.txt]` → `edit_deck.py clean --client-safe` |
| Binary VLM review checklist | `scripts/qa_checklist.py outline.md --output assets/qa-checklist.md` |
| One machine-readable QA record | `scripts/qa_summary.py deck.pptx --outline outline.md --palette <p> --output qa-summary.json` |

---

## Phase 0 — Detect Mode

Determine the correct mode before proceeding:

**Mode A: From Scratch**
No template file provided. Claude generates a fully styled deck using the outline + design system.

**Mode B: Template-Based**
User provides an existing .pptx file to use as the visual template. Read `references/template-mode.md`.

**Mode C: Image-Rich Deck**
User has shared one or more images (photos, diagrams, screenshots, brand assets) to embed in specific slides. Read `references/image-handling.md`. This mode can be combined with A or B.

**Mode D: Edit Existing Deck**
User provides a finished .pptx and wants changes (fix text, swap images, add/remove/reorder slides, or split/merge decks via `edit_deck.py extract`/`append`). Do NOT rebuild — edit surgically. Read `references/editing.md`. Skip Phases 1–3 and go straight to the editing workflow + Phase 4 QA. Exception: when the user wants the whole deck restyled/normalized ("make it match our brand", "clean this up"), use the restyle path — `deck_to_outline.py` extracts the content into an outline (with a fidelity report), then Phases 2–4 run as normal on that outline.

**Mode E: Revise a Deck This Toolchain Built**
The deck came back with reviewer comments, emailed feedback, or fresh data, and its outline still exists. Do NOT rebuild the whole deck (that clobbers hand edits) and do NOT hand-edit the .pptx (the outline stops being the source of truth). Fix the **outline**, then patch only the affected slides — `edit_deck.py comments plan|apply` for comments, `refresh_deck.py` for new data, `outline_ops.py` for structure, `patch_deck.py` for the swap. Read `references/editing.md` § "Revision loop (decks this toolchain built)", then run Phase 4 QA.

---

## Phase 1 — Content & Style Discovery

Ask ONE clarifying message covering all of the below before starting. Combine into a single natural prompt, not a numbered list.

1. **Topic / title** — What is the presentation about? Who is the audience?
2. **Length** — Approximately how many slides? (Short 5–10 / Medium 10–20 / Long 20+)
3. **Purpose** — Pitch deck / Internal strategy / Client presentation / Conference talk / Tutorial
4. **Tone** — Bold & dark / Clean & corporate / Warm & editorial / Technical & minimal
5. **User images** — Has the user shared any images? If yes, ask which slides to place them on (see `references/image-handling.md`).
6. **Template** — Has the user provided a .pptx template file? If yes, switch to Mode B.
7. **Source documents** — Users may supply PDF/docx/xlsx/CSV source documents to draft the outline from; CSV/xlsx files can also feed charts directly via `- Data-File: path.csv` (provenance is stamped into speaker notes). For decks where numbers must be defensible (board/consulting), offer an evidence ledger (`**Evidence-File:**` + `- Evidence:` refs — `references/generation-guide.md`).
8. **Recurring client / engagement** — if this deck is for a client the skill has built for before, check `<assets>/profiles/` for a matching profile. When one exists, build with `**Profile:** <name>` (or `--profile <name>`) and **read `<assets>/profiles/<name>.lessons.md` before writing the outline** — it lists the QA findings that recurred on past decks for this client, so you can avoid them this time instead of rediscovering them in Phase 4. See `references/generation-guide.md`.

After gathering answers, research the topic if needed using available tools.

---

## Phase 2 — Outline (User Approval Gate)

0. **Strategy/consulting decks first draft the storyline** (`references/storyline.md`): SCQA the executive summary, dot-dash the body, then run the titles test (`build_deck.py outline.md --titles`). Action titles only — full-sentence takeaways, never topic labels.
1. Write the slide outline in markdown (one `## Slide N:` section per slide) using the exact syntax in `references/generation-guide.md`. Follow its Design Rules: vary layouts, use icons (`icon:name`) on bullet slides, write `- Notes:` speaker notes for every slide, and `- Source:` on every exhibit.
2. Validate it: `python3 scripts/build_deck.py outline.md --check` — fix errors; address warnings (missing notes, house-style findings such as marketing adjectives, hedges, passive or duplicate titles, etc.).
3. **Show the outline to the user and ask for approval or edits before building.** The outline is the reviewable artifact; changing a slide here is cheap, regenerating a deck is not. Skip this gate only if the user explicitly asked for a one-shot build.

---

## Phase 3 — Generation

1. If user images are involved, normalize first: `python3 scripts/prep_images.py assets/user-images/`
2. Run `python3 scripts/build_deck.py outline.md --output deck.pptx [--palette X] [--template T.pptx] [--assets-dir DIR] [--density compact|comfortable] [--variant a|b|c] [--ghost] [--profile NAME] [--verbatim]`
   - `--ghost` builds a skeleton deck (real action titles, grey labeled exhibit placeholders) for storyline sign-off before investing in content.
   - `--profile NAME` applies a saved engagement profile (palette, footer, stamp, banned phrases, mandatory slides); `--verbatim` locks the copy so outline text lands byte-identical.
   - Custom brand palettes: drop `<name>.json` into `<assets>/palettes/` and use `--palette <name>` (schema: `references/generation-guide.md`).
   - Brand-match decks: `python3 scripts/brand_kit.py <domain> --name <n>` generates a custom palette (+ logo) from a company's web presence, then build with `--palette <n>`.
3. The build fails fast on validation errors. If any slide fails during build, **no .pptx is written**.

Optional: `python3 scripts/gen_appendix.py outline.md` for pitch/strategy appendix skeleton; `python3 scripts/gen_handout.py outline.md` for a pre-read markdown handout.

---

## Phase 4 — QA (Required)

Never skip QA. Seven complementary checks (see `references/qa-guide.md`):

1. **Programmatic:** `python3 scripts/qa_check.py deck.pptx` — bounds, tiny text, contrast, font-true overflow, and editability warnings (text parked on a shape, free connectors, tiny fonts + dead space). `--accessibility` for WCAG AA strict mode; `--integrity` for Keynote/macOS package hygiene (always runs) plus optional OOXML schema validation
2. **Deck lint:** `python3 scripts/pptx_lint.py deck.pptx --palette <palette>` — cross-slide consistency in fix order (page sequence, jiggle, axis honesty, chart honesty, numeric consistency, off-palette colors, off-brand fonts/type scale, AI tells, layout monotony, density rhythm)
3. **Content diff:** `python3 scripts/diff_deck.py outline.md deck.pptx` — dropped outline content is an error, untraceable "added" deck text a warning (`--strict` to escalate)
4. **Geometry (pre-render):** `python3 scripts/geometry_report.py deck.pptx` — deterministic overlap/spacing/whitespace/balance metrics; fix findings before spending a render
5. **Visual:** `python3 scripts/render_slides.py deck.pptx --grid --out assets/qa-thumbs/` + fresh-eyes subagent answering the binary checklist from `python3 scripts/qa_checklist.py outline.md` (see `references/qa-guide.md`; grid cells are numbered)
6. **Consistency (LLM):** `python3 scripts/qa_check.py deck.pptx --numbers` + titles test — cross-check totals, repeated KPIs, title claims (see `references/qa-guide.md`)
7. **Fix loop:** edit outline → rebuild → re-run all checks until clean; on revision cycles, diff against the last delivered render with `python3 scripts/visual_regress.py assets/qa-baseline/ assets/qa-current/`

Before a deck leaves the building: `python3 scripts/pptx_lint.py deck.pptx --client-safe [--denylist terms.txt]` — comments, hidden slides and denylisted client/internal terms are errors; clean them with `edit_deck.py clean --client-safe` (`references/editing.md`).

Final gate: `python3 scripts/qa_summary.py deck.pptx --outline outline.md --palette <p> --output qa-summary.json` unifies checks 1-4 into one machine-readable record (exit 1 on any error) — its stats plus the deck's `editability` level (L0–L5) double as the delivery note. Add `--comments reviewed.pptx` to get a copy of the deck with every finding pinned to its slide as a PowerPoint comment. On a profiled deck add `--profile <name>` so findings that recur across runs are appended to `<assets>/profiles/<name>.lessons.md` — the file Phase 1 tells you to read next time.

Deliver the `.pptx` and the thumbnail grid together.

---

## Dependencies

```bash
pip install -r requirements.txt
pip install pillow-heif   # only needed for HEIC (iPhone) image inputs
pip install cairosvg      # optional: Lucide SVG icon fallback (Tabler PNG needs nothing)
# Optional: rendered diagrams (references/diagrams.md)
# brew install graphviz            # .dot/.gv support
# npm i -g @mermaid-js/mermaid-cli # .mmd support
# Optional for high-fidelity QA rendering:
# LibreOffice (soffice) + poppler (pdftoppm)
# Optional: pip install unoserver, then run `unoserver` in the background —
# render_slides.py auto-uses it (much faster repeated QA renders)
```

Smoke test (validate + build + QA on the example outline):

```bash
python3 scripts/smoke_test.py
pytest tests/
```
