# Editing Existing Decks (Mode D)

Use this when the user provides a .pptx and wants changes — fix text, swap an image, add/remove/reorder slides — while preserving the existing design exactly. Never rebuild a deck the user already has; edit it surgically.

**First, check where the deck came from.** If this toolchain built it and the outline still exists, skip the OOXML surgery below and go to **"Revision loop"** further down: fix the outline, then patch only the affected slides (`comments plan|apply`, `outline_ops`, `refresh_deck`, `patch_deck`). Hand-editing a deck whose outline exists forks the two, and the next rebuild silently loses the edit. The unpack/pack path below is for **foreign** decks — no outline, or an outline that no longer describes the file.

## Workflow

```bash
python3 scripts/render_slides.py original.pptx --grid --out review/   # see what exists
python3 scripts/edit_deck.py unpack original.pptx unpacked/           # extract + pretty-print
python3 scripts/edit_deck.py list unpacked/                           # slide order + titles
# 1) structural changes FIRST (N = 1-based position from list, NOT slide file number):
python3 scripts/edit_deck.py duplicate unpacked/ 3                    # clone slide at position 3
python3 scripts/edit_deck.py remove unpacked/ 7                       # drop slide at position 7
python3 scripts/edit_deck.py reorder unpacked/ 3,1,2,4               # new slide order by position
python3 scripts/edit_deck.py clean unpacked/ [--client-safe]          # orphan media (+ pre-send scrub)
# 2) then content edits: Edit tool on unpacked/ppt/slides/slideN.xml
python3 scripts/edit_deck.py pack unpacked/ edited.pptx               # validate + re-zip
python3 scripts/qa_check.py edited.pptx                               # QA gate
```

`pack` refuses to build if any XML fails to parse, so syntax mistakes surface immediately instead of as a corrupt file PowerPoint rejects.

## Editing slide XML

- Text lives in `<a:t>` elements inside runs (`<a:r>`). Edit the text between the tags; never touch the surrounding `<a:rPr>` (run formatting) unless changing formatting is the goal.
- One paragraph = one `<a:p>`. To bold a header run: `<a:rPr b="1" .../>`.
- **Never insert literal unicode bullets (•)** — bullet glyphs come from `<a:buChar>`/`<a:buAutoNum>` in paragraph properties.
- Leading/trailing spaces need `xml:space="preserve"` on the `<a:t>`.
- Smart quotes in new text: use XML entities `&#x201C; &#x201D; &#x2018; &#x2019;`.
- To reorder slides, rearrange `<p:sldId>` elements in `ppt/presentation.xml` → `<p:sldIdLst>` (the order there IS the deck order).
- To swap an image: replace the file in `ppt/media/` keeping the same filename and format.

## Gotchas

| Issue | Rule |
|-------|------|
| Manual slide file copying | Never — `duplicate` updates Content_Types, rels, and sldIdLst; hand-copying misses them and corrupts the deck |
| Structural vs content edits | Do ALL duplicate/remove/reorder operations BEFORE text edits |
| Position vs file number | `duplicate` / `remove` take the **1-based position** shown by `list`, not `slideN.xml` file numbers |
| Template slots ≠ your items | When a layout has 4 boxes and you have 3 items, delete the whole leftover group (shape + text), not just its text |
| Longer replacement text | May overflow the original box — always run visual QA after edits |
| Notes on duplicated slides | `duplicate` intentionally drops the notes link; add new notes if needed |
| Tags on duplicated slides | `duplicate` copies the slide's `ppt/tags/` part and re-keys the clone's `PK_KEY` to `<key>-copy` (`-copy2`, …) so `patch_deck`/`refresh_deck` still resolve the source slide; untagged (foreign) decks are unaffected |

## Quick text-only fixes

For a pure find/replace (a typo, a number), prefer python-pptx in place of unpacking:

```python
from pptx import Presentation
prs = Presentation("original.pptx")
for slide in prs.slides:
    for shape in slide.shapes:
        if shape.has_text_frame:
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    run.text = run.text.replace("2025", "2026")
prs.save("edited.pptx")
```

This preserves all formatting because only run text changes.

## Batch text edits: inventory → replace

For text-only changes across many slides, prefer the JSON workflow over
hand-editing XML:

    python3 scripts/edit_deck.py unpack deck.pptx work/
    python3 scripts/edit_deck.py inventory work/ > inventory.json
    # edit the "text" fields you want to change (keep slide/run addresses),
    # save the changed entries (only those) as edits.json
    python3 scripts/edit_deck.py replace work/ edits.json
    python3 scripts/edit_deck.py pack work/ deck-edited.pptx

Addresses are slideN.xml file numbers (these differ from deck position
after a reorder — correlate with the `list` subcommand). Run formatting
(bold/size/color) is preserved — only the text changes. Then run the
standard Phase 4 QA on the result.

## Slide operations playbook

Precise recipes for common edit requests, built from the existing
primitives. Nothing here needs new tooling — the discipline is in the
sequencing.

### condense / rewrite / formalize / translate

All four are the same mechanical recipe — only the text you write differs
(shorter, reworded, formal register, target language):

```bash
python3 scripts/edit_deck.py unpack deck.pptx work/
python3 scripts/edit_deck.py inventory work/ > inventory.json
```

Inventory output (snippet) — every text run with its stable address:

```json
[
  {"slide": 2, "run": 0, "text": "Three priorities driving $14.2M in value"},
  {"slide": 2, "run": 1, "text": "Phase 1 GPU cluster deployed ahead of schedule, saving $2.1M annually"},
  {"slide": 2, "run": 2, "text": "Hybrid cloud migration 62% complete — on track for Q4"}
]
```

Write `edits.json` containing ONLY the runs you change, keeping the
`slide`/`run` addresses (here: condensing run 1):

```json
[
  {"slide": 2, "run": 1, "text": "Phase 1 GPUs live early — $2.1M/yr saved"}
]
```

```bash
python3 scripts/edit_deck.py replace work/ edits.json
python3 scripts/edit_deck.py pack work/ deck-edited.pptx
python3 scripts/qa_check.py deck-edited.pptx
```

**Run-granularity caveat:** one visual sentence can be split across several
runs (e.g. a bold figure mid-sentence produces three runs: before / bold /
after). `replace` works at run level, so rewrite each piece in place — or
put the whole new sentence in the first run and set the leftover runs of
that sentence to `""`. Never merge runs by hand-editing XML for this; you
lose the formatting boundaries. For translate, keep numbers/units as their
own runs when they already are — only the prose runs change.

### split (one overloaded slide → two)

**Outline-backed deck?** Use `outline_ops.py outline.md split N` instead (see
"Revision loop" below) — it splits the *outline*, so the two halves are real
built slides rather than a clone with half its runs blanked. The recipe below
is for foreign decks only.

1. `python3 scripts/edit_deck.py list work/` — find the slide's position N.
2. `python3 scripts/edit_deck.py duplicate work/ N` — the clone lands at
   position N+1 (notes link intentionally dropped).
3. `inventory` → write two edit sets: in the original, blank (`""`) the runs
   for the half that moves; in the clone (a NEW slideN.xml file number —
   re-run `inventory` to see it), blank the half that stayed, and retitle
   both (e.g. "… (1/2)" / "… (2/2)" or sharper per-half titles).
4. `replace` → `pack` → visual QA. Blanked runs leave empty shapes behind;
   if a whole box/group is now empty, delete that shape's `<p:sp>` in the
   slide XML rather than leaving a ghost slot.

### remix (re-layout a slide)

- **Deck built from an outline (Mode A/B):** edit that slide's section in
  the outline — change `**Layout:**` to the new layout — and rebuild. When
  the delivered deck has picked up hand edits a full rebuild would clobber,
  patch just that slide instead (see "Patching slides" below):

  ```bash
  python3 scripts/patch_deck.py outline.md delivered.pptx N --output patched.pptx
  ```

- **Template-mode decks:** `add_slide.py deck.pptx out.pptx --layout K`
  adds a blank slide from the template's own layout K; fill placeholders
  with python-pptx, then `remove`/`reorder` the old slide as above.
- **Foreign decks (no outline, no usable template layouts):** there is no
  automated remix — the slide's design exists only as its own XML. Rebuild
  the slide content on the nearest skill layout (it will carry skill
  styling, not the source deck's), or restructure within the existing
  layout via `replace` + shape edits. To re-author the *whole* deck onto
  skill layouts, use the restyle path (see "Restyling a foreign deck").

## Revision loop (decks this toolchain built)

When the deck came back from review and its **outline still exists**, do not
rebuild the whole deck (that clobbers every hand edit it picked up) and do
not hand-edit the `.pptx` (the outline stops being the source of truth).
Fix the outline, then patch only the slides that changed:

```bash
# 1. read what the reviewers actually said
python3 scripts/edit_deck.py comments list reviewed.pptx

# 2. comments (+ emailed feedback) → a revision plan mapped onto the outline
python3 scripts/edit_deck.py comments plan reviewed.pptx \
    --outline outline.md [--feedback feedback.md] [--assets-dir assets] \
    --output revision-plan                  # writes revision-plan.json + .md

# 3. apply the mechanical items to the outline (idempotent)
#    the outline is copied to outline.md.bak first, and is only rewritten
#    when something actually applied (line endings are preserved)
python3 scripts/edit_deck.py comments apply revision-plan.json outline.md

# 4. YOU write the `manual` items into the outline by hand — that is most of them

# 5. rebuild and swap only the affected slides
python3 scripts/build_deck.py outline.md --check
python3 scripts/patch_deck.py outline.md reviewed.pptx "3,7,9" --output revised.pptx

# 6. drop the comments that are now handled, then run the QA gate
python3 scripts/edit_deck.py comments strip revised.pptx --slides 5,8 --output clean.pptx
python3 scripts/qa_summary.py clean.pptx --outline outline.md --palette X
```

**What the plan contains.** One item per comment thread (and per addressed
paragraph of `feedback.md`), each carrying the author, the deck slide, the
outline slide and field it maps to, the anchored shape (by its semantic
name, e.g. `s04-text-2`) and text snippet, the thread's replies, its
resolved state, and a parsed `action`. Reading order in `revision-plan.md`
is deck order, grouped by slide. Resolved threads are rendered but skipped
on `apply`.

**What `apply` will do for you** — only this grammar, case-insensitively,
after stripping a leading `@claude`/`please`:

| Comment | Action |
|---|---|
| `delete bullet 3` / `delete this` | remove that bullet (`delete this` uses the anchored one) |
| `replace "X" with "Y"` or `X -> Y` | in-place text replacement (on the heading it also rewrites a `- Heading:`/`- Title:` that echoed it) |
| `retitle: T` / `title: T` | rewrite the `## Slide N:` heading and any `- Heading:`/`- Title:` |
| `add bullet: T` | append a bullet |
| `move to slide N`, `kill slide`, `split slide`, `merge with slide N` | slide ops via `outline_ops` |

**Everything else is a `manual` item** — "make this punchier", "the CFO will
push back here", "wrong chart". That is by design: the tool does the
mechanical edits and leaves the judgment to you. Work the `manual` list in
`revision-plan.md` yourself, editing the outline.

Guarantees worth relying on: `apply` is **idempotent** (a delete verifies
the text it recorded before removing anything, and skips rather than
deleting the wrong line on a re-run; an already-applied replace is a skip,
not an error). Idempotency is exact, not best-effort: every run records
what it applied in `<plan-stem>.applied.json` (merged with what was already
there), and a later run skips those as `already applied` — so even
`replace "X" with "X."`, whose replacement still contains what it replaced,
fires exactly once. The ledger is keyed on **comment identity** (the
thread id; emailed feedback by the text of the note), not on the item's
`c3`-style number, which is positional and shifts as soon as a comment is
stripped, resolved or added. So the same `revision-plan.applied.json` can
sit next to a plan you re-planned for round 2: an unchanged comment is
still skipped, and round 2's brand-new items are applied even when they
land on numbers round 1 used. A ledger written by an older version (keyed
by number) is ignored with a notice rather than trusted, and one written
for a different deck is warned about, not refused. It reports every skip
with a reason, and after a `move`,
`merge` or `kill` renumbers the outline it **stops** applying further
slide-level ops and tells you to re-plan against the rewritten outline.

Comments on slides you rebuild disappear with the slide (patch replaces it);
`comments strip` is for the ones on slides you did *not* rebuild. Note
`--resolved` only strips modern comments — the legacy `p:cm` format carries
no resolved state, so those always survive; use `--slides` or `--all` for
them.

### Restructuring the outline: outline_ops

`outline_ops.py` rewrites the markdown outline itself, leaving every byte
outside the target slides untouched:

```bash
python3 scripts/outline_ops.py outline.md split 4 [--at 3] [--heading "…"] --in-place
python3 scripts/outline_ops.py outline.md merge 5,6 [--heading "…"] [--layout bullet-list] --output new.md
python3 scripts/outline_ops.py outline.md move 9 2 --in-place
python3 scripts/outline_ops.py outline.md renumber --in-place
```

`split` divides the slide's list content (bullets, cards, items, steps,
table rows, tiles) at `--at` (default: half), copies layout, palette,
kicker and source to both halves, and appends " (1/2)" / " (2/2)" to the
headings unless `--heading` is given. It **refuses** on single-exhibit
layouts (a waterfall has nothing to split) and on slides carrying a
`**Data:**` block. `merge` concatenates list content and refuses when the
two slides declare different layouts unless you pass `--layout`. Every
subcommand renumbers `## Slide N:` sequentially and then runs `--check` on
the result: a validation failure is reported and **exits 1, but the
rewritten file is still written** — you asked for the edit, so you get it
and the error list. `renumber` on a well-formed outline is a byte-level
no-op.

### Refreshing numbers only: refresh_deck

The night-before scenario: the partner sends revised Q4 figures, the CSVs
behind two exhibits are overwritten, nothing else about the deck moved.

```bash
python3 scripts/refresh_deck.py outline.md delivered.pptx \
    --output refreshed.pptx --log refresh-log.md [--dry-run] [--all] \
    [--palette X] [--assets-dir assets] [--template T.pptx] \
    [--profile NAME] [--verbatim]
```

The build options are handed straight to `patch_deck` — repeat whatever the
deck was delivered with (see "Repeat the build options" below).

Every slide this toolchain builds carries a `PK_DATA_SHA` tag — the hash of
the `- Data-File:` bytes (or the inline `- Data:` rows) that produced its
exhibit. `refresh_deck` replays the outline, recomputes that hash per slide,
and rebuilds **only** the slides where the two disagree (through
`patch_deck`, so hand edits everywhere else survive byte-for-byte). It also
answers the question the partner actually asks:

```
## Slide 4 — revenue-compounds-through-fy25
- Revenue@FY25: 12 → 14 (+2, +16.7%)
```

Numbers that exist on only one side are listed under **unpaired** rather
than guessed at — a renamed series or an added category is a content change,
not a data refresh, and belongs in the outline. Chart values are keyed
`series@category`, table cells `rowlabel@columnheader`, stat tiles by their
caption; a caption or row label repeated on one slide is numbered in tile /
row order on both sides (`Run-rate`, `Run-rate (2)`, …) so each pairs with
its own counterpart. `--dry-run` prints the
changed list and the log and writes nothing; `--all` rebuilds every tagged
slide. Two scope limits: refresh **never inserts slides** (an outline slide
with no tagged match is reported and skipped — adding a slide is
`patch_deck` / `edit_deck` work), and a delivered deck with no slide tags
at all cannot be matched safely, so the CLI exits **2** and asks for one
rebuild with this version first.

## Patching slides: patch_deck

`patch_deck.py` rebuilds one slide — or a batch — from the outline and swaps
them into an already-delivered deck, preserving hand edits everywhere else:
the alternative to a full rebuild once a delivered deck has been touched by
humans.

```bash
python3 scripts/patch_deck.py outline.md delivered.pptx 4 \
    --output patched.pptx [--palette X] [--assets-dir assets] \
    [--template T.pptx] [--profile NAME] [--verbatim]
python3 scripts/patch_deck.py outline.md delivered.pptx "3,7,9" --output patched.pptx
python3 scripts/patch_deck.py outline.md delivered.pptx "2-4" --output patched.pptx
```

**Repeat the build options.** The scratch rebuild has to be given the same
`--palette` / `--template` / `--profile` / `--verbatim` the deck was
delivered with, or the patched slides come back styled differently from
their neighbours — a profile's palette and footer, or locked copy, quietly
lost on exactly the slides you replaced. Options the *outline* declares
(`**Palette:**`, `**Profile:**`, `**Copy:** locked`) need nothing. Verbatim
is the one option the deck itself records (slide tag `PK_FLAGS`): patching a
verbatim-built deck without `--verbatim` (or a `**Copy:** locked` outline) is
refused rather than silently unlocking the copy, and the mirror is refused
too — a deck that was *not* built verbatim cannot be patched from a locked
rebuild (`rebuild the whole deck once`), which would leave the copy locked on
only the patched slides. An outline the guard cannot read is reported as such
(`outline <path> unreadable: …`) rather than as a verbatim mismatch. A CLI-only
`--profile` is *not* recorded anywhere — repeating it is on you.

The selection is 1-based **outline** positions: a single number, a comma
list, or a range. The outline is rebuilt once in a scratch dir (palette,
trackers, and page numbers get full deck context), the selected slides are
extracted from that rebuild, merged into the delivered deck via the
rId-remapping append engine, moved into position, and the stale slides
removed.

**Slide matching.** Every built slide carries a `ppt/tags/` identity — an
outline slug key (`PK_KEY`) plus a title hash. When the delivered deck has
those tags, each selected outline slide is located by identity, so a slide a
reviewer dragged to a new position, or that now sits after an inserted
slide, is still found and replaced correctly; the summary prints
`outline slide 3 (pricing-bridge) -> position 5`. Guardrails: an
outline slide with no match is an error naming the key (nothing is
guessed), and **duplicate keys are refused** rather than silently
first-matched — a deck where some other tool cloned a slide's tags part
verbatim has two slides with one key, so patch that deck `--by-position`
or reconcile it first. (`edit_deck duplicate` itself gives the clone its
own tags part re-keyed `<key>-copy`, `-copy2`, … — unique across the deck —
so the source slide keeps patching; the clone is never an outline slide and
is left alone.) `--by-position`, and an untagged (foreign or pre-tags) deck, fall
back to the older contract: same 1-based position in both files, which
requires the outline and delivered slide counts to match exactly.

`--output` must differ from the delivered deck (rollback copy). After
patching, run `qa_check.py` and `visual_regress.py` against the delivered
baseline — only the patched slides should differ.

## Restyling a foreign deck: deck_to_outline

The extraction path for "make this deck not ugly" requests: recover an
existing deck's content into the outline format, then rebuild it on this
toolchain's layouts and palettes:

```bash
python3 scripts/deck_to_outline.py old.pptx --output outline.md \
    --palette midnight-executive --report fidelity.json
# review/edit outline.md (headings, layouts, dropped items), then:
python3 scripts/build_deck.py outline.md --check
python3 scripts/build_deck.py outline.md --output restyled.pptx
```

What it recovers: slide titles (largest text), body paragraphs as bullets,
`Source:` footnotes, speaker notes, tables as markdown rows, native-chart
data (categories/series/values, mapped back to `chart:` kinds), and
embedded pictures (exported to `<assets>/extracted/` and referenced by
path). Layouts are classified heuristically (title / section-divider /
table / two-column-split / full-image / bullet-list) — treat them as a
starting draft and upgrade to consulting layouts where the content fits.

Extraction is lossy **by design** — the goal is recovering content, not
cloning geometry. Everything that cannot carry over (SmartArt, unsupported
chart types, extra charts/tables on one slide, freeform art) is listed per
slide in the fidelity report (`--report`) rather than dropped silently.
Review the report and the `--check` output before delivering; re-author
dropped exhibits on the nearest skill layout.

## Splitting a deck: extract

Pull a contiguous or cherry-picked subset into a new standalone deck —
positions are the 1-based order `list` shows:

```bash
python3 scripts/edit_deck.py extract deck.pptx 3-7 --output sub.pptx
python3 scripts/edit_deck.py extract deck.pptx 1,3-5,9 --output sub.pptx
```

The subset keeps deck order regardless of how the selection is written.
Unreferenced leftovers (notes, charts, embedded workbooks, media of dropped
slides) are garbage-collected so the output is clean. `--output` must
differ from the input. Speaker notes of KEPT slides survive.

## Merging decks: append

Copy slides from one deck into another, after the destination's last slide,
preserving the source slides' look:

```bash
python3 scripts/edit_deck.py append dst.pptx src.pptx --output merged.pptx
python3 scripts/edit_deck.py append dst.pptx src.pptx --slides 2-4 --output merged.pptx
```

Each copied slide brings its full dependency graph — slide layout, slide
master, theme, images, charts (including chart colors/style and the
embedded .xlsx data) — renamed to avoid collisions and re-registered in
content types and `presentation.xml`. Slides sharing a layout/master are
deduplicated within one append run.

**Known limitations:**

- Source slides keep their own master/theme — faithful import, no
  restyling. The merged deck shows each slide as it looked in its source
  deck; expect a visible style seam unless both decks share a design. To
  restyle instead, re-author the content on the destination deck's layouts.
- Speaker notes on copied slides are dropped (same policy as `duplicate`).
- If the decks have different slide sizes (16:9 vs 4:3), copied shapes keep
  their source geometry on the destination's slide size — check visually.
- Slide-to-slide hyperlinks pointing outside the copied selection are
  dropped (a warning is printed).
- Appended slides keep their original baked page numbers — re-number via
  `inventory`/`replace` if needed.

After any merge, run the standard QA gate (`qa_check.py`, render review) —
and flip through the seam slides specifically.

## Pre-send scrub: client-safe clean

Audit first (read-only, on the packed file), then clean:

```bash
python3 scripts/pptx_lint.py delivered.pptx --client-safe --denylist terms.txt
python3 scripts/client_safe.py scan delivered.pptx --denylist terms.txt   # same audit, standalone
```

Comments, hidden slides and denylist hits are errors; core-props identity,
off-slide/hidden shapes, unused layouts/masters, embedded chart workbooks,
the package thumbnail and the presence of speaker notes are warnings (full
list in `references/qa-guide.md`, Step 0b).

To fix them, unpack and clean:

```bash
python3 scripts/edit_deck.py unpack delivered.pptx unpacked/
python3 scripts/edit_deck.py clean unpacked/ --client-safe [--drop-notes] [--drop-hidden]
python3 scripts/edit_deck.py pack unpacked/ client-ready.pptx
```

`--client-safe` drops comment parts and the package thumbnail, removes
unused layouts and masters, and scrubs the core-properties identity while
keeping `dc:title`. `--drop-notes` removes every speaker-notes part and
`--drop-hidden` removes hidden slides — each implies `--client-safe`. Both
are destructive and neither is on by default: notes are often the point of
the deck, and a hidden slide may be a deliberate backup. Plain `clean` with
no flags is still just the orphan-media sweep.

## Sanitizing numbers (turning a client deck into a sample)

```bash
python3 scripts/edit_deck.py sanitize-numbers unpacked/ [--seed 7]
```

Replaces every run of 3+ digits in slide text and in chart caches with
random digits; years 1900–2100 and 1–2 digit tokens survive so labels and
axes still read as a deck. `--seed` makes it reproducible. The embedded
`.xlsx` workbooks behind native charts are **left untouched by design** —
the cached values are what renders, and rewriting the workbooks would
desynchronise them. Delete the embeddings (or rebuild the deck from a
sanitized outline) if the workbook data itself must not ship.

## Keynote-safe hygiene on a foreign deck

```bash
python3 scripts/qa_check.py foreign.pptx --integrity          # detect
python3 scripts/save_hygiene.py foreign.pptx --output fixed.pptx \
    [--author "Alex Kim"] [--title "Q3 Review"]               # fix
```

Decks written by python-pptx (ours included, before Cycle 5) ship a 4:3
`sldSz` type on a 16:9 canvas, a `notesMaster` relationship with no
`notesMasterIdLst`, a leftover `printerSettings1.bin`, and stock core
properties — the trio that makes Keynote and Quick Look refuse the file.
Builds from this toolchain are sanitized at save, so this path is for decks
that came from elsewhere. Without `--output` the file is fixed in place;
`--title` defaults to the filename stem; omitting `--author` preserves a
real existing `dc:creator` and replaces only an empty or stock one. Exit 0
when the saved package is clean.

## Patching decks that use navigation (Cycle 7)

`patch_deck.py` re-points other slides' navigation hyperlinks (agenda rows,
tracker chips) at the replacement slide, so swapping a slide that is a jump
target never leaves dangling links. The *patched slide's own* navigation
links are stripped during extraction (each prints a `stripping dangling
slide link` warning) — rebuild the full deck when a patched slide must keep
live agenda/tracker jumps.
