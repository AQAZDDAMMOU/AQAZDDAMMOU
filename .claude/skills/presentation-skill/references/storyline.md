# Storyline First (Consulting Discipline)

For strategy/consulting decks, write the argument before writing slides. This is the McKinsey dot-dash method: **dots become action titles, dashes become each slide's exhibit.**

## Step 1 — SCQA the executive summary

- **Situation**: the stable context everyone agrees on.
- **Complication**: what changed / why act now.
- **Question**: the question the deck answers (implicit).
- **Answer**: your recommendation — stated FIRST (Pyramid Principle: lead with the answer, then support).

## Step 2 — Dot-dash the body

```
• Cloud costs will double by 2028 on the current trajectory      ← dot = action title
   – line chart: cost projection vs budget                       ← dash = exhibit
• Three levers can hold costs flat without slowing delivery
   – waterfall: $18M gross savings bridge
• Lever 1 (workload tiering) is the fastest path to value
   – matrix-2x2: effort vs value, 8 initiatives plotted
```

Each dot is one slide. MECE-check the dots: no overlaps, nothing missing.

## Step 3 — The titles test

```bash
python3 scripts/build_deck.py outline.md --titles
```

Read the titles top to bottom. If they don't work as a standalone argument a partner could skim, fix the titles before building. The validator warns on topic-label headings (<5 words) on content slides.

## Action title rules

- Complete sentence stating the takeaway: "GPU consolidation saves $2.1M annually" ✔ / "Cost savings" ✘
- One claim per slide; the exhibit below must directly prove the title.
- Quantify when possible; the number belongs in the title.

## Structure conventions

- Open with `exec-summary-scqa` (Situation | Findings | Recommendation).
- `agenda` after the exec summary; repeat it at section breaks with `- Current:` set.
- Every exhibit slide carries `- Source:` (rendered as a footnote; the validator warns on unsourced statistics).
- End on `next-steps` (action/owner/when) — it stays on screen during Q&A. Never end on "Thank You".
- Backup material goes after a `## Appendix` line: slides are marked BACKUP, numbered `B·n`, and exempt from variety/notes warnings.

## Number formatting

`$4.2M` not `$4,200,000` · one decimal max on percentages (`38.2%`) · use `bps` for sub-percent moves · round to the precision the decision needs, not the data has.
