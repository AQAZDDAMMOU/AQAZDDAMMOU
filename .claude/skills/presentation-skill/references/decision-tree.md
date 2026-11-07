# Decision Tree

Use this to pick the correct workflow before building.

```
User request involves .pptx?
├── User has finished deck + wants edits?
│   └── Mode D → references/editing.md (unpack → edit → pack)
├── User provided .pptx template to match?
│   └── Mode B → profile_template.py → references/template-mode.md
│       └── build_deck.py --template T.pptx outline.md
├── User shared images for specific slides?
│   └── Mode C (+ A or B) → prep_images.py → references/image-handling.md
└── Create new deck from scratch
    └── Mode A → references/generation-guide.md
        ├── Pick purpose → references/narratives/*.md
        ├── Domain deck (AI governance / TOM / portfolio)? → references/recipes/*.md
        ├── Write outline.md → build_deck.py --check
        ├── User approves outline
        ├── build_deck.py --output deck.pptx [--variant b]
        └── QA: qa_check.py + diff_deck.py + render_slides.py + subagent
```

## Tool selection

| Need | Tool |
|------|------|
| Validate outline | `build_deck.py --check` |
| Build deck | `build_deck.py outline.md --output deck.pptx` |
| A/B variant | `--variant b` (aurora + comfortable) |
| Template brand colors | `--template corp.pptx` (+ optional `.config.json`) |
| Content proofread | `diff_deck.py outline.md deck.pptx` |
| Geometric QA | `qa_check.py deck.pptx` |
| Accessibility AA | `qa_check.py deck.pptx --accessibility` |
| OOXML schema check | `qa_check.py deck.pptx --integrity` |
| Layout metrics (pre-render) | `geometry_report.py deck.pptx` |
| Visual QA | `render_slides.py deck.pptx --grid` |
| Revision regression | `visual_regress.py qa-baseline/ qa-current/` |
| Brand palette from URL | `brand_kit.py <domain> --name <n>` |
| Render mermaid/graphviz diagram | `render_diagram.py src.mmd` (auto during build) |
| Pre-read handout | `gen_handout.py outline.md` |
| Appendix skeleton | `gen_appendix.py outline.md` |
| Add slide to existing | `add_slide.py in.pptx out.pptx --layout N` |
| Split / merge decks | `edit_deck.py extract` / `edit_deck.py append` |
| Reorder unpacked slides | `edit_deck.py reorder unpacked/ 3,1,2,4` |
| Clean orphan media | `edit_deck.py clean unpacked/` |
