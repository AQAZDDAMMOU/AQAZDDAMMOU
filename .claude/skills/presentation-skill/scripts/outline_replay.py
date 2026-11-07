#!/usr/bin/env python3
"""
outline_replay.py — the one parse -> expand -> profile/density -> data ->
evidence -> agenda -> references -> scale-groups sequence shared by
build_deck.build(), diff_deck.diff_outline, qa_checklist.build_checklist,
gen_handout.gen_handout, patch_deck and refresh_deck.

Two process-globals are installed as a side effect, on purpose:
`builders.set_density()` (the deck's density, which the profile layout repick
and every consumer's truncation caps read) — `helpers.set_verbatim()` stays
build()'s job.  tests/conftest.py resets both after every test.

Each caller used to hand-reimplement this sequence, which let them drift out
of step (slide numbering that no longer matched build_deck.build(), a
transform one caller forgot to apply). replay() is the single source of
truth; callers vary only in which pieces of ReplayResult they use and how
they report .errors/.warnings.
"""
from dataclasses import dataclass, field
from pathlib import Path

from outline_data import load_data_files
from outline_parse import parse_outline
from outline_transforms import (apply_agenda_columns, apply_auto_agenda,
                                apply_references, apply_scale_groups)
import builders
from evidence import apply_evidence, load_for_outline, validate_evidence
from outline_expand import expand_repeats
from palettes import apply_variant
import profiles
from smart_layout import repick


@dataclass
class ReplayResult:
    meta: dict
    slides: list
    errors: list = field(default_factory=list)     # data-file + evidence + expand errors
    warnings: list = field(default_factory=list)
    ctx: dict = field(default_factory=dict)
    ledger: object = None   # evidence ledger or None
    # Subset of .errors that came from evidence-ledger resolution/validation
    # (malformed **Evidence-File:**, unknown ids, the traceability gate) —
    # callers that print .errors with a data-file-specific message use this
    # to give evidence problems an accurate message instead. Always a subset
    # of .errors; not a separate error channel.
    evidence_errors: list = field(default_factory=list)
    # Resolved build knobs (CLI > variant preset > front-matter > profile):
    # the palette key build() should style with (None = the default palette)
    # and the density replay installed via builders.set_density().
    palette_key: str | None = None
    density: str = ""


def make_ctx(outline_path, assets_dir=None, template_path=None, **extra):
    """Build the ctx dict replay()/load_data_files()/evidence resolution need.

    Optional CLI knobs replay() honours when present: `profile` (--profile),
    `palette`, `density`, `variant` (each outranks the outline's own
    front-matter, exactly as build() applies them) and `verbatim`."""
    return {
        "outline_dir": Path(outline_path).parent,
        "assets_dir": Path(assets_dir) if assets_dir else Path("assets"),
        "template_path": template_path,
        **extra,
    }


def replay(outline_path, *, assets_dir=None, ctx=None, md_text=None):
    """Replay parse -> data -> evidence -> agenda -> references -> scale-groups.

    outline_path may be None only when both md_text and ctx are given
    (gen_handout's in-memory path, which has no outline file to resolve a
    directory from — ctx already carries outline_dir).
    """
    if ctx is None:
        ctx = make_ctx(outline_path, assets_dir=assets_dir)
    if md_text is None:
        md_text = Path(outline_path).read_text(encoding="utf-8")

    # verbatim: a caller (--verbatim, diff_deck) puts True/False in ctx to
    # force the mode; None/absent lets parse_outline detect the outline's own
    # `**Copy:** locked` front matter.
    meta, slides = parse_outline(md_text, verbatim=ctx.get("verbatim"))
    slides, expand_errors, expand_warnings = expand_repeats(slides, ctx)

    # Engagement profile: deck-meta defaults below the front-matter and the
    # resolved profile in ctx["_profile"] (validate() and the layout repick
    # read it there).  Then the density — CLI flag > variant preset >
    # front-matter/profile > default — is installed process-wide exactly as
    # build() does, because the layout repick below is capacity-checked
    # against it and every consumer's truncation caps
    # (diff_deck.LAYOUT_CONTENT) read builders.density() too.
    meta, profile_errors = profiles.apply(meta, ctx)
    variant_key = ctx.get("variant") or meta.get("variant", "")
    palette_key, density = apply_variant(variant_key, ctx.get("palette"),
                                         ctx.get("density"))
    density = density or meta.get("density", "")
    builders.set_density(density)

    data_errors, data_warnings = load_data_files(slides, ctx)
    data_errors = expand_errors + profile_errors + data_errors
    data_warnings = expand_warnings + data_warnings

    # Evidence ledger: validate the gate on authored slides (before source
    # synthesis), then flow citations into footnotes and notes.
    ledger, ev_errors = load_for_outline(meta, ctx)
    if ledger is not None and not ev_errors:
        gate_errors, ev_warnings = validate_evidence(slides, ledger)
        ev_errors.extend(gate_errors)
        data_warnings.extend(ev_warnings)
        apply_evidence(slides, ledger)
    evidence_errors = list(ev_errors)
    data_errors.extend(ev_errors)

    slides = apply_auto_agenda(meta, slides)
    slides = apply_agenda_columns(meta, slides, ctx)
    slides = apply_references(meta, slides)
    slides = apply_scale_groups(meta, slides)
    # Profile layout repick last, so its "Slide N" warnings use the same
    # numbering as validate() (after the agenda/sources slides are inserted);
    # it only reads bullets/cards, so nothing above depends on it.
    data_warnings.extend(
        repick(slides, ctx["_profile"].get("preferred_layouts")))

    return ReplayResult(meta=meta, slides=slides, errors=data_errors,
                        warnings=data_warnings, ctx=ctx, ledger=ledger,
                        evidence_errors=evidence_errors,
                        palette_key=palette_key, density=density)
