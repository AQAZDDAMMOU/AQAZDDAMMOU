"""Evidence ledger: per-deck source traceability for material numbers.

A deck declares `**Evidence-File:** evidence.json` in front-matter; slides
reference ledger entries with `- Evidence: E1, E2`. Each ledger entry maps an
id to a claim's value and its source:

    {
      "E1": {"value": "42%", "source": "IDC MarketScape 2026", "page": "12"},
      "E2": {"value": "$4.2B", "source": "Company 10-K FY25",
             "url": "https://..."}
    }

`source` is required; `value` (the number as it appears on the slide),
`page`, `url`, and `claim` are optional. With a ledger active the
unsourced-stat check escalates from warning to a fail-fast error: every
slide with material statistics must carry `- Source:` or `- Evidence:`.
Referenced entries flow into the slide's source footnote (when no explicit
`- Source:` is given) and full citations are appended to speaker notes.
diff_deck.py verifies referenced ledger values survive into the built deck.
"""
import json
import re
from pathlib import Path

from narrative import STAT_RX

REQUIRED_ENTRY_KEYS = {"source"}
KNOWN_ENTRY_KEYS = {"source", "value", "page", "url", "claim"}
_ID_RX = re.compile(r"^[A-Za-z0-9_.-]+$")


def resolve_evidence_file(spec, ctx):
    """Resolve an Evidence-File path: outline dir, then assets dir, then
    as-is (mirrors Data-File resolution)."""
    candidates = [Path(ctx.get("outline_dir", ".")) / spec,
                  Path(ctx.get("assets_dir", "assets")) / spec, Path(spec)]
    for c in candidates:
        if c.is_file():
            return c
    return None


def load_ledger(path):
    """Parse and validate an evidence JSON file.

    Returns (ledger dict, errors). A non-empty errors list means the ledger
    is unusable — callers must fail fast (no .pptx written).
    """
    errors = []
    try:
        raw = json.loads(Path(path).read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as e:
        return {}, [f"evidence file unreadable ({e})"]
    if not isinstance(raw, dict) or not raw:
        return {}, ["evidence file must be a non-empty JSON object "
                    "mapping ids to entries"]
    ledger = {}
    for eid, entry in raw.items():
        where = f"evidence entry {eid!r}"
        if not _ID_RX.match(str(eid)):
            errors.append(f"{where}: id must be alphanumeric/_-. only")
            continue
        if not isinstance(entry, dict):
            errors.append(f"{where}: must be an object with a 'source' key")
            continue
        missing = REQUIRED_ENTRY_KEYS - set(entry)
        if missing:
            errors.append(f"{where}: missing required key(s) "
                          f"{sorted(missing)}")
            continue
        unknown = set(entry) - KNOWN_ENTRY_KEYS
        if unknown:
            errors.append(f"{where}: unknown key(s) {sorted(unknown)} "
                          f"(supported: {', '.join(sorted(KNOWN_ENTRY_KEYS))})")
            continue
        bad = [k for k, v in entry.items() if not isinstance(v, str)]
        if bad:
            errors.append(f"{where}: value(s) for {sorted(bad)} must be "
                          "strings")
            continue
        ledger[str(eid)] = entry
    return ledger, errors


def parse_refs(spec):
    """'E1, E2' -> ['E1', 'E2'] (order kept, blanks dropped)."""
    return [t.strip() for t in (spec or "").split(",") if t.strip()]


def _slide_text_blob(p):
    return " ".join([
        p.get("heading", ""), p.get("title", ""), p.get("notes", ""),
        p.get("kicker", ""), p.get("callout", ""), p.get("value", ""),
        " ".join(p.get("bullets", [])),
        " ".join(str(v) for s in p.get("stats", []) for v in s.values()),
        " ".join(c for row in p.get("table_rows", []) for c in row),
    ])


def validate_evidence(slides, ledger):
    """Ledger-vs-outline checks. Returns (errors, warnings).

    Errors: unknown ids; slides with material statistics but neither
    `- Source:` nor `- Evidence:` (the traceability gate — active only
    because a ledger is declared). Warnings: referenced values not found in
    the slide's text; ledger entries no slide references.
    """
    errors, warnings = [], []
    used = set()
    for n, p in enumerate(slides, 1):
        where = (f"Slide {n} (line {p['_line']})" if p.get("_line")
                 else f"Slide {n}")
        refs = parse_refs(p.get("evidence"))
        unknown = [r for r in refs if r not in ledger]
        if unknown:
            errors.append(f"{where}: unknown evidence id(s) "
                          f"{', '.join(unknown)} — not in the ledger")
        used.update(r for r in refs if r in ledger)
        blob = _slide_text_blob(p)
        for r in refs:
            value = ledger.get(r, {}).get("value")
            if value and value not in blob:
                warnings.append(
                    f"{where}: evidence {r} value {value!r} does not appear "
                    "on the slide — cite the number as ledgered")
        if p.get("_auto") or p.get("_appendix"):
            continue
        if not refs and not p.get("source") and STAT_RX.search(blob):
            errors.append(
                f"{where}: material statistics but no '- Source:' or "
                "'- Evidence:' — the evidence ledger requires every "
                "material stat to trace")
    unused = sorted(set(ledger) - used)
    if unused:
        warnings.append(f"evidence entries never referenced: "
                        f"{', '.join(unused)}")
    return errors, warnings


def _citation(eid, entry):
    parts = [entry["source"]]
    if entry.get("page"):
        parts.append(f"p.{entry['page']}")
    if entry.get("url"):
        parts.append(entry["url"])
    return f"{eid} — {', '.join(parts)}"


def apply_evidence(slides, ledger):
    """Flow ledger citations into slides (before build).

    Slides with `- Evidence:` refs get: a synthesized `- Source:` footnote
    from the referenced entries' sources (unless one is already set), and a
    full citation line appended to speaker notes. Unknown ids are skipped —
    validate_evidence reports them. Returns the number of slides touched.
    """
    touched = 0
    for p in slides:
        refs = [r for r in parse_refs(p.get("evidence")) if r in ledger]
        if not refs:
            continue
        touched += 1
        if not p.get("source"):
            seen, sources = set(), []
            for r in refs:
                src = ledger[r]["source"]
                if src not in seen:
                    seen.add(src)
                    sources.append(src)
            p["source"] = "; ".join(sources)
        citations = "Evidence: " + "; ".join(
            _citation(r, ledger[r]) for r in refs)
        p["notes"] = (f"{p['notes']} {citations}" if p.get("notes")
                      else citations)
    return touched


def load_for_outline(meta, ctx):
    """Resolve and load the deck's ledger from parsed outline meta.

    Returns (ledger dict or None when no **Evidence-File:** is declared,
    errors list).
    """
    spec = meta.get("evidence_file")
    if not spec:
        return None, []
    path = resolve_evidence_file(spec, ctx)
    if path is None:
        return None, [f"Evidence-File not found: '{spec}' "
                      "(tried outline dir, assets dir, literal path)"]
    ledger, errors = load_ledger(path)
    return ledger, [f"Evidence-File '{spec}': {e}" for e in errors]
