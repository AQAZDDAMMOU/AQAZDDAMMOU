#!/usr/bin/env python3
"""
numfmt.py — one number-format policy for a whole deck.

Consulting decks are judged on numeric hygiene: a figure must be spelled the
same way on every slide it appears (same decimals, same unit suffix, same
negative convention) and rounded parts must add to the rounded total.  This
module owns all three concerns:

  * **Policy** — `**Number-Format:** units=M decimals=1 negatives=parens`
    front-matter, parsed by `parse_policy` and installed process-wide with
    `set_policy` (a module-level knob, reset by tests/conftest.py).
  * **Formatting** — `fmt_abs`, `fmt_pct`, `fmt_multiple`; `NM` for
    non-meaningful values (None, or a division by zero).
  * **Consistent rounding** — `round_consistent` (largest-remainder, works
    for signed parts), the engine behind `charts.round_to_sum`.

Negative-sign rendering, precisely:

    no policy declared          -3.2   (plain ASCII hyphen-minus — this is
                                        what the builders have always drawn,
                                        so decks without the directive render
                                        byte-for-byte as before)
    negatives=minus (default)   –3.2   (en dash, the typographic minus)
    negatives=parens            (3.2)

A `units` suffix sits inside the parentheses — `(8M)`, never `(8)M`.

`units` only *appends* a suffix — it never rescales a value.  Builders pass
values already scaled to the unit the author chose in the outline.

Policy boundary (as of cycle 5): `units`, `decimals` and `negatives` reach
every builder value label through `fmt_abs`; `pct` reaches every percentage
label a builder draws — funnel step conversions, the waterfall bracket, the
CAGR arrow and the stacked-100/mekko share labels — through `fmt_pct`
(rounded at `pct_decimals()` so a column's shares still add to 100).  Those
labels are the one place where "no policy" is not quite the old byte string:
`fmt_pct` rounds ties half-up (12.5% -> 13%, not 12%) and groups thousands
(+1,204%), where the hard-coded `f"{x:.0f}%"` did neither.  `units`/`decimals`/
`negatives` also reach *native* chart data labels as a spreadsheet format
code (`chart_number_format`, applied by `charts.add_native_chart`).
`multiples` is honoured by `fmt_multiple` only: the outline parser drops the
`x` suffix when it reads a data point, so no builder can tell a multiple from
a plain number and nothing calls it yet.

With no policy installed the formatters reproduce the two legacy helpers
exactly: STYLE_PLAIN is the old `_fmt_plain`/`_fmt_label_num` (thousands
separator, one decimal for non-integers) and STYLE_COMPACT is the old
`builders_exhibits._fmt_num` (no decimals at all above 1000).
"""
import re
from decimal import ROUND_HALF_UP, Decimal
from typing import NamedTuple

NM = "NM"                 # non-meaningful (undefined / divide-by-zero)
AUTO = "auto"             # decimals: 0 for integers, 1 otherwise
EN_DASH = "–"

STYLE_PLAIN = "plain"     # 1,500.5 — funnel/mekko/chart labels
STYLE_COMPACT = "compact"  # 1,500   — waterfall/tornado bar labels

UNITS = ("none", "K", "M", "B")
NEGATIVE_STYLES = ("minus", "parens")
MULTIPLE_STYLES = ("1x", "1.0x")
MAX_DECIMALS = 3

POLICY_DEFAULTS = {
    "units": "none",
    "decimals": AUTO,
    "negatives": "minus",
    "multiples": "1x",
    "pct": 0,
}

# magnitude multipliers, for comparing "14.2M" with "14,200,000"
UNIT_SCALE = {"": 1.0, "K": 1e3, "M": 1e6, "B": 1e9}

_POLICY = None  # process-global; see set_policy


# ── policy ──────────────────────────────────────────────────────────────────
def _parse_int(key, raw, lo, hi):
    try:
        value = int(raw)
    except ValueError:
        raise ValueError(
            f"Number-Format {key}={raw!r} is not an integer") from None
    if not lo <= value <= hi:
        raise ValueError(f"Number-Format {key}={raw!r} out of range {lo}-{hi}")
    return value


def parse_policy(spec):
    """Parse a `**Number-Format:**` spec into a policy dict (None if empty).

    spec: `units=M decimals=1 negatives=parens multiples=1x pct=1`.
    Raises ValueError on an unknown key or an out-of-range value.
    """
    if not spec or not str(spec).strip():
        return None
    policy = dict(POLICY_DEFAULTS)
    for token in str(spec).split():
        if "=" not in token:
            raise ValueError(
                f"Number-Format expects key=value pairs, got {token!r}")
        key, _, raw = token.partition("=")
        key, raw = key.strip(), raw.strip()
        if key not in POLICY_DEFAULTS:
            raise ValueError(
                f"unknown Number-Format key {key!r} "
                f"(expected one of {', '.join(sorted(POLICY_DEFAULTS))})")
        if key == "units":
            if raw not in UNITS:
                raise ValueError(f"Number-Format units={raw!r} "
                                 f"(expected one of {', '.join(UNITS)})")
            policy[key] = raw
        elif key == "negatives":
            if raw not in NEGATIVE_STYLES:
                raise ValueError(
                    f"Number-Format negatives={raw!r} "
                    f"(expected one of {', '.join(NEGATIVE_STYLES)})")
            policy[key] = raw
        elif key == "multiples":
            if raw not in MULTIPLE_STYLES:
                raise ValueError(
                    f"Number-Format multiples={raw!r} "
                    f"(expected one of {', '.join(MULTIPLE_STYLES)})")
            policy[key] = raw
        elif key == "decimals":
            policy[key] = (AUTO if raw == AUTO
                           else _parse_int(key, raw, 0, MAX_DECIMALS))
        else:  # pct
            policy[key] = _parse_int(key, raw, 0, MAX_DECIMALS)
    return policy


def set_policy(policy):
    """Install (or clear, with None) the deck-wide policy. Process-global."""
    global _POLICY
    if policy is not None and not isinstance(policy, dict):
        raise ValueError("number-format policy must be a dict or None")
    _POLICY = dict(policy) if policy else None
    return _POLICY


def get_policy():
    """The installed policy dict, or None when no policy was declared."""
    return _POLICY


# ── rounding ────────────────────────────────────────────────────────────────
def round_half_up(value, decimals=0):
    """Round ties away from zero (what readers expect; float round() does not).

    Returns an int for decimals=0 so labels never render '-0.0'.
    """
    quant = Decimal(1).scaleb(-decimals)
    out = Decimal(repr(float(value))).quantize(quant, rounding=ROUND_HALF_UP)
    return int(out) if decimals <= 0 else float(out)


def round_consistent(parts, total, decimals=0):
    """Round `parts` so they sum to exactly `round(total)` (largest remainder).

    Each part is rounded half-up; the shortfall (or excess) against the
    rounded total is then given to the parts with the largest absolute
    remainders, one unit of the last decimal each.  Signed parts are handled:
    a positive shortfall bumps the parts that were rounded *down* the most.

    `total=None` — or an all-zero `parts` — means "no total to honour": every
    part is simply rounded half-up.
    """
    if not parts:
        return []
    factor = 10 ** decimals
    scaled = [float(v) * factor for v in parts]
    rounded = [round_half_up(s) for s in scaled]
    if total is None or not any(parts):
        return [r / factor if decimals else r for r in rounded]

    diff = round_half_up(float(total) * factor) - sum(rounded)
    if diff:
        # remainder in [-0.5, 0.5]: how far each part was rounded down (+) or up (-)
        rest = [scaled[i] - rounded[i] for i in range(len(parts))]
        order = sorted(range(len(parts)), key=lambda i: (rest[i], i),
                       reverse=diff > 0)
        step = 1 if diff > 0 else -1
        for i in order[:abs(diff)]:
            rounded[i] += step
    return [r / factor if decimals else r for r in rounded]


# ── formatting ──────────────────────────────────────────────────────────────
def _legacy(value, signed, style):
    """Byte-for-byte reproduction of the pre-policy builder formatters."""
    sign = "+" if signed else ""
    if style == STYLE_COMPACT:
        if abs(value) >= 1000:
            return f"{value:{sign},.0f}"
        if value != int(value):
            return f"{value:{sign}.1f}"
        return f"{int(value):{sign}d}"
    if value == int(value):
        return f"{value:{sign},.0f}"
    return f"{value:{sign},.1f}"


def _decimals_for(value, policy):
    dec = policy["decimals"]
    if dec != AUTO:
        return dec
    return 0 if float(value) == int(value) else 1


def label_decimals(values, style=STYLE_PLAIN):
    """Decimals `fmt_abs` will print for a *group* of values.

    Rounding a set of labels consistently (round_consistent) and printing them
    (fmt_abs) must agree on precision, so both sides call this.  With a policy
    installed the policy's `decimals` wins (`auto` falls back to "0 when every
    value is whole, else 1"); without one, each legacy style keeps its own
    rule.  STYLE_COMPACT prints no decimals at all above 1000, so a mixed set
    like [1500.5, 3.5] still prints one label at a coarser precision than it
    was rounded at — an edge case the compact style has always had.
    """
    values = [float(v) for v in values]
    whole = all(v == int(v) for v in values)
    policy = get_policy()
    if policy is not None:
        dec = policy["decimals"]
        return (0 if whole else 1) if dec == AUTO else dec
    if style == STYLE_COMPACT:
        return 0 if all(v == int(v) or abs(v) >= 1000 for v in values) else 1
    return 0 if whole else 1


def _wrap_sign(body, negative, signed, policy):
    if negative:
        if policy["negatives"] == "parens":
            return f"({body})"
        return f"{EN_DASH}{body}"
    return f"+{body}" if signed else body


def fmt_abs(value, signed=False, style=STYLE_PLAIN):
    """Format an absolute quantity under the deck policy.

    signed: prefix non-negative values with '+' (delta labels).
    style:  STYLE_PLAIN or STYLE_COMPACT — only consulted when no policy is
            installed, where each reproduces its legacy builder formatter.
    """
    if value is None:
        return NM
    policy = get_policy()
    if policy is None:
        return _legacy(value, signed, style)
    dec = _decimals_for(value, policy)
    magnitude = round_half_up(abs(float(value)), dec)
    suffix = "" if policy["units"] == "none" else policy["units"]
    return _wrap_sign(f"{magnitude:,.{dec}f}{suffix}",
                      float(value) < 0, signed, policy)


def fmt_pct(value, signed=False, legacy_decimals=0):
    """Format a percentage already expressed in percent units (12 -> '12%').

    legacy_decimals: precision to use *only when no policy is installed*, so a
    caller whose hard-coded format printed a decimal (the CAGR arrow's
    `{cagr:+.1%}`) keeps printing one.  The policy's `pct` always wins.
    """
    if value is None:
        return NM
    policy = get_policy()
    dec = policy["pct"] if policy else legacy_decimals
    magnitude = round_half_up(abs(float(value)), dec)
    body = f"{magnitude:,.{dec}f}%"
    if float(value) >= 0:
        return f"+{body}" if signed else body
    if policy is None:
        return f"-{body}"
    return _wrap_sign(body, True, signed, policy)


def pct_decimals(default=0):
    """Decimals a percentage label prints under the installed policy.

    `default` is what no policy means for this caller — the mirror of
    `fmt_pct(..., legacy_decimals=...)`, which it must match.

    Share labels are rounded largest-remainder (`charts.round_to_sum`) before
    `fmt_pct` prints them, so the rounding and the printing must agree on
    precision — both read it here.
    """
    policy = get_policy()
    return policy["pct"] if policy else default


def fmt_multiple(value):
    """Format a multiple: 11.234 -> '11.2x' ('11.0x' under multiples=1.0x)."""
    if value is None:
        return NM
    policy = get_policy()
    style = policy["multiples"] if policy else POLICY_DEFAULTS["multiples"]
    body = f"{round_half_up(abs(float(value)), 1):,.1f}"
    if style == "1x" and body.endswith(".0"):
        body = body[:-2]
    if float(value) < 0:
        return _wrap_sign(f"{body}x", True, False, policy or POLICY_DEFAULTS) \
            if policy else f"-{body}x"
    return f"{body}x"


# ── native chart number format ──────────────────────────────────────────────
def chart_number_format(policy):
    """The spreadsheet format code a native chart's data labels should use.

    Returns None when no policy is declared, so charts.py leaves the chart XML
    exactly as python-pptx wrote it (no `c:dLbls` element at all).  Otherwise
    the code mirrors what `fmt_abs` draws on shape labels: a thousands
    separator, the policy's decimals (`auto` -> an optional digit), the unit
    suffix, and a negative section spelling negatives the policy's way —
    `(1.2M)` under parens, an en dash under minus.
    """
    if not policy:
        return None
    dec = policy["decimals"]
    if dec == AUTO:
        body = "#,##0.#"
    elif dec == 0:
        body = "#,##0"
    else:
        body = "#,##0." + "0" * dec
    if policy["units"] != "none":
        body += f'"{policy["units"]}"'
    if policy["negatives"] == "parens":
        return f"{body};({body})"
    return f'{body};"{EN_DASH}"{body}'


# ── token parsing ───────────────────────────────────────────────────────────
class NumToken(NamedTuple):
    """One number as it is written on a slide."""
    value: float          # signed, as written (14.2 for "$14.2M")
    unit: str             # normalised magnitude: '' | K | M | B | bps | pp
    kind: str             # abs | pct | multiple | currency
    decimals: int         # digits after the decimal point, as written
    neg_style: str        # parens | minus | dash | None
    raw: str              # the matched text
    suffix: str = ""      # the magnitude suffix as written ('MM', 'bn', …)


# suffix spelling -> normalised unit
UNIT_SPELLINGS = {
    "k": "K", "K": "K", "thousand": "K", "thousands": "K",
    "m": "M", "M": "M", "mm": "M", "MM": "M", "mn": "M", "Mn": "M",
    "million": "M", "millions": "M",
    "b": "B", "B": "B", "bn": "B", "Bn": "B", "BN": "B",
    "billion": "B", "billions": "B",
}
MEASURE_SUFFIXES = ("bps", "pp")
CURRENCY = "$€£¥"
MINUS_CHARS = "-‐‑‒–—―−"
YEAR_MIN, YEAR_MAX = 1900, 2099
PERIOD_PREFIX_RX = re.compile(r"(?i)(?:^|[^A-Za-z])(fy|cy|q|h)$")

_TOKEN_RX = re.compile(r"""
    (?:(?P<open>\()[ ]*)?
    (?:(?P<cur>[$€£¥])[ ]*)?
    (?:(?P<sign>[""" + MINUS_CHARS + r"""+])[ ]*)?
    (?P<num>\d{1,3}(?:,\d{3})+(?:\.\d+)?|\d+(?:\.\d+)?|\.\d+)
    (?:
        [ ]?(?P<pct>%)
      | (?P<suf>bps|pp|MM|mn|bn|Bn|BN|[KkMmBbxX])(?![A-Za-z0-9])
      | [ ](?P<word>million|millions|billion|billions|thousand|thousands)\b
    )?
    (?:[ ]*(?P<close>\)))?
""", re.VERBOSE)


def _neg_style(match, text):
    """parens | minus | dash | None — None when the '-' is a range/hyphen."""
    if match.group("open") and match.group("close"):
        return "parens"
    sign = match.group("sign")
    if not sign or sign == "+":
        return None
    before = text[:match.start("sign")].rstrip(" ")
    if before and (before[-1].isalnum() or before[-1] in "/:"):
        return None  # '2020-2024', 'A-3' — a range or a hyphenation, not a sign
    return "minus" if sign == "-" else "dash"


def _skip(match, text, suffix, has_pct):
    """True when the match is a year, a period label, or glued to a word."""
    tail = text[match.end():match.end() + 1]
    if tail.isalnum():
        return True
    head = text[:match.start()]
    if head and (head[-1].isalnum() or head[-1] in "/:"):
        return True  # '3x3', 'H100', '1/2' — glued to what precedes it
    if PERIOD_PREFIX_RX.search(head):
        return True
    if suffix or has_pct or match.group("cur"):
        return False
    num = match.group("num")
    if "." not in num and "," not in num and len(num) == 4 \
            and YEAR_MIN <= int(num) <= YEAR_MAX:
        return True
    return False


def find_tokens(text):
    """Every number written in `text`, skipping years and FY/Q period labels."""
    tokens = []
    if not text:
        return tokens
    for match in _TOKEN_RX.finditer(text):
        raw_suffix = match.group("suf") or match.group("word") or ""
        has_pct = bool(match.group("pct"))
        if _skip(match, text, raw_suffix, has_pct):
            continue
        num = match.group("num")
        value = float(num.replace(",", ""))
        neg = _neg_style(match, text)
        if neg:
            value = -value
        low = raw_suffix.lower()
        if low in ("x",):
            kind, unit = "multiple", ""
        elif has_pct:
            kind, unit = "pct", ""
        elif low in MEASURE_SUFFIXES:
            kind, unit = "abs", low
        elif match.group("cur"):
            kind, unit = "currency", UNIT_SPELLINGS.get(raw_suffix, "")
        else:
            kind, unit = "abs", UNIT_SPELLINGS.get(raw_suffix, "")
        decimals = len(num.partition(".")[2])
        start = match.start("open") if match.group("open") is not None \
            else match.start("cur") if match.group("cur") is not None \
            else match.start()
        raw = text[start:match.end()].strip()
        tokens.append(NumToken(value, unit, kind, decimals, neg, raw,
                               raw_suffix))
    return tokens


def parse_token(text):
    """The single NumToken `text` consists of, or None if it is not one."""
    if text is None:
        return None
    stripped = text.strip()
    if not stripped:
        return None
    tokens = find_tokens(stripped)
    if len(tokens) != 1 or tokens[0].raw != stripped:
        return None
    return tokens[0]


def base_value(token):
    """A token's value in base units — '14.2M' and '14,200,000' both 1.42e7."""
    if token is None:
        return None
    return token.value * UNIT_SCALE.get(token.unit, 1.0)
