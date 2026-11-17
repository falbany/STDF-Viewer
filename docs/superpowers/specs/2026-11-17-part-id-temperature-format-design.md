# Part-ID Temperature Tag — Loose Format Parsing

Date: 2026-11-17
Status: approved-pending (short design, small change)
Scope: `deps/exporter/helpers.py` (`TEMPERATURE_RE`, `parse_temperature`,
`parse_temperature_text`, `part_id_base`) + `tests/exporter/test_helpers.py`

## Intent and why

`TEMPERATURE_RE` currently accepts exactly `_T<n>C` (case-insensitive,
optional decimal, signed negative), and `parse_temperature_text` grabs the
first number anywhere in the MIR `TST_TEMP` free text. Real Part IDs and
temperature strings vary: the unit may be spelled out, bracketed, carry a
degree sign, be separated from the number by spaces, or be absent entirely
(`_T25_` when the default `_` separator follows). The goal is to stop
constraining the format more than necessary while keeping per-DUT temperature
resolution (`resolve_dut_temperature`) and the temperature split grouping
correct — including files that tag Kelvin.

## Goals

- **Default separator is a module global**: `PART_ID_SEPARATOR = "_"`; both
  temperature patterns are built from it (formatted into the pattern string,
  `re.escape`d) so a future separator change is a one-line edit.
- **Loose unit grammar** after `_T<number>` (Part ID) or `<number>`
  (free text), case-insensitive throughout, optional spaces between number
  and unit, `+`/`-` signs, optional decimal:
  - bare letters: `C`, `K`
  - words, any case: `CELSIUS`, `Celsius`, `KELVIN`, `Kelvin`, …
  - bracketed: `[C]`, `[K]`, `[CELSIUS]`, `[Celsius]`, `[KELVIN]`, `[Kelvin]`
  - degree sign: `°` alone, `°C`, `°K`, also after a space
  - an optional `_` separator may sit between the number and the unit
    (`_T25_C`, `_T25_CELSIUS`, `_T25_K`)
  - Part-ID only: **no unit at all** when followed by the `_` separator or
    end of string: `_T25_`, `_T25_DIE1`, `_T25`
- **Kelvin converts to Celsius** everywhere: `K - 273.15` rounded to 2
  decimals (`_T298K` -> `24.85`, `"298 K"` -> `24.85`). The split groups key
  on the converted value — a file mixing `25C` and `298K` gets separate
  groups (only exactly-equivalent tags, `298.15K` / `25C`, collapse into
  one) — and the `T{n}°C` sheet label stays physically honest.
- **`parse_temperature_text` adopts the same loose grammar**: number +
  optional unit (first match wins, no separator requirement — free text),
  Kelvin converted, bare numbers still parse. Composite/subfile-prefix
  handling (`#1 → …`) and `None` cases unchanged.
- Last-match-wins (Part ID) and `part_id_base` stripping semantics unchanged.
- **High performance**: both patterns are compiled exactly once at module
  import; the parse functions run a single scan pass (Part ID: one
  `finditer` keeping the last match; free text: one `search`); Kelvin
  detection is a membership test on a frozenset after a single
  `strip/upper` of the unit group — no per-call compilation, no list
  building, no extra regex passes.

## Non-goals

- No configurable separator *setting* — `PART_ID_SEPARATOR` is a code-level
  constant. A UI setting can come later if a file needs another separator.
- No rounding of authentically authored Celsius values.
- Free-text heuristics beyond the unit grammar stay as they are.

## Anti-goals

- No false positives on Part IDs that merely contain `_T<digits>` mid-token
  with trailing junk: the no-unit form requires a following `_` or end of
  string; a unit form must actually match a unit (`SN541_T255X` is NOT a
  25° tag).
- Accepted trade-off (loose free text): `"298 K ohms"`-style prose could
  read as Kelvin. `TST_TEMP` values don't carry such prose in practice;
  looseness was explicitly requested.

## Approach (the one sensible way)

Shared unit sub-pattern + two compiled patterns, all built once from
`PART_ID_SEPARATOR`:

    # unit alternation, longest-first (word before letter)
    _TEMP_UNIT = r"(?:\[?\s*°?\s*(?:CELSIUS|KELVIN)\s*\]?"
                 r"|\[\s*°?\s*[CK]\s*\]|°\s*[CK]?|[CK])"

    TEMPERATURE_RE = re.compile(
        rf"{_SEP}T\s*(?P<temp>[+-]?\d+(?:\.\d+)?)\s*"
        rf"(?:{_SEP}?\s*(?P<unit>{_TEMP_UNIT})|(?=\s*{_SEP}|\s*$))",
        re.IGNORECASE)

    _TEMPERATURE_TEXT_RE = re.compile(
        rf"[+-]?\d+(?:\.\d+)?\s*{_SEP}?\s*(?P<unit>{_TEMP_UNIT})?",
        re.IGNORECASE)

- Part-ID pattern: unit branch consumes an optional `_` + spaces + unit; the
  no-unit branch is a zero-width lookahead so
  `part_id_base("SN541_T25_DIE1")` -> `"SN541_DIE1"` (separator survives)
  while `part_id_base("SN541_T25C")` -> `"SN541"` (unit consumed) — matching
  today's strip behavior for the tagged case.
- Text pattern: unit optional with no separator/anchor requirement —
  `"25"` stays `25.0`, `"25.5 °C"` stays `25.5`, `"298 K"` converts.
- `parse_temperature` switches from `findall` to `finditer` (last match
  wins) and routes the value through a shared converter
  (`_to_celsius(value, unit)`): unit `None` or C/CELSIUS passes through;
  K/KELVIN becomes `round(value - 273.15, 2)`.
- `part_id_base` reuses `TEMPERATURE_RE` unchanged (it strips the whole
  match span).

### Calls I made (check while reading)

- Conversion rounding: `round(value - 273.15, 2)` in BOTH parsers (same
  value from the same tag either way).
- `_T25_C` reads as Celsius `25` (separator before a unit, not "no unit").
- Unit-then-junk (`_T25CAB`) matches as `25` — same as today; base becomes
  `SN541AB`.
- The free-text Kelvin rule keys on the unit token adjacent to the number,
  so `"Temp = 298 K"` converts but `"Bin 298"` (no unit) stays a bare number.
- No function-signature changes: `float | None`, `None` when absent.

## What's left to the builder

- Update the temperature docstrings (`_T<n>C`-only wording, new formats,
  Kelvin conversion).
- TDD: extend `TestParseTemperature` (unit matrix, `_T25_`, `_T25_DIE1`,
  `_T25` end-of-string, Kelvin conversion, `+` sign, spaces),
  `TestPartIdBase` (trailing-separator strip), and `TestParseTemperatureText`
  (`"298 K"` -> `24.85`, `"25 kelvin"`, `"25 °"`, bare `"25"` and every
  existing case unchanged). All existing cases must keep passing.
