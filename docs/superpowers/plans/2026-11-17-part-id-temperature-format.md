# Part-ID Temperature Loose Format Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Relax the Part-ID/MIR temperature tag grammar (units, spaces, brackets, degree sign, no-unit `_T25_` form) and convert Kelvin values to Celsius everywhere, with the default `_` separator as a module global.

**Architecture:** One shared unit-alternation sub-pattern plus two module-level compiled regexes (`TEMPERATURE_RE` for Part IDs with a no-unit lookahead branch, `_TEMPERATURE_TEXT_RE` for free text), both built once at import from `PART_ID_SEPARATOR`. A single `_to_celsius` converter applies the Kelvin shift; `parse_temperature` and `parse_temperature_text` route their captures through it.

**Tech Stack:** Python 3.10+ stdlib `re`, pytest.

**Spec:** `docs/superpowers/specs/2026-11-17-part-id-temperature-format-design.md`

## Global Constraints

- Patterns compile exactly once at module import; parse functions run one scan pass (no per-call compilation, no intermediate lists) — user performance ruling.
- Kelvin conversion: `round(value - 273.15, 2)` in both parsers.
- No function-signature changes: all parsers stay `(value) -> float | None`.
- Existing test cases keep passing unchanged.
- No new dependencies; GPL header and file style untouched.

## Review Focus

1. `SN541_T255X` must NOT parse (no-unit form needs `_` or end-of-string) — pinned in Task 1 Step 1.
2. Kelvin float noise must not leak into labels: `format_temperature_label(parse_temperature("SN541_T298K")) == "24.85"` (not `24.850000000000023`) — pinned in Task 1 Step 1.
3. `"SN541_T25C_T298K"` — last match wins even across mixed units → `24.85` — pinned in Task 1 Step 1.
4. Free-text `""`, `"N/A"`, bare `"25"` must keep today's behavior (None / None / 25.0) — pinned in Task 2 Step 1 (existing cases).
5. Composite free text `"#2 → -40C"` must never read the `#2` index — pinned in Task 2 Step 1 (existing case).

---

### Task 1: Part-ID tag parsing (separator global, loose `TEMPERATURE_RE`, Kelvin converter)

**Files:**
- Modify: `deps/exporter/helpers.py:45-77` (constants block + `parse_temperature`), `deps/exporter/helpers.py:130-142` (`part_id_base` docstring only)
- Test: `tests/exporter/test_helpers.py:53-78` (`TestParseTemperature`), `tests/exporter/test_helpers.py:114-127` (`TestPartIdBase`)

**Interfaces:**
- Consumes: nothing new (stdlib `re`).
- Produces: `PART_ID_SEPARATOR = "_"` (module global, `str`); `TEMPERATURE_RE` with named groups `temp` (str) and `unit` (str | absent → `None` via `group()`); `_to_celsius(value: float, unit: str | None) -> float`; `parse_temperature(part_id: str | None) -> float | None` (unchanged signature). Task 2 consumes `_TEMP_UNIT`, `_SEP`, `_to_celsius`.

- [ ] **Step 1: Write the failing tests**

In `TestParseTemperature.test_spec_cases` parametrize, add rows (existing rows stay):

```python
("SN541_T298K", 24.85),                      # Kelvin -> Celsius
("SN541_T298Kelvin", 24.85),
("SN541_T298[Kelvin]", 24.85),
("SN541_T298_K", 24.85),                     # separator before unit
("SN541_T25K", -248.15),
("SN541_T25_KELVIN", -248.15),
("SN541_T25 °K", -248.15),
("SN541_T25_", 25.0),                        # no unit, trailing separator
("SN541_T25_DIE1", 25.0),                    # no unit, separator + suffix
("SN541_T25", 25.0),                         # no unit, end of string
("SN541_T25 C", 25.0),                       # spaces around unit
("SN541_T 25C", 25.0),                       # space after T
("SN541_T25°C", 25.0),                       # degree sign
("SN541_T25°", 25.0),                        # bare degree sign
("SN541_T25_C", 25.0),                       # separator before unit
("SN541_T25_CELSIUS", 25.0),
("SN541_T25Celsius", 25.0),
("SN541_T25[CELSIUS]", 25.0),
("SN541_T25[C]", 25.0),
("SN541_T+25C", 25.0),                       # plus sign
("SN541_T255X", None),                       # anti-goal: junk suffix, no tag
```

Add two test methods to `TestParseTemperature`:

```python
def test_multi_match_mixed_units_last_wins(self):
    assert parse_temperature("SN541_T25C_T298K") == 24.85

def test_kelvin_label_free_of_float_noise(self):
    assert format_temperature_label(parse_temperature("SN541_T298K")) == "24.85"
```

In `TestPartIdBase.test_spec_cases` parametrize, add:

```python
("SN541_T25_DIE1", "SN541_DIE1"),   # separator survives the strip
("SN541_T25_", "SN541_"),           # end-of-string separator stays (documented)
("SN541_T25_CELSIUS", "SN541"),
("SN541_T25°C", "SN541"),
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/exporter/test_helpers.py::TestParseTemperature tests/exporter/test_helpers.py::TestPartIdBase -v`
Expected: FAIL on the new rows (old rows pass).

- [ ] **Step 3: Implement in `deps/exporter/helpers.py`**

Replace the `TEMPERATURE_RE` block (helpers.py:44-49) with the global + shared unit sub-pattern + compiled Part-ID pattern, exactly:

```python
# Default Part-ID separator. The temperature tag patterns below are built
# from it once at import — change this constant to retarget both.
PART_ID_SEPARATOR = "_"
_ESCAPED_SEPARATOR = re.escape(PART_ID_SEPARATOR)

# Loose temperature unit after the number, longest-first, case-insensitive
# (applied by the compiling flags): words, bracketed words, bracketed
# letters, degree sign, bare letters.
_TEMP_UNIT = (
    r"\[?\s*°?\s*(?:CELSIUS|KELVIN)\s*\]?"
    r"|\[\s*°?\s*[CK]\s*\]"
    r"|°\s*[CK]?"
    r"|[CK]"
)

# Part-ID temperature tag, e.g. the "_T25C" in "SN541_T25C": optional sign
# and decimal, spaces allowed around the number and unit, unit optional
# (word / [word] / [C] / [K] / ° / °C / °K / C / K, optionally after a
# separator); without a unit the tag must be followed by the separator or
# end of string ("_T25_", "_T25_DIE1", "_T25"). Case-insensitive; signed
# so cold temperatures ("_T-5.5C") match.
TEMPERATURE_RE = re.compile(
    rf"{_ESCAPED_SEPARATOR}T\s*(?P<temp>[+-]?\d+(?:\.\d+)?)\s*"
    rf"(?:{_ESCAPED_SEPARATOR}?\s*(?P<unit>{_TEMP_UNIT})"
    rf"|(?=\s*{_ESCAPED_SEPARATOR}|\s*$))",
    re.IGNORECASE,
)

# Kelvin unit tokens, normalized (strip brackets/degree/spaces, upper).
_KELVIN_UNITS = frozenset(("K", "KELVIN"))
_KELVIN_OFFSET = 273.15
```

Add below it:

```python
def _to_celsius(value: float, unit: str | None) -> float:
    """Value as Celsius: Kelvin units shift by 273.15, rounded to 2
    decimals so float noise never reaches a sheet label; Celsius units and
    an absent unit pass through."""
    if unit is None:
        return value
    if unit.strip("[]° \t").upper() in _KELVIN_UNITS:
        return round(value - _KELVIN_OFFSET, 2)
    return value
```

Rewrite `parse_temperature` (single `finditer` pass, last match wins):

```python
def parse_temperature(part_id: str | None) -> float | None:
    """Return the temperature encoded in a part ID as Celsius, last match
    wins. Accepts the loose unit grammar of ``TEMPERATURE_RE`` (C/K words,
    brackets, degree sign, spaces, or no unit before the separator/end) and
    converts Kelvin tags (``"_T298K"`` -> ``24.85``). ``None`` when the ID
    is ``None`` or carries no tag."""
    if part_id is None:
        return None
    last = None
    for match in TEMPERATURE_RE.finditer(part_id):
        last = match
    if last is None:
        return None
    return _to_celsius(float(last.group("temp")), last.group("unit"))
```

`part_id_base` body is unchanged (same regex, span strip); refresh its docstring to say the strip removes the tag including a matched unit, and the separator survives the no-unit lookahead form.

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/exporter/test_helpers.py -v`
Expected: PASS (whole file — confirms no regression on the untouched tests).

Run: `python -m pytest tests/exporter -q`
Expected: PASS (matrix/summary consumers unchanged).

- [ ] **Step 5: Commit**

```bash
git add deps/exporter/helpers.py tests/exporter/test_helpers.py
git commit -m "feat(exporter): loose Part-ID temperature tag grammar with Kelvin conversion"
```

### Task 2: `parse_temperature_text` adopts the loose grammar + Kelvin conversion

**Files:**
- Modify: `deps/exporter/helpers.py` (`parse_temperature_text` and the pattern block from Task 1)
- Test: `tests/exporter/test_helpers.py:80-111` (`TestParseTemperatureText`)

**Interfaces:**
- Consumes: `_TEMP_UNIT`, `_ESCAPED_SEPARATOR`, `_to_celsius` (Task 1).
- Produces: `_TEMPERATURE_TEXT_RE` (module-level compiled `re.Pattern`, named groups `temp`/`unit`); `parse_temperature_text(value) -> float | None` (unchanged signature; consumer `deps/exporter/worksheets/file_info.py:118` unaffected).

- [ ] **Step 1: Write the failing tests**

In `TestParseTemperatureText.test_spec_cases` parametrize, add rows (existing rows stay):

```python
("298 K", 24.85),          # Kelvin -> Celsius
("25 kelvin", -248.15),
("25 [KELVIN]", -248.15),
("25°", 25.0),             # bare degree sign
("25 °C", 25.0),           # existing case already; keep and add:
("25 °K", -248.15),
```

(If `"25 °C"` is already present, do not duplicate it.)

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/exporter/test_helpers.py::TestParseTemperatureText -v`
Expected: FAIL on the new rows only.

- [ ] **Step 3: Implement in `deps/exporter/helpers.py`**

Add the compiled free-text pattern next to `TEMPERATURE_RE` (Task 1 block), then rewrite `parse_temperature_text`:

```python
# Free-text temperature (MIR TST_TEMP): first number, optional loose unit,
# no separator/anchor requirement — "25", "25C", "25.5 °C", "298 K".
_TEMPERATURE_TEXT_RE = re.compile(
    rf"(?P<temp>[+-]?\d+(?:\.\d+)?)\s*{_ESCAPED_SEPARATOR}?\s*"
    rf"(?P<unit>{_TEMP_UNIT})?",
    re.IGNORECASE,
)
```

```python
def parse_temperature_text(value) -> float | None:
    """Temperature as a float in Celsius from a File Info "Test
    Temperature:" cell. MIR TST_TEMP is free text: the first number wins,
    an adjacent loose unit is honored (``"298 K"`` -> ``24.85``), a bare
    number parses as-is. A multi-subfile group cell collapses per
    formatFileGroupValues into ``"#1 → 25C\\n#2 → -40C"`` — the first
    subfile's line is taken. ``None`` when the value is absent or carries
    no number (blank cell, "N/A")."""
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return None if math.isnan(value) else float(value)  # NaN -> None
    text = str(value).strip()
    if "\n" in text:
        # multi-subfile composite ("#1 → 25C\n#2 → -40C"): first subfile wins
        text = text.split("\n", 1)[0]
    if text.startswith("#") and "→" in text:
        # drop the "#n →" source prefix so its index never parses as a number
        text = text.split("→", 1)[1]
    match = _TEMPERATURE_TEXT_RE.search(text)
    if match is None:
        return None
    return _to_celsius(float(match.group("temp")), match.group("unit"))
```

(`_FIRST_FLOAT_RE` keeps its other users; no other changes.)

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/exporter -q`
Expected: PASS (including `worksheets/file_info.py` consumers).

- [ ] **Step 5: Commit**

```bash
git add deps/exporter/helpers.py tests/exporter/test_helpers.py
git commit -m "feat(exporter): loose unit grammar + Kelvin conversion in parse_temperature_text"
```
