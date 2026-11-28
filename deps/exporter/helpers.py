"""
#
# helpers.py - exporter package - STDF Viewer
#
# Author: falbany (fork)
# Created Date: November 9th 2026
# -----
# Last Modified: Mon Nov 9 2026
# Modified By: falbany (fork)
# -----
# Copyright (c) 2026 falbany (fork)
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <https://www.gnu.org/licenses/>.
#

Pure parsing / formula / chunking helpers for the modern xlsxwriter
exporter: the pure spec/letter utilities are engine-neutral, and the
cached-value engine below mirrors the sheet formulas for previews."""

import math
import re
import statistics

from xlsxwriter.utility import xl_col_to_name

from deps.SharedSrc import isPass


def get_column_letter(col: int) -> str:
    """1-based column letter ("A" for 1), wrapping xlsxwriter's 0-based
    ``xl_col_to_name`` — every spec/formula letter in the package stays
    1-based Excel notation regardless of engine index conventions."""
    return xl_col_to_name(col - 1)


# Default Part-ID separator. The temperature tag patterns below are built
# from it once at import — change this constant to retarget both.
PART_ID_SEPARATOR = "_"
_ESCAPED_SEPARATOR = re.escape(PART_ID_SEPARATOR)

# Loose temperature unit after the number, longest-first, case-insensitive
# (applied by the compiling flags): words, bracketed words, bracketed
# letters, degree sign, bare letters.
_TEMP_UNIT = (
    r"\[?\s*+°?\s*+(?:CELSIUS|KELVIN)\s*+\]?"
    r"|\[\s*+°?\s*+[CK]\s*+\]"
    r"|°\s*+[CK]?"
    r"|[CK]"
)

# Part-ID temperature tag, e.g. the "_T25C" in "SN541_T25C": optional sign
# and decimal, spaces allowed around the number and unit, unit optional
# (word / [word] / [C] / [K] / ° / °C / °K / C / K, optionally after a
# separator); without a unit the tag must be followed by the separator or
# end of string ("_T25_", "_T25_DIE1", "_T25"). Case-insensitive; signed
# so cold temperatures ("_T-5.5C") match.
TEMPERATURE_RE = re.compile(
    rf"{_ESCAPED_SEPARATOR}T\s*+(?P<temp>[+-]?\d+(?:\.\d+)?)\s*+"
    rf"(?:{_ESCAPED_SEPARATOR}?\s*+(?P<unit>{_TEMP_UNIT})"
    rf"|(?=\s*+{_ESCAPED_SEPARATOR}|\s*+$))",
    re.IGNORECASE,
)

# Kelvin unit tokens, normalized (strip brackets/degree/spaces, upper).
_KELVIN_UNITS = frozenset(("K", "KELVIN"))
_KELVIN_OFFSET = 273.15

# Free-text temperature (MIR TST_TEMP): first number, optional loose unit,
# no separator/anchor requirement — "25", "25C", "25.5 °C", "298 K".
_TEMPERATURE_TEXT_RE = re.compile(
    rf"(?P<temp>[+-]?\d+(?:\.\d+)?)\s*+{_ESCAPED_SEPARATOR}?\s*+"
    rf"(?P<unit>{_TEMP_UNIT})?",
    re.IGNORECASE,
)

# First integer anywhere in a string (fallback extractor for the DB-shaped
# "N ms" / "Bin N" metadata values below).
_FIRST_INT_RE = re.compile(r"-?\d+")

# Default test/param split: "workflow/rest/of/name/param" -> test + last segment.
DEFAULT_PARAM_PATTERN = re.compile(r"^(?P<test>.+)/(?P<param>[^/]+)$")

# Excel hard limit is 16384 columns per sheet. A Test Summary chunk sheet is
# 13 fixed columns (A..L base block) + n visible die columns + 1 spacer +
# 1 Trend sparkline column + 1 spacer + n hidden flag mirrors, so a chunk
# fits only while 13 + n + 3 + n <= 16384, i.e.
# n <= (16384 - 13 - 3) // 2 = 8184 DUTs per sheet.
DUT_CHUNK_CAP = 8184


def _to_celsius(value: float, unit: str | None) -> float:
    """Value as Celsius: Kelvin units shift by 273.15, rounded to 2
    decimals so float noise never reaches a sheet label; Celsius units and
    an absent unit pass through."""
    if unit is None:
        return value
    if unit.strip("[]° \t").upper() in _KELVIN_UNITS:
        return round(value - _KELVIN_OFFSET, 2)
    return value


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


def parse_temperature_text(value) -> float | None:
    """Temperature as a float in Celsius from a File Info "Test
    Temperature:" cell. MIR TST_TEMP is free text: the first number wins,
    an adjacent loose unit is honored (``"298 K"`` -> ``24.85``), a bare
    number parses as-is. A multi-subfile group cell collapses per
    formatFileGroupValues into ``"#1 → 25C\\n#2 → -40C"`` — the first
    subfile's line is taken. ``None`` when the value is absent or carries
    no number (blank cell, "N/A").
    """
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


def resolve_dut_temperature(part_id: str | None, file_temperature: float | None) -> float | None:
    """One resolver, two consumers (DUT Summary design §7.4): the
    Temperature metadata row and the temperature split group by the SAME
    per-DUT value. A ``_T<n>C`` tag in the Part ID wins; the file's
    recorded MIR TST_TEMP covers untagged IDs; ``None`` when neither
    source carries a number. ``0.0`` is a value, never a fallback
    trigger."""
    tagged = parse_temperature(part_id)
    return tagged if tagged is not None else file_temperature


def format_temperature_label(value: float) -> str:
    """The ``T<n>`` label of a split sheet name: the float with a trimmed
    trailing ``.0`` — ``25.0`` -> ``"25"``, ``125.5`` -> ``"125.5"``,
    ``-55.0`` -> ``"-55"``. Float rendering cannot produce a forbidden
    sheet-name character, and every label stays far under Excel's
    31-character sheet-name cap."""
    if value == int(value):
        return str(int(value))
    return repr(float(value))


def part_id_base(part_id: str | None) -> str:
    """Return the part ID without its (last) temperature tag.

    ``"SN541_T25C"`` -> ``"SN541"`` — the strip removes the tag including a
    matched unit; with the no-unit lookahead form the separator survives
    (``"SN541_T25_DIE1"`` -> ``"SN541_DIE1"``, ``"SN541_T25_"`` stays
    ``"SN541_"``). Returns the full original ID unchanged when no tag is
    present (``""`` for a ``None`` ID).
    """
    if part_id is None:
        return ""
    matches = list(TEMPERATURE_RE.finditer(part_id))
    if not matches:
        return part_id
    last = matches[-1]
    return part_id[: last.start()] + part_id[last.end() :]


def parse_test_time(value) -> int | str:
    """Milliseconds as an int from a DB ``"N ms"`` string.

    The dutInfo Test Time is ``TestTime || ' ms'`` (schema.rs
    FETCH_SELECT_DUT_SUMMARY): strip the ``ms`` suffix and int() it; when
    that fails, fall back to the first integer found anywhere. Absent or
    unparseable values yield ``""`` (blank cell) so the numeric Min/Max/
    Mean formulas skip the DUT.
    """
    if value is None or value == "":
        return ""
    if isinstance(value, (int, float)):
        return int(value)
    text = str(value).strip()
    if text.lower().endswith("ms"):
        text = text[:-2].strip()
    try:
        return int(text)
    except ValueError:
        pass
    match = _FIRST_INT_RE.search(str(value))
    return int(match.group(0)) if match else ""


def parse_bin(value) -> int | str:
    """Bin number as an int from a DB ``"Bin N"`` string.

    The dutInfo HW/SW Bin values are ``'Bin ' || HBIN`` / ``'Bin ' || SBIN``
    (schema.rs): extract the first integer so the cells carry bare numbers
    (no "Bin" repetition) and the Min/Max/Mean formulas stay numeric.
    Absent or unparseable values yield ``""``.
    """
    if value is None or value == "":
        return ""
    if isinstance(value, (int, float)):
        return int(value)
    match = _FIRST_INT_RE.search(str(value))
    return int(match.group(0)) if match else ""


def compile_param_pattern(pattern: str | None) -> re.Pattern:
    """Compile a test/param split pattern, falling back to the default.

    ``None`` or an invalid regex yields ``DEFAULT_PARAM_PATTERN``.
    """
    if pattern is None:
        return DEFAULT_PARAM_PATTERN
    try:
        return re.compile(pattern)
    except re.error:
        return DEFAULT_PARAM_PATTERN


def split_test_name(name: str, pattern: str | None = None) -> tuple[str, str]:
    """Split a test name into ``(test, param)``; no match -> ``(name, "")``.

    Prefers the pattern's named ``test``/``param`` groups ("A/B/C" with the
    default pattern -> ("A/B", "C")); otherwise the param is the last group
    with its span stripped from the name.
    """
    pat = compile_param_pattern(pattern)
    match = pat.search(name)
    if match is None:
        return (name, "")
    groups = match.groupdict()
    if groups.get("test") is not None and groups.get("param") is not None:
        return (groups["test"], groups["param"])
    last = pat.groups
    if last and match.group(last) is not None:
        return (name[: match.start(last)] + name[match.end(last) :], match.group(last))
    return (name, match.group(0))


def workflow_key(test_name: str) -> str | None:
    """Return the first "/" segment of a test name, or ``None`` without "/".

    Used to detect workflow group boundaries ("Func/vdd/min" -> "Func").
    """
    head, sep, _ = test_name.partition("/")
    return head if sep else None


def flag_value(flag) -> int | None:
    """Map a raw STDF test flag to a cell value: ``None`` (not tested) /
    ``0`` (pass) / ``1`` (fail), mirroring ``deps.SharedSrc.isPass``.

    Negative flags mean "not tested" (isPass would call them pass, so the
    ``None`` check comes first).
    """
    if flag < 0:
        return None
    return 0 if isPass(flag) else 1


def split_dut_chunks(dut_ids: list, cap: int = DUT_CHUNK_CAP) -> tuple[tuple, ...]:
    """Split a list into consecutive chunks of at most ``cap`` per sheet
    (generic: named for the matrix's DUT axis, reused for the DUT Summary
    sheet's test axis).

    An empty list still yields one (empty) chunk so the caller renders one
    sheet; small lists yield a single chunk.
    """
    if len(dut_ids) <= cap:
        return (tuple(dut_ids),)
    return tuple(tuple(dut_ids[i : i + cap]) for i in range(0, len(dut_ids), cap))


def die_col(index: int) -> str:
    """Excel column letter of the ``index``-th die column (0 -> "M")."""
    return get_column_letter(13 + index)


def col_range(letter_first: str, letter_last: str, row: int) -> str:
    """Excel range string spanning the die columns of one row ("M17:AZ17")."""
    return f"{letter_first}{row}:{letter_last}{row}"


def stat_formula(name: str, row: int, first: str, last: str) -> str:
    """Row statistic over the die range, IFERROR-guarded (2026-11-11 user
    ruling: resilience): ``=IFERROR(AVERAGE(M17:AZ17),"")`` style.

    AVERAGE/STDEV error on all-blank / too-few ranges; MIN/MAX never error
    (an empty range computes 0) but take the same guard for uniformity —
    the error branch always blanks the cell."""
    return f'=IFERROR({name}({first}{row}:{last}{row}),"")'


def cp_formula(row: int, first_die: str, last_die: str) -> str:
    """Cp of one Test Summary data row: (USL-LSL)/(6σ) with σ embedded as
    STDEV over the row's die range — F=SpecMax, E=SpecMin (fixed letters).

    Ninth user-feedback wave (2026-11-10): the J column stopped being
    St. Dev. and became Cp itself, so the sample stdev is computed inline
    over ``{first}{row}:{last}{row}``. IFERROR blanks σ=0 (constant die
    values) and missing limits.
    """
    return f'=IFERROR((F{row}-E{row})/(6*STDEV({first_die}{row}:{last_die}{row})),"")'


def cpk_formula(row: int, first_die: str, last_die: str) -> str:
    """Cpk of one Test Summary data row (F=SpecMax, E=SpecMin, I=Mean), σ
    embedded as STDEV over the row's die range.

    Ninth user-feedback wave: the signature gains the die range — the old
    shape read σ from the removed J (St. Dev.) column, so both ratio arms
    now compute STDEV(``{first}{row}:{last}{row}``) inline. IFERROR blanks
    σ=0 and missing/one-sided limits.
    """
    rng = f"STDEV({first_die}{row}:{last_die}{row})"
    return f'=IFERROR(MIN((F{row}-I{row})/(3*{rng}),(I{row}-E{row})/(3*{rng})),"")'


def fails_formula_limits(row: int, first: str, last: str) -> str:
    """FAILS count vs low/high limits (E/F), ignoring blank DUT cells —
    IFERROR-guarded (2026-11-11) like every Test Summary formula."""
    rng = col_range(first, last, row)
    return (
        f'=IFERROR(SUMPRODUCT(({rng}<>"")'
        f'*((E{row}<>"")*({rng}<E{row})+(F{row}<>"")*({rng}>F{row}))),"")'
    )


def fails_formula_flags(row: int, first: str, last: str) -> str:
    """FAILS count of 0/1 flag cells over the hidden mirror range —
    IFERROR-guarded (2026-11-11) like every Test Summary formula."""
    mrng = col_range(first, last, row)
    return f'=IFERROR(SUMPRODUCT(({mrng}<>"")*({mrng}=1)),"")'


# --------------------------------------------------------------------------
# Column twins of the five builders above (DUT Summary wave): the DUT Summary
# sheet is the Test Summary matrix transposed, so the attribute block's
# Min/Max/Mean/Cp/Cpk/Fails cells run DOWN each test column over the DUT
# rows, and the spec-min/max/mean cells are attribute-block cells in the
# SAME column ({col}{spec_min_row} etc.) — the transpose of the matrix's
# fixed E/F/I letters. The caller passes the BASE_HEADERS-derived rows
# (dut_summary.SPEC_MIN_ROW / SPEC_MAX_ROW / MEAN_ROW), so the attribute
# layout stays the one source of truth.
# --------------------------------------------------------------------------


# SUBTOTAL visible-cells codes (filter-aware attribute block, ruling
# 2026-11-16): 101 AVERAGE / 104 MAX / 105 MIN / 107 STDEV (sample) —
# same semantics as the plain functions, restricted to visible rows;
# 103 COUNTA drives the OFFSET visible-mask idiom on the Fails counts.
# (106 is PRODUCT in the 101-111 series, NOT STDEV — 107 is.)
_STAT_CODES = {"MIN": 105, "MAX": 104, "AVERAGE": 101}


def stat_formula_col(name: str, col: str, first_row: int, last_row: int) -> str:
    """Column twin of stat_formula (DUT Summary attribute block), filter-aware:
    ``=IFERROR(SUBTOTAL(101,P14:P1000),"")`` style — SUBTOTAL codes ignore
    AutoFilter-hidden rows, one IFERROR guard, the range running down the
    test column."""
    return f'=IFERROR(SUBTOTAL({_STAT_CODES[name]},{col}{first_row}:{col}{last_row}),"")'


def cp_formula_col(
    col: str, first_row: int, last_row: int, spec_min_row: int, spec_max_row: int
) -> str:
    """Column twin of cp_formula (DUT Summary attribute block), filter-aware:
    (USL-LSL)/(6σ) with σ embedded as SUBTOTAL(107, ...) (sample STDEV over
    the visible cells) down the test column; the limits read the attribute
    block's SpecMax/SpecMin cells ({col}{spec_max_row} / {col}{spec_min_row}).
    IFERROR blanks σ=0 and missing limits."""
    return (
        f'=IFERROR(({col}{spec_max_row}-{col}{spec_min_row})'
        f'/(6*SUBTOTAL(107,{col}{first_row}:{col}{last_row})),"")'
    )


def cpk_formula_col(
    col: str,
    first_row: int,
    last_row: int,
    spec_min_row: int,
    spec_max_row: int,
    mean_row: int,
) -> str:
    """Column twin of cpk_formula (DUT Summary attribute block), filter-aware:
    the smaller of the two ratio arms over σ = SUBTOTAL(107, ...) (visible
    cells); limits and mean read the attribute block's cells
    ({col}{spec_max_row} / {col}{mean_row} / {col}{spec_min_row}). IFERROR
    blanks σ=0 and missing/one-sided limits."""
    rng = f"SUBTOTAL(107,{col}{first_row}:{col}{last_row})"
    return (
        f"=IFERROR(MIN(({col}{spec_max_row}-{col}{mean_row})/(3*{rng}),"
        f'({col}{mean_row}-{col}{spec_min_row})/(3*{rng})),"")'
    )


def fails_formula_limits_col(
    col: str, first_row: int, last_row: int, spec_min_row: int, spec_max_row: int
) -> str:
    """Column twin of fails_formula_limits (DUT Summary attribute block),
    filter-aware: FAILS count down one test column vs the attribute block's
    SpecMin/SpecMax cells, ignoring blank DUT cells and rows hidden by the
    AutoFilter (SUBTOTAL(103, OFFSET(...)) visible-mask idiom — volatile,
    accepted 2026-11-16; cached results stay full-population)."""
    rng = f"{col}{first_row}:{col}{last_row}"
    anchor = f"{col}${first_row}"
    visible = f"SUBTOTAL(103,OFFSET({anchor},ROW({rng})-ROW({anchor}),0))"
    return (
        f'=IFERROR(SUMPRODUCT(({rng}<>"")'
        f'*(({col}{spec_min_row}<>"")*({rng}<{col}{spec_min_row})'
        f'+({col}{spec_max_row}<>"")*({rng}>{col}{spec_max_row}))'
        f"*{visible})"
        ',"")'
    )


def fails_formula_flags_col(
    mirror_first: str, mirror_last: str, first_row: int, last_row: int
) -> str:
    """Column twin of fails_formula_flags (DUT Summary attribute block),
    filter-aware: FAILS count of 0/1 flag cells down the hidden mirror
    range, restricted to rows visible under the AutoFilter (SUBTOTAL(103,
    OFFSET(...)) visible-mask idiom — volatile, accepted 2026-11-16)."""
    mrng = f"{mirror_first}{first_row}:{mirror_last}{last_row}"
    anchor = f"{mirror_first}${first_row}"
    visible = f"SUBTOTAL(103,OFFSET({anchor},ROW({mrng})-ROW({anchor}),0))"
    return f'=IFERROR(SUMPRODUCT(({mrng}<>"")*({mrng}=1)*{visible}),"")'


def metadata_stat_formulas(
    row: int, first: str, last: str, mean: bool = True, std: bool = False
) -> tuple[str, ...]:
    """(MIN, MAX[, AVERAGE][, STDEV]) formula set for a numeric metadata row.

    ``mean=False`` drops the AVERAGE formula; ``std=True`` appends the
    STDEV formula (sample stdev — the same ``STDEV`` semantic the data
    rows' St. Dev. column uses). Fourth wave 2026-11-10: only Temperature
    keeps ``mean=True`` alone; the Test Time row takes ``mean=True`` with
    ``std=True``; every other numeric row takes ``(False, False)``. The
    caller pads the Mean/St. Dev. cells with dashes for the dropped ones.
    """
    formulas = [
        stat_formula("MIN", row, first, last),
        stat_formula("MAX", row, first, last),
    ]
    if mean:
        formulas.append(stat_formula("AVERAGE", row, first, last))
    if std:
        formulas.append(stat_formula("STDEV", row, first, last))
    return tuple(formulas)


# --------------------------------------------------------------------------
# Fifth user-feedback wave (2026-11-10): live Test Statistics Cpk formula
# --------------------------------------------------------------------------


def _is_number(value) -> bool:
    """True for real int/float cells (bools and NaN excluded — a bool would
    silently average as 0/1 and a NaN would poison every statistic)."""
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return False
    return not math.isnan(value)


def numeric_values(cells) -> list:
    """The numeric cells of ``cells``: int/float kept, bools/None/NaN
    dropped, strings kept only when float-parseable ("2.5" -> 2.5)."""
    out = []
    for cell in cells:
        if _is_number(cell):
            out.append(cell)
        elif isinstance(cell, str):
            try:
                out.append(float(cell.strip()))
            except ValueError:
                continue
    return out


def stat_cached(name: str, values) -> "float | None":
    """MIN/MAX/AVERAGE/STDEV over the numeric cells (sample STDEV, the
    same semantic the sheet's STDEV() uses); None when uncomputable —
    no values, or fewer than 2 for STDEV; unknown names are None too."""
    nums = numeric_values(values)
    if name == "MIN":
        return min(nums) if nums else None
    if name == "MAX":
        return max(nums) if nums else None
    if name == "AVERAGE":
        return statistics.fmean(nums) if nums else None
    if name == "STDEV":
        return statistics.stdev(nums) if len(nums) >= 2 else None
    return None


def _ratio_cached(spec_min, spec_max, values, mode) -> "float | None":
    """Shared Cp/Cpk computation over the row's numeric die values.

    Mirrors the sheet formulas: sigma = STDEV(range), mu = AVERAGE(range),
    Cp = (max-min)/(6*sigma), Cpk = min((max-mu)/(3*sigma), (mu-min)/(3*sigma)).
    None on missing/one-sided limits, sigma == 0, or fewer than 2 values —
    the exact cases the IFERROR wrappers blank on the sheet."""
    if spec_min is None or spec_max is None:
        return None
    nums = numeric_values(values)
    if len(nums) < 2:
        return None
    sigma = statistics.stdev(nums)
    if sigma == 0:
        return None
    mu = statistics.fmean(nums)
    spread = spec_max - spec_min
    if mode == "cp":
        return spread / (6 * sigma)
    return min((spec_max - mu) / (3 * sigma), (mu - spec_min) / (3 * sigma))


def cp_cached(spec_min, spec_max, values) -> "float | None":
    """Cached Cp over a die-value row (see ``_ratio_cached``)."""
    return _ratio_cached(spec_min, spec_max, values, "cp")


def cpk_cached(spec_min, spec_max, values) -> "float | None":
    """Cached Cpk over a die-value row (see ``_ratio_cached``)."""
    return _ratio_cached(spec_min, spec_max, values, "cpk")


def fails_cached_limits(values, lo, hi) -> int:
    """Count of numeric values strictly outside (lo, hi) — the same
    exclusive comparison the SUMPRODUCT fail formulas encode
    (rng<lo or rng>hi); blank/non-numeric cells never count. Both limits
    must be present (the both-limits branch)."""
    return sum(1 for v in numeric_values(values) if v < lo or v > hi)


def fails_cached_limits_one_sided(values, lo, hi) -> int:
    """Count of numeric values failing a ONE-SIDED window — the faithful
    cache for `has_limits` rows (OR semantics): each SUMPRODUCT side is
    guarded by ``(E16<>"")*(rng<E16)`` / ``(F16<>"")*(rng>F16)``, so a
    missing side never fails. Mirrors exactly that: below a PRESENT lo,
    or above a PRESENT hi. The matrix layer falls back to the flag cache
    only when BOTH sides are missing (the formula uses the flag mirror
    there)."""
    fails = 0
    for v in numeric_values(values):
        if (lo is not None and v < lo) or (hi is not None and v > hi):
            fails += 1
    return fails


def fails_cached_flags(flags) -> int:
    """Count of recorded fail flags (1s) — the flag-mirror SUMPRODUCT."""
    return sum(1 for f in numeric_values(flags) if f == 1)


def has_numeric(values) -> bool:
    """True when ``values`` carries at least one numeric cell - the
    sparkline gate (rows without numeric die values get no sparkline)."""
    return bool(numeric_values(values))


def row_stat_caches(nums, lo, hi, mirror_cells) -> tuple:
    """One-pass row statistics for a summary row: the 6-tuple
    ``(mn, mx, avg, cp, cpk, fails)``.

    ``nums`` is the ALREADY-filtered numeric list (a ``numeric_values``
    output the caller builds once — never re-filtered here, so no
    ``numeric_values`` pass over it). Gating copied verbatim from today's
    helpers: MIN/MAX/AVG are None on an empty list (``stat_cached``);
    Cp/Cpk share ONE ``statistics.stdev`` + ONE ``statistics.fmean`` and
    are None when either limit is missing, ``len(nums) < 2`` or
    sigma == 0 (the ``_ratio_cached`` cases the IFERROR wrappers blank).
    FAILS is branch-selected inside: with a limit present on either side,
    the one-sided count over ``nums`` (``fails_cached_limits_one_sided``
    OR semantics — a missing side never fails); with BOTH limits absent,
    the flag count ``sum(1 for f in numeric_values(mirror_cells)
    if f == 1)`` (``fails_cached_flags``, the flag-mirror SUMPRODUCT —
    only the small mirror list is filtered). Callers keep the MIN/MAX
    faithful-0 and None->"" mapping.
    """
    mn = min(nums) if nums else None
    mx = max(nums) if nums else None
    avg = statistics.fmean(nums) if nums else None

    cp = None
    cpk = None
    if lo is not None and hi is not None and len(nums) >= 2:
        sigma = statistics.stdev(nums)
        if sigma != 0:
            cp = (hi - lo) / (6 * sigma)
            cpk = min((hi - avg) / (3 * sigma), (avg - lo) / (3 * sigma))

    if lo is not None or hi is not None:
        fails = sum(
            1 for v in nums
            if (lo is not None and v < lo) or (hi is not None and v > hi)
        )
    else:
        fails = sum(1 for f in numeric_values(mirror_cells) if f == 1)
    return mn, mx, avg, cp, cpk, fails


# Control characters Excel forbids inside sheet XML strings (copied out of
# openpyxl when the package left that dependency behind — tab \011 and
# newline \012 stay legal).
ILLEGAL_CHARACTERS_RE = re.compile(r"[\000-\010]|[\013-\014]|[\016-\037]")


def sanitize_cell(value):
    """Make a value safe for Excel sheet XML: strip illegal control
    characters from strings, map NaN (and None) to blank, pass scalars
    through unchanged."""
    if isinstance(value, str):
        return ILLEGAL_CHARACTERS_RE.sub("", value)
    if value is None:
        return ""
    if isinstance(value, float) and math.isnan(value):
        return ""
    return value


def reorder_sheets_to_end(wss: list, sheets: list) -> list:
    """The 2026-11-28 tab-order ruling: move the listed worksheets to the
    END of the workbook list, in listed order — GDR & DTR Summary and Bin
    Chart close the strip. Identity-based (the generator's pre-created
    worksheet objects), ``None`` entries (sheet not selected) and missing
    sheets are skipped; everything else keeps its relative order. Pure —
    the generator calls it in the ``finally`` between ``order_dut_block``
    and the ``ws.index`` renumber, so the renumber still sees the final
    order."""
    out = list(wss)
    for sheet in sheets:
        if sheet is None or sheet not in out:
            continue
        out.remove(sheet)
        out.append(sheet)
    return out
