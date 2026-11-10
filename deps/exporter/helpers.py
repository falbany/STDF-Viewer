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

# Part-ID temperature tag, e.g. the "_T25C" in "SN541_T25C"; case-insensitive,
# optional decimal, and signed so cold temperatures ("_T-5.5C") match.
TEMPERATURE_RE = re.compile(r"_T(-?\d+(?:\.\d+)?)C", re.IGNORECASE)

# First integer anywhere in a string (fallback extractor for the DB-shaped
# "N ms" / "Bin N" metadata values below).
_FIRST_INT_RE = re.compile(r"-?\d+")

# First signed decimal number anywhere in a string (File Info temperature
# fallback: MIR TST_TEMP arrives as free text like "25C" / "-40C" / "25.5 °C").
_FIRST_FLOAT_RE = re.compile(r"-?\d+(?:\.\d+)?")

# Default test/param split: "workflow/rest/of/name/param" -> test + last segment.
DEFAULT_PARAM_PATTERN = re.compile(r"^(?P<test>.+)/(?P<param>[^/]+)$")

# Excel hard limit is 16384 columns per sheet. A DUT Summary chunk sheet is
# 13 fixed columns (A..L base block) + n visible die columns + 1 spacer +
# 1 Trend sparkline column + 1 spacer + n hidden flag mirrors, so a chunk
# fits only while 13 + n + 3 + n <= 16384, i.e.
# n <= (16384 - 13 - 3) // 2 = 8184 DUTs per sheet.
DUT_CHUNK_CAP = 8184


def parse_temperature(part_id: str | None) -> float | None:
    """Return the temperature encoded in a part ID, last match wins.

    ``"SN541_T25C"`` -> ``25.0``; ``"SN541_T-5.5C"`` -> ``-5.5``;
    ``None`` when the ID is ``None`` or carries no ``_T<n>C`` tag.
    """
    if part_id is None:
        return None
    matches = TEMPERATURE_RE.findall(part_id)
    if not matches:
        return None
    return float(matches[-1])


def parse_temperature_text(value) -> float | None:
    """Temperature as a float from a File Info "Test Temperature:" cell.

    MIR TST_TEMP is free text — ``"25C"``, ``"-40C"``, ``"25.5 °C"`` and a
    bare ``"25"`` all parse; the first signed decimal wins. A multi-subfile
    group cell collapses per formatFileGroupValues into
    ``"#1 → 25C\\n#2 → -40C"`` — the first subfile's line is taken. ``None``
    when the value is absent or carries no number (blank cell, "N/A").
    """
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value) if value == value else None  # NaN -> None
    text = str(value).strip()
    if "\n" in text:
        # multi-subfile composite ("#1 → 25C\n#2 → -40C"): first subfile wins
        text = text.split("\n", 1)[0]
    if text.startswith("#") and "→" in text:
        # drop the "#n →" source prefix so its index never parses as a number
        text = text.split("→", 1)[1]
    match = _FIRST_FLOAT_RE.search(text)
    return float(match.group(0)) if match else None


def part_id_base(part_id: str | None) -> str:
    """Return the part ID without its (last) temperature tag.

    ``"SN541_T25C"`` -> ``"SN541"``; returns the full original ID unchanged
    when no tag is present (``""`` for a ``None`` ID).
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


def split_dut_chunks(
    dut_ids: list[str], cap: int = DUT_CHUNK_CAP
) -> tuple[tuple[str, ...], ...]:
    """Split DUT IDs into consecutive chunks of at most ``cap`` per sheet.

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
    """Row statistic over the die range: ``=MIN(M17:AZ17)`` style."""
    return f"={name}({first}{row}:{last}{row})"


def cp_formula(row: int, first_die: str, last_die: str) -> str:
    """Cp of one DUT Summary data row: (USL-LSL)/(6σ) with σ embedded as
    STDEV over the row's die range — F=SpecMax, E=SpecMin (fixed letters).

    Ninth user-feedback wave (2026-11-10): the J column stopped being
    St. Dev. and became Cp itself, so the sample stdev is computed inline
    over ``{first}{row}:{last}{row}``. IFERROR blanks σ=0 (constant die
    values) and missing limits.
    """
    return f'=IFERROR((F{row}-E{row})/(6*STDEV({first_die}{row}:{last_die}{row})),"")'


def cpk_formula(row: int, first_die: str, last_die: str) -> str:
    """Cpk of one DUT Summary data row (F=SpecMax, E=SpecMin, I=Mean), σ
    embedded as STDEV over the row's die range.

    Ninth user-feedback wave: the signature gains the die range — the old
    shape read σ from the removed J (St. Dev.) column, so both ratio arms
    now compute STDEV(``{first}{row}:{last}{row}``) inline. IFERROR blanks
    σ=0 and missing/one-sided limits.
    """
    rng = f"STDEV({first_die}{row}:{last_die}{row})"
    return f'=IFERROR(MIN((F{row}-I{row})/(3*{rng}),(I{row}-E{row})/(3*{rng})),"")'


def fails_formula_limits(row: int, first: str, last: str) -> str:
    """FAILS count vs low/high limits (E/F), ignoring blank DUT cells."""
    rng = col_range(first, last, row)
    return f'=SUMPRODUCT(({rng}<>"")*((E{row}<>"")*({rng}<E{row})+(F{row}<>"")*({rng}>F{row})))'


def fails_formula_flags(row: int, first: str, last: str) -> str:
    """FAILS count of 0/1 flag cells over the hidden mirror range."""
    mrng = col_range(first, last, row)
    return f'=SUMPRODUCT(({mrng}<>"")*({mrng}=1))'


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


def build_stat_cpk_formula(
    spec_min_l: str, spec_max_l: str, avg_l: str, std_l: str, row_number: int
) -> str:
    """Live Cpk formula for one Test Statistics row, DUT-Summary shaped.

    ``=IFERROR(MIN((<max>-<avg>)/(3*<std>),(<avg>-<min>)/(3*<std>)),"")``
    with the column letters substituted: substituting the DUT Summary's
    fixed E/F/I/J letters reproduces :func:`cpk_formula` verbatim (one
    formula shape, two sheets). ``row_number`` is the 1-based Excel row.
    The IFERROR blank covers σ=0 and missing/one-sided limits, so the
    "∞" / "N/A" string sentinels disappear from the column.
    """
    return (
        f"=IFERROR(MIN(({spec_max_l}{row_number}-{avg_l}{row_number})"
        f"/(3*{std_l}{row_number}),({avg_l}{row_number}-{spec_min_l}"
        f'{row_number})/(3*{std_l}{row_number})),"")'
    )


def apply_cpk_formula(stat_row: list, header_info: dict, row_number: int) -> None:
    """Overwrite the Cpk cell of one transformed Test Statistics row in
    place with :func:`build_stat_cpk_formula`, computed from the row's
    SpecMin/SpecMax/Average/St. Dev. columns.

    ``header_info`` carries the 0-based anchors from
    ``generator.transform_stat_header`` (``spec_cols``, ``avg_col``,
    ``std_col``, ``cpk_col``). No-op when the transform is disabled or any
    of the four anchors is missing (translated-UI degradation — the cell
    keeps its upstream string), and when a degraded row is shorter than
    the Cpk column (no IndexError, no phantom cells).
    """
    if not header_info.get("enabled"):
        return
    spec_cols = header_info.get("spec_cols")
    avg_col = header_info.get("avg_col")
    std_col = header_info.get("std_col")
    cpk_col = header_info.get("cpk_col")
    if spec_cols is None or avg_col is None or std_col is None or cpk_col is None:
        return
    if not cpk_col < len(stat_row):
        return
    stat_row[cpk_col] = build_stat_cpk_formula(
        get_column_letter(spec_cols[0] + 1),
        get_column_letter(spec_cols[1] + 1),
        get_column_letter(avg_col + 1),
        get_column_letter(std_col + 1),
        row_number,
    )


# --------------------------------------------------------------------------
# Ninth user-feedback wave (2026-11-10): live Test Statistics Cp formula
# --------------------------------------------------------------------------


def build_stat_cp_formula(
    spec_min_l: str, spec_max_l: str, std_l: str, row_number: int
) -> str:
    """Live Cp formula for one Test Statistics row: (USL-LSL)/(6σ).

    ``=IFERROR(({max}{r}-{min}{r})/(6*{std}{r}),"")`` with the column
    letters substituted. Unlike the DUT Summary's :func:`cp_formula` —
    which embeds STDEV over the die range because its J column is now Cp
    itself — the Test Statistics σ comes from the numeric "St. Dev."
    column (upstream values are numeric since the third wave).
    ``row_number`` is the 1-based Excel row; the IFERROR blank covers
    σ=0 and missing/one-sided limits.
    """
    return (
        f"=IFERROR(({spec_max_l}{row_number}-{spec_min_l}{row_number})"
        f'/(6*{std_l}{row_number}),"")'
    )


def apply_cp_formula(stat_row: list, header_info: dict, row_number: int) -> None:
    """Overwrite the Cp cell of one transformed Test Statistics row in
    place with :func:`build_stat_cp_formula`, computed from the row's
    SpecMin/SpecMax/St. Dev. columns.

    ``header_info`` carries the 0-based anchors from
    ``generator.transform_stat_header`` (``spec_cols``, ``std_col``,
    ``cp_col``). No-op when the transform is disabled or any of the three
    anchors is missing (translated-UI degradation — the cell keeps its
    placeholder), and when a degraded row is shorter than the Cp column
    (no IndexError, no phantom cells).
    """
    if not header_info.get("enabled"):
        return
    spec_cols = header_info.get("spec_cols")
    std_col = header_info.get("std_col")
    cp_col = header_info.get("cp_col")
    if spec_cols is None or std_col is None or cp_col is None:
        return
    if not cp_col < len(stat_row):
        return
    stat_row[cp_col] = build_stat_cp_formula(
        get_column_letter(spec_cols[0] + 1),
        get_column_letter(spec_cols[1] + 1),
        get_column_letter(std_col + 1),
        row_number,
    )


# --------------------------------------------------------------------------
# cached-value engine (xlsxwriter migration, spec §7.5 upgrade 1)
#
# xlsxwriter stores no computed result for a formula unless one is supplied;
# previews (QuickLook/Numbers/Google Sheets) show that cache. These helpers
# recompute in Python exactly what the sheet formulas compute, so every
# write_formula can pass an explicit value=. Uncomputable results are None
# at this layer (the spec builder maps None -> "" so the cache stays blank
# instead of a stray 0).
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
    (rng<lo or rng>hi); blank/non-numeric cells never count."""
    return sum(1 for v in numeric_values(values) if v < lo or v > hi)


def fails_cached_flags(flags) -> int:
    """Count of recorded fail flags (1s) — the flag-mirror SUMPRODUCT."""
    return sum(1 for f in numeric_values(flags) if f == 1)


def stat_cp_value(spec_min, spec_max, std) -> "float | None":
    """Cached Cp for one Test Statistics row: sigma comes from the row's
    St. Dev. COLUMN (not a die range), so only the limits and sigma gate:
    None on missing/one-sided limits or sigma <= 0/None."""
    if spec_min is None or spec_max is None:
        return None
    if not _is_number(std) or std <= 0:
        return None
    return (spec_max - spec_min) / (6 * std)


def stat_cpk_value(spec_min, spec_max, avg, std) -> "float | None":
    """Cached Cpk for one Test Statistics row (sigma from the St. Dev.
    column); None on missing/one-sided limits or sigma <= 0/None."""
    if spec_min is None or spec_max is None:
        return None
    if not _is_number(std) or std <= 0:
        return None
    return min((spec_max - avg) / (3 * std), (avg - spec_min) / (3 * std))


def has_numeric(values) -> bool:
    """True when ``values`` carries at least one numeric cell — the
    sparkline gate (rows without numeric die values get no sparkline)."""
    return bool(numeric_values(values))
