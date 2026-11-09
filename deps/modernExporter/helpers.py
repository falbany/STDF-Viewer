"""
#
# helpers.py - modernExporter package - STDF Viewer
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
"""

"""Pure parsing / formula / chunking helpers for the modern openpyxl exporter."""

import re

from openpyxl.utils import get_column_letter

from deps.SharedSrc import isPass

# Part-ID temperature tag, e.g. the "_T25C" in "SN541_T25C"; case-insensitive,
# optional decimal, and signed so cold temperatures ("_T-5.5C") match.
TEMPERATURE_RE = re.compile(r"_T(-?\d+(?:\.\d+)?)C", re.IGNORECASE)

# First integer anywhere in a string (fallback extractor for the DB-shaped
# "N ms" / "Bin N" metadata values below).
_FIRST_INT_RE = re.compile(r"-?\d+")

# Default test/param split: "workflow/rest/of/name/param" -> test + last segment.
DEFAULT_PARAM_PATTERN = re.compile(r"^(?P<test>.+)/(?P<param>[^/]+)$")

# Excel hard limit is 16384 columns per sheet. A DUT Summary chunk sheet is
# 13 fixed columns (A..L base block) + n visible die columns + 1 spacer + n
# hidden flag mirrors, so a chunk fits only while 13 + n + 1 + n <= 16384,
# i.e. n <= (16384 - 13 - 1) // 2 = 8185 DUTs per sheet.
DUT_CHUNK_CAP = 8185


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


def split_dut_chunks(dut_ids: list[str], cap: int = DUT_CHUNK_CAP) -> tuple[tuple[str, ...], ...]:
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


def cpk_formula(row: int) -> str:
    """Cpk from the row's Min/Max/Average/StdDev cells (E/F/I/J)."""
    return f'=IFERROR(MIN((F{row}-I{row})/(3*J{row}),(I{row}-E{row})/(3*J{row})),"")'


def fails_formula_limits(row: int, first: str, last: str) -> str:
    """FAILS count vs low/high limits (E/F), ignoring blank DUT cells."""
    rng = col_range(first, last, row)
    return f'=SUMPRODUCT(({rng}<>"")*((E{row}<>"")*({rng}<E{row})+(F{row}<>"")*({rng}>F{row})))'


def fails_formula_flags(row: int, first: str, last: str) -> str:
    """FAILS count of 0/1 flag cells over the hidden mirror range."""
    mrng = col_range(first, last, row)
    return f'=SUMPRODUCT(({mrng}<>"")*({mrng}=1))'


def fails_formula_ftr(row: int, first: str, last: str) -> str:
    """FAILS count of 0/1 FTR cells over the visible die range."""
    rng = col_range(first, last, row)
    return f'=SUMPRODUCT(({rng}<>"")*({rng}=1))'


def metadata_stat_formulas(row: int, first: str, last: str) -> tuple[str, str, str]:
    """(MIN, MAX, AVERAGE) formula triple for a numeric metadata block row."""
    return (stat_formula("MIN", row, first, last),
            stat_formula("MAX", row, first, last),
            stat_formula("AVERAGE", row, first, last))
