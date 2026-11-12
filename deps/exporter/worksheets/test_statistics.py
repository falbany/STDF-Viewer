"""
#
# test_statistics.py - exporter worksheets - STDF Viewer
#
# Author: falbany (fork)
# Created Date: November 12th 2026
# -----
# Last Modified: Thu Nov 12 2026
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

Test Statistics worksheet: header/row transforms and conditional-formatting
ranges for the modern sheet.

Pure module-level code moved verbatim from generator.py (worksheets split,
2026-11-12); the TestStatisticsMixin writer joins this file in the
extraction wave.
"""

import re

import xlsxwriter.utility as xlutil

from deps.SharedSrc import getSetting, tab
from deps.uic_stdExporter import ReportSelection
from deps.exporter.dedup import drops_to_exclusion
from deps.exporter.helpers import (
    apply_cp_formula,
    apply_cpk_formula,
    get_column_letter,
    split_test_name,
    stat_cp_value,
    stat_cpk_value,
)
from deps.exporter.sheets import plain_table_spec
from deps.exporter.theme import TAB_COLOR_HEXES
from deps.exporter.writer import render_plain_table, write_plain_row
from deps.exporter.worksheets.dut_summary import FTR_SPEC_MAX, FTR_SPEC_MIN

# The GUI's waitForDataList rows (DataInterface.getStatistics) arrive as
# header = [""] + translated hHeader and data rows = [VHeader, TestName,
# ...value cells]. The transforms below reshape them for the modern sheet:
# split the VHeader column, rename the limit columns, insert a Param
# column, and paste the stat values as numeric cells. Header matching is
# English-only (whitespace-normalized), same ruling as the File Info wave:
# under translated UIs the rename/Param transforms degrade to untransformed
# columns, while the VHeader split and the numericize stay structural.

# The received leading "" VHeader header cell expands into these labels.
STAT_VHEADER_LABELS = ("Test#", "Head", "Sites", "File ID")

# VHeader shape (DataInterface.py:521-524):
# "{test_num} / Head{head} / {All Sites|Site<site>}[ / File{fid}]"
VHEADER_RE = re.compile(
    r"^(\d+)\s*/\s*Head(\d+)\s*/\s*(All Sites|Site(-?\d+))(?:\s*/\s*File(\d+))?$"
)

# SpecMin/SpecMax stylization on the Test Statistics sheet (B4), mirroring
# the DUT Summary E/F columns.
STAT_SPEC_FORMAT = "0.00E+00"
STAT_SPEC_FILL_HEX = "FFF3CD"

# Live Cpk display format (fifth wave 2026-11-10, controller ruling): the
# formula results render with 2 decimals like the DUT Summary's Cpk column.
STAT_CPK_FORMAT = "0.00"

# Fixed width for the Test Statistics value columns G..Q (sixth wave
# 2026-11-10; ninth wave extends P -> Q for the inserted Cp column):
# applied AFTER the sheet-wide autofit, which keeps sizing A:F (the
# test-name column benefits).
STAT_VALUE_WIDTH_RANGE = ("G", "Q")
STAT_VALUE_COLUMN_WIDTH = 8

# English-only header anchors (whitespace-normalized matching).
_STAT_RENAMES = {"Low Limit": "SpecMin", "High Limit": "SpecMax"}
_STAT_TEST_NAME = "Test Name"
_STAT_FAIL_NUM = "Fail Num"
_STAT_CPK = "Cpk"
_STAT_CP = "Cp"
_STAT_AVERAGE = "Average"
_STAT_STD = "St. Dev."
_STAT_PATTERN_NAME = "Pattern Name"

# Group-banding fill hex (theme token) passed to write_plain_row as a
# plain string — xlsxwriter formats compose per cell.
theme_group_fill_hex = "E9EEF6"


def _norm_header_text(cell) -> str:
    """Whitespace-normalized header text for the English-only matching."""
    return " ".join(str(cell).split())


def parse_vheader(raw) -> tuple[int, int, "int | str", "int | str"] | None:
    """(test_num, head, site, fid) from one VHeader string, or None.

    ``site`` is the string ``"All Sites"`` when the DUT row aggregates all
    sites (site == -1), else an int; ``fid`` is an int when the VHeader
    carries a ``/ File<n>`` segment (multi-file export) and ``""`` for the
    single-file shape. Unparseable / non-string input returns None (the
    caller degrades the row gracefully, never crashes).
    """
    if not isinstance(raw, str):
        return None
    match = VHEADER_RE.match(raw.strip())
    if match is None:
        return None
    site = "All Sites" if match.group(3) == "All Sites" else int(match.group(4))
    fid = int(match.group(5)) if match.group(5) is not None else ""
    return int(match.group(1)), int(match.group(2)), site, fid


def numericize_value(value):
    """int for int-like strings, float for float-like strings, else the
    value unchanged (identity on non-strings).

    ``"12"`` -> 12, ``"1.500"`` -> 1.5, ``"1e3"`` -> 1000.0; the upstream
    sentinels ``"N/A"`` / ``"∞"`` / ``"?"`` and empty cells stay strings,
    so the sheet holds real numbers wherever the source was numeric.
    """
    if not isinstance(value, str):
        return value
    text = value.strip()
    try:
        return int(text)
    except ValueError:
        pass
    try:
        return float(text)
    except ValueError:
        return value


def transform_stat_header(header_row) -> dict:
    """Transform the received Test Statistics header row.

    - the structural leading ``""`` VHeader cell expands into the four
      split labels (Test#, Head, Sites, File ID);
    - ``"Low Limit"`` / ``"High Limit"`` rename to ``"SpecMin"`` /
      ``"SpecMax"`` in place (normalized, English-only);
    - ``"Param"`` is inserted right after the ``"Test Name"`` entry
      (English-only anchor);
    - a ``"Cp"`` column is inserted immediately BEFORE the ``"Cpk"``
      entry (normalized English anchor — ninth wave; degraded/disabled
      headers skip the insert, so no formulas are written either).

    Returns ``{"header": [...], "enabled": bool, "param_inserted": bool,
    "spec_cols": (i, j) | None, "avg_col": int | None, "std_col": int | None,
    "cp_col": int | None, "cpk_col": int | None}`` where ``spec_cols`` is
    the 0-based (SpecMin, SpecMax) column pair in the TRANSFORMED header
    (None when either limit anchor was missing — translated UIs degrade to
    untransformed columns), the fifth-wave keys carry the 0-based Average /
    St. Dev. / Cpk anchors and the ninth-wave ``cp_col`` the inserted Cp
    column (all normalized matching, None when missing). Every anchor scan
    runs AFTER the Cp insertion, so ``cpk_col`` points at the real Cpk
    column; the spec columns sit LEFT of the insertion and keep their
    indices. An unexpected non-blank first cell disables the whole
    transform (header returned untouched, all anchors None).
    """
    cells = [str(cell) if cell is not None else "" for cell in header_row]
    if not cells or _norm_header_text(cells[0]) != "":
        return {
            "header": list(header_row),
            "enabled": False,
            "param_inserted": False,
            "spec_cols": None,
            "avg_col": None,
            "std_col": None,
            "cp_col": None,
            "cpk_col": None,
        }
    body = [str(cell) for cell in cells[1:]]
    normalized = [_norm_header_text(cell) for cell in body]
    renamed = [_STAT_RENAMES.get(norm, cell) for norm, cell in zip(normalized, body)]
    header = list(STAT_VHEADER_LABELS) + renamed
    param_inserted = _STAT_TEST_NAME in normalized
    if param_inserted:
        header.insert(4 + normalized.index(_STAT_TEST_NAME) + 1, "Param")
    # ninth wave: "Cp" immediately left of "Cpk" (normalized anchor); the
    # anchor scans below run AFTER this insertion so cpk_col points at the
    # real Cpk column and spec_cols reflects the final positions
    cp_col = None
    for index, cell in enumerate(header):
        if _norm_header_text(cell) == _STAT_CPK:
            header.insert(index, "Cp")
            cp_col = index
            break
    spec_cols = None
    if "SpecMin" in renamed and "SpecMax" in renamed:
        spec_cols = (header.index("SpecMin"), header.index("SpecMax"))
    norm_header = [_norm_header_text(cell) for cell in header]

    def _anchor(text):
        return norm_header.index(text) if text in norm_header else None

    return {
        "header": header,
        "enabled": True,
        "param_inserted": param_inserted,
        "spec_cols": spec_cols,
        "avg_col": _anchor(_STAT_AVERAGE),
        "std_col": _anchor(_STAT_STD),
        "cp_col": cp_col,
        "cpk_col": _anchor(_STAT_CPK),
    }


def transform_stat_row(row, header_info: dict) -> list:
    """Transform one received data row to mirror ``transform_stat_header``.

    - cell 0 (VHeader) parses into (Test#, Head, Sites, File ID); a
      parse failure degrades to [raw VHeader, "", "", ""] — never a crash;
    - the Test Name cell splits into (prefix, Param leaf) via
      ``split_test_name`` when the Param column exists (English anchor);
      otherwise the full name stays untransformed;
    - every remaining value cell (Unit onward, including Pattern Name /
      MPR mix columns) goes through ``numericize_value``;
    - a ``""`` placeholder is inserted at the Cp column index (ninth
      wave) so the upstream Cpk value and everything right of it shift
      one cell right — rows stay header-aligned; the applier overwrites
      the placeholder with the live formula. Rows that end before the Cp
      column (truncated upstream) skip the insert — no phantom cells.
    """
    if not header_info.get("enabled"):
        return list(row)
    cells = list(row)
    raw_vheader = cells[0] if cells else ""
    parsed = parse_vheader(raw_vheader)
    out = list(parsed) if parsed is not None else [raw_vheader, "", "", ""]
    rest = cells[1:]
    if header_info.get("param_inserted"):
        name = rest[0] if rest else ""
        prefix, leaf = split_test_name(name if isinstance(name, str) else str(name))
        out.extend([prefix, leaf])
    elif rest:
        out.append(rest[0])
    out.extend(numericize_value(value) for value in rest[1:])
    cp_col = header_info.get("cp_col")
    if cp_col is not None and cp_col < len(out):
        out.insert(cp_col, "")
    return out


def stat_cf_ranges(header_info: dict, data_row_count: int) -> dict:
    """Conditional-formatting ranges for the Test Statistics sheet.

    Computes the Fail Num / Cpk / Cp column letters from the TRANSFORMED
    header (normalized English anchors — positions shift right of the
    inserted FTR "Pattern Name" / MPR mix columns and the ninth-wave Cp
    column, never hard-coded) and spans the DATA rows only:
    ``{"fail_num_range": "J2:Jn", "cp_range": "K2:Kn", "cpk_range":
    "L2:Ln"}``. An empty dict when either the Fail Num or the Cpk anchor
    is missing (translated-UI degradation), when the transform is
    disabled, or when there are no data rows; ``cp_range`` is present
    whenever the Cp anchor is (the transform inserts it with Cpk).
    """
    header = header_info.get("header") or []
    normalized = [_norm_header_text(cell) for cell in header]
    if (
        _STAT_FAIL_NUM not in normalized
        or _STAT_CPK not in normalized
        or data_row_count <= 0
    ):
        return {}
    last_row = 1 + data_row_count
    fail_letter = get_column_letter(normalized.index(_STAT_FAIL_NUM) + 1)
    cpk_letter = get_column_letter(normalized.index(_STAT_CPK) + 1)
    ranges = {
        "fail_num_range": f"{fail_letter}2:{fail_letter}{last_row}",
        "cpk_range": f"{cpk_letter}2:{cpk_letter}{last_row}",
    }
    if _STAT_CP in normalized:
        cp_letter = get_column_letter(normalized.index(_STAT_CP) + 1)
        ranges["cp_range"] = f"{cp_letter}2:{cp_letter}{last_row}"
    return ranges


def _stat_full_test_name(transformed_row, header_info) -> str:
    """Full upstream test name from a transformed row: prefix + "/" + leaf
    when the Param column exists and the leaf is non-empty — the exact
    inverse of ``split_test_name``; the untransformed cell otherwise."""
    cells = transformed_row
    if header_info.get("param_inserted"):
        prefix = cells[4] if len(cells) > 4 else ""
        leaf = cells[5] if len(cells) > 5 else ""
        return f"{prefix}/{leaf}" if leaf else prefix
    return cells[4] if len(cells) > 4 else ""


def stat_row_is_ftr(transformed_row, header_info, ftr_tests) -> bool:
    """FTR detection for one transformed Test Statistics row (fourth wave).

    PRIMARY: ``(test_num, full test name)`` sits in the cache that
    ``writeDUT_Summary`` builds from the summaryTable's SUB_CODE — needs a
    parsed VHeader (the row's leading int). FALLBACK (covers Stat-selected-
    without-DUT runs, where the cache stays empty): the transformed header
    carries a "Pattern Name" column (normalized match) and the row's cell
    there is a non-empty string — upstream sets VECT_NAM only for FTR rows
    and "" otherwise (DataInterface.py:540), so PTR/MPR never misfire.
    Disabled transforms are never FTR.
    """
    if not header_info.get("enabled"):
        return False
    cells = transformed_row
    test_num = cells[0] if cells else None
    if isinstance(test_num, int):
        full_name = _stat_full_test_name(cells, header_info)
        if (test_num, full_name) in ftr_tests:
            return True
    header = header_info.get("header") or []
    for index, cell in enumerate(header):
        if _norm_header_text(cell) == _STAT_PATTERN_NAME:
            if len(cells) > index:
                value = cells[index]
                if isinstance(value, str) and value != "":
                    return True
            break
    return False


def apply_ftr_pseudo_limits(stat_row, header_info) -> None:
    """Overwrite the row's SpecMin/SpecMax cells in place with the DUT
    Summary's synthetic FTR floats (-0.999999 / 0.999999, matrix.FTR_SPEC_MIN/
    FTR_SPEC_MAX). No-op when the limit anchors are missing (translated UI)
    or a degraded row is shorter than the spec columns."""
    spec_cols = header_info.get("spec_cols")
    if spec_cols is None:
        return
    min_index, max_index = spec_cols
    if min_index < len(stat_row):
        stat_row[min_index] = FTR_SPEC_MIN
    if max_index < len(stat_row):
        stat_row[max_index] = FTR_SPEC_MAX


def _is_number(value) -> bool:
    """Real numeric cell (bools and NaN excluded)."""
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return False
    return value == value  # NaN != NaN


def _stat_formula_values(stat_row: list, header_info: dict) -> dict:
    """Cached Cp/Cpk results for one transformed Test Statistics row,
    keyed by 1-based column: the same values the live formulas compute
    (sigma/mu from the row's numericized Average/St. Dev. cells, limits
    from the EFFECTIVE SpecMin/SpecMax cells — FTR pseudo-limits already
    applied by ``apply_ftr_pseudo_limits``). Uncomputable -> "" so the
    cache stays blank (never xlsxwriter's default cached 0)."""
    cp_col = header_info.get("cp_col")
    cpk_col = header_info.get("cpk_col")
    spec_cols = header_info.get("spec_cols")
    if cp_col is None or cpk_col is None or spec_cols is None:
        return {}

    def _cell(index):
        value = stat_row[index] if index < len(stat_row) else None
        return value if _is_number(value) else None

    spec_min = _cell(spec_cols[0])
    spec_max = _cell(spec_cols[1])
    avg = _cell(header_info.get("avg_col"))
    std = _cell(header_info.get("std_col"))
    cp_cache = stat_cp_value(spec_min, spec_max, std)
    cpk_cache = stat_cpk_value(spec_min, spec_max, avg, std)
    return {
        cp_col + 1: "" if cp_cache is None else cp_cache,
        cpk_col + 1: "" if cpk_cache is None else cpk_cache,
    }


class TestStatisticsMixin:
    """Test Statistics worksheet writer (moved verbatim from generator.py;
    the former staticmethod `_stat_formula_values` demotes to this module's
    `_stat_formula_values` function)."""

    def writeStatistic(self, testTuples: list, heads: list[int], sites: list[int]):
        # Sheet for statistics of test items, e.g. cpk, mean, etc.
        StatSheet = self.sheetDict.get(ReportSelection.Stat, None)
        if StatSheet is None:
            return

        # re-measure filter (§7.2): Test Statistics honors the same kept-DUT
        # set as the DUT Summary — the stored drop-set when the DUT sheet
        # ran, a light dutInfo-only fetch otherwise (stat-only exports)
        kargs = {"testTuples": testTuples}
        if getSetting().gen.hide_remeasured:
            drops = self._resolve_remeasure_drops()
            if drops:
                kargs["dut_exclusion"] = drops_to_exclusion(drops)

        dataTable = self.waitForDataList(
            heads, sites, self.selectedFiles, tab.Trend, kargs
        )
        if not dataTable:
            return
        # the GUI sends the (already translated) header row as the first
        # element: [""] + headers — transform it for the modern sheet
        # (third wave): the "" VHeader cell splits into Test#/Head/Sites/
        # File ID, Low/High Limit rename to SpecMin/SpecMax, "Param" joins
        # after "Test Name". English-only anchors — translated UIs degrade
        # to untransformed columns (same ruling as the File Info wave).
        headerRow, dataRows = dataTable[0], dataTable[1:]
        header_info = transform_stat_header(headerRow)

        # SpecMin/SpecMax stylization (B4): the transformed header gives the
        # two column indices -> letters; the spec carries the scientific
        # format + static yellow fill (render_plain_table's column_fills
        # mechanism) and the per-row loop below applies the same to the rows
        # it writes itself (the classic per-row cadence bypasses spec rows).
        number_formats: dict[str, str] = {}
        column_fills: dict[str, str] = {}
        if header_info["spec_cols"] is not None:
            for column_index in header_info["spec_cols"]:
                letter = get_column_letter(column_index + 1)
                number_formats[letter] = STAT_SPEC_FORMAT
                column_fills[letter] = STAT_SPEC_FILL_HEX
        # Cpk display format (fifth wave, controller ruling): "0.00" on the
        # formula column, letter computed from the transformed header like
        # the spec_cols formats above; ninth wave: the same "0.00" on the
        # inserted Cp formula column
        cpk_col = header_info.get("cpk_col")
        if cpk_col is not None:
            number_formats[get_column_letter(cpk_col + 1)] = STAT_CPK_FORMAT
        cp_col = header_info.get("cp_col")
        if cp_col is not None:
            number_formats[get_column_letter(cp_col + 1)] = STAT_CPK_FORMAT

        # Fail Num / Cpk conditional formatting (fourth wave): letters from
        # the transformed header, data rows only — render_plain_table adds
        # the rules from the spec
        cf_ranges = stat_cf_ranges(header_info, len(dataRows))

        # header only (identity tr: no re-translation of the GUI header)
        spec = plain_table_spec(
            self.tr("Test Statistics"), header_info["header"], [], tr=lambda s: s
        )
        spec["number_formats"] = number_formats
        spec["column_fills"] = column_fills
        spec["tab_color"] = TAB_COLOR_HEXES["stat"]
        spec["repeat_header_row"] = 1
        if cf_ranges:
            spec["conditional_formatting"] = cf_ranges
        render_plain_table(StatSheet, spec, self._theme)

        # data rows keep the classic cadence: forceQuit check, one loopCount
        # and one sendProgress per row — the transforms happen per row
        # before writing, without altering the progress semantics
        written_rows = []
        for offset, dataRow in enumerate(dataRows):
            if self.forceQuit:
                return
            stat_row = transform_stat_row(dataRow, header_info)
            if stat_row_is_ftr(stat_row, header_info, self._ftr_tests):
                apply_ftr_pseudo_limits(stat_row, header_info)
            # live Cpk formula (fifth wave): AFTER the pseudo-limits, so FTR
            # rows compute from the ±0.999999 synthetic limits exactly like
            # the DUT Summary; row_number = 2 + data_row_index (header row 1).
            # Ninth wave: the Cp formula lands on the placeholder cell the
            # row transform inserted, reading σ from the numeric St. Dev.
            # column (same limits feed both ratios).
            apply_cpk_formula(stat_row, header_info, 2 + offset)
            apply_cp_formula(stat_row, header_info, 2 + offset)
            # cached formula results (xlsxwriter migration): the same
            # values the Cp/Cpk formulas compute — σ/μ from the row's
            # numericized cells, limits from the EFFECTIVE spec cells
            # (FTR pseudo-limits already applied); uncomputable -> ""
            formula_values = _stat_formula_values(stat_row, header_info)
            written_rows.append(stat_row)
            write_plain_row(
                StatSheet,
                2 + offset,
                stat_row,
                self._theme,
                band=(theme_group_fill_hex if offset % 2 == 0 else None),
                number_formats=number_formats,
                column_fills=column_fills,
                formula_values=formula_values,
            )
            self._row[ReportSelection.Stat] = 2 + offset
            self.loopCount += 1
            self.sendProgress()

        # extend the auto-filter over the written rows (render_plain_table
        # only knew about the header when it drew the sheet)
        width = max([len(header_info["header"])] + [len(row) for row in written_rows])
        StatSheet.autofilter(0, 0, len(written_rows), width - 1)
        self._autofit_columns(StatSheet)
        # sixth wave: value columns G..Q get a compact fixed width (autofit
        # stays for A:F — the test-name column benefits)
        StatSheet.set_column(
            xlutil.xl_cell_to_rowcol(f"{STAT_VALUE_WIDTH_RANGE[0]}1")[1],
            xlutil.xl_cell_to_rowcol(f"{STAT_VALUE_WIDTH_RANGE[1]}1")[1],
            STAT_VALUE_COLUMN_WIDTH,
        )
