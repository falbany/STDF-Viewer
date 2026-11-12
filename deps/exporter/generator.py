"""
#
# generator.py - exporter package - STDF Viewer
#
# Author: falbany (fork)
# Created Date: November 9th 2026
# -----
# Last Modified: Tue Nov 10 2026
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
# Adapted from uic_stdExporter.py by noonchen (Copyright (c) 2020 noonchen)
#

Modern xlsxwriter report generator: plain, DUT Summary and chart sheets.

``ModernReportGenerator`` subclasses the classic ``reportGenerator`` from
``deps.uic_stdExporter`` and keeps its threading contract verbatim — the
inherited ``generate_report`` slot, the ``waitFor*`` data handshake, the
``forceQuit`` early-exits and the ``loopCount`` / ``sendProgress`` cadence —
while writing the workbook with xlsxwriter through the exporter theme and
renderers.

Ninth wave (2026-11-10, A4 print design): the Bin Chart sheet gains a
merged A:F "Title" banner (``_bin_banner`` — "{Product ID} / {Sublot ID}"
extracted from the File Info table by ``bin_banner_from_file_info``, or
the sheet-name fallback), a "Headline 1" "Charts" band above the donuts
and a "Detailed Bin Contribution - HW Bin / SW Bin" band above the two
tables. ``report_core`` also pins the workbook default (Normal) font to
the theme's Times New Roman 8 via ``set_workbook_default_font``.

Tenth wave (2026-11-10, layout compaction): the banner wraps in a 50 pt
row 1, the per-block title row is gone (KPI cards one blank row under
the banner), the "Bin Charts" headline carries a blank row on both
sides, the contribution headline a blank row below, and the charts band
is one 250 pt row with 8.5 cm donuts cell-anchored A{row} / D{row}.

Instance state (``self._wb``, ``self._theme``, ``self._row``) replaces the
classic module-level ``sv`` statics, so several generators can coexist.
The DUT Summary matrix is rendered through the Task 4/5 spec builder and
matrix renderer (die chunks beyond the Excel column cap get their own
"(n)" sheets); the chart sheets (bin, wafer, trend, histogram) embed the
GUI-provided PNGs and write their stat rows with the themed body style,
mirroring the classic cursor gaps and progress cadence.
"""

from itertools import product
import re

import xlsxwriter
import xlsxwriter.utility as xlutil
from PyQt5 import QtCore

from deps.SharedSrc import REC, getSetting, tab  # fork(falbany): getSetting for the re-measure filter
from deps.uic_stdExporter import (
    ReportSelection,
    dataChannel,
    reportGenerator,
    signals,
)
from deps.exporter.charts import (
    BIN_BANNER_ROW_HEIGHT_PT,
    BIN_CHART_BAND_HEIGHT_PT,
    BIN_CHARTS_HEADLINE,
    BIN_CONTRIBUTION_HEADLINE,
    BIN_PAGE_COLUMN_WIDTH,
    BIN_PAGE_WIDTH_RANGE,
    aggregate_bin_counts,
    aggregate_kpi,
    bin_table_ranges,
    build_bin_donut,
    complete_chart_flags,
    write_bin_table,
    write_kpi_cards,
    write_merged_headline,
)
from deps.exporter.helpers import (
    apply_cp_formula,
    apply_cpk_formula,
    format_temperature_label,
    get_column_letter,
    parse_temperature_text,
    split_test_name,
    stat_cp_value,
    stat_cpk_value,
)
from deps.exporter.worksheets.dut_summary import (
    FTR_SPEC_MAX,
    FTR_SPEC_MIN,
    build_dut_summary_spec,
    group_order_by_temperature,
    prune_summary_view,
)
from deps.exporter.dedup import (  # fork(falbany): re-measure filter
    compute_remeasure_drops,
    dut_info_from_partial,
    drops_to_exclusion,
)
from deps.exporter.sheets import (
    insert_report_image,
    plain_table_spec,
)
from deps.exporter.theme import ModernTheme, TAB_COLOR_HEXES
from deps.exporter.writer import (
    FOOTER_TEXT,
    apply_autofit,
    render_matrix_sheet,
    render_plain_table,
    reset_widths,
    write_body_cell,
    write_plain_row,
)

# auto-fit column width = max content length * 1.1, capped (classic uses the
# same *1.1 factor via xlsxwriter set_column; the cap is the modern bound).
# Plain sheets only — the DUT Summary widths are pinned by the renderer.
AUTOFIT_FACTOR = 1.1
AUTOFIT_CAP = 60

# File Info rows whose values must be written as numeric cells (fix wave 2).
# Labels arrive TRANSLATED from the main window (getFileInfoForReport) with
# a trailing space after the colon; matching runs on the whitespace-
# normalized label, so it fires in the English locale — under other locales
# the values degrade to their classic string cells (harmless).
FILE_INFO_NUMERIC_LABELS = frozenset(
    {
        "Yield:",
        "DUTs Tested:",
        "DUTs Passed:",
        "DUTs Failed:",
        "DUTs Superseded:",
        "DUTs Unknown:",
    }
)
YIELD_FORMAT = "0.00%"

# Group-banding fill hex (theme token) passed to write_plain_row as a
# plain string — xlsxwriter formats compose per cell.
theme_group_fill_hex = "E9EEF6"


# English-only File Info anchors feeding the Bin-sheet banner (ninth
# wave): whitespace-normalized label, trailing ":" stripped.
_BANNER_PRODUCT_LABEL = "Product ID"
_BANNER_SUBLOT_LABEL = "Sublot ID"

# English-only File Info anchor feeding the DUT Summary temperature
# fallback: the MIR TST_TEMP values ("Test Temperature: " row, one cell
# per selected file column).
_FILE_INFO_TEMPERATURE_LABEL = "Test Temperature"


def order_dut_block(worksheets: list, block: list) -> list:
    """Stable reorder placing the DUT-derived `block` contiguously directly
    after its first member (the base DUT sheet) — §7.4's tab-order ruling.

    The base sheet is already in `worksheets` at its creation position;
    chunk-overflow / split sheets sit at the workbook end. Everything
    outside the block keeps its relative order. Identity when the block
    has no more than one member (base-only exports render in place) or
    nothing to place."""
    if len(block) <= 1:
        return worksheets
    base, derived = block[0], block[1:]
    base_at = worksheets.index(base)
    without_derived = [ws for ws in worksheets if ws not in derived]
    return without_derived[: base_at + 1] + derived + without_derived[base_at + 1 :]


def bin_banner_from_file_info(rows: list) -> "str | None":
    """ "{Product ID} / {Sublot ID}" per file column, files joined " | "
    (ninth wave banner text for the Bin sheet's Title row).

    Labels match whitespace-normalized with a trailing ":" stripped,
    English-only — ``None`` when either anchor is missing (translated
    UIs) or no file column carries both values (blank fields), so the
    caller falls back to the sheet name. Never raises.
    """
    products: list = []
    sublots: list = []
    for row in rows:
        label = " ".join(str(row[0] if row else "").split())
        if label.endswith(":"):
            label = label[:-1]
        if label == _BANNER_PRODUCT_LABEL:
            products = list(row[1:])
        elif label == _BANNER_SUBLOT_LABEL:
            sublots = list(row[1:])
    pairs = []
    for index in range(max(len(products), len(sublots))):
        product = products[index] if index < len(products) else ""
        sublot = sublots[index] if index < len(sublots) else ""
        product = product.strip() if isinstance(product, str) else ""
        sublot = sublot.strip() if isinstance(sublot, str) else ""
        if product and sublot:
            pairs.append(f"{product} / {sublot}")
    return " | ".join(pairs) if pairs else None


def file_temperatures_from_file_info(rows: list, fids: list) -> dict:
    """{fid: temperature float | None} from the File Info table's
    "Test Temperature:" row (MIR TST_TEMP per file column).

    Values are position-matched with `fids` (getFileInfoForReport's
    per-file value order — the same contract as the "File #n" header).
    Labels match whitespace-normalized with a trailing ":" stripped,
    English-only like the banner anchors — a translated UI (or a missing
    row) yields all ``None`` so the DUT Summary keeps its blank cells.
    Composite multi-subfile cells parse to the first subfile's value.
    """
    values: list = []
    for row in rows:
        label = " ".join(str(row[0] if row else "").split())
        if label.endswith(":"):
            label = label[:-1]
        if label == _FILE_INFO_TEMPERATURE_LABEL:
            values = list(row[1:])
            break
    return {
        fid: parse_temperature_text(values[index] if index < len(values) else None)
        for index, fid in enumerate(fids)
    }


# --------------------------------------------------------------------------
# Test Statistics transforms (third improvement wave, 2026-11-10)
# --------------------------------------------------------------------------

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


def _numeric_file_info_cell(value):
    """(cell_value, number_format) for one File Info metadata value.

    Count values arrive as str(int) -> int cell (unconvertible strings are
    kept); Yield arrives as "98.50%" -> 0.985 with the "0.00%" display
    format (numeric AND showing 98.50%); "?" (no DUTs) and any
    unconvertible value stay strings.
    """
    text = str(value)
    if text.endswith("%"):
        try:
            return float(text[:-1]) / 100, YIELD_FORMAT
        except ValueError:
            return value, None
    try:
        return int(text), None
    except ValueError:
        return value, None


def _is_number(value) -> bool:
    """Real numeric cell (bools and NaN excluded)."""
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return False
    return value == value  # NaN != NaN


class ModernReportGenerator(reportGenerator):
    def __init__(
        self,
        signals: signals,
        mutex: QtCore.QMutex,
        conditionWait: QtCore.QWaitCondition,
        reportInfo: tuple,
        channel: dataChannel,
    ):
        super().__init__(signals, mutex, conditionWait, reportInfo, channel)
        # instance state replacing the classic module-level sv statics
        self._wb: xlsxwriter.Workbook | None = None
        self._theme = None  # ModernTheme binds to the workbook in report_core
        self._row: dict[ReportSelection, int] = {}  # per-sheet row cursors

    # ----------------------------------------------------------------------
    # helpers
    # ----------------------------------------------------------------------

    def _autofit_columns(self, ws) -> None:
        """Apply the writer's width ledger as ``max content length * 1.1``
        capped at AUTOFIT_CAP (xlsxwriter cells cannot be re-read, so the
        renderer notes widths at write time)."""
        apply_autofit(ws, AUTOFIT_FACTOR, AUTOFIT_CAP)

    def _write_body_row(self, ws, row: int, values: list, fill=None) -> None:
        """One themed body row at 1-based ``row``: body font/alignment and
        thin border via writer.write_body_cell, optional band fill hex."""
        for column, value in enumerate(values, start=1):
            write_body_cell(ws, row, column, value, self._theme, fill=fill)

    @staticmethod
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

    # ----------------------------------------------------------------------
    # plain sheets (Task 7)
    # ----------------------------------------------------------------------

    def writeFileInfo(self):
        InfoSheet = self.sheetDict.get(ReportSelection.FileInfo, None)
        if InfoSheet is None:
            return

        # get file info table
        dataTable = self.waitForTableData(isFileInfo=True)
        if self.forceQuit:
            return

        # numeric metadata rows (fix wave 2): labels are translated
        # upstream with a trailing space after the colon — match on the
        # whitespace-normalized label; count values become int cells, Yield
        # becomes a 0.xx fraction with the "0.00%" display format
        cell_number_formats: dict[tuple[int, int], str] = {}
        rows = []
        for row_index, row in enumerate(dataTable):
            label = row[0] if row else ""
            if " ".join(str(label).split()) not in FILE_INFO_NUMERIC_LABELS:
                rows.append(row)
                continue
            converted = [label]
            for col_index, value in enumerate(row[1:], start=1):
                cell, fmt = _numeric_file_info_cell(value)
                converted.append(cell)
                if fmt is not None:
                    cell_number_formats[(row_index, col_index)] = fmt
            rows.append(converted)

        # header spans every selected file column (fix wave 2): one
        # "File #n" per selected fid, position-matched with
        # getFileInfoForReport's per-file value order — the labels carry
        # the ACTUAL file ids (selecting files 0 and 2 -> "File #1",
        # "File #3"); applied unconditionally (single file -> "File #1")
        spec = plain_table_spec(
            self.tr("File Info"),
            [self.tr("Property Name")]
            + [f"File #{fid + 1}" for fid in self.selectedFiles],
            rows,
            self.tr,
            cell_number_formats=cell_number_formats,
        )
        spec["tab_color"] = TAB_COLOR_HEXES["file_info"]
        spec["repeat_header_row"] = 1
        render_plain_table(InfoSheet, spec, self._theme)
        self._autofit_columns(InfoSheet)

        # ninth wave: stash the Bin-sheet banner text when the English
        # Product ID / Sublot ID anchors resolve (missing labels keep the
        # report_core fallback — translated UIs degrade)
        banner = bin_banner_from_file_info(dataTable)
        if banner:
            self._bin_banner = banner

        self.loopCount += 1
        self.sendProgress()

    def writeGDR_DTR(self):
        GDR_DTR_Sheet = self.sheetDict.get(ReportSelection.GDR_DTR, None)
        if GDR_DTR_Sheet is None:
            return

        dataTable = self.waitForTableData(isFileInfo=False)
        if self.forceQuit:
            return
        # remove \n that added besides values
        dataTable = [[ele.strip("\n") for ele in dataRow] for dataRow in dataTable]

        spec = plain_table_spec(
            self.tr("GDR & DTR Summary"),
            [self.tr("Record Type"), self.tr("Value"), self.tr("Approx. Location")],
            dataTable,
            self.tr,
        )
        spec["tab_color"] = TAB_COLOR_HEXES["gdr"]
        spec["repeat_header_row"] = 1
        render_plain_table(GDR_DTR_Sheet, spec, self._theme)
        self._autofit_columns(GDR_DTR_Sheet)

        self.loopCount += 1
        self.sendProgress()

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
            formula_values = self._stat_formula_values(stat_row, header_info)
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

    # ----------------------------------------------------------------------
    # DUT Summary matrix + chart sheets (Task 8)
    # ----------------------------------------------------------------------

    def _resolve_remeasure_drops(self) -> "set | set()":
        """The filter's drop-set, either direction of the sheet selection.

        writeDUT_Summary stores its drop-set before writeStatistic runs
        (DUT sheet selected → reuse it; an empty set means "ran, nothing
        to drop" — never refetch). None means the DUT sheet did not
        render: the light dutInfo-only fetch answers, and an empty fetch
        degrades gracefully to an empty drop-set."""
        if self._remeasure_drops is not None:
            return self._remeasure_drops
        partial = self.waitForDutInfo(
            self.selectedHeads, self.selectedSites, self.selectedFiles
        )
        return compute_remeasure_drops(
            dut_info_from_partial(partial or {}), self.selectedFiles
        )

    def _order_and_ids(self, summaryTable, order_filter=None):
        """The classic `(fid, dutIndex)` display order and its per-column
        part IDs (moved verbatim out of writeDUT_Summary).

        `order_filter` (temperature split §7.4) keeps only the listed
        pairs — None keeps everything. The "DUT #n" fallback numbers
        within the KEPT order."""
        order = []
        for fid in sorted(summaryTable["dutInfo"].keys()):
            for dutIndex in sorted(summaryTable["dutInfo"][fid].keys()):
                pair = (fid, dutIndex)
                if order_filter is None or pair in order_filter:
                    order.append(pair)

        # full part ID per ordered DUT (classic dataRow[1] indexing); a
        # blank part ID falls back to "DUT #n"
        dut_ids = []
        for i, (fid, dutIndex) in enumerate(order):
            info = summaryTable["dutInfo"][fid][dutIndex]
            part_id = info[1] if len(info) > 1 and info[1] else ""
            dut_ids.append(part_id if part_id else f"DUT #{i + 1}")
        return order, dut_ids

    def writeDUT_Summary(self):
        DutSheet = self.sheetDict.get(ReportSelection.DUT, None)
        if DutSheet is None:
            return

        # a complex dict structure from the GUI thread
        # key: Data, value:     testTuple -> {fid, testDataDict}
        # key: TestInfo, value: testTuple -> testInfoDict
        # key: dut2ind, value:  fid -> {dutIndex -> dataIndex}
        # key: dutInfo, value:  fid -> {dutIndex -> dutSummary}
        summaryTable = self.waitForDutSummary(
            self.selectedHeads, self.selectedSites, self.selectedFiles, self.testTuples
        )
        if self.forceQuit:
            return

        # FTR cache (fourth wave): (test_num, full test name) of every test
        # tuple whose record SUB_CODE is FTR (any fid) — the PRIMARY FTR
        # detection signal for writeStatistic's pseudo-limits. Runs before
        # the sheet rendering; a Stat-only selection never reaches here and
        # keeps the report_core-initialized empty set (fallback signal).
        self._ftr_tests = set()
        for test_tuple in self.testTuples:
            fid_data = summaryTable.get("Data", {}).get(test_tuple) or {}
            if any(
                (fid_data.get(fid) or {}).get("SUB_CODE") == REC.FTR for fid in fid_data
            ):
                self._ftr_tests.add((test_tuple[0], test_tuple[2]))

        # re-measure filter (DUT Summary enhancements §7.2): the drop-set
        # is computed BEFORE rendering and stored for writeStatistic's
        # stats hook (Task 6) — the DUT sheet runs first in the report
        # sequence, so a DUT-selected export never refetches here. An
        # empty set means "ran, nothing to drop"; None (never stored)
        # means the DUT sheet did not render (stat-only fallback fetch).
        if getSetting().gen.hide_remeasured:
            drops = compute_remeasure_drops(summaryTable["dutInfo"], self.selectedFiles)
        else:
            drops = set()
        self._remeasure_drops = drops
        if drops:
            summaryTable = prune_summary_view(summaryTable, drops)

        order, dut_ids = self._order_and_ids(summaryTable)

        # per-file MIR Test Temperature (fix wave: the Temperature metadata
        # row falls back to it when a part ID carries no _T<n>C tag). The
        # File Info table is fetched even when that sheet is not selected —
        # the GUI side answers from its immutable fileMetaData; the File
        # Info sheet, when selected, fetched the identical table moments ago.
        file_info_rows = self.waitForTableData(isFileInfo=True)
        if self.forceQuit:
            return
        file_temperatures = file_temperatures_from_file_info(
            file_info_rows, self.selectedFiles
        )

        specs = build_dut_summary_spec(
            summaryTable,
            self.testTuples,
            order,
            dut_ids,
            sheet_name=self.tr("DUT Summary"),
            # numeric cell formats are hardcoded scientific in the spec
            # builder (2026-11-10) — no settings-derived float format anymore
            tr=self.tr,
            file_temperatures=file_temperatures,
        )

        # chunk 0 renders on the existing DUT Summary sheet; chunks 1+
        # (die overflow beyond the Excel column cap) get their own sheets —
        # spec names already carry the " (n)" suffixes. Column widths are
        # pinned inside render_matrix_sheet (A=8, B=30, C..L=8, dies=8).
        # Every sheet created here joins _dut_sheet_block (base sheet
        # first, report_core appended it at creation) so the close-time
        # reorder keeps the DUT block adjacent to the base sheet.
        for chunk_index, spec in enumerate(specs):
            ws = DutSheet if chunk_index == 0 else self._wb.add_worksheet(spec["name"])
            render_matrix_sheet(ws, spec, self._theme)
            if chunk_index > 0:
                self._dut_sheet_block.append(ws)

        # temperature split (§7.4): one extra worksheet per distinct
        # temperature, ascending — additive; the base sheet above is
        # untouched. Groups of the filtered view only; DUTs with no
        # resolvable temperature stay in the base sheet only. No new
        # progress ticks (splitting adds sheets, not fetches).
        if getSetting().gen.split_temperature:
            groups = group_order_by_temperature(order, summaryTable["dutInfo"], file_temperatures)
            resolved = [temp for temp in groups if temp is not None]
            if len(resolved) >= 2:
                # sorted(resolved), not sorted(groups): the None key does
                # not compare against floats
                for temp in sorted(resolved):
                    pairs = set(groups[temp])
                    subset = prune_summary_view(summaryTable, set(order) - pairs)
                    sub_order, sub_ids = self._order_and_ids(subset, order_filter=pairs)
                    split_specs = build_dut_summary_spec(
                        subset,
                        self.testTuples,
                        sub_order,
                        sub_ids,
                        sheet_name=f'{self.tr("DUT Summary")} - T{format_temperature_label(temp)}°C',
                        tr=self.tr,
                        file_temperatures=file_temperatures,
                    )
                    for spec in split_specs:
                        ws = self._wb.add_worksheet(spec["name"])
                        render_matrix_sheet(ws, spec, self._theme)
                        self._dut_sheet_block.append(ws)

        # classic cadence: one tick after the DUT info part...
        self.loopCount += 1
        self.sendProgress()
        # ...and one per test column (rows span every chunk)
        for _ in self.testTuples:
            self.loopCount += 1
            self.sendProgress()

    def writeBinChart(self):
        BinSheet = self.sheetDict.get(ReportSelection.Bin, None)
        if BinSheet is None:
            return

        # ninth wave: row 1 carries the merged A:F Title-lookalike banner
        # (Product ID / Sublot ID from File Info, or the sheet-name
        # fallback); tenth wave: the banner wraps in a 50 pt row so a
        # 2-line title fits, the per-block title row is GONE and the KPI
        # cards sit one blank row below the banner
        BinSheet.merge_range(
            0,
            0,
            0,
            5,
            self._bin_banner,
            self._theme.lookalike("title", wrap=True),
        )
        BinSheet.set_row(0, BIN_BANNER_ROW_HEIGHT_PT)
        key = ReportSelection.Bin
        row = 2
        for head, site in product(self.selectedHeads, self.selectedSites):
            if self.forceQuit:
                return
            # get hard bin list, then soft bin list — the same
            # waitForDataList calls as before; each table feeds BOTH the
            # donut chart and the structured table; the HB row list also
            # feeds the KPI cards (upstream writes the SAME summary cells
            # into the SB rows — aggregating one list avoids double-count)
            tables = [
                self.waitForDataList(
                    [head], [site], self.selectedFiles, tab.Bin, {"isHBIN": isHBIN}
                )
                for isHBIN in (True, False)
            ]
            aggregated = [aggregate_bin_counts(dataTable) for dataTable in tables]
            kpi = aggregate_kpi(tables[0])
            n_files = len(tables[0])
            site_label = "All Sites" if site == -1 else f"Site{site}"
            site_tag = "AllSites" if site == -1 else f"Site{site}"

            # KPI cards (label/value/subtitle, each with a dark accent
            # outline riding the row-merge format) after ONE blank
            # separator row — the head/site identification lives in the
            # chart and table titles only
            row = write_kpi_cards(BinSheet, row + 1, kpi, n_files, self._theme)

            # "Bin Charts" headline (tenth wave rename) with a blank row
            # on BOTH sides, then the charts band: ONE 250 pt row with
            # the 8.5 cm donuts cell-anchored HB at A{chart_row} and SB
            # at D{chart_row} (user ruling); empty bin lists leave the
            # band reserved, keeping the dashboard above intact
            row = write_merged_headline(
                BinSheet, row + 1, BIN_CHARTS_HEADLINE, "headline1", self._theme
            )
            row += 1  # blank between the headline and the band
            chart_row = row
            BinSheet.set_row(chart_row - 1, BIN_CHART_BAND_HEIGHT_PT)
            # one blank separator after the band, then the "Detailed Bin
            # Contribution" headline (ninth wave) with a blank row below it
            # (tenth wave) — the HB and SW tables stack below with a
            # single blank row between them. The TABLES are written FIRST
            # (2026-11-11 user ruling): the donuts reference the visible
            # table ranges (Bin column B / Count column D), so the charts
            # display exactly the data in the tables below them
            row = write_merged_headline(
                BinSheet,
                chart_row + 2,
                BIN_CONTRIBUTION_HEADLINE,
                "headline1",
                self._theme,
            )
            row += 1  # blank below the contribution headline
            section_refs = []
            for index, (isHBIN, bins) in enumerate(zip((True, False), aggregated)):
                kind = "Hardware" if isHBIN else "Software"
                prefix = "HB" if isHBIN else "SB"
                title_row = row
                row = write_bin_table(
                    BinSheet,
                    row,
                    bins,
                    self._theme,
                    f"{kind} Bins — Head{head} / {site_label}",
                    f"{prefix}_Head{head}_{site_tag}",
                )
                section_refs.append(
                    bin_table_ranges(title_row, len(bins)) if bins else None
                )
                if index == 0:
                    row += 1  # single blank separator before the SB section

            # the donuts sit in the charts band ABOVE their tables, fed by
            # them (write order is irrelevant — refs are strings; xlsxwriter
            # resolves the numCache from the already-written cells at close)
            for (kind, anchor_letter), refs, bins in zip(
                (("Hardware", "A"), ("Software", "D")), section_refs, aggregated
            ):
                if refs is None:
                    continue
                name_cell, cat_range, val_range = refs
                build_bin_donut(
                    self._wb,
                    BinSheet,
                    f"{kind} Bin — Head{head} / {site_label}",
                    bins,
                    f"{anchor_letter}{chart_row}",
                    name_cell,
                    cat_range,
                    val_range,
                )
            # block end: the cursor sits on the blank row after the SB
            # Total; the next block's KPI cards start after it
        self._row[key] = row

        self.loopCount += 1
        self.sendProgress()
        # gridlines off (every other sheet's look; the chart sheets bypass
        # render_plain_table, which is where the plain sheets get it)
        BinSheet.hide_gridlines(2)
        BinSheet.set_footer(FOOTER_TEXT)
        BinSheet.set_tab_color(f"#{TAB_COLOR_HEXES['bin']}")
        self._autofit_columns(BinSheet)
        # eighth wave (R2), tenth-wave width: the Bin sheet is a fixed
        # A:F page — all six columns share width 15 (KPI grid + tables;
        # NOTE the Calibri-11 metric recalibration is a USER CHECKLIST item:
        # xlsxwriter stores width + ~0.71 vs the openpyxl era)
        first_col, last_col = (
            xlutil.xl_cell_to_rowcol(f"{letter}1")[1] for letter in BIN_PAGE_WIDTH_RANGE
        )
        BinSheet.set_column(first_col, last_col, BIN_PAGE_COLUMN_WIDTH)
        # eighth wave (R1): the Bin sheet prints on A4, fit to one page
        # wide with unlimited height (same fitToPage pattern as the DUT
        # Summary matrix setup in writer.py); other sheets are untouched
        BinSheet.set_paper(9)  # A4
        BinSheet.fit_to_pages(1, 0)

    def writeWaferMap(self):
        WaferSheet = self.sheetDict.get(ReportSelection.Wafer, None)
        if WaferSheet is None:
            return

        key = ReportSelection.Wafer
        row = self._row.get(key, 0)
        for wafer in self.waferTuples:
            if self.forceQuit:
                return
            # wafer map is already associated with head
            image_io = self.waitForImage(
                wafer, 0, self.selectedSites, self.selectedFiles, tab.Wafer
            )
            row += insert_report_image(WaferSheet, image_io, row) + 2
            # add wafer info
            dataTable = self.waitForDataList(
                [],
                self.selectedSites,
                self.selectedFiles,
                tab.Wafer,
                {"testTuples": [wafer]},
            )
            for dataList in dataTable:
                self._write_body_row(WaferSheet, row, dataList)
                row += 1
            # gap between images
            row += 3
            self.loopCount += 1
            self.sendProgress()
        self._row[key] = row
        self._autofit_columns(WaferSheet)
        WaferSheet.hide_gridlines(2)
        WaferSheet.set_footer(FOOTER_TEXT)
        WaferSheet.set_tab_color(f"#{TAB_COLOR_HEXES['wafer']}")

    def writeTrendPlot(self, testTuple: tuple, head: int, site: int):
        TrendSheet = self.sheetDict.get(ReportSelection.Trend, None)
        if TrendSheet is None:
            return

        # get image and stat from main thread
        image_io = self.waitForImage(
            testTuple, head, [site], self.selectedFiles, tab.Trend
        )
        dataTable = self.waitForDataList(
            [head], [site], self.selectedFiles, tab.Trend, {"testTuples": [testTuple]}
        )
        # insert into the work sheet
        key = ReportSelection.Trend
        row = self._row.get(key, 0)
        row += insert_report_image(TrendSheet, image_io, row)
        for dataList in dataTable:
            self._write_body_row(TrendSheet, row, dataList)
            row += 1
        row += 2
        self._row[key] = row

        self.loopCount += 1
        self.sendProgress()
        self._autofit_columns(TrendSheet)
        TrendSheet.hide_gridlines(2)
        TrendSheet.set_footer(FOOTER_TEXT)
        TrendSheet.set_tab_color(f"#{TAB_COLOR_HEXES['trend']}")

    def writeHistogram(self, testTuple: tuple, head: int, site: int):
        HistoSheet = self.sheetDict.get(ReportSelection.Histo, None)
        if HistoSheet is None:
            return

        image_io = self.waitForImage(
            testTuple, head, [site], self.selectedFiles, tab.Histo
        )
        dataTable = self.waitForDataList(
            [head], [site], self.selectedFiles, tab.Histo, {"testTuples": [testTuple]}
        )
        key = ReportSelection.Histo
        row = self._row.get(key, 0)
        row += insert_report_image(HistoSheet, image_io, row)
        for dataList in dataTable:
            self._write_body_row(HistoSheet, row, dataList)
            row += 1
        row += 2
        self._row[key] = row

        self.loopCount += 1
        self.sendProgress()
        self._autofit_columns(HistoSheet)
        HistoSheet.hide_gridlines(2)
        HistoSheet.set_footer(FOOTER_TEXT)
        HistoSheet.set_tab_color(f"#{TAB_COLOR_HEXES['histo']}")

    # ----------------------------------------------------------------------
    # report assembly
    # ----------------------------------------------------------------------

    def report_core(self):
        # strings_to_urls MUST stay off: URL-looking strings (File Info
        # paths) would silently become hyperlinks (xlsxwriter's default)
        self._wb = xlsxwriter.Workbook(
            self.reportPath, {"strings_to_urls": False, "in_memory": False}
        )
        self._theme = ModernTheme(self._wb)
        reset_widths()
        self.sheetDict = {}
        self.loopCount = 0
        self._row = {}
        # Bin-sheet banner (ninth wave): the report_core default is the
        # sheet name; writeFileInfo overwrites it with
        # "{Product ID} / {Sublot ID}" when the File Info sheet is
        # selected and the English anchors resolve
        self._bin_banner = self.tr("Bin Chart")
        # FTR cache (fourth wave) initialized BEFORE any writer runs:
        # writeDUT_Summary fills it when the DUT sheet is selected; a
        # Stat-only selection leaves it empty (the Pattern-Name fallback
        # covers detection there)
        self._ftr_tests: set = set()
        # re-measure filter hand-off (§7.2): None until writeDUT_Summary
        # stores its drop-set — None (vs an empty set) is the stat-only
        # signal that triggers the light dutInfo-only fetch
        self._remeasure_drops: "set | None" = None
        # the DUT-derived sheet block (base + chunk overflow + temperature
        # splits, base first) for the close-time tab-order reorder (§7.4)
        self._dut_sheet_block: list = []

        # the try/finally mirrors the classic `with Workbook(...)` block:
        # the (partial) workbook is saved on forceQuit early-returns and on
        # exceptions alike; only the success path emits closeSignal
        try:
            # 1. create sheets for each selected contents, in the fixed
            #    classic order; 2./3./4. below write them in classic order
            for cont in [
                ReportSelection.FileInfo,
                ReportSelection.DUT,
                ReportSelection.GDR_DTR,
                ReportSelection.Stat,
                ReportSelection.Trend,
                ReportSelection.Histo,
                ReportSelection.Bin,
                ReportSelection.Wafer,
                ReportSelection.PPQQ,
                ReportSelection.Correlation,
            ]:
                # sheet order in the xlsx is fixed
                if cont == ReportSelection.FileInfo:
                    sheetName = self.tr("File Info")
                elif cont == ReportSelection.DUT:
                    sheetName = self.tr("DUT Summary")
                elif cont == ReportSelection.GDR_DTR:
                    sheetName = self.tr("GDR & DTR Summary")
                elif cont == ReportSelection.Stat:
                    sheetName = self.tr("Test Statistics")
                elif cont == ReportSelection.Trend:
                    sheetName = self.tr("Trend Chart")
                elif cont == ReportSelection.Histo:
                    sheetName = self.tr("Histogram")
                elif cont == ReportSelection.Bin:
                    sheetName = self.tr("Bin Chart")
                elif cont == ReportSelection.Wafer:
                    sheetName = self.tr("Wafer Map")
                elif cont == ReportSelection.PPQQ:
                    sheetName = self.tr("Normal Validation")
                else:
                    sheetName = self.tr("Correlation")
                if cont in self.selectedContents:
                    self.sheetDict[cont] = self._wb.add_worksheet(sheetName)
                    if cont == ReportSelection.DUT:
                        # block[0] = the base DUT sheet (§7.4 reorder)
                        self._dut_sheet_block.append(self.sheetDict[cont])

            # ** write contents independent of test numbers
            # file info
            self.writeFileInfo()
            if self.forceQuit:
                return
            # GDR & DTR Summary
            self.writeGDR_DTR()
            if self.forceQuit:
                return
            # dut summary
            self.writeDUT_Summary()
            if self.forceQuit:
                return

            # bin sheet
            self.writeBinChart()
            if self.forceQuit:
                return
            # wafer sheet
            self.writeWaferMap()
            if self.forceQuit:
                return

            # ** write contents related to test numbers
            self.writeStatistic(self.testTuples, self.selectedHeads, self.selectedSites)
            if self.forceQuit:
                return

            for testTuple in self.testTuples:
                for head, site in product(self.selectedHeads, self.selectedSites):
                    if self.forceQuit:
                        return

                    self.writeTrendPlot(testTuple, head, site)
                    self.writeHistogram(testTuple, head, site)

                # add gap between different test items
                for key in (ReportSelection.Trend, ReportSelection.Histo):
                    self._row[key] = self._row.get(key, 0) + 2
        finally:
            # §7.4 placement ruling: sheets xlsxwriter appends mid-report
            # land at the workbook end — reposition the DUT-derived block
            # (base + chunk overflow + temperature splits) so it sits
            # adjacent to the base sheet. Tab order derives from the
            # worksheet list at close(), so the reposition is the whole
            # job. PRIVATE-LIST DEPENDENCY: Workbook.worksheets_objs is
            # pinned by TestDutBlockTabOrder (pure + integration) —
            # re-verify on xlsxwriter version bumps.
            if self._dut_sheet_block:
                self._wb.worksheets_objs = order_dut_block(
                    self._wb.worksheets_objs, self._dut_sheet_block
                )
            self._wb.close()
            # xlsxwriter gap completion (charts.py): inject the data-label
            # show* zeros + <c:plotVisOnly val="0"/> it cannot emit — run on
            # every exit path so even a forceQuit partial file stays clean
            try:
                complete_chart_flags(self.reportPath)
            except OSError:
                pass  # a failed completion must not mask the report itself

        self.closeSignal.emit(True)
