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

Modern openpyxl report generator: plain, DUT Summary and chart sheets.

``ModernReportGenerator`` subclasses the classic ``reportGenerator`` from
``deps.uic_stdExporter`` and keeps its threading contract verbatim — the
inherited ``generate_report`` slot, the ``waitFor*`` data handshake, the
``forceQuit`` early-exits and the ``loopCount`` / ``sendProgress`` cadence —
while writing the workbook with openpyxl through the exporter theme and
renderers instead of xlsxwriter.

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
is one 250 pt row with square ``band_anchor()``-pinned donuts.

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

from openpyxl import Workbook
from openpyxl.styles import Alignment
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties
from PyQt5 import QtCore

from deps.SharedSrc import REC, tab
from deps.uic_stdExporter import (
    ReportSelection,
    dataChannel,
    reportGenerator,
    signals,
)
from deps.exporter.charts import (
    BIN_BANNER_ROW_HEIGHT_PT,
    BIN_CHART_BAND_HEIGHT_PT,
    BIN_CHART_GAP_CM,
    BIN_CHART_SIZE_CM,
    BIN_CHARTS_HEADLINE,
    BIN_CONTRIBUTION_HEADLINE,
    BIN_PAGE_COLUMN_WIDTH,
    BIN_PAGE_WIDTH_RANGE,
    aggregate_bin_counts,
    aggregate_kpi,
    band_anchor,
    bin_label,
    build_bin_donut,
    write_bin_table,
    write_kpi_cards,
    write_merged_headline,
)
from deps.exporter.helpers import apply_cpk_formula, split_test_name
from deps.exporter.matrix import (
    FTR_SPEC_MAX,
    FTR_SPEC_MIN,
    build_dut_summary_spec,
)
from deps.exporter.sheets import (
    insert_report_image,
    plain_table_spec,
)
from deps.exporter.theme import ModernTheme, set_workbook_default_font
from deps.exporter.writer import (
    render_matrix_sheet,
    render_plain_table,
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


# English-only File Info anchors feeding the Bin-sheet banner (ninth
# wave): whitespace-normalized label, trailing ":" stripped.
_BANNER_PRODUCT_LABEL = "Product ID"
_BANNER_SUBLOT_LABEL = "Sublot ID"


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

# Fixed width for the Test Statistics value columns G..P (sixth wave
# 2026-11-10): applied AFTER the sheet-wide autofit, which keeps sizing
# A:F (the test-name column benefits).
STAT_VALUE_WIDTH_RANGE = ("G", "P")
STAT_VALUE_COLUMN_WIDTH = 8

# Bin Chart pie helper blocks live in this fixed far column (labels) and
# the next one (counts); both are hidden by the chart factory. A fixed far
# column keeps them clear of the auto-fitted table columns (BJ/BK > cap 60).
BIN_HELPER_COL = 62

# English-only header anchors (whitespace-normalized matching).
_STAT_RENAMES = {"Low Limit": "SpecMin", "High Limit": "SpecMax"}
_STAT_TEST_NAME = "Test Name"
_STAT_FAIL_NUM = "Fail Num"
_STAT_CPK = "Cpk"
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
      (English-only anchor).

    Returns ``{"header": [...], "enabled": bool, "param_inserted": bool,
    "spec_cols": (i, j) | None, "avg_col": int | None, "std_col": int | None,
    "cpk_col": int | None}`` where ``spec_cols`` is the 0-based (SpecMin,
    SpecMax) column pair in the TRANSFORMED header (None when either limit
    anchor was missing — translated UIs degrade to untransformed columns)
    and the fifth-wave keys carry the 0-based Average / St. Dev. / Cpk
    anchors for the live Cpk formula (normalized matching, None when
    missing). An unexpected non-blank first cell disables the whole
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
            "cpk_col": None,
        }
    body = [str(cell) for cell in cells[1:]]
    normalized = [_norm_header_text(cell) for cell in body]
    renamed = [_STAT_RENAMES.get(norm, cell) for norm, cell in zip(normalized, body)]
    header = list(STAT_VHEADER_LABELS) + renamed
    param_inserted = _STAT_TEST_NAME in normalized
    if param_inserted:
        header.insert(4 + normalized.index(_STAT_TEST_NAME) + 1, "Param")
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
      MPR mix columns) goes through ``numericize_value``.
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
    return out


def stat_cf_ranges(header_info: dict, data_row_count: int) -> dict:
    """Conditional-formatting ranges for the Test Statistics sheet.

    Computes the Fail Num / Cpk column letters from the TRANSFORMED header
    (normalized English anchors — positions shift right of the inserted
    FTR "Pattern Name" / MPR mix columns, never hard-coded) and spans the
    DATA rows only: ``{"fail_num_range": "J2:Jn", "cpk_range": "K2:Kn"}``.
    An empty dict when either anchor is missing (translated-UI degradation),
    when the transform is disabled, or when there are no data rows.
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
    return {
        "fail_num_range": f"{fail_letter}2:{fail_letter}{last_row}",
        "cpk_range": f"{cpk_letter}2:{cpk_letter}{last_row}",
    }


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
        self._wb: Workbook | None = None
        self._theme = ModernTheme()
        self._row: dict[ReportSelection, int] = {}  # per-sheet row cursors

    # ----------------------------------------------------------------------
    # helpers
    # ----------------------------------------------------------------------

    def _autofit_columns(self, ws) -> None:
        """Set each used column's width to max content length * 1.1, capped
        at AUTOFIT_CAP (the modern counterpart of the classic
        ``set_column(col, col, width*1.1)`` calls). Plain sheets only — the
        DUT Summary matrix pins its column widths in the renderer."""
        widths: dict[int, int] = {}
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is None:
                    continue
                length = len(str(cell.value))
                if length > widths.get(cell.column, 0):
                    widths[cell.column] = length
        for column, length in widths.items():
            ws.column_dimensions[get_column_letter(column)].width = min(
                length * AUTOFIT_FACTOR, AUTOFIT_CAP
            )

    def _write_body_row(self, ws, row: int, values: list, fill=None) -> None:
        """One themed body row at 1-based ``row``: body font/alignment and
        thin border via writer.write_body_cell, optional band fill."""
        for column, value in enumerate(values, start=1):
            write_body_cell(ws, row, column, value, self._theme, fill=fill)

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
        render_plain_table(GDR_DTR_Sheet, spec, self._theme)
        self._autofit_columns(GDR_DTR_Sheet)

        self.loopCount += 1
        self.sendProgress()

    def writeStatistic(self, testTuples: list, heads: list[int], sites: list[int]):
        # Sheet for statistics of test items, e.g. cpk, mean, etc.
        StatSheet = self.sheetDict.get(ReportSelection.Stat, None)
        if StatSheet is None:
            return

        dataTable = self.waitForDataList(
            heads, sites, self.selectedFiles, tab.Trend, {"testTuples": testTuples}
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
        # the spec_cols formats above
        cpk_col = header_info.get("cpk_col")
        if cpk_col is not None:
            number_formats[get_column_letter(cpk_col + 1)] = STAT_CPK_FORMAT

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
            # the DUT Summary; row_number = 2 + data_row_index (header row 1)
            apply_cpk_formula(stat_row, header_info, 2 + offset)
            written_rows.append(stat_row)
            write_plain_row(
                StatSheet,
                2 + offset,
                stat_row,
                self._theme,
                band=(self._theme.group_fill if offset % 2 == 0 else None),
                number_formats=number_formats,
                column_fills=column_fills,
            )
            self._row[ReportSelection.Stat] = 2 + offset
            self.loopCount += 1
            self.sendProgress()

        # extend the auto-filter over the written rows (render_plain_table
        # only knew about the header when it drew the sheet)
        width = max([len(header_info["header"])] + [len(row) for row in written_rows])
        StatSheet.auto_filter.ref = (
            f"A1:{get_column_letter(width)}{1 + len(written_rows)}"
        )
        self._autofit_columns(StatSheet)
        # sixth wave: value columns G..P get a compact fixed width (autofit
        # stays for A:F — the test-name column benefits)
        for index in range(
            column_index_from_string(STAT_VALUE_WIDTH_RANGE[0]),
            column_index_from_string(STAT_VALUE_WIDTH_RANGE[1]) + 1,
        ):
            StatSheet.column_dimensions[
                get_column_letter(index)
            ].width = STAT_VALUE_COLUMN_WIDTH

    # ----------------------------------------------------------------------
    # DUT Summary matrix + chart sheets (Task 8)
    # ----------------------------------------------------------------------

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

        # use a list to determine the dut order (classic order-building)
        order = []
        for fid in sorted(summaryTable["dutInfo"].keys()):
            for dutIndex in sorted(summaryTable["dutInfo"][fid].keys()):
                order.append((fid, dutIndex))

        # full part ID per ordered DUT (classic dataRow[1] indexing); a
        # blank part ID falls back to "DUT #n"
        dut_ids = []
        for i, (fid, dutIndex) in enumerate(order):
            info = summaryTable["dutInfo"][fid][dutIndex]
            part_id = info[1] if len(info) > 1 and info[1] else ""
            dut_ids.append(part_id if part_id else f"DUT #{i + 1}")

        specs = build_dut_summary_spec(
            summaryTable,
            self.testTuples,
            order,
            dut_ids,
            sheet_name=self.tr("DUT Summary"),
            # numeric cell formats are hardcoded scientific in the spec
            # builder (2026-11-10) — no settings-derived float format anymore
            tr=self.tr,
        )

        # chunk 0 renders on the existing DUT Summary sheet; chunks 1+
        # (die overflow beyond the Excel column cap) get their own sheets —
        # spec names already carry the " (n)" suffixes. Column widths are
        # pinned inside render_matrix_sheet (A=8, B=30, C..L=8, dies=8).
        for chunk_index, spec in enumerate(specs):
            ws = DutSheet if chunk_index == 0 else self._wb.create_sheet(spec["name"])
            render_matrix_sheet(ws, spec, self._theme)

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

        # ninth wave: row 1 carries the merged A:F "Title"-style banner
        # (Product ID / Sublot ID from File Info, or the sheet-name
        # fallback); tenth wave: the banner wraps in a 50 pt row so a
        # 2-line title fits, the per-block title row is GONE and the KPI
        # cards sit one blank row below the banner
        write_merged_headline(BinSheet, 1, self._bin_banner, "Title")
        BinSheet.row_dimensions[1].height = BIN_BANNER_ROW_HEIGHT_PT
        BinSheet.cell(row=1, column=1).alignment = Alignment(
            wrap_text=True, vertical="center"
        )
        key = ReportSelection.Bin
        row = 2
        # hidden helper blocks (fixed far column) advance per chart so the
        # series references of earlier blocks stay intact
        helper_row = 1
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
            # outline) after ONE blank separator row — the head/site
            # identification lives in the chart and table titles only
            row = write_kpi_cards(BinSheet, row + 1, kpi, n_files, self._theme)

            # "Bin Charts" headline (tenth wave rename) with a blank row
            # on BOTH sides, then the charts band: ONE 250 pt row with
            # both donuts square-sized to the line (HB from column A, SB
            # offset past the HB chart's right edge); empty bin lists
            # leave the band reserved, keeping the dashboard above intact
            row = write_merged_headline(
                BinSheet, row + 1, BIN_CHARTS_HEADLINE, "Headline 1"
            )
            row += 1  # blank between the headline and the band
            chart_row = row
            BinSheet.row_dimensions[chart_row].height = BIN_CHART_BAND_HEIGHT_PT
            for (kind, _prefix, offset_cm), bins in zip(
                (
                    ("Hardware", "HB", 0.0),
                    ("Software", "SB", BIN_CHART_SIZE_CM + BIN_CHART_GAP_CM),
                ),
                aggregated,
            ):
                chart_bins = [
                    (bin_num, bin_label(bin_num, name), count)
                    for bin_num, name, count in bins
                ]
                build_bin_donut(
                    BinSheet,
                    f"{kind} Bin — Head{head} / {site_label}",
                    chart_bins,
                    band_anchor(chart_row, offset_cm),
                    BIN_HELPER_COL,
                    helper_row=helper_row,
                )
                helper_row += 2 + len(chart_bins)  # header + data rows + gap

            # one blank separator after the band, then the "Detailed Bin
            # Contribution" headline (ninth wave) with a blank row below it
            # (tenth wave) — the HB and SW tables stack below with a
            # single blank row between them
            row = write_merged_headline(
                BinSheet, chart_row + 2, BIN_CONTRIBUTION_HEADLINE, "Headline 1"
            )
            row += 1  # blank below the contribution headline
            for index, (isHBIN, bins) in enumerate(zip((True, False), aggregated)):
                kind = "Hardware" if isHBIN else "Software"
                prefix = "HB" if isHBIN else "SB"
                row = write_bin_table(
                    BinSheet,
                    row,
                    bins,
                    self._theme,
                    f"{kind} Bins — Head{head} / {site_label}",
                    f"{prefix}_Head{head}_{site_tag}",
                )
                if index == 0:
                    row += 1  # single blank separator before the SB section
            # block end: the cursor sits on the blank row after the SB
            # Total; the next block's KPI cards start after it
        self._row[key] = row

        self.loopCount += 1
        self.sendProgress()
        # gridlines off (every other sheet's look; the chart sheets bypass
        # render_plain_table, which is where the plain sheets get it)
        BinSheet.sheet_view.showGridLines = False
        self._autofit_columns(BinSheet)
        # eighth wave (R2), tenth-wave width: the Bin sheet is a fixed
        # A:F page — all six columns share width 15 (KPI grid + tables;
        # the 250 pt band charts overflow the grid on absolute anchors);
        # the autofit keeps sizing only the hidden helper columns
        first_col, last_col = (
            column_index_from_string(letter) for letter in BIN_PAGE_WIDTH_RANGE
        )
        for index in range(first_col, last_col + 1):
            BinSheet.column_dimensions[
                get_column_letter(index)
            ].width = BIN_PAGE_COLUMN_WIDTH
        # eighth wave (R1): the Bin sheet prints on A4, fit to one page
        # wide with unlimited height (same fitToPage pattern as the DUT
        # Summary matrix setup in writer.py); other sheets are untouched
        BinSheet.page_setup.paperSize = BinSheet.PAPERSIZE_A4
        BinSheet.page_setup.fitToWidth = 1
        BinSheet.page_setup.fitToHeight = 0
        BinSheet.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)

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

    # ----------------------------------------------------------------------
    # report assembly
    # ----------------------------------------------------------------------

    def report_core(self):
        self._wb = Workbook()
        set_workbook_default_font(self._wb)
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
        defaultSheet = self._wb.active

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
                    self.sheetDict[cont] = self._wb.create_sheet(sheetName)

            # openpyxl starts with a default sheet; drop it once the real
            # sheets exist (keep it when nothing is selected so the save
            # still yields a valid workbook)
            if self.sheetDict and defaultSheet in self._wb.worksheets:
                self._wb.remove(defaultSheet)

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
            self._wb.save(self.reportPath)
            self._wb.close()

        self.closeSignal.emit(True)
