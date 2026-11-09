"""
#
# generator.py - modernExporter package - STDF Viewer
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
# Adapted from uic_stdExporter.py by noonchen (Copyright (c) 2020 noonchen)
#
"""

"""Modern openpyxl report generator: plain, DUT Summary and chart sheets.

``ModernReportGenerator`` subclasses the classic ``reportGenerator`` from
``deps.uic_stdExporter`` and keeps its threading contract verbatim — the
inherited ``generate_report`` slot, the ``waitFor*`` data handshake, the
``forceQuit`` early-exits and the ``loopCount`` / ``sendProgress`` cadence —
while writing the workbook with openpyxl through the modernExporter theme and
renderers instead of xlsxwriter.

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
from openpyxl.utils import get_column_letter
from PyQt5 import QtCore

from deps.SharedSrc import tab
from deps.uic_stdExporter import (
    ReportSelection,
    dataChannel,
    reportGenerator,
    signals,
)
from deps.modernExporter.helpers import split_test_name
from deps.modernExporter.matrix import build_dut_summary_spec
from deps.modernExporter.sheets import insert_report_image, plain_table_spec
from deps.modernExporter.theme import ModernTheme
from deps.modernExporter.writer import (
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
FILE_INFO_NUMERIC_LABELS = frozenset({
    "Yield:", "DUTs Tested:", "DUTs Passed:", "DUTs Failed:",
    "DUTs Superseded:", "DUTs Unknown:",
})
YIELD_FORMAT = "0.00%"


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
    r"^(\d+)\s*/\s*Head(\d+)\s*/\s*(All Sites|Site(-?\d+))(?:\s*/\s*File(\d+))?$")

# SpecMin/SpecMax stylization on the Test Statistics sheet (B4), mirroring
# the DUT Summary E/F columns.
STAT_SPEC_FORMAT = "0.00E+00"
STAT_SPEC_FILL_HEX = "FFF3CD"

# English-only header anchors (whitespace-normalized matching).
_STAT_RENAMES = {"Low Limit": "SpecMin", "High Limit": "SpecMax"}
_STAT_TEST_NAME = "Test Name"


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
    site = ("All Sites" if match.group(3) == "All Sites"
            else int(match.group(4)))
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
    "spec_cols": (i, j) | None}`` where ``spec_cols`` is the 0-based
    (SpecMin, SpecMax) column pair in the TRANSFORMED header (None when
    either limit anchor was missing — translated UIs degrade to
    untransformed columns). An unexpected non-blank first cell disables
    the whole transform (header returned untouched).
    """
    cells = [str(cell) if cell is not None else "" for cell in header_row]
    if not cells or _norm_header_text(cells[0]) != "":
        return {"header": list(header_row), "enabled": False,
                "param_inserted": False, "spec_cols": None}
    body = [str(cell) for cell in cells[1:]]
    normalized = [_norm_header_text(cell) for cell in body]
    renamed = [_STAT_RENAMES.get(norm, cell)
               for norm, cell in zip(normalized, body)]
    header = list(STAT_VHEADER_LABELS) + renamed
    param_inserted = _STAT_TEST_NAME in normalized
    if param_inserted:
        header.insert(4 + normalized.index(_STAT_TEST_NAME) + 1, "Param")
    spec_cols = None
    if "SpecMin" in renamed and "SpecMax" in renamed:
        spec_cols = (header.index("SpecMin"), header.index("SpecMax"))
    return {"header": header, "enabled": True,
            "param_inserted": param_inserted, "spec_cols": spec_cols}


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
    out = (list(parsed) if parsed is not None
           else [raw_vheader, "", "", ""])
    rest = cells[1:]
    if header_info.get("param_inserted"):
        name = rest[0] if rest else ""
        prefix, leaf = split_test_name(name if isinstance(name, str)
                                       else str(name))
        out.extend([prefix, leaf])
    elif rest:
        out.append(rest[0])
    out.extend(numericize_value(value) for value in rest[1:])
    return out


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
    def __init__(self, signals: signals, mutex: QtCore.QMutex,
                 conditionWait: QtCore.QWaitCondition, reportInfo: tuple,
                 channel: dataChannel):
        super().__init__(signals, mutex, conditionWait, reportInfo, channel)
        # instance state replacing the classic module-level sv statics
        self._wb: Workbook | None = None
        self._theme = ModernTheme()
        self._row: dict[ReportSelection, int] = {}   # per-sheet row cursors

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
            ws.column_dimensions[get_column_letter(column)].width = (
                min(length * AUTOFIT_FACTOR, AUTOFIT_CAP))

    def _write_body_row(self, ws, row: int, values: list,
                        fill=None) -> None:
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
            rows, self.tr,
            cell_number_formats=cell_number_formats)
        render_plain_table(InfoSheet, spec, self._theme)
        self._autofit_columns(InfoSheet)

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
        dataTable = [[ele.strip("\n") for ele in dataRow]
                     for dataRow in dataTable]

        spec = plain_table_spec(
            self.tr("GDR & DTR Summary"),
            [self.tr("Record Type"), self.tr("Value"),
             self.tr("Approx. Location")],
            dataTable, self.tr)
        render_plain_table(GDR_DTR_Sheet, spec, self._theme)
        self._autofit_columns(GDR_DTR_Sheet)

        self.loopCount += 1
        self.sendProgress()

    def writeStatistic(self, testTuples: list, heads: list[int],
                       sites: list[int]):
        # Sheet for statistics of test items, e.g. cpk, mean, etc.
        StatSheet = self.sheetDict.get(ReportSelection.Stat, None)
        if StatSheet is None:
            return

        dataTable = self.waitForDataList(heads, sites, self.selectedFiles,
                                         tab.Trend,
                                         {"testTuples": testTuples})
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

        # header only (identity tr: no re-translation of the GUI header)
        spec = plain_table_spec(self.tr("Test Statistics"),
                                header_info["header"], [], tr=lambda s: s)
        spec["number_formats"] = number_formats
        spec["column_fills"] = column_fills
        render_plain_table(StatSheet, spec, self._theme)

        # data rows keep the classic cadence: forceQuit check, one loopCount
        # and one sendProgress per row — the transforms happen per row
        # before writing, without altering the progress semantics
        written_rows = []
        for offset, dataRow in enumerate(dataRows):
            if self.forceQuit:
                return
            stat_row = transform_stat_row(dataRow, header_info)
            written_rows.append(stat_row)
            write_plain_row(
                StatSheet, 2 + offset, stat_row, self._theme,
                band=(self._theme.group_fill if offset % 2 == 0 else None),
                number_formats=number_formats, column_fills=column_fills)
            self._row[ReportSelection.Stat] = 2 + offset
            self.loopCount += 1
            self.sendProgress()

        # extend the auto-filter over the written rows (render_plain_table
        # only knew about the header when it drew the sheet)
        width = max([len(header_info["header"])]
                    + [len(row) for row in written_rows])
        StatSheet.auto_filter.ref = (
            f"A1:{get_column_letter(width)}{1 + len(written_rows)}")
        self._autofit_columns(StatSheet)

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
        summaryTable = self.waitForDutSummary(self.selectedHeads,
                                              self.selectedSites,
                                              self.selectedFiles,
                                              self.testTuples)
        if self.forceQuit:
            return

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
            summaryTable, self.testTuples, order, dut_ids,
            sheet_name=self.tr("DUT Summary"),
            # numeric cell formats are hardcoded scientific in the spec
            # builder (2026-11-10) — no settings-derived float format anymore
            tr=self.tr)

        # chunk 0 renders on the existing DUT Summary sheet; chunks 1+
        # (die overflow beyond the Excel column cap) get their own sheets —
        # spec names already carry the " (n)" suffixes. Column widths are
        # pinned inside render_matrix_sheet (A=8, B=30, C..L=8, dies=8).
        for chunk_index, spec in enumerate(specs):
            ws = (DutSheet if chunk_index == 0
                  else self._wb.create_sheet(spec["name"]))
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

        key = ReportSelection.Bin
        row = self._row.get(key, 0)
        for head, site in product(self.selectedHeads, self.selectedSites):
            if self.forceQuit:
                return
            # get bin image from GUI thread and insert to the sheet
            image_io = self.waitForImage(testTuple=(0, 0, ""), head=head,
                                         sites=[site],
                                         fids=self.selectedFiles,
                                         chartType=tab.Bin)
            row += insert_report_image(BinSheet, image_io, row)
            # get hard bin list, then soft bin list
            for isHBIN in (True, False):
                dataTable = self.waitForDataList([head], [site],
                                                 self.selectedFiles, tab.Bin,
                                                 {"isHBIN": isHBIN})
                for dataList in dataTable:
                    self._write_body_row(BinSheet, row, dataList)
                    row += 1
            # gap between head/site blocks
            row += 3
        self._row[key] = row

        self.loopCount += 1
        self.sendProgress()
        self._autofit_columns(BinSheet)

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
            image_io = self.waitForImage(wafer, 0, self.selectedSites,
                                         self.selectedFiles, tab.Wafer)
            row += insert_report_image(WaferSheet, image_io, row) + 2
            # add wafer info
            dataTable = self.waitForDataList([], self.selectedSites,
                                             self.selectedFiles, tab.Wafer,
                                             {"testTuples": [wafer]})
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
        image_io = self.waitForImage(testTuple, head, [site],
                                     self.selectedFiles, tab.Trend)
        dataTable = self.waitForDataList([head], [site], self.selectedFiles,
                                         tab.Trend,
                                         {"testTuples": [testTuple]})
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

        image_io = self.waitForImage(testTuple, head, [site],
                                     self.selectedFiles, tab.Histo)
        dataTable = self.waitForDataList([head], [site], self.selectedFiles,
                                         tab.Histo,
                                         {"testTuples": [testTuple]})
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
        self.sheetDict = {}
        self.loopCount = 0
        self._row = {}
        defaultSheet = self._wb.active

        # the try/finally mirrors the classic `with Workbook(...)` block:
        # the (partial) workbook is saved on forceQuit early-returns and on
        # exceptions alike; only the success path emits closeSignal
        try:
            # 1. create sheets for each selected contents, in the fixed
            #    classic order; 2./3./4. below write them in classic order
            for cont in [ReportSelection.FileInfo, ReportSelection.DUT,
                         ReportSelection.GDR_DTR, ReportSelection.Stat,
                         ReportSelection.Trend, ReportSelection.Histo,
                         ReportSelection.Bin, ReportSelection.Wafer,
                         ReportSelection.PPQQ, ReportSelection.Correlation]:
                # sheet order in the xlsx is fixed
                if   cont == ReportSelection.FileInfo:  sheetName = self.tr("File Info")
                elif cont == ReportSelection.DUT:       sheetName = self.tr("DUT Summary")
                elif cont == ReportSelection.GDR_DTR:   sheetName = self.tr("GDR & DTR Summary")
                elif cont == ReportSelection.Stat:      sheetName = self.tr("Test Statistics")
                elif cont == ReportSelection.Trend:     sheetName = self.tr("Trend Chart")
                elif cont == ReportSelection.Histo:     sheetName = self.tr("Histogram")
                elif cont == ReportSelection.Bin:       sheetName = self.tr("Bin Chart")
                elif cont == ReportSelection.Wafer:     sheetName = self.tr("Wafer Map")
                elif cont == ReportSelection.PPQQ:      sheetName = self.tr("Normal Validation")
                else:                                   sheetName = self.tr("Correlation")
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
            if self.forceQuit: return
            # GDR & DTR Summary
            self.writeGDR_DTR()
            if self.forceQuit: return
            # dut summary
            self.writeDUT_Summary()
            if self.forceQuit: return

            # bin sheet
            self.writeBinChart()
            if self.forceQuit: return
            # wafer sheet
            self.writeWaferMap()
            if self.forceQuit: return

            # ** write contents related to test numbers
            self.writeStatistic(self.testTuples, self.selectedHeads,
                                self.selectedSites)
            if self.forceQuit: return

            for testTuple in self.testTuples:
                for head, site in product(self.selectedHeads,
                                          self.selectedSites):
                    if self.forceQuit: return

                    self.writeTrendPlot(testTuple, head, site)
                    self.writeHistogram(testTuple, head, site)

                # add gap between different test items
                for key in (ReportSelection.Trend, ReportSelection.Histo):
                    self._row[key] = self._row.get(key, 0) + 2
        finally:
            self._wb.save(self.reportPath)
            self._wb.close()

        self.closeSignal.emit(True)
