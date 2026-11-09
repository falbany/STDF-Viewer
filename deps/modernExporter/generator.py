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

from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from PyQt5 import QtCore

from deps.SharedSrc import getSetting, tab
from deps.uic_stdExporter import (
    ReportSelection,
    dataChannel,
    reportGenerator,
    signals,
)
from deps.modernExporter.matrix import build_dut_summary_spec
from deps.modernExporter.sheets import insert_report_image, plain_table_spec
from deps.modernExporter.theme import ModernTheme
from deps.modernExporter.writer import (
    BASE_COLUMN_COUNT,
    render_matrix_sheet,
    render_plain_table,
    write_body_cell,
)

# auto-fit column width = max content length * 1.1, capped (classic uses the
# same *1.1 factor via xlsxwriter set_column; the cap is the modern bound)
AUTOFIT_FACTOR = 1.1
AUTOFIT_CAP = 60
# DUT Summary die columns keep a pinned width — only the fixed A..L block
# is auto-fitted
DIE_COLUMN_WIDTH = 10


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

    def _autofit_columns(self, ws, max_column: int | None = None) -> None:
        """Set each used column's width to max content length * 1.1, capped
        at AUTOFIT_CAP (the modern counterpart of the classic
        ``set_column(col, col, width*1.1)`` calls). ``max_column`` limits
        the fit to the first N columns — the DUT Summary matrix auto-fits
        only its fixed A..L block, its die columns stay pinned."""
        widths: dict[int, int] = {}
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is None:
                    continue
                if max_column is not None and cell.column > max_column:
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

    def _write_plain_body_row(self, ws, row: int, offset: int,
                              values: list) -> None:
        """Write one body row of a plain table at 1-based ``row`` with the
        same styling as writer.render_plain_table (body font/alignment,
        thin border, alternating banding on even offsets)."""
        self._write_body_row(
            ws, row, values,
            fill=self._theme.group_fill if offset % 2 == 0 else None)

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

        spec = plain_table_spec(self.tr("File Info"),
                                [self.tr("Property Name"), self.tr("Value")],
                                dataTable, self.tr)
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
        # element: [""] + headers — write it verbatim, untranslated, as the
        # plain-table header row; everything after it is data
        headerRow, dataRows = dataTable[0], dataTable[1:]

        # header only (identity tr: no re-translation of the GUI header)
        spec = plain_table_spec(self.tr("Test Statistics"), headerRow, [],
                                tr=lambda s: s)
        render_plain_table(StatSheet, spec, self._theme)

        # data rows keep the classic cadence: forceQuit check, one loopCount
        # and one sendProgress per row
        for offset, dataRow in enumerate(dataRows):
            if self.forceQuit:
                return
            self._write_plain_body_row(StatSheet, 2 + offset, offset,
                                       dataRow)
            self._row[ReportSelection.Stat] = 2 + offset
            self.loopCount += 1
            self.sendProgress()

        # extend the auto-filter over the written rows (render_plain_table
        # only knew about the header when it drew the sheet)
        width = max([len(headerRow)] + [len(row) for row in dataRows])
        StatSheet.auto_filter.ref = (
            f"A1:{get_column_letter(width)}{1 + len(dataRows)}")
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
            float_format=getSetting().getFloatFormat(), tr=self.tr)

        # chunk 0 renders on the existing DUT Summary sheet; chunks 1+
        # (die overflow beyond the Excel column cap) get their own sheets —
        # spec names already carry the " (n)" suffixes
        for chunk_index, spec in enumerate(specs):
            ws = (DutSheet if chunk_index == 0
                  else self._wb.create_sheet(spec["name"]))
            render_matrix_sheet(ws, spec, self._theme)
            # auto-fit the fixed A..L block only; die columns stay pinned
            self._autofit_columns(ws, max_column=BASE_COLUMN_COUNT)
            for column in spec["die_columns"]:
                ws.column_dimensions[column].width = DIE_COLUMN_WIDTH

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
