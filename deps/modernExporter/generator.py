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

"""Modern openpyxl report generator: skeleton with the plain styled sheets.

``ModernReportGenerator`` subclasses the classic ``reportGenerator`` from
``deps.uic_stdExporter`` and keeps its threading contract verbatim — the
inherited ``generate_report`` slot, the ``waitFor*`` data handshake, the
``forceQuit`` early-exits and the ``loopCount`` / ``sendProgress`` cadence —
while writing the workbook with openpyxl through the modernExporter theme and
renderers instead of xlsxwriter.

Instance state (``self._wb``, ``self._theme``, ``self._row``) replaces the
classic module-level ``sv`` statics, so several generators can coexist.
Sheet writers replaced wholesale by Task 8 (DUT Summary matrix, bin, wafer,
trend, histogram) are stubbed here but keep the classic progress contract
where the classic totalLoopCnt formula counts them.
"""

from itertools import product

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
from deps.modernExporter.matrix import sanitize_cell
from deps.modernExporter.sheets import plain_table_spec
from deps.modernExporter.theme import ModernTheme
from deps.modernExporter.writer import render_plain_table

# auto-fit column width = max content length * 1.1, capped (classic uses the
# same *1.1 factor via xlsxwriter set_column; the cap is the modern bound)
AUTOFIT_FACTOR = 1.1
AUTOFIT_CAP = 60


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
        ``set_column(col, col, width*1.1)`` calls)."""
        widths: dict[int, int] = {}
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is not None:
                    length = len(str(cell.value))
                    if length > widths.get(cell.column, 0):
                        widths[cell.column] = length
        for column, length in widths.items():
            ws.column_dimensions[get_column_letter(column)].width = (
                min(length * AUTOFIT_FACTOR, AUTOFIT_CAP))

    def _write_plain_body_row(self, ws, row: int, offset: int,
                              values: list) -> None:
        """Write one body row of a plain table at 1-based ``row`` with the
        same styling as writer.render_plain_table (body font/alignment,
        thin border, alternating banding on even offsets)."""
        band = self._theme.group_fill if offset % 2 == 0 else None
        for column, value in enumerate(values, start=1):
            cell = ws.cell(row=row, column=column)
            cell.value = (value if isinstance(value, str)
                          and value.startswith("=") else sanitize_cell(value))
            cell.font = self._theme.body_font
            cell.alignment = self._theme.body_alignment
            cell.border = self._theme.border
            if band is not None:
                cell.fill = band

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
    # Task 8 stubs — replaced wholesale when the matrix/chart sheets land
    # ----------------------------------------------------------------------

    def writeDUT_Summary(self):
        """Stub: the DUT Summary matrix sheet is implemented in Task 8.
        Keeps the classic progress contract — the dut info part plus one
        column per test tuple — so loopCount still totals totalLoopCnt
        whenever DUT is selected."""
        DutSheet = self.sheetDict.get(ReportSelection.DUT, None)
        if DutSheet is None:
            return

        self.loopCount += 1 + len(self.testTuples)
        self.sendProgress()

    def writeBinChart(self):
        """Stub: the bin chart sheet is implemented in Task 8 (no progress
        bookkeeping yet — totalLoopCnt only counts it once that lands)."""
        if self.sheetDict.get(ReportSelection.Bin, None) is None:
            return

    def writeWaferMap(self):
        """Stub: the wafer map sheet is implemented in Task 8 (no progress
        bookkeeping yet — totalLoopCnt only counts it once that lands)."""
        if self.sheetDict.get(ReportSelection.Wafer, None) is None:
            return

    def writeTrendPlot(self, testTuple: tuple, head: int, site: int):
        """Stub: the trend sheet is implemented in Task 8 (no progress
        bookkeeping yet — totalLoopCnt only counts it once that lands)."""
        if self.sheetDict.get(ReportSelection.Trend, None) is None:
            return

    def writeHistogram(self, testTuple: tuple, head: int, site: int):
        """Stub: the histogram sheet is implemented in Task 8 (no progress
        bookkeeping yet — totalLoopCnt only counts it once that lands)."""
        if self.sheetDict.get(ReportSelection.Histo, None) is None:
            return

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
