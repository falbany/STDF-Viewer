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

import xlsxwriter
from PyQt5 import QtCore

from deps.SharedSrc import tab
from deps.uic_stdExporter import (
    ReportSelection,
    dataChannel,
    reportGenerator,
    signals,
)
from deps.exporter.charts import complete_chart_flags
from deps.exporter.worksheets.bin_chart import BinChartMixin
from deps.exporter.worksheets.dut_summary import DUTSummaryMixin, order_dut_block
from deps.exporter.worksheets.histogram import HistogramMixin
from deps.exporter.worksheets.trend_chart import TrendChartMixin
from deps.exporter.worksheets.wafer_map import WaferMapMixin
from deps.exporter.sheets import (
    insert_report_image,
)
from deps.exporter.theme import ModernTheme, TAB_COLOR_HEXES
from deps.exporter.writer import (
    FOOTER_TEXT,
    apply_autofit,
    reset_widths,
    write_body_cell,
)
from deps.exporter.worksheets.file_info import (
    FileInfoMixin,
)
from deps.exporter.worksheets.gdr_dtr import GdrDtrMixin
from deps.exporter.worksheets.test_statistics import TestStatisticsMixin

# auto-fit column width = max content length * 1.1, capped (classic uses the
# same *1.1 factor via xlsxwriter set_column; the cap is the modern bound).
# Plain sheets only — the DUT Summary widths are pinned by the renderer.
AUTOFIT_FACTOR = 1.1
AUTOFIT_CAP = 60


class ModernReportGenerator(
    FileInfoMixin,
    GdrDtrMixin,
    DUTSummaryMixin,
    TestStatisticsMixin,
    BinChartMixin,
    WaferMapMixin,
    TrendChartMixin,
    HistogramMixin,
    reportGenerator,
):
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
