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
import xlsxwriter.utility as xlutil
from PyQt5 import QtCore

from deps.SharedSrc import tab
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
from deps.exporter.worksheets.dut_summary import DUTSummaryMixin, order_dut_block
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
    FileInfoMixin, GdrDtrMixin, DUTSummaryMixin, TestStatisticsMixin, reportGenerator
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
