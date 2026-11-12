"""
#
# bin_chart.py - exporter worksheets - STDF Viewer
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
Bin Chart worksheet writer: banner, KPI cards, donuts and contribution
tables.

(Moved verbatim from generator.py, worksheets split 2026-11-12.)
"""
from itertools import product

import xlsxwriter.utility as xlutil

from deps.SharedSrc import tab
from deps.uic_stdExporter import ReportSelection
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
    write_bin_table,
    write_kpi_cards,
    write_merged_headline,
)
from deps.exporter.theme import TAB_COLOR_HEXES
from deps.exporter.writer import FOOTER_TEXT


class BinChartMixin:
    """Bin Chart worksheet writer (moved verbatim from generator.py)."""

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
