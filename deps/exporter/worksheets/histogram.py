"""
#
# histogram.py - exporter worksheets - STDF Viewer
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
Histogram worksheet writer (moved verbatim from generator.py,
worksheets split 2026-11-12).
"""

from deps.SharedSrc import tab
from deps.uic_stdExporter import ReportSelection
from deps.exporter.sheets import insert_report_image
from deps.exporter.theme import TAB_COLOR_HEXES
from deps.exporter.writer import FOOTER_TEXT


class HistogramMixin:
    """Histogram worksheet writer."""

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
