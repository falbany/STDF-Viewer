"""
#
# wafer_map.py - exporter worksheets - STDF Viewer
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
Wafer Map worksheet writer (moved verbatim from generator.py,
worksheets split 2026-11-12; issue #2 rewires the per-wafer body from the
static PNG to the live cell-grid bin map — stacked wafer keeps the image).
"""
from deps.SharedSrc import tab
from deps.uic_stdExporter import ReportSelection
from deps.exporter.charts import wafer_table_name, write_wafer_map_block
from deps.exporter.sheets import insert_report_image
from deps.exporter.theme import TAB_COLOR_HEXES
from deps.exporter.writer import FOOTER_TEXT


class WaferMapMixin:
    """Wafer Map worksheet writer (moved verbatim from generator.py)."""

    def writeWaferMap(self):
        WaferSheet = self.sheetDict.get(ReportSelection.Wafer, None)
        if WaferSheet is None:
            return

        key = ReportSelection.Wafer
        row = self._row.get(key, 0)
        for index, wafer in enumerate(self.waferTuples):
            if self.forceQuit:
                return
            if wafer[0] == -1:
                # stacked wafer (fail-count colormap): image + classic info
                # rows — out of the cell-map scope (issue #2 non-goal)
                image_io = self.waitForImage(
                    wafer, 0, self.selectedSites, self.selectedFiles, tab.Wafer
                )
                row += insert_report_image(WaferSheet, image_io, row) + 2
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
                row += 3
            else:
                # flagged fetch: one bundled payload (info rows + map data
                # + palette) per wafer; the block writes title, info rows,
                # SW table A:D and the die grid from column F
                payload = self.waitForDataList(
                    [],
                    self.selectedSites,
                    self.selectedFiles,
                    tab.Wafer,
                    {"testTuples": [wafer], "waferMapData": wafer},
                )
                start_row = max(row, 1)  # block rows are 1-based
                row = (
                    write_wafer_map_block(
                        WaferSheet,
                        self._theme,
                        payload["rows"],
                        payload["map"],
                        start_row,
                        wafer_table_name(index, wafer[2]),
                    )
                    + 3
                )
            self.loopCount += 1
            self.sendProgress()
        self._row[key] = row
        self._autofit_columns(WaferSheet)
        WaferSheet.hide_gridlines(2)
        WaferSheet.set_footer(FOOTER_TEXT)
        WaferSheet.set_tab_color(f"#{TAB_COLOR_HEXES['wafer']}")
