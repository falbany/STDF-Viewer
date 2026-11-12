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
static PNG to the live cell-grid bin map. 2026-11-12 user ruling: NO PNG
on the sheet at all — the stacked-wafer fail-count map (waferInd == -1)
has no SW-bin equivalent, so its tuple is skipped (progress still ticks);
the sheet therefore starts at A1).
"""
from deps.SharedSrc import tab
from deps.uic_stdExporter import ReportSelection
from deps.exporter.charts import wafer_table_name, write_wafer_map_block
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
        first_block = True
        for index, wafer in enumerate(self.waferTuples):
            if self.forceQuit:
                return
            if wafer[0] == -1:
                # stacked wafer (fail-count colormap): no SW-bin equivalent,
                # and the sheet carries NO PNG at all (user ruling) — skip
                # the block, keep the progress tick
                self.loopCount += 1
                self.sendProgress()
                continue
            # flagged fetch: one bundled payload (info rows + map data
            # + palette) per wafer; the block writes title, KPI cards,
            # info rows, SW table A:D and the die grid from column F
            payload = self.waitForDataList(
                [],
                self.selectedSites,
                self.selectedFiles,
                tab.Wafer,
                {"testTuples": [wafer], "waferMapData": wafer},
            )
            start_row = max(row, 1)  # block rows are 1-based; sheet starts A1
            row = (
                write_wafer_map_block(
                    WaferSheet,
                    self._theme,
                    payload["rows"],
                    payload["map"],
                    start_row,
                    wafer_table_name(index, wafer[2]),
                    freeze_panes=first_block,
                )
                + 3
            )
            first_block = False
            self.loopCount += 1
            self.sendProgress()
        self._row[key] = row
        self._autofit_columns(WaferSheet)
        WaferSheet.hide_gridlines(2)
        WaferSheet.set_footer(FOOTER_TEXT)
        WaferSheet.set_tab_color(f"#{TAB_COLOR_HEXES['wafer']}")
