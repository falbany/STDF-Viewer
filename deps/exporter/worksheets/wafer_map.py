"""
#
# wafer_map.py - exporter worksheets - STDF Viewer
#
# Author: falbany (fork)
# Created Date: November 12th 2026
# -----
# Last Modified: Thu Nov 13 2026
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
Wafer Map worksheet writer: ONE WORKSHEET PER WAFER (2026-11-13 user
ruling) — each wafer gets its own styled sheet ("Wafer " + wafer ID,
sanitized/unique), replacing the pre-created single "Wafer Map" sheet.
No PNG anywhere (stacked wafer skipped); every sheet renders the live
cell-grid bin map with its own freeze panes, pinned columns and
cosmetics.

2026-11-13 wafer merge (issue #4): with "Hide re-measured DUTs (keep
latest)" checked, a real WAFER_ID spanning 2+ wafer tuples gets ONE
extra "(Merged)" sheet right after the last tuple of its group — the
per-file payloads fused die-by-die (latest file wins), re-measured dies
thick-orange-bordered, counts recomputed from the merged grid.
"""
from deps.SharedSrc import getSetting, tab
from deps.uic_stdExporter import ReportSelection
from deps.exporter.charts import (
    WAFER_LABEL_COL_WIDTH,
    WAFER_SPACER_PINNED_WIDTH,
    WAFER_VALUE_COL_WIDTH,
    wafer_sheet_name,
    wafer_table_name,
    write_wafer_map_block,
)
from deps.exporter.theme import TAB_COLOR_HEXES
from deps.exporter.wafer_merge import merge_wafer_payloads, merged_wafer_groups
from deps.exporter.writer import FOOTER_TEXT


class WaferMapMixin:
    """Wafer Map worksheet writer (one worksheet per wafer)."""

    def _style_wafer_sheet(self, WaferSheet):
        """The per-wafer cosmetics (2026-11-13 column plan — pinned AFTER
        the autofit so these win over the ledger (only A:D ever get
        notes, from the SW table): A=30 (keys/labels), B:D=8 (values),
        E=3.5 (spacer)) + gridlines off, footer, tab color. Shared by
        the per-file sheets and the merged sheet."""
        self._autofit_columns(WaferSheet)
        WaferSheet.set_column(0, 0, WAFER_LABEL_COL_WIDTH)
        WaferSheet.set_column(1, 3, WAFER_VALUE_COL_WIDTH)
        WaferSheet.set_column(4, 4, WAFER_SPACER_PINNED_WIDTH)
        WaferSheet.hide_gridlines(2)
        WaferSheet.set_footer(FOOTER_TEXT)
        WaferSheet.set_tab_color(f"#{TAB_COLOR_HEXES['wafer']}")

    def writeWaferMap(self):
        key = ReportSelection.Wafer
        base_sheet = self.sheetDict.get(key, None)
        # the pre-created single sheet is replaced by per-wafer sheets:
        # xlsxwriter 3.3.1 cannot rename a worksheet, so the (always
        # still-empty) base sheet is dropped from the workbook list and
        # one sheet per wafer is added — they land at the workbook end,
        # which is exactly the "Wafer Map" tab position of the classic
        # creation order
        if self._wb is not None and base_sheet is not None:
            self._wb.worksheets_objs = [
                ws for ws in self._wb.worksheets_objs if ws is not base_sheet
            ]
        taken = {ws.get_name() for ws in self._wb.worksheets_objs}
        row = self._row.get(key, 0)
        # merged groups only exist when the filter is ON (2026-11-13
        # wafer merge): real IDs spanning 2+ tuples; NULL/empty excluded
        merge_ids = (
            set(merged_wafer_groups(self.waferTuples))
            if getSetting().gen.hide_remeasured
            else set()
        )
        fetched = []  # (tuple index, wafer tuple, payload) in tuple order
        # pass 1 — fetch payloads in tuple order (no ticks here); the
        # stacked tuple keeps its tick exactly as today
        for index, wafer in enumerate(self.waferTuples):
            if self.forceQuit:
                return
            if wafer[0] == -1:
                # stacked wafer (fail-count colormap): no SW-bin equivalent,
                # and the report carries NO PNG at all (user ruling) — skip
                # the sheet, keep the progress tick
                self.loopCount += 1
                self.sendProgress()
                continue
            # flagged fetch: one bundled payload (info rows + map data
            # + palette) per wafer; the block fills the wafer's sheet
            payload = self.waitForDataList(
                [],
                self.selectedSites,
                self.selectedFiles,
                tab.Wafer,
                {"testTuples": [wafer], "waferMapData": wafer},
            )
            fetched.append((index, wafer, payload))
        # pass 2 — write per-file sheets in the SAME order as today; after
        # the LAST tuple of a merge group, write its merged sheet
        group_last = {
            wafer[2]: index
            for index, wafer, _ in fetched  # later tuples overwrite: last wins
        }
        for index, wafer, payload in fetched:
            name = wafer_sheet_name(index, wafer[2], taken)
            taken.add(name)
            WaferSheet = self._wb.add_worksheet(name)
            start_row = 1  # block rows are 1-based; the sheet starts A1
            write_wafer_map_block(
                WaferSheet,
                self._theme,
                payload["rows"],
                payload["map"],
                start_row,
                wafer_table_name(index, wafer[2]),
                freeze_panes=True,  # every sheet freezes its own grid
            )
            self._style_wafer_sheet(WaferSheet)
            row = start_row  # kept for _row bookkeeping symmetry
            self.loopCount += 1
            self.sendProgress()
            if wafer[2] in merge_ids and group_last[wafer[2]] == index:
                self._write_merged_wafer_sheet(
                    wafer[2], index, taken,
                    [(w, p) for _i, w, p in fetched if w[2] == wafer[2]],
                )
                row = start_row
                self.loopCount += 1
                self.sendProgress()
        self._row[key] = row

    def _write_merged_wafer_sheet(self, wafer_id, index, taken, group):
        """One merged sheet for a re-measured wafer ID (issue #4):
        `group` is that ID's [(wafer, payload)] in tuple order, ordered
        ascending by the shared supersede ranks — the LAST payload wins
        each die coordinate and owns the identity rows; re-measured dies
        get the thick orange border (merged sheets only)."""
        ranks = self._fid_ranks()
        ordered = sorted(group, key=lambda wp: ranks.get(wp[0][1], wp[0][1]))
        merged_map, superseded = merge_wafer_payloads([p for _, p in ordered])
        name = wafer_sheet_name(index, wafer_id, taken, suffix=" (Merged)")
        taken.add(name)
        WaferSheet = self._wb.add_worksheet(name)
        start_row = 1
        write_wafer_map_block(
            WaferSheet,
            self._theme,
            ordered[-1][1]["rows"],  # identity rows from the winner
            merged_map,
            start_row,
            wafer_table_name(index, wafer_id, suffix="_M"),
            freeze_panes=True,
            superseded_coords=superseded,
            source_files=len(ordered),
        )
        self._style_wafer_sheet(WaferSheet)
