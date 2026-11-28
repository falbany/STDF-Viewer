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
No PNG anywhere; every sheet renders the live cell-grid bin map with its
own freeze panes, pinned columns and cosmetics.

2026-11-13 wafer merge (issue #4): with "Hide re-measured DUTs (keep
latest)" checked, a real WAFER_ID spanning 2+ wafer tuples gets ONE
extra "(Merged)" sheet right after the last tuple of its group — the
per-file payloads fused die-by-die (latest file wins), re-measured dies
thick-orange-bordered, counts recomputed from the merged grid.

2026-11-13 stacked wafer map (issue #5): the stacked tuple
(`waferIndex == -1`) rendered ONE sheet — "Wafer Stacked Map (All
Wafers)" — last among the wafer sheets (2026-11-14 user ruling: an
extended engineering view of the wafer, appended after the per-wafer
and merged sheets), the cross-wafer fail-count heatmap as a live cell
grid under one green→yellow→red color-scale. 2026-11-18 (issue #13):
that sheet moved to `worksheets/wafer_heat_map.py` — the self-gated
WaferHeatMapMixin composed right after this mixin; this writer fetches
REAL tuples only and stashes their payloads for the heat sheet's
roster.
"""

from deps.exporter.charts import (
    WAFER_LABEL_COL_WIDTH,
    WAFER_SPACER_PINNED_WIDTH,
    WAFER_VALUE_COL_WIDTH,
    aggregate_kpi,
    wafer_sheet_name,
    wafer_table_name,
    write_wafer_map_block,
)
from deps.exporter.contract import ReportHost
from deps.exporter.theme import TAB_COLOR_HEXES
from deps.exporter.wafer_merge import merge_wafer_payloads, merged_wafer_groups
from deps.exporter.writer import FOOTER_TEXT
from deps.SharedSrc import getSetting, tab
from deps.uic_stdExporter import ReportSelection


class WaferMapMixin(ReportHost):
    """Wafer Map worksheet writer (one worksheet per wafer)."""

    def _style_wafer_sheet(self, WaferSheet):
        """The per-wafer cosmetics (2026-11-13 column plan — pinned AFTER
        the autofit so these win over the ledger (only A:D ever get
        notes, from the SW table): A=30 (keys/labels), B:D=8 (values),
        E=3.5 (spacer)) + gridlines off, footer, tab color, and the
        2026-11-14 print setup (A4 landscape, one page wide / unlimited
        tall; rows 1-3 repeat on every printed page — a span coupled to
        every wafer block anchoring X-headers at row 3 via start_row=1,
        so moving a block's anchor moves this span). Shared by the
        per-file sheets, the merged sheet and the stacked sheet."""
        self._autofit_columns(WaferSheet)
        WaferSheet.set_column(0, 0, WAFER_LABEL_COL_WIDTH)
        WaferSheet.set_column(1, 3, WAFER_VALUE_COL_WIDTH)
        WaferSheet.set_column(4, 4, WAFER_SPACER_PINNED_WIDTH)
        WaferSheet.hide_gridlines(2)
        # 2026-11-14 print design (approach A): A4 landscape, one page
        # wide / unlimited tall — fit-to-width never enlarges, so small
        # maps print at 100% — and rows 1-3 repeat on every printed page
        # (title / identity-legend / X-headers: every wafer block anchors
        # them at rows 1-3 because every caller passes start_row=1 —
        # moving a block's anchor moves this span too)
        WaferSheet.set_paper(9)  # A4
        WaferSheet.set_landscape()
        WaferSheet.fit_to_pages(1, 0)
        WaferSheet.repeat_rows(0, 2)
        WaferSheet.set_footer(FOOTER_TEXT)
        WaferSheet.set_tab_color(f"#{TAB_COLOR_HEXES['wafer']}")

    def writeWaferMap(self):
        """The Wafer Map sheets: one sheet per wafer tuple, plus a merged
        sheet per re-measured ID (issue #4); payloads fetched in tuple
        order. Wafer Map content not selected (no pre-created base
        sheet) → return before any fetch or sheet write (the classic's
        guard); the fetched (index, wafer, payload) triples of the REAL
        tuples land on ``self._wafer_fetched`` — the WaferHeatMapMixin's
        roster stash (issue #13)."""
        key = ReportSelection.Wafer
        base_sheet = self.sheetDict.get(key, None)
        if base_sheet is None:
            return
        # report_core binds workbook and theme before any write* runs
        wb = self._wb
        theme = self._theme
        assert wb is not None and theme is not None
        # the pre-created single sheet is replaced by per-wafer sheets:
        # xlsxwriter 3.3.1 cannot rename a worksheet, so the (always
        # still-empty) base sheet is dropped from the workbook list and
        # one sheet per wafer is added — they land at the workbook end,
        # which is exactly the "Wafer Map" tab position of the classic
        # creation order
        wb.worksheets_objs = [ws for ws in wb.worksheets_objs if ws is not base_sheet]
        taken = {ws.get_name() for ws in wb.worksheets_objs}
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
        # stacked tuple is skipped — the WaferHeatMapMixin owns its fetch
        # and its tick now (issue #13)
        for index, wafer in enumerate(self.waferTuples):
            if self.forceQuit:
                return
            if wafer[0] == -1:
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
        # the heat map's roster stash (issue #13): real tuples only, so
        # the heat sheet's roster never re-fetches a payload
        self._wafer_fetched = fetched
        # pass 2 — per-file sheets in the SAME order as today; after the
        # LAST tuple of a merge group, write its merged sheet (the
        # stacked all-wafers sheet moved to the WaferHeatMapMixin,
        # issue #13)
        group_last = {
            wafer[2]: index
            for index, wafer, _ in fetched  # later tuples overwrite: last wins
        }
        for index, wafer, payload in fetched:
            name = wafer_sheet_name(index, wafer[2], taken)
            taken.add(name)
            WaferSheet = wb.add_worksheet(name)
            start_row = 1  # block rows are 1-based; the sheet starts A1
            write_wafer_map_block(
                WaferSheet,
                theme,
                payload["rows"],
                payload["map"],
                start_row,
                wafer_table_name(index, wafer[2]),
                freeze_panes=True,  # every sheet freezes its own grid
                wb=wb,  # issue #7: the SW table gets its palette donut
            )
            self._style_wafer_sheet(WaferSheet)
            # fork(falbany): Report Overview DUTs column — the wafer's own
            # DUT total (the same KPI this sheet displays)
            self._ov_facts.setdefault("wafer_sheet_duts", []).append(
                (name, aggregate_kpi(payload["rows"])["total"])
            )
            row = start_row  # kept for _row bookkeeping symmetry
            self.loopCount += 1
            self.sendProgress()
            if wafer[2] in merge_ids and group_last[wafer[2]] == index:
                self._write_merged_wafer_sheet(
                    wafer[2],
                    index,
                    taken,
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
        wb = self._wb
        theme = self._theme
        assert wb is not None and theme is not None  # report_core binds them
        ranks = self._fid_ranks()
        ordered = sorted(group, key=lambda wp: ranks.get(wp[0][1], wp[0][1]))
        merged_map, superseded = merge_wafer_payloads([p for _, p in ordered])
        name = wafer_sheet_name(index, wafer_id, taken, suffix=" (Merged)")
        taken.add(name)
        WaferSheet = wb.add_worksheet(name)
        start_row = 1
        write_wafer_map_block(
            WaferSheet,
            theme,
            ordered[-1][1]["rows"],  # identity rows from the winner
            merged_map,
            start_row,
            wafer_table_name(index, wafer_id, suffix="_M"),
            freeze_panes=True,
            superseded_coords=superseded,
            source_files=len(ordered),
            wb=wb,  # issue #7: the merged sheet's SW table gets its donut
        )
        self._style_wafer_sheet(WaferSheet)
        # fork(falbany): Report Overview merged-wafer facts (values in hand)
        self._ov_facts.setdefault("merged_wafers", []).append(
            (wafer_id, name, len(ordered), len(superseded))
        )
        # the merged sheet covers the WHOLE group — Σ over every file's
        # payloads, not just the winner's identity rows
        self._ov_facts.setdefault("wafer_sheet_duts", []).append(
            (name, sum(aggregate_kpi(p["rows"])["total"] for _w, p in ordered))
        )
