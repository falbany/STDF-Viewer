"""
#
# wafer_heat_map.py - exporter worksheets - STDF Viewer
#
# Author: falbany (fork)
# Created Date: November 18th 2026
# -----
# Last Modified: Wed Nov 18 2026
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
Stacked cross-wafer renderer, moved verbatim from ``charts.py``
(2026-11-18, issue #13: single consumer) — the stacked wafer sheet's
reading-legend constants, the per-wafer roster and the ONE-sheet stacked
block writer, plus the ``WaferHeatMapMixin`` that composes it into
``ModernReportGenerator`` as a self-gated sheet writer (the split:
``writeWaferMap`` no longer writes anything stacked). ``charts.py``
keeps the shared wafer-grid machinery.
"""

import xlsxwriter.utility as xlutil

from deps.exporter.charts import (
    WAFER_DIE_ROW_PT,
    WAFER_DIE_WIDTH,
    WAFER_FIRST_DIE_COL0,
    WAFER_SPACER_COL0,
    WAFER_STACKED_KPI_ROWS,
    WAFER_STACKED_KPI_SPANS,
    WAFER_YAXIS_COL0,
    WAFER_YAXIS_WIDTH,
    _wafer_kv_row,
    _wafer_prop_fmt,
    _wafer_site_info,
    aggregate_kpi,
    bin_table_ranges,
    build_bin_donut,
    build_wafer_grid,
    wafer_sheet_name,
    wafer_table_name,
    write_bin_table,
    write_kpi_cards,
)
from deps.exporter.contract import ReportHost
from deps.exporter.helpers import sanitize_cell
from deps.SharedSrc import tab
from deps.uic_stdExporter import ReportSelection

# issue #5 reading legend (2026-11-13 user ask; 2026-11-14 sbin-uniform
# layout: ONE line on row 2, the identity slot — the map anchor at F3
# leaves a single full-width row, and the "what it aggregates" moved into
# the headline card's subtitle): how to read the gradient plus the
# blank-vs-0 distinction (0 = probed & passed, blank = never probed).
# Sized to render inside the A:map-width merge without wrapping; the
# empty branch merges across a 5-die map's span (A:K).
# The fail-count gradient (issue #5, colors pinned verbatim): green low
# → yellow mid → red hot — the sbin palette's pass/fail hexes with
# Excel's canonical mid; shared verbatim by the die map and the Fail Map
# Count legend column.
WAFER_GRADIENT_MIN_HEX = "#00CC00"
WAFER_GRADIENT_MID_HEX = "#FFFF00"
WAFER_GRADIENT_MAX_HEX = "#CC0000"

WAFER_STACKED_LEGEND = "Green: few fails → Red: many · 0: probed, passed · blank: never probed"


def stacked_wafer_ids(payloads: list) -> "tuple[list[str], int]":
    """Per-wafer IDs for the stacked sheet's roster (2026-11-16 ruling —
    names only; the ID table's replacement): ``[wafer_id]`` per
    ``(wafer_id, payload)`` pair, in tuple order (duplicates preserved —
    a roster, not a set), plus the Σ of the per-wafer DUT totals (the
    stacked sheet's "Total DUTs Tested" KPI card). No grid walk, no
    yield — per-wafer detail lives on each wafer's own sheet."""
    ids = []
    total_duts = 0
    for wafer_id, payload in payloads:
        ids.append(wafer_id)
        total_duts += aggregate_kpi(payload.get("rows") or [])["total"]
    return ids, total_duts


WAFER_STACKED_ROSTER_LABEL = "Stacked Wafers"


def write_stacked_wafer_roster(
    ws,
    start_row: int,
    wafer_ids: list,
    theme,
) -> int:
    """The stacked sheet's wafer roster (2026-11-16 ruling — the ID
    table's replacement): a small italic-gray label line, then ONE
    merged A:D line per wafer ID in tuple order, in the per-wafer
    sheets' metadata-box look — Explanatory-Text format, light-gray
    external border only (first line: top edge, last: bottom, every
    line: left/right; no internal separators). Pinned columns — never
    note_width. Returns the blank row after the last line."""
    ws.merge_range(
        start_row - 1,
        0,
        start_row - 1,
        3,
        WAFER_STACKED_ROSTER_LABEL,
        _wafer_prop_fmt(theme, ""),
    )
    last = len(wafer_ids) - 1
    for offset, wafer_id in enumerate(wafer_ids):
        sides = ("t" if offset == 0 else "") + "lr" + ("b" if offset == last else "")
        ws.merge_range(
            start_row + offset,
            0,
            start_row + offset,
            3,
            sanitize_cell(str(wafer_id)),
            _wafer_prop_fmt(theme, sides, align="center"),
        )
    return start_row + len(wafer_ids) + 1


# Fail Map Count (2026-11-14 user ask): the fail-count distribution
# table below the roster — write_bin_table with live COUNTIF counts
# over the die map, one row per DISTINCT count value in the grid
WAFER_FAIL_TABLE_TITLE = "Fail Map Count — All Wafers"
WAFER_FAIL_TABLE_HEADERS = ("Fails per Die", "Value", "Count", "Distribution")


def write_stacked_wafer_block(  # pylint: disable=too-many-locals
    ws,
    theme,
    rows: list,
    map_data: dict,
    start_row: int,
    freeze_panes: bool = False,
    wafer_ids: "list | None" = None,
    total_duts: int = 0,
    fail_table_name: "str | None" = None,
    wb=None,
) -> int:
    """One stacked wafer block (issue #5; 2026-11-14 sbin-uniform layout):
    title band, the reading legend on row 2 (the sbin identity slot), the
    live fail-count die grid anchored at F{start + 2} EXACTLY like the
    sbin sheets (X headers on the anchor row, Y-axis header in F, dies
    from the next one — counts as the cell VALUES from G) under ONE
    3_color_scale rule (green #00CC00 low → yellow #FFFF00 → red #CC0000
    high; default min / percentile-50 / max anchors), and the info items
    as KPI cards across A:D with the sbin geometry (Total Failed DUTs
    full-width rows 4-6 =SUM over the die grid; Fail Coordinates col A +
    Files B:D side by side rows 7-9, =COUNTIF ">0" / plain count; slot 2
    rows 10-12 splits in two — DUTs per Wafer col A lavender =COUNT over
    the die grid (the sbin Total Parts analog) + Total DUTs Tested B:D
    light green (Σ of the per-wafer DUT totals) — cached results ride
    the formulas). Real aggregations carry a
    count-0 bucket (getStackedWaferData seeds every probed coordinate):
    probed-but-passing dies render as 0 at the green end of the gradient,
    wafer-outline style; only never-probed gap coordinates stay unwritten
    (CF ignores blanks). ``wafer_ids`` renders the per-wafer roster
    (write_stacked_wafer_roster — 2026-11-16 ruling, the ID table's
    replacement) at the row-14 slot (ONE blank row after the cards);
    ``fail_table_name`` renders the Fail Map Count below it (one blank
    row after the roster's last line) — write_bin_table with
    live COUNTIF counts over the die map, one row per DISTINCT count
    value in the grid ("0 fails" = probed & passed dies, Distribution =
    share of the probed coordinates, Total = Σ = probed coordinates).
    Empty grid → title + legend + the info items as plain kv lines, no
    cards, no tables (the sbin empty shape). Returns the next free row.

    2026-11-15 donut (issue #7): ``wb`` given AND the fail table written
    → ONE donut ONE blank row below its Total (anchor ``A{end+1}``, the
    A:D zone), standard cycled ``BIN_SLICE_PALETTE`` slices — fail-count
    buckets carry no user color — reading the VISIBLE table refs."""
    info = map_data.get("Info") or (1.0, "", False, False, "", [])
    wafer_id = info[4] if len(info) > 4 else ""
    sites = info[5] if len(info) > 5 else []
    site_info = _wafer_site_info(sites)
    title = f"{wafer_id} – {site_info}" if wafer_id else site_info

    x_axis, y_axis, grid, _total = build_wafer_grid(map_data)
    info_cells = [cell for row in (rows or []) for cell in row]

    if not grid:
        # empty stacked aggregation (no failed DUTs in the selection):
        # title + legend + info lines — the export must not break
        ws.merge_range(start_row - 1, 0, start_row - 1, 3, title, theme.lookalike("title"))
        # no grid → no map-width span: merge across a 5-die map's span
        # (A:K) so both branches carry the same legend
        ws.merge_range(
            start_row,
            0,
            start_row,
            WAFER_FIRST_DIE_COL0 + 4,
            WAFER_STACKED_LEGEND,
            _wafer_prop_fmt(theme, ""),
        )
        for offset, cell in enumerate(info_cells):
            key, _, value = cell.partition(":")
            _wafer_kv_row(ws, start_row + 2 + offset, key.strip() + ":", value.strip(), theme)
        return start_row + len(info_cells) + 2

    last_col0 = WAFER_FIRST_DIE_COL0 + len(x_axis) - 1
    ws.merge_range(start_row - 1, 0, start_row - 1, last_col0, title, theme.lookalike("title"))

    # reading legend on row 2, the sbin identity slot (full width, the
    # "Explanatory Text" lookalike)
    ws.merge_range(
        start_row,
        0,
        start_row,
        last_col0,
        WAFER_STACKED_LEGEND,
        _wafer_prop_fmt(theme, ""),
    )

    xh_row = start_row + 2  # 1-based X-header row — the sbin F3 anchor
    first_letter = xlutil.xl_col_to_name(WAFER_FIRST_DIE_COL0)
    last_letter = xlutil.xl_col_to_name(last_col0)

    if freeze_panes:
        # keep title + legend + X-header rows and columns A:F (the
        # Y-axis header) visible while scrolling the die grid — same
        # cosmetics as the sbin sheets (spec Goal 4)
        ws.freeze_panes(xh_row, WAFER_FIRST_DIE_COL0)

    # fixed squares: spacer E, Y-axis header F, die columns G.. — explicit
    # widths, never the ledger (same as the sbin block)
    ws.set_column(WAFER_SPACER_COL0, WAFER_SPACER_COL0, WAFER_DIE_WIDTH)
    ws.set_column(WAFER_YAXIS_COL0, WAFER_YAXIS_COL0, WAFER_YAXIS_WIDTH)
    ws.set_column(WAFER_FIRST_DIE_COL0, last_col0, WAFER_DIE_WIDTH)

    axis_fmt = theme.cell_format(font="bold_body", align="center", border=True, fill_hex="EDEFF2")
    for col_offset, x in enumerate(x_axis):
        ws.write(xh_row - 1, WAFER_FIRST_DIE_COL0 + col_offset, str(x), axis_fmt)

    die_fmt = theme.cell_format(font="body", align="center")
    for row_offset, y in enumerate(y_axis):
        row0 = xh_row + row_offset  # 0-based; first die row = 1-based xh_row+1
        ws.set_row(row0, WAFER_DIE_ROW_PT)
        ws.write(row0, WAFER_YAXIS_COL0, str(y), axis_fmt)
        for col_offset, x in enumerate(x_axis):
            count = grid.get((x, y))
            if count is not None:
                ws.write(row0, WAFER_FIRST_DIE_COL0 + col_offset, int(count), die_fmt)

    # ONE gradient definition shared by every color-scale rule on the
    # sheet (2026-11-14 legend ruling): explicit num anchors — min /
    # PERCENTILE.INC-50 / max of the die values — so the Fail Map Count
    # Value column renders EXACTLY the map's colors for equal values
    # (the legend's distinct-value population would otherwise get its
    # own percentile mid, drifting from the map's weighted mid on the
    # skewed real distributions). The die range itself renders exactly
    # as before: Excel derives the same anchors from these values.
    die_values = sorted(grid.values())
    n = len(die_values)
    idx = 0.5 * (n - 1)
    lo = int(idx)
    die_mid = die_values[lo] + (die_values[min(lo + 1, n - 1)] - die_values[lo]) * (idx - lo)
    gradient = {
        "type": "3_color_scale",
        "min_color": WAFER_GRADIENT_MIN_HEX,
        "mid_color": WAFER_GRADIENT_MID_HEX,
        "max_color": WAFER_GRADIENT_MAX_HEX,
        "min_type": "num",
        "min_value": die_values[0],
        "mid_type": "num",
        "mid_value": die_mid,
        "max_type": "num",
        "max_value": die_values[-1],
    }
    ws.conditional_format(
        f"{first_letter}{xh_row + 1}:{last_letter}{xh_row + len(y_axis)}",
        dict(gradient),
    )

    # KPI cards across A:D with the sbin geometry (WAFER_KPI_SPANS/ROWS):
    # the three info items as cards, values LIVE over the die grid —
    # SUM = total failed DUTs (count-0 dies contribute nothing),
    # COUNTIF ">0" = coordinates with at least one fail; Files has no
    # grid source — parsed from the payload rows. Cached results ride
    # the formulas (xlsxwriter-migration convention).
    map_range = f"${first_letter}${xh_row + 1}:${last_letter}${xh_row + len(y_axis)}"
    total = sum(grid.values())
    coords = sum(1 for v in grid.values() if v > 0)
    files = 0
    for cell in info_cells:
        if isinstance(cell, str) and cell.strip().startswith("Files:"):
            try:
                files = int(cell.partition(":")[2])
            except ValueError:
                files = 0
    card_defs = (
        (
            "Total Failed DUTs",
            (f"=SUM({map_range})", total),
            "fail count per die, summed across all files",
            False,
        ),
        (
            "Fail Coordinates",
            (f'=COUNTIF({map_range},">0")', coords),
            "dies failing at least once",
            False,
        ),
        ("Files", files, "stacked hot-spot view", False),
        # 2026-11-14 user ask: slot 2 (rows 10-12) splits in two — the
        # Total Parts analog (live COUNT over the die grid, lavender) +
        # the per-wafer DUT total (light green)
        (
            "DUTs per Wafer",
            (f"=COUNT({map_range})", len(grid)),
            "probed dies across all wafers",
            False,
        ),
        ("Total DUTs Tested", total_duts, "Σ of the per-wafer DUT totals", False),
    )
    write_kpi_cards(
        ws,
        start_row + 3,  # legend row 2 → cards rows 4-12 (sbin geometry)
        {},  # kpi / n_files unused: card_defs replaces the built cards
        1,
        theme,
        note_widths=False,  # pinned columns — never the ledger
        card_spans=WAFER_STACKED_KPI_SPANS,
        card_rows=WAFER_STACKED_KPI_ROWS,
        card_defs=card_defs,
    )

    next_free = max(start_row + 12, xh_row + len(y_axis) + 1)
    table_row = start_row + 13  # 1-based row 14: ONE blank row after the cards
    if wafer_ids:
        roster_end = write_stacked_wafer_roster(ws, table_row, wafer_ids, theme)
        next_free = max(next_free, roster_end)
        table_row = roster_end + 1  # Fail Map Count sits below, one blank row
    if fail_table_name:
        # one row per DISTINCT count value in the grid (ascending; the
        # count-0 bucket = probed & passed dies) — counts cached against
        # the live COUNTIF over the die map
        counts = {}
        for v in grid.values():
            counts[v] = counts.get(v, 0) + 1
        fail_bins = [(v, "1 fail" if v == 1 else f"{v} fails", counts[v]) for v in sorted(counts)]
        # table_end = the writer's returned row — the blank spacer after
        # the Total; the issue #7 donut sits ONE blank row below that
        table_end = write_bin_table(
            ws,
            table_row,
            fail_bins,
            theme,
            WAFER_FAIL_TABLE_TITLE,
            fail_table_name,
            first_col=1,
            count_map_range=map_range,
            headers=WAFER_FAIL_TABLE_HEADERS,
        )
        next_free = max(next_free, table_end)
        # the Value column IS the legend (2026-11-14 user ask): the same
        # gradient as the map, over the Value data rows only
        ws.conditional_format(
            f"B{table_row + 2}:B{table_row + 1 + len(fail_bins)}",
            dict(gradient),
        )
        if wb is not None:
            name_cell, cat_range, val_range = bin_table_ranges(
                table_row, len(fail_bins), first_col=1
            )
            build_bin_donut(
                wb,
                ws,
                WAFER_FAIL_TABLE_TITLE,
                fail_bins,
                f"A{table_end + 1}",
                name_cell,
                cat_range,
                val_range,
            )
    return next_free


class WaferHeatMapMixin(ReportHost):
    """The stacked all-wafers heat map sheet (issue #13): self-gated on
    selection AND stacked-tuple presence; appends "Wafer Heat Map" last,
    after every wafer-dedicated sheet. The roster rides writeWaferMap's
    payload stash when both are selected (zero extra round-trips) and
    fetches the missing tuples data-only (no sheets, no ticks) alone."""

    def writeWaferHeatMap(self) -> None:
        """The "Wafer Heat Map" sheet: the stacked tuple's cross-wafer
        fail-count heatmap, rendered by write_stacked_wafer_block exactly
        as writeWaferMap used to (argument-for-argument call, issue #13).
        Gated on the WaferHeatMap selection AND a stacked tuple
        (`waferIndex == -1`) actually existing — no wafer data means no
        sheet, no fetch, no tick. Called by report_core right after
        writeWaferMap; per-wafer payloads come from its ``_wafer_fetched``
        stash when it ran, else data-only fetches here."""
        # gate 1 — the selection (the sheet is NOT pre-created in
        # report_core's fixed loop; it must never appear unselected)
        if ReportSelection.WaferHeatMap not in self.selectedContents:
            return
        # gate 2 — presence (doc-review call): no stacked tuple in the
        # database → no sheet, no fetch, no tick
        stacked_index = next(
            (i for i, w in enumerate(self.waferTuples) if w[0] == -1), None
        )
        if stacked_index is None:
            return
        # report_core binds workbook and theme before any write* runs
        wb = self._wb
        theme = self._theme
        assert wb is not None and theme is not None
        # stacked fetch — no tick here: the tick is the sheet write below
        stacked = self.waferTuples[stacked_index]
        stacked_payload = self.waitForDataList(
            [],
            self.selectedSites,
            self.selectedFiles,
            tab.Wafer,
            {"testTuples": [stacked], "waferMapData": stacked},
        )
        # roster pairs — ride writeWaferMap's stash (both selected: zero
        # extra round-trips); alone, fetch the missing per-wafer payloads
        # data-only (no sheets, no ticks). A partial stash is consumed
        # defensively: only the uncovered real tuples fetch.
        stash = list(self._wafer_fetched)
        covered = {index for index, _w, _p in stash}
        fetched = []
        for index, wafer in enumerate(self.waferTuples):
            if self.forceQuit:
                return
            if wafer[0] == -1 or index in covered:
                continue
            payload = self.waitForDataList(
                [],
                self.selectedSites,
                self.selectedFiles,
                tab.Wafer,
                {"testTuples": [wafer], "waferMapData": wafer},
            )
            fetched.append((index, wafer, payload))
        self._wafer_fetched = stash + fetched
        # per-wafer roster: every payload is now in hand — no new fetch
        # round-trip for the names-only roster, and the Total DUTs Tested
        # card sums the per-wafer DUT totals
        wafer_ids, total_duts = stacked_wafer_ids(
            [(w[2], p) for _i, w, p in stash + fetched]
        )
        taken = {ws.get_name() for ws in wb.worksheets_objs}
        name = wafer_sheet_name(stacked_index, "Heat Map", taken)
        taken.add(name)
        HeatSheet = wb.add_worksheet(name)
        write_stacked_wafer_block(
            HeatSheet,
            theme,
            stacked_payload["rows"],
            stacked_payload["map"],
            1,
            freeze_panes=True,  # same grid freeze as the sbin sheets
            wafer_ids=wafer_ids,
            total_duts=total_duts,
            fail_table_name=wafer_table_name(stacked_index, "FailMapCount"),
            wb=wb,  # issue #7: the Fail Map Count table gets its donut
        )
        # per-wafer cosmetics (cross-mixin by convention, contract-declared)
        self._style_wafer_sheet(HeatSheet)
        self.loopCount += 1
        self.sendProgress()
