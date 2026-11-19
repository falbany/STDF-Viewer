"""
#
# test_correlation.py - exporter worksheets - STDF Viewer
#
# Author: falbany (fork)
# Created Date: November 19th 2026
# -----
# Last Modified: Thu Nov 19 2026
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

Test Correlation Map sheet (2026-11-19 user ruling): one N x N
co-failure matrix — cell = DUTs failing BOTH the row and the column
test, diagonal = each test's own failing DUTs — computed from the
DUT-block context's flag mirrors, one map per temperature partition
when split-temp mode resolves >= 2 temperatures, one all-DUTs map
otherwise.

Computed, not formula'd: ALL tests ride the map (non-failing tests
included), and live SUMPRODUCT formulas die at that scale (N=200 ->
20k formulas x 8184-cell mirror arrays, seconds of Excel recalc plus a
multi-MB calc chain). The whole matrix is ONE numpy matmul over the
0/1 fail vectors the mirrors already carry ("" = not executed enters
both operands as 0 — exactly what a SUMPRODUCT would say), written as
plain cached values under ONE 3_color_scale: the wafer fail grid's
rendering path, same green/yellow/red family and the same min /
percentile-50 / max anchor computation.

Ordering: failing DUTs, descending (top-failing test top-left),
first-appearance order breaks ties; never-executed tests (no mirror
entry in the partition) stay off the map; the displayed matrix is
capped at 100 x 100 tests (2026-11-19 user ruling — programs run 30 to
200 tests with a mean of 100, so the cap bites only on the largest).
"""

from __future__ import annotations

import numpy as np

from deps.exporter.charts import WAFER_DIE_WIDTH
from deps.exporter.contract import ReportHost
from deps.exporter.helpers import format_temperature_label, get_column_letter
from deps.exporter.worksheets.test_summary import group_order_by_temperature
from deps.SharedSrc import getSetting
from deps.uic_stdExporter import ReportSelection

CORRELATION_CAP = 100  # displayed matrix cap, rows AND columns (2026-11-19 user ruling)

# the wafer fail grid's gradient family, verbatim (ONE scale per sheet,
# explicit num anchors — the 2026-11-14 legend ruling's shape)
CORRELATION_GRADIENT_MIN_HEX = "#00CC00"
CORRELATION_GRADIENT_MID_HEX = "#FFFF00"
CORRELATION_GRADIENT_MAX_HEX = "#CC0000"

# the wafer map's grid geometry: the Y-label column sits in F, the grid
# runs from G (the sbin-family anchor layout the ruling references) at
# the wafer die width — the column labels read top-to-bottom in 90°-
# rotated header cells (2026-11-19 user ruling), so the narrow width
# only has to fit the counts
CORRELATION_LABEL_COL0 = 5
CORRELATION_FIRST_GRID_COL0 = 6
CORRELATION_LABEL_WIDTH = 34
CORRELATION_GRID_COL_WIDTH = WAFER_DIE_WIDTH
_HEADER_ROW_HEIGHT_PER_CHAR = 5.5  # pt of rotated 11pt text per label char
_HEADER_ROW_HEIGHT_MAX = 409  # Excel's hard row-height cap

_HEADER_LABEL = "Test"  # the F-column corner cell (the Y-axis header slot)


def _percentile50(values: list[int]) -> float:
    """PERCENTILE.INC-50 of an ascending list (n >= 1) — the wafer fail
    grid's mid-anchor computation, shared shape."""
    n = len(values)
    idx = 0.5 * (n - 1)
    lo = int(idx)
    return values[lo] + (values[min(lo + 1, n - 1)] - values[lo]) * (idx - lo)


def build_correlation_spec(sheet_name: str, tests: list, cap: int = CORRELATION_CAP) -> dict | None:
    """The sheet spec for ONE correlation map, or ``None`` when nothing
    maps.

    ``tests`` carries ``(test_num, test_name, mirror_list)`` per test,
    the mirror lists aligned to the partition's DUT order (values are
    the context mirrors: ``""`` not executed / ``0`` pass / ``1`` fail).

    Kept tests: at least one executed DUT (a never-executed test has no
    information); non-failing tests ride at the bottom (2026-11-19
    ruling). Order: failing DUTs descending, first-appearance order
    breaks ties — the top-failing test lands top-left. Capped at ``cap``
    rows/columns; the legend says so when the cap bites.
    """
    rows = []
    for index, (num, name, mirrors) in enumerate(tests):
        if not any(m != "" for m in mirrors):
            continue  # never executed in this partition
        fails = sum(1 for m in mirrors if m == 1)
        vector = [1 if m == 1 else 0 for m in mirrors]
        rows.append((num, name, fails, vector, index))
    rows.sort(key=lambda row: (-row[2], row[4]))
    kept = len(rows)
    rows = rows[:cap]
    if not rows:
        return None
    vec = np.array([row[3] for row in rows], dtype=np.int64)
    matrix = (vec @ vec.T).tolist()
    values = sorted(value for row in matrix for value in row)
    legend = (
        "Cell = DUTs failing both the row and the column test; diagonal = "
        "the test's own failing DUTs; ordered by failing DUTs, descending"
    )
    if kept > cap:
        legend += f"; showing the first {cap} of {kept} tests"
    labels = [f"{row[0]} {row[1]}".strip() for row in rows]
    # the 90°-rotated header cells read top-to-bottom: the header row
    # must be tall enough for the longest label (Excel caps at 409 pt)
    header_row_height = min(
        _HEADER_ROW_HEIGHT_MAX,
        int(_HEADER_ROW_HEIGHT_PER_CHAR * max(len(label) for label in labels)) + 8,
    )
    return {
        "sheet_name": sheet_name,
        "legend": legend,
        "nums": [str(row[0]) for row in rows],
        "labels": labels,
        "header_row_height": header_row_height,
        "fails": [row[2] for row in rows],
        "matrix": matrix,
        "capped": kept > cap,
        "kept": kept,
        # all-pass maps force the max anchor to 1 so the all-zero grid
        # renders at the green end instead of collapsing onto mid yellow
        "gradient": {
            "type": "3_color_scale",
            "min_type": "num",
            "min_value": values[0],
            "mid_type": "num",
            "mid_value": _percentile50(values),
            "max_type": "num",
            "max_value": max(values[-1], 1),
            "min_color": CORRELATION_GRADIENT_MIN_HEX,
            "mid_color": CORRELATION_GRADIENT_MID_HEX,
            "max_color": CORRELATION_GRADIENT_MAX_HEX,
        },
    }


def render_correlation_sheet(ws, spec: dict, theme) -> None:
    """The wafer fail grid's layout carrying pair counts: title band row
    1, reading legend row 2, the column-test labels rotated 90° on the
    anchor row (2026-11-19 ruling — the wafer-die-width columns only
    have to fit the counts), F = row-test labels ("num name"), the
    counts as plain cached values from G, ONE 3_color_scale over the
    whole grid, panes frozen at the grid's top-left corner (headers stay
    visible on a 100 x 100 map)."""
    ws.hide_gridlines(2)
    labels, matrix = spec["labels"], spec["matrix"]
    n = len(labels)
    last_col0 = CORRELATION_FIRST_GRID_COL0 + n - 1
    ws.merge_range(0, 0, 0, last_col0, spec["sheet_name"], theme.lookalike("title"))
    ws.merge_range(1, 0, 1, last_col0, spec["legend"], theme.lookalike("explanatory"))

    ws.set_column(CORRELATION_LABEL_COL0, CORRELATION_LABEL_COL0, CORRELATION_LABEL_WIDTH)
    ws.set_column(CORRELATION_FIRST_GRID_COL0, last_col0, CORRELATION_GRID_COL_WIDTH)

    header_row0 = 2
    axis_fmt = theme.cell_format(font="bold_body", align="center", border=True, fill_hex="EDEFF2")
    axis_fmt_rotated = theme.cell_format(
        font="bold_body",
        align="center",
        border=True,
        fill_hex="EDEFF2",
        extra={"rotation": 90, "valign": "bottom"},
    )
    label_fmt = theme.cell_format(font="bold_body", align="left", border=True, fill_hex="EDEFF2")
    body_fmt = theme.cell_format(font="body", align="center")
    ws.write(header_row0, CORRELATION_LABEL_COL0, _HEADER_LABEL, axis_fmt)
    ws.set_row(header_row0, spec["header_row_height"])
    for offset, label in enumerate(labels):
        ws.write(header_row0, CORRELATION_FIRST_GRID_COL0 + offset, label, axis_fmt_rotated)
    for row_offset, label in enumerate(labels):
        row0 = header_row0 + 1 + row_offset
        ws.write(row0, CORRELATION_LABEL_COL0, label, label_fmt)
        for col_offset, value in enumerate(matrix[row_offset]):
            ws.write(row0, CORRELATION_FIRST_GRID_COL0 + col_offset, int(value), body_fmt)

    first_letter = get_column_letter(CORRELATION_FIRST_GRID_COL0)
    last_letter = get_column_letter(last_col0)
    ws.conditional_format(
        f"{first_letter}{header_row0 + 2}:{last_letter}{header_row0 + 1 + n}",
        dict(spec["gradient"]),
    )
    ws.freeze_panes(header_row0 + 1, CORRELATION_FIRST_GRID_COL0)


class TestCorrelationMixin(ReportHost):
    """The Test Correlation Map sheets (2026-11-19 user ruling):
    self-gated on the Test Correlation selection AND a prepared DUT-block
    context (the flag mirrors + temperatures are the map's only inputs).
    One map per temperature partition when split-temp mode resolves >= 2
    temperatures (the split sheets' own gate), one all-DUTs map
    otherwise; each sheet joins `_dut_sheet_block` so the close-time
    reorder keeps it inside the DUT-derived tab family."""

    def writeTestCorrelationMap(self) -> None:
        """Write the correlation map sheet(s); no selection, no context
        or no executed tests -> no sheet, no fetch. The sheets are
        tickless (2026-11-19 fix — they add sheets, not fetches, like
        the temperature splits; the context tick they ride is already in
        the GUI denominator via dut_block_loop_count's
        correlation_selected flag)."""
        # gate 1 — the selection (the sheet is NOT pre-created in
        # report_core's fixed loop; it must never appear unselected)
        if ReportSelection.TestCorrelation not in self.selectedContents:
            return
        # gate 2 — the DUT-block context (never ran -> no mirrors, no map)
        if self._remeasure_drops is None:
            return
        wb = self._wb
        theme = self._theme
        assert wb is not None and theme is not None
        order = self._dut_ctx_order
        base_name = self.tr("Test Correlation Map")

        def collect(pairs):
            """(test_num, test_name, mirror list) per test tuple, the
            mirrors sliced through the context pos into the partition's
            DUT order (a pair the table misses reads as not executed —
            the builders' degraded-pair blank semantics)."""
            collected = []
            for test_tuple in self.testTuples:
                entry = self._dut_ctx_stats.get(test_tuple)
                if entry is None:
                    continue
                mirrors = []
                for pair in pairs:
                    index = self._dut_ctx_pos.get(pair)
                    mirrors.append("" if index is None else entry[1][index])
                collected.append((test_tuple[0], test_tuple[2], mirrors))
            return collected

        maps = []
        if getSetting().gen.split_temperature:
            # the split sheets' own partition + >= 2 resolved gate; DUTs
            # with no resolvable temperature stay unmapped (they keep
            # the base sheet there — here they simply have no map)
            groups = group_order_by_temperature(
                order, self._dut_ctx_summary["dutInfo"], self._dut_ctx_temps
            )
            resolved = sorted(temp for temp in groups if temp is not None)
            if len(resolved) >= 2:
                maps = [
                    (
                        f"{base_name} - T{format_temperature_label(temp)}°C",
                        collect(groups[temp]),
                    )
                    for temp in resolved
                ]
        if not maps:
            maps = [(base_name, collect(order))]

        for name, tests in maps:
            spec = build_correlation_spec(name, tests)
            if spec is None:
                continue
            ws = wb.add_worksheet(name)
            render_correlation_sheet(ws, spec, theme)
            self._dut_sheet_block.append(ws)
