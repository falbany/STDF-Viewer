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
rendering path and the same min / percentile-50 / max anchor
computation, colored with the Test Summary Cpk bands' modern soft
green / amber / red palette (2026-11-19 user ruling).

Ordering: failing DUTs, descending (top-failing test top-left),
first-appearance order breaks ties; never-executed tests (no mirror
entry in the partition) stay off the map; the displayed matrix is
capped at 100 x 100 tests (2026-11-19 user ruling — programs run 30 to
200 tests with a mean of 100, so the cap bites only on the largest).

2026-11-19 follow-up rulings: the grid carries ONE Excel Table
(style: None, so the hand-tuned theme survives) with AutoFilter, so
sorting/filtering a test column surfaces its correlations; the N x N
diagonal gets a formula-based border rule ("this test compared to
itself") that follows table sorts; the sheet prints A4 landscape, fit
to one page wide with rows 1-3 repeated; its tab is light water blue
at 40%. Repeated test numbers (MPR/PMR) become ``num#k`` tokens —
Excel Table headers must be unique — and blank columns B:E fold shut.
"""

from __future__ import annotations

import re

import numpy as np

from deps.exporter.charts import WAFER_DIE_WIDTH
from deps.exporter.contract import ReportHost
from deps.exporter.helpers import format_temperature_label, get_column_letter
from deps.exporter.theme import TAB_COLOR_HEXES
from deps.exporter.worksheets.test_statistics import group_order_by_temperature
from deps.exporter.writer import FOOTER_TEXT
from deps.SharedSrc import getSetting
from deps.uic_stdExporter import ReportSelection

CORRELATION_CAP = 100  # displayed matrix cap, rows AND columns (2026-11-19 user ruling)

# ONE scale per sheet with explicit num anchors (the 2026-11-14 legend
# ruling's shape). 2026-11-19 user ruling: use the modern Test Summary Cpk
# band palette's green / amber / red (soft fills, no saturated primaries).
CORRELATION_GRADIENT_MIN_HEX = "#A7F3D0"  # Cpk >= 1.67 band
CORRELATION_GRADIENT_MID_HEX = "#FEF3C7"  # Cpk >= 1 band
CORRELATION_GRADIENT_MAX_HEX = "#FECACA"  # Cpk < 1 band

# the self-comparison diagonal's border: dark slate, border-only (no fill)
# so the color scale underneath stays visible
CORRELATION_DIAGONAL_BORDER_HEX = "1F2937"

# the wafer map's grid geometry: the Y-label column sits in F, the grid
# runs from G (the sbin-family anchor layout the ruling references) at
# the wafer die width — the column headers carry the TEST ID only (2026-
# 11-19 follow-up ruling: the full "num name" lives on the row labels),
# read top-to-bottom in 90°-rotated cells, so the narrow width only has
# to fit the counts and the short header row stays compact
CORRELATION_LABEL_COL0 = 5
CORRELATION_FIRST_GRID_COL0 = 6
CORRELATION_LABEL_WIDTH = 34
CORRELATION_GRID_COL_WIDTH = WAFER_DIE_WIDTH
_HEADER_ROW_HEIGHT_PER_CHAR = 5.5  # pt of rotated 11pt text per ID char
_HEADER_ROW_HEIGHT_MAX = 409  # Excel's hard row-height cap
_HEADER_ROW_HEIGHT_MIN = 24  # the table's filter dropdown buttons need room

_HEADER_LABEL = "Test"  # the F-column corner cell (the Y-axis header slot)


def _table_name(sheet_name: str) -> str:
    """Workbook-unique, letter-starting Excel Table name for one map.

    Table names carry no spaces/punctuation and must not look like a cell
    reference; sheet names are workbook-unique, so sanitizing the sheet
    name keeps the name unique and readable. The 40-char cap mirrors
    ``charts.wafer_table_name``."""
    safe = re.sub(r"\W+", "_", sheet_name).strip("_")[:40]
    return f"Corr_{safe}" if safe else "Corr_Map"


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
    rows/columns; the legend says so when the cap bites. Repeated test
    numbers (MPR/PMR) get unique ``num#k`` tokens so Excel Table headers
    stay legal while the diagonal still identifies each row.
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
    # MPR/PMR tests legitimately share a test number, but Excel Table
    # headers must be unique (add_table rejects the WHOLE table on a
    # duplicate, leaving an empty header row). Tokenize repeats as
    # "num#2", "num#3" in display order; the same token prefixes the row
    # label, so the diagonal rule still matches row to column.
    nums = []
    seen_nums: dict[str, int] = {}
    for row in rows:
        base = str(row[0])
        occurrence = seen_nums.get(base, 0) + 1
        seen_nums[base] = occurrence
        nums.append(base if occurrence == 1 else f"{base}#{occurrence}")
    labels = [f"{num} {row[1]}".strip() for num, row in zip(nums, rows, strict=True)]
    # the 90°-rotated header cells read top-to-bottom: the header row must
    # be tall enough for the longest TEST ID only — the long "num name"
    # labels live on the rows, so the map keeps a compact footprint
    # (Excel caps the height at 409 pt; the 24 pt floor keeps the table's
    # filter dropdown buttons usable on short IDs)
    header_row_height = min(
        _HEADER_ROW_HEIGHT_MAX,
        max(
            _HEADER_ROW_HEIGHT_MIN,
            int(_HEADER_ROW_HEIGHT_PER_CHAR * max(len(num) for num in nums)) + 8,
        ),
    )
    return {
        "sheet_name": sheet_name,
        "legend": legend,
        "nums": nums,
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
    1, reading legend row 2, the test IDs rotated 90° on the anchor row
    (2026-11-19 follow-up ruling — the full "num name" lives in the
    F-column row labels, so the narrow columns only have to fit the
    counts and the header stays short), the counts as plain cached
    values from G, ONE 3_color_scale over the data grid and a
    formula-based border rule on the N x N diagonal. The whole map is
    ONE Excel Table (style: None keeps the hand-tuned theme) so the
    header filters/sorts the matrix; panes freeze at the grid's
    top-left corner and the sheet prints A4 landscape, one page wide,
    with rows 1-3 repeated."""
    ws.hide_gridlines(2)
    nums, labels, matrix = spec["nums"], spec["labels"], spec["matrix"]
    n = len(labels)
    header_row0 = 2
    last_col0 = CORRELATION_FIRST_GRID_COL0 + n - 1
    last_row0 = header_row0 + n  # header row + one row per test
    ws.merge_range(0, 0, 0, last_col0, spec["sheet_name"], theme.lookalike("title"))
    ws.merge_range(1, 0, 1, last_col0, spec["legend"], theme.lookalike("explanatory"))

    ws.set_column(CORRELATION_LABEL_COL0, CORRELATION_LABEL_COL0, CORRELATION_LABEL_WIDTH)
    ws.set_column(CORRELATION_FIRST_GRID_COL0, last_col0, CORRELATION_GRID_COL_WIDTH)
    # columns B:E carry nothing (the map anchors at F): fold them so the
    # screen layout and the printout do not waste horizontal space
    ws.set_column(1, CORRELATION_LABEL_COL0 - 1, None, None, {"hidden": True})

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
    ws.set_row(header_row0, spec["header_row_height"])

    # one Excel Table spans label column + grid, header row through the
    # last data row: the AutoFilter arrows sort/filter the matrix while
    # ``style: None`` leaves the hand-tuned formats untouched (the
    # charts.py bin tables' pattern)
    ws.add_table(
        header_row0,
        CORRELATION_LABEL_COL0,
        last_row0,
        last_col0,
        {
            "name": _table_name(spec["sheet_name"]),
            "style": None,
            "autofilter": True,
            "columns": [
                {
                    "header": header,
                    "header_format": axis_fmt if offset == 0 else axis_fmt_rotated,
                }
                for offset, header in enumerate([_HEADER_LABEL, *nums])
            ],
        },
    )
    for row_offset, label in enumerate(labels):
        row0 = header_row0 + 1 + row_offset
        ws.write(row0, CORRELATION_LABEL_COL0, label, label_fmt)
        for col_offset, value in enumerate(matrix[row_offset]):
            ws.write(row0, CORRELATION_FIRST_GRID_COL0 + col_offset, int(value), body_fmt)

    # get_column_letter is 1-based: the old 0-based args shifted the range
    # one column left and dropped the last grid column (2026-11-19 fix)
    first_letter = get_column_letter(CORRELATION_FIRST_GRID_COL0 + 1)
    last_letter = get_column_letter(last_col0 + 1)
    grid_range = f"{first_letter}{header_row0 + 2}:{last_letter}{header_row0 + 1 + n}"
    # border-only diagonal rule, anchored at the range's top-left cell:
    # the row label's leading test number vs the column header. Relative
    # refs follow Excel table sorts; no fill, so the scale shows through.
    label_letter = get_column_letter(CORRELATION_LABEL_COL0 + 1)
    label_anchor = f"${label_letter}{header_row0 + 2}"
    header_anchor = f"{first_letter}${header_row0 + 1}"
    ws.conditional_format(
        grid_range,
        {
            "type": "formula",
            "criteria": (
                f'LEFT({label_anchor},FIND(" ",{label_anchor}&" ")-1)'
                f"={header_anchor}"
            ),
            # xlsxwriter border code 3 = dashed ("hyphen" style): visible
            # without overpowering the tiny colored cells
            "format": theme.cf_border(CORRELATION_DIAGONAL_BORDER_HEX, style=3),
            "stop_if_true": False,
        },
    )
    ws.conditional_format(grid_range, dict(spec["gradient"]))
    ws.freeze_panes(header_row0 + 1, CORRELATION_FIRST_GRID_COL0)

    # print design (2026-11-19 user ruling): A4 landscape, one page wide /
    # unlimited tall, rows 1-3 (title / legend / rotated headers) repeat
    ws.set_paper(9)  # A4
    ws.set_landscape()
    ws.fit_to_pages(1, 0)
    ws.repeat_rows(0, header_row0)
    ws.print_area(0, 0, last_row0, last_col0)
    ws.set_footer(FOOTER_TEXT)
    ws.set_tab_color(f"#{TAB_COLOR_HEXES['test_correlation']}")


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
