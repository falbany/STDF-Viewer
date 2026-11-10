"""Tests for deps/modernExporter/charts.py: Bin-row parsers, count
aggregation and the native openpyxl pie-chart factory (fifth wave — the
Bin Chart sheet swaps the embedded PNGs for two real PieCharts fed by
hidden helper columns).

Data shapes confirmed upstream (DataInterface.py:596-608 +
STDF-Viewer.py:1600-1604): one waitForDataList(tab.Bin) row PER FILE,
all strings — [VHeader, "Yield: {pct}%", "Total: {n}", "Pass: {c}",
"Failed: {c}", "Unknown: {c}", "Superseded: {c}", then one cell per
non-zero bin: '"{BIN_NAME}"\\nBin {num}: {pct}%']. Percentages carry ONE
decimal (upstream ``{:.1f}%`` formatting).

Chart assertions run on in-memory openpyxl objects (charts survive a
save/reload, but the in-memory object is the artifact under test here).
"""

import pytest
from openpyxl import Workbook
from openpyxl.chart import PieChart

from deps.modernExporter.charts import (
    aggregate_bin_counts,
    bin_label,
    build_bin_pie,
    parse_bin_row,
)


def bin_row(vheader="Hardware Bin / Head1 / All Sites", total=10,
            bins=(("PASS_BIN", 1, 90.0), ("FAIL_BIN", 3, 10.0))):
    """One classic waitForDataList(tab.Bin) row (all strings)."""
    return ([vheader, "Yield: 90.00%", f"Total: {total}", "Pass: 9",
             "Failed: 1", "Unknown: 0", "Superseded: 0"]
            + [f'"{name}"\nBin {num}: {pct}%' for name, num, pct in bins])


BINS = [(1, "Bin 1 - PASS_BIN", 9), (3, "Bin 3 - FAIL_BIN", 1)]
AGG_BINS = [(1, "PASS_BIN", 9), (3, "FAIL_BIN", 1)]   # raw names from the rows
COLORS = {1: "#00CC00", 3: "#CC0000"}


# --------------------------------------------------------------------------
# parse_bin_row
# --------------------------------------------------------------------------

class TestParseBinRow:
    def test_classic_row(self):
        parsed = parse_bin_row(bin_row())
        assert parsed["total"] == 10
        assert parsed["bins"] == [(1, "PASS_BIN", 90.0), (3, "FAIL_BIN", 10.0)]

    def test_bins_keep_row_order(self):
        # upstream iterates sorted(binSummary.keys()) — the parser preserves
        # whatever order arrives and never re-sorts
        parsed = parse_bin_row(bin_row(bins=(("B9", 9, 5.0), ("B2", 2, 95.0))))
        assert [b[0] for b in parsed["bins"]] == [9, 2]

    def test_bin_names_with_spaces_and_punctuation(self):
        parsed = parse_bin_row(bin_row(bins=(("Hot Bin - v1.2 (rev A)", 5, 100.0),)))
        assert parsed["bins"] == [(5, "Hot Bin - v1.2 (rev A)", 100.0)]

    def test_missing_total_degrades_to_none(self):
        row = bin_row()
        del row[2]   # drop the "Total: 10" cell
        parsed = parse_bin_row(row)
        assert parsed["total"] is None
        assert len(parsed["bins"]) == 2

    def test_unparseable_leading_cells_skipped(self):
        # "Yield: nan%" (no DUTs), the VHeader, garbage — none parse
        row = ["Hardware Bin / Head1 / All Sites", "Yield: nan%",
               "Total: 7", "Pass: ?", "Failed: 0"]
        parsed = parse_bin_row(row)
        assert parsed == {"total": 7, "bins": []}

    def test_unparseable_trailing_cells_skipped(self):
        row = bin_row(bins=()) + ["Total: 5", '"Broken" Bin x: 1%',
                                  "no shape at all", 42, None]
        parsed = parse_bin_row(row)
        assert parsed == {"total": 5, "bins": []}

    def test_non_string_cells_ignored(self):
        parsed = parse_bin_row(["VH", 10, None, 4.5, True])
        assert parsed == {"total": None, "bins": []}

    def test_empty_row(self):
        assert parse_bin_row([]) == {"total": None, "bins": []}

    def test_zero_total_parsed(self):
        # "Total: 0" parses (the aggregation layer decides to skip it)
        parsed = parse_bin_row(bin_row(total=0, bins=()))
        assert parsed["total"] == 0


# --------------------------------------------------------------------------
# aggregate_bin_counts
# --------------------------------------------------------------------------

class TestAggregateBinCounts:
    def test_single_row_counts_from_percentage(self):
        # per-fid count = round(pct/100 * total): 90.0% of 10 -> 9
        assert aggregate_bin_counts([bin_row()]) == AGG_BINS

    def test_rows_summed_per_bin_across_files(self):
        rows = [bin_row(vheader="Hardware Bin / Head1 / All Sites / File0",
                        total=10),
                bin_row(vheader="Hardware Bin / Head1 / All Sites / File2",
                        total=20,
                        bins=(("PASS_BIN", 1, 75.0), ("FAIL_BIN", 3, 25.0)))]
        assert aggregate_bin_counts(rows) == [
            (1, "PASS_BIN", 9 + 15), (3, "FAIL_BIN", 1 + 5)]

    def test_first_seen_name_kept(self):
        rows = [bin_row(bins=(("First Name", 1, 100.0),)),
                bin_row(bins=(("Second Name", 1, 100.0),))]
        assert aggregate_bin_counts(rows) == [(1, "First Name", 20)]

    def test_sorted_by_bin_number(self):
        rows = [bin_row(bins=(("Late", 9, 50.0), ("Early", 2, 50.0)))]
        assert aggregate_bin_counts(rows) == [
            (2, "Early", 5), (9, "Late", 5)]

    def test_rows_without_total_skipped(self):
        row = bin_row()
        del row[2]
        assert aggregate_bin_counts([row, bin_row()]) == AGG_BINS

    def test_zero_total_row_skipped(self):
        # total 0 (no DUTs) cannot reconstruct counts — skipped, not div-by-zero
        assert aggregate_bin_counts([bin_row(total=0, bins=()), bin_row()]) == AGG_BINS

    def test_rounding_from_one_decimal_percentages(self):
        # 33.3% of 100 -> 33 (upstream rounds to ONE decimal first, so the
        # reconstruction can be off by ±1 vs the true counts — documented)
        rows = [bin_row(total=100,
                        bins=(("A", 1, 33.3), ("B", 2, 66.7)))]
        assert aggregate_bin_counts(rows) == [(1, "A", 33), (2, "B", 67)]

    def test_empty_input(self):
        assert aggregate_bin_counts([]) == []

    def test_bin_number_parsed_as_int(self):
        result = aggregate_bin_counts([bin_row()])
        assert all(isinstance(num, int) for num, _, _ in result)


# --------------------------------------------------------------------------
# bin_label
# --------------------------------------------------------------------------

class TestBinLabel:
    def test_label_shape(self):
        assert bin_label(1, "PASS_BIN") == "Bin 1 - PASS_BIN"

    def test_name_with_spaces(self):
        assert bin_label(12, "Hot Bin - v1.2") == "Bin 12 - Hot Bin - v1.2"


# --------------------------------------------------------------------------
# build_bin_pie
# --------------------------------------------------------------------------

class TestBuildBinPie:
    @pytest.fixture
    def ws(self):
        return Workbook().active

    def test_chart_added_and_configured(self, ws):
        build_bin_pie(ws, "Hardware Bin — Head1 / All Sites", BINS, COLORS,
                      "A6", 62)
        assert len(ws._charts) == 1
        chart = ws._charts[0]
        assert isinstance(chart, PieChart)
        assert chart.anchor == "A6"
        assert chart.title is not None
        assert chart.style == 10
        assert chart.dataLabels.showPercent is True
        assert chart.width == 9
        assert chart.height == 7
        # hidden helper columns must still be plotted (openpyxl 3.1.5:
        # ChartBase.visible_cells_only serializes to plotVisOnly)
        assert chart.visible_cells_only is False

    def test_helper_block_written_from_anchor_row(self, ws):
        build_bin_pie(ws, "t", BINS, COLORS, "A6", 62)
        # header row + one row per bin, derived from the anchor's row
        assert ws.cell(row=6, column=62).value == "Bin"
        assert ws.cell(row=6, column=63).value == "Count"
        assert ws.cell(row=7, column=62).value == "Bin 1 - PASS_BIN"
        assert ws.cell(row=7, column=63).value == 9
        assert isinstance(ws.cell(row=7, column=63).value, int)
        assert ws.cell(row=8, column=62).value == "Bin 3 - FAIL_BIN"
        assert ws.cell(row=8, column=63).value == 1

    def test_explicit_helper_row_overrides_anchor_derivation(self, ws):
        # the generator gives the two charts of one block distinct helper
        # row ranges (both anchors sit on the same visible row)
        build_bin_pie(ws, "t", BINS, COLORS, "A6", 62, helper_row=20)
        assert ws.cell(row=20, column=62).value == "Bin"
        assert ws.cell(row=21, column=63).value == 9
        assert ws.cell(row=22, column=63).value == 1
        assert ws.cell(row=6, column=62).value is None

    def test_helper_columns_hidden(self, ws):
        build_bin_pie(ws, "t", BINS, COLORS, "A6", 62)
        assert ws.column_dimensions["BJ"].hidden is True
        assert ws.column_dimensions["BK"].hidden is True

    def test_series_references(self, ws):
        build_bin_pie(ws, "t", BINS, COLORS, "A6", 62)
        series = ws._charts[0].series[0]
        # counts column (incl. the "Count" header via titles_from_data=True
        # for a series name) + categories over the labels column
        assert series.val.numRef.f == "'Sheet'!$BK$7:$BK$8"
        # openpyxl keeps the series-name ref relative (no $ signs)
        assert series.tx.strRef.f == "'Sheet'!BK6"
        # set_categories always stores numRef in memory (openpyxl
        # convention; Excel resolves it against the string cells)
        assert series.cat.numRef.f == "'Sheet'!$BJ$7:$BJ$8"

    def test_per_slice_colors_from_settings_dict(self, ws):
        build_bin_pie(ws, "t", BINS, COLORS, "A6", 62)
        points = ws._charts[0].series[0].data_points
        assert [p.idx for p in points] == [0, 1]
        assert points[0].graphicalProperties.solidFill.srgbClr == "00CC00"
        assert points[1].graphicalProperties.solidFill.srgbClr == "CC0000"

    def test_bins_missing_from_colors_keep_theme_palette(self, ws):
        # no DataPoint -> Excel falls back to the default slice colors
        build_bin_pie(ws, "t", BINS, {}, "A6", 62)
        assert ws._charts[0].series[0].data_points in ([], None)

    def test_partial_color_dict(self, ws):
        build_bin_pie(ws, "t", BINS, {3: "#CC0000"}, "A6", 62)
        points = ws._charts[0].series[0].data_points
        assert [p.idx for p in points] == [1]
        assert points[0].graphicalProperties.solidFill.srgbClr == "CC0000"

    def test_alpha_hex_truncated_to_six_digits(self, ws):
        # isHexColor accepts #RRGGBB[AA]; srgbClr takes exactly six digits
        build_bin_pie(ws, "t", BINS, {1: "#00CC00FF"}, "A6", 62)
        points = ws._charts[0].series[0].data_points
        assert points[0].graphicalProperties.solidFill.srgbClr == "00CC00"

    def test_empty_bins_noop(self, ws):
        build_bin_pie(ws, "t", [], COLORS, "A6", 62)
        assert ws._charts == []
        assert ws.cell(row=6, column=62).value is None
        assert ws.column_dimensions["BJ"].hidden is not True
