"""Tests for deps/exporter/charts.py: Bin-row parsers, count and
KPI aggregation, the native openpyxl donut-chart factory, the structured
bin-table renderer and the dashboard KPI cards (seventh wave — per-block
dashboard header: title row, stylized KPI cards; tables use Excel's
built-in "Output" style; charts drop the per-slice color overrides for
the built-in palette).

Eighth wave (2026-11-10, printability + dashboard refinement): the two
donuts tile a fixed charts band (A{r}/D{r}, 7x6 cm, built-in style 7),
tables move right to columns B:E, KPI cards gain a medium accent outline
(``apply_card_border``), and the band geometry derives from the chart
height constant (``BIN_CHART_ROW_SPAN == 12``).

Data shapes confirmed upstream (DataInterface.py:596-608 +
STDF-Viewer.py:1600-1604): one waitForDataList(tab.Bin) row PER FILE,
all strings — [VHeader, "Yield: {pct}%", "Total: {n}", "Pass: {c}",
"Failed: {c}", "Unknown: {c}", "Superseded: {c}", then one cell per
non-zero bin: '"{BIN_NAME}"\\nBin {num}: {pct}%']. Percentages carry ONE
decimal (upstream ``{:.1f}%`` formatting). The HB and SB lists carry
IDENTICAL summary cells (both derive from getDUTCountOnConditions).

Chart assertions run on in-memory openpyxl objects (charts survive a
save/reload, but the in-memory object is the artifact under test here);
table assertions cover both the in-memory styling and the save/reload
round-trip where Excel Tables are concerned. Merged-cell styling is
asserted IN MEMORY only — openpyxl's reader recreates MergedCell objects
without styles, though the styles themselves are in the saved XML.
"""

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.chart import DoughnutChart

from deps.exporter.charts import (
    BIN_CHART_BAND_HEIGHT_PT,
    BIN_CHART_HEIGHT_CM,
    BIN_CHART_WIDTH_CM,
    BIN_DONUT_HOLE_SIZE,
    BIN_PAGE_COLUMN_WIDTH,
    BIN_TABLE_HEADERS,
    aggregate_bin_counts,
    aggregate_kpi,
    apply_card_border,
    bin_label,
    build_bin_donut,
    parse_bin_row,
    parse_kpi_row,
    write_bin_table,
    write_kpi_cards,
    write_merged_headline,
)
from deps.exporter.theme import ModernTheme


def sides(border):
    """``{side: (style, rgb) | None}`` for one Border — reload-safe (the
    reader recreates missing sides as empty ``Side()`` objects, not None)."""
    out = {}
    for name in ("top", "bottom", "left", "right"):
        side = getattr(border, name)
        out[name] = (
            None if side is None or side.style is None else (side.style, side.color.rgb)
        )
    return out


def bin_row(
    vheader="Hardware Bin / Head1 / All Sites",
    total=10,
    bins=(("PASS_BIN", 1, 90.0), ("FAIL_BIN", 3, 10.0)),
):
    """One classic waitForDataList(tab.Bin) row (all strings)."""
    return [
        vheader,
        "Yield: 90.00%",
        f"Total: {total}",
        "Pass: 9",
        "Failed: 1",
        "Unknown: 0",
        "Superseded: 0",
    ] + [f'"{name}"\nBin {num}: {pct}%' for name, num, pct in bins]


BINS = [(1, "Bin 1 - PASS_BIN", 9), (3, "Bin 3 - FAIL_BIN", 1)]
AGG_BINS = [(1, "PASS_BIN", 9), (3, "FAIL_BIN", 1)]  # raw names from the rows


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
        del row[2]  # drop the "Total: 10" cell
        parsed = parse_bin_row(row)
        assert parsed["total"] is None
        assert len(parsed["bins"]) == 2

    def test_unparseable_leading_cells_skipped(self):
        # "Yield: nan%" (no DUTs), the VHeader, garbage — none parse
        row = [
            "Hardware Bin / Head1 / All Sites",
            "Yield: nan%",
            "Total: 7",
            "Pass: ?",
            "Failed: 0",
        ]
        parsed = parse_bin_row(row)
        assert parsed == {"total": 7, "bins": []}

    def test_unparseable_trailing_cells_skipped(self):
        row = bin_row(bins=()) + [
            "Total: 5",
            '"Broken" Bin x: 1%',
            "no shape at all",
            42,
            None,
        ]
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
        rows = [
            bin_row(vheader="Hardware Bin / Head1 / All Sites / File0", total=10),
            bin_row(
                vheader="Hardware Bin / Head1 / All Sites / File2",
                total=20,
                bins=(("PASS_BIN", 1, 75.0), ("FAIL_BIN", 3, 25.0)),
            ),
        ]
        assert aggregate_bin_counts(rows) == [
            (1, "PASS_BIN", 9 + 15),
            (3, "FAIL_BIN", 1 + 5),
        ]

    def test_first_seen_name_kept(self):
        rows = [
            bin_row(bins=(("First Name", 1, 100.0),)),
            bin_row(bins=(("Second Name", 1, 100.0),)),
        ]
        assert aggregate_bin_counts(rows) == [(1, "First Name", 20)]

    def test_sorted_by_bin_number(self):
        rows = [bin_row(bins=(("Late", 9, 50.0), ("Early", 2, 50.0)))]
        assert aggregate_bin_counts(rows) == [(2, "Early", 5), (9, "Late", 5)]

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
        rows = [bin_row(total=100, bins=(("A", 1, 33.3), ("B", 2, 66.7)))]
        assert aggregate_bin_counts(rows) == [(1, "A", 33), (2, "B", 67)]

    def test_empty_input(self):
        assert aggregate_bin_counts([]) == []

    def test_bin_number_parsed_as_int(self):
        result = aggregate_bin_counts([bin_row()])
        assert all(isinstance(num, int) for num, _, _ in result)


# --------------------------------------------------------------------------
# parse_kpi_row + aggregate_kpi (seventh wave: dashboard KPI cards)
# --------------------------------------------------------------------------


class TestParseKpiRow:
    def test_classic_row(self):
        parsed = parse_kpi_row(bin_row())
        assert parsed == {
            "total": 10,
            "pass": 9,
            "failed": 1,
            "unknown": 0,
            "superseded": 0,
        }

    def test_only_recognized_labels(self):
        # the VHeader and the "Yield: ..%" cell must not contribute keys
        parsed = parse_kpi_row(bin_row())
        assert set(parsed) == {"total", "pass", "failed", "unknown", "superseded"}

    def test_partial_row_keeps_found_keys_only(self):
        parsed = parse_kpi_row(["VHeader", "Yield: nan%", "Total: 7"])
        assert parsed == {"total": 7}

    def test_unparseable_counts_skipped(self):
        # "Pass: ?" / "Failed: nan" (no DUTs) — skipped, never raises
        parsed = parse_kpi_row(["Total: 3", "Pass: ?", "Failed: nan"])
        assert parsed == {"total": 3}

    def test_non_string_cells_ignored(self):
        assert parse_kpi_row([10, None, 4.5, True]) == {}

    def test_empty_row(self):
        assert parse_kpi_row([]) == {}

    def test_zero_counts_parse(self):
        # "Total: 0" (no DUTs) parses; the aggregation layer decides yield
        parsed = parse_kpi_row(["Total: 0", "Pass: 0", "Failed: 0"])
        assert parsed == {"total": 0, "pass": 0, "failed": 0}

    def test_tolerates_surrounding_whitespace(self):
        assert parse_kpi_row(["  Total:   12  "]) == {"total": 12}

    def test_values_are_ints(self):
        parsed = parse_kpi_row(bin_row())
        assert all(isinstance(v, int) for v in parsed.values())


class TestAggregateKpi:
    def test_single_row(self):
        kpi = aggregate_kpi([bin_row()])
        assert kpi == {
            "total": 10,
            "pass": 9,
            "failed": 1,
            "unknown": 0,
            "superseded": 0,
            "yield": 0.9,
        }
        assert isinstance(kpi["yield"], float)

    def test_rows_summed_across_files(self):
        first = bin_row(vheader="Hardware Bin / Head1 / All Sites / File0", total=10)
        second = bin_row(vheader="Hardware Bin / Head1 / All Sites / File2", total=20)
        kpi = aggregate_kpi([first, second])
        assert kpi["total"] == 30
        assert kpi["pass"] == 18
        assert kpi["failed"] == 2
        assert kpi["yield"] == pytest.approx(18 / 30)

    def test_missing_keys_default_zero(self):
        # a file whose cells failed to parse contributes nothing, not NaN
        kpi = aggregate_kpi([bin_row(), ["VHeader", "Yield: nan%"]])
        assert kpi["total"] == 10
        assert kpi["pass"] == 9

    def test_empty_rows_give_zeros_and_no_yield(self):
        assert aggregate_kpi([]) == {
            "total": 0,
            "pass": 0,
            "failed": 0,
            "unknown": 0,
            "superseded": 0,
            "yield": None,
        }

    def test_zero_total_gives_yield_none(self):
        kpi = aggregate_kpi([bin_row(total=0)])
        assert kpi["total"] == 0
        assert kpi["yield"] is None

    def test_yield_computed_from_summed_pass_and_total(self):
        first = bin_row(total=10)  # 9 pass
        second = bin_row(total=5, bins=(("A", 1, 100.0),))
        second[3] = "Pass: 5"  # all 5 pass
        second[4] = "Failed: 0"
        kpi = aggregate_kpi([first, second])
        assert kpi["yield"] == pytest.approx(14 / 15)

    def test_superseded_excluded_from_yield_denominator(self):
        row = bin_row()
        row[6] = "Superseded: 2"  # total 10, pass 9 -> effective 8
        kpi = aggregate_kpi([row])
        assert kpi["superseded"] == 2
        assert kpi["yield"] == pytest.approx(9 / 8)

    def test_kpi_superseded_ge_total(self):
        # degenerate counts: effective total <= 0 -> yield None (placeholder)
        row = bin_row()
        row[2] = "Total: 2"
        row[6] = "Superseded: 2"
        assert aggregate_kpi([row])["yield"] is None


# --------------------------------------------------------------------------
# write_kpi_cards (seventh wave: stylized dashboard cards, rows 3-5)
# --------------------------------------------------------------------------

KPI = {"total": 10, "pass": 9, "failed": 1, "unknown": 2, "superseded": 3, "yield": 0.9}


class TestWriteKpiCards:
    @pytest.fixture
    def ws(self):
        return Workbook().active

    @pytest.fixture
    def theme(self):
        return ModernTheme()

    def written(self, ws, theme, kpi=KPI, n_files=4, start_row=3):
        return write_kpi_cards(ws, start_row, kpi, n_files, theme)

    # card content -----------------------------------------------------------

    def test_next_row_returned(self, ws, theme):
        # label 3 + value 4 + subtitle 5 -> 6
        assert self.written(ws, theme) == 6

    def test_card1_yield_label_value_subtitle(self, ws, theme):
        self.written(ws, theme)
        assert ws["A3"].value == "Yield"
        assert ws["A4"].value == pytest.approx(0.9)
        assert isinstance(ws["A4"].value, float)
        assert ws["A4"].number_format == "0.00%"  # numeric, wave-3 spirit
        assert ws["A5"].value == "9 pass / 7 total parts"  # 10 - 3 superseded

    def test_card2_fails(self, ws, theme):
        self.written(ws, theme)
        assert ws["C3"].value == "Fails"
        assert ws["C4"].value == 1
        assert isinstance(ws["C4"].value, int)
        assert ws["C5"].value == "2 unknown · 3 superseded"

    def test_card3_total_parts(self, ws, theme):
        self.written(ws, theme)
        assert ws["E3"].value == "Total Parts"
        assert ws["E4"].value == 7  # 10 total - 3 superseded
        assert isinstance(ws["E4"].value, int)
        assert ws["E5"].value == "across 4 file(s)"

    def test_total_parts_card_excludes_superseded(self, ws, theme):
        kpi = {**KPI, "total": 5, "superseded": 5, "yield": None}
        self.written(ws, theme, kpi=kpi)
        assert ws["E4"].value == 0
        assert ws["A5"].value == "9 pass / 0 total parts"

    def test_zero_total_yield_cell_is_dash(self, ws, theme):
        kpi = dict(KPI, total=0, **{"pass": 0, "yield": None})
        self.written(ws, theme, kpi=kpi)
        assert ws["A4"].value == "-"

    # fonts ------------------------------------------------------------------

    def test_label_font(self, ws, theme):
        self.written(ws, theme)
        font = ws["A3"].font
        assert font.bold is True
        assert font.size == 9
        assert font.color.rgb.endswith("64748B")
        assert font.name == "Times New Roman"

    def test_value_font(self, ws, theme):
        self.written(ws, theme)
        for coordinate in ("A4", "C4", "E4"):
            font = ws[coordinate].font
            assert font.bold is True
            assert font.size == 16
            assert font.color.rgb.endswith("1F2937")

    def test_subtitle_font(self, ws, theme):
        self.written(ws, theme)
        for coordinate in ("A5", "C5", "E5"):
            font = ws[coordinate].font
            assert font.italic is True
            assert font.size == 9
            assert font.color.rgb.endswith("94A3B8")

    # fills (every cell of each 3x2 card block) ------------------------------

    def test_card_fills_on_every_cell(self, ws, theme):
        # Yield = neutral yellow, Fails = fail red, Total = group blue-grey;
        # merge FIRST then style — openpyxl's merge wipes styles set before
        self.written(ws, theme)
        expected = {"A": "FFF3CD", "C": "FFCCCC", "E": "E9EEF6"}
        for letter, hex6 in expected.items():
            for row in (3, 4, 5):
                for column in (letter, chr(ord(letter) + 1)):
                    cell = ws[f"{column}{row}"]
                    assert cell.fill.fgColor.rgb.endswith(hex6), f"{column}{row}"
                    assert cell.fill.fill_type == "solid"

    def test_cards_merged_in_three_stacked_rows(self, ws, theme):
        self.written(ws, theme)
        ranges = {str(r) for r in ws.merged_cells.ranges}
        for span in (
            "A3:B3",
            "A4:B4",
            "A5:B5",
            "C3:D3",
            "C4:D4",
            "C5:D5",
            "E3:F3",
            "E4:F4",
            "E5:F5",
        ):
            assert span in ranges, span

    def test_card_text_centered(self, ws, theme):
        self.written(ws, theme)
        for coordinate in ("A3", "A4", "A5"):
            assert ws[coordinate].alignment.horizontal == "center"

    # accent borders (eighth wave, R4) ----------------------------------------

    def test_card1_border_outline(self, ws, theme):
        # each card gets a medium outline in its OWN DARK accent (ninth
        # wave — dark yellow for Yield): corners combine two sides, edge
        # cells one side
        self.written(ws, theme)
        rgb = "FF9C6500"
        assert sides(ws["A3"].border) == {
            "top": ("medium", rgb),
            "bottom": None,
            "left": ("medium", rgb),
            "right": None,
        }
        assert sides(ws["B3"].border)["top"] == ("medium", rgb)
        assert sides(ws["B3"].border)["right"] == ("medium", rgb)
        assert sides(ws["A4"].border) == {
            "top": None,
            "bottom": None,
            "left": ("medium", rgb),
            "right": None,
        }
        assert sides(ws["B4"].border)["right"] == ("medium", rgb)
        assert sides(ws["A5"].border)["bottom"] == ("medium", rgb)
        assert sides(ws["B5"].border)["bottom"] == ("medium", rgb)
        assert sides(ws["B5"].border)["right"] == ("medium", rgb)
        # no foreign accent leaks in
        assert sides(ws["C3"].border)["top"] == ("medium", "FF9C0006")

    def test_card2_and_card3_border_accents(self, ws, theme):
        # Fails card -> dark red, Total card -> dark blue (ninth wave)
        self.written(ws, theme)
        assert sides(ws["C3"].border)["top"] == ("medium", "FF9C0006")
        assert sides(ws["C3"].border)["left"] == ("medium", "FF9C0006")
        assert sides(ws["D3"].border)["right"] == ("medium", "FF9C0006")
        assert sides(ws["C5"].border)["bottom"] == ("medium", "FF9C0006")
        assert sides(ws["D5"].border)["bottom"] == ("medium", "FF9C0006")
        assert sides(ws["E3"].border)["top"] == ("medium", "FF1E40AF")
        assert sides(ws["E3"].border)["left"] == ("medium", "FF1E40AF")
        assert sides(ws["F3"].border)["right"] == ("medium", "FF1E40AF")
        assert sides(ws["E5"].border)["bottom"] == ("medium", "FF1E40AF")
        assert sides(ws["F5"].border)["right"] == ("medium", "FF1E40AF")

    def test_card_borders_survive_save_reload(self, ws, theme, tmp_path):
        # styles set POST-merge reach the saved XML on every edge cell;
        # openpyxl's reader re-propagates the anchor sides across each
        # merged row (MergedCellRange._get_borders/format), so the anchors
        # keep their outer corner sides after the round-trip
        self.written(ws, theme)
        path = tmp_path / "cards.xlsx"
        ws.parent.save(path)
        reloaded = load_workbook(path)["Sheet"]
        assert reloaded["A3"].border.top.style == "medium"
        assert reloaded["A3"].border.left.color.rgb == "FF9C6500"
        assert reloaded["C3"].border.top.color.rgb == "FF9C0006"
        assert reloaded["E3"].border.left.color.rgb == "FF1E40AF"
        assert reloaded["A5"].border.bottom.style == "medium"
        assert reloaded["A5"].border.left.color.rgb == "FF9C6500"


# --------------------------------------------------------------------------
# apply_card_border (eighth wave, R4: KPI card outline helper)
# --------------------------------------------------------------------------


class TestApplyCardBorder:
    @pytest.fixture
    def ws(self):
        return Workbook().active

    def test_corners_combine_edges_single_interior_untouched(self, ws):
        # 3x3 block B2:D4: corners carry two sides, edge-mids one side,
        # the interior C3 stays fully untouched
        apply_card_border(ws, 2, 4, 2, 4, "E9EEF6")
        rgb = "FFE9EEF6"
        assert sides(ws["B2"].border) == {
            "top": ("medium", rgb),
            "bottom": None,
            "left": ("medium", rgb),
            "right": None,
        }
        assert sides(ws["D2"].border)["top"] == ("medium", rgb)
        assert sides(ws["D2"].border)["right"] == ("medium", rgb)
        assert sides(ws["B4"].border)["bottom"] == ("medium", rgb)
        assert sides(ws["B4"].border)["left"] == ("medium", rgb)
        assert sides(ws["D4"].border) == {
            "top": None,
            "bottom": ("medium", rgb),
            "left": None,
            "right": ("medium", rgb),
        }
        assert sides(ws["C2"].border) == {
            "top": ("medium", rgb),
            "bottom": None,
            "left": None,
            "right": None,
        }
        assert sides(ws["B3"].border)["left"] == ("medium", rgb)
        assert sides(ws["D3"].border)["right"] == ("medium", rgb)
        assert sides(ws["C3"].border) == {
            "top": None,
            "bottom": None,
            "left": None,
            "right": None,
        }
        # (untouched cells hold openpyxl's default border: sides present
        # but style-less — the sides() None mapping above proves it)

    def test_single_cell_block_all_four_sides(self, ws):
        apply_card_border(ws, 1, 1, 1, 1, "FFCCCC")
        rgb = "FFFFCCCC"
        assert sides(ws["A1"].border) == {
            "top": ("medium", rgb),
            "bottom": ("medium", rgb),
            "left": ("medium", rgb),
            "right": ("medium", rgb),
        }

    def test_unmerged_block_round_trip(self, ws, tmp_path):
        # without merges every edge cell keeps its exact sides after a
        # save/reload
        apply_card_border(ws, 1, 2, 1, 3, "FFF3CD")
        path = tmp_path / "outline.xlsx"
        ws.parent.save(path)
        reloaded = load_workbook(path)["Sheet"]
        rgb = "FFFFF3CD"
        assert sides(reloaded["A1"].border) == {
            "top": ("medium", rgb),
            "bottom": None,
            "left": ("medium", rgb),
            "right": None,
        }
        assert sides(reloaded["B3"].border) == {
            "top": None,
            "bottom": ("medium", rgb),
            "left": None,
            "right": ("medium", rgb),
        }


# --------------------------------------------------------------------------
# chart band geometry (eighth wave: table-driven, derived from the height)
# --------------------------------------------------------------------------


class TestChartGeometry:
    def test_band_line_size_and_grid(self):
        # tenth wave: the two donuts sit on ONE 250 pt band row, 8.5 cm
        # wide (user rulings) x the line tall; the A:F grid drops to
        # width 15 and the charts are CELL-anchored A{row} / D{row} —
        # Excel measures the width-15 columns ≈ 2.78 cm on this machine,
        # so A..C ≈ 8.33 cm and the 8.5 cm frames tile it edge-to-edge
        assert BIN_CHART_BAND_HEIGHT_PT == 250
        assert BIN_CHART_WIDTH_CM == 8.5
        assert BIN_CHART_HEIGHT_CM == pytest.approx(250 * 2.54 / 72)
        assert BIN_PAGE_COLUMN_WIDTH == 15


# --------------------------------------------------------------------------
# bin_label
# --------------------------------------------------------------------------


class TestBinLabel:
    def test_label_shape(self):
        assert bin_label(1, "PASS_BIN") == "Bin 1 - PASS_BIN"

    def test_name_with_spaces(self):
        assert bin_label(12, "Hot Bin - v1.2") == "Bin 12 - Hot Bin - v1.2"


# --------------------------------------------------------------------------
# build_bin_donut (sixth wave donut; seventh wave: built-in style 6 with
# the FIRST Excel palette — per-slice color overrides REMOVED, bins are
# distinguished by labels/legend only; revert = re-add the DataPoint loop)
# --------------------------------------------------------------------------


class TestBuildBinDonut:
    @pytest.fixture
    def ws(self):
        return Workbook().active

    def test_chart_added_and_configured(self, ws):
        build_bin_donut(ws, "Hardware Bin — Head1 / All Sites", BINS, "A6", 62)
        assert len(ws._charts) == 1
        chart = ws._charts[0]
        assert isinstance(chart, DoughnutChart)
        assert chart.anchor == "A6"
        assert chart.title is not None
        # ninth-wave print fix: style 1 is Excel's MONOCHROME GRAY scheme
        # (the donuts rendered as flat gray rings); style 2 is the colorful
        # palette-1 scheme the design meant (verified via an Excel render)
        assert chart.style == 2
        assert chart.dataLabels.showPercent is True
        # ninth-wave print fix: Excel renders MISSING dLbls show* flags as
        # TRUE (labels stacked "Count; Bin N - name; count; pct" over the
        # slices) — every flag is pinned explicitly
        assert chart.dataLabels.showVal is False
        assert chart.dataLabels.showCatName is False
        assert chart.dataLabels.showSerName is False
        assert chart.dataLabels.showLegendKey is False
        assert chart.dataLabels.showBubbleSize is False
        # white bold label font — readable on the saturated slice colors
        assert chart.dataLabels.txPr is not None
        assert chart.width == 8.5  # tenth wave: user ruling 8.5 cm wide
        assert chart.height == pytest.approx(BIN_CHART_HEIGHT_CM)  # 250 pt line
        # mockup: donut hole (openpyxl 3.1.5 DoughnutChart.holeSize,
        # NestedMinMax 1..90 — the factory constant sits in the 50-60 band)
        assert chart.holeSize == BIN_DONUT_HOLE_SIZE
        # mockup: legend on the right (explicit, not the default)
        assert chart.legend is not None
        assert chart.legend.position == "r"
        # hidden helper columns must still be plotted (openpyxl 3.1.5:
        # ChartBase.visible_cells_only serializes to plotVisOnly)
        assert chart.visible_cells_only is False

    def test_helper_block_written_from_anchor_row(self, ws):
        build_bin_donut(ws, "t", BINS, "A6", 62)
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
        build_bin_donut(ws, "t", BINS, "A6", 62, helper_row=20)
        assert ws.cell(row=20, column=62).value == "Bin"
        assert ws.cell(row=21, column=63).value == 9
        assert ws.cell(row=22, column=63).value == 1
        assert ws.cell(row=6, column=62).value is None

    def test_helper_columns_hidden(self, ws):
        build_bin_donut(ws, "t", BINS, "A6", 62)
        assert ws.column_dimensions["BJ"].hidden is True
        assert ws.column_dimensions["BK"].hidden is True

    def test_series_references(self, ws):
        build_bin_donut(ws, "t", BINS, "A6", 62)
        series = ws._charts[0].series[0]
        # counts column (incl. the "Count" header via titles_from_data=True
        # for a series name) + categories over the labels column
        assert series.val.numRef.f == "'Sheet'!$BK$7:$BK$8"
        # openpyxl keeps the series-name ref relative (no $ signs)
        assert series.tx.strRef.f == "'Sheet'!BK6"
        # set_categories always stores numRef in memory (openpyxl
        # convention; Excel resolves it against the string cells)
        assert series.cat.numRef.f == "'Sheet'!$BJ$7:$BJ$8"

    def test_no_per_slice_data_points_builtin_palette(self, ws):
        # seventh wave (ruled consequence): the per-slice DataPoint fills
        # from the settings bin-color dicts are REMOVED so Excel's built-in
        # style-6 palette (first gallery scheme) colors the slices; bins
        # are distinguished by labels/legend only
        build_bin_donut(ws, "t", BINS, "A6", 62)
        assert not ws._charts[0].series[0].data_points

    def test_empty_bins_noop(self, ws):
        build_bin_donut(ws, "t", [], "A6", 62)
        assert ws._charts == []
        assert ws.cell(row=6, column=62).value is None
        assert ws.column_dimensions["BJ"].hidden is not True


# --------------------------------------------------------------------------
# write_merged_headline (ninth wave: merged built-in-style headline bands)
# --------------------------------------------------------------------------


class TestWriteMergedHeadline:
    @pytest.fixture
    def ws(self):
        return Workbook().active

    def test_merge_style_value_return(self, ws):
        assert write_merged_headline(ws, 1, "Charts", "Headline 1") == 2
        assert "A1:F1" in {str(r) for r in ws.merged_cells.ranges}
        assert ws["A1"].value == "Charts"
        for column in range(1, 7):
            assert ws.cell(row=1, column=column).style == "Headline 1"

    def test_custom_span(self, ws):
        write_merged_headline(ws, 4, "t", "Title", first_col=2, last_col=5)
        assert "B4:E4" in {str(r) for r in ws.merged_cells.ranges}
        for column in range(2, 6):
            assert ws.cell(row=4, column=column).style == "Title"

    def test_style_survives_save_reload(self, ws, tmp_path):
        write_merged_headline(ws, 1, "t", "Headline 2")
        path = tmp_path / "headline.xlsx"
        ws.parent.save(path)
        reloaded = load_workbook(path)["Sheet"]
        assert reloaded["A1"].style == "Headline 2"
        named = {s.name: s for s in reloaded.parent._named_styles}
        assert named["Headline 2"].builtinId == 17


# --------------------------------------------------------------------------
# write_bin_table (sixth wave: structured tables replace the classic rows)
# --------------------------------------------------------------------------

TABLE_TITLE = "Hardware Bins — Head1 / All Sites"
TABLE_NAME = "HB_Head1_AllSites"


class TestWriteBinTable:
    @pytest.fixture
    def ws(self):
        return Workbook().active

    @pytest.fixture
    def theme(self):
        return ModernTheme()

    def written(self, ws, theme, bins=None, start_row=15):
        """write_bin_table with the standard HB title/name; returns next row."""
        return write_bin_table(
            ws,
            start_row,
            bins if bins is not None else AGG_BINS,
            theme,
            TABLE_TITLE,
            TABLE_NAME,
        )

    # layout ---------------------------------------------------------------

    def test_next_row_returned(self, ws, theme):
        # title 15 + header 16 + 2 data rows (17, 18) + total 19 -> 20
        assert self.written(ws, theme) == 20

    def test_section_title_uses_builtin_headline_2_style(self, ws, theme):
        # ninth wave: the section title band carries Excel's BUILT-IN
        # "Headline 2" cell style (user's "Title 3"; openpyxl 3.1.5
        # registers it with builtinId=17, so Excel's gallery shows it as
        # the built-in style); the band spans the table width B:E (eighth
        # wave — column A belongs to the donut/KPI page grid)
        self.written(ws, theme)
        assert ws["B15"].style == "Headline 2"
        assert ws["B15"].value == TABLE_TITLE
        # resolved look of the builtin style: bold dark font
        assert ws["B15"].font.b is True

    def test_section_title_band_merged_and_styled_across_table(self, ws, theme):
        # merge B15:E15 and style EVERY cell of the band (styles must be
        # applied AFTER merging — openpyxl's merge_cells wipes styles set
        # before it)
        self.written(ws, theme)
        ranges = {str(r) for r in ws.merged_cells.ranges}
        assert "B15:E15" in ranges
        for column in range(2, 6):
            assert ws.cell(row=15, column=column).style == "Headline 2"

    def test_section_title_headline_2_style_round_trip(self, ws, theme, tmp_path):
        # the named style carries builtinId=17 into styles.xml and the
        # assignment survives a save/reload
        self.written(ws, theme)
        path = tmp_path / "out.xlsx"
        ws.parent.save(path)
        reloaded = load_workbook(path)["Sheet"]
        assert reloaded["B15"].style == "Headline 2"
        named = {s.name: s for s in reloaded.parent._named_styles}
        assert named["Headline 2"].builtinId == 17

    def test_header_row_blue_with_white_bold_centered(self, ws, theme):
        self.written(ws, theme)
        assert [ws.cell(row=16, column=c).value for c in range(2, 6)] == list(
            BIN_TABLE_HEADERS
        )
        for column in range(2, 6):
            cell = ws.cell(row=16, column=column)
            assert cell.fill.fgColor.rgb == "FF1E40AF"  # theme header fill
            assert cell.font.bold is True
            assert cell.font.color.rgb == "FFFFFFFF"  # white bold
            assert cell.alignment.horizontal == "center"
            assert cell.border.left.style == "thin"  # theme border

    def test_data_rows_name_number_count_fraction(self, ws, theme):
        self.written(ws, theme)
        assert [ws[cell].value for cell in ("B17", "C17", "D17", "E17")] == [
            "PASS_BIN",
            1,
            9,
            0.9,
        ]
        assert isinstance(ws["D17"].value, int)
        assert isinstance(ws["E17"].value, float)
        assert [ws[cell].value for cell in ("B18", "C18", "D18", "E18")] == [
            "FAIL_BIN",
            3,
            1,
            0.1,
        ]
        # fraction cells carry the 0.00% display format
        assert ws["E17"].number_format == "0.00%"
        assert ws["E18"].number_format == "0.00%"
        # data cells: body font + theme border
        for coordinate in ("B17", "C17", "D17", "E17", "E18"):
            cell = ws[coordinate]
            assert cell.border.left.style == "thin"
            assert cell.alignment.horizontal == "center"

    def test_total_row_bold_sum_and_100_percent(self, ws, theme):
        self.written(ws, theme)
        assert ws["B19"].value == "Total"
        assert ws["C19"].value is None
        total = ws["D19"]
        assert total.value == 10
        assert isinstance(total.value, int)
        assert total.font.bold is True
        dist = ws["E19"]
        assert dist.value == 1.0
        assert dist.number_format == "0.00%"
        assert dist.font.bold is True
        assert ws["B19"].font.bold is True

    def test_fraction_denominator_is_the_count_sum(self, ws, theme):
        # counts are the wave-5 reconstruction; the denominator is their own
        # sum (may differ from the true DUT total by the rounding caveat)
        bins = [(1, "A", 33), (2, "B", 67)]
        self.written(ws, theme, bins=bins)
        assert ws["E17"].value == pytest.approx(0.33)
        assert ws["E18"].value == pytest.approx(0.67)
        assert ws["D19"].value == 100
        assert ws["E19"].value == 1.0

    def test_empty_bins_noop(self, ws, theme):
        assert self.written(ws, theme, bins=[]) == 15
        assert ws.max_row == 1  # nothing written
        assert ws.conditional_formatting is not None and not list(
            ws.conditional_formatting
        )
        assert list(ws.tables) == []

    # conditional formatting -------------------------------------------------

    def test_data_bar_over_distribution_data_rows_only(self, ws, theme):
        self.written(ws, theme)
        ranges = {str(cf.sqref): list(cf.rules) for cf in ws.conditional_formatting}
        assert list(ranges) == ["E17:E18"]  # Total row (19) outside
        (rule,) = ranges["E17:E18"]
        assert rule.type == "dataBar"
        assert rule.dataBar.showValue is True  # bar AND text in the cell
        assert rule.dataBar.color.rgb.endswith("5B9BD5")
        assert [(f.type, f.val) for f in rule.dataBar.cfvo] == [("num", 0), ("num", 1)]

    # Excel Table (filter arrows; wave-2 F5 repair rule respected) ----------

    def test_table_filter_header_to_last_data_row(self, ws, theme):
        self.written(ws, theme)
        (name,) = ws.tables
        assert name == TABLE_NAME
        table = ws.tables[TABLE_NAME]
        assert table.ref == "B16:E18"  # Total row outside
        assert table.autoFilter is not None
        assert table.autoFilter.ref == table.ref

    def test_table_column_names_match_header_cells(self, ws, theme):
        # wave-2 F5 lesson: tableColumn names MUST equal the header-cell text
        self.written(ws, theme)
        table = ws.tables[TABLE_NAME]
        assert [c.name for c in table.tableColumns] == list(BIN_TABLE_HEADERS)
        assert [ws.cell(row=16, column=c).value for c in range(2, 6)] == [
            c.name for c in table.tableColumns
        ]

    def test_round_trip_keeps_table_and_data_bar(self, ws, theme, tmp_path):
        self.written(ws, theme)
        path = tmp_path / "bin.xlsx"
        ws.parent.save(path)
        reloaded = load_workbook(path)["Sheet"]
        table = reloaded.tables[TABLE_NAME]
        assert table.ref == "B16:E18"
        assert table.autoFilter.ref == table.ref
        assert [c.name for c in table.tableColumns] == list(BIN_TABLE_HEADERS)
        (rule,) = [
            r
            for cf in reloaded.conditional_formatting
            for r in cf.rules
            if str(cf.sqref) == "E17:E18"
        ]
        assert rule.type == "dataBar"
        assert rule.dataBar.color.rgb.endswith("5B9BD5")
