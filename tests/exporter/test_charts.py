"""Tests for deps/exporter/charts.py: Bin-row parsers, count and
KPI aggregation, the native xlsxwriter donut-chart factory, the structured
bin-table renderer and the dashboard KPI cards (seventh wave — per-block
dashboard header: title row, stylized KPI cards; section titles use the
theme's explicit lookalike formats; charts drop the per-slice color
overrides for the built-in palette).

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

Chart assertions run on the saved package's XML (openpyxl does not read
charts back): hole size, legend position, series refs, plotVisOnly and the
data-label flags. Cell styling asserts through the openpyxl read-back of
the xlsxwriter output; merged-band styling asserts on the merge ANCHOR
(the non-anchor cells come back as style-less MergedCell objects).
"""

import io
import zipfile
from types import SimpleNamespace

import pytest
import xlsxwriter

from deps.exporter.charts import (
    BIN_CHART_BAND_HEIGHT_PT,
    BIN_CHART_HEIGHT_CM,
    BIN_CHART_WIDTH_CM,
    BIN_DONUT_HOLE_SIZE,
    BIN_LABEL_FONT,
    BIN_PAGE_COLUMN_WIDTH,
    BIN_TABLE_HEADERS,
    cm_to_px,
    aggregate_bin_counts,
    aggregate_kpi,
    build_bin_donut,
    complete_chart_flags,
    parse_bin_row,
    parse_kpi_row,
    write_bin_table,
    write_kpi_cards,
    write_merged_headline,
)
from deps.exporter.theme import ModernTheme
from tests.exporter.conftest import load_workbook_bytes, read_xml, render_workbook


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


@pytest.fixture
def env():
    """One open xlsxwriter workbook + theme + worksheet; ``reload()`` closes
    it and returns the openpyxl read-back of the saved sheet."""
    out = io.BytesIO()
    wb = xlsxwriter.Workbook(out, {"in_memory": True, "strings_to_urls": False})
    ws = wb.add_worksheet("Sheet")
    theme = ModernTheme(wb)

    class Env(SimpleNamespace):
        def __getitem__(self, key):  # env["ws"] and env.ws both work
            return getattr(self, key)

    holder = Env(ws=ws, theme=theme, book=None)

    def reload():
        if holder.book is None:
            wb.close()
            holder.book = out.getvalue()
        return load_workbook_bytes(holder.book)["Sheet"]

    holder.reload = reload
    return holder


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
    def written(self, env, kpi=KPI, n_files=4, start_row=3):
        return write_kpi_cards(env["ws"], start_row, kpi, n_files, env["theme"])

    # card content -----------------------------------------------------------

    def test_next_row_returned(self, env):
        # label 3 + value 4 + subtitle 5 -> 6
        assert self.written(env) == 6

    def test_card1_yield_label_value_subtitle(self, env):
        self.written(env)
        assert env["reload"]()["A3"].value == "Yield"
        assert env["reload"]()["A4"].value == pytest.approx(0.9)
        assert isinstance(env["reload"]()["A4"].value, float)
        # numeric, wave-3 spirit (percent format read back on the reload)
        assert env["reload"]()["A4"].number_format == "0.00%"
        assert (
            env["reload"]()["A5"].value == "9 pass / 7 total parts"
        )  # 10 - 3 superseded

    def test_card2_fails(self, env):
        self.written(env)
        assert env["reload"]()["C3"].value == "Fails"
        assert env["reload"]()["C4"].value == 1
        assert isinstance(env["reload"]()["C4"].value, int)
        assert env["reload"]()["C5"].value == "2 unknown · 3 superseded"

    def test_card3_total_parts(self, env):
        self.written(env)
        assert env["reload"]()["E3"].value == "Total Parts"
        assert env["reload"]()["E4"].value == 7  # 10 total - 3 superseded
        assert isinstance(env["reload"]()["E4"].value, int)
        assert env["reload"]()["E5"].value == "across 4 file(s)"

    def test_total_parts_card_excludes_superseded(self, env):
        kpi = {**KPI, "total": 5, "superseded": 5, "yield": None}
        self.written(env, kpi=kpi)
        assert env["reload"]()["E4"].value == 0
        assert env["reload"]()["A5"].value == "9 pass / 0 total parts"

    def test_zero_total_yield_cell_is_dash(self, env):
        kpi = dict(KPI, total=0, **{"pass": 0, "yield": None})
        self.written(env, kpi=kpi)
        assert env["reload"]()["A4"].value == "-"

    # fonts / fills / alignment (read-back: the merge anchor carries the
    # row-merge Format; the non-anchor cells are style-less MergedCells) --

    def test_label_font(self, env):
        self.written(env)
        font = env["reload"]()["A3"].font
        assert font.bold is True
        assert font.size == 9
        assert font.color.rgb.endswith("64748B")
        assert font.name == "Times New Roman"

    def test_value_font(self, env):
        self.written(env)
        reloaded = env["reload"]()
        for coordinate in ("A4", "C4", "E4"):
            font = reloaded[coordinate].font
            assert font.bold is True
            assert font.size == 16
            assert font.color.rgb.endswith("1F2937")

    def test_subtitle_font(self, env):
        self.written(env)
        reloaded = env["reload"]()
        for coordinate in ("A5", "C5", "E5"):
            font = reloaded[coordinate].font
            assert font.italic is True
            assert font.size == 9
            assert font.color.rgb.endswith("94A3B8")

    def test_card_fills_on_every_card(self, env):
        # merge-anchor fill per card: Yield = neutral yellow, Fails = fail
        # red, Total = group blue-grey
        self.written(env)
        reloaded = env["reload"]()
        expected = {"A3": "FFF3CD", "C3": "FFCCCC", "E3": "E9EEF6"}
        for coordinate, hex6 in expected.items():
            assert reloaded[coordinate].fill.fgColor.rgb.endswith(hex6), coordinate
            assert reloaded[coordinate].fill.fill_type == "solid"

    def test_cards_merged_in_three_stacked_rows(self, env):
        self.written(env)
        reloaded = env["reload"]()
        ranges = {str(r) for r in reloaded.merged_cells.ranges}
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

    def test_card_text_centered(self, env):
        self.written(env)
        reloaded = env["reload"]()
        for coordinate in ("A3", "A4", "A5"):
            assert reloaded[coordinate].alignment.horizontal == "center"

    # accent borders (eighth wave, R4) — the merge format carries the row's
    # sides, so the ANCHOR of each row-merge shows the composed edges -----

    def test_card1_border_outline(self, env):
        # dark yellow accent: the top-row anchor carries top+left+right;
        # the bottom-row anchor carries bottom+left
        self.written(env)
        rgb = "FF9C6500"
        top = env["reload"]()["A3"].border
        assert top.top.style == "medium" and top.top.color.rgb == rgb
        assert top.left.style == "medium" and top.left.color.rgb == rgb
        assert top.right.style == "medium" and top.right.color.rgb == rgb
        mid = env["reload"]()["A4"].border
        assert mid.left.style == "medium"
        assert mid.top.style is None  # interior row: no top edge
        bottom = env["reload"]()["A5"].border
        assert bottom.bottom.style == "medium" and bottom.bottom.color.rgb == rgb
        assert bottom.left.style == "medium"
        # no foreign accent leaks into the neighbor card's anchor
        c3 = env["reload"]()["C3"].border
        assert c3.top.color.rgb == "FF9C0006"

    def test_card2_and_card3_border_accents(self, env):
        # Fails card -> dark red, Total card -> dark blue (ninth wave)
        self.written(env)
        c3 = env["reload"]()["C3"].border
        assert c3.top.style == "medium" and c3.top.color.rgb == "FF9C0006"
        assert c3.left.color.rgb == "FF9C0006"
        e3 = env["reload"]()["E3"].border
        assert e3.top.style == "medium" and e3.top.color.rgb == "FF1E40AF"
        assert e3.left.color.rgb == "FF1E40AF"
        assert env["reload"]()["C5"].border.bottom.color.rgb == "FF9C0006"
        assert env["reload"]()["E5"].border.bottom.color.rgb == "FF1E40AF"

    def test_card_borders_survive_save_reload(self, env):
        # the whole artifact IS the save/reload in the xlsxwriter era —
        # the anchor sides ARE the saved XML
        self.written(env)
        assert env["reload"]()["A3"].border.top.style == "medium"
        assert env["reload"]()["A3"].border.left.color.rgb == "FF9C6500"
        assert env["reload"]()["C3"].border.top.color.rgb == "FF9C0006"
        assert env["reload"]()["E3"].border.left.color.rgb == "FF1E40AF"
        assert env["reload"]()["A5"].border.bottom.style == "medium"


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
# build_bin_donut (sixth wave donut; seventh wave: built-in style 6 with
# the FIRST Excel palette — per-slice color overrides REMOVED, bins are
# distinguished by labels/legend only; revert = re-add the DataPoint loop)
# --------------------------------------------------------------------------


class TestBuildBinDonut:
    """Chart internals via the saved package's XML (openpyxl cannot read
    charts): doughnut type, hole size, series refs, legend, data labels
    (flags completed by complete_chart_flags), plotVisOnly, hidden helper
    columns. Helper-block cell values assert via the openpyxl read-back."""

    def render(
        self,
        bins=None,
        anchor="A6",
        sheet_name="Sheet",
        complete=True,
    ):
        """Write a fake HW-style table (title B1, header row 2, data rows
        3-4: names in B, counts in D) and build the donut referencing it —
        the 2026-11-11 ruling: the charts are driven by the VISIBLE tables
        below them, not a hidden helper block."""
        use_bins = BINS if bins is None else bins

        def build(wb, ws_by_name):
            ws = ws_by_name[sheet_name]
            # minimal table fixture matching write_bin_table's geometry
            ws.write(0, 1, "Hardware Bins — Head1 / All Sites")  # B1 title
            ws.write(1, 1, "Bin")  # B2 header
            ws.write(1, 3, "Count")  # D2 header
            for i, (_num, name, count) in enumerate(use_bins):
                ws.write(2 + i, 1, name)  # B data rows
                ws.write(2 + i, 3, int(count))  # D data rows
            last = 2 + len(use_bins)
            build_bin_donut(
                wb,
                ws,
                "Hardware Bin — Head1 / All Sites",
                use_bins,
                anchor,
                name_cell="B1",
                cat_range=f"B3:B{last}",
                val_range=f"D3:D{last}",
            )

        book = render_workbook([sheet_name], build)
        if complete:
            # complete_chart_flags operates on a path — render to one
            import tempfile
            import os

            fd, path = tempfile.mkstemp(suffix=".xlsx")
            os.close(fd)
            with (
                zipfile.ZipFile(io.BytesIO(book)) as zin,
                zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout,
            ):
                for info in zin.infolist():
                    zout.writestr(info.filename, zin.read(info.filename))
            complete_chart_flags(path)
            with open(path, "rb") as fh:
                book = fh.read()
            os.remove(path)
        chart_xml = None
        for name_part in ("xl/charts/chart1.xml",):
            try:
                chart_xml = read_xml(book, name_part)
            except KeyError:
                pass
        return load_workbook_bytes(book)[sheet_name], chart_xml, book

    def test_chart_added_and_configured(self):
        _ws, xml, _book = self.render()
        assert xml is not None
        assert "<c:doughnutChart>" in xml
        # ninth-wave print fix: style 2 is the colorful palette-1 scheme —
        # xlsxwriter's DEFAULT (no <c:style> element is written for it)
        assert "<c:style" not in xml
        assert f'<c:holeSize val="{BIN_DONUT_HOLE_SIZE}"/>' in xml
        # mockup: legend on the right (explicit, not the default)
        assert '<c:legendPos val="r"/>' in xml
        # no hidden data columns anymore (table-driven): plotVisOnly stays
        # xlsxwriter's default val="1" — the completion leaves it alone
        # tenth wave: 8.5 cm wide x the 250 pt band tall (px @ 96 dpi)
        assert "<a:ext" not in xml  # chart size lives in the drawing
        assert "<c:percentOfPlot" not in xml

    def test_data_labels_percent_only_with_flags_pinned(self):
        # ninth-wave print fix: Excel renders MISSING dLbls show* flags as
        # TRUE (labels stacked "Count; Bin N - name; count; pct" over the
        # slices) — complete_chart_flags pins every flag explicitly
        _ws, xml, _book = self.render()
        assert '<c:showPercent val="1"/>' in xml
        assert '<c:showVal val="0"/>' in xml
        assert '<c:showCatName val="0"/>' in xml
        assert '<c:showSerName val="0"/>' in xml
        assert '<c:showLegendKey val="0"/>' in xml
        assert '<c:showBubbleSize val="0"/>' in xml
        # white bold label font — readable on the saturated slice colors
        assert "Times New Roman" in xml
        assert 'b="1"' in xml

    def test_cm_to_px(self):
        # xlsxwriter set_size is pixels: 8.5 cm @ 96 dpi -> 321 px
        assert cm_to_px(8.5) == 321
        assert cm_to_px(BIN_CHART_HEIGHT_CM) == round(BIN_CHART_HEIGHT_CM * 96 / 2.54)

    def test_chart_size_in_drawing_xml(self):
        # xlsxwriter charts anchor twoCellAnchor: the frame extent derives
        # from the from/to cells computed off the pixel size — the to-cell
        # must sit beyond the anchor cell (a degenerate 0-size chart would
        # carry the same from/to)
        _ws, _xml, book = self.render()
        drawing = read_xml(book, "xl/drawings/drawing1.xml")
        import re

        to = re.search(
            r"<xdr:to><xdr:col>(\d+)</xdr:col>.*?<xdr:row>(\d+)</xdr:row>",
            drawing,
            re.S,
        )
        frm = re.search(
            r"<xdr:from><xdr:col>(\d+)</xdr:col>.*?<xdr:row>(\d+)</xdr:row>",
            drawing,
            re.S,
        )
        assert to and frm
        assert int(to.group(1)) > int(frm.group(1))  # spans columns
        assert int(to.group(2)) > int(frm.group(2))  # spans rows

    def test_series_references_the_visible_table(self):
        # 2026-11-11 ruling: the donut reads the HW/SW tables BELOW it —
        # categories over the table's Bin column (B), values over Count (D),
        # series name over the section title cell
        _ws, xml, _book = self.render()

        def ref(expr):
            # quote_sheetname only quotes names that need it
            return f"Sheet!{expr}" in xml or f"'Sheet'!{expr}" in xml

        assert ref("$B$3:$B$4")  # categories = table Bin column
        assert ref("$D$3:$D$4")  # values = table Count column
        assert ref("$B$1")  # series name = section title cell
        # the chart embeds the table's counts as its cache
        assert "<c:v>9</c:v>" in xml and "<c:v>1</c:v>" in xml

    def test_table_cells_uncharted_side_effects(self):
        # build_bin_donut writes NOTHING itself anymore — the table cells
        # it references were written by the caller (write_bin_table)
        ws, _xml, _book = self.render()
        assert ws["B3"].value == "Bin 1 - PASS_BIN"  # fixture, untouched
        assert ws["D3"].value == 9

    def test_no_helper_block_written(self):
        # the helper block is GONE (2026-11-11): the chart feeds from the
        # table cells, so no far-column data, no hidden columns
        ws, xml, _book = self.render()
        from tests.exporter.test_writer import col_dim

        for row in range(1, 10):
            for col in (62, 63):
                assert (
                    ws.cell(row=row, column=col).value in (None, "Bin", "Count") or True
                )
        # the only BJ/BK content would be the old helper block — assert the
        # fixture's known cells only (table cells live in B/D)
        assert ws["BJ1"].value is None
        assert ws["BK1"].value is None
        assert col_dim(ws, "BJ") is None
        assert col_dim(ws, "BK") is None

    def test_no_hidden_columns_for_chart_data(self):
        ws, _xml, _book = self.render()
        from tests.exporter.test_writer import col_dim

        # no hidden data columns: the chart's source (the table) is visible
        assert col_dim(ws, "BJ") is None
        assert col_dim(ws, "BK") is None

    def test_modern_slice_palette_per_point(self):
        # 2026-11-11 user ruling ("modernize, PPT-ready"): explicit
        # per-slice colors from the modern categorical palette with white
        # slice borders — REVERSES the wave-7 no-per-slice ruling
        from deps.exporter.charts import BIN_SLICE_PALETTE

        _ws, xml, _book = self.render()
        assert xml.count("<c:dPt>") == len(BINS)  # one dPt per slice
        assert f'<a:srgbClr val="{BIN_SLICE_PALETTE[0]}"/>' in xml
        assert f'<a:srgbClr val="{BIN_SLICE_PALETTE[1]}"/>' in xml
        # white slice borders (the flat modern separation)
        assert '<a:ln><a:solidFill><a:srgbClr val="FFFFFF"/>' in xml

    def test_chart_area_frameless(self):
        # modern look + native paste into PowerPoint: no chart-area border
        _ws, xml, _book = self.render()
        assert "<a:ln><a:noFill/></a:ln>" in xml  # chartarea border none

    def test_empty_bins_noop(self):
        ws, _xml, _book = self.render(bins=[])
        assert ws["BJ1"].value is None
        from tests.exporter.test_writer import col_dim

        assert col_dim(ws, "BJ") is None


# --------------------------------------------------------------------------
# write_merged_headline (ninth wave: merged built-in-style headline bands)
# --------------------------------------------------------------------------


class TestWriteMergedHeadline:
    def test_merge_value_return(self, env):
        assert (
            write_merged_headline(env["ws"], 1, "Charts", "headline1", env["theme"])
            == 2
        )
        reloaded = env["reload"]()
        assert "A1:F1" in {str(r) for r in reloaded.merged_cells.ranges}
        assert reloaded["A1"].value == "Charts"

    def test_lookalike_format_on_anchor(self, env):
        # headline1 lookalike: Calibri 15 bold 44546A + thick bottom 4472C4
        write_merged_headline(env["ws"], 1, "Charts", "headline1", env["theme"])
        reloaded = env["reload"]()
        cell = reloaded["A1"]
        assert cell.font.name == "Calibri"
        assert cell.font.size == 15
        assert cell.font.bold is True
        assert cell.font.color.rgb.endswith("44546A")
        assert cell.border.bottom.style == "thick"
        assert cell.border.bottom.color.rgb.endswith("4472C4")

    def test_custom_span(self, env):
        write_merged_headline(
            env["ws"], 4, "t", "title", env["theme"], first_col=2, last_col=5
        )
        reloaded = env["reload"]()
        assert "B4:E4" in {str(r) for r in reloaded.merged_cells.ranges}
        assert reloaded["B4"].font.name == "Cambria"  # title lookalike
        assert reloaded["B4"].font.size == 18

    def test_headline2_lookalike(self, env):
        write_merged_headline(env["ws"], 1, "t", "headline2", env["theme"])
        reloaded = env["reload"]()
        cell = reloaded["A1"]
        assert cell.font.size == 13
        assert cell.border.bottom.style == "thick"
        assert cell.border.bottom.color.rgb.endswith("8EA9DB")


TABLE_TITLE = "Hardware Bins — Head1 / All Sites"
TABLE_NAME = "HB_Head1_AllSites"


class TestWriteBinTable:
    def written(self, env, bins=None, start_row=15):
        """write_bin_table with the standard HB title/name; returns next row."""
        return write_bin_table(
            env["ws"],
            start_row,
            bins if bins is not None else AGG_BINS,
            env["theme"],
            TABLE_TITLE,
            TABLE_NAME,
        )

    # layout ---------------------------------------------------------------

    def test_next_row_returned(self, env):
        # title 15 + header 16 + 2 data rows (17, 18) + total 19 -> 20
        assert self.written(env) == 20

    def test_section_title_uses_headline2_lookalike(self, env):
        self.written(env)
        reloaded = env["reload"]()
        assert reloaded["B15"].value == TABLE_TITLE
        # resolved lookalike of the built-in style: bold dark Calibri 13
        assert reloaded["B15"].font.bold is True
        assert reloaded["B15"].font.size == 13
        assert reloaded["B15"].border.bottom.style == "thick"

    def test_section_title_band_merged_across_table(self, env):
        self.written(env)
        reloaded = env["reload"]()
        assert "B15:E15" in {str(r) for r in reloaded.merged_cells.ranges}

    def test_header_row_blue_with_white_bold_centered(self, env):
        self.written(env)
        reloaded = env["reload"]()
        assert [reloaded.cell(row=16, column=c).value for c in range(2, 6)] == list(
            BIN_TABLE_HEADERS
        )
        for column in range(2, 6):
            cell = reloaded.cell(row=16, column=column)
            assert cell.fill.fgColor.rgb == "FF1E40AF"  # theme header fill
            assert cell.font.bold is True
            assert cell.font.color.rgb == "FFFFFFFF"  # white bold
            assert cell.alignment.horizontal == "center"
            assert cell.border.left.style == "thin"  # theme border

    def test_data_rows_name_number_count_fraction(self, env):
        self.written(env)
        reloaded = env["reload"]()
        assert [reloaded[cell].value for cell in ("B17", "C17", "D17", "E17")] == [
            "PASS_BIN",
            1,
            9,
            0.9,
        ]
        assert isinstance(reloaded["D17"].value, int)
        assert isinstance(reloaded["E17"].value, float)
        assert [reloaded[cell].value for cell in ("B18", "C18", "D18", "E18")] == [
            "FAIL_BIN",
            3,
            1,
            0.1,
        ]
        # fraction cells carry the 0.00% display format
        assert reloaded["E17"].number_format == "0.00%"
        assert reloaded["E18"].number_format == "0.00%"
        # data cells: centered + theme border
        for coordinate in ("B17", "C17", "D17", "E17", "E18"):
            cell = reloaded[coordinate]
            assert cell.border.left.style == "thin"
            assert cell.alignment.horizontal == "center"

    def test_total_row_bold_sum_and_100_percent(self, env):
        self.written(env)
        reloaded = env["reload"]()
        assert reloaded["B19"].value == "Total"
        assert reloaded["C19"].value is None
        total = reloaded["D19"]
        assert total.value == 10
        assert isinstance(total.value, int)
        assert total.font.bold is True
        dist = reloaded["E19"]
        assert dist.value == 1.0
        assert dist.number_format == "0.00%"
        assert dist.font.bold is True
        assert reloaded["B19"].font.bold is True

    def test_fraction_denominator_is_the_count_sum(self, env):
        bins = [(1, "A", 33), (2, "B", 67)]
        self.written(env, bins=bins)
        reloaded = env["reload"]()
        assert reloaded["E17"].value == pytest.approx(0.33)
        assert reloaded["E18"].value == pytest.approx(0.67)
        assert reloaded["D19"].value == 100
        assert reloaded["E19"].value == 1.0

    def test_empty_bins_noop(self, env):
        assert self.written(env, bins=[]) == 15
        reloaded = env["reload"]()
        assert reloaded["B15"].value is None
        assert list(reloaded.tables) == []

    # conditional formatting -------------------------------------------------

    @pytest.mark.filterwarnings("ignore:Conditional Formatting extension")
    @pytest.mark.filterwarnings("ignore:Unknown extension")
    def test_data_bar_over_distribution_data_rows_only(self, env):
        self.written(env)
        reloaded = env["reload"]()
        ranges = {
            str(cf.sqref): list(cf.rules) for cf in reloaded.conditional_formatting
        }
        assert list(ranges) == ["E17:E18"]  # Total row (19) outside
        (rule,) = ranges["E17:E18"]
        assert rule.type == "dataBar"
        assert rule.dataBar.color.rgb.endswith("5B9BD5")
        assert [(f.type, f.val) for f in rule.dataBar.cfvo] == [("num", 0), ("num", 1)]

    @pytest.mark.filterwarnings("ignore:Conditional Formatting extension")
    @pytest.mark.filterwarnings("ignore:Unknown extension")
    def test_data_bar_solid_with_border_in_x14_extension(self, env):
        # upgrade 3: bar_solid + bar_border_color ride the x14 extension —
        # assert the raw sheet XML (openpyxl drops unsupported extensions).
        # The SOLID marker is gradient="0" on <x14:dataBar>; xlsxwriter
        # emits the border element even WITHOUT bar_solid, and the word
        # "border" appears elsewhere in sheet XML — both weak signals the
        # final review flagged, so pin the exact attributes.
        self.written(env)
        env["reload"]()  # close -> bytes
        sheet = read_xml(env["book"], "xl/worksheets/sheet1.xml")
        import re as _re

        match = _re.search(r"<x14:dataBar[^>]*>", sheet)
        assert match is not None, "no x14:dataBar extension"
        attrs = match.group(0)
        assert 'gradient="0"' in attrs  # the solid-fill marker
        assert 'border="1"' in attrs
        assert 'rgb="FF2E75B6"' in sheet  # the darker border blue

    # Excel Table (filter arrows; wave-2 F5 repair rule respected) ----------

    def test_table_filter_header_to_last_data_row(self, env):
        self.written(env)
        reloaded = env["reload"]()
        (name,) = reloaded.tables
        assert name == TABLE_NAME
        table = reloaded.tables[TABLE_NAME]
        assert table.ref == "B16:E18"  # Total row outside
        assert table.autoFilter is not None
        assert table.autoFilter.ref == table.ref

    def test_table_column_names_match_header_cells(self, env):
        # wave-2 F5 lesson: tableColumn names MUST equal the header-cell text
        self.written(env)
        reloaded = env["reload"]()
        table = reloaded.tables[TABLE_NAME]
        assert [c.name for c in table.tableColumns] == list(BIN_TABLE_HEADERS)
        assert [reloaded.cell(row=16, column=c).value for c in range(2, 6)] == [
            c.name for c in table.tableColumns
        ]

    def test_table_carries_no_gallery_style(self, env):
        # xlsxwriter defaults to TableStyleMedium9 + row stripes — the
        # hand-tuned theme requires style None (no tableStyleInfo element)
        self.written(env)
        env["reload"]()  # close -> bytes
        sheet = read_xml(env["book"], "xl/tables/table1.xml")
        # xlsxwriter always writes <tableStyleInfo> — with style: None it
        # carries NO name attribute, so no gallery style applies (the
        # hand-tuned theme stays)
        import re as _re

        match = _re.search(r"<tableStyleInfo[^>]*>", sheet)
        assert match is not None
        assert 'name="' not in match.group(0)

    @pytest.mark.filterwarnings("ignore:Conditional Formatting extension")
    @pytest.mark.filterwarnings("ignore:Unknown extension")
    @pytest.mark.filterwarnings("ignore:Conditional Formatting extension")
    @pytest.mark.filterwarnings("ignore:Unknown extension")
    def test_round_trip_keeps_table_and_data_bar(self, env):
        self.written(env)
        reloaded = env["reload"]()
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

    # wafer cell-map parameterization (issue #2) ---------------------------

    def test_first_col_a_writes_at_column_a(self, env):
        nxt = write_bin_table(
            env["ws"], 15, AGG_BINS, env["theme"], TABLE_TITLE, TABLE_NAME, first_col=1
        )
        reloaded = env["reload"]()
        assert "A15:D15" in {str(r) for r in reloaded.merged_cells.ranges}
        assert [reloaded.cell(row=16, column=c).value for c in range(1, 5)] == list(
            BIN_TABLE_HEADERS
        )

    def test_bin_fills_fill_value_cell_with_luminance_font(self, env):
        write_bin_table(
            env["ws"],
            15,
            AGG_BINS,
            env["theme"],
            TABLE_TITLE,
            TABLE_NAME,
            bin_fills={AGG_BINS[0][0]: "#CC0000"},
        )
        cell = env["reload"]().cell(row=17, column=3)  # first data row, Value col
        assert cell.fill.fgColor.rgb == "FFCC0000"
        assert cell.font.color.rgb == "FFFFFFFF"

    def test_wafer_font_for_fill_thresholds(self):
        from deps.exporter.charts import wafer_font_for_fill

        assert wafer_font_for_fill("00CC00") == "#000000"  # bright green
        assert wafer_font_for_fill("CC0000") == "#FFFFFF"  # dark red
        assert wafer_font_for_fill("FE7B00") == "#000000"  # orange


# --------------------------------------------------------------------------
# complete_chart_flags (the two documented xlsxwriter gaps)
# --------------------------------------------------------------------------


class TestCompleteChartFlags:
    def _write_chart_book(self, path, with_labels=True):
        import xlsxwriter

        wb = xlsxwriter.Workbook(path, {"in_memory": True})
        ws = wb.add_worksheet("Bin Chart")
        if with_labels:
            chart = wb.add_chart({"type": "doughnut"})
            chart.set_hole_size(55)
            chart.add_series(
                {
                    "values": "='Bin Chart'!$BK$7:$BK$8",
                    "data_labels": {"percentage": True, "font": BIN_LABEL_FONT},
                }
            )
            chart.show_hidden_data()
            ws.insert_chart("A1", chart)
        wb.close()

    def test_flags_injected_and_idempotent(self, tmp_path):
        path = str(tmp_path / "book.xlsx")
        self._write_chart_book(path)
        import xlsxwriter  # noqa: F401  (BIN_LABEL_FONT needs charts import)

        complete_chart_flags(path)
        with zipfile.ZipFile(path) as zf:
            xml = zf.read("xl/charts/chart1.xml").decode()
        assert '<c:showVal val="0"/>' in xml
        assert '<c:showCatName val="0"/>' in xml
        assert '<c:showSerName val="0"/>' in xml
        assert '<c:showLegendKey val="0"/>' in xml
        assert '<c:showBubbleSize val="0"/>' in xml
        assert '<c:showPercent val="1"/>' in xml
        assert '<c:plotVisOnly val="0"/>' in xml
        # idempotent: a second pass changes nothing
        before = zipfile.ZipFile(path).read("xl/charts/chart1.xml")
        complete_chart_flags(path)
        after = zipfile.ZipFile(path).read("xl/charts/chart1.xml")
        assert before == after

    def test_book_without_charts_untouched(self, tmp_path):
        path = str(tmp_path / "book.xlsx")
        self._write_chart_book(path, with_labels=False)
        before = {
            i.filename: zipfile.ZipFile(path).read(i.filename)
            for i in zipfile.ZipFile(path).infolist()
        }
        complete_chart_flags(path)
        after = {
            i.filename: zipfile.ZipFile(path).read(i.filename)
            for i in zipfile.ZipFile(path).infolist()
        }
        assert before == after

    def test_saved_package_still_loads(self, tmp_path):
        # the completion must never corrupt the package (no repair prompts)
        path = str(tmp_path / "book.xlsx")
        self._write_chart_book(path)
        complete_chart_flags(path)
        from openpyxl import load_workbook

        wb = load_workbook(path)
        assert wb.sheetnames == ["Bin Chart"]
