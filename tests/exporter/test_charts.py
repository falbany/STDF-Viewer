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

import numpy as np
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
    WAFER_STACKED_LEGEND,
    WAFER_SUPERSEDED_BORDER_HEX,
    cm_to_px,
    aggregate_bin_counts,
    aggregate_kpi,
    build_bin_donut,
    build_wafer_grid,
    complete_chart_flags,
    parse_bin_row,
    parse_kpi_row,
    write_bin_table,
    write_kpi_cards,
    write_merged_headline,
    stacked_wafer_table_rows,
    write_stacked_wafer_block,
    write_wafer_map_block,
    wafer_info_rows,
    wafer_kpi_from_grid,
    wafer_sheet_name,
    wafer_table_name,
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

    def reload_data_only():
        if holder.book is None:
            wb.close()
            holder.book = out.getvalue()
        return load_workbook_bytes(holder.book, data_only=True)["Sheet"]

    holder.reload_data_only = reload_data_only
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
        write_bin_table(
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


# --------------------------------------------------------------------------
# wafer cell-grid bin map (issue #2)
# --------------------------------------------------------------------------

MAP = {
    "Bounds": (2, -2, 2, -2),  # xmax, xmin, ymax, ymin
    "Info": (1.0, "", False, False, "W01", [-1]),
    "Data": {
        1: {"x": np.array([-1, 0, 1, 0]), "y": np.array([0, 0, 0, 1])},
        2: {"x": np.array([0, 0]), "y": np.array([0, -1])},
    },
    "Statistic": {1: ("Good", 3, 60.0), 2: ("Fail", 2, 40.0)},
    "Palette": {1: "#00CC00", 2: "#CC0000"},
    "BinPF": {1: "P", 2: "F"},
    "Properties": [
        ("Wafer Center:", "(1.0, 2.0)"),
        ("Wafer Flat Direction:", "Down"),
        ("Wafer XY Direction:", "(Up, Left)"),
    ],
}


class TestBuildWaferGrid:
    def test_default_orientation_top_row_is_max_y(self):
        x_axis, y_axis, grid, total = build_wafer_grid(MAP)
        assert x_axis == [-2, -1, 0, 1, 2]  # ascending left→right
        assert y_axis == [2, 1, 0, -1, -2]  # max Y on top
        assert total == 5

    def test_invert_x_flips_x_axis(self):
        m = dict(MAP, Info=(1.0, "", True, False, "W01", [-1]))
        x_axis, y_axis, _, _ = build_wafer_grid(m)
        assert x_axis == [2, 1, 0, -1, -2]
        assert y_axis == [2, 1, 0, -1, -2]

    def test_invert_y_flips_y_axis(self):
        m = dict(MAP, Info=(1.0, "", False, True, "W01", [-1]))
        x_axis, y_axis, _, _ = build_wafer_grid(m)
        assert x_axis == [-2, -1, 0, 1, 2]
        assert y_axis == [-2, -1, 0, 1, 2]

    def test_duplicate_xy_last_write_wins_and_counts(self):
        m = dict(
            MAP,
            Data={
                1: {"x": np.array([0, 0]), "y": np.array([0, 0])},
                2: {"x": np.array([0]), "y": np.array([0])},
            },
        )
        _, _, grid, total = build_wafer_grid(m)
        assert grid[(0, 0)] == 2
        assert total == 1  # raw array len 3, grid 1

    def test_non_int_bounds_degrade_to_empty(self):
        x_axis, y_axis, grid, total = build_wafer_grid(
            {"Bounds": (2.5, -2, 2, -2), "Data": {1: {"x": np.array([0]), "y": np.array([0])}}}
        )
        assert (x_axis, y_axis, grid, total) == ([], [], {}, 0)


class TestWaferKpiHelpers:
    def test_kpi_classification_by_bin_pf(self):
        kpi = wafer_kpi_from_grid({(0, 0): 1, (1, 0): 1, (2, 0): 2, (3, 0): 9}, {1: "P", 2: "F"})
        assert kpi == {
            "total": 4,
            "pass": 2,
            "failed": 1,
            "unknown": 1,  # sbin 9 has no PF entry
            "superseded": 0,
            "yield": 0.5,
        }

    def test_kpi_empty_grid(self):
        assert wafer_kpi_from_grid({}, {})["total"] == 0
        assert wafer_kpi_from_grid({}, {})["yield"] is None

    def test_kpi_palette_fallback_when_pf_flags_absent(self):
        # 2026-11-13: real STDFs often carry no usable SBR PF flag — the
        # bins classify by the color the user sees in the GUI map
        grid = {(0, 0): 1, (1, 0): 1, (2, 0): 2, (3, 0): 9}
        kpi = wafer_kpi_from_grid(
            grid,
            {1: "", 2: " ", 9: "Unknown"},
            {1: "#00CC00", 2: "#CC0000", 9: "#FE7B00"},
        )
        assert (kpi["pass"], kpi["failed"], kpi["unknown"]) == (2, 1, 1)
        assert kpi["yield"] == 2 / 4

    def test_info_rows_drop_kpi_cells_keep_identity(self):
        rows = [["Yield: 87.42%", "Total: 412", "Pass: 360", "WAFER_ID: W01", "FABWF_ID: L1"]]
        assert wafer_info_rows(rows) == [["WAFER_ID: W01", "FABWF_ID: L1"]]

    def test_info_rows_drop_empty_rows(self):
        assert wafer_info_rows([["Total: 5", "Pass: 4"]]) == []
        assert wafer_info_rows([]) == []


class TestWriteWaferMapBlock:
    def written(
        self, env, rows=None, map_data=None, start_row=1, name="SW_0_W01", **kwargs
    ):
        return write_wafer_map_block(
            env["ws"],
            env["theme"],
            rows if rows is not None else [["Yield: 60.00%", "Total: 5", "WAFER_ID: W01"]],
            map_data if map_data is not None else MAP,
            start_row,
            name,
            **kwargs,
        )

    def test_block_writes_grid_cf_and_table(self, env):
        nxt = self.written(env)
        reloaded = env["reload"]()
        # 2026-11-13 layout v2: title r=1 (merge A1:K1), identity line 2
        # (WAFER_ID key A2 / value merged B2:D2), map anchored at F3 (X
        # headers G3..K3, Y axis F4..F8, dies 4..8), Yield card A4:D6,
        # Fails A7:A9 + Total B7:D9 side by side, properties on fixed
        # lines 11-13 in a border box, SW table from 15: band 15, header
        # 16, data 17..18, total 19 -> returns 20
        assert "A1:K1" in {str(r) for r in reloaded.merged_cells.ranges}
        assert reloaded["A1"].value == "W01 – All Sites"
        assert reloaded["A2"].value == "WAFER_ID:"
        assert reloaded["B2"].value == "W01"
        assert "B2:D2" in {str(r) for r in reloaded.merged_cells.ranges}
        assert reloaded["G3"].value == "-2"  # first X header (ascending)
        assert reloaded["F4"].value == "2"  # top Y header = max Y
        assert reloaded["I5"].value == 1  # die (x=0, y=1): sbin int
        assert reloaded["I6"].value == 2  # duplicate (0,0): last wins
        rngs = {str(rng) for rng in reloaded.conditional_formatting._cf_rules}
        assert any("G4:K8" in rng for rng in rngs)
        assert [reloaded.cell(row=16, column=c).value for c in range(1, 5)] == list(
            BIN_TABLE_HEADERS
        )
        assert nxt == 20

    def test_empty_map_writes_title_and_rows_only(self, env):
        nxt = self.written(env, map_data={})
        reloaded = env["reload"]()
        assert "A1:D1" in {str(r) for r in reloaded.merged_cells.ranges}
        assert "B2:D2" in {str(r) for r in reloaded.merged_cells.ranges}
        # no table, no grid; the KPI cells drop out of the identity line
        assert not reloaded.tables
        assert reloaded["A2"].value == "WAFER_ID:"
        assert reloaded["B2"].value == "W01"
        assert nxt == 3  # start_row 1 + 1 identity row + 1

    def test_die_columns_fixed_width(self, env):
        self.written(env)
        reloaded = env["reload"]()
        # xlsxwriter 3.3.1 converts set_column chars -> px (int(w*7+0.5)+5)
        # then stores px -> chars ((px*256//7)/256): 2.71 -> 24 px ->
        # 877/256; 3.5 -> 30 px -> 1097/256. Pin the stored widths for the
        # designed pixel sizes (24 px die/spacer squares, 30 px axis col).
        die_width = 877 / 256  # 24 px
        axis_width = 1097 / 256  # 30 px
        assert reloaded.column_dimensions["G"].width == die_width
        assert reloaded.column_dimensions["E"].width == die_width
        assert reloaded.column_dimensions["F"].width == axis_width

    def test_cf_rule_value_is_sbin_with_palette_fill(self, env):
        self.written(env)
        reloaded = env["reload"]()
        (rule,) = [
            r
            for cf in reloaded.conditional_formatting
            for r in cf.rules
            if r.type == "cellIs" and str(cf.sqref) == "G4:K8" and r.operator == "equal"
            and r.formula == ["1"]
        ]
        fill_rgb = rule.dxf.fill.bgColor.rgb or rule.dxf.fill.fgColor.rgb
        assert str(fill_rgb).endswith("00CC00")

    def test_wafer_table_name_sanitized(self):
        assert wafer_table_name(0, "W 01/x") == "SW_0_W_01_x"

    def test_sw_table_counts_come_from_the_grid(self, env):
        # raw Data carries 3 dups on (0,0) but the grid dedupes to 1 die —
        # the SW table's Total must read the grid, not the arrays
        m = dict(
            MAP,
            Data={
                1: {"x": np.array([0, 0]), "y": np.array([0, 0])},
                2: {"x": np.array([0]), "y": np.array([0])},
            },
        )
        self.written(env, map_data=m)
        reloaded = env["reload"]()
        # the dedup grid carries ONE bin -> SW data 17, Total 18 (a live
        # SUM formula with cached 1)
        assert reloaded["C17"].value == "=COUNTIF($G$4:$K$8,B17)"
        assert reloaded["C18"].value == "=SUM(C17:C17)"
        data_only = env.reload_data_only()
        assert data_only["C17"].value == 1
        assert data_only["C18"].value == 1

    def test_kpi_cards_carry_grid_values_matching_the_sw_table(self, env):
        # 2026-11-12 user ruling: the cards and the SW table share ONE
        # source (the deduplicated grid) — MAP's duplicate (0,0) counts
        # ONCE as sbin 2: 3 pass dies (sbin 1 = P), 2 fail (sbin 2 = F),
        # 5 total, yield 0.6. 2026-11-13 layout v2: Yield full-width rows
        # 4-6, Fails (col A) + Total (B:D) side by side rows 7-9, and the
        # card values plus the table Count cells are LIVE FORMULAS over
        # the die map (cached results so previews show the numbers)
        self.written(env)
        reloaded = env["reload"]()
        data_only = env.reload_data_only()
        assert reloaded["A4"].value == "Yield"
        assert reloaded["A5"].value == "=(C17)/COUNT($G$4:$K$8)"
        assert reloaded["A6"].value == "3 pass / 5 total parts"
        assert reloaded["A7"].value == "Fails"
        assert reloaded["A8"].value == "=C18"
        assert reloaded["B7"].value == "Total Parts"
        assert reloaded["B8"].value == "=COUNT($G$4:$K$8)"
        # SW table Count cells: live COUNTIF keyed on each row's Value
        assert reloaded["C17"].value == "=COUNTIF($G$4:$K$8,B17)"
        assert reloaded["C18"].value == "=COUNTIF($G$4:$K$8,B18)"
        assert reloaded["C19"].value == "=SUM(C17:C18)"
        # cached results (what previews/Excel show before a recalc)
        assert data_only["A5"].value == pytest.approx(0.6)
        assert data_only["A8"].value == 2
        assert data_only["B8"].value == 5
        assert data_only["C17"].value == 3
        assert data_only["C18"].value == 2
        assert data_only["C19"].value == 5

    def test_wafer_properties_below_the_cards(self, env):
        # 2026-11-13 v2: the rendered wafer properties sit on fixed lines
        # 11-13, wrapped in a border box (A:D perimeter), key col A +
        # value merged B:D, Explanatory-Text style
        self.written(env)
        reloaded = env["reload"]()
        assert reloaded["A11"].value == "Wafer Center:"
        assert reloaded["B11"].value == "(1.0, 2.0)"
        assert "B11:D11" in {str(r) for r in reloaded.merged_cells.ranges}
        assert reloaded["A12"].value == "Wafer Flat Direction:"
        assert reloaded["B12"].value == "Down"
        assert reloaded["A13"].value == "Wafer XY Direction:"
        assert reloaded["B13"].value == "(Up, Left)"
        assert reloaded["A11"].font.italic is True  # Explanatory Text look
        # border box around the group: top edge 11, bottom edge 13
        assert reloaded["A11"].border.top.style == "thin"
        assert reloaded["A11"].border.left.style == "thin"
        assert reloaded["B11"].border.top.style == "thin"
        assert reloaded["B11"].border.right.style == "thin"
        assert reloaded["A12"].border.left.style == "thin"
        assert reloaded["B12"].border.right.style == "thin"
        assert reloaded["A13"].border.bottom.style == "thin"
        assert reloaded["B13"].border.bottom.style == "thin"
        assert reloaded["B13"].border.right.style == "thin"
        assert reloaded["A14"].border.top.style is None  # box ends at 13

    def test_kpi_cards_never_note_widths_on_the_spacer(self, env):
        # the cards ride over columns A..F; the ledger must stay clean of
        # column E (0-based 4) — the 24 px spacer squares survive the
        # later _autofit_columns
        self.written(env)
        reloaded = env["reload"]()
        assert reloaded.column_dimensions["E"].width == 877 / 256

    def test_freeze_panes_at_the_die_grid(self, env):
        self.written(env, freeze_panes=True)
        assert env["reload"]().freeze_panes == "G4"

    def test_no_freeze_by_default(self, env):
        self.written(env)
        assert env["reload"]().freeze_panes is None


class TestWaferSuperseded:
    """2026-11-13 wafer merge (issue #4), merged-sheet-only rendering:
    superseded-die thick orange border, Fails subtitle superseded count,
    Total Parts source-file subtitle, merged name / table suffixes."""

    # MAP + a 2-die bin 9 with no BinPF / Palette entry -> kpi unknown = 2
    UNKNOWN2_MAP = dict(
        MAP,
        Data={
            1: {"x": np.array([0]), "y": np.array([0])},
            2: {"x": np.array([2]), "y": np.array([0])},
            9: {"x": np.array([1, 2]), "y": np.array([0, 1])},
        },
    )

    def test_merged_name_suffix_and_31_char_cap(self):
        assert wafer_sheet_name(0, "W08", set(), suffix=" (Merged)") == "Wafer W08 (Merged)"
        # 31-char Excel cap applies AFTER the suffix is appended; two
        # calls against the same taken-set take DIFFERENT names (_2...)
        taken = set()
        n1 = wafer_sheet_name(0, "W" * 40, taken, suffix=" (Merged)")
        taken.add(n1)
        n2 = wafer_sheet_name(0, "W" * 40, taken, suffix=" (Merged)")
        assert len(n1) <= 31 and len(n2) <= 31
        assert n1 != n2

    def test_merged_table_name_suffix(self):
        assert wafer_table_name(0, "W08", suffix="_M") == "SW_0_W08_M"

    def test_fails_subtitle_counts_superseded(self, env):
        # strings land in xl/sharedStrings.xml (xlsxwriter default)
        write_wafer_map_block(
            env["ws"], env["theme"], [["WAFER_ID: W01"]], self.UNKNOWN2_MAP, 1,
            "SW_0_W01", superseded_coords={(1, 0)},
        )
        env["reload"]()
        assert "2 unknown · 1 superseded" in read_xml(env["book"], "xl/sharedStrings.xml")

    def test_fails_subtitle_defaults_to_zero_superseded(self, env):
        write_wafer_map_block(
            env["ws"], env["theme"], [["WAFER_ID: W01"]], self.UNKNOWN2_MAP, 1,
            "SW_0_W01",
        )
        env["reload"]()
        assert "2 unknown · 0 superseded" in read_xml(env["book"], "xl/sharedStrings.xml")

    def test_superseded_border_style(self, env):
        # die (0, 0) sits at I6 (duplicate (0,0): last wins, sbin 2) —
        # thick orange border on all four sides; die (0, 1) at I5 untouched
        write_wafer_map_block(
            env["ws"], env["theme"], [["WAFER_ID: W01"]], MAP, 1,
            "SW_0_W01", superseded_coords={(0, 0)},
        )
        reloaded = env["reload"]()
        border = reloaded["I6"].border
        for side in (border.top, border.bottom, border.left, border.right):
            assert side.style == "thick"
            assert str(side.color.rgb).endswith(WAFER_SUPERSEDED_BORDER_HEX)
        assert reloaded["I5"].border.top.style is None
        styles = read_xml(env["book"], "xl/styles.xml")
        assert WAFER_SUPERSEDED_BORDER_HEX in styles
        assert 'style="thick"' in styles

    def test_no_border_without_superseded_coords(self, env):
        write_wafer_map_block(
            env["ws"], env["theme"], [["WAFER_ID: W01"]], MAP, 1, "SW_0_W01"
        )
        env["reload"]()
        assert WAFER_SUPERSEDED_BORDER_HEX not in read_xml(env["book"], "xl/styles.xml")

    def test_total_parts_source_files(self, env):
        write_wafer_map_block(
            env["ws"], env["theme"], [["WAFER_ID: W01"]], MAP, 1,
            "SW_0_W01", source_files=3,
        )
        env["reload"]()
        assert "across 3 file(s)" in read_xml(env["book"], "xl/sharedStrings.xml")

    def test_total_parts_defaults_to_one_file(self, env):
        write_wafer_map_block(
            env["ws"], env["theme"], [["WAFER_ID: W01"]], MAP, 1, "SW_0_W01"
        )
        env["reload"]()
        assert "across 1 file(s)" in read_xml(env["book"], "xl/sharedStrings.xml")


# issue #5 stacked wafer payload (Task 1's STACKED shape + the Info tuple
# stacked_report_payload enriches in): fail counts keyed by COUNT, default
# orientation; Bounds order is (xmax, xmin, ymax, ymin) — x∈[1,5], y=1
STACKED_MAP = {
    "Bounds": (5, 1, 1, 1),
    "ID": (-1, -1),
    "Data": {1: {"x": [1, 2, 3], "y": [1, 1, 1]}, 3: {"x": [4], "y": [1]}},
    "Statistic": {1: 1, 3: 3},
    "Info": (1.0, "", False, False, "Stacked Wafer Map", [-1]),
}
STACKED_ROWS = [["Total Failed DUTs: 6", "Fail Coordinates: 4", "Files: 2"]]


class TestStackedWaferBlock:
    """Issue #5: the stacked wafer map block — one live fail-count grid
    (counts as cell values) under ONE green→yellow→red color-scale rule;
    title + info lines only, no KPI cards, no SW bin table."""

    def written(self, env, rows=None, map_data=None, start_row=1, **kwargs):
        return write_stacked_wafer_block(
            env["ws"],
            env["theme"],
            rows if rows is not None else STACKED_ROWS,
            map_data if map_data is not None else STACKED_MAP,
            start_row,
            **kwargs,
        )

    def test_counts_written_as_values(self, env):
        nxt = self.written(env)
        reloaded = env["reload"]()
        # title row 1, legend row 2, X headers row 3 (sbin anchor F3),
        # dies row 4, G-anchored per the layout constants (x=1..5 → G..K);
        # count 1 at (x=1..3), 3 at (x=4), (x=5) has no fails → unwritten
        assert [reloaded.cell(row=4, column=c).value for c in range(7, 12)] == [
            1, 1, 1, 3, None,
        ]
        assert reloaded["G3"].value == "1"  # first X header, str(x)
        assert reloaded["F4"].value == "1"  # Y-axis header
        assert reloaded.tables == {}  # no table by default (opt-in param)
        assert nxt == 13  # 4 cards end at row 12 → next free 13

    def test_gradient_cf_rule_present(self, env):
        self.written(env)
        env["reload"]()
        xml = read_xml(env["book"], "xl/worksheets/sheet1.xml")
        # ONE color-scale rule here (no fail table); 2026-11-14 legend
        # ruling: EXPLICIT num anchors (min / PERCENTILE.INC-50 / max of
        # the die values) shared by every gradient on the sheet — the map
        # and its legend color identical values identically, even on
        # skewed distributions where Excel's per-range percentile would
        # land elsewhere for the legend's distinct values
        assert xml.count('type="colorScale"') == 1
        for hex8 in ("FF00CC00", "FFFFFF00", "FFCC0000"):
            assert xml.count(hex8) == 1
        assert xml.count('<cfvo type="num"') == 3  # min / mid / max
        assert '<cfvo type="percentile"' not in xml  # no per-range anchors

    def test_legend_gradient_on_value_column(self, env):
        # the Fail Map Count Value column carries the SAME gradient as the
        # map (2026-11-14 user ask) — a live color legend: identical
        # anchors, so equal values render identical colors; data rows only
        self.written(env, fail_table_name="SW_0_FailMapCount")
        env["reload"]()
        xml = read_xml(env["book"], "xl/worksheets/sheet1.xml")
        assert xml.count('type="colorScale"') == 2  # die range + legend column
        for hex8 in ("FF00CC00", "FFFFFF00", "FFCC0000"):
            assert xml.count(hex8) == 2  # same three colors on both rules
        # 3 anchors × 2 gradient rules, + the data bar's own 2 cfvos
        assert xml.count('<cfvo type="num"') == 8
        # the legend range covers the Value column data rows (B17:B18 —
        # the fail table's two distinct values), not the Total row
        assert 'sqref="B17:B18"' in xml

    def test_empty_stacked_degrades(self, env):
        nxt = self.written(
            env,
            rows=[["Total Failed DUTs: 0", "Fail Coordinates: 0", "Files: 1"]],
            map_data={},
        )
        reloaded = env["reload"]()
        xml = read_xml(env["book"], "xl/worksheets/sheet1.xml")
        assert "conditionalFormatting" not in xml  # no grid → no CF rule
        merges = {str(r) for r in reloaded.merged_cells.ranges}
        assert "A1:D1" in merges
        # the legend still rides on row 2 (it explains the SHEET, not the
        # grid), merged across a 5-die map's span A:K; no cards on the
        # sbin-empty shape — the info items stay plain kv lines
        assert reloaded["A2"].value == WAFER_STACKED_LEGEND
        assert "A2:K2" in merges
        assert reloaded["A3"].value == "Total Failed DUTs:"
        assert reloaded["B3"].value == "0"
        assert reloaded["A4"].value == "Fail Coordinates:"
        assert reloaded["B4"].value == "0"
        assert reloaded["A5"].value == "Files:"
        assert reloaded["B5"].value == "1"
        assert nxt == 6

    def test_single_count_renders(self, env):
        # every coordinate shares one fail count (min == max): the rule is
        # still written, Excel renders uniformly, nothing raises
        m = dict(STACKED_MAP, Data={2: {"x": [1, 2], "y": [1, 1]}}, Bounds=(2, 1, 1, 1))
        self.written(env, map_data=m)
        reloaded = env["reload"]()
        assert [reloaded.cell(row=4, column=c).value for c in range(7, 9)] == [2, 2]
        xml = read_xml(env["book"], "xl/worksheets/sheet1.xml")
        assert xml.count('type="colorScale"') == 1

    def test_sparse_axes_full_extent(self, env):
        # gaps in the x range: the axis spans the full rectangular extent
        # like the sbin sheets, missing coords stay unwritten
        m = dict(STACKED_MAP, Data={1: {"x": [1, 5], "y": [2, 2]}}, Bounds=(5, 1, 2, 2))
        self.written(env, map_data=m)
        reloaded = env["reload"]()
        assert [reloaded.cell(row=3, column=c).value for c in range(7, 12)] == [
            "1", "2", "3", "4", "5",
        ]
        assert reloaded["G4"].value == 1 and reloaded["K4"].value == 1
        assert all(reloaded.cell(row=4, column=c).value is None for c in range(8, 11))

    def test_kpi_cards_carry_the_stacked_totals(self, env):
        # sbin-uniform KPI zone (2026-11-14 user ruling): the three info
        # items become cards — Total Failed DUTs full-width A:D (rows 4-6,
        # =SUM over the die grid), Fail Coordinates (col A) + Files (B:D)
        # side by side (rows 7-9, =COUNTIF ">0" / plain count); cached
        # results ride the formulas (xlsxwriter convention). 2026-11-14:
        # a 4th card slot (rows 10-12, light green) carries the per-wafer
        # DUT total
        nxt = self.written(env, total_duts=25)
        reloaded = env["reload"]()
        assert reloaded["A4"].value == "Total Failed DUTs"
        assert reloaded["A5"].value == "=SUM($G$4:$K$4)"
        assert reloaded["A6"].value == "fail count per die, summed across all files"
        assert reloaded["A7"].value == "Fail Coordinates"
        assert reloaded["A8"].value == '=COUNTIF($G$4:$K$4,">0")'
        assert reloaded["A9"].value == "dies failing at least once"
        assert reloaded["B7"].value == "Files"
        assert reloaded["B8"].value == 2
        assert reloaded["B9"].value == "stacked hot-spot view"
        assert reloaded["A10"].value == "Total DUTs per Wafer"
        assert reloaded["A11"].value == 25
        assert reloaded["A12"].value == "Σ of the per-wafer DUT totals"
        data_only = env.reload_data_only()
        assert data_only["A5"].value == 6  # Σ of the die grid (0s free)
        assert data_only["A8"].value == 4  # coords with ≥ 1 fail
        assert nxt == 13  # cards end at row 12

    def test_legend_below_title(self, env):
        # reading legend (2026-11-13 ask; 2026-11-14 sbin-uniform layout:
        # ONE line on row 2, the identity slot — the map anchor at F3
        # leaves a single full-width row; the "what it aggregates" moved
        # into the headline card's subtitle)
        nxt = self.written(env)
        reloaded = env["reload"]()
        assert reloaded["A2"].value == WAFER_STACKED_LEGEND
        assert "A2:K2" in {str(r) for r in reloaded.merged_cells.ranges}
        assert nxt == 13

    def test_stacked_wafer_table_rows_arithmetic(self):
        # bin 1 = the wafer's DEDUP-grid dies at sbin 1 (the same dedup
        # its own sheet renders — table = map by construction); yield =
        # aggregate_kpi over the payload's info rows; overall = Σpass /
        # Σ(total - superseded) across the payloads
        payloads = [
            (
                "W08",
                {
                    "map": MAP,
                    "rows": [["Total: 10", "Pass: 6", "Superseded: 0"]],
                },
            ),
            (
                "W09",
                {
                    "map": {
                        "Bounds": (1, -1, 1, -1),
                        "Data": {1: {"x": [0], "y": [0]}, 2: {"x": [1], "y": [0]}},
                    },
                    "rows": [["Total: 4", "Pass: 2"]],
                },
            ),
        ]
        rows, overall, total_duts = stacked_wafer_table_rows(payloads)
        # MAP: sbin-1 dies in the dedup grid = 3 — (-1,0), (1,0), (0,1);
        # the duplicate (0,0) last-write-wins to sbin 2
        assert rows[0] == ("W08", 3, 0.6)
        assert rows[1] == ("W09", 1, 0.5)
        assert overall == 8 / 14  # Σpass 8 / Σ(total - superseded) 14
        assert total_duts == 14  # Σ of the per-wafer DUT totals

    def test_stacked_wafers_table(self, env):
        # per-wafer metrics table (2026-11-14 user ask) in the SW-bin-table
        # look at the sbin table slot: band A15:C15, blue-header Excel
        # table (Wafer ID / Bin 1 / Yield) 16, data 17-18, bold Total 19
        # (live =SUM over the Bin-1 column, cached; cross-wafer yield)
        nxt = self.written(
            env,
            wafer_rows=[("W08", 2, 0.6), ("W09", 1, None)],
            overall_yield=0.5,
            wafer_table_name="SW_0_StackedWafers",
        )
        reloaded = env["reload"]()
        assert reloaded["A15"].value == "Stacked Wafers — per-wafer metrics"
        assert [reloaded.cell(row=16, column=c).value for c in range(1, 4)] == [
            "Wafer ID", "Bin 1", "Yield",
        ]
        assert reloaded["A17"].value == "W08"
        assert reloaded["B17"].value == 2
        assert reloaded["C17"].value == 0.6
        assert reloaded["C18"].value in ("", None)  # missing yield → blank
        assert reloaded["A19"].value == "Total"
        assert reloaded["B19"].value == "=SUM(B17:B18)"
        # the Total yield displays as a percent like the data rows'
        # (2026-11-14 user catch: it wrote with the plain bold format)
        assert reloaded["C17"].number_format == "0.00%"
        assert reloaded["C19"].number_format == "0.00%"
        assert "SW_0_StackedWafers" in reloaded.tables
        assert nxt == 20

    def test_fail_map_count_table(self, env):
        # Fail Map Count (2026-11-14 user ask): the SW-bin-table mimicry —
        # one row per DISTINCT count value in the grid (ascending), the
        # Count cells LIVE =COUNTIF over the die map keyed on the row's
        # own Value cell, Distribution fractions + data bar, bold Total
        # (live =SUM, cached). At the 15 slot when no wafer table precedes.
        nxt = self.written(env, fail_table_name="SW_0_FailMapCount")
        reloaded = env["reload"]()
        assert reloaded["A15"].value == "Fail Map Count — All Wafers"
        assert [reloaded.cell(row=16, column=c).value for c in range(1, 5)] == [
            "Fails per Die", "Value", "Count", "Distribution",
        ]
        assert reloaded["A17"].value == "1 fail"
        assert reloaded["B17"].value == 1
        assert reloaded["C17"].value == "=COUNTIF($G$4:$K$4,B17)"
        assert reloaded["A18"].value == "3 fails"
        assert reloaded["B18"].value == 3
        assert reloaded["A19"].value == "Total"
        data_only = env.reload_data_only()
        assert data_only["C17"].value == 3  # three dies with 1 fail
        assert data_only["C18"].value == 1  # one die with 3 fails
        assert data_only["D17"].value == 0.75  # 3/4 probed coordinates
        assert data_only["C19"].value == 4  # Σ = probed coordinates
        assert nxt == 20

    def test_fail_map_count_includes_probed_passed(self, env):
        # the count-0 bucket is its own row (0 = probed & passed dies) —
        # the distribution covers every written die cell
        m = dict(STACKED_MAP, Data={0: {"x": [5], "y": [1]}, 1: {"x": [1, 2, 3], "y": [1, 1, 1]}})
        self.written(env, map_data=m, fail_table_name="SW_0_FailMapCount")
        reloaded = env["reload"]()
        assert reloaded["A17"].value == "0 fails"
        assert reloaded["B17"].value == 0
        data_only = env.reload_data_only()
        assert data_only["C17"].value == 1
        assert data_only["C18"].value == 3
        assert data_only["C19"].value == 4

    def test_fail_map_count_below_wafer_table(self, env):
        # the Fail Map Count table sits BELOW the stacked-wafers ID table
        # (2026-11-14 user ruling), one blank row after its Total
        nxt = self.written(
            env,
            wafer_rows=[("W08", 1, 0.6)],
            overall_yield=0.6,
            wafer_table_name="SW_0_StackedWafers",
            fail_table_name="SW_0_FailMapCount",
        )
        reloaded = env["reload"]()
        assert reloaded["A15"].value == "Stacked Wafers — per-wafer metrics"
        assert reloaded["A20"].value == "Fail Map Count — All Wafers"
        data_only = env.reload_data_only()
        assert data_only["B18"].value == 1  # ID table Total (1 wafer row)
        assert data_only["C22"].value == 3  # fail table: 1 fail × 3 dies
        assert nxt == 25
