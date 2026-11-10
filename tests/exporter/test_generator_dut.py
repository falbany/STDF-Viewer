"""Tests for deps/exporter/generator.py: DUT Summary + chart sheets.

Task 8 integration: selectedContents={DUT, Trend, Bin, Wafer} (Histogram
deliberately NOT selected), two test items (one parametric with limits, one
FTR), three DUTs in one file. waitForDutSummary returns a classic-protocol
summaryTable (fixture style of test_matrix.py::make_summary), waitForImage a
small real PNG, waitForDataList simple stat rows. report_core() runs against
tmp_path; cell/style assertions run on the reloaded workbook, image
assertions on the in-memory one (openpyxl does not reload images from saved
files — see test_sheets.py).
"""

import io
from pathlib import Path

import numpy as np
import pytest
from openpyxl import load_workbook
from openpyxl.chart import DoughnutChart
from PIL import Image as PILImage
from PyQt5 import QtCore

from deps.SharedSrc import REC, getSetting, tab
from deps.exporter.generator import ModernReportGenerator, bin_banner_from_file_info
from deps.exporter.theme import number_format_from_settings
from deps.uic_stdExporter import ReportSelection, dataChannel, signals


def chart_title_text(chart) -> str:
    """Concatenated rich-text runs of a chart title (reload-safe)."""
    return "".join(run.t for p in chart.title.tx.rich.p for run in (p.r or []) if run.t)


def bin_table(bin_type: str) -> list[list]:
    """One classic waitForDataList(tab.Bin) row (single file): VHeader +
    yield/total/count cells + one '"name"\\nBin N: P.P%' cell per non-zero
    bin (DataInterface.py:596-608 flattened by STDF-Viewer.py:1600-1604;
    all strings)."""
    label = "Hardware Bin" if bin_type == "HB" else "Software Bin"
    return [
        [
            f"{label} / Head1 / All Sites",
            "Yield: 90.00%",
            "Total: 10",
            "Pass: 9",
            "Failed: 1",
            "Unknown: 0",
            "Superseded: 0",
            '"PASS_BIN"\nBin 1: 90.0%',
            '"FAIL_BIN"\nBin 3: 10.0%',
        ]
    ]


@pytest.fixture(scope="module")
def qapp():
    """QObject.tr needs a QCoreApplication instance (no widgets required)."""
    app = QtCore.QCoreApplication.instance()
    if app is None:
        app = QtCore.QCoreApplication([])
    yield app


# --------------------------------------------------------------------------
# synthetic GUI data (classic protocol shapes)
# --------------------------------------------------------------------------


def make_dut_info(
    fid=1,
    part_id="",
    part_txt="",
    head_site="1--1",
    tests_executed=42,
    test_time="123 ms",
    hw_bin="Bin 1",
    sw_bin="Bin 2",
    wafer="",
    xy=(3, -5),
    dut_flag="Passed",
):
    """dutInfo tuple: [File ID, Part ID, Part Text, Head-Site, Tests Executed,
    Test Time, HW Bin, SW Bin, Wafer ID, (X, Y), DUT Flag] (indices 0..10).

    Test Time / HW Bin / SW Bin use the raw DB shapes ("123 ms", "Bin 1",
    "Bin 2" — see schema.rs FETCH_SELECT_DUT_SUMMARY)."""
    return (
        fid,
        part_id,
        part_txt,
        head_site,
        tests_executed,
        test_time,
        hw_bin,
        sw_bin,
        wafer,
        xy,
        dut_flag,
    )


def make_summary():
    """Classic waitForDutSummary dict: 2 tests x 3 DUTs in file 1."""
    duts = [
        (1, 0, make_dut_info(part_id="SN541_T25C")),
        (
            1,
            1,
            make_dut_info(
                part_id="SN542_T25C", sw_bin="Bin 3", xy=(4, -6), dut_flag="Failed"
            ),
        ),
        (1, 2, make_dut_info(part_id="SN543_T85C")),
    ]
    tests = [
        # parametric with limits (data row 16)
        dict(
            name="0_TXCC/VCCC_ASIC",
            num=10,
            unit="V",
            llimit=0.5,
            hlimit=1.5,
            rec=REC.PTR,
            data={1: ([1.0, 1.2, np.nan], [0, 0, -1])},
        ),
        # FTR (data row 17): die cells from the flag list
        dict(
            name="2_Func/scanchain",
            num=12,
            unit="",
            llimit=np.nan,
            hlimit=np.nan,
            rec=REC.FTR,
            vect_nam="PAT_TOP",
            data={1: ([0.0, 1.0, 0.0], [0, 0x80, 0])},
        ),
    ]
    data, test_info, dut2ind, dut_info = {}, {}, {}, {}
    dut2ind[1] = {0: 0, 1: 1, 2: 2}
    dut_info[1] = {d[1]: d[2] for d in duts}
    for t in tests:
        test_tuple = (t["num"], 0, t["name"])
        # TestInfo order: [TEST_NAME, TEST_NUM, HLimit, LLimit, Unit]
        test_info[test_tuple] = [
            t["name"],
            t["num"],
            t["hlimit"],
            t["llimit"],
            t["unit"],
        ]
        values, flags = t["data"][1]
        entry = {
            "SUB_CODE": t["rec"],
            "dataList": np.asarray(values, dtype=float),
            "flagList": np.asarray(flags, dtype=int),
        }
        if t.get("vect_nam") is not None:
            entry["VECT_NAM"] = t["vect_nam"]
        data[test_tuple] = {1: entry}
    return {
        "Data": data,
        "TestInfo": test_info,
        "dut2ind": dut2ind,
        "dutInfo": dut_info,
    }


def png_bytes(width=96, height=48):
    """A small real PNG (96x48 scales to 1152x576 px at 12 in width)."""
    buf = io.BytesIO()
    PILImage.new("RGB", (width, height), color=(0, 0, 0)).save(buf, format="PNG")
    return buf.getvalue()


PNG_BYTES = png_bytes()


def build_generated(tmp_path, qapp, with_stat=False, with_file_info=False):
    """Run report_core() once for {DUT, Trend, Bin, Wafer} (plus Stat when
    `with_stat`, FileInfo when `with_file_info`) and return (generator,
    reloaded workbook, report path, close/error capture)."""
    testTuples = [(10, 0, "0_TXCC/VCCC_ASIC"), (12, 0, "2_Func/scanchain")]
    selectedFiles = [1]
    selectedHeads = [1]
    selectedSites = [-1]
    waferTuples = [(0, 1, "W01")]  # (wafer index, file id, wafer name)
    reportPath = str(tmp_path / "report.xlsx")
    # classic totalLoopCnt formula (stdfExporter.start): Trend 2*1*1
    # + Bin 1*1 + DUT (1 + 2) + Wafer len(waferTuples)
    totalLoopCnt = 2 + 1 + 3 + 1
    selection = [
        ReportSelection.DUT,
        ReportSelection.Trend,
        ReportSelection.Bin,
        ReportSelection.Wafer,
    ]
    if with_stat:
        # classic sheet-creation order: DUT < Stat < Trend
        selection.insert(1, ReportSelection.Stat)
        totalLoopCnt += 2  # Stat: +1 per data row
    if with_file_info:
        # classic sheet-creation order: FileInfo first (+1 loop tick)
        selection.insert(0, ReportSelection.FileInfo)
        totalLoopCnt += 1
    reportInfo = (
        selection,
        testTuples,
        selectedFiles,
        selectedHeads,
        selectedSites,
        reportPath,
        waferTuples,
        totalLoopCnt,
    )

    sig = signals()
    channel = dataChannel()
    gen = ModernReportGenerator(
        sig, QtCore.QMutex(), QtCore.QWaitCondition(), reportInfo, channel
    )

    # GUI side: classic-protocol summary table, chart PNGs, stat rows
    gen.waitForDutSummary = lambda heads, sites, fids, testTuples: make_summary()

    def fake_wait_for_image(testTuple, head, sites, fids, chartType):
        return io.BytesIO(PNG_BYTES)

    gen.waitForImage = fake_wait_for_image

    # classic DataInterface.getStatistics shape (containsFTR=True: the
    # "Pattern Name" column follows "Test Name"; VECT_NAM only on the FTR row)
    stat_table = [
        [
            "",
            "Test Name",
            "Pattern Name",
            "Unit",
            "Low Limit",
            "High Limit",
            "Fail Num",
            "Cpk",
            "Average",
            "Median",
            "St. Dev.",
            "Min",
            "Max",
        ],
        [
            "10 / Head1 / All Sites",
            "0_TXCC/VCCC_ASIC",
            "",
            "V",
            "0.500",
            "1.500",
            "1",
            "1.23",
            "1.000",
            "1.000",
            "0.010",
            "0.980",
            "1.020",
        ],
        [
            "12 / Head1 / All Sites",
            "2_Func/scanchain",
            "PAT_TOP",
            "P/F",
            "N/A",
            "N/A",
            "0",
            "∞",
            "1.000",
            "1.000",
            "0.000",
            "0.000",
            "1.000",
        ],
    ]

    def fake_wait_for_data_list(heads, sites, fids, chartType, kargs):
        if chartType == tab.Trend and kargs.get("testTuples") == testTuples:
            # writeStatistic passes the FULL testTuples list (per-test trend
            # rows pass a single-tuple list — never equal here, two tests)
            return stat_table
        if chartType == tab.Bin:
            return bin_table("HB") if kargs.get("isHBIN") else bin_table("SB")
        if chartType == tab.Wafer:
            return [["WAFER-W01-ROW", "Total DUTs", 3]]
        return [["TREND-ROW", "mean", 1.1]]

    gen.waitForDataList = fake_wait_for_data_list

    if with_file_info:
        # File Info raw shape (buildFileMetaData): MIR description labels
        # ("Product ID" / "Sublot ID" feed the Bin-sheet banner, ninth wave)
        file_info_table = [
            ["File Name", "demo.stdf"],
            ["Product ID", "ASIC-X"],
            ["Sublot ID", "LOT123"],
            ["Yield: ", "98.50%"],
            ["DUTs Tested: ", "10"],
            ["DUTs Passed: ", "9"],
            ["DUTs Failed: ", "1"],
            ["DUTs Superseded: ", "0"],
            ["DUTs Unknown: ", "0"],
        ]
        gen.waitForTableData = lambda isFileInfo: file_info_table

    closed, errors = [], []
    sig.closeSignal.connect(closed.append)
    sig.msgSignal.connect(errors.append)

    gen.report_core()

    assert Path(reportPath).exists(), "report file was not written"
    wb = load_workbook(reportPath)
    return gen, wb, reportPath, closed, errors


@pytest.fixture
def generated(tmp_path, qapp):
    return build_generated(tmp_path, qapp)


@pytest.fixture
def generated_with_file_info(tmp_path, qapp):
    return build_generated(tmp_path, qapp, with_file_info=True)


# --------------------------------------------------------------------------
# sheet inventory
# --------------------------------------------------------------------------


class TestSheetInventory:
    def test_sheets_classic_order_without_histogram(self, generated):
        _, wb, _, _, _ = generated
        assert wb.sheetnames == ["DUT Summary", "Trend Chart", "Bin Chart", "Wafer Map"]

    def test_unselected_and_overflow_sheets_absent(self, generated):
        _, wb, _, _, _ = generated
        assert "Histogram" not in wb.sheetnames  # not selected
        assert "DUT Summary (2)" not in wb.sheetnames  # 3 DUTs, single chunk


# --------------------------------------------------------------------------
# DUT Summary matrix sheet
# --------------------------------------------------------------------------


class TestDutSummarySheet:
    def test_banner_row_one(self, generated):
        _, wb, _, _, _ = generated
        ws = wb["DUT Summary"]
        assert ws["A1"].value == "Test#"  # Test# leads the base labels
        assert ws["M1"].value == 1  # banner labels are plain ints
        assert ws["A1"].fill.fgColor.rgb == "FFCBD5E1"

    def test_repeated_header_row_fifteen(self, generated):
        _, wb, _, _, _ = generated
        ws = wb["DUT Summary"]
        assert ws["A15"].value == "Test#"
        assert ws["A15"].fill.fgColor.rgb == "FF1E40AF"
        assert ws["M15"].value == "SN541"  # part ID base

    def test_metadata_block_layout(self, generated):
        _, wb, _, _, _ = generated
        ws = wb["DUT Summary"]
        # 13 metadata rows (2..14), Test# 10000.., numeric cells parsed bare
        assert ws["A2"].value == 10000
        assert ws["B2"].value == "METADATA"
        assert ws["C2"].value == "File ID"
        # Test Time row (13): parsed to bare ints under MIN/MAX formulas
        assert ws["C13"].value == "Test Time"
        assert ws["D13"].value == "ms"
        assert ws["G13"].value == "=MIN(M13:O13)"
        # always-fold rows stay hidden (Head-Site 11, Test Time 13); the DUT
        # Flag row (14) is VISIBLE since the fourth wave (status CF)
        for row in (11, 13):
            assert ws.row_dimensions[row].hidden is True
        assert ws.row_dimensions[14].hidden is not True
        assert ws["C14"].value == "DUT Flag"
        assert ws["M14"].value == "Passed"

    def test_first_test_formulas_at_row_sixteen(self, generated):
        _, wb, _, _, _ = generated
        ws = wb["DUT Summary"]
        assert ws["G16"].value.startswith("=MIN(M16:")
        assert ws["H16"].value.startswith("=MAX(M16:")
        assert ws["L16"].value.startswith("=SUMPRODUCT(")

    def test_die_values_parametric_and_ftr(self, generated):
        _, wb, _, _, _ = generated
        ws = wb["DUT Summary"]
        assert ws["M16"].value == 1.0
        assert ws["N16"].value == 1.2
        # FTR row: die cells from the flag list (0 pass / 1 fail)
        assert ws["B17"].value == "2_Func"
        assert ws["C17"].value == "PAT_TOP"
        assert ws["D17"].value == "P/F"
        assert ws["M17"].value == 0
        assert ws["N17"].value == 1

    def test_mirror_columns_hidden_after_spacer(self, generated):
        _, wb, _, _, _ = generated
        dims = wb["DUT Summary"].column_dimensions
        assert not dims["P"].hidden  # spacer stays visible
        for letter in ("Q", "R", "S"):
            assert dims[letter].hidden is True

    def test_fixed_column_widths_no_autofit(self, generated):
        # 2026-11-10 R9: DUT-sheet widths are fixed by the renderer
        _, wb, _, _, _ = generated
        dims = wb["DUT Summary"].column_dimensions
        assert dims["A"].width == 8  # Test#
        assert dims["B"].width == 30  # Test name
        assert dims["C"].width == 8  # Param
        assert dims["M"].width == 8  # die columns pinned


# --------------------------------------------------------------------------
# chart sheets
# --------------------------------------------------------------------------


class TestChartSheets:
    def test_images_embedded(self, generated):
        # fifth wave: Wafer/Trend keep the embedded PNGs — Bin no longer
        # calls waitForImage at all
        gen, _, _, _, _ = generated
        for name in ("Wafer Map", "Trend Chart"):
            assert len(gen._wb[name]._images) >= 1, name

    def test_bin_sheet_has_no_images(self, generated):
        gen, _, _, _, _ = generated
        assert len(gen._wb["Bin Chart"]._images) == 0

    # -- dashboard header (seventh wave) ------------------------------------

    def test_bin_banner_and_block_title(self, generated):
        # ninth wave: row 1 = merged A:F "Title"-style banner (fallback
        # "Bin Chart" — this fixture selects no File Info), row 3 = block
        # title "Bin Chart — Head{h} / {site}", 14pt bold theme-primary
        _, wb, _, _, _ = generated
        ws = wb["Bin Chart"]
        assert ws["A1"].value == "Bin Chart"
        assert ws["A1"].style == "Title"
        assert "A1:F1" in {str(r) for r in ws.merged_cells.ranges}
        assert ws["A2"].value is None  # row 2 blank
        cell = ws["A3"]
        assert cell.value == "Bin Chart — Head1 / All Sites"
        assert cell.font.size == 14
        assert cell.font.bold is True
        assert cell.font.color.rgb.endswith("1E40AF")
        assert ws["A4"].value is None  # row 4 blank

    def test_bin_kpi_cards_written(self, generated):
        # R2: three cards on rows 5-7 (label/value/subtitle in A, C, E),
        # aggregated from the one-file HB row list
        _, wb, _, _, _ = generated
        ws = wb["Bin Chart"]
        assert ws["A5"].value == "Yield"
        assert ws["A6"].value == pytest.approx(0.9)  # float, wave-3
        assert isinstance(ws["A6"].value, float)
        assert ws["A6"].number_format == "0.00%"
        assert ws["A7"].value == "9 pass / 10 total parts"
        assert ws["C5"].value == "Fails"
        assert ws["C6"].value == 1
        assert ws["C7"].value == "0 unknown · 0 superseded"
        assert ws["E5"].value == "Total Parts"
        assert ws["E6"].value == 10
        assert ws["E7"].value == "across 1 file(s)"
        assert ws["B8"].value is None  # row 8 blank

    def test_bin_kpi_card_fills_and_merges(self, generated):
        # card fills on the anchor cells (openpyxl's READER recreates
        # merged non-anchor cells without styles, but the styles are in
        # the saved XML); every card row merged across its 2 columns
        _, wb, _, _, _ = generated
        ws = wb["Bin Chart"]
        for anchor, hex6 in (("A5", "FFF3CD"), ("C5", "FFCCCC"), ("E5", "E9EEF6")):
            assert ws[anchor].fill.fgColor.rgb.endswith(hex6), anchor
        ranges = {str(r) for r in ws.merged_cells.ranges}
        for span in (
            "A5:B5",
            "A6:B6",
            "A7:B7",
            "C5:D5",
            "C6:D6",
            "C7:D7",
            "E5:F5",
            "E6:F6",
            "E7:F7",
        ):
            assert span in ranges, span
        # eighth wave (R4) borders on the ninth-wave dark accents: each
        # card carries a medium outline in its own dark accent — post-merge
        # styles survive the save/reload on the anchors
        assert ws["A5"].border.top.style == "medium"
        assert ws["A5"].border.left.color.rgb == "FF9C6500"
        assert ws["C5"].border.top.color.rgb == "FF9C0006"
        assert ws["E5"].border.left.color.rgb == "FF1E40AF"
        assert ws["A7"].border.bottom.style == "medium"

    # -- donuts tiled in their own band (eighth wave) -------------------------

    def test_bin_native_donut_charts(self, generated):
        # charts survive the save/reload round-trip (holeSize / legend
        # position included); anchors: HB donut at A10, SB donut at D10 —
        # below the merged "Charts" headline (row 9); the two 8.5 cm
        # donuts tile the A:F page (0-based row 9, cols 0 / 3)
        _, wb, _, _, _ = generated
        charts = wb["Bin Chart"]._charts
        assert len(charts) == 2
        assert all(isinstance(c, DoughnutChart) for c in charts)
        titles = [chart_title_text(c) for c in charts]
        assert titles == [
            "Hardware Bin — Head1 / All Sites",
            "Software Bin — Head1 / All Sites",
        ]
        assert [(c.anchor._from.row, c.anchor._from.col) for c in charts] == [
            (9, 0),
            (9, 3),
        ]
        # hidden helper columns must still be plotted (survives the
        # round-trip: plotVisOnly val="0")
        assert all(c.visible_cells_only is False for c in charts)
        # ninth-wave print fix: the pinned show* flags survive the reload
        # (Excel renders MISSING flags as TRUE — label garbage over slices)
        assert all(c.dataLabels.showVal is False for c in charts)
        assert all(c.dataLabels.showCatName is False for c in charts)
        assert all(c.dataLabels.showSerName is False for c in charts)
        assert all(c.dataLabels.showLegendKey is False for c in charts)
        assert all(c.dataLabels.showBubbleSize is False for c in charts)
        # mockup: donut hole in the 50-60 band, right-hand legend
        assert all(c.holeSize == 55 for c in charts)
        assert all(c.legend is not None and c.legend.position == "r" for c in charts)

    def test_bin_chart_construction_in_memory(self, generated):
        # in-memory-only attributes (style / size are defaults after a
        # reload): colorful palette 1 (ninth-wave fix — style id 2; id 1 is
        # Excel's gray monochrome scheme), 8.5x7 cm tiles, percent labels
        gen, _, _, _, _ = generated
        for chart in gen._wb["Bin Chart"]._charts:
            assert chart.style == 2
            assert chart.width == 8.5
            assert chart.height == 7
            assert chart.dataLabels.showPercent is True

    def test_bin_charts_use_builtin_palette(self, generated):
        # seventh wave (ruled consequence): no per-slice DataPoint fills —
        # the settings bin-color dicts no longer reach the Bin sheet, the
        # built-in style-6 palette colors the slices
        gen, _, _, _, _ = generated
        for chart in gen._wb["Bin Chart"]._charts:
            assert not chart.series[0].data_points

    def test_bin_helper_block_hidden_columns(self, generated):
        # helper data in fixed far columns BJ/BK, header "Bin"/"Count" +
        # one row per bin; both columns hidden; the two charts of the block
        # use distinct helper row ranges (HB 1..3, SB 5..7)
        _, wb, _, _, _ = generated
        ws = wb["Bin Chart"]
        assert ws["BJ1"].value == "Bin"
        assert ws["BK1"].value == "Count"
        assert ws["BJ2"].value == "Bin 1 - PASS_BIN"
        assert ws["BK2"].value == 9
        assert isinstance(ws["BK2"].value, int)
        assert ws["BJ3"].value == "Bin 3 - FAIL_BIN"
        assert ws["BK3"].value == 1
        assert ws["BJ5"].value == "Bin"
        assert ws["BJ6"].value == "Bin 1 - PASS_BIN"
        assert ws["BJ7"].value == "Bin 3 - FAIL_BIN"
        for letter in ("BJ", "BK"):
            assert ws.column_dimensions[letter].hidden is True

    # -- structured tables after the fixed charts band (eighth wave) ---------

    def test_bin_structured_tables_written(self, generated):
        # ninth-wave layout per block: banner 1, blank 2, title 3, blank 4,
        # KPIs 5-7, blank 8, "Charts" headline 9, charts band 10-23 (14
        # rows = BIN_CHART_ROW_SPAN at 7 cm), blank 24, "Detailed Bin
        # Contribution" headline 25, HB title 26 + header 27 + 2 data rows
        # 28-29 + Total 30, blank 31, SB title 32 + header 33 + data 34-35
        # + Total 36; tables sit in columns B:E (column A belongs to the
        # page grid)
        _, wb, _, _, _ = generated
        ws = wb["Bin Chart"]
        # HB section: Headline-2-styled title band, blue header, two data
        # rows, bold Total
        assert ws["B26"].value == "Hardware Bins — Head1 / All Sites"
        assert ws["B26"].style == "Headline 2"
        assert [ws.cell(row=27, column=c).value for c in range(2, 6)] == [
            "Bin",
            "Value",
            "Count",
            "Distribution",
        ]
        assert [ws[cell].value for cell in ("B28", "C28", "D28")] == [
            "PASS_BIN",
            1,
            9,
        ]
        assert ws["E28"].value == pytest.approx(0.9)
        assert ws["E28"].number_format == "0.00%"
        assert ws["E29"].value == pytest.approx(0.1)
        assert ws["B30"].value == "Total"
        assert ws["D30"].value == 10
        assert isinstance(ws["D30"].value, int)
        assert ws["E30"].value == 1.0
        assert ws["E30"].number_format == "0.00%"
        # SB section after ONE blank row (31): title 32, header 33,
        # data 34-35, Total 36
        assert ws["B32"].value == "Software Bins — Head1 / All Sites"
        assert ws["B32"].style == "Headline 2"
        assert [ws.cell(row=33, column=c).value for c in range(2, 6)] == [
            "Bin",
            "Value",
            "Count",
            "Distribution",
        ]
        assert ws["B36"].value == "Total"
        assert ws["D36"].value == 10

    def test_bin_headlines_charts_and_contribution(self, generated):
        # ninth wave: merged A:F "Headline 1" bands — "Charts" above the
        # donuts (row 9), "Detailed Bin Contribution - HW Bin / SW Bin"
        # above BOTH tables (row 25, NOT replacing the per-table bands)
        _, wb, _, _, _ = generated
        ws = wb["Bin Chart"]
        assert ws["A9"].value == "Charts"
        assert ws["A9"].style == "Headline 1"
        assert ws["A25"].value == "Detailed Bin Contribution - HW Bin / SW Bin"
        assert ws["A25"].style == "Headline 1"
        assert "A9:F9" in {str(r) for r in ws.merged_cells.ranges}
        assert "A25:F25" in {str(r) for r in ws.merged_cells.ranges}

    def test_bin_charts_band_geometry_single_blank_separators(self, generated):
        # the fixed band rows 10..23 hold ONLY the two donuts (no cell
        # values in A:F — helper data lives in BJ/BK); the single blank
        # rows 2, 4, 8, 24, 31 are the ONLY separators between the bands
        _, wb, _, _, _ = generated
        ws = wb["Bin Chart"]
        for row in (2, 4, 8, 24, 31):
            assert all(ws.cell(row=row, column=c).value is None for c in range(1, 7)), (
                row
            )
        for row in range(10, 24):
            assert all(ws.cell(row=row, column=c).value is None for c in range(1, 7)), (
                row
            )
        # content rows sandwich the separators
        for row in (
            1,
            3,
            5,
            6,
            7,
            9,
            25,
            26,
            27,
            28,
            29,
            30,
            32,
            33,
            34,
            35,
            36,
        ):
            assert any(
                ws.cell(row=row, column=c).value is not None for c in range(1, 7)
            ), row

    def test_bin_section_title_band_spans_table(self, generated):
        # the Headline-2 band is merged across the table width B:E (eighth wave)
        _, wb, _, _, _ = generated
        ws = wb["Bin Chart"]
        ranges = {str(r) for r in ws.merged_cells.ranges}
        assert "B26:E26" in ranges
        assert "B32:E32" in ranges

    def test_bin_tables_autofilter_on_both(self, generated):
        # ruled: filter arrows on BOTH tables (mockup shows the SB arrows;
        # one worksheet-level auto_filter cannot cover two ranges, so each
        # table is a minimal Excel Table whose column names equal the header
        # cells verbatim — the wave-2 F5 repair hazard avoided)
        _, wb, _, _, _ = generated
        ws = wb["Bin Chart"]
        assert sorted(ws.tables) == ["HB_Head1_AllSites", "SB_Head1_AllSites"]
        # header..last-data rows at B:E; Total rows (30 / 36) outside
        assert ws.tables["HB_Head1_AllSites"].ref == "B27:E29"
        assert ws.tables["SB_Head1_AllSites"].ref == "B33:E35"
        for name in ("HB_Head1_AllSites", "SB_Head1_AllSites"):
            table = ws.tables[name]
            assert table.autoFilter is not None
            assert table.autoFilter.ref == table.ref
            assert [c.name for c in table.tableColumns] == [
                "Bin",
                "Value",
                "Count",
                "Distribution",
            ]

    def test_bin_tables_data_bars(self, generated):
        # one DataBarRule per table over the Distribution (column E) DATA
        # rows only
        _, wb, _, _, _ = generated
        ws = wb["Bin Chart"]
        for data_range in ("E28:E29", "E34:E35"):
            (rule,) = [
                r
                for cf in ws.conditional_formatting
                for r in cf.rules
                if str(cf.sqref) == data_range
            ]
            assert rule.type == "dataBar"
            assert rule.dataBar.showValue is True
            assert rule.dataBar.color.rgb.endswith("5B9BD5")
            assert [(f.type, f.val) for f in rule.dataBar.cfvo] == [
                ("num", 0),
                ("num", 1),
            ]

    def test_bin_sheet_gridlines_off(self, generated):
        # the brief lists gridlines-off as existing Bin-sheet behavior; the
        # chart sheets bypass render_plain_table (where plain sheets get
        # it), so writeBinChart pins it explicitly
        _, wb, _, _, _ = generated
        assert wb["Bin Chart"].sheet_view.showGridLines is False

    def test_wafer_and_trend_stat_rows_written(self, generated):
        _, wb, _, _, _ = generated
        wafer = {
            c.value
            for row in wb["Wafer Map"].iter_rows()
            for c in row
            if c.value is not None
        }
        trend = {
            c.value
            for row in wb["Trend Chart"].iter_rows()
            for c in row
            if c.value is not None
        }
        assert "WAFER-W01-ROW" in wafer
        assert "TREND-ROW" in trend

    def test_bin_sheet_uniform_page_widths(self, generated):
        # ninth wave: the Bin sheet is a fixed A:F page — all six columns
        # share width 20.4 (KPI grid + two tiled 8.5 cm donuts = 3 columns
        # each at MDW=5 for TNR 8); G+ keep their defaults (no
        # column-dimension entry in the file — openpyxl would answer 13,
        # its default width, on access), the autofit keeps sizing the
        # hidden helper columns
        gen, wb, _, _, _ = generated
        dims = wb["Bin Chart"].column_dimensions
        for letter in ("A", "B", "C", "D", "E", "F"):
            assert dims[letter].width == 20.4, letter
        assert "G" not in wb["Bin Chart"].column_dimensions  # beyond the page
        assert "G" not in gen._wb["Bin Chart"].column_dimensions
        assert dims["BJ"].hidden is True
        assert dims["BK"].hidden is True

    def test_bin_sheet_a4_print_setup(self, generated):
        # eighth wave (R1): the Bin sheet prints on A4, fit to one page
        # wide with unlimited height; other sheets keep their defaults
        # (the DUT Summary sets fit-to-page but no paper size). The class
        # constant PAPERSIZE_A4 is the string "9"; the saved attribute
        # round-trips as int 9 through openpyxl's Integer descriptor.
        _, wb, _, _, _ = generated
        ws = wb["Bin Chart"]
        assert ws.page_setup.paperSize == int(ws.PAPERSIZE_A4)
        assert ws.page_setup.fitToWidth == 1
        assert ws.page_setup.fitToHeight == 0
        assert ws.sheet_properties.pageSetUpPr.fitToPage is True
        assert wb["Trend Chart"].page_setup.paperSize is None
        assert wb["DUT Summary"].page_setup.paperSize is None


# --------------------------------------------------------------------------
# progress / threading contract
# --------------------------------------------------------------------------


class TestProgressContract:
    def test_loop_count_matches_classic_total(self, generated):
        gen, _, _, _, _ = generated
        # DUT 1 + 2 tests, Trend 2 (2 tests x 1 head x 1 site), Bin 1, Wafer 1
        assert gen.totalLoopCnt == 7
        assert gen.loopCount == gen.totalLoopCnt

    def test_close_signal_emitted_no_error(self, generated):
        _, _, _, closed, errors = generated
        assert closed == [True]
        assert errors == []


# --------------------------------------------------------------------------
# die-cell number formats (final review critical 2 + 2026-11-10 R3)
# --------------------------------------------------------------------------


class TestDieCellNumberFormat:
    """Parametric die cells carry the hardcoded scientific format — decoupled
    from the exporter settings (which used to derive a float format via
    number_format_from_settings; the printf-template trap stays documented
    in the design addendum)."""

    def test_die_cells_scientific_hardcoded(self, generated):
        _, wb, _, _, _ = generated
        ws = wb["DUT Summary"]
        assert ws["M16"].number_format == "0.00E+00"
        assert ws["N16"].number_format == "0.00E+00"
        assert ws["M17"].number_format == "0"  # FTR die cells stay integer

    def test_settings_notation_no_longer_affects_die_cells(
        self, tmp_path, qapp, monkeypatch
    ):
        settings = getSetting()
        monkeypatch.setattr(settings.gen, "data_precision", 3)
        monkeypatch.setattr(settings.gen, "data_notation", "F")
        _, wb, _, _, _ = build_generated(tmp_path, qapp)
        ws = wb["DUT Summary"]
        assert number_format_from_settings(3, "F") == "0.000"  # utility intact
        assert ws["M16"].number_format == "0.00E+00"  # hardcode wins
        assert ws["N16"].number_format == "0.00E+00"


# --------------------------------------------------------------------------
# fourth wave: DUT-Summary FTR cache -> Test Statistics pseudo-limits
# --------------------------------------------------------------------------


def cf_block(ws, ref):
    """Rules attached to one conditional-formatting range."""
    for cf in ws.conditional_formatting:
        if str(cf.sqref) == ref:
            return list(cf.rules)
    raise AssertionError(f"no conditional formatting on {ref}")


class TestStatFtrCacheDetection:
    """DUT + Stat selection: writeDUT_Summary runs FIRST and caches the FTR
    test keys (PRIMARY detection signal); the Test Statistics FTR row gets
    the 0 / 0.999999 pseudo-limits, and the Fail Num / Cpk CF letters shift
    right of the Pattern Name column."""

    @pytest.fixture
    def generated(self, tmp_path, qapp):
        return build_generated(tmp_path, qapp, with_stat=True)

    def test_sheet_inventory_with_stat(self, generated):
        _, wb, _, _, _ = generated
        assert wb.sheetnames == [
            "DUT Summary",
            "Test Statistics",
            "Trend Chart",
            "Bin Chart",
            "Wafer Map",
        ]

    def test_ftr_cache_built_from_dut_summary(self, generated):
        gen, _, _, _, _ = generated
        # (test_num, full test name) of every SUB_CODE==FTR testTuple; the
        # PTR test (10, "0_TXCC/VCCC_ASIC") is not cached
        assert gen._ftr_tests == {(12, "2_Func/scanchain")}

    def test_ftr_row_pseudo_limits_via_cache(self, generated):
        _, wb, _, _, _ = generated
        ws = wb["Test Statistics"]
        # FTR-shifted transformed header: SpecMin=I, SpecMax=J. FTR row 3
        # carries the synthetic floats; the PTR row keeps its real limits.
        # (-0.999999 survives the xlsx round-trip as a float — non-integral;
        # 0.999999 likewise.)
        assert ws["I3"].value == -0.999999
        assert isinstance(ws["I3"].value, float)
        assert not isinstance(ws["I3"].value, str)
        assert ws["J3"].value == 0.999999
        assert isinstance(ws["J3"].value, float)
        assert ws["I2"].value == 0.5
        assert ws["J2"].value == 1.5

    def test_fail_num_cpk_cf_shifted_letters(self, generated):
        _, wb, _, _, _ = generated
        ws = wb["Test Statistics"]
        # Fail Num index 10 -> K, Cpk index 11 -> L (Pattern Name shift);
        # data rows only (header row 1 excluded)
        fail_rules = cf_block(ws, "K2:K3")
        assert [r.formula for r in fail_rules] == [
            ["AND(ISNUMBER(K2),K2>0)"],
            ["AND(ISNUMBER(K2),K2=0)"],
        ]
        cpk_rules = cf_block(ws, "L2:L3")
        assert [r.formula[0] for r in cpk_rules] == [
            "AND(ISNUMBER(L2),L2>=1.67)",
            "AND(ISNUMBER(L2),L2>=1.33)",
            "AND(ISNUMBER(L2),L2>=1)",
            "AND(ISNUMBER(L2),L2<1)",
        ]

    def test_loop_count_includes_stat_rows(self, generated):
        gen, _, _, _, _ = generated
        # DUT 3 + Stat 2 + Trend 2 + Bin 1 + Wafer 1
        assert gen.loopCount == gen.totalLoopCnt == 9

    def test_close_signal_emitted_no_error(self, generated):
        _, _, _, closed, errors = generated
        assert closed == [True]
        assert errors == []


# --------------------------------------------------------------------------
# Bin-sheet banner (ninth wave: Product ID / Sublot ID from File Info)
# --------------------------------------------------------------------------


class TestBinBanner:
    def test_banner_from_file_info(self, generated_with_file_info):
        _, wb, _, _, _ = generated_with_file_info
        ws = wb["Bin Chart"]
        assert ws["A1"].value == "ASIC-X / LOT123"
        assert ws["A1"].style == "Title"
        assert "A1:F1" in {str(r) for r in ws.merged_cells.ranges}

    def test_bin_banner_fallback_without_file_info(self, generated):
        # Bin selected WITHOUT File Info -> the report_core default
        # (Review Focus 5: no AttributeError, sheet-name fallback)
        _, wb, _, _, _ = generated
        assert wb["Bin Chart"]["A1"].value == "Bin Chart"
        assert wb["Bin Chart"]["A1"].style == "Title"


class TestBinBannerFromFileInfo:
    def test_single_file(self):
        rows = [["Product ID", "ASIC-X"], ["Sublot ID", "LOT123"]]
        assert bin_banner_from_file_info(rows) == "ASIC-X / LOT123"

    def test_multi_file_joined_in_column_order(self):
        # Review Focus 4: one "product / sublot" pair per file column,
        # files joined " | " in column order
        rows = [["Product ID", "P1", "P2"], ["Sublot ID", "S1", "S2"]]
        assert bin_banner_from_file_info(rows) == "P1 / S1 | P2 / S2"

    def test_banner_missing_anchors_returns_none(self):
        # translated UI: no English labels -> None -> caller falls back
        # to the sheet name (Review Focus 1)
        assert bin_banner_from_file_info([["Produit", "X"], ["Lot", "Y"]]) is None

    def test_banner_partial_values(self):
        # Review Focus 2: a file column missing one field is dropped;
        # all pairs dropped -> None
        rows = [["Product ID", "P1", "P2"], ["Sublot ID", "S1", ""]]
        assert bin_banner_from_file_info(rows) == "P1 / S1"
        assert bin_banner_from_file_info([["Product ID", "P1"]]) is None

    def test_label_matching_normalizes_and_strips_colon(self):
        rows = [["Product ID: ", "P"], ["Sublot ID:", "S"]]
        assert bin_banner_from_file_info(rows) == "P / S"
