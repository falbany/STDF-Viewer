"""Tests for deps/modernExporter/generator.py: DUT Summary + chart sheets.

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
from deps.modernExporter.generator import ModernReportGenerator
from deps.modernExporter.theme import number_format_from_settings
from deps.uic_stdExporter import ReportSelection, dataChannel, signals


def chart_title_text(chart) -> str:
    """Concatenated rich-text runs of a chart title (reload-safe)."""
    return "".join(run.t for p in chart.title.tx.rich.p
                   for run in (p.r or []) if run.t)


def bin_table(bin_type: str) -> list[list]:
    """One classic waitForDataList(tab.Bin) row (single file): VHeader +
    yield/total/count cells + one '"name"\\nBin N: P.P%' cell per non-zero
    bin (DataInterface.py:596-608 flattened by STDF-Viewer.py:1600-1604;
    all strings)."""
    label = "Hardware Bin" if bin_type == "HB" else "Software Bin"
    return [[f"{label} / Head1 / All Sites", "Yield: 90.00%", "Total: 10",
             "Pass: 9", "Failed: 1", "Unknown: 0", "Superseded: 0",
             '"PASS_BIN"\nBin 1: 90.0%', '"FAIL_BIN"\nBin 3: 10.0%']]


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

def make_dut_info(fid=1, part_id="", part_txt="", head_site="1--1",
                  tests_executed=42, test_time="123 ms", hw_bin="Bin 1",
                  sw_bin="Bin 2", wafer="", xy=(3, -5), dut_flag="Passed"):
    """dutInfo tuple: [File ID, Part ID, Part Text, Head-Site, Tests Executed,
    Test Time, HW Bin, SW Bin, Wafer ID, (X, Y), DUT Flag] (indices 0..10).

    Test Time / HW Bin / SW Bin use the raw DB shapes ("123 ms", "Bin 1",
    "Bin 2" — see schema.rs FETCH_SELECT_DUT_SUMMARY)."""
    return (fid, part_id, part_txt, head_site, tests_executed, test_time,
            hw_bin, sw_bin, wafer, xy, dut_flag)


def make_summary():
    """Classic waitForDutSummary dict: 2 tests x 3 DUTs in file 1."""
    duts = [
        (1, 0, make_dut_info(part_id="SN541_T25C")),
        (1, 1, make_dut_info(part_id="SN542_T25C", sw_bin="Bin 3", xy=(4, -6),
                             dut_flag="Failed")),
        (1, 2, make_dut_info(part_id="SN543_T85C")),
    ]
    tests = [
        # parametric with limits (data row 16)
        dict(name="0_TXCC/VCCC_ASIC", num=10, unit="V", llimit=0.5,
             hlimit=1.5, rec=REC.PTR,
             data={1: ([1.0, 1.2, np.nan], [0, 0, -1])}),
        # FTR (data row 17): die cells from the flag list
        dict(name="2_Func/scanchain", num=12, unit="", llimit=np.nan,
             hlimit=np.nan, rec=REC.FTR, vect_nam="PAT_TOP",
             data={1: ([0.0, 1.0, 0.0], [0, 0x80, 0])}),
    ]
    data, test_info, dut2ind, dut_info = {}, {}, {}, {}
    dut2ind[1] = {0: 0, 1: 1, 2: 2}
    dut_info[1] = {d[1]: d[2] for d in duts}
    for t in tests:
        test_tuple = (t["num"], 0, t["name"])
        # TestInfo order: [TEST_NAME, TEST_NUM, HLimit, LLimit, Unit]
        test_info[test_tuple] = [t["name"], t["num"], t["hlimit"],
                                 t["llimit"], t["unit"]]
        values, flags = t["data"][1]
        entry = {"SUB_CODE": t["rec"],
                 "dataList": np.asarray(values, dtype=float),
                 "flagList": np.asarray(flags, dtype=int)}
        if t.get("vect_nam") is not None:
            entry["VECT_NAM"] = t["vect_nam"]
        data[test_tuple] = {1: entry}
    return {"Data": data, "TestInfo": test_info, "dut2ind": dut2ind,
            "dutInfo": dut_info}


def png_bytes(width=96, height=48):
    """A small real PNG (96x48 scales to 1152x576 px at 12 in width)."""
    buf = io.BytesIO()
    PILImage.new("RGB", (width, height), color=(0, 0, 0)).save(buf,
                                                              format="PNG")
    return buf.getvalue()


PNG_BYTES = png_bytes()


def build_generated(tmp_path, qapp, with_stat=False):
    """Run report_core() once for {DUT, Trend, Bin, Wafer} (plus Stat when
    `with_stat`) and return (generator, reloaded workbook, report path,
    close/error capture)."""
    testTuples = [(10, 0, "0_TXCC/VCCC_ASIC"), (12, 0, "2_Func/scanchain")]
    selectedFiles = [1]
    selectedHeads = [1]
    selectedSites = [-1]
    waferTuples = [(0, 1, "W01")]     # (wafer index, file id, wafer name)
    reportPath = str(tmp_path / "report.xlsx")
    # classic totalLoopCnt formula (stdfExporter.start): Trend 2*1*1
    # + Bin 1*1 + DUT (1 + 2) + Wafer len(waferTuples)
    totalLoopCnt = 2 + 1 + 3 + 1
    selection = [ReportSelection.DUT, ReportSelection.Trend,
                 ReportSelection.Bin, ReportSelection.Wafer]
    if with_stat:
        # classic sheet-creation order: DUT < Stat < Trend
        selection.insert(1, ReportSelection.Stat)
        totalLoopCnt += 2   # Stat: +1 per data row
    reportInfo = (selection,
                  testTuples, selectedFiles, selectedHeads, selectedSites,
                  reportPath, waferTuples, totalLoopCnt)

    sig = signals()
    channel = dataChannel()
    gen = ModernReportGenerator(sig, QtCore.QMutex(),
                                QtCore.QWaitCondition(), reportInfo, channel)

    # GUI side: classic-protocol summary table, chart PNGs, stat rows
    gen.waitForDutSummary = (lambda heads, sites, fids, testTuples:
                             make_summary())

    def fake_wait_for_image(testTuple, head, sites, fids, chartType):
        return io.BytesIO(PNG_BYTES)
    gen.waitForImage = fake_wait_for_image

    # classic DataInterface.getStatistics shape (containsFTR=True: the
    # "Pattern Name" column follows "Test Name"; VECT_NAM only on the FTR row)
    stat_table = [
        ["", "Test Name", "Pattern Name", "Unit", "Low Limit", "High Limit",
         "Fail Num", "Cpk", "Average", "Median", "St. Dev.", "Min", "Max"],
        ["10 / Head1 / All Sites", "0_TXCC/VCCC_ASIC", "", "V", "0.500",
         "1.500", "1", "1.23", "1.000", "1.000", "0.010", "0.980", "1.020"],
        ["12 / Head1 / All Sites", "2_Func/scanchain", "PAT_TOP", "P/F",
         "N/A", "N/A", "0", "∞", "1.000", "1.000", "0.000", "0.000", "1.000"],
    ]

    def fake_wait_for_data_list(heads, sites, fids, chartType, kargs):
        if (chartType == tab.Trend
                and kargs.get("testTuples") == testTuples):
            # writeStatistic passes the FULL testTuples list (per-test trend
            # rows pass a single-tuple list — never equal here, two tests)
            return stat_table
        if chartType == tab.Bin:
            return (bin_table("HB") if kargs.get("isHBIN")
                    else bin_table("SB"))
        if chartType == tab.Wafer:
            return [["WAFER-W01-ROW", "Total DUTs", 3]]
        return [["TREND-ROW", "mean", 1.1]]
    gen.waitForDataList = fake_wait_for_data_list

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


# --------------------------------------------------------------------------
# sheet inventory
# --------------------------------------------------------------------------

class TestSheetInventory:
    def test_sheets_classic_order_without_histogram(self, generated):
        _, wb, _, _, _ = generated
        assert wb.sheetnames == ["DUT Summary", "Trend Chart",
                                 "Bin Chart", "Wafer Map"]

    def test_unselected_and_overflow_sheets_absent(self, generated):
        _, wb, _, _, _ = generated
        assert "Histogram" not in wb.sheetnames      # not selected
        assert "DUT Summary (2)" not in wb.sheetnames  # 3 DUTs, single chunk


# --------------------------------------------------------------------------
# DUT Summary matrix sheet
# --------------------------------------------------------------------------

class TestDutSummarySheet:
    def test_banner_row_one(self, generated):
        _, wb, _, _, _ = generated
        ws = wb["DUT Summary"]
        assert ws["A1"].value == "Test#"              # Test# leads the base labels
        assert ws["M1"].value == 1                    # banner labels are plain ints
        assert ws["A1"].fill.fgColor.rgb == "FFCBD5E1"

    def test_repeated_header_row_fifteen(self, generated):
        _, wb, _, _, _ = generated
        ws = wb["DUT Summary"]
        assert ws["A15"].value == "Test#"
        assert ws["A15"].fill.fgColor.rgb == "FF1E40AF"
        assert ws["M15"].value == "SN541"             # part ID base

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
        assert not dims["P"].hidden                   # spacer stays visible
        for letter in ("Q", "R", "S"):
            assert dims[letter].hidden is True

    def test_fixed_column_widths_no_autofit(self, generated):
        # 2026-11-10 R9: DUT-sheet widths are fixed by the renderer
        _, wb, _, _, _ = generated
        dims = wb["DUT Summary"].column_dimensions
        assert dims["A"].width == 8                   # Test#
        assert dims["B"].width == 30                  # Test name
        assert dims["C"].width == 8                   # Param
        assert dims["M"].width == 8                   # die columns pinned


# --------------------------------------------------------------------------
# chart sheets
# --------------------------------------------------------------------------

class TestChartSheets:
    @pytest.fixture(autouse=True)
    def seed_bin_colors(self):
        """The global bin-color dicts are empty until a STDF loads
        (setSettingDefaultColor); seed both so the pies carry per-slice
        DataPoint colors, restored afterwards."""
        settings = getSetting()
        seeded = {1: "#00CC00", 3: "#CC0000"}
        settings.color.hbin_colors.update(seeded)
        settings.color.sbin_colors.update(seeded)
        yield
        for key in seeded:
            settings.color.hbin_colors.pop(key, None)
            settings.color.sbin_colors.pop(key, None)

    def test_images_embedded(self, generated):
        # fifth wave: Wafer/Trend keep the embedded PNGs — Bin no longer
        # calls waitForImage at all
        gen, _, _, _, _ = generated
        for name in ("Wafer Map", "Trend Chart"):
            assert len(gen._wb[name]._images) >= 1, name

    def test_bin_sheet_has_no_images(self, generated):
        gen, _, _, _, _ = generated
        assert len(gen._wb["Bin Chart"]._images) == 0

    def test_bin_native_donut_charts(self, generated):
        # sixth wave: the fifth-wave pies become doughnuts — charts survive
        # the save/reload round-trip (holeSize / legend position included)
        _, wb, _, _, _ = generated
        charts = wb["Bin Chart"]._charts
        assert len(charts) == 2
        assert all(isinstance(c, DoughnutChart) for c in charts)
        titles = [chart_title_text(c) for c in charts]
        assert titles == ["Hardware Bin — Head1 / All Sites",
                          "Software Bin — Head1 / All Sites"]
        # side-by-side anchors: HB at A1 (row 0 / col 0), SB at H1 (col 7)
        assert [(c.anchor._from.row, c.anchor._from.col) for c in charts] == [
            (0, 0), (0, 7)]
        # hidden helper columns must still be plotted (survives the
        # round-trip: plotVisOnly val="0")
        assert all(c.visible_cells_only is False for c in charts)
        # mockup: donut hole in the 50-60 band, right-hand legend
        assert all(c.holeSize == 55 for c in charts)
        assert all(c.legend is not None and c.legend.position == "r"
                   for c in charts)

    def test_bin_chart_construction_in_memory(self, generated):
        # in-memory-only attributes (style / size are defaults after a
        # reload): built-in style 10, 9x7 cm, percent data labels
        gen, _, _, _, _ = generated
        for chart in gen._wb["Bin Chart"]._charts:
            assert chart.style == 10
            assert chart.width == 9
            assert chart.height == 7
            assert chart.dataLabels.showPercent is True

    def test_bin_slice_colors_from_settings(self, generated):
        # per-slice DataPoint colors from the settings bin-color dicts,
        # 6-digit hex without the leading "#"
        gen, _, _, _, _ = generated
        settings = getSetting()
        for chart, color_dict in (
                (gen._wb["Bin Chart"]._charts[0], settings.color.hbin_colors),
                (gen._wb["Bin Chart"]._charts[1], settings.color.sbin_colors)):
            points = chart.series[0].data_points
            assert [p.idx for p in points] == [0, 1]
            assert [p.graphicalProperties.solidFill.srgbClr for p in points] \
                == [color_dict[1].lstrip("#")[:6],
                    color_dict[3].lstrip("#")[:6]]

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

    def test_bin_structured_tables_written(self, generated):
        # sixth wave: structured HB/SB tables replace the classic text rows
        # (first block: charts consume 15 rows = ceil(7cm/0.21in) + 1 gap;
        # HB section rows 15-19, one-row gap, SB section rows 21-25)
        _, wb, _, _, _ = generated
        ws = wb["Bin Chart"]
        # HB section: title, blue header, two data rows, bold Total
        assert ws["A15"].value == "Hardware Bins — Head1 / All Sites"
        assert [ws.cell(row=16, column=c).value for c in range(1, 5)] == [
            "Bin", "Value", "Count", "Distribution"]
        assert [ws[cell].value for cell in ("A17", "B17", "C17")] == [
            "PASS_BIN", 1, 9]
        assert ws["D17"].value == pytest.approx(0.9)
        assert ws["D17"].number_format == "0.00%"
        assert ws["D18"].value == pytest.approx(0.1)
        assert ws["A19"].value == "Total"
        assert ws["C19"].value == 10
        assert isinstance(ws["C19"].value, int)
        assert ws["D19"].value == 1.0
        assert ws["D19"].number_format == "0.00%"
        # the classic raw upstream rows are gone (old B15 "Yield: 90.00%")
        assert ws["B15"].value is None
        # SB section: title at 21 (19 total + 1-row gap -> title 21)
        assert ws["A21"].value == "Software Bins — Head1 / All Sites"
        assert [ws.cell(row=22, column=c).value for c in range(1, 5)] == [
            "Bin", "Value", "Count", "Distribution"]
        assert ws["A25"].value == "Total"
        assert ws["C25"].value == 10

    def test_bin_tables_autofilter_on_both(self, generated):
        # ruled: filter arrows on BOTH tables (mockup shows the SB arrows;
        # one worksheet-level auto_filter cannot cover two ranges, so each
        # table is a minimal Excel Table whose column names equal the header
        # cells verbatim — the wave-2 F5 repair hazard avoided)
        _, wb, _, _, _ = generated
        ws = wb["Bin Chart"]
        assert sorted(ws.tables) == ["HB_Head1_AllSites", "SB_Head1_AllSites"]
        # header..last-data rows; Total rows (19 / 25) outside the ranges
        assert ws.tables["HB_Head1_AllSites"].ref == "A16:D18"
        assert ws.tables["SB_Head1_AllSites"].ref == "A22:D24"
        for name in ("HB_Head1_AllSites", "SB_Head1_AllSites"):
            table = ws.tables[name]
            assert table.autoFilter is not None
            assert table.autoFilter.ref == table.ref
            assert [c.name for c in table.tableColumns] == [
                "Bin", "Value", "Count", "Distribution"]

    def test_bin_tables_data_bars(self, generated):
        # one DataBarRule per table over the Distribution DATA rows only
        _, wb, _, _, _ = generated
        ws = wb["Bin Chart"]
        for data_range in ("D17:D18", "D23:D24"):
            (rule,) = [r for cf in ws.conditional_formatting
                       for r in cf.rules if str(cf.sqref) == data_range]
            assert rule.type == "dataBar"
            assert rule.dataBar.showValue is True
            assert rule.dataBar.color.rgb.endswith("5B9BD5")
            assert [(f.type, f.val) for f in rule.dataBar.cfvo] == [
                ("num", 0), ("num", 1)]

    def test_bin_sheet_gridlines_off(self, generated):
        # the brief lists gridlines-off as existing Bin-sheet behavior; the
        # chart sheets bypass render_plain_table (where plain sheets get
        # it), so writeBinChart pins it explicitly
        _, wb, _, _, _ = generated
        assert wb["Bin Chart"].sheet_view.showGridLines is False

    def test_wafer_and_trend_stat_rows_written(self, generated):
        _, wb, _, _, _ = generated
        wafer = {c.value for row in wb["Wafer Map"].iter_rows()
                 for c in row if c.value is not None}
        trend = {c.value for row in wb["Trend Chart"].iter_rows()
                 for c in row if c.value is not None}
        assert "WAFER-W01-ROW" in wafer
        assert "TREND-ROW" in trend

    def test_bin_table_column_widths(self, generated):
        # sixth wave: the four table columns get fixed widths (pinned AFTER
        # the autofit, which keeps sizing the hidden helper columns)
        _, wb, _, _, _ = generated
        dims = wb["Bin Chart"].column_dimensions
        assert dims["A"].width == 24
        assert dims["B"].width == 8
        assert dims["C"].width == 8
        assert dims["D"].width == 14


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
        assert ws["M17"].number_format == "0"   # FTR die cells stay integer

    def test_settings_notation_no_longer_affects_die_cells(self, tmp_path, qapp, monkeypatch):
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
        assert wb.sheetnames == ["DUT Summary", "Test Statistics",
                                 "Trend Chart", "Bin Chart", "Wafer Map"]

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
            ["AND(ISNUMBER(K2),K2>0)"], ["AND(ISNUMBER(K2),K2=0)"]]
        cpk_rules = cf_block(ws, "L2:L3")
        assert [r.formula[0] for r in cpk_rules] == [
            "AND(ISNUMBER(L2),L2>=1.67)", "AND(ISNUMBER(L2),L2>=1.33)",
            "AND(ISNUMBER(L2),L2>=1)", "AND(ISNUMBER(L2),L2<1)"]

    def test_loop_count_includes_stat_rows(self, generated):
        gen, _, _, _, _ = generated
        # DUT 3 + Stat 2 + Trend 2 + Bin 1 + Wafer 1
        assert gen.loopCount == gen.totalLoopCnt == 9

    def test_close_signal_emitted_no_error(self, generated):
        _, _, _, closed, errors = generated
        assert closed == [True]
        assert errors == []
