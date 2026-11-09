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
from PIL import Image as PILImage
from PyQt5 import QtCore

from deps.SharedSrc import REC, tab
from deps.modernExporter.generator import ModernReportGenerator
from deps.uic_stdExporter import ReportSelection, dataChannel, signals


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
                  tests_executed=42, test_time="0.5", hw_bin=1, sw_bin=2,
                  wafer="", xy=(3, -5), dut_flag="Passed"):
    """dutInfo tuple: [File ID, Part ID, Part Text, Head-Site, Tests Executed,
    Test Time, HW Bin, SW Bin, Wafer ID, (X, Y), DUT Flag] (indices 0..10)."""
    return (fid, part_id, part_txt, head_site, tests_executed, test_time,
            hw_bin, sw_bin, wafer, xy, dut_flag)


def make_summary():
    """Classic waitForDutSummary dict: 2 tests x 3 DUTs in file 1."""
    duts = [
        (1, 0, make_dut_info(part_id="SN541_T25C")),
        (1, 1, make_dut_info(part_id="SN542_T25C", sw_bin=3, xy=(4, -6),
                             dut_flag="Failed")),
        (1, 2, make_dut_info(part_id="SN543_T85C")),
    ]
    tests = [
        # parametric with limits (data row 17)
        dict(name="0_TXCC/VCCC_ASIC", num=10, unit="V", llimit=0.5,
             hlimit=1.5, rec=REC.PTR,
             data={1: ([1.0, 1.2, np.nan], [0, 0, -1])}),
        # FTR (data row 18): die cells from the flag list
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


@pytest.fixture
def generated(tmp_path, qapp):
    """Run report_core() once for {DUT, Trend, Bin, Wafer} and return
    (generator, reloaded workbook, report path, close/error capture)."""
    testTuples = [(10, 0, "0_TXCC/VCCC_ASIC"), (12, 0, "2_Func/scanchain")]
    selectedFiles = [1]
    selectedHeads = [1]
    selectedSites = [-1]
    waferTuples = [(0, 1, "W01")]     # (wafer index, file id, wafer name)
    reportPath = str(tmp_path / "report.xlsx")
    # classic totalLoopCnt formula (stdfExporter.start): Trend 2*1*1
    # + Bin 1*1 + DUT (1 + 2) + Wafer len(waferTuples)
    totalLoopCnt = 2 + 1 + 3 + 1
    reportInfo = ([ReportSelection.DUT, ReportSelection.Trend,
                   ReportSelection.Bin, ReportSelection.Wafer],
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

    def fake_wait_for_data_list(heads, sites, fids, chartType, kargs):
        if chartType == tab.Bin:
            return ([["BIN-HB-ROW", "count", 3]] if kargs.get("isHBIN")
                    else [["BIN-SB-ROW", "count", 3]])
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
        assert ws["A1"].value == "Test"
        assert ws["M1"].value == "SN541_T25C"        # full part ID banner
        assert ws["A1"].fill.fgColor.rgb == "FFCBD5E1"

    def test_repeated_header_row_sixteen(self, generated):
        _, wb, _, _, _ = generated
        ws = wb["DUT Summary"]
        assert ws["A16"].value == "Test"
        assert ws["A16"].fill.fgColor.rgb == "FF1E40AF"
        assert ws["M16"].value == "SN541"            # part ID base

    def test_first_test_formulas_at_row_seventeen(self, generated):
        _, wb, _, _, _ = generated
        ws = wb["DUT Summary"]
        assert ws["G17"].value.startswith("=MIN(M17:")
        assert ws["H17"].value.startswith("=MAX(M17:")
        assert ws["L17"].value.startswith("=SUMPRODUCT(")

    def test_die_values_parametric_and_ftr(self, generated):
        _, wb, _, _, _ = generated
        ws = wb["DUT Summary"]
        assert ws["M17"].value == 1.0
        assert ws["N17"].value == 1.2
        # FTR row: die cells from the flag list (0 pass / 1 fail)
        assert ws["B18"].value == "PAT_TOP"
        assert ws["D18"].value == "P/F"
        assert ws["M18"].value == 0
        assert ws["N18"].value == 1

    def test_mirror_columns_hidden_after_spacer(self, generated):
        _, wb, _, _, _ = generated
        dims = wb["DUT Summary"].column_dimensions
        assert not dims["P"].hidden                   # spacer stays visible
        for letter in ("Q", "R", "S"):
            assert dims[letter].hidden is True

    def test_die_columns_pinned_base_autofit(self, generated):
        _, wb, _, _, _ = generated
        dims = wb["DUT Summary"].column_dimensions
        assert dims["M"].width == 10                  # die columns pinned
        assert dims["A"].width == pytest.approx(len("METADATA") * 1.1)


# --------------------------------------------------------------------------
# chart sheets
# --------------------------------------------------------------------------

class TestChartSheets:
    def test_images_embedded(self, generated):
        gen, _, _, _, _ = generated
        for name in ("Bin Chart", "Wafer Map", "Trend Chart"):
            assert len(gen._wb[name]._images) >= 1, name

    def test_bin_stat_rows_written(self, generated):
        _, wb, _, _, _ = generated
        values = {c.value for row in wb["Bin Chart"].iter_rows()
                  for c in row if c.value is not None}
        assert "BIN-HB-ROW" in values
        assert "BIN-SB-ROW" in values

    def test_wafer_and_trend_stat_rows_written(self, generated):
        _, wb, _, _, _ = generated
        wafer = {c.value for row in wb["Wafer Map"].iter_rows()
                 for c in row if c.value is not None}
        trend = {c.value for row in wb["Trend Chart"].iter_rows()
                 for c in row if c.value is not None}
        assert "WAFER-W01-ROW" in wafer
        assert "TREND-ROW" in trend

    def test_chart_columns_autofit(self, generated):
        _, wb, _, _, _ = generated
        dims = wb["Bin Chart"].column_dimensions
        assert dims["A"].width == pytest.approx(len("BIN-SB-ROW") * 1.1)


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
