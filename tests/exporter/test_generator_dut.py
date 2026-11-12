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
import zipfile
from pathlib import Path

import numpy as np
import pytest
from openpyxl import load_workbook
from PIL import Image as PILImage
from PyQt5 import QtCore

from deps.SharedSrc import REC, getSetting, tab
from deps.exporter.charts import BIN_TABLE_HEADERS
from deps.exporter.generator import ModernReportGenerator
from deps.exporter.worksheets.file_info import bin_banner_from_file_info
from deps.exporter.theme import number_format_from_settings
from deps.uic_stdExporter import ReportSelection, dataChannel, signals
from tests.exporter.test_writer import col_dim

# openpyxl's reader drops features it cannot parse (xlsxwriter's sparkline
# x14 extensions, solid data-bar extensions) — the suite asserts those via
# the raw package XML instead; silence the reader-side noise here.
pytestmark = [
    pytest.mark.filterwarnings("ignore:Sparkline Group extension"),
    pytest.mark.filterwarnings("ignore:Conditional Formatting extension"),
    pytest.mark.filterwarnings("ignore:Unknown extension"),
]


def chart_title_text(xml: str) -> str:
    """The title text of a chart XML document (first <a:t> run)."""
    import re

    match = re.search(r"<a:t>([^<]*)</a:t>", xml)
    return match.group(1) if match else ""


def zip_charts_anchor_cols(drawing_xml: str):
    """(from_col, from_row) of every chart anchor in a drawing XML."""
    import re

    out = []
    for m in re.finditer(
        r"<xdr:from><xdr:col>(\d+)</xdr:col><xdr:colOff>\d+</xdr:colOff>"
        r"<xdr:row>(\d+)</xdr:row>",
        drawing_xml,
    ):
        out.append((int(m.group(1)), int(m.group(2))))
    return out


def ws_row_height(generated, sheet, row_1based) -> float:
    _, wb, _, _, _ = generated
    dim = wb[sheet].row_dimensions.get(row_1based)
    return dim.height if dim is not None else None


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


def make_summary(tagged=True, duplicate_sn=False, blank_ids=False, part_ids=None):
    """Classic waitForDutSummary dict: 2 tests x 3 DUTs in file 1.

    `tagged` controls the _T<n>C temperature tag on the part IDs — the
    untagged variant exercises the MIR Test Temperature fallback.
    `duplicate_sn` re-uses SN541 on the third DUT (dutIndex 2) with the
    SAME temperature tag as the first — identical raw Part IDs, so the
    filter's exact-match rule drops the FIRST occurrence (1, 0).
    `blank_ids` empties every Part ID (the filter must never match).
    `part_ids` overrides the whole per-DUT Part ID list verbatim (its
    length IS the DUT count; data arrays pad with pass/NaN) — the
    temperature-split tests build mixed tag/untagged rosters with it."""

    def part_id(base, temp):
        if blank_ids:
            return ""
        return f"{base}_T{temp}C" if tagged else base

    if part_ids is not None:
        ids = part_ids
        flagged = ["Passed", "Failed", "Passed"] + ["Passed"] * (len(ids) - 3)
        duts = [(1, i, make_dut_info(part_id=pid, dut_flag=flag)) for i, (pid, flag) in enumerate(zip(ids, flagged))]
    else:
        duts = [
            (1, 0, make_dut_info(part_id=part_id("SN541", 25))),
            (
                1,
                1,
                make_dut_info(
                    part_id=part_id("SN542", 25),
                    sw_bin="Bin 3",
                    xy=(4, -6),
                    dut_flag="Failed",
                ),
            ),
            # duplicate_sn: dutIndex 2 carries dutIndex 0's EXACT part ID
            # (base AND tag) — the filter matches raw strings, not SNs
            (
                1,
                2,
                make_dut_info(
                    part_id=part_id("SN541" if duplicate_sn else "SN543", 25 if duplicate_sn else 85)
                ),
            ),
        ]
    n = len(duts)

    def pad(values, flags):
        values = list(values) + [np.nan] * (n - len(values))
        flags = list(flags) + [0] * (n - len(flags))
        return values, flags

    ptr_values, ptr_flags = pad([1.0, 1.2, np.nan], [0, 0, -1])
    ftr_values, ftr_flags = pad([0.0, 1.0, 0.0], [0, 0x80, 0])
    tests = [
        # parametric with limits (data row 16)
        dict(
            name="0_TXCC/VCCC_ASIC",
            num=10,
            unit="V",
            llimit=0.5,
            hlimit=1.5,
            rec=REC.PTR,
            data={1: (ptr_values, ptr_flags)},
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
            data={1: (ftr_values, ftr_flags)},
        ),
    ]
    data, test_info, dut2ind, dut_info = {}, {}, {}, {}
    dut2ind[1] = {dut_index: dut_index for dut_index in range(n)}
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

# Flagged tab.Wafer payload (issue #2 cell-grid bin map): the GUI branch
# returns {"rows", "map"} with palette hexes merged into the map dict.
# 3x3 bounds; dies: (-1,0)->1, (0,0)->1, (0,-1)->2.
WAFER_MAP_PAYLOAD = {
    "Bounds": (1, -1, 1, -1),  # xmax, xmin, ymax, ymin
    "Info": (1.0, "", False, False, "W01", [-1]),
    "Data": {
        1: {"x": np.array([-1, 0]), "y": np.array([0, 0])},
        2: {"x": np.array([0]), "y": np.array([-1])},
    },
    "Statistic": {1: ("Good", 2, 66.7), 2: ("Fail", 1, 33.3)},
    "Palette": {1: "#00CC00", 2: "#CC0000"},
    "BinPF": {1: "P", 2: "F"},
    "Properties": [
        ("Wafer Center:", "(10, 20)"),
        ("Wafer Flat Direction:", "Down"),
        ("Wafer XY Direction:", "(Up, Right)"),
    ],
}


def build_generated(
    tmp_path,
    qapp,
    with_stat=False,
    with_file_info=False,
    part_id_tagged=True,
    file_temperature=None,
    duplicate_sn=False,
    blank_ids=False,
    part_ids=None,
    split_temperature=False,
    hide_remeasured=False,
    with_dut=True,
    wafer_tuples=None,
    monkeypatch=None,
):
    """Run report_core() once for {DUT, Trend, Bin, Wafer} (plus Stat when
    `with_stat`, FileInfo when `with_file_info`) and return (generator,
    reloaded workbook, report path, close/error capture). `file_temperature`
    appends a File Info "Test Temperature:" row (the DUT Summary fallback
    source); `part_id_tagged` drops the _T<n>C tag from the part IDs.
    `duplicate_sn` shares SN541 between the first and third DUT;
    `hide_remeasured` flips the filter setting for the run (pass the test's
    `monkeypatch` fixture so it auto-restores). `wafer_tuples` overrides the
    wafer selection (e.g. [(-1, -1, "STACK")] for the stacked-wafer path)."""
    testTuples = [(10, 0, "0_TXCC/VCCC_ASIC"), (12, 0, "2_Func/scanchain")]
    selectedFiles = [1]
    selectedHeads = [1]
    selectedSites = [-1]
    waferTuples = [(0, 1, "W01")] if wafer_tuples is None else wafer_tuples
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
    if not with_dut:
        # stat-only selection: the DUT sheet (its info tick + 2 test
        # columns) never renders — the filter's stat-only fallback fetch
        # path (§7.2) is what still honors the drop-set
        selection.remove(ReportSelection.DUT)
        totalLoopCnt -= 3
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
    gen.waitForDutSummary = lambda heads, sites, fids, testTuples: make_summary(
        tagged=part_id_tagged, duplicate_sn=duplicate_sn, blank_ids=blank_ids, part_ids=part_ids
    )
    # the filter path must never reach the GUI for dutInfo when the DUT
    # sheet rendered; the stat-only selection stubs the light fetch with
    # the partial dutInfo shape (getPartialDUTInfoOnCondition rows)
    gen.waitForDutInfo = lambda *args: (_ for _ in ()).throw(
        AssertionError("waitForDutInfo must not be called when the DUT sheet ran")
    )
    gen._captured_stat_kargs = []
    if not with_dut:
        gen.waitForDutInfo = lambda heads, sites, fids: {
            1: {
                0: ("SN541", "", "Head 1 - Site 1", "Failed - 0x08"),
                2: ("SN541", "", "Head 1 - Site 1", "Passed - 0x00"),
            }
        }
    if hide_remeasured:
        assert monkeypatch is not None, "hide_remeasured=True requires monkeypatch"
        monkeypatch.setattr(getSetting().gen, "hide_remeasured", True)
    if split_temperature:
        assert monkeypatch is not None, "split_temperature=True requires monkeypatch"
        monkeypatch.setattr(getSetting().gen, "split_temperature", True)

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
            if kargs.get("waferMapData") is not None:
                # flagged cell-map fetch (issue #2): bundled payload
                return {
                    "rows": [["Yield: 60.00%", "Total: 5", "WAFER_ID: W01"]],
                    "map": dict(WAFER_MAP_PAYLOAD),
                }
            # unflagged classic shape (legacy exporter + stacked wafer)
            return [["WAFER-W01-ROW", "Total DUTs", 3]]
        return [["TREND-ROW", "mean", 1.1]]

    gen.waitForDataList = fake_wait_for_data_list

    real_fake = gen.waitForDataList

    def capturing_wait_for_data_list(heads, sites, fids, chartType, kargs):
        gen._captured_stat_kargs.append((chartType, dict(kargs)))
        return real_fake(heads, sites, fids, chartType, kargs)

    gen.waitForDataList = capturing_wait_for_data_list

    # writeDUT_Summary fetches the File Info table for the temperature
    # fallback regardless of the sheet selection — the stub answers from
    # the fixture's table (empty when the sheet is not selected, plus the
    # "Test Temperature:" row when `file_temperature` is set)
    file_info_table = []
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
    if file_temperature is not None:
        file_info_table = file_info_table + [["Test Temperature: ", file_temperature]]
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
        assert ws["G13"].value == '=IFERROR(MIN(M13:O13),"")'
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
        assert ws["G16"].value.startswith("=IFERROR(MIN(M16:")
        assert ws["H16"].value.startswith("=IFERROR(MAX(M16:")
        assert ws["L16"].value.startswith("=IFERROR(SUMPRODUCT(")

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
        # xlsxwriter migration: spacer P, Trend Q (visible), spacer R, then
        # the hidden flag mirrors S..U
        _, wb, _, _, _ = generated
        ws = wb["DUT Summary"]
        for letter in ("P", "Q", "R"):
            dim = col_dim(ws, letter)
            assert dim is None or not dim.hidden
        for letter in ("S", "T", "U"):
            assert col_dim(ws, letter).hidden is True

    def test_fixed_column_widths_no_autofit(self, generated):
        # 2026-11-10 R9: DUT-sheet widths are fixed by the renderer
        # (xlsxwriter stores width + ~0.71 — max-digit-width compensation)
        _, wb, _, _, _ = generated
        ws = wb["DUT Summary"]
        assert col_dim(ws, "A").width == pytest.approx(8, abs=0.75)  # Test#
        assert col_dim(ws, "B").width == pytest.approx(30, abs=0.75)  # Test name
        assert col_dim(ws, "C").width == pytest.approx(8, abs=0.75)  # Param
        assert col_dim(ws, "M").width == pytest.approx(8, abs=0.75)  # dies


# --------------------------------------------------------------------------
# Temperature metadata row: Part ID tag vs recorded MIR fallback
# --------------------------------------------------------------------------


class TestDutTemperatureFallback:
    """writeDUT_Summary wires the File Info "Test Temperature:" row into the
    Temperature metadata row's fallback (untagged part IDs)."""

    @pytest.fixture
    def generated_fallback(self, tmp_path, qapp):
        # part IDs carry no tag; the file records 85C
        return build_generated(
            tmp_path, qapp, part_id_tagged=False, file_temperature="85C"
        )

    @pytest.fixture
    def generated_tagged(self, tmp_path, qapp):
        # part IDs keep their tags even though the file records one too
        return build_generated(tmp_path, qapp, file_temperature="85C")

    @pytest.fixture
    def generated_no_source(self, tmp_path, qapp):
        # no tag and no recorded temperature: blank cells
        return build_generated(tmp_path, qapp, part_id_tagged=False)

    def test_untagged_part_ids_take_file_temperature(self, generated_fallback):
        _, wb, _, _, _ = generated_fallback
        ws = wb["DUT Summary"]
        assert ws["M5"].value == 85.0  # Temperature row (5): die cells M..O
        assert ws["N5"].value == 85.0
        assert ws["O5"].value == 85.0

    def test_tagged_part_ids_keep_encoded_temperature(self, generated_tagged):
        _, wb, _, _, _ = generated_tagged
        ws = wb["DUT Summary"]
        assert ws["M5"].value == 25.0  # part-ID tag wins over the file value
        assert ws["N5"].value == 25.0
        assert ws["O5"].value == 85.0

    def test_no_tag_no_file_row_stays_blank(self, generated_no_source):
        _, wb, _, _, _ = generated_no_source
        ws = wb["DUT Summary"]
        assert ws["M5"].value in ("", None)
        assert ws["O5"].value in ("", None)

    def test_fallback_temperatures_fold_like_tags(self, generated_fallback):
        # constant 85C across all DUTs -> the Temperature row hides
        _, wb, _, _, _ = generated_fallback
        assert wb["DUT Summary"].row_dimensions[5].hidden is True


# --------------------------------------------------------------------------
# chart sheets
# --------------------------------------------------------------------------


class TestChartSheets:
    def test_images_embedded(self, generated):
        # issue #2: the Wafer sheet renders the cell-grid bin map — the
        # per-wafer PNG is gone; only Trend keeps embedded PNGs. xlsxwriter:
        # images live in xl/media/ (the fixture sends the SAME PNG —
        # xlsxwriter deduplicates the bytes to one media entry, referenced
        # by two <xdr:pic> anchors across the drawings)
        _, _, report_path, _, _ = generated
        with zipfile.ZipFile(report_path) as zf:
            media = [n for n in zf.namelist() if n.startswith("xl/media/")]
            pics = sum(
                zf.read(n).count(b"<xdr:pic>")
                for n in zf.namelist()
                if n.startswith("xl/drawings/drawing") and n.endswith(".xml")
            )
        assert len(media) >= 1
        assert pics == 2  # 2 trend anchors (wafer PNG replaced by the grid)

    def test_bin_sheet_has_no_images(self, generated):
        # the package's drawings carry exactly the two trend image anchors
        # and the two donut graphicFrames — Bin contributes charts only
        _, _, report_path, _, _ = generated
        with zipfile.ZipFile(report_path) as zf:
            pics, frames = 0, 0
            for n in zf.namelist():
                if n.startswith("xl/drawings/drawing") and n.endswith(".xml"):
                    data = zf.read(n)
                    pics += data.count(b"<xdr:pic>")
                    frames += data.count(b"<xdr:graphicFrame")  # tag carries attributes
        assert pics == 2
        assert frames == 2

    # -- dashboard header (seventh wave) ------------------------------------

    def test_bin_banner_row_one(self, generated):
        # ninth wave: row 1 = merged A:F "Title"-style banner (fallback
        # "Bin Chart" — this fixture selects no File Info); tenth wave:
        # wrapped in a 50 pt row so a 2-line title fits, the block title
        # row is GONE and the KPI cards start after one blank row
        _, wb, _, _, _ = generated
        ws = wb["Bin Chart"]
        assert ws["A1"].value == "Bin Chart"
        # Title LOOKALIKE (xlsxwriter has no named styles): Cambria 18 bold
        assert ws["A1"].font.name == "Cambria"
        assert ws["A1"].font.size == 18
        assert ws["A1"].font.bold is True
        assert "A1:F1" in {str(r) for r in ws.merged_cells.ranges}
        assert ws.row_dimensions[1].height == pytest.approx(50, abs=0.3)
        assert ws["A1"].alignment.wrap_text is True
        assert ws["A1"].alignment.vertical == "center"  # baseline parity
        assert ws["A2"].value is None  # row 2 blank separator
        assert ws["A3"].value == "Yield"  # KPI cards directly below

    def test_bin_kpi_cards_written(self, generated):
        # tenth wave: three cards on rows 3-5 (label/value/subtitle in A,
        # C, E), aggregated from the one-file HB row list
        _, wb, _, _, _ = generated
        ws = wb["Bin Chart"]
        assert ws["A3"].value == "Yield"
        assert ws["A4"].value == pytest.approx(0.9)  # float, wave-3
        assert isinstance(ws["A4"].value, float)
        assert ws["A4"].number_format == "0.00%"
        assert ws["A5"].value == "9 pass / 10 total parts"
        assert ws["C3"].value == "Fails"
        assert ws["C4"].value == 1
        assert ws["C5"].value == "0 unknown · 0 superseded"
        assert ws["E3"].value == "Total Parts"
        assert ws["E4"].value == 10
        assert ws["E5"].value == "across 1 file(s)"
        assert ws["B6"].value is None  # row 6 blank

    def test_bin_kpi_card_fills_and_merges(self, generated):
        # card fills on the anchor cells (openpyxl's READER recreates
        # merged non-anchor cells without styles, but the styles are in
        # the saved XML); every card row merged across its 2 columns
        _, wb, _, _, _ = generated
        ws = wb["Bin Chart"]
        for anchor, hex6 in (("A3", "FFF3CD"), ("C3", "FFCCCC"), ("E3", "E9EEF6")):
            assert ws[anchor].fill.fgColor.rgb.endswith(hex6), anchor
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
        # eighth wave (R4) borders on the ninth-wave dark accents: each
        # card carries a medium outline in its own dark accent — post-merge
        # styles survive the save/reload on the anchors
        assert ws["A3"].border.top.style == "medium"
        assert ws["A3"].border.left.color.rgb == "FF9C6500"
        assert ws["C3"].border.top.color.rgb == "FF9C0006"
        assert ws["E3"].border.left.color.rgb == "FF1E40AF"
        assert ws["A5"].border.bottom.style == "medium"

    # -- donuts tiled in their own band (eighth wave) -------------------------

    def test_bin_native_donut_charts(self, generated):
        # chart internals via the saved package XML (openpyxl cannot read
        # charts): two doughnuts titled HB/SB, hole 55, right legend,
        # plotVisOnly=0 (completion), pinned show* zeros, anchored A9/D9
        _, _, report_path, _, _ = generated
        with zipfile.ZipFile(report_path) as zf:
            charts = [
                zf.read(n).decode()
                for n in sorted(zf.namelist())
                if n.startswith("xl/charts/chart") and n.endswith(".xml")
            ]
        assert len(charts) == 2
        titles = sorted(chart_title_text(xml) for xml in charts)
        assert titles == [
            "Hardware Bin — Head1 / All Sites",
            "Software Bin — Head1 / All Sites",
        ]
        for xml in charts:
            assert "<c:doughnutChart>" in xml
            assert '<c:holeSize val="55"/>' in xml
            assert '<c:legendPos val="r"/>' in xml
            # no hidden data anymore: plotVisOnly stays xlsxwriter's
            # default (visible cells only — the tables are visible)
            assert '<c:plotVisOnly val="0"/>' not in xml
            assert '<c:showVal val="0"/>' in xml  # completion pinned the flags
            assert '<c:showPercent val="1"/>' in xml
            # 2026-11-11 modern palette: per-slice dPt colors (wave-7
            # no-per-slice ruling reversed)
            assert "<c:dPt>" in xml
            assert "<c:style" not in xml  # style 2 = xlsxwriter default
        with zipfile.ZipFile(report_path) as zf:
            # the chart anchors live in the Bin sheet's drawing (drawing1
            # belongs to Trend) — scan every drawing with a graphicFrame
            anchors = []
            for n in zf.namelist():
                if n.startswith("xl/drawings/drawing") and n.endswith(".xml"):
                    drawing = zf.read(n).decode()
                    if "graphicFrame" in drawing:
                        anchors += zip_charts_anchor_cols(drawing)
        # both donuts cell-anchored on the band row 9 (0-based 8): cols 0/3
        assert sorted(anchors) == [(0, 8), (3, 8)]
        assert ws_row_height(generated, "Bin Chart", 9) == pytest.approx(250, abs=0.3)
        # the series feed the TABLES: categories B15:B16, counts D15:D16,
        # series name = the HB section title cell
        hb = next(c for c in charts if "Hardware" in c)
        assert "'Bin Chart'!$B$15:$B$16" in hb
        assert "'Bin Chart'!$D$15:$D$16" in hb
        assert "'Bin Chart'!$B$13" in hb
        sb = next(c for c in charts if "Software" in c)
        assert "'Bin Chart'!$B$21:$B$22" in sb
        assert "'Bin Chart'!$D$21:$D$22" in sb
        assert "'Bin Chart'!$B$19" in sb

    def test_bin_charts_use_modern_palette(self, generated):
        # 2026-11-11: explicit per-slice modern palette colors
        from deps.exporter.charts import BIN_SLICE_PALETTE

        _, _, report_path, _, _ = generated
        with zipfile.ZipFile(report_path) as zf:
            for name in zf.namelist():
                if name.startswith("xl/charts/chart") and name.endswith(".xml"):
                    data = zf.read(name).decode()
                    assert "<c:dPt>" in data, name
                    assert BIN_SLICE_PALETTE[0] in data, name

    def test_donuts_reference_the_visible_tables(self, generated):
        # 2026-11-11 ruling: the donuts feed from the HW/SW TABLES below
        # them (Bin column B, Count column D, section title as series
        # name) — the old hidden BJ/BK helper block is gone
        _, wb, _, _, _ = generated
        ws = wb["Bin Chart"]
        # no helper block anywhere in the far columns
        assert ws["BJ1"].value is None and ws["BK1"].value is None
        for letter in ("BJ", "BK"):
            assert col_dim(ws, letter) is None
        # the tables hold the data the charts reference (2-bin fixture:
        # title 13, header 14, data 15-16)
        assert ws["B15"].value == "PASS_BIN" and ws["D15"].value == 9
        assert ws["B16"].value == "FAIL_BIN" and ws["D16"].value == 1

    # -- structured tables after the fixed charts band (eighth wave) ---------

    @pytest.mark.filterwarnings("ignore:Conditional Formatting extension")
    @pytest.mark.filterwarnings("ignore:Unknown extension")
    def test_bin_structured_tables_written(self, generated):
        # tenth-wave layout per block: banner 1 (50 pt), blank 2, KPIs 3-5,
        # blank 6, "Bin Charts" headline 7, blank 8, charts band 9 (ONE
        # 250 pt row), blank 10, "Detailed Bin Contribution" headline 11,
        # blank 12, HB title 13 + header 14 + 2 data rows 15-16 + Total 17,
        # blank 18, SB title 19 + header 20 + data 21-22 + Total 23;
        # tables sit in columns B:E (column A belongs to the page grid)
        _, wb, _, _, _ = generated
        ws = wb["Bin Chart"]
        # HB section: Headline-2-styled title band, blue header, two data
        # rows, bold Total
        assert ws["B13"].value == "Hardware Bins — Head1 / All Sites"
        assert ws["B13"].font.bold is True  # Headline-2 lookalike
        assert ws["B13"].font.size == 13
        assert [ws.cell(row=14, column=c).value for c in range(2, 6)] == [
            "Bin",
            "Value",
            "Count",
            "Distribution",
        ]
        assert [ws[cell].value for cell in ("B15", "C15", "D15")] == [
            "PASS_BIN",
            1,
            9,
        ]
        assert ws["E15"].value == pytest.approx(0.9)
        assert ws["E15"].number_format == "0.00%"
        assert ws["E16"].value == pytest.approx(0.1)
        assert ws["B17"].value == "Total"
        assert ws["D17"].value == 10
        assert isinstance(ws["D17"].value, int)
        assert ws["E17"].value == 1.0
        assert ws["E17"].number_format == "0.00%"
        # SB section after ONE blank row (18): title 19, header 20,
        # data 21-22, Total 23
        assert ws["B19"].value == "Software Bins — Head1 / All Sites"
        assert ws["B19"].font.bold is True
        assert ws["B19"].font.size == 13
        assert [ws.cell(row=20, column=c).value for c in range(2, 6)] == [
            "Bin",
            "Value",
            "Count",
            "Distribution",
        ]
        assert ws["B23"].value == "Total"
        assert ws["D23"].value == 10

    def test_bin_headlines_charts_and_contribution(self, generated):
        # ninth wave, tenth-wave rows/renames: merged A:F "Headline 1"
        # bands — "Bin Charts" above the donuts (row 7), "Detailed Bin
        # Contribution - HW Bin / SW Bin" above BOTH tables (row 11, NOT
        # replacing the per-table bands)
        _, wb, _, _, _ = generated
        ws = wb["Bin Chart"]
        assert ws["A7"].value == "Bin Charts"
        assert ws["A7"].font.name == "Calibri"  # Headline-1 lookalike
        assert ws["A7"].font.size == 15
        assert ws["A11"].value == "Detailed Bin Contribution - HW Bin / SW Bin"
        assert ws["A11"].font.size == 15
        assert "A7:F7" in {str(r) for r in ws.merged_cells.ranges}
        assert "A11:F11" in {str(r) for r in ws.merged_cells.ranges}

    def test_bin_charts_band_geometry_single_blank_separators(self, generated):
        # the ONE band row 9 holds ONLY the two donuts (no cell values in
        # A:F — helper data lives in BJ/BK); the single blank rows 2, 6,
        # 8, 10, 12, 18 are the ONLY separators between the bands
        _, wb, _, _, _ = generated
        ws = wb["Bin Chart"]
        for row in (2, 6, 8, 10, 12, 18):
            assert all(ws.cell(row=row, column=c).value is None for c in range(1, 7)), (
                row
            )
        # content rows sandwich the separators (row 9 = the charts band:
        # values live in the floating donuts and the hidden helper block)
        for row in (1, 3, 4, 5, 7, 11, 13, 14, 15, 16, 17, 19, 20, 21, 22, 23):
            assert any(
                ws.cell(row=row, column=c).value is not None for c in range(1, 7)
            ), row

    @pytest.mark.filterwarnings("ignore:Conditional Formatting extension")
    @pytest.mark.filterwarnings("ignore:Unknown extension")
    def test_bin_section_title_band_spans_table(self, generated):
        # the Headline-2 band is merged across the table width B:E (eighth wave)
        _, wb, _, _, _ = generated
        ws = wb["Bin Chart"]
        ranges = {str(r) for r in ws.merged_cells.ranges}
        assert "B13:E13" in ranges
        assert "B19:E19" in ranges

    @pytest.mark.filterwarnings("ignore:Conditional Formatting extension")
    @pytest.mark.filterwarnings("ignore:Unknown extension")
    def test_bin_tables_autofilter_on_both(self, generated):
        # ruled: filter arrows on BOTH tables (mockup shows the SB arrows;
        # one worksheet-level auto_filter cannot cover two ranges, so each
        # table is a minimal Excel Table whose column names equal the header
        # cells verbatim — the wave-2 F5 repair hazard avoided)
        _, wb, _, _, _ = generated
        ws = wb["Bin Chart"]
        assert sorted(ws.tables) == ["HB_Head1_AllSites", "SB_Head1_AllSites"]
        # header..last-data rows at B:E; Total rows (17 / 23) outside
        assert ws.tables["HB_Head1_AllSites"].ref == "B14:E16"
        assert ws.tables["SB_Head1_AllSites"].ref == "B20:E22"
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

    @pytest.mark.filterwarnings("ignore:Conditional Formatting extension")
    @pytest.mark.filterwarnings("ignore:Unknown extension")
    def test_bin_tables_data_bars(self, generated):
        # one DataBarRule per table over the Distribution (column E) DATA
        # rows only
        _, wb, _, _, _ = generated
        ws = wb["Bin Chart"]
        for data_range in ("E15:E16", "E21:E22"):
            (rule,) = [
                r
                for cf in ws.conditional_formatting
                for r in cf.rules
                if str(cf.sqref) == data_range
            ]
            assert rule.type == "dataBar"
            # xlsxwriter omits showValue=1 (OOXML default: shown)
            assert rule.dataBar.showValue in (True, None)
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
        # issue #2: the wafer block's info rows come from the flagged
        # payload (sbin cells dropped — the SW table carries the counts)
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
        assert "WAFER_ID:" in wafer  # identity key (value merged B2:D2)
        assert "Yield" in wafer  # KPI card label — the % string is gone
        assert "TREND-ROW" in trend

    def test_bin_sheet_uniform_page_widths(self, generated):
        # tenth wave: the Bin sheet is a fixed A:F page — all six columns
        # share width 15 (KPI grid + tables; the 250 pt band charts
        # overflow the grid on absolute anchors); G+ keep their defaults
        # (no column-dimension entry in the file — openpyxl would answer
        # 13, its default width, on access), the autofit keeps sizing the
        # hidden helper columns
        _, wb, _, _, _ = generated
        ws = wb["Bin Chart"]
        for letter in ("A", "B", "C", "D", "E", "F"):
            # xlsxwriter stores width + ~0.71 (MDW compensation); the Bin
            # grid recalibration in Excel is a USER CHECKLIST item
            assert col_dim(ws, letter).width == pytest.approx(15, abs=0.75), letter
        assert "G" not in ws.column_dimensions  # beyond the page
        # no helper columns exist anymore (table-driven donuts, 2026-11-11)
        assert col_dim(ws, "BJ") is None
        assert col_dim(ws, "BK") is None

    def test_bin_sheet_a4_print_setup(self, generated):
        # eighth wave (R1): the Bin sheet prints on A4, fit to one page
        # wide with unlimited height; other sheets keep their defaults
        # (the DUT Summary sets fit-to-page but no paper size). The class
        # constant PAPERSIZE_A4 is the string "9"; the saved attribute
        # round-trips as int 9 through openpyxl's Integer descriptor.
        _, wb, _, _, _ = generated
        ws = wb["Bin Chart"]
        assert ws.page_setup.paperSize == int(ws.PAPERSIZE_A4)
        assert ws.page_setup.fitToWidth in (1, None)  # 1 = OOXML default, omitted
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


class TestDuplicatePartIdHighlight:
    """Always-on Excel duplicate-value CF over the Part ID metadata row's
    DIE CELLS ONLY (row 3) — builtin yellow preset, no option flag."""

    @pytest.fixture
    def generated(self, tmp_path, qapp):
        return build_generated(tmp_path, qapp)

    def test_one_duplicate_rule_on_part_id_die_cells(self, generated):
        _, wb, _, _, _ = generated
        rules = cf_block(wb["DUT Summary"], "M3:O3")
        assert len(rules) == 1
        # xlsxwriter emits type "duplicate"; openpyxl reads its expanded
        # form "duplicateValues"
        assert rules[0].type == "duplicateValues"

    def test_rule_uses_builtin_yellow(self, generated):
        _, wb, _, _, _ = generated
        rules = cf_block(wb["DUT Summary"], "M3:O3")
        assert str(rules[0].dxf.fill.bgColor.rgb).endswith("FFFF00")

    def test_base_cells_not_covered(self, generated):
        _, wb, _, _, _ = generated
        # the CF range is the die span only — never the full row
        with pytest.raises(AssertionError, match="no conditional formatting"):
            cf_block(wb["DUT Summary"], "A3:O3")


class TestRemeasureFilterPrune:
    """Filter ON drops re-measured DUT columns before the spec builds;
    OFF is the byte-identical identity path."""

    def test_filter_on_drops_first_duplicate_occurrence(self, tmp_path, qapp, monkeypatch):
        generated = build_generated(
            tmp_path, qapp, duplicate_sn=True, hide_remeasured=True, monkeypatch=monkeypatch
        )
        _, wb, _, _, _ = generated
        ws = wb["DUT Summary"]
        # kept order: dutIndex 1 (SN542) then 2 (the later SN541)
        assert ws["M15"].value == "SN542"
        assert ws["N15"].value == "SN541"
        # the third die column is gone: O15 is the spacer region
        assert ws["O15"].value in (None, "")

    def test_filter_off_keeps_all_columns(self, tmp_path, qapp):
        generated = build_generated(tmp_path, qapp, duplicate_sn=True)
        _, wb, _, _, _ = generated
        ws = wb["DUT Summary"]
        assert ws["M15"].value == "SN541"
        assert ws["N15"].value == "SN542"
        assert ws["O15"].value == "SN541"

    def test_filter_on_blank_ids_never_match(self, tmp_path, qapp, monkeypatch):
        generated = build_generated(
            tmp_path, qapp, blank_ids=True, hide_remeasured=True, monkeypatch=monkeypatch
        )
        _, wb, _, _, _ = generated
        ws = wb["DUT Summary"]
        # nothing dropped: all three die columns survive
        assert ws["M15"].value == "DUT #1"
        assert ws["N15"].value == "DUT #2"
        assert ws["O15"].value == "DUT #3"

    def test_prune_summary_view_removes_exactly_the_drops(self):
        from deps.exporter.worksheets.dut_summary import prune_summary_view

        summary = make_summary()
        dropped = {(1, 0)}
        pruned = prune_summary_view(summary, dropped)
        assert 0 not in pruned["dutInfo"][1]
        assert 1 in pruned["dutInfo"][1]
        assert 0 not in pruned["dut2ind"][1]
        # Data / TestInfo are shared, not rebuilt
        assert pruned["Data"] is summary["Data"]
        assert pruned["TestInfo"] is summary["TestInfo"]
        # the original is untouched (copies, not in-place)
        assert 0 in summary["dutInfo"][1]

    def test_prune_summary_view_empty_set_is_identity(self):
        from deps.exporter.worksheets.dut_summary import prune_summary_view

        summary = make_summary()
        assert prune_summary_view(summary, set()) is summary


class TestStatsExclusionChain:
    """G3: Test Statistics honors the drop-set in every selection — via
    the kargs hand-off when the DUT sheet ran, via the light dutInfo-only
    fetch when it did not."""

    def test_dut_selected_kargs_carry_exclusion(self, tmp_path, qapp, monkeypatch):
        generated = build_generated(
            tmp_path,
            qapp,
            with_stat=True,
            duplicate_sn=True,
            hide_remeasured=True,
            monkeypatch=monkeypatch,
        )
        gen, _, _, _, _ = generated
        stat_kargs = [k for chart, k in gen._captured_stat_kargs if chart == tab.Trend]
        assert stat_kargs and stat_kargs[0].get("dut_exclusion") == {1: {0}}

    def test_dut_selected_filter_off_no_exclusion_key(self, tmp_path, qapp):
        generated = build_generated(tmp_path, qapp, with_stat=True, duplicate_sn=True)
        gen, _, _, _, _ = generated
        stat_kargs = [k for chart, k in gen._captured_stat_kargs if chart == tab.Trend]
        assert stat_kargs and "dut_exclusion" not in stat_kargs[0]

    def test_dut_selected_never_fetches_dutinfo(self, tmp_path, qapp, monkeypatch):
        # waitForDutInfo raises when the DUT sheet ran — the run succeeding
        # is the assertion
        build_generated(
            tmp_path,
            qapp,
            with_stat=True,
            duplicate_sn=True,
            hide_remeasured=True,
            monkeypatch=monkeypatch,
        )

    def test_stat_only_fetches_dutinfo_and_excludes(self, tmp_path, qapp, monkeypatch):
        generated = build_generated(
            tmp_path,
            qapp,
            with_stat=True,
            duplicate_sn=True,
            hide_remeasured=True,
            with_dut=False,
            monkeypatch=monkeypatch,
        )
        gen, wb, _, _, _ = generated
        # the DUT sheet is absent; the fallback fetch ran (stub did not raise)
        assert "DUT Summary" not in wb.sheetnames
        stat_kargs = [k for chart, k in gen._captured_stat_kargs if chart == tab.Trend]
        # partial fetch: SN541 at dutIndex 0 and 2 → first occurrence drops
        assert stat_kargs and stat_kargs[0].get("dut_exclusion") == {1: {0}}


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
        # Fail Num index 10 -> K (left of the ninth-wave Cp insert,
        # unchanged); Cp index 11 -> L (the old Cpk slot); Cpk index 12
        # -> M (Pattern Name shift + Cp insert); data rows only (header
        # row 1 excluded)
        fail_rules = cf_block(ws, "K2:K3")
        assert [r.formula for r in fail_rules] == [
            ["AND(ISNUMBER(K2),K2>0)"],
            ["AND(ISNUMBER(K2),K2=0)"],
        ]
        cp_rules = cf_block(ws, "L2:L3")
        assert [r.formula[0] for r in cp_rules] == [
            "AND(ISNUMBER(L2),L2>=1.67)",
            "AND(ISNUMBER(L2),L2>=1.33)",
            "AND(ISNUMBER(L2),L2>=1)",
            "AND(ISNUMBER(L2),L2<1)",
        ]
        cpk_rules = cf_block(ws, "M2:M3")
        assert [r.formula[0] for r in cpk_rules] == [
            "AND(ISNUMBER(M2),M2>=1.67)",
            "AND(ISNUMBER(M2),M2>=1.33)",
            "AND(ISNUMBER(M2),M2>=1)",
            "AND(ISNUMBER(M2),M2<1)",
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
        assert ws["A1"].font.name == "Cambria"  # Title lookalike
        assert ws["A1"].font.size == 18
        assert "A1:F1" in {str(r) for r in ws.merged_cells.ranges}

    def test_bin_banner_fallback_without_file_info(self, generated):
        # Bin selected WITHOUT File Info -> the report_core default
        # (Review Focus 5: no AttributeError, sheet-name fallback)
        _, wb, _, _, _ = generated
        assert wb["Bin Chart"]["A1"].value == "Bin Chart"
        assert wb["Bin Chart"]["A1"].font.name == "Cambria"


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


class TestTemperatureSplit:
    """Split ON renders one extra sheet per distinct temperature (§7.4);
    the base sheet and Test Statistics stay exactly as with the option OFF."""

    SPLIT_IDS = ["SN541_T25C", "SN542_T-55C", "SN543_T125.5C"]

    def split_generated(self, tmp_path, qapp, monkeypatch, part_ids=None, **kwargs):
        return build_generated(
            tmp_path, qapp,
            part_ids=list(self.SPLIT_IDS) if part_ids is None else part_ids,
            split_temperature=True, monkeypatch=monkeypatch, **kwargs
        )

    def test_split_sheets_exist_ascending_among_themselves(self, tmp_path, qapp, monkeypatch):
        generated = self.split_generated(tmp_path, qapp, monkeypatch)
        _, wb, _, _, _ = generated
        names = wb.sheetnames
        assert "DUT Summary - T-55°C" in names
        assert "DUT Summary - T125.5°C" in names
        assert "DUT Summary - T25°C" in names
        # ascending by temperature among the splits: -55 < 25 < 125.5
        # (pre-Task-8 truth: appended after the other sheets — Task 8
        # pins the adjacent order)
        i55 = names.index("DUT Summary - T-55°C")
        i125 = names.index("DUT Summary - T125.5°C")
        i25 = names.index("DUT Summary - T25°C")
        assert i55 < i25 < i125

    def test_each_split_carries_only_its_temperature(self, tmp_path, qapp, monkeypatch):
        generated = self.split_generated(tmp_path, qapp, monkeypatch)
        _, wb, _, _, _ = generated
        assert wb["DUT Summary - T25°C"]["M15"].value == "SN541"
        assert wb["DUT Summary - T-55°C"]["M15"].value == "SN542"
        assert wb["DUT Summary - T125.5°C"]["M15"].value == "SN543"
        # one die column per sheet: N15 is the spacer region everywhere
        for name in ("DUT Summary - T25°C", "DUT Summary - T-55°C", "DUT Summary - T125.5°C"):
            assert wb[name]["N15"].value in (None, "")

    def test_unresolvable_dut_base_only(self, tmp_path, qapp, monkeypatch):
        generated = self.split_generated(
            tmp_path, qapp, monkeypatch,
            part_ids=self.SPLIT_IDS + ["SN544"],  # no tag, no file temp
        )
        _, wb, _, _, _ = generated
        # base sheet keeps all four DUTs
        base_headers = [wb["DUT Summary"][f"{col}15"].value for col in "MNOP"]
        assert base_headers == ["SN541", "SN542", "SN543", "SN544"]
        # SN544 appears in NO split sheet
        for name in ("DUT Summary - T25°C", "DUT Summary - T-55°C", "DUT Summary - T125.5°C"):
            headers = [wb[name][f"{col}15"].value for col in "MN"]
            assert "SN544" not in headers

    def test_split_off_unchanged(self, tmp_path, qapp):
        generated = build_generated(tmp_path, qapp, part_ids=list(self.SPLIT_IDS))
        _, wb, _, _, _ = generated
        assert wb.sheetnames == ["DUT Summary", "Trend Chart", "Bin Chart", "Wafer Map"]

    def test_single_temperature_no_split_sheets(self, tmp_path, qapp, monkeypatch):
        generated = build_generated(
            tmp_path, qapp,
            part_ids=["SN541_T25C", "SN542_T25C", "SN543_T25C"],
            split_temperature=True, monkeypatch=monkeypatch,
        )
        _, wb, _, _, _ = generated
        assert wb.sheetnames == ["DUT Summary", "Trend Chart", "Bin Chart", "Wafer Map"]

    def test_split_does_not_touch_stats_kargs(self, tmp_path, qapp, monkeypatch):
        generated = self.split_generated(tmp_path, qapp, monkeypatch, with_stat=True)
        gen, _, _, _, _ = generated
        stat_kargs = [k for chart, k in gen._captured_stat_kargs if chart == tab.Trend]
        assert stat_kargs and stat_kargs[0] == {
            "testTuples": [(10, 0, "0_TXCC/VCCC_ASIC"), (12, 0, "2_Func/scanchain")]
        }


class TestDutBlockTabOrder:
    """§7.4 placement ruling: the DUT-derived block (base + chunk overflow +
    temperature splits) sits ADJACENT to the base sheet, ahead of Bin /
    Trend / Wafer — sheets xlsxwriter appends mid-report otherwise land at
    the workbook end."""

    def test_order_dut_block_places_block_after_base(self):
        from deps.exporter.worksheets.dut_summary import order_dut_block

        fileinfo, dut, gdr, stat, trend, chunk, split_a = object(), object(), object(), object(), object(), object(), object()
        worksheets = [fileinfo, dut, gdr, stat, trend, chunk, split_a]
        block = [dut, chunk, split_a]
        assert order_dut_block(worksheets, block) == [fileinfo, dut, chunk, split_a, gdr, stat, trend]

    def test_single_member_block_is_identity(self):
        from deps.exporter.worksheets.dut_summary import order_dut_block

        fileinfo, dut, stat = object(), object(), object()
        worksheets = [fileinfo, dut, stat]
        assert order_dut_block(worksheets, [dut]) == worksheets

    def test_empty_block_is_identity(self):
        from deps.exporter.worksheets.dut_summary import order_dut_block

        worksheets = [object(), object()]
        assert order_dut_block(worksheets, []) == worksheets

    def test_split_sheets_adjacent_to_base_sheet(self, tmp_path, qapp, monkeypatch):
        generated = build_generated(
            tmp_path, qapp,
            part_ids=["SN541_T25C", "SN542_T-55C", "SN543_T125.5C"],
            split_temperature=True, monkeypatch=monkeypatch,
        )
        _, wb, _, _, _ = generated
        assert wb.sheetnames == [
            "DUT Summary",
            "DUT Summary - T-55°C",
            "DUT Summary - T25°C",
            "DUT Summary - T125.5°C",
            "Trend Chart",
            "Bin Chart",
            "Wafer Map",
        ]


# --------------------------------------------------------------------------
# wafer cell-grid bin map (issue #2)
# --------------------------------------------------------------------------


class TestWaferCellMap:
    def test_wafer_sheet_has_no_image_and_two_pics_total(self, generated):
        # the per-wafer PNG is replaced by the cell grid: only the two
        # trend anchors remain in the package drawings
        _, _, report_path, _, _ = generated
        with zipfile.ZipFile(report_path) as zf:
            pics = sum(
                zf.read(n).count(b"<xdr:pic>")
                for n in zf.namelist()
                if n.startswith("xl/drawings/drawing") and n.endswith(".xml")
            )
        assert pics == 2

    def test_block_title_and_sw_table(self, generated):
        # 2026-11-13 layout: title 1, identity line 2 (WAFER_ID key A2 /
        # value merged B2:D2), blank 3, STACKED KPI cards 4-12 (Yield
        # 4-6, Fails 7-9, Total 10-12, each A:D), map anchored at row 4
        # (X headers G4.., Y axis F5.., dies 5..7), properties 14/16/18,
        # SW table from 20 (band 20, header 21, data 22..23, total 24)
        _, wb, _, _, _ = generated
        ws = wb["Wafer Map"]
        assert ws["A1"].value == "W01 – All Sites"
        assert ws["A2"].value == "WAFER_ID:"
        assert ws["B2"].value == "W01"
        assert ws["A4"].value == "Yield"  # bin-chart-style KPI cards
        assert ws["A14"].value == "Wafer Center:"
        assert ws["B14"].value == "(10, 20)"
        assert [ws.cell(row=21, column=c).value for c in range(1, 5)] == list(
            BIN_TABLE_HEADERS
        )
        assert "SW_0_W01" in ws.tables

    def test_kpi_cards_derive_from_the_grid(self, generated):
        # 2026-11-12 user ruling: cards == SW table (grid is the single
        # source). Payload grid: 2 dies sbin 1 (P), 1 die sbin 2 (F).
        # 2026-11-13 dynamic counts: the card values and the table's
        # Count cells are FORMULAS over the die map, with cached results
        _, wb, report_path, _, _ = generated
        ws = wb["Wafer Map"]
        assert ws["A5"].value == "=(C22)/COUNT($G$5:$I$7)"
        assert ws["A8"].value == "=C23"
        assert ws["A11"].value == "=COUNT($G$5:$I$7)"
        assert ws["C22"].value == "=COUNTIF($G$5:$I$7,B22)"
        assert ws["C23"].value == "=COUNTIF($G$5:$I$7,B23)"
        assert ws["C24"].value == "=SUM(C22:C23)"
        from openpyxl import load_workbook

        cached = load_workbook(report_path, data_only=True)["Wafer Map"]
        assert cached["A5"].value == pytest.approx(2 / 3)
        assert cached["A8"].value == 1
        assert cached["A11"].value == 3
        assert cached["C22"].value == 2
        assert cached["C23"].value == 1
        assert cached["C24"].value == 3  # SW table Total

    def test_freeze_panes_on_the_die_grid(self, generated):
        # first block freezes at its die grid (title + identity + X-header
        # rows and columns A:F stay visible while scrolling)
        _, wb, _, _, _ = generated
        assert wb["Wafer Map"].freeze_panes == "G5"

    def test_die_cell_holds_sbin_value(self, generated):
        _, wb, _, _, _ = generated
        ws = wb["Wafer Map"]
        assert ws["G4"].value == "-1"  # first X header (ascending)
        assert ws["F5"].value == "1"  # top Y header = max Y
        assert ws["G6"].value == 1  # die (-1, 0)
        assert ws["H6"].value == 1  # die (0, 0)
        assert ws["H7"].value == 2  # die (0, -1)

    def test_die_columns_stay_square_after_autofit(self, generated):
        # 2026-11-13 column plan pinned AFTER the autofit: A=30, B:D=8,
        # E=3.5 (stored through the xlsxwriter chars->px->chars round trip:
        # 7862/256, 2230/256, 1097/256); F keeps its 30 px and the die
        # columns their 24 px squares (877/256)
        _, wb, _, _, _ = generated
        ws = wb["Wafer Map"]
        assert col_dim(ws, "A").width == 7862 / 256
        assert col_dim(ws, "B").width == 2230 / 256
        assert col_dim(ws, "D").width == 2230 / 256
        assert col_dim(ws, "E").width == 1097 / 256
        assert col_dim(ws, "F").width == 1097 / 256
        assert col_dim(ws, "G").width == 877 / 256

    def test_stacked_wafer_skipped_no_png(self, tmp_path, qapp):
        # 2026-11-12 user ruling: NO PNG on the Wafer sheet at all — the
        # stacked fail-count map (waferInd == -1) has no SW-bin equivalent,
        # so its tuple renders no block (the progress tick still fires)
        _, wb, report_path, _, _ = build_generated(
            tmp_path, qapp, wafer_tuples=[(-1, -1, "STACK")]
        )
        with zipfile.ZipFile(report_path) as zf:
            pics = sum(
                zf.read(n).count(b"<xdr:pic>")
                for n in zf.namelist()
                if n.startswith("xl/drawings/drawing") and n.endswith(".xml")
            )
        assert pics == 2  # 2 trend anchors only
        ws = wb["Wafer Map"]
        values = {c.value for row in ws.iter_rows() for c in row if c.value is not None}
        assert "WAFER-W01-ROW" not in values  # no classic info row either
        assert ws.max_row == 1  # empty sheet: the next wafer would start at A1

    def test_empty_payload_writes_title_and_rows_only(self, tmp_path, qapp):
        # Review Focus pin: a WIR/WCR-only wafer (empty map) must not break
        # the export — title + info rows, no grid, no table
        payload = dict(WAFER_MAP_PAYLOAD)
        empty = dict(payload, Data={}, Statistic={}, Palette={})

        def build(tmp_path, qapp):
            # patch the module the HARNESS lives in: pytest's prepend import
            # mode loads this file as exporter.test_generator_dut (no
            # tests/__init__.py), so `import tests.exporter...` would create
            # a second copy and the patch would land on the wrong one
            import sys

            mod = sys.modules[build_generated.__module__]
            original = mod.WAFER_MAP_PAYLOAD
            mod.WAFER_MAP_PAYLOAD = empty
            try:
                return build_generated(tmp_path, qapp)
            finally:
                mod.WAFER_MAP_PAYLOAD = original

        _, wb, _, _, _ = build(tmp_path, qapp)
        ws = wb["Wafer Map"]
        assert "A1:D1" in {str(r) for r in ws.merged_cells.ranges}
        assert "B2:D2" in {str(r) for r in ws.merged_cells.ranges}
        assert not ws.tables
        assert ws["A2"].value == "WAFER_ID:"  # KPI cells drop out
        assert ws["B2"].value == "W01"
