"""Tests for deps/exporter/worksheets/test_correlation.py — the Test
Correlation Map (2026-11-19 ruling): ALL tests, non-failing included,
failing-DUTs-descending order, capped 100 x 100, ONE numpy matmul
written as plain cached values under the wafer grid's 3_color_scale,
one map per temperature partition when split-temp mode resolves >= 2
temperatures, gated on the Test Correlation selection + a prepared
DUT-block context. Pure spec tests, a renderer roundtrip, and
generator-level runs through build_generated (the DUT-block harness).
"""

import pytest
from PyQt5 import QtCore

from deps.exporter.theme import ModernTheme
from deps.exporter.worksheets.test_correlation import (
    CORRELATION_CAP,
    build_correlation_spec,
    render_correlation_sheet,
)
from tests.exporter.conftest import load_workbook_bytes, render_workbook
from tests.exporter.test_generator_dut import build_generated

# WAFER_DIE_WIDTH (2.71 chars) as openpyxl reads it back: xlsxwriter
# stores the 24 px width as 877/256 (the test_charts.py die-width pin)
WAFER_DIE_WIDTH_ROUNDTRIP = 877 / 256

# openpyxl's reader drops features it cannot parse (xlsxwriter's x14
# extensions) — the suite asserts those via the raw package XML instead;
# silence the reader-side noise here (the sibling modules' pytestmark).
pytestmark = [
    pytest.mark.filterwarnings("ignore:Sparkline Group extension"),
    pytest.mark.filterwarnings("ignore:Conditional Formatting extension"),
    pytest.mark.filterwarnings("ignore:Unknown extension"),
]


@pytest.fixture(scope="module")
def qapp():
    """QObject.tr needs a QCoreApplication instance (no widgets required)."""
    app = QtCore.QCoreApplication.instance()
    if app is None:
        app = QtCore.QCoreApplication([])
    yield app


class TestCorrelationSpec:
    """The pure builder: selection, ordering, cap, the matmul, anchors."""

    def test_ordering_non_failing_in_never_executed_out(self):
        tests = [
            (10, "A", [0, 1, 0]),  # fails 1
            (12, "B", [1, 1, 0]),  # fails 2 -> top-left
            (14, "C", [0, 0, 0]),  # non-failing, executed -> rides at the bottom
            (16, "D", ["", "", ""]),  # never executed -> off the map
        ]
        spec = build_correlation_spec("Test Correlation Map", tests)
        assert spec["nums"] == ["12", "10", "14"]
        assert spec["fails"] == [2, 1, 0]
        assert spec["labels"] == ["12 B", "10 A", "14 C"]
        assert spec["capped"] is False

    def test_ties_keep_first_appearance(self):
        tests = [(20, "X", [0, 0]), (10, "Y", [0, 0]), (30, "Z", [1, 0])]
        spec = build_correlation_spec("s", tests)
        assert spec["nums"] == ["30", "20", "10"]

    def test_cap_100(self):
        tests = [
            (i, f"T{i:03d}", [1] * (i % 3 + 1) + [0] * (5 - (i % 3 + 1)))
            for i in range(120)
        ]
        spec = build_correlation_spec("s", tests)
        assert len(spec["nums"]) == CORRELATION_CAP == 100
        assert spec["capped"] is True and spec["kept"] == 120
        assert "first 100 of 120 tests" in spec["legend"]
        # the cap keeps the TOP-failing tests (descending)
        assert spec["fails"] == [3] * 40 + [2] * 40 + [1] * 20

    def test_matrix_matches_brute_force_and_diagonal(self):
        mirrors = [
            [1, 0, "", 1],
            [0, 1, 1, ""],
            ["", "", 1, 1],
            [0, 0, 0, 0],
        ]
        tests = [(i, f"t{i}", m) for i, m in enumerate(mirrors)]
        spec = build_correlation_spec("s", tests)
        brute = [
            [
                sum(
                    1 for d in range(4) if mirrors[a][d] == 1 and mirrors[b][d] == 1
                )
                for b in range(4)
            ]
            for a in range(4)
        ]
        assert spec["matrix"] == brute
        # the diagonal IS each test's own failing DUTs
        assert [spec["matrix"][i][i] for i in range(4)] == spec["fails"]

    def test_all_pass_grid_forces_green_anchor(self):
        spec = build_correlation_spec("s", [(1, "A", [0, 0]), (2, "B", [0, 0])])
        gradient = spec["gradient"]
        assert gradient["min_value"] == 0
        assert gradient["max_value"] == 1  # the all-zero grid sits at the green end

    def test_no_executed_tests_returns_none(self):
        assert build_correlation_spec("s", [(1, "A", ["", ""])]) is None
        assert build_correlation_spec("s", []) is None


def _render(sheet_name, tests):
    spec = build_correlation_spec(sheet_name, tests)
    assert spec is not None

    def build(wb, ws_by_name):
        render_correlation_sheet(ws_by_name[sheet_name], spec, ModernTheme(wb))

    wb = load_workbook_bytes(render_workbook([sheet_name], build))
    return wb[sheet_name], spec


class TestCorrelationRender:
    """The wafer-grid layout: F = row labels, G.. = counts, ONE scale."""

    @pytest.fixture
    def rendered(self):
        tests = [
            (12, "2_Func/scanchain", [0, 1, 0]),
            (10, "0_TXCC/VCCC_ASIC", [0, 0, 0]),
            (14, "4_Continuity", ["", 0, 1]),
        ]
        return _render("Test Correlation Map", tests)

    def test_layout_and_values(self, rendered):
        sheet, spec = rendered
        assert sheet["F3"].value == "Test"  # the corner (Y-axis header slot)
        # column headers: the full labels, rotated 90° so the wafer-die-
        # width columns only carry the counts
        assert [sheet.cell(3, 7 + i).value for i in range(3)] == spec["labels"]
        assert sheet.cell(3, 7).alignment.textRotation == 90
        # xlsxwriter quantizes row heights to 0.75 pt steps
        assert sheet.row_dimensions[3].height == pytest.approx(spec["header_row_height"], abs=1)
        assert sheet["F4"].value == spec["labels"][0]
        assert [sheet.cell(4, 7 + i).value for i in range(3)] == spec["matrix"][0]
        assert sheet["F5"].value == spec["labels"][1]
        # panes frozen at the grid's top-left corner: headers stay visible
        assert sheet.freeze_panes == "G4"

    def test_wafer_die_column_width(self, rendered):
        sheet, _ = rendered
        # the wafer map's die width (24 px — xlsxwriter stores 877/256)
        assert sheet.column_dimensions["G"].width == WAFER_DIE_WIDTH_ROUNDTRIP

    def test_one_color_scale_over_the_grid(self, rendered):
        sheet, spec = rendered
        rules = [
            rule for cf in sheet.conditional_formatting for rule in cf.rules if rule.colorScale
        ]
        assert len(rules) == 1
        cfvo = rules[0].colorScale.cfvo
        assert [str(v.type) for v in cfvo] == ["num", "num", "num"]
        assert [float(v.val) for v in cfvo] == [
            float(spec["gradient"]["min_value"]),
            float(spec["gradient"]["mid_value"]),
            float(spec["gradient"]["max_value"]),
        ]


class TestCorrelationMapSheets:
    """Generator-level: gating, tab placement, split-temp partitions."""

    def test_sheet_when_selected(self, tmp_path, qapp, monkeypatch):
        gen, wb, _, _, _ = build_generated(
            tmp_path, qapp, correlation_selected=True, monkeypatch=monkeypatch
        )
        assert wb.sheetnames == [
            "Test Summary",
            "DUT Summary",
            "Test Correlation Map",
            "Trend Chart",
            "Bin Chart",
            "Wafer W01",
        ]
        sheet = wb["Test Correlation Map"]
        assert sheet["F3"].value == "Test"
        # fails-descending: test 12 (one failing DUT) precedes test 10
        assert sheet["G3"].value == "12 2_Func/scanchain"
        assert sheet["H3"].value == "10 0_TXCC/VCCC_ASIC"
        assert sheet["F4"].value == "12 2_Func/scanchain"
        assert sheet["G4"].value == 1 and sheet["H4"].value == 0
        assert sheet["F5"].value == "10 0_TXCC/VCCC_ASIC"
        assert sheet["G5"].value == 0 and sheet["H5"].value == 0
        assert gen.loopCount == gen.totalLoopCnt

    def test_absent_by_default(self, tmp_path, qapp):
        _, wb, _, _, _ = build_generated(tmp_path, qapp)
        assert "Test Correlation Map" not in wb.sheetnames

    def test_correlation_alone_standalone_export(self, tmp_path, qapp, monkeypatch):
        # the 2026-11-19 fix: NO Test Summary / DUT Summary / Test
        # Statistics selected — the map still builds (its context tick is
        # in the GUI denominator, so the progress math never divides by 0)
        gen, wb, _, _, _ = build_generated(
            tmp_path, qapp, with_dut=False, correlation_selected=True, monkeypatch=monkeypatch
        )
        assert "Test Correlation Map" in wb.sheetnames
        sheet = wb["Test Correlation Map"]
        assert sheet["F4"].value == "12 2_Func/scanchain"
        assert sheet["G4"].value == 1
        assert gen.loopCount == gen.totalLoopCnt

    def test_split_two_temps_two_maps(self, tmp_path, qapp, monkeypatch):
        _, wb, _, _, _ = build_generated(
            tmp_path,
            qapp,
            correlation_selected=True,
            split_temperature=True,
            monkeypatch=monkeypatch,
            part_ids=["SN541_T25C", "SN542_T85C", "SN543_T25C"],
        )
        # split mode: one map per resolved temperature, no all-DUTs map
        assert "Test Correlation Map" not in wb.sheetnames
        t25 = wb["Test Correlation Map - T25°C"]
        t85 = wb["Test Correlation Map - T85°C"]
        # 25C: tests 10 + 12 both non-failing on DUTs {0, 2} — zeros grid
        # in first-appearance order
        assert [t25.cell(3, 7 + i).value for i in range(2)] == [
            "10 0_TXCC/VCCC_ASIC",
            "12 2_Func/scanchain",
        ]
        assert [t25.cell(4, 7 + i).value for i in range(2)] == [0, 0]
        # 85C: the single DUT fails test 12 only
        assert t85["G3"].value == "12 2_Func/scanchain"
        assert t85["F4"].value == "12 2_Func/scanchain"
        assert t85["G4"].value == 1
