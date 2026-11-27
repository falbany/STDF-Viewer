"""Tests for deps/exporter/overview.py — the Report Overview first sheet
(2026-11-27 design, IDEAS.md item 7): pure derivation tests (roles,
banner chain, temperatures, notes, spec assembly), a renderer roundtrip,
and generator-level runs through build_generated (the DUT-block harness).
"""

import io

import pytest
import xlsxwriter
from PyQt5 import QtCore

from deps.exporter.overview import (
    derive_notes,
    format_temperatures,
    resolve_banner,
    sheet_roles,
)
from deps.exporter.worksheets.file_info import file_names_from_file_info
from deps.uic_stdExporter import ReportSelection

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


@pytest.fixture(scope="module")
def tr(qapp):
    """English-locale QObject.tr — the same callable the generator hands
    the overview at close time."""
    return QtCore.QObject().tr


def make_book(names):
    """A real in-memory workbook with one worksheet per name; returns the
    worksheet list in order (roles consume live xlsxwriter objects)."""
    wb = xlsxwriter.Workbook(io.BytesIO(), {"in_memory": True})
    return [wb.add_worksheet(name) for name in names]


class TestSheetRoles:
    def test_precreated_map_by_selection_identity(self, tr):
        wss = make_book(
            [
                "File Info",
                "Test Statistics",
                "DUT Summary",
                "GDR & DTR Summary",
                "Trend Chart",
                "Histogram",
                "Bin Chart",
                "Normal Validation",
                "Correlation",
            ]
        )
        # the pre-created subset the generator's sheetDict carries
        sheet_dict = {
            ReportSelection.FileInfo: wss[0],
            ReportSelection.Stat: wss[1],
            ReportSelection.DUT: wss[2],
            ReportSelection.GDR_DTR: wss[3],
            ReportSelection.Trend: wss[4],
            ReportSelection.Histo: wss[5],
            ReportSelection.Bin: wss[6],
            ReportSelection.PPQQ: wss[7],
            ReportSelection.Correlation: wss[8],
        }
        assert sheet_roles(wss, None, sheet_dict, tr) == [
            "Report",  # File Info
            "Statistics",  # Test Statistics
            "Per-DUT",  # DUT Summary
            "Report",  # GDR & DTR Summary
            "Chart",  # Trend Chart
            "Chart",  # Histogram
            "Bins",  # Bin Chart
            "Validation",  # Normal Validation
            "Correlation",  # Correlation
        ]

    def test_chunks_prefix_match_family_role(self, tr):
        wss = make_book(["Test Statistics", "Test Statistics (2)", "DUT Summary (3)"])
        sheet_dict = {ReportSelection.Stat: wss[0]}
        roles = sheet_roles(wss, None, sheet_dict, tr)
        assert roles == ["Statistics", "Statistics", "Per-DUT"]

    def test_temperature_splits_prefix_match_family_role(self, tr):
        wss = make_book(
            ["Test Statistics - T25°C", "Test Correlation Map - T-40°C"]
        )
        assert sheet_roles(wss, None, {}, tr) == ["Statistics", "Correlation"]

    def test_wafer_family(self, tr):
        wss = make_book(["Wafer Heat Map", "Wafer 25 (Merged)", "Wafer 1", "Wafer 1_2"])
        assert sheet_roles(wss, None, {}, tr) == ["Wafer"] * 4

    def test_unmatched_degrades_to_sheet(self, tr):
        wss = make_book(["Mystery Sheet"])
        assert sheet_roles(wss, None, {}, tr) == ["Sheet"]

    def test_overview_excluded_by_identity(self, tr):
        wss = make_book(["Report Overview", "Test Statistics"])
        sheet_dict = {ReportSelection.Stat: wss[1]}
        assert sheet_roles(wss, wss[0], sheet_dict, tr) == ["", "Statistics"]

    def test_translated_ui_prefixes_use_live_tr(self, tr):
        # the writer built "Statistiques (2)" from tr("Test Statistics")
        # in the SAME locale — the role matcher must prefix-match the
        # live strings, never hardcoded English
        def french(s):
            return {
                "Test Statistics": "Statistiques",
                "DUT Summary": "Bilan DUT",
                "Test Correlation Map": "Carte de Corrélation",
            }.get(s, s)

        wss = make_book(["Statistiques (2)", "Bilan DUT - T25°C", "Feuille mystère"])
        assert sheet_roles(wss, None, {}, french) == [
            "Statistics",
            "Per-DUT",
            "Sheet",
        ]


class TestFileNamesFromFileInfo:
    def test_names_from_the_file_name_row(self):
        rows = [["File Name: ", "a.stdf", " b.stdf "], ["Yield: ", "96.4%", "98.1%"]]
        assert file_names_from_file_info(rows) == ["a.stdf", "b.stdf"]

    def test_missing_anchor_yields_empty(self):
        assert file_names_from_file_info([["Yield: ", "96%"]]) == []

    def test_translated_anchor_yields_empty(self):
        rows = [["Nom du fichier: ", "a.stdf"]]
        assert file_names_from_file_info(rows) == []


class TestResolveBanner:
    def test_stage1_anchors_resolve(self, tr):
        rows = [
            ["Product ID: ", "TOP-3491", "TOP-3491"],
            ["Sublot ID: ", "2A", "2B"],
        ]
        assert resolve_banner(rows, tr) == ("TOP-3491 / 2A | TOP-3491 / 2B", True)

    def test_stage2_falls_back_to_file_names(self, tr):
        rows = [["File Name: ", "a.stdf", "b.stdf"]]
        assert resolve_banner(rows, tr) == ("2 — a.stdf · b.stdf", False)

    def test_stage3_falls_back_to_sheet_name(self, tr):
        assert resolve_banner(None, tr) == ("Report Overview", False)
        assert resolve_banner([], tr) == ("Report Overview", False)

    def test_rows_present_without_any_anchor(self, tr):
        rows = [["Yield: ", "96.4%"]]
        assert resolve_banner(rows, tr) == ("Report Overview", False)


class TestFormatTemperatures:
    def test_sorted_distinct_labels(self):
        assert format_temperatures({1: 25.0, 2: -40.0, 3: None}) == "T-40°C · T25°C"

    def test_all_none_yields_none(self):
        assert format_temperatures({1: None, 2: None}) is None

    def test_empty_or_missing_dict_yields_none(self):
        assert format_temperatures({}) is None
        assert format_temperatures(None) is None

    def test_fractional_temperature_keeps_its_decimal(self):
        assert format_temperatures({1: 25.5}) == "T25.5°C"


def notes_kwargs(**overrides):
    """derive_notes' baseline: a quiet, English, nothing-happened export."""
    kwargs = dict(
        sheet_names=["File Info", "Test Statistics", "DUT Summary", "Wafer 1"],
        facts={},
        drop_count=None,
        hide_remeasured=False,
        wafer_heat_selected=False,
        correlation_selected=False,
        anchors_resolved=True,
        file_info_present=False,
        tr=None,
    )
    kwargs.update(overrides)
    return kwargs


class TestDeriveNotes:
    def test_quiet_export_has_no_notes(self, tr):
        assert derive_notes(**{**notes_kwargs(), "tr": tr}) == []

    def test_note1_drops_when_filter_on(self, tr):
        notes = derive_notes(**{**notes_kwargs(), "tr": tr, "hide_remeasured": True, "drop_count": 312})
        assert notes == ["312 re-measured DUTs dropped (kept latest)"]

    def test_note1_idle_filter_on_nothing_dropped_suppressed(self, tr):
        notes = derive_notes(**{**notes_kwargs(), "tr": tr, "hide_remeasured": True, "drop_count": 0})
        assert notes == []

    def test_note2_matrix_chunks(self, tr):
        notes = derive_notes(**{**notes_kwargs(), "tr": tr, "facts": {"stat_chunks": 2}})
        assert notes == ["Test Statistics split into 2 sheets (Excel column cap)"]

    def test_note2_single_chunk_stays_silent(self, tr):
        notes = derive_notes(**{**notes_kwargs(), "tr": tr, "facts": {"stat_chunks": 1}})
        assert notes == []

    def test_note3_dut_summary_chunks_from_names(self, tr):
        notes = derive_notes(
            **{
                **notes_kwargs(),
                "tr": tr,
                "sheet_names": ["DUT Summary", "DUT Summary (2)", "DUT Summary (3)"],
            }
        )
        assert notes == ["DUT Summary split into 3 sheets (Excel column cap)"]

    def test_note4_splits_with_unresolved_clause(self, tr):
        facts = {"stat_splits": [(-40.0, 9216), (25.0, 9204)], "stat_unresolved": 12}
        notes = derive_notes(**{**notes_kwargs(), "tr": tr, "facts": facts})
        assert notes == [
            "Temperature splits: T-40°C, T25°C — 12 DUTs without a temperature stay on the base sheet"
        ]

    def test_note4_without_unresolved_clause(self, tr):
        facts = {"stat_splits": [(-40.0, 9216), (25.0, 9204)], "stat_unresolved": 0}
        notes = derive_notes(**{**notes_kwargs(), "tr": tr, "facts": facts})
        assert notes == ["Temperature splits: T-40°C, T25°C"]

    def test_note5_merged_wafer_per_entry(self, tr):
        facts = {"merged_wafers": [("W25", "Wafer 25 (Merged)", 2, 14)]}
        notes = derive_notes(**{**notes_kwargs(), "tr": tr, "facts": facts})
        assert notes == ["Wafer W25 merged from 2 files (14 dies superseded)"]

    def test_note6_heat_map_present(self, tr):
        notes = derive_notes(
            **{
                **notes_kwargs(),
                "tr": tr,
                "sheet_names": ["Wafer 1", "Wafer Heat Map"],
            }
        )
        assert notes == ["Stacked wafer view added (all wafers)"]

    def test_note7_heat_selected_but_absent(self, tr):
        notes = derive_notes(**{**notes_kwargs(), "tr": tr, "wafer_heat_selected": True})
        assert notes == ["Wafer Heat Map not exported — selection has no all-wafers data"]

    def test_note7_heat_present_not_fired(self, tr):
        notes = derive_notes(
            **{
                **notes_kwargs(),
                "tr": tr,
                "wafer_heat_selected": True,
                "sheet_names": ["Wafer 1", "Wafer Heat Map"],
            }
        )
        assert notes == ["Stacked wafer view added (all wafers)"]

    def test_note8_correlation_selected_but_absent(self, tr):
        notes = derive_notes(**{**notes_kwargs(), "tr": tr, "correlation_selected": True})
        assert notes == ["Test Correlation Map not exported — no DUT context"]

    def test_note8_correlation_present_not_fired(self, tr):
        notes = derive_notes(
            **{
                **notes_kwargs(),
                "tr": tr,
                "correlation_selected": True,
                "sheet_names": ["Test Correlation Map", "Test Correlation Map - T25°C"],
            }
        )
        assert notes == []

    def test_note9_translated_ui_with_table(self, tr):
        notes = derive_notes(
            **{**notes_kwargs(), "tr": tr, "anchors_resolved": False, "file_info_present": True}
        )
        assert notes == ["Non-English UI — some columns keep raw values"]

    def test_note9_no_table_stays_silent(self, tr):
        notes = derive_notes(
            **{**notes_kwargs(), "tr": tr, "anchors_resolved": False, "file_info_present": False}
        )
        assert notes == []

    def test_maximal_scenario_keeps_the_pinned_order(self, tr):
        facts = {
            "stat_chunks": 2,
            "stat_splits": [(-40.0, 9216), (25.0, 9204)],
            "stat_unresolved": 12,
            "merged_wafers": [("W25", "Wafer 25 (Merged)", 2, 14)],
        }
        notes = derive_notes(
            **{
                **notes_kwargs(),
                "tr": tr,
                "sheet_names": ["File Info", "Test Statistics", "DUT Summary", "DUT Summary (2)", "Wafer 1"],
                "facts": facts,
                "drop_count": 312,
                "hide_remeasured": True,
                "correlation_selected": True,
                "anchors_resolved": False,
                "file_info_present": True,
            }
        )
        assert notes == [
            "312 re-measured DUTs dropped (kept latest)",
            "Test Statistics split into 2 sheets (Excel column cap)",
            "DUT Summary split into 2 sheets (Excel column cap)",
            "Temperature splits: T-40°C, T25°C — 12 DUTs without a temperature stay on the base sheet",
            "Wafer W25 merged from 2 files (14 dies superseded)",
            "Test Correlation Map not exported — no DUT context",
            "Non-English UI — some columns keep raw values",
        ]
