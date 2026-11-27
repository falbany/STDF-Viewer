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
    build_overview_spec,
    derive_notes,
    format_temperatures,
    render_overview,
    resolve_banner,
    sheet_roles,
)
from deps.exporter.theme import ModernTheme
from deps.exporter.worksheets.file_info import file_names_from_file_info
from deps.uic_stdExporter import ReportSelection, dataChannel, signals
from tests.exporter.conftest import load_workbook_bytes, render_workbook
from openpyxl import load_workbook

from deps.exporter.generator import ModernReportGenerator
from tests.exporter.test_generator_dut import build_generated

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


def spec_kwargs(**overrides):
    """build_overview_spec baseline: full-ish selection, everything in hand."""
    wss = make_book(
        [
            "Report Overview",
            "File Info",
            "Test Statistics",
            "Test Statistics (2)",
            "Test Statistics - T-40°C",
            "DUT Summary",
            "Wafer 1",
            "Wafer Heat Map",
        ]
    )
    file_info_rows = [
        ["File Name: ", "a.stdf", "b.stdf"],
        ["Product ID: ", "TOP-3491", "TOP-3491"],
        ["Sublot ID: ", "2A", "2B"],
        ["Test Temperature: ", "25", "-40"],
    ]
    facts = {
        "remeasured_groups": 156,
        "stat_chunks": 2,
        "stat_splits": [(-40.0, 9216)],
        "stat_unresolved": 12,
        "stat_sheet_duts": [
            ("Test Statistics", 16384),
            ("Test Statistics (2)", 2048),
            ("Test Statistics - T-40°C", 9216),
        ],
        "merged_wafers": [("W25", "Wafer 25 (Merged)", 2, 14)],
    }
    kwargs = dict(
        wss=wss,
        overview_ws=wss[0],
        sheet_dict={
            ReportSelection.FileInfo: wss[1],
            ReportSelection.Stat: wss[2],
            ReportSelection.DUT: wss[5],
        },
        file_info_rows=file_info_rows,
        facts=facts,
        drop_count=312,
        hide_remeasured=True,
        split_temperature=True,
        wafer_heat_selected=True,
        correlation_selected=False,
        ctx_temps={1: 25.0, 2: -40.0},
        selected_fids=[1, 2],
        test_count=148,
        dut_count=18432,
        heads=[1],
        sites=list(range(8)),
        wafer_count=25,
        generated="2026-11-27 14:32",
        tr=None,
    )
    kwargs.update(overrides)
    return kwargs


class TestBuildOverviewSpec:
    def test_full_spec_shape_and_values(self, tr):
        spec = build_overview_spec(**{**spec_kwargs(), "tr": tr})
        assert set(spec) == {"title", "facts", "sheets", "notes"}
        assert spec["title"] == "TOP-3491 / 2A | TOP-3491 / 2B"
        labels = [label for label, _value in spec["facts"]]
        assert labels == [
            "Generated",
            "Files",
            "Temperatures",
            "Tests",
            "DUTs",
            "Heads / Sites",
            "Wafers",
            "Filters",
        ]
        values = dict(spec["facts"])
        assert values["Generated"] == "2026-11-27 14:32"
        assert values["Files"] == "2 — a.stdf · b.stdf"
        assert values["Temperatures"] == "T-40°C · T25°C"
        assert values["Tests"] == "148"
        assert (
            values["DUTs"] == "18,432 tested · 312 superseded · 156 re-measured (kept latest)"
        )
        assert values["Heads / Sites"] == "1 head · 8 sites"
        assert values["Wafers"] == "25 (1 merged)"
        assert values["Filters"] == "Re-measure filter ON (keep latest) · Temperature split ON"

    def test_sheet_rows_hexes_duts_notes(self, tr):
        spec = build_overview_spec(**{**spec_kwargs(), "tr": tr})
        rows = spec["sheets"]
        assert [row["name"] for row in rows] == [
            "File Info",
            "Test Statistics",
            "Test Statistics (2)",
            "Test Statistics - T-40°C",
            "DUT Summary",
            "Wafer 1",
            "Wafer Heat Map",
        ]  # the overview itself is excluded
        by_name = {row["name"]: row for row in rows}
        assert by_name["File Info"]["hex"] == "94A3B8"
        assert by_name["Test Statistics"]["hex"] == "16A34A"
        assert by_name["Test Statistics"]["duts"] == "16,384"
        assert by_name["Test Statistics (2)"]["duts"] == "2,048"
        assert by_name["Test Statistics (2)"]["note"] == "chunk 2 of 2"
        assert by_name["Test Statistics - T-40°C"]["duts"] == "9,216"
        assert by_name["Test Statistics - T-40°C"]["note"] == "temperature split"
        assert by_name["DUT Summary"]["hex"] == "1E40AF"
        assert by_name["DUT Summary"]["duts"] == "18,432"
        assert by_name["Wafer 1"]["hex"] == "7C3AED"
        assert by_name["Wafer 1"]["duts"] == "—"
        assert by_name["Wafer Heat Map"]["note"] == "all wafers, stacked"
        assert all(row["role"] for row in rows)

    def test_notes_flow_through(self, tr):
        spec = build_overview_spec(**{**spec_kwargs(), "tr": tr})
        assert "312 re-measured DUTs dropped (kept latest)" in spec["notes"]
        assert spec["notes"][0] == "312 re-measured DUTs dropped (kept latest)"

    def test_spec_without_dut_context(self, tr):
        wss = make_book(
            ["Report Overview", "File Info", "Test Statistics", "DUT Summary"]
        )
        spec = build_overview_spec(
            **{
                **spec_kwargs(
                    wss=wss,
                    overview_ws=wss[0],
                    sheet_dict={
                        ReportSelection.FileInfo: wss[1],
                        ReportSelection.Stat: wss[2],
                        ReportSelection.DUT: wss[3],
                    },
                    file_info_rows=None,
                    facts={},
                    drop_count=None,
                    hide_remeasured=False,
                    split_temperature=False,
                    wafer_heat_selected=False,
                    ctx_temps=None,
                    dut_count=None,
                    wafer_count=None,
                ),
                "tr": tr,
            }
        )
        values = dict(spec["facts"])
        assert values["Files"] == "—"
        assert values["Temperatures"] == "—"
        assert values["DUTs"] == "—"
        assert values["Wafers"] == "—"
        assert values["Filters"] == "None"
        assert spec["notes"] == []
        assert all(row["duts"] == "—" for row in spec["sheets"])

    def test_ctx_temps_win_over_the_table(self, tr):
        # ctx says {T25} only; the table says 25/-40 — the context wins
        spec = build_overview_spec(**{**spec_kwargs(), "tr": tr, "ctx_temps": {1: 25.0}})
        assert dict(spec["facts"])["Temperatures"] == "T25°C"

    def test_table_temps_used_when_ctx_absent(self, tr):
        spec = build_overview_spec(**{**spec_kwargs(), "tr": tr, "ctx_temps": None})
        assert dict(spec["facts"])["Temperatures"] == "T-40°C · T25°C"


def render_spec(sheets, notes=None):
    """A hand-built spec for renderer tests (build_overview_spec's output
    shape, headings included)."""
    return {
        "title": "TOP-3491 / 2A",
        "facts": [
            ("Generated", "2026-11-27 14:32"),
            ("Files", "2 — a.stdf · b.stdf"),
            ("Temperatures", "T-40°C · T25°C"),
            ("Tests", "148"),
            ("DUTs", "18,432 tested"),
            ("Heads / Sites", "1 head · 8 sites"),
            ("Wafers", "25 (1 merged)"),
            ("Filters", "None"),
        ],
        "sheets": sheets,
        "notes": notes if notes is not None else [],
        "headings": {"report": "Report", "sheets": "Sheets", "notes": "Export Notes"},
    }


RENDER_SHEETS = [
    {"name": "File Info", "hex": "94A3B8", "role": "Report", "duts": "—", "note": ""},
    {"name": "Test Statistics", "hex": "16A34A", "role": "Statistics", "duts": "16,384", "note": ""},
    {"name": "Wafer 25 (Merged)", "hex": "7C3AED", "role": "Wafer", "duts": "—", "note": "merged from 2 files"},
    {"name": "Wafer 1_2", "hex": "7C3AED", "role": "Wafer", "duts": "—", "note": ""},
]


def _build_overview(wb, ws_by_name):
    theme = ModernTheme(wb)
    render_overview(ws_by_name["Report Overview"], render_spec(RENDER_SHEETS), theme)


class TestRenderOverview:
    def _load(self, spec=None):
        names = ["Report Overview", "File Info", "Test Statistics", "Wafer 25 (Merged)", "Wafer 1_2"]
        book = render_workbook(
            names,
            lambda wb, ws_by_name: render_overview(
                ws_by_name["Report Overview"], spec or render_spec(RENDER_SHEETS), ModernTheme(wb)
            ),
        )
        return load_workbook_bytes(book)

    def test_banner_and_facts_readable(self):
        ws = self._load()["Report Overview"]
        assert ws.cell(row=1, column=1).value == "TOP-3491 / 2A"
        assert ws.cell(row=3, column=1).value == "Report"
        assert ws.cell(row=4, column=1).value == "Generated"
        assert ws.cell(row=4, column=3).value == "2026-11-27 14:32"
        assert ws.cell(row=11, column=1).value == "Filters"

    def test_hyperlink_targets_resolve(self):
        from urllib.parse import unquote

        book = self._load()
        ws = book["Report Overview"]
        # openpyxl exposes internal hyperlinks per-CELL (ws._hyperlinks only
        # carries the r:id external ones)
        targets = [
            cell.hyperlink.location
            for row in ws.iter_rows()
            for cell in row
            if cell.hyperlink
        ]
        assert len(targets) == 4
        sheet_names = set(book.sheetnames)
        for location in targets:
            target = unquote(location.split("!")[0]).strip("'")
            assert target in sheet_names, location

    def test_chips_carry_tab_fill_and_tab_color(self):
        ws = self._load()["Report Overview"]
        # sheet rows: 1-based 15 = first (header band sits at 14)
        assert ws.cell(row=15, column=1).fill.fgColor.rgb.endswith("94A3B8")
        assert ws.cell(row=16, column=1).fill.fgColor.rgb.endswith("16A34A")
        assert ws.sheet_properties.tabColor.rgb.endswith("1E40AF")

    def test_print_setup_a4_portrait_one_page_wide(self):
        ws = self._load()["Report Overview"]
        assert ws.page_setup.paperSize == 9  # A4
        assert ws.page_setup.orientation == "portrait"
        assert ws.sheet_properties.pageSetUpPr.fitToPage is True
        # xlsxwriter omits fitToWidth when it is Excel's default 1
        assert ws.page_setup.fitToWidth in (1, None)
        assert ws.page_setup.fitToHeight == 0

    def test_gridlines_hidden_and_footer(self):
        ws = self._load()["Report Overview"]
        assert ws.sheet_view.showGridLines is False
        footer = ws.oddFooter.center.text or ""
        assert "&P" in footer and "&N" in footer

    def test_notes_block_omitted_when_empty(self):
        ws = self._load()["Report Overview"]
        values = [cell.value for row in ws.iter_rows() for cell in row if cell.value]
        assert all(value != "Export Notes" for value in values)

    def test_notes_block_renders_when_present(self):
        spec = render_spec(RENDER_SHEETS, notes=["something happened"])
        ws = self._load(spec)["Report Overview"]
        values = [cell.value for row in ws.iter_rows() for cell in row if cell.value]
        assert "Export Notes" in values
        # the renderer prepends the "– " bullet itself
        assert "– something happened" in values


class TestOverviewStashHooks:
    """Task 8: the _ov_facts/_file_info_table stash hooks, populated by
    real build_generated runs (the DUT-block harness)."""

    def test_plain_run_stashes_matrix_facts(self, tmp_path, qapp):
        gen, _wb, _path, _closed, _errors = build_generated(tmp_path, qapp)
        # 3 DUTs, no chunk overflow, no splits, no merges in the baseline
        assert gen._ov_facts["stat_chunks"] == 1
        assert gen._ov_facts["stat_splits"] == []
        assert gen._ov_facts["stat_unresolved"] is None
        names = dict(gen._ov_facts["stat_sheet_duts"])
        assert names["Test Statistics"] == 3
        assert "merged_wafers" not in gen._ov_facts
        # the context's File Info fetch rides the same stash (empty table:
        # the sheet is not selected and no temperature was set)
        assert gen._file_info_table == []

    def test_remeasured_groups_stash(self, tmp_path, qapp, monkeypatch):
        gen, wb, _path, _closed, _errors = build_generated(
            tmp_path, qapp, duplicate_sn=True, hide_remeasured=True, monkeypatch=monkeypatch
        )
        assert len(gen._remeasure_drops) >= 1
        assert gen._ov_facts["remeasured_groups"] >= 1

    def test_split_facts_from_a_split_run(self, tmp_path, qapp, monkeypatch):
        gen, wb, _path, _closed, _errors = build_generated(
            tmp_path,
            qapp,
            split_temperature=True,
            monkeypatch=monkeypatch,
            part_ids=["SN541_T25C", "SN542_T85C", "SN543_T25C"],
        )
        assert "Test Statistics - T25°C" in wb.sheetnames
        assert [temp for temp, _count in gen._ov_facts["stat_splits"]] == [25.0, 85.0]
        names = dict(gen._ov_facts["stat_sheet_duts"])
        assert names["Test Statistics - T25°C"] == 2
        assert names["Test Statistics - T85°C"] == 1
        assert gen._ov_facts["stat_chunks"] == 1

    def test_file_info_table_stash_with_sheet(self, tmp_path, qapp):
        gen, _wb, _path, _closed, _errors = build_generated(
            tmp_path, qapp, with_file_info=True
        )
        assert gen._file_info_table and gen._file_info_table[0][0] == "File Name"

    def test_overview_selected_adds_no_tick(self, tmp_path, qapp):
        # the overview is tickless: a full run with it selected still
        # completes (closed) and stashes the facts
        gen, wb, _path, closed, _errors = build_generated(
            tmp_path, qapp, overview_selected=True
        )
        assert closed == [True]
        assert gen._ov_facts["stat_chunks"] == 1


class TestOverviewIntegration:
    """Task 9: the close-time render through real build_generated runs."""

    @staticmethod
    def _overview_values(wb):
        return [
            cell.value
            for row in wb["Report Overview"].iter_rows()
            for cell in row
            if cell.value is not None
        ]

    def test_overview_is_first_sheet_with_full_selection(self, tmp_path, qapp):
        gen, wb, _path, closed, _errors = build_generated(
            tmp_path, qapp, overview_selected=True, with_file_info=True
        )
        assert closed == [True]
        assert wb.sheetnames[0] == "Report Overview"
        assert wb.sheetnames[1] == "File Info"
        assert wb.sheetnames[2] == "Test Statistics"
        assert wb.sheetnames.index("DUT Summary") > wb.sheetnames.index("Test Statistics")
        assert wb.sheetnames[-1] == "Wafer W01"
        # the stage-1 banner from the File Info anchors (harness table)
        assert wb["Report Overview"].cell(row=1, column=1).value == "ASIC-X / LOT123"

    def test_toc_targets_resolve(self, tmp_path, qapp):
        from urllib.parse import unquote

        _gen, wb, _path, _closed, _errors = build_generated(
            tmp_path, qapp, overview_selected=True
        )
        ws = wb["Report Overview"]
        targets = [
            cell.hyperlink.location
            for row in ws.iter_rows()
            for cell in row
            if cell.hyperlink
        ]
        assert len(targets) == len(wb.sheetnames) - 1
        for location in targets:
            assert unquote(location.split("!")[0]).strip("'") in wb.sheetnames, location

    def test_overview_only_selection(self, tmp_path, qapp):
        # Review Focus 1: nothing else selected — no context, no writers,
        # a one-sheet workbook that still loads cleanly with "—" facts
        sig = signals()
        report_path = str(tmp_path / "overview-only.xlsx")
        report_info = (
            [ReportSelection.Overview],
            [],
            [1],
            [1],
            [-1],
            report_path,
            [],
            0,
        )
        gen = ModernReportGenerator(
            sig, QtCore.QMutex(), QtCore.QWaitCondition(), report_info, dataChannel()
        )
        gen.report_core()
        wb = load_workbook(report_path)
        assert wb.sheetnames == ["Report Overview"]
        values = self._overview_values(wb)
        assert "—" in values
        assert "Report Overview" in values  # the stage-3 banner fallback

    def test_notes_fire_for_scenario(self, tmp_path, qapp, monkeypatch):
        _gen, wb, _path, closed, _errors = build_generated(
            tmp_path,
            qapp,
            overview_selected=True,
            hide_remeasured=True,
            split_temperature=True,
            monkeypatch=monkeypatch,
            part_ids=["SN541_T25C", "SN542_T85C", "SN541_T25C"],
        )
        assert closed == [True]
        values = self._overview_values(wb)
        # the renderer prepends the "– " bullet — match substrings
        assert any("1 re-measured DUTs dropped (kept latest)" in value for value in values)
        assert any("Temperature splits: T25°C, T85°C" in value for value in values)

    def test_force_quit_still_renders_overview(self, tmp_path, qapp, monkeypatch):
        # Review Focus 5: a writer bails mid-run; the finally still paints
        monkeypatch.setattr(
            ModernReportGenerator,
            "writeBinChart",
            lambda self: setattr(self, "forceQuit", True),
        )
        _gen, wb, _path, _closed, errors = build_generated(
            tmp_path, qapp, overview_selected=True
        )
        assert wb.sheetnames[0] == "Report Overview"
        assert not errors

    def test_heat_note_fires_for_stacked_run(self, tmp_path, qapp):
        _gen, wb, _path, _closed, _errors = build_generated(
            tmp_path,
            qapp,
            overview_selected=True,
            heat_selected=True,
            wafer_tuples=[(-1, -1, "STACK")],
        )
        assert any(
            "Stacked wafer view added (all wafers)" in value
            for value in self._overview_values(wb)
        )

    def test_overview_off_is_regression_free(self, tmp_path, qapp):
        _gen, wb, _path, _closed, _errors = build_generated(tmp_path, qapp)
        assert "Report Overview" not in wb.sheetnames
