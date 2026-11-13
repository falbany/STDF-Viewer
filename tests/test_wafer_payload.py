"""Tests for the flagged tab.Wafer payload builders (issue #2 cell-grid
wafer map): the pure helpers `wafer_bin_palette` (SharedSrc) and
`wafer_payload_rows` (DataInterface) — the flagged fetch branch composes
them GUI-side; the unflagged classic rows keep their trailing sbin cells.
Issue #5 adds `stacked_report_payload` for the stacked wafer tuple
(waferIndex == -1): cross-wafer fail-count aggregation, no sbin machinery.
"""

from deps.DataInterface import stacked_report_payload, wafer_payload_rows
from deps.SharedSrc import wafer_bin_palette

STATS = {
    "VHeader": ["#0 / All Sites"],
    "Rows": [
        [
            ("Yield: 87.42%", -1, False),
            ("Total: 412", -1, False),
            ("Good Dies\nBin1: 87.4%", 1, False),
        ]
    ],
}


def test_rows_drop_sbin_cells_keep_info_cells():
    assert wafer_payload_rows(STATS) == [["Yield: 87.42%", "Total: 412"]]


def test_unflagged_rows_keep_bin_cells():
    # Review Focus pin: the classic getWaferStatistics shape (trailing
    # per-sbin cells) must stay untouched for the legacy exporter.
    assert [e[0] for e in STATS["Rows"][0]] == [
        "Yield: 87.42%",
        "Total: 412",
        "Good Dies\nBin1: 87.4%",
    ]


def test_palette_user_color_wins():
    assert wafer_bin_palette({"Data": {3: {}}}, {}, {3: "#123456"}) == {3: "#123456"}


def test_palette_bin_pf_fallbacks():
    sbin_dict = {1: {"BIN_PF": "P"}, 2: {"BIN_PF": "F"}, 9: {"BIN_PF": ""}, 7: {}}
    pal = wafer_bin_palette({"Data": {k: {} for k in sbin_dict}}, sbin_dict, {})
    assert pal == {1: "#00CC00", 2: "#CC0000", 9: "#FE7B00", 7: "#FE7B00"}


def test_palette_empty_data():
    assert wafer_bin_palette({}, {}, {}) == {}


STACKED = {
    "Bounds": (5, 1, 1, 1),
    "ID": (-1, -1),
    "Data": {
        0: {"x": [7, 8], "y": [7, 8]},  # probed, never failed (count-0 key)
        1: {"x": [1, 2, 3], "y": [1, 1, 1]},
        3: {"x": [4], "y": [1]},
    },
    "Statistic": {0: 0, 1: 1, 3: 3},
}


def test_stacked_rows_arithmetic():
    # counts are the dict KEYS: total fails = 0*2 + 1*3 + 3*1 = 6, not 4;
    # Fail Coordinates counts FAILING coordinates only — the count-0
    # bucket (probed dies that never failed) is excluded, else a normal
    # wafer would report its whole probed-die count
    out = stacked_report_payload(STACKED, [-1], [0, 1])
    assert out["rows"] == [
        ["Total Failed DUTs: 6", "Fail Coordinates: 4", "Files: 2"]
    ]


def test_stacked_payload_enriches_info():
    out = stacked_report_payload(STACKED, [-1], [0])
    assert out["map"]["Info"] == (1.0, "", False, False, "Stacked Wafer Map", [-1])
    assert "Palette" not in out["map"] and "BinPF" not in out["map"]


def test_stacked_payload_empty():
    out = stacked_report_payload({}, [-1], [0])
    assert out["rows"] == [["Total Failed DUTs: 0", "Fail Coordinates: 0", "Files: 1"]]
    assert out["map"]["Info"][4] == "Stacked Wafer Map"
