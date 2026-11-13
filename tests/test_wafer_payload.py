"""Tests for the flagged tab.Wafer payload builders (issue #2 cell-grid
wafer map): the pure helpers `wafer_bin_palette` (SharedSrc) and
`wafer_payload_rows` (DataInterface) — the flagged fetch branch composes
them GUI-side; the unflagged classic rows keep their trailing sbin cells.
"""

from deps.DataInterface import wafer_payload_rows
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
