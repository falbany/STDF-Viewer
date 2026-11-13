"""Tests for deps/exporter/wafer_merge.py: the pure per-die payload merge.

Payload shape: {"rows": list, "map": dict} as DataInterface.getWaferReportPayload
returns; map keys Bounds / ID / Info / Data / Statistic / Palette / BinPF /
Properties. Real Data x/y values are NUMPY ARRAYS (getWaferCoordsDict) —
staging must mirror that or the tests lie about the production path.
No Qt, no workbook — pure dicts.
"""

import numpy as np

from deps.exporter.wafer_merge import merge_wafer_payloads, merged_wafer_groups


def payload(data, bounds=(5, 1, 5, 1), wafer_id="W08", palette=None, bin_pf=None):
    sbin = next(iter(data), 1)
    stat = {s: (f"Bin {s}", len(v["x"]), 100.0 * len(v["x"]) / max(1, sum(len(d["x"]) for d in data.values())))
            for s, v in data.items()}
    m = {"Bounds": bounds, "ID": (0, 0),
         "Info": (1.0, "1 mm × 1 mm", False, False, wafer_id, [-1]),
         "Data": data, "Statistic": stat,
         "Palette": palette or {s: "#00CC00" for s in data},
         "BinPF": bin_pf if bin_pf is not None else {s: "P" for s in data},
         "Properties": [("WAFER ID", wafer_id)]}
    return {"rows": [["WAFER ID: " + wafer_id]], "map": m}


class TestMergeWaferPayloads:
    def test_later_payload_wins_shared_coord(self):
        old = payload({1: {"x": [1, 2], "y": [3, 4]}})                 # dies (1,3) (2,4) bin1
        new = payload({2: {"x": [1], "y": [3]}})                       # die (1,3) re-measured bin2
        merged, sup = merge_wafer_payloads([old, new])
        coords = {(x, y): s for s, xy in merged["Data"].items() for x, y in zip(xy["x"], xy["y"])}
        assert coords == {(1, 3): 2, (2, 4): 1}
        assert sup == {(1, 3)}

    def test_three_payloads_keep_last(self):
        p1 = payload({1: {"x": [1], "y": [1]}})
        p2 = payload({2: {"x": [1], "y": [1]}})
        p3 = payload({3: {"x": [1], "y": [1]}})
        merged, sup = merge_wafer_payloads([p1, p2, p3])
        assert merged["Data"] == {3: {"x": [1], "y": [1]}}
        assert sup == {(1, 1)}

    def test_bounds_are_unioned(self):
        old = payload({1: {"x": [1], "y": [1]}}, bounds=(5, 1, 5, 1))
        new = payload({1: {"x": [7], "y": [6]}}, bounds=(7, 2, 6, 2))
        merged, _ = merge_wafer_payloads([old, new])
        assert merged["Bounds"] == (7, 1, 6, 1)

    def test_union_palette_and_binpf_later_wins(self):
        old = payload({1: {"x": [1], "y": [1]}}, palette={1: "#00CC00"}, bin_pf={1: "P"})
        new = payload({2: {"x": [2], "y": [1]}}, palette={2: "#CC0000"}, bin_pf={2: "F"})
        merged, _ = merge_wafer_payloads([old, new])
        assert merged["Palette"] == {1: "#00CC00", 2: "#CC0000"}
        assert merged["BinPF"] == {1: "P", 2: "F"}

    def test_conflicting_palette_later_wins(self):
        old = payload({1: {"x": [1], "y": [1]}}, palette={1: "#111111"})
        new = payload({1: {"x": [1], "y": [1]}}, palette={1: "#222222"})
        merged, _ = merge_wafer_payloads([old, new])
        assert merged["Palette"] == {1: "#222222"}

    def test_statistic_recomputed_from_merged_grid(self):
        old = payload({1: {"x": [1, 2], "y": [1, 1]}})
        new = payload({2: {"x": [1], "y": [1]}})
        merged, _ = merge_wafer_payloads([old, new])
        assert merged["Statistic"][1][1] == 1      # die (2,1) only
        assert merged["Statistic"][2][1] == 1
        assert abs(sum(v[2] for v in merged["Statistic"].values()) - 100.0) < 1e-9

    def test_identity_from_last_payload(self):
        old = payload({1: {"x": [1], "y": [1]}}, wafer_id="OLD")
        new = payload({1: {"x": [1], "y": [1]}}, wafer_id="W08")
        merged, _ = merge_wafer_payloads([old, new])
        assert merged["Info"][4] == "W08"
        assert merged["Properties"] == [("WAFER ID", "W08")]

    def test_empty_older_payload_merges(self):
        # WIR/WCR-only older file: no Data, non-int Bounds
        empty = {"rows": [["WAFER ID: W08"]], "map": {"Bounds": ("",), "ID": (0, 0)}}
        full = payload({1: {"x": [1], "y": [1]}})
        merged, sup = merge_wafer_payloads([empty, full])
        assert merged["Data"] == {1: {"x": [1], "y": [1]}}
        assert sup == set()

    def test_all_empty_degrades(self):
        empty = {"rows": [], "map": {}}
        merged, sup = merge_wafer_payloads([empty, empty])
        assert merged["Data"] == {}
        assert sup == set()

    def test_numpy_coordinate_arrays_merge(self):
        # 2026-11-13 production crash: real getWaferCoordsDict payloads
        # carry numpy arrays — truth-testing one (`arr or []`) raises
        # ValueError for len > 1 (report dialog: "The truth value of an
        # array ... is ambiguous"); the merge reads them like
        # build_wafer_grid does, with explicit None checks
        old = payload({1: {"x": np.array([1, 2]), "y": np.array([3, 4])}})
        new = payload({2: {"x": np.array([1]), "y": np.array([3])}})
        merged, sup = merge_wafer_payloads([old, new])
        coords = {(x, y): s for s, xy in merged["Data"].items() for x, y in zip(xy["x"], xy["y"])}
        assert coords == {(1, 3): 2, (2, 4): 1}
        assert sup == {(1, 3)}


class TestMergedWaferGroups:
    def test_grouping_skips_stacked_and_singles(self):
        tuples = [(-1, -1, "Stacked Wafer Map"), (0, 0, "W08"), (1, 1, "W08"), (2, 1, "W09")]
        assert merged_wafer_groups(tuples) == ["W08"]

    def test_null_ids_excluded(self):
        tuples = [(0, 0, "None"), (1, 1, "None"), (2, 1, ""), (3, 1, ""), (4, 2, None), (5, 2, None)]
        assert merged_wafer_groups(tuples) == []

    def test_first_appearance_order(self):
        tuples = [(0, 0, "B"), (1, 0, "A"), (2, 1, "B"), (3, 1, "A")]
        assert merged_wafer_groups(tuples) == ["B", "A"]
