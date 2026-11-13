"""Tests for deps/exporter/dedup.py: the re-measure drop-set (pure dicts).

`dutInfo` shape: {fid: {dutIndex: info}} with info[1] = Part ID and
info[10] = DUT Flag text ("Superseded - 0x04" / "Failed - 0x08" / ...),
the classic tuple the DUT Summary fetch returns (schema.rs
FETCH_SELECT_DUT_SUMMARY). No Qt, no workbook — pure functions.
"""

from deps.exporter.dedup import (
    compute_fid_ranks,
    compute_remeasure_drops,
    drops_to_exclusion,
    dut_info_from_partial,
    file_dates_from_rows,
)


def make_info(part_id="", dut_flag="Passed - 0x00"):
    """11-slot dutInfo tuple; only Part ID (1) and DUT Flag (10) matter."""
    return (
        1,  # 0: File ID
        part_id,  # 1: Part ID
        "",  # 2: Part Text
        "1 - 1",  # 3: Head-Site
        42,  # 4: Tests Executed
        "123 ms",  # 5: Test Time
        "Bin 1",  # 6: HW Bin
        "Bin 2",  # 7: SW Bin
        "",  # 8: Wafer ID
        (3, -5),  # 9: (X, Y)
        dut_flag,  # 10: DUT Flag
    )


def dut_info(entries):
    """{fid: {dutIndex: part_id/flag kwargs}} -> classic dutInfo dict."""
    out = {}
    for fid, dut_index, part_id, flag in entries:
        out.setdefault(fid, {})[dut_index] = make_info(part_id, flag)
    return out


class TestComputeRemeasureDrops:
    def test_stdf_flagged_supersede_dropped(self):
        info = dut_info([(1, 0, "SN002", "Superseded - 0x04")])
        assert compute_remeasure_drops(info, [1]) == {(1, 0)}

    def test_cross_file_fail_to_pass_drops_earlier(self):
        # SN002 fails in file 1, re-measured in file 2: the file-1
        # occurrence drops even though nothing is STDF-flagged
        info = dut_info(
            [
                (1, 0, "SN002", "Failed - 0x08"),
                (2, 0, "SN002", "Passed - 0x00"),
            ]
        )
        assert compute_remeasure_drops(info, [1, 2]) == {(1, 0)}

    def test_three_occurrences_drop_first_two(self):
        info = dut_info(
            [
                (1, 0, "SN", "Failed - 0x08"),
                (2, 3, "SN", "Failed - 0x08"),
                (3, 1, "SN", "Passed - 0x00"),
            ]
        )
        assert compute_remeasure_drops(info, [1, 2, 3]) == {(1, 0), (2, 3)}

    def test_latest_wins_even_when_latest_fails(self):
        info = dut_info(
            [
                (1, 0, "SN002", "Passed - 0x00"),
                (2, 0, "SN002", "Failed - 0x08"),
            ]
        )
        # "remeasured", not "best result": the pass is dropped
        assert compute_remeasure_drops(info, [1, 2]) == {(1, 0)}

    def test_empty_and_none_part_ids_never_match(self):
        info = dut_info(
            [
                (1, 0, "", "Passed - 0x00"),
                (2, 0, None, "Passed - 0x00"),
                (1, 5, "", "Failed - 0x08"),
            ]
        )
        assert compute_remeasure_drops(info, [1, 2]) == set()

    def test_case_sensitive_match(self):
        info = dut_info(
            [
                (1, 0, "sn001", "Passed - 0x00"),
                (2, 0, "SN001", "Passed - 0x00"),
            ]
        )
        assert compute_remeasure_drops(info, [1, 2]) == set()

    def test_selected_fids_subset_ignores_unselected(self):
        # a later occurrence in an UNSELECTED fid neither suppresses nor drops
        info = dut_info(
            [
                (1, 0, "SN002", "Failed - 0x08"),
                (2, 0, "SN002", "Passed - 0x00"),
            ]
        )
        assert compute_remeasure_drops(info, [1]) == set()
        # and an unselected occurrence is itself never dropped
        assert compute_remeasure_drops(info, [2]) == set()

    def test_unflagged_same_file_duplicate_drops_earlier(self):
        # testers that never set PART_FLG: same SN twice in one file,
        # no Supersede flag — rule (2) still dedups (latest wins)
        info = dut_info(
            [
                (1, 0, "SN002", "Failed - 0x08"),
                (1, 1, "SN002", "Passed - 0x00"),
            ]
        )
        assert compute_remeasure_drops(info, [1]) == {(1, 0)}

    def test_empty_input(self):
        assert compute_remeasure_drops({}, [1]) == set()


DATES = {
    0: (1, "2026-01-03 10:00:00 (UTC+0)", "2026-01-03 08:00:00 (UTC+0)"),
    1: (1, "2026-01-01 10:00:00 (UTC+0)", "2026-01-01 08:00:00 (UTC+0)"),
    2: (1, "2026-01-02 10:00:00 (UTC+0)", ""),
}


class TestComputeFidRanks:
    def test_merge_mode_ranks_by_fid(self):
        dates = {0: (2, "t2", ""), 1: (1, "t1", "")}  # fid 0 has 2 subfids
        assert compute_fid_ranks(dates) == {0: 0, 1: 1}

    def test_open_mode_ranks_by_finish_t(self):
        # later FINISH_T -> higher rank, regardless of fid
        assert compute_fid_ranks(DATES) == {1: 0, 2: 1, 0: 2}

    def test_missing_finish_falls_back_to_start(self):
        dates = {0: (1, "", "2026-01-05"), 1: (1, "2026-01-01", "")}
        assert compute_fid_ranks(dates) == {1: 0, 0: 1}

    def test_missing_timestamps_fall_back_to_fid_order(self):
        dates = {5: (1, "", ""), 2: (1, "", "")}
        assert compute_fid_ranks(dates) == {2: 0, 5: 1}

    def test_fid_tiebreaks_identical_timestamps(self):
        dates = {3: (1, "same", "same"), 1: (1, "same", "same")}
        assert compute_fid_ranks(dates) == {1: 0, 3: 1}


class TestFileDatesFromRows:
    def test_rows_shaped_to_dates(self):
        rows = [(0, "FINISH_T", 0, "2026-01-03"), (0, "START_T", 0, "2026-01-03"),
                (1, "FINISH_T", 0, "2026-01-01")]
        out = file_dates_from_rows(rows, {0: 1, 1: 2})
        assert out[0] == (1, "2026-01-03", "2026-01-03")
        assert out[1] == (2, "2026-01-01", "")  # missing fields -> ""

    def test_multisubfile_group_takes_latest_value(self):
        rows = [(0, "FINISH_T", 0, "2026-01-01"), (0, "FINISH_T", 1, "2026-01-09")]
        assert file_dates_from_rows(rows, {0: 2})[0][1] == "2026-01-09"


class TestRankAwareDrops:
    def test_open_mode_later_timestamp_wins(self):
        # same Part ID; fid 1 measured EARLIER in time than fid 2 but has
        # a higher fid — with open-mode ranks the fid-1 row must drop
        info = dut_info([(1, 0, "SN", "Passed - 0x00"), (2, 0, "SN", "Failed - 0x08")])
        ranks = compute_fid_ranks({1: (1, "2026-01-05", ""), 2: (1, "2026-01-09", "")})
        assert compute_remeasure_drops(info, [1, 2], fid_ranks=ranks) == {(1, 0)}

    def test_no_ranks_preserves_current_fid_order(self):
        info = dut_info([(1, 0, "SN", "Passed - 0x00"), (2, 0, "SN", "Failed - 0x08")])
        assert compute_remeasure_drops(info, [1, 2]) == {(1, 0)}

    def test_ranks_resort_unsorted_input(self):
        # dict insertion order fid 2 then fid 1: the rank scan must still
        # pick the rank-latest occurrence
        info = dut_info([(2, 0, "SN", "Failed - 0x08"), (1, 0, "SN", "Passed - 0x00")])
        ranks = compute_fid_ranks({1: (1, "2026-01-09", ""), 2: (1, "2026-01-01", "")})
        assert compute_remeasure_drops(info, [1, 2], fid_ranks=ranks) == {(2, 0)}


class TestDutInfoFromPartial:
    def test_partial_padded_to_dutinfo_shape(self):
        partial = {
            1: {
                2: ("SN541", "txt", "Head 1 - Site 1", "Passed - 0x00"),
            }
        }
        assert dut_info_from_partial(partial) == {
            1: {
                2: (
                    "",  # 0: File ID
                    "SN541",  # 1: Part ID
                    "txt",  # 2: Part Text
                    "",  # 3: Head-Site
                    "",  # 4
                    "",  # 5
                    "",  # 6
                    "",  # 7
                    "",  # 8
                    "",  # 9
                    "Passed - 0x00",  # 10: DUT Flag
                )
            }
        }

    def test_empty_fetch_degrades_to_empty(self):
        assert dut_info_from_partial({}) == {}


class TestDropsToExclusion:
    def test_pairs_group_by_fid(self):
        assert drops_to_exclusion({(1, 0), (2, 5)}) == {1: {0}, 2: {5}}

    def test_empty_set_empty_dict(self):
        assert drops_to_exclusion(set()) == {}
