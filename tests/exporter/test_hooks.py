"""Tests for Modern Export hooks: persisted settings fields + the stats
exclusion hook.

GeneralConfig gains `modern_export` (alias "Modern Export", default True)
so the export dialog's "Modern Report" checkbox state survives restarts;
the re-measure feature adds `hide_remeasured` / `split_temperature` and
the pure `apply_dut_exclusion` numpy mask used by
DataInterface.getTestStatistics. The UI/uic wiring is verified by import
checks plus the full suite; these tests are pure config/numpy, no
Qt/GUI required.
"""

import numpy as np

from deps.DataInterface import DataInterface, apply_dut_exclusion
from deps.SharedSrc import REC, SettingParams, calc_cpk


def test_modern_export_default_true():
    assert SettingParams().gen.modern_export is True


def test_modern_export_alias_in_dump():
    dumped = SettingParams().model_dump(by_alias=True)
    assert dumped["General"]["Modern Export"] is True


def test_modern_export_false_round_trips():
    params = SettingParams()
    params.gen.modern_export = False
    dumped = params.model_dump(by_alias=True)
    assert dumped["General"]["Modern Export"] is False
    assert params.gen.modern_export is False


def test_hide_remeasured_default_false():
    assert SettingParams().gen.hide_remeasured is False


def test_hide_remeasured_alias_in_dump():
    dumped = SettingParams().model_dump(by_alias=True)
    assert dumped["General"]["Hide Remeasured"] is False


def test_hide_remeasured_round_trips():
    params = SettingParams()
    params.gen.hide_remeasured = True
    dumped = params.model_dump(by_alias=True)
    assert dumped["General"]["Hide Remeasured"] is True
    assert params.gen.hide_remeasured is True


def test_split_temperature_default_false():
    assert SettingParams().gen.split_temperature is False


def test_split_temperature_alias_in_dump():
    dumped = SettingParams().model_dump(by_alias=True)
    assert dumped["General"]["Split Temperature"] is False


def test_split_temperature_round_trips():
    params = SettingParams()
    params.gen.split_temperature = True
    dumped = params.model_dump(by_alias=True)
    assert dumped["General"]["Split Temperature"] is True
    assert params.gen.split_temperature is True


def ptr_dict():
    """A classic getTestDataFromHeadSite PTR result: three DUT indices
    with matching data/flag lists."""
    return {
        "dutList": np.asarray([0, 1, 2]),
        "dataList": np.asarray([1.0, 2.0, 3.0]),
        "flagList": np.asarray([0, 0, 1]),
    }


class TestApplyDutExclusion:
    def test_ptr_dict_masks_all_lists(self):
        d = ptr_dict()
        masked = apply_dut_exclusion(d, 1, {1: {1}})
        assert masked["dutList"].tolist() == [0, 2]
        assert masked["dataList"].tolist() == [1.0, 3.0]
        assert masked["flagList"].tolist() == [0, 1]

    def test_mpr_state_list_masked_too(self):
        d = {
            "dutList": np.asarray([0, 1]),
            "dataList": np.asarray([1.0, 2.0]),
            "stateList": np.asarray([1, 0]),
            "flagList": np.asarray([0, 1]),
        }
        masked = apply_dut_exclusion(d, 1, {1: {1}})
        assert masked["stateList"].tolist() == [1]

    def test_none_exclusion_returns_same_object(self):
        d = ptr_dict()
        assert apply_dut_exclusion(d, 1, None) is d

    def test_unknown_fid_returns_same_object(self):
        d = ptr_dict()
        assert apply_dut_exclusion(d, 7, {1: {0}}) is d

    def test_all_excluded_empties_the_dict(self):
        d = ptr_dict()
        assert apply_dut_exclusion(d, 1, {1: {0, 1, 2}}) == {}


def full_stat_dict():
    """A classic getTestDataFromHeadSite PTR result with the PRECOMPUTED
    stat keys testDataProcessCore produces at fetch time — computed over
    the FULL population (that is exactly the C1 shape: the scalars
    predate the exclusion)."""
    data = np.asarray([10.0, 20.0, 80.0])
    mean, sdev, cpk = calc_cpk(0.0, 100.0, data)
    return {
        "dutList": np.asarray([0, 1, 2]),
        "dataList": data,
        "flagList": np.asarray([0, 1, 0]),
        "Unit": "V",
        "LLimit": 0.0,
        "HLimit": 100.0,
        "SUB_CODE": REC.PTR,
        "Cpk": cpk,
        "Mean": mean,
        "SDev": sdev,
        "Median": np.nanmedian(data),
        "Min": np.nanmin(data),
        "Max": np.nanmax(data),
    }


class TestExclusionRecomputesStats:
    """C1 (branch review): masking the raw lists never recomputed the
    precomputed stat keys — Cpk/Average/Median/St. Dev./Min/Max kept
    counting dropped DUTs while Fail Num did not (G3 violated)."""

    def make_interface(self):
        interface = DataInterface.__new__(DataInterface)
        interface.num_files = 1
        interface.testRecTypeDict = {(10, "0_TXCC/VCCC_ASIC"): REC.PTR}
        interface.getTestDataFromHeadSite = (
            lambda testTuple, heads, sites, fid: full_stat_dict()
        )
        return interface

    def test_scalar_stats_reflect_only_kept_duts(self):
        interface = self.make_interface()
        d = interface.getTestStatistics(
            [(10, 0, "0_TXCC/VCCC_ASIC")], [1], [-1], [1], dut_exclusion={1: {1}}
        )
        row = d["Rows"][0]
        # kept DUTs are [10.0, 80.0] — the dropped 20.0 (the failing flag)
        # must not count. Row scalars are strings through the app's float
        # format (1 decimal here): tolerance 0.051 covers the rounding.
        assert row[4] == "0"  # Fail Num over KEPT flags only
        assert abs(float(row[6]) - 45.0) < 0.051  # Average
        assert abs(float(row[7]) - 45.0) < 0.051  # Median
        assert abs(float(row[9]) - 10.0) < 0.051  # Min
        assert abs(float(row[10]) - 80.0) < 0.051  # Max
        mean, sdev, cpk = calc_cpk(0.0, 100.0, np.asarray([10.0, 80.0]))
        assert abs(float(row[8]) - sdev) < 0.051  # St. Dev.
        assert abs(float(row[5]) - cpk) < 0.051  # Cpk

    def test_no_exclusion_keeps_upstream_scalars(self):
        interface = self.make_interface()
        d = interface.getTestStatistics([(10, 0, "0_TXCC/VCCC_ASIC")], [1], [-1], [1])
        row = d["Rows"][0]
        assert abs(float(row[6]) - 110.0 / 3.0) < 0.051  # full-population mean
