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

from deps.DataInterface import apply_dut_exclusion
from deps.SharedSrc import SettingParams


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
