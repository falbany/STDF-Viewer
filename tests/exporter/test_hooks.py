"""Tests for Modern Export hooks: persisted settings field (Task 9).

GeneralConfig gains `modern_export` (alias "Modern Export", default True)
so the export dialog's "Modern Report" checkbox state survives restarts.
The UI/uic wiring is verified by import checks plus the full suite; these
tests are pure config, no Qt/GUI required.
"""

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
