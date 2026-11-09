"""Tests for deps/modernExporter/helpers.py: parsing, chunking, formula builders."""

import re

import pytest

from deps.modernExporter.helpers import (
    DEFAULT_PARAM_PATTERN,
    DUT_CHUNK_CAP,
    TEMPERATURE_RE,
    col_range,
    compile_param_pattern,
    cpk_formula,
    die_col,
    fails_formula_flags,
    fails_formula_ftr,
    fails_formula_limits,
    flag_value,
    metadata_minmax_formulas,
    parse_temperature,
    part_id_base,
    split_dut_chunks,
    split_test_name,
    stat_formula,
    workflow_key,
)


class TestParseTemperature:
    @pytest.mark.parametrize(
        ("part_id", "expected"),
        [
            ("SN541_T25C", 25.0),
            ("SN541_T-5C", -5.0),
            ("SN541_T85C", 85.0),
            ("SN541_T-5.5C", -5.5),
            ("SN541_t85c", 85.0),  # IGNORECASE
            ("SN541", None),
            (None, None),
        ],
    )
    def test_spec_cases(self, part_id, expected):
        assert parse_temperature(part_id) == expected

    def test_multi_match_takes_last(self):
        assert parse_temperature("SN541_T25C_T85C") == 85.0

    def test_returns_float(self):
        assert isinstance(parse_temperature("SN541_T25C"), float)

    def test_module_level_regex(self):
        assert isinstance(TEMPERATURE_RE, re.Pattern)
        assert TEMPERATURE_RE.flags & re.IGNORECASE


class TestPartIdBase:
    @pytest.mark.parametrize(
        ("part_id", "expected"),
        [
            ("SN541_T25C", "SN541"),
            ("SN541", "SN541"),
        ],
    )
    def test_spec_cases(self, part_id, expected):
        assert part_id_base(part_id) == expected

    def test_strips_last_match_only(self):
        assert part_id_base("SN541_T25C_T85C") == "SN541_T25C"


class TestParamPattern:
    def test_default_pattern_matches_test_slash_param(self):
        match = DEFAULT_PARAM_PATTERN.search("A/B/C")
        assert match.group("test") == "A/B"
        assert match.group("param") == "C"

    def test_compile_none_gives_default(self):
        assert compile_param_pattern(None) == DEFAULT_PARAM_PATTERN

    def test_compile_invalid_falls_back_to_default(self):
        assert compile_param_pattern("([unclosed") == DEFAULT_PARAM_PATTERN

    def test_compile_valid_pattern(self):
        pattern = compile_param_pattern(r"^(?P<test>.+)-(?P<param>[^-]+)$")
        assert pattern.search("A-B") is not None


class TestSplitTestName:
    def test_spec_case_nested(self):
        assert split_test_name("A/B/C") == ("A/B", "C")

    def test_spec_case_no_match(self):
        assert split_test_name("ABC") == ("ABC", "")

    def test_custom_pattern(self):
        pattern = r"^(?P<test>.+)-(?P<param>[^-]+)$"
        assert split_test_name("A-B", pattern) == ("A", "B")

    def test_custom_pattern_no_match_returns_full(self):
        assert split_test_name("A/B/C", r"^(?P<test>.+)-(?P<param>[^-]+)$") == ("A/B/C", "")


class TestWorkflowKey:
    def test_first_segment(self):
        assert workflow_key("A/B/C") == "A"

    def test_no_slash_is_none(self):
        assert workflow_key("ABC") is None


class TestFlagValue:
    def test_negative_is_none(self):
        assert flag_value(-1) is None

    def test_zero_is_pass(self):
        assert flag_value(0) == 0

    def test_fail_bit_is_fail(self):
        # isPass (SharedSrc.py:306) fails only on the bit7=1, bit6=0 pattern (0x80).
        assert flag_value(0x80) == 1

    def test_no_pass_fail_info_treated_as_pass(self):
        # SharedSrc comment: "x1 none, treated as pass" — 0x40 alone is not a fail.
        assert flag_value(0x40) == 0
        assert flag_value(0xC0) == 0


class TestSplitDutChunks:
    def test_empty_gives_single_empty_chunk(self):
        assert split_dut_chunks([]) == ((),)

    def test_small_list_single_chunk(self):
        ids = ["a", "b", "c"]
        assert split_dut_chunks(ids) == (("a", "b", "c"),)

    def test_exact_cap_single_chunk(self):
        ids = [str(i) for i in range(DUT_CHUNK_CAP)]
        chunks = split_dut_chunks(ids)
        assert len(chunks) == 1
        assert chunks[0] == tuple(ids)

    def test_one_over_cap_two_chunks(self):
        ids = list(range(DUT_CHUNK_CAP + 1))
        chunks = split_dut_chunks(ids)
        assert len(chunks) == 2
        assert len(chunks[0]) == DUT_CHUNK_CAP
        assert len(chunks[1]) == 1
        assert chunks[0][0] == 0
        assert chunks[0][-1] == DUT_CHUNK_CAP - 1
        assert chunks[1][0] == DUT_CHUNK_CAP

    def test_chunks_are_consecutive(self):
        ids = list(range(DUT_CHUNK_CAP * 2 + 5))
        chunks = split_dut_chunks(ids)
        assert [len(c) for c in chunks] == [DUT_CHUNK_CAP, DUT_CHUNK_CAP, 5]
        flat = [x for c in chunks for x in c]
        assert flat == ids

    def test_cap_derivation_fits_excel_column_limit_with_mirrors(self):
        # A chunk sheet is 13 base columns + n dies + 1 spacer + n hidden
        # mirrors, all within Excel's 16384-column cap:
        # 13 + n + 1 + n <= 16384  ->  n <= (16384 - 13 - 1) // 2 = 8185.
        assert DUT_CHUNK_CAP == (16384 - 13 - 1) // 2 == 8185
        # the arithmetic boundary: one DUT over the cap spills to chunk 2
        assert [len(c) for c in split_dut_chunks(list(range(8186)))] == [8185, 1]


class TestColumns:
    def test_die_col_first_is_m(self):
        assert die_col(0) == "M"

    def test_die_col_index(self):
        assert die_col(1) == "N"
        assert die_col(15) == "AB"

    def test_col_range(self):
        assert col_range("M", "AZ", 17) == "M17:AZ17"


class TestStatFormula:
    @pytest.mark.parametrize("name", ["MIN", "MAX", "AVERAGE", "STDEV"])
    def test_spec_names(self, name):
        assert stat_formula(name, 17, "M", "AZ") == f"={name}(M17:AZ17)"

    def test_other_row(self):
        assert stat_formula("MIN", 18, "M", "N") == "=MIN(M18:N18)"


class TestCpkFormula:
    def test_spec_string_row_17(self):
        assert cpk_formula(17) == '=IFERROR(MIN((F17-I17)/(3*J17),(I17-E17)/(3*J17)),"")'

    def test_other_row(self):
        assert cpk_formula(5) == '=IFERROR(MIN((F5-I5)/(3*J5),(I5-E5)/(3*J5)),"")'


class TestFailsFormulas:
    def test_limits_variant(self):
        assert (
            fails_formula_limits(17, "M", "AZ")
            == '=SUMPRODUCT((M17:AZ17<>"")*((E17<>"")*(M17:AZ17<E17)+(F17<>"")*(M17:AZ17>F17)))'
        )

    def test_flags_variant_over_mirror_range(self):
        assert (
            fails_formula_flags(17, "Q", "S")
            == '=SUMPRODUCT((Q17:S17<>"")*(Q17:S17=1))'
        )

    def test_ftr_variant_over_visible_range(self):
        assert (
            fails_formula_ftr(17, "M", "AZ")
            == '=SUMPRODUCT((M17:AZ17<>"")*(M17:AZ17=1))'
        )


class TestMetadataMinMaxFormulas:
    def test_returns_min_max_tuple(self):
        assert metadata_minmax_formulas(17, "M", "AZ") == ("=MIN(M17:AZ17)", "=MAX(M17:AZ17)")

    def test_other_row(self):
        assert metadata_minmax_formulas(20, "M", "O") == ("=MIN(M20:O20)", "=MAX(M20:O20)")
