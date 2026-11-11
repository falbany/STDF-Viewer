"""Tests for deps/exporter/helpers.py: parsing, chunking, formula builders."""

import re

import pytest

from deps.exporter.helpers import (
    DEFAULT_PARAM_PATTERN,
    DUT_CHUNK_CAP,
    TEMPERATURE_RE,
    cp_cached,
    cpk_cached,
    fails_cached_flags,
    fails_cached_limits,
    fails_cached_limits_one_sided,
    format_temperature_label,
    get_column_letter,
    has_numeric,
    numeric_values,
    stat_cached,
    stat_cp_value,
    stat_cpk_value,
    apply_cp_formula,
    apply_cpk_formula,
    build_stat_cp_formula,
    build_stat_cpk_formula,
    col_range,
    compile_param_pattern,
    cp_formula,
    cpk_formula,
    die_col,
    fails_formula_flags,
    fails_formula_limits,
    flag_value,
    metadata_stat_formulas,
    parse_bin,
    parse_temperature,
    parse_temperature_text,
    parse_test_time,
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


class TestParseTemperatureText:
    @pytest.mark.parametrize(
        ("value", "expected"),
        [
            ("25C", 25.0),
            ("25", 25.0),
            ("-40C", -40.0),
            ("25.5 °C", 25.5),
            ("+25C", 25.0),  # sign dropped by the regex, number still found
            ("#1 → 25C\n#2 → -40C", 25.0),  # composite: first subfile wins
            ("Set to 25C on chuck", 25.0),  # free text: first number wins
            ("", None),
            ("N/A", None),
            (None, None),
        ],
    )
    def test_spec_cases(self, value, expected):
        assert parse_temperature_text(value) == expected

    def test_composite_prefix_index_never_parses_as_number(self):
        # the "#1 →" source index must not leak into the parsed value
        assert parse_temperature_text("#2 → -40C") == -40.0

    def test_numeric_passthrough(self):
        assert parse_temperature_text(25) == 25.0
        assert isinstance(parse_temperature_text(25), float)

    def test_nan_becomes_none(self):
        assert parse_temperature_text(float("nan")) is None

    def test_returns_float(self):
        assert isinstance(parse_temperature_text("25C"), float)


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
        assert split_test_name("A/B/C", r"^(?P<test>.+)-(?P<param>[^-]+)$") == (
            "A/B/C",
            "",
        )


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
        # 13 + n + 3 + n <= 16384  ->  n <= (16384 - 13 - 3) // 2 = 8184
        # (Trend sparkline column + second spacer, xlsxwriter migration).
        assert DUT_CHUNK_CAP == (16384 - 13 - 3) // 2 == 8184
        # the arithmetic boundary: one DUT over the cap spills to chunk 2
        assert [len(c) for c in split_dut_chunks(list(range(8185)))] == [8184, 1]


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
        assert stat_formula(name, 17, "M", "AZ") == f'=IFERROR({name}(M17:AZ17),"")'

    def test_other_row(self):
        assert stat_formula("MIN", 18, "M", "N") == '=IFERROR(MIN(M18:N18),"")'


class TestCpFormula:
    """Ninth wave: Cp = (USL-LSL)/(6*sigma) with sigma embedded as STDEV over
    the row's die range (the St. Dev. column no longer exists to reference)."""

    def test_spec_string_row_17(self):
        assert cp_formula(17, "M", "AZ") == '=IFERROR((F17-E17)/(6*STDEV(M17:AZ17)),"")'

    def test_other_row_and_range(self):
        assert cp_formula(5, "M", "N") == '=IFERROR((F5-E5)/(6*STDEV(M5:N5)),"")'


class TestCpkFormula:
    """Ninth wave: same Cpk semantics, but sigma comes from STDEV over the
    die range (the signature gains the die range — the J column is now Cp)."""

    def test_spec_string_row_17(self):
        assert cpk_formula(17, "M", "AZ") == (
            "=IFERROR(MIN((F17-I17)/(3*STDEV(M17:AZ17)),"
            '(I17-E17)/(3*STDEV(M17:AZ17))),"")'
        )

    def test_other_row_and_range(self):
        assert cpk_formula(5, "M", "N") == (
            '=IFERROR(MIN((F5-I5)/(3*STDEV(M5:N5)),(I5-E5)/(3*STDEV(M5:N5))),"")'
        )

    def test_sigma_from_range_not_cell(self):
        # ninth wave divergence from build_stat_cpk_formula: the DUT Summary
        # embeds STDEV(range) because the J column no longer holds St. Dev.
        formula = cpk_formula(17, "M", "AZ")
        assert "STDEV(M17:AZ17)" in formula
        assert "J17" not in formula


class TestFailsFormulas:
    def test_limits_variant(self):
        assert (
            fails_formula_limits(17, "M", "AZ")
            == '=IFERROR(SUMPRODUCT((M17:AZ17<>"")*((E17<>"")*(M17:AZ17<E17)+(F17<>"")*(M17:AZ17>F17))),"")'
        )

    def test_flags_variant_over_mirror_range(self):
        assert (
            fails_formula_flags(17, "Q", "S")
            == '=IFERROR(SUMPRODUCT((Q17:S17<>"")*(Q17:S17=1)),"")'
        )

    def test_ftr_variant_removed(self):
        # third wave: FTR rows carry real limits and use the limits variant;
        # the FTR-specific formula was removed from helpers (grep: no refs)
        import deps.exporter.helpers as helpers_module

        assert not hasattr(helpers_module, "fails_formula_ftr")


class TestParseTestTime:
    """DB Test Time arrives as `TestTime || ' ms'` (schema.rs) — "N ms" strings."""

    @pytest.mark.parametrize(
        ("raw", "expected"),
        [
            ("123 ms", 123),
            ("123ms", 123),  # no space variant
            ("  45  ms ", 45),  # stray whitespace
            ("0 ms", 0),
            ("123", 123),  # already bare integer
            (123, 123),  # already numeric (defensive)
            (12.0, 12),
            ("", ""),
            (None, ""),
            ("ms", ""),  # no number at all
            ("garbage", ""),
        ],
    )
    def test_spec_cases(self, raw, expected):
        assert parse_test_time(raw) == expected

    def test_regex_fallback_extracts_first_integer(self):
        # odd formatting: strip-suffix path fails, the integer regex recovers
        assert parse_test_time("about 250 msecs") == 250

    def test_returns_int_type(self):
        assert isinstance(parse_test_time("42 ms"), int)


class TestParseBin:
    """DB HW/SW Bin arrives as `'Bin ' || HBIN` (schema.rs) — "Bin N" strings."""

    @pytest.mark.parametrize(
        ("raw", "expected"),
        [
            ("Bin 3", 3),
            ("Bin 0", 0),
            ("bin 7", 7),  # case-insensitive via regex
            ("Bin 127", 127),
            ("3", 3),  # bare integer
            (3, 3),  # already numeric (defensive)
            (0, 0),
            ("", ""),
            (None, ""),
            ("Bin", ""),  # no number at all
            ("garbage", ""),
        ],
    )
    def test_spec_cases(self, raw, expected):
        assert parse_bin(raw) == expected

    def test_extracts_first_integer(self):
        assert parse_bin("Hardware Bin: 5") == 5

    def test_returns_int_type(self):
        assert isinstance(parse_bin("Bin 2"), int)


class TestMetadataStatFormulas:
    def test_mean_true_returns_min_max_mean_triple(self):
        assert metadata_stat_formulas(17, "M", "AZ", mean=True) == (
            '=IFERROR(MIN(M17:AZ17),"")',
            '=IFERROR(MAX(M17:AZ17),"")',
            '=IFERROR(AVERAGE(M17:AZ17),"")',
        )

    def test_mean_false_drops_average(self):
        # third wave: only Temperature keeps a Mean cell — other numeric
        # metadata rows take the (MIN, MAX) pair and pad Mean with a dash
        assert metadata_stat_formulas(20, "M", "O", mean=False) == (
            '=IFERROR(MIN(M20:O20),"")',
            '=IFERROR(MAX(M20:O20),"")',
        )

    def test_std_true_appends_stdev_after_mean(self):
        # fourth wave: the Test Time row keeps Mean AND St. Dev. (sample
        # stdev, same STDEV semantic as the data rows' St. Dev. column)
        assert metadata_stat_formulas(21, "M", "O", mean=True, std=True) == (
            '=IFERROR(MIN(M21:O21),"")',
            '=IFERROR(MAX(M21:O21),"")',
            '=IFERROR(AVERAGE(M21:O21),"")',
            '=IFERROR(STDEV(M21:O21),"")',
        )

    def test_std_true_without_mean(self):
        # the helper stays general: STDEV without AVERAGE is legal even
        # though no metadata row uses that combination today
        assert metadata_stat_formulas(22, "M", "O", mean=False, std=True) == (
            '=IFERROR(MIN(M22:O22),"")',
            '=IFERROR(MAX(M22:O22),"")',
            '=IFERROR(STDEV(M22:O22),"")',
        )


# --------------------------------------------------------------------------
# Fifth user-feedback wave: live Test Statistics Cpk formula
# --------------------------------------------------------------------------


class TestBuildStatCpkFormula:
    """DUT-Summary-shaped Cpk formula with dynamic column letters (the
    Test Statistics columns shift with the Param / Pattern Name / MPR mix
    inserts, so the letters are parameters, never literals)."""

    def test_spec_string(self):
        # standard transformed header: SpecMin H, SpecMax I, Average L,
        # St. Dev. N, first data row 2
        assert build_stat_cpk_formula("H", "I", "L", "N", 2) == (
            '=IFERROR(MIN((I2-L2)/(3*N2),(L2-H2)/(3*N2)),"")'
        )

    def test_sigma_from_column_diverges_from_dut_formula(self):
        # ninth wave divergence: the Test Statistics sigma comes from the
        # (numeric) St. Dev. COLUMN, while the DUT Summary formula embeds
        # STDEV over the die range — the old letter-substitution equivalence
        # (build_stat_cpk_formula("E","F","I","J") == cpk_formula) is gone
        assert "3*N2" in build_stat_cpk_formula("H", "I", "L", "N", 2)
        assert "STDEV(" not in build_stat_cpk_formula("H", "I", "L", "N", 2)
        assert "STDEV(" in cpk_formula(2, "M", "O")

    def test_double_digit_rows(self):
        formula = build_stat_cpk_formula("A", "B", "C", "D", 123)
        assert "(B123-C123)" in formula
        assert "(C123-A123)/(3*D123)" in formula
        assert formula.endswith('),"")')


class TestApplyCpkFormula:
    """Overwrites the Cpk cell of a transformed stat row with the live
    formula; graceful no-ops for the degradation cases."""

    # standard transformed header indices: SpecMin 7, SpecMax 8, Cpk 10,
    # Average 11, Median 12, St. Dev. 13
    INFO = {
        "enabled": True,
        "spec_cols": (7, 8),
        "avg_col": 11,
        "std_col": 13,
        "cpk_col": 10,
    }

    def test_overwrites_cpk_cell_in_place(self):
        row = [0] * 14
        row[10] = "1.23"  # upstream-formatted string
        apply_cpk_formula(row, self.INFO, 2)
        assert row[10] == '=IFERROR(MIN((I2-L2)/(3*N2),(L2-H2)/(3*N2)),"")'

    def test_row_number_is_the_excel_row(self):
        row = [0] * 14
        apply_cpk_formula(row, self.INFO, 7)
        assert row[10] == '=IFERROR(MIN((I7-L7)/(3*N7),(L7-H7)/(3*N7)),"")'

    def test_disabled_transform_noop(self):
        info = dict(self.INFO, enabled=False)
        row = [0] * 14
        apply_cpk_formula(row, info, 2)
        assert row[10] == 0

    @pytest.mark.parametrize("key", ["spec_cols", "avg_col", "std_col", "cpk_col"])
    def test_missing_anchor_noop(self, key):
        # translated UI: any of the four anchors missing degrades to a no-op
        info = dict(self.INFO)
        info[key] = None
        row = [0] * 14
        apply_cpk_formula(row, info, 2)
        assert row[10] == 0

    def test_missing_anchor_keys_tolerated(self):
        # header_info dicts without the fifth-wave keys (older call sites)
        # degrade to a no-op, never a KeyError
        row = [0] * 14
        apply_cpk_formula(row, {"enabled": True}, 2)
        assert row[10] == 0

    def test_short_degraded_row_no_crash(self):
        # unparseable VHeader shrinks the row; a cpk_col off the end writes
        # nothing (no IndexError, no phantom cells)
        row = [0] * 5
        apply_cpk_formula(row, self.INFO, 2)
        assert row == [0] * 5


# --------------------------------------------------------------------------
# Ninth user-feedback wave: Test Statistics Cp formula (left of Cpk)
# --------------------------------------------------------------------------


class TestBuildStatCpFormula:
    """Cp = (USL-LSL)/(6*sigma) over one Test Statistics row, sigma taken
    from the numeric St. Dev. column (unlike the DUT Summary's embedded
    STDEV — the upstream St. Dev. values are numeric since wave 3)."""

    def test_spec_string(self):
        # SpecMin H, SpecMax I, St. Dev. O, first data row 2
        assert build_stat_cp_formula("H", "I", "O", 2) == '=IFERROR((I2-H2)/(6*O2),"")'

    def test_row_number_is_the_excel_row(self):
        assert build_stat_cp_formula("A", "B", "C", 123) == (
            '=IFERROR((B123-A123)/(6*C123),"")'
        )

    def test_argument_order_max_first(self):
        # the USL (SpecMax) leads, mirroring the Cpk formula's (max-avg) arm
        formula = build_stat_cp_formula("E", "F", "J", 17)
        assert formula == '=IFERROR((F17-E17)/(6*J17),"")'


class TestApplyCpFormula:
    """Overwrites the Cp cell of a transformed stat row with the live
    formula; graceful no-ops for the degradation cases."""

    # standard transformed header after the ninth-wave Cp insert:
    # SpecMin 7, SpecMax 8, Cp 10, Cpk 11, Average 12, Median 13, St. Dev. 14
    INFO = {
        "enabled": True,
        "spec_cols": (7, 8),
        "std_col": 14,
        "cp_col": 10,
    }

    def test_overwrites_cp_cell_in_place(self):
        row = [0] * 17
        row[10] = ""  # the transform's placeholder
        apply_cp_formula(row, self.INFO, 2)
        assert row[10] == '=IFERROR((I2-H2)/(6*O2),"")'

    def test_row_number_is_the_excel_row(self):
        row = [0] * 17
        apply_cp_formula(row, self.INFO, 7)
        assert row[10] == '=IFERROR((I7-H7)/(6*O7),"")'

    def test_disabled_transform_noop(self):
        info = dict(self.INFO, enabled=False)
        row = [0] * 17
        apply_cp_formula(row, info, 2)
        assert row[10] == 0

    @pytest.mark.parametrize("key", ["spec_cols", "std_col", "cp_col"])
    def test_missing_anchor_noop(self, key):
        # translated UI: any of the three anchors missing degrades to a no-op
        info = dict(self.INFO)
        info[key] = None
        row = [0] * 17
        apply_cp_formula(row, info, 2)
        assert row[10] == 0

    def test_missing_anchor_keys_tolerated(self):
        # header_info dicts without the ninth-wave key (older call sites)
        # degrade to a no-op, never a KeyError
        row = [0] * 17
        apply_cp_formula(row, {"enabled": True}, 2)
        assert row[10] == 0

    def test_short_degraded_row_no_crash(self):
        # a cp_col off the end of a degraded row writes nothing (no
        # IndexError, no phantom cells)
        row = [0] * 5
        apply_cp_formula(row, self.INFO, 2)
        assert row == [0] * 5


class TestGetColumnLetterXlsxwriter:
    """1-based wrapper over xlsxwriter's 0-based xl_col_to_name."""

    def test_single_columns(self):
        assert get_column_letter(1) == "A"
        assert get_column_letter(12) == "L"
        assert get_column_letter(13) == "M"

    def test_double_columns(self):
        assert get_column_letter(26) == "Z"
        assert get_column_letter(27) == "AA"
        assert get_column_letter(52) == "AZ"

    def test_matches_xlsxwriter_zero_based(self):
        from xlsxwriter.utility import xl_col_to_name

        assert get_column_letter(28) == xl_col_to_name(27)


class TestNumericValues:
    def test_keeps_numbers_and_numeric_strings_drops_rest(self):
        assert numeric_values([1, "2.5", None, "", "abc", 3.25]) == [1, 2.5, 3.25]

    def test_drops_bools_and_nan(self):
        assert numeric_values([True, False, float("nan"), 5.0]) == [5.0]

    def test_empty(self):
        assert numeric_values([None, ""]) == []


class TestStatCached:
    def test_min_max_average(self):
        assert stat_cached("MIN", [3.0, 1.0, 2.0]) == 1.0
        assert stat_cached("MAX", [3.0, 1.0, 2.0]) == 3.0
        assert stat_cached("AVERAGE", [1, 2, 3]) == 2.0

    def test_stdev_sample_semantics(self):
        assert stat_cached("STDEV", [1.0, 3.0]) == pytest.approx(2.0**0.5)

    def test_stdev_needs_two_values(self):
        assert stat_cached("STDEV", [2.0]) is None

    def test_empty_values(self):
        assert stat_cached("MIN", []) is None
        assert stat_cached("AVERAGE", [None, ""]) is None

    def test_unknown_name(self):
        assert stat_cached("MEDIAN", [1.0, 2.0]) is None


class TestCpCpkCached:
    def test_cp_symmetric_limits(self):
        # limits (0, 1) over [0.25, 0.75]: mu=0.5, sigma=sqrt(0.125)
        values = [0.25, 0.75]
        sigma = 0.125**0.5
        assert cp_cached(0.0, 1.0, values) == pytest.approx(1.0 / (6 * sigma))

    def test_cpk_symmetric_limits_both_ratios_equal(self):
        # limits (-1, 1) over [-1, 1]: mu=0, sigma=sqrt(2); both ratios equal
        sigma = 2.0**0.5
        expected = 1.0 / (3 * sigma)
        assert cpk_cached(-1.0, 1.0, [-1.0, 1.0]) == pytest.approx(expected)
        # and equals cp for the symmetric case
        assert cp_cached(-1.0, 1.0, [-1.0, 1.0]) == pytest.approx(expected)

    def test_cpk_takes_worse_side(self):
        # limits (0, 2) over [1.9, 2.1]: mu=2.0, sigma=sqrt(0.02)/... worse at max
        values = [1.9, 2.1]
        sigma = 0.02**0.5
        mu = 2.0
        expected = min((2.0 - mu) / (3 * sigma), (mu - 0.0) / (3 * sigma))
        assert cpk_cached(0.0, 2.0, values) == pytest.approx(expected)

    def test_none_cases(self):
        assert cp_cached(None, 1.0, [0.5, 0.6]) is None  # one-sided
        assert cp_cached(0.0, None, [0.5, 0.6]) is None
        assert cp_cached(0.0, 0.999999, [0.0, 0.0]) is None  # sigma == 0
        assert cp_cached(0.0, 1.0, [0.5]) is None  # n < 2
        assert cpk_cached(0.0, 1.0, []) is None


class TestFailsCached:
    def test_limits_exclusive_compare(self):
        assert fails_cached_limits([0.5, 1.2, ""], 0.0, 0.999999) == 1

    def test_limits_both_sides_counted(self):
        assert fails_cached_limits([-0.5, 0.5, 1.5], 0.0, 1.0) == 2

    def test_flags_count_ones(self):
        assert fails_cached_flags([0, 1, "", None, 1]) == 2


class TestStatColumnVariants:
    def test_stat_cp_value(self):
        assert stat_cp_value(0.0, 1.0, 0.5) == pytest.approx(1.0 / 3.0)

    def test_stat_cpk_value(self):
        assert stat_cpk_value(0.0, 1.0, 0.5, 0.5) == pytest.approx(1.0 / 3.0)

    def test_none_rules(self):
        assert stat_cp_value(None, 1.0, 0.5) is None
        assert stat_cp_value(0.0, None, 0.5) is None
        assert stat_cp_value(0.0, 1.0, 0.0) is None
        assert stat_cp_value(0.0, 1.0, None) is None
        assert stat_cpk_value(0.0, 1.0, 0.5, None) is None
        assert stat_cpk_value(0.0, 1.0, 0.5, 0.0) is None
        assert stat_cpk_value(0.0, None, 0.5, 0.5) is None


class TestHasNumeric:
    def test_true_and_false(self):
        assert has_numeric([None, "", 5]) is True
        assert has_numeric(["", None]) is False


class TestOneSidedFailsCache:
    """Final-review fix 2: a one-sided-limit row (has_limits is OR-semantics;
    the SUMPRODUCT guards each side with (E<>\"\")/(F<>\"\")) counts fails
    on the PRESENT side only — the cache must mirror that, not require both
    limits."""

    def test_low_limit_only_counts_below(self):
        # LL=1.0, no HL; dies [0.5, 2.0] -> 0.5 fails low; 2.0 has no HL to breach
        assert fails_cached_limits_one_sided([0.5, 2.0], 1.0, None) == 1

    def test_high_limit_only_counts_above(self):
        assert fails_cached_limits_one_sided([0.5, 2.0], None, 1.0) == 1

    def test_both_limits_present_same_as_before(self):
        assert fails_cached_limits_one_sided([0.5, 1.2, ""], 0.0, 0.999999) == 1

    def test_no_limits_falls_back_to_flags(self):
        # the matrix layer passes None/None -> the caller uses flags; the
        # helper itself reports everything inside an empty window (0 fails
        # on limits) so the CALLER decides the fallback
        assert fails_cached_limits_one_sided([0.5], None, None) == 0


class TestFormatTemperatureLabel:
    def test_int_temperatures_trim_the_decimal(self):
        assert format_temperature_label(25.0) == "25"
        assert format_temperature_label(-55.0) == "-55"

    def test_fractional_temperatures_keep_their_decimals(self):
        assert format_temperature_label(125.5) == "125.5"
        assert format_temperature_label(-5.5) == "-5.5"
