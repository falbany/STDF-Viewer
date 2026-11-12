# Worksheets Package Split — Design (short)

Date: 2026-11-12
Status: approved approach (mixins), pending doc sign-off

## Intent & why

`deps/exporter/generator.py` (1358 lines) carries ~8 worksheet writers in one
file, plus their pure transforms. One file per worksheet type in a dedicated
`deps/exporter/worksheets/` folder improves maintainability: future waves
touch one sheet's file without wading through the others.

## Goals / non-goals

- **Goal:** pure code motion — zero behavior change, zero output change.
- **Goal:** every worksheet type gets one file; `matrix.py` joins the family
  as `worksheets/dut_summary.py` (it is the DUT Summary generator).
- **Non-goal:** no logic refactors, renames of public helpers, or signature
  changes beyond the move.
- **Anti-goal:** breaking the classic `reportGenerator` threading contract
  (`waitFor*` handshake, `forceQuit`, `loopCount`/`sendProgress` cadence).

## Approach

**Mixins** (picked over free-functions-with-delegation): each worksheet file
defines a small `XxxMixin` class holding its `write*` method(s); module-level
constants and pure transform functions stay module-level in the same file.
Rationale: the writers are tightly coupled to generator state (`self._wb`,
`forceQuit`, `waitFor*`) and cross-sheet handoffs (`_bin_banner`,
`_ftr_tests`, `_remeasure_drops` ride the instance) — delegation would add a
plumbing layer without decoupling anything. Mixins keep `report_core`'s
`self.write*()` calls and the test suite's `gen.write*()` calls untouched.

`ModernReportGenerator` lists the mixins **before** `reportGenerator` in its
bases so the modern overrides win the MRO:

```python
class ModernReportGenerator(
    FileInfoMixin, GdrDtrMixin, DUTSummaryMixin, TestStatisticsMixin,
    BinChartMixin, WaferMapMixin, TrendChartMixin, HistogramMixin,
    reportGenerator,
): ...
```

## File map

| New file | Contents (moved from generator.py unless noted) |
|---|---|
| `worksheets/__init__.py` | package marker |
| `worksheets/file_info.py` | `FILE_INFO_NUMERIC_LABELS`, `YIELD_FORMAT`, `_numeric_file_info_cell`, banner anchors, `bin_banner_from_file_info`, `file_temperatures_from_file_info`, `FileInfoMixin.writeFileInfo` |
| `worksheets/gdr_dtr.py` | `GdrDtrMixin.writeGDR_DTR` |
| `worksheets/dut_summary.py` | **all of `matrix.py`** (spec builder, FTR constants, `prune_summary_view`, `group_order_by_temperature`) + `order_dut_block` + `DUTSummaryMixin` (`writeDUT_Summary`, `_order_and_ids`, `_resolve_remeasure_drops`) |
| `worksheets/test_statistics.py` | stat transforms (`transform_stat_header/row`, `parse_vheader`, `numericize_value`, `stat_cf_ranges`, `stat_row_is_ftr`, `apply_ftr_pseudo_limits`, `_stat_full_test_name`, `_norm_header_text`), `STAT_*` constants, `theme_group_fill_hex`, `_is_number`, `_stat_formula_values` (as module fn), `TestStatisticsMixin.writeStatistic` |
| `worksheets/bin_chart.py` | `BinChartMixin.writeBinChart` |
| `worksheets/wafer_map.py` | `WaferMapMixin.writeWaferMap` |
| `worksheets/trend_chart.py` | `TrendChartMixin.writeTrendPlot` |
| `worksheets/histogram.py` | `HistogramMixin.writeHistogram` |

**Stays in `generator.py`:** `ModernReportGenerator` composition, `__init__`
instance state, shared helpers `_autofit_columns` / `_write_body_row`
(used by several sheets), `AUTOFIT_FACTOR/CAP`, and `report_core` (assembly,
sheet creation, tab-order reorder — importing `order_dut_block` from
`worksheets.dut_summary`).

## Import fallout (update, do not shim)

- `matrix.py` is **deleted**. `sanitize_cell` + `ILLEGAL_CHARACTERS_RE`
  move to `helpers.py` (pure utils): `charts.py`, `writer.py`,
  `sheets.py` repoint there. Rationale: `worksheets/dut_summary.py`
  hosts the DUT mixin, which imports `render_matrix_sheet` from
  `writer.py` — a `writer → dut_summary` import would then be circular.
- `uic_stdExporter.py`'s lazy `ModernReportGenerator` import is unchanged.
- Tests repoint: `deps.exporter.matrix` → `...worksheets.dut_summary`;
  `generator` imports of moved module functions → their new homes.

## Verification

Full `pytest tests/exporter` green; `python -c "from deps.exporter.generator
import ModernReportGenerator"` smoke; grep confirms no remaining
`deps.exporter.matrix` references.
