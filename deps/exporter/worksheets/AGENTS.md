# AGENTS.md — the sheet-builder mixins

*Conventions for the worksheet modules in `deps/exporter/worksheets/`.
Global rules, architecture and the lint gates live in
[../AGENTS.md](../AGENTS.md) — read that first.*

---

## What these modules are

One mixin per sheet type. Each module defines a mixin class composed into
`ModernReportGenerator` (`generator.py`) and exposes **one public `write*`
entry point** (pylint's `min-public-methods=1` admits exactly this shape):

| module | entry point |
|---|---|
| `bin_chart.py` | `writeBinChart()` |
| `dut_summary.py` | `writeDUT_Summary()` |
| `file_info.py` | `writeFileInfo()` |
| `gdr_dtr.py` | `writeGDR_DTR()` |
| `histogram.py` | `writeHistogram(testTuple, head, site)` |
| `test_statistics.py` | `writeStatistic(testTuples, heads, sites)` |
| `trend_chart.py` | `writeTrendPlot(testTuple, head, site)` |
| `wafer_map.py` | `writeWaferMap()` |

(The names are camelCase by ruling — see the parent's Settled rulings.)

## Conventions

- **Host attributes come from `ReportHost`.** Every mixin inherits
  `ReportHost` (`../contract.py`), which declares — as PEP 526 annotations
  only — the workbook, theme, progress plumbing and selection state the
  host binds before any `write*` runs. A mixin never sets these itself.
- **New cross-mixin or host members get declared in `contract.py`.**
  Nothing else: the contract is annotation-only (`__slots__ = ()`), no
  runtime state, no methods. Adding a member here means adding its
  declaration there, or pyright and pylint cannot see it.
- **Styling goes through the theme** (`../theme.py`): the workbook-scoped
  format factory builds immutable Formats; per-cell styling composes at
  write time. Don't mutate a `Format`, don't hand-roll parallel format
  sets.
- **Charts are built by the `charts.py` builders** (donut factory, bin
  tables, KPI cards), not inline in the mixins — the mixin decides *what*
  to show and *where to anchor it*; `charts.py` decides *how* it renders.
- **Live formulas are sourced from table ranges.** Cells that show
  computed numbers carry Excel formulas referencing the table ranges
  (e.g. `=SUM`, `=COUNTIF`), so they recalculate when the user edits the
  workbook. Formulas are written in English syntax (see parent ruling).
- **Specs before pixels.** The DUT Summary sheet is built as a spec dict
  (`dut_summary.build_dut_summary_spec`) rendered by
  `writer.render_matrix_sheet`; plain tables go through
  `sheets.plain_table_spec` → `writer.render_plain_table`. Follow that
  split: mixins decide data, `writer.py` renders.

## Sheet-local gotchas

- **`localSheetId` follows creation order.** When a mixin splices or
  reorders `workbook.worksheets_objs` (the per-wafer sheets do), renumber
  `sheet.index` to match the new order — xlsxwriter writes the creation
  index into `localSheetId`, and skipping the renumber corrupts the
  workbook. This surgery happens in this directory, so this is where the
  rule bites.
- **`merge_range()` needs a ≥2-cell span.** Single cells go through
  `write()`. The KPI-card sites in `charts.py` carry the guard (commit
  `2020ec3`); new merge sites need one too.
- **Cross-mixin calls exist only by convention.** `test_statistics` calls
  `dut_summary`'s `_resolve_remeasure_drops` (its route to the shared
  supersede ranks); `dut_summary` and `wafer_map` call
  `reportGenerator._fid_ranks` directly. They work only because both
  sides inherit `ReportHost`, where those members are declared.

## Adding a sheet type

1. New module here, named after the sheet type, with one mixin class and
   one public `write*` entry point.
2. Add the mixin to the bases of `ModernReportGenerator` in
   `../generator.py` (composition root).
3. Inherit `ReportHost`; declare any new host/cross-mixin members in
   `../contract.py`.
4. New test file in `tests/exporter/` named after the module
   (`test_<module>.py`), fixtures from `tests/exporter/conftest.py`.
5. Run the three gates (parent's First hour) — all must stay clean.

## Pointer back

Global rules, architecture map, rulings and gates:
[../AGENTS.md](../AGENTS.md).
