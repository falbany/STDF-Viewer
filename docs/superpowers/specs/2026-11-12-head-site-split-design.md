# DUT Summary Head-Site Split — Design (short)

Date: 2026-11-12 (implemented 2026-11-13)
Status: implemented on `feature/xlsxwriter-exporter`, closes issue #1

## Intent & why

The DUT Summary sheet's combined **Head-Site** metadata row (Excel row 11,
`METADATA_ROWS[9]`, `dutinfo_index: 3`) renders the SQL-concatenated text
`"Head 1 - Site 4"` per DUT. Text-in-one-cell cannot be sorted or compared
numerically. Issue #1 asks for two metadata rows — **Head** and **Site** —
carrying bare integers so the user can sort the DUT columns by head or by
site.

## Decisions

**Parse in the exporter, no SQL/Rust change.** The issue offered two routes:
expose `HEAD_NUM`/`SITE_NUM` as two SQL columns, or parse the rendered text.
Parsing wins, on precedent and blast radius:

- X/Y already share one `dutinfo_index` (9) and parse the preformatted
  `"(x, y)"` string; Head/Site share index 3 and parse `"Head n - Site m"`
  the same way (`_HEAD_SITE_RE`). HW/SW Bin / Test Time likewise parse
  DB-shaped strings ("Bin 1" / "123 ms").
- The SQL→`DutSummaryRow`→py-api→`dutInfo` chain stays byte-identical: the
  GUI DUT Data Table keeps its combined text, the re-measure partial fetch
  and dedup are untouched, and upstream-shared files (schema.rs, fetcher.rs,
  py_api) see no diff — fork-pull-friendly.

**Mask policy: `MASK_CONSTANT` for both rows** (X/Y policy, replacing the
old unconditional `MASK_ALWAYS`). Single-head/single-site lots — the common
case — fold both rows and keep the block compact; the numeric values only
add information when they vary, which is exactly when the rows stay visible.
Blank parses count as nothing, so an all-blank row folds too (all-blank is a
constant of nothing).

**Layout: computed, never hard-coded.** The block grows 13 → 14 rows, so
everything below shifts (Tests Executed 13, Test Time 14, DUT Flag 15;
repeated header 16, data 17+):

- `dut_summary.py`: `HEADER_ROW = METADATA_START + len(METADATA_ROWS)`,
  `DATA_START = HEADER_ROW + 1` — placed after the tuple, beside the already
  computed `DUT_FLAG_ROW`/`PART_ID_ROW` (those self-adjust unchanged).
- `writer.py`: drops its mirrored `HEADER_ROW=15`/`DATA_START=16` constants
  for `_header_row(spec)` / `_data_start(spec)` — derived from
  `len(spec["metadata_rows"])` at render time, so renderer and spec builder
  can never drift apart again.

**Head/Site are numeric rows**: `unit_type: "INT"`, live MIN/MAX formulas +
cached values like HW/SW Bin (semantically odd numbers, but harmless and
consistent — every numeric row gets the same stat treatment). Head/Site rows
keep `mean/std: False` → dash cells in Mean/St. Dev.

## Out of scope

- The GUI DUT Data Table keeps the combined "Head h - Site s" column.
- The rationale's "filter the Data worksheet" lands as sortability: the
  metadata rows sit above the AutoFilter, so Excel row-wise filtering can't
  reach them; numeric cells make left-to-right column sorts (and formulas)
  over the die range meaningful instead.

## Tests

Stub `dutInfo` tuples switch to the real DB shape (`head_site="Head 1 - Site
1"` — the old `"1--1"`/`"1-1"` stubs never matched what the DB emits).
Updated: `test_matrix.py` (14-entry constant pin, Head/Site parse + mask
tests, every shifted formula string), `test_writer.py` (14-name fixture,
shifted layout constants), `test_generator_dut.py` (stub + hidden-row
assertions), `test_dedup.py` (fixture comment only — dedup reads indices
1/10).
