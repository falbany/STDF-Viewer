# Exporter performance design — export wall-clock (issue #11)

Date: 2026-11-17 · Branch: feature/xlsxwriter-exporter · Status: approved direction, awaiting plan

## Intent and why

Big DUT-block exports (Test Summary + DUT Summary, matrix-heavy files) take
too long to finish: measured **36.3 s at 2000 tests × 500 DUTs** and
**128.5 s at 4000 × 1000** on the synthetic-but-realistic harness
(`perf/profile_export.py`, untracked — drives the real
`ModernReportGenerator.report_core()` headlessly with stubbed `waitFor*`
fetches). The goal is roughly **2× wall-clock reduction** by doing less
work in the fork's own exporter code — no concurrency, no protocol
changes, no visual changes.

This design is measurement-driven: every lever below cites its measured
share of wall time. Re-run the harness after each wave to confirm.

## Measured ground truth (2026-11-17)

Per-pool wall-clock shares (harness, 2000×500 = 1M die cells / 4.04M
written cells / 36.3 s total; cProfile pass instrumented, 92.6 s):

| Pool | Share | Contents (cProfile cum) |
|---|---|---|
| per-cell write wrapper (ours) | **≈ 40 %** | `_write_cell` × 4.07M: `theme.cell_format` 12.9 s, `_note_width` 4.7 s, `sanitize_cell` 3.9 s + loop overhead |
| xlsxwriter internals | **≈ 34 %** | `write()` bookkeeping 13.3 s, close-time XML assembly 16.1 s, zlib 2.1 s |
| spec builders | **17 %** | 6.25 s real; `numeric_values` alone 10.3 s profiled (the ~7× row walk), stdev×2, fmean×2 |
| rezip (`complete_chart_flags`) | **0.4 %** | decompress-only path (no charts); re-deflate would cost 7 s at 120 MB if charts selected |
| misc render | ~5 % | `xl_col_to_name` per cell in the DUT-row loop (2.9 s), metadata re-parses |

Scaling is ~linear in die cells (4× cells → 3.5× time). Peak RAM 867 MB
→ 2.1 GB. Sheet dimensions verified: matrix row = 12 static + D die + 3
tail + D mirrors, **all cells through `_write_cell`** (no bulk path);
DUT Summary row = 15 attribute + T tests + 1 spacer + T mirrors likewise.
Micro-benchmarks: `isinstance(float) + math.isnan` is 5.4× faster than
today's `np.isnan` try/except (0.08 s vs 0.41 s per 1M cells);
`len(str(v))` costs 0.32 s per 1M cells (dead work on pinned sheets).

## Goals

1. ~2× wall-clock on DUT-block exports at the harness scales
   (36.3 s → ≤ 20 s at 2000×500; 128.5 s → ≤ 75 s at 4000×1000).
2. All fixes confined to fork-owned code (`deps/exporter/**`).
3. Every change independently verifiable; harness re-run per wave.

## Non-goals

- Trend/Histogram export path (finding 1) — upstream-inherited.
- Excel open/filter recalc cost (finding 6) and file weight (sparklines).
- The `waitFor*` handshake protocol — call counts may not change either
  (batching was only a Trend/Histo idea; nothing here touches fetches).
- STDF parsing / GUI-side import time.

## Anti-goals (a failure even if faster)

- ANY visual or byte-identity regression to rendered sheets: cell values,
  cached formula results, number formats, fills, borders, CF rules,
  sparklines, hidden columns/rows, widths — all pinned by the existing
  exporter test suite (782 tests) and the wave-verification rules.
- Threading or multiprocessing (measured: GIL-bound pure Python
  everywhere that matters; zlib releases the GIL but close()'s deflate is
  terminal; multiprocessing loses to summaryTable pickling under macOS
  spawn and fights the fork-merge constraint). Decided 2026-11-17.
- Changing what the user sees in Excel, including the hidden flag mirrors
  (they stay: CF rules and no-limit FAILS formulas reference them).

## Constraints

- Fork-merge constraint: new files + tiny hooks only upstream-side; this
  work lives entirely in `deps/exporter/**` (already fork-owned).
- Lint gates stay clean: ruff, pylint 10.00, pyright 0 errors.
- The 800-test exporter suite stays green; new behavior gets pin tests
  mirroring the existing style (openpyxl read-back + raw-XML asserts).
- User runs French Excel: nothing here changes formula text, but keep
  the rule in mind if any test fixtures touch formulas.

## Decided (user rulings, 2026-11-17)

- Pain dimension = export wall-clock only.
- Scope = the fork's exporter code only (levers below); threading
  rejected; `constant_memory` parked (see Deferred).
- Wave order by measured share: wrapper → bulk mirrors → builders →
  misc. **Wave 5 (streaming rezip) PARKED** — 0.4 % measured; revisit
  as a standalone effort if charts-selected large exports bite (see
  Deferred).
- **Split-temperature sheets consume** the Wave 3 shared per-test
  compute (sliced by kept pairs) — no export path stays on the slow
  walk.
- `perf/profile_export.py` is **committed to the fork** — every wave
  gate and success criterion stays reproducible from a clean checkout
  (new file only; no upstream conflict under the fork-merge constraint).

## The waves

### Wave 1 — `_write_cell` fast path (biggest lever, ≈ 40 % pool)

1. **Format hoisting**: `theme.cell_format` is memoized, but its cache
   KEY (font, align, border, fill_hex, num_fmt, extra) is rebuilt per
   cell — 4.07M times. In the two hot renderers
   (`writer.py _render_data_rows`, `render_dut_summary_sheet`'s DUT-row
   loop) the key varies only along a small bounded set
   (per-column format, per-row band/boundary/neutral fill). Resolve the
   Format once per (column × style-variant) BEFORE the row loop and pass
   the built Format straight into a fast write; exact mechanism (new
   `fmt=` fast path on `_write_cell` vs a sibling `_write_cell_fast`)
   is the builder's call. Requirement: `theme.cell_format` invocations
   drop from per-cell to a bounded per-sheet count; the memo in
   `theme.py` stays as-is for everyone else.
2. **Sanitize fast path**: `sanitize_cell`'s numeric branch becomes
   `isinstance(value, float) and math.isnan(value)` (`np.float64` is a
   `float` subclass, so numpy scalars keep working; `np.int64` can't be
   NaN). Behavior identical: NaN → `""`, strings still stripped via
   `ILLEGAL_CHARACTERS_RE`, None → `""`.
3. **`_note_width` skip on pinned sheets**: skip width noting wherever
   the ledger result is dead work — for DIFFERENT reasons per family
   (builder-check finding): matrix and DUT Summary ledger entries are
   never popped (only `apply_autofit` pops, and neither sheet calls
   it); wafer sheets DO pop via `_style_wafer_sheet`'s autofit
   (wafer_map.py:79) but the autofit widths are then overridden by the
   pinned widths — also dead work. Skipping notes on all three families
   is behavior-neutral; File Info and any genuinely-autofit sheet keeps
   today's behavior.
4. **`xl_col_to_name` hoist**: `render_dut_summary_sheet`'s DUT-row loop
   calls it per cell (writer.py:726) for a `test_formats` dict lookup —
   precompute the per-column format list once per sheet.

Expected: the ≈40 % wrapper pool drops to ≈12-15 %.

### Wave 2 — bulk-write the uniform mirror cells (≈ 2M cells)

Both sheets' flag mirrors are written cell-by-cell through the wrapper,
yet every mirror cell carries the IDENTICAL style (body/center/border,
no band — banding is gated on `fixed_column`, no fill, no num_fmt —
verified in both renderers). Write each mirror segment with one
`ws.write_row(...)` per row (matrix) / per DUT row (DUT Summary) and a
single prebuilt Format. `write_row` maps `""` to blank — same XML as
today. The matrix's 3-cell tail (`"", Trend, ""`) can ride along or stay
as-is (builder's call, must stay visually identical). Verification: the
existing hidden-mirror / CF-anchor tests must stay green unchanged.

Expected: ≈2M wrapper calls removed (−5-6 s at the 2000×500 scale).

### Wave 3 — builders: one pass, one sigma, one compute for both sheets

1. **One numeric pass per row**: `_test_rows` (test_summary.py) and
   `_test_column` (dut_summary.py) rebuild `numeric_values(die_cells)`
   inside `fails_cached_*`, MIN, MAX, AVERAGE, Cp, Cpk and
   `has_numeric` (7 list rebuilds + per-cell isinstance per row), and
   compute stdev twice + fmean twice (Cp and Cpk each). Build `nums`
   once per test, compute sigma/mu once. The cached-value gating must
   stay byte-identical: None → `""`, MIN/MAX faithful-0 on empty,
   one-sided OR semantics for `has_limits`, flag-mirror fallback only
   when both limits absent. Add a row-level helper in `helpers.py`
   (naming the builder's call) and keep the existing public helpers for
   other callers (metadata block).
2. **Shared compute between both sheets** (finding 4): both builders
   re-derive identical per-test `(die_cells, mirror_cells, six caches)`
   from the same pruned summaryTable + order. Compute once, consume
   twice. Plumbing follows the existing `_dut_ctx_*` host-context
   pattern: build the per-test table in `_prepare_dut_context` (it
   already owns the fetch-once context) and let both `build_*_spec`
   consume it; when absent (direct calls, tests) they compute inline as
   today. Placement concerns (matrix chunks by DUT columns, DUT Summary
   chunks by test columns) are slicing at consumption time — the shared
   compute stays per-test over the FULL ordered DUT list.
   **Split-temperature sheets consume it too** (2026-11-17 ruling): the
   split path calls `build_test_summary_spec` with a pruned subset table
   and `sub_order` (an order-preserving subset of the context order) —
   slice the shared table by the kept `(fid, dutIndex)` pairs instead of
   recomputing; split exports get pin tests proving byte-identity with
   today's inline computation.
3. **`dut2ind` hoist**: resolution is test-independent — resolve
   `(fid, dutIndex) → array_index` once per chunk/context, leaving
   per-cell work at `entry.get(fid)` + two list indexes.
4. **Metadata per-DUT precompute**: resolve `dutInfo[fid][dutIndex]`
   once per DUT (14 fields re-fetch it today), parse the Part-ID
   temperature regex and `part_id_base` once per DUT (2-3 passes today),
   shared by matrix metadata rows, secondary header and DUT Summary's
   `meta_columns`.

Expected: builders 17 % → ≈4-5 %; the both-sheets duplication
disappears on Test Summary + DUT Summary exports.

### Wave 4 — small verified items (do while in there)

- `stat_row_is_ftr` (test_statistics.py): precompute the Pattern-Name
  header index once per sheet instead of re-scanning the header per row.
- `write_wafer_map_block` (charts.py): `bins_present`/`counts`/`pf_map`/
  `kpi` walk the die grid ~5× with an O(bins × dies) nested count — one
  `Counter(grid.values())` pass feeding all four (≤10k dies per wafer,
  so this is hygiene more than seconds).

## Deferred (explicitly parked, not in this effort)

- **Streaming rezip** (parked 2026-11-17 ruling): `complete_chart_flags`
  (charts.py) decompresses EVERY zip member into RAM, then re-deflates
  all — though only `xl/charts/chartN.xml` changes, and its "nothing
  changed" exit happens after the full decompress. Measured 0.4 % of
  wall-clock on matrix exports, but it scales with file size (7 s
  re-deflate + 640 MB inflated at a 120 MB file IF charts are selected)
  and GB-scale RAM on huge reports. The fix when it is picked up:
  1. Read ONLY the tiny chart members first; if none change, return —
     chart-less exports then skip ALL zip work.
  2. When charts changed: stream-copy unchanged members
     (`zin.open(name)` → `zout` with the ORIGINAL `ZipInfo`),
     re-compress only the chart XMLs, `os.replace` as today. Peak
     memory becomes O(largest member).
- **xlsxwriter `constant_memory` mode**: the only lever on the ≈34 %
  xlsxwriter floor; moves XML generation into the write loop (roughly
  CPU-neutral, large RAM win) but breaks `apply_autofit`-after-write and
  `merge_range` ordering. Revisit only if peak RAM (2.1 GB at 4M die
  cells) becomes a user-reported pain.
- Batched chart rendering / protocol batching (Trend/Histo — upstream).

## What's left to the builder

- Exact fast-path API shape in `writer.py` (parameter vs sibling
  function) and format-precompute granularity per loop.
- The row-level stats helper's name/home in `helpers.py` and how the
  shared per-test table rides the host context (follow `_dut_ctx_*`).
- Whether the matrix's 3-cell tail rides Wave 2's `write_row`.
- Test additions pinning each wave (mirror the existing pin style).

## Success criteria

1. Harness: **~2× vs a fresh same-day baseline** re-run of
   `perf/profile_export.py` immediately before the first wave (machine
   load drifts, so the gate is relative; the 2026-11-17 absolutes —
   36.3 s at 2000×500, 128.5 s at 4000×1000 — are sanity references,
   targets ≤ 20 s / ≤ 75 s on this machine class).
2. Full exporter suite green; lint gates clean (ruff / pylint 10.00 /
   pyright 0).
3. Visual identity: no assertion changes to existing sheet-content,
   CF, sparkline, hidden-column or width tests; final acceptance is the
   user's live-export checklist on a real file (the established
   close-out ritual for exporter waves).
