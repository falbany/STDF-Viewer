# AGENTS.md — the exporter package

*Fork-owned Excel report exporter for STDF Viewer. Written to orient LLM
agents working in `deps/exporter/`; humans are welcome too.*

---

## What this is

`deps/exporter/` is the fork's report exporter for STDF Viewer: it writes
the Excel workbook with **xlsxwriter**, builds **native Excel charts**
(donuts), and fills the tables with **live formulas** — cells recalculate
when the user edits the workbook. The package is adapted from upstream's
`deps/uic_stdExporter.py` (noonchen, GPL) and is fork-owned; the upstream
GUI code lives elsewhere in the repo and is not ours to rewrite (see
Settled rulings).

## Read first: `contract.py`

`contract.py` (85 lines) declares `ReportHost` — the typing-only contract
between the sheet mixins and their host. Every worksheet mixin inherits it,
so pyright and pylint see what the host guarantees: the workbook, theme,
progress plumbing, selection state, and cross-mixin methods. It carries no
runtime state: every member is a PEP 526 annotation (no values,
`__slots__ = ()`), so lookups still fall through to the real host in the
combined MRO. Read this file before anything else — it is the key to how
the mixins typecheck at all.

## Architecture map

Composition flow: `ModernReportGenerator` (`generator.py`) subclasses the
classic `reportGenerator` from `deps/uic_stdExporter.py`, keeps its
threading contract verbatim (the `generate_report` slot, the `waitFor*`
data handshake, the `forceQuit` early-exits, the `loopCount` /
`sendProgress` cadence), and composes **one mixin per sheet type**. The
mixins build *spec dicts* and call renderers. Nothing picks sheets from a
registry — the generator decides which `write*` mixins run.

| module | role |
|---|---|
| `generator.py` | Composition root. Owns the workbook, theme and progress plumbing and the selection state the mixins consume; subclasses the classic exporter and preserves its threading contract. |
| `contract.py` | `ReportHost` typing-only host contract (PEP 526 annotations, `__slots__ = ()`). **Read first.** |
| `worksheets/*.py` | One mixin per sheet type — `bin_chart`, `dut_summary`, `file_info`, `gdr_dtr`, `histogram`, `test_statistics`, `test_summary`, `trend_chart`, `wafer_map` — each exposing a `write*` entry point. Conventions: [worksheets/AGENTS.md](worksheets/AGENTS.md). |
| `writer.py` | Pure xlsxwriter renderer: Test Summary matrix sheets and plain tables, from spec dicts built by the mixins. Writes values (formulas verbatim with their cached results), styles (through the theme's workbook-scoped format factory), conditional formatting, native AutoFilter, hidden mirror columns, sparklines and print setup. Every data decision was made upstream of it. |
| `sheets.py` | Plain-table spec builder (`plain_table_spec`: headers translated, cells sanitized, generator rows materialized — feeding `writer.render_plain_table`) and scaled image embedding (`insert_report_image`: Pillow probes the PNG size, scale reaches `insert_image`, returns the sheet rows it spans). **Not a sheet picker** — nothing picks sheets here. |
| `charts.py` | Bin-distribution parsers, the native xlsxwriter donut-chart factory, the structured bin-table renderer and the dashboard KPI cards (headline bands, merged fills via one composed Format). |
| `theme.py` | The theme / workbook-scoped format factory. Formats are immutable — per-cell styling composes at write time; don't mutate a `Format`. |
| `helpers.py` | Pure parsing / formula / chunking helpers; the cached-value engine mirrors the sheet formulas for previews. |
| `dedup.py` | The optional "Hide re-measured DUTs" filter: keeps only the latest measurement per serial number, returns a drop-set of `(fid, dutIndex)` pairs. Touches no Rust core and no GUI. |
| `wafer_merge.py` | With re-measure-hide on, a real WAFER_ID measured in 2+ loaded files gets one extra **merged** worksheet: per-file payloads fused die-by-die, the latest file's result superseding the older ones. |

Rename note: the Test Summary module's first name was `matrix.py`
(records of it survive in writer.py comments and the git history); it is
now `worksheets/test_summary.py`, which builds the spec dicts
(`build_test_summary_spec` → `writer.render_matrix_sheet`).

## First hour

1. Read `contract.py`.
2. Skim the class statement of `ModernReportGenerator` in `generator.py` —
   the bases list *is* the sheet inventory.
3. Run the tests: they live in `tests/exporter/`, one test file per area
   plus `conftest.py` — the sheet mixins are covered by the
   `test_generator*.py` integration files and `test_test_summary.py`
   (the Test Summary module's first name was `matrix.py`), not by per-mixin files:

   ```bash
   uv run pytest tests/exporter        # 812 tests collected as of 2026-11-17
   ```

4. Run the three gates — **all must stay clean** after any change
   (configured in `pyproject.toml`, scoped to `deps/exporter`):

   ```bash
   uv run ruff check deps/exporter     # clean as of 2026-11-16
   uv run pylint deps/exporter         # 10.00/10 as of 2026-11-16
   uv run pyright                      # 0 errors as of 2026-11-16
   ```

## Settled rulings

Do not relitigate these; they were ruled by the maintainer:

- **camelCase identifiers stay** (2026-11-14). The package mirrors the
  upstream exporter API, whose names are camelCase. Ruff's pep8-naming
  rules are deliberately not selected and pylint's C0103 is deliberately
  disabled — do not "fix" the naming.
- **`ReportHost` stays annotation-only** (2026-11-15). PEP 526
  annotations, `__slots__ = ()`, no runtime state, no methods. Do not
  convert it into a runtime base class — the lookups must fall through
  the MRO to the real host.
- **Fork merge constraint** (standing). This fork must pull upstream
  cleanly: fork-owned work means **new files + tiny hooks**; never rewrite
  upstream files in place. `deps/exporter/`, `tests/exporter/` and the
  lint-gate config in `pyproject.toml` are fork-owned; the rest of the
  repo is upstream's.
- **Formula strings are written in English** (standing). New formulas go
  into files as English syntax — `IFERROR`, `,` separators — even though
  the maintainer's Excel runs in French and rule requests arrive as
  `SIERREUR` / `;`. Translate the request, not the file.

## Gotchas

Traps that already cost debugging waves — read before touching the
affected areas:

- **`localSheetId` follows creation order, not display order.** xlsxwriter
  writes the creation index into each sheet's `localSheetId`. If you
  reorder or splice `workbook.worksheets_objs` after creation (e.g.
  inserting per-wafer sheets), you must renumber `sheet.index` to match —
  skipping this corrupts the workbook (the wave-6 bug, 2026-11-14).
- **`merge_range()` needs a ≥2-cell span.** A 1-cell merge is invalid:
  single cells go through `write()`. The KPI-card merge sites carry a
  guard (commit `2020ec3`); any new merge site needs the same guard.
- **Cross-mixin couplings exist only by convention.**
  `test_statistics` calls `test_summary`'s `_resolve_remeasure_drops`
  (its route to the shared supersede ranks); `report_core` invokes
  `test_summary`'s `_prepare_dut_context()` once per export (the shared
  DUT-block context: `_dut_ctx_*`, `_ftr_tests`, `_remeasure_drops`);
  `test_summary` and `wafer_map` call `reportGenerator._fid_ranks`
  directly. These are real calls with no runtime base: they must be
  declared in `contract.py` (annotation-only) or pyright and pylint
  cannot see them. Any new
  coupling gets declared there too.

## Pointers

- `docs/superpowers/specs/` and `docs/superpowers/plans/` — design
  history: read for the *why* behind the current shape; don't duplicate it
  here.
- `deps/exporter/IDEAS.md` — feature backlog, ranked candidates.
- `pyproject.toml` — the three gate configurations (ruff / pylint /
  pyright) and their scoping.
- `tests/exporter/` — the test suite; `tests/exporter/conftest.py` holds
  the shared fixtures.
- [worksheets/AGENTS.md](worksheets/AGENTS.md) — conventions for the
  sheet-builder mixins, including how to add a sheet type.
