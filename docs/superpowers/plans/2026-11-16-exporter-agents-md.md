# Exporter AGENTS.md Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Two new AGENTS.md files (`deps/exporter/AGENTS.md` + `deps/exporter/worksheets/AGENTS.md`) that onboard a fresh LLM agent to the exporter and forbid the settled-ruling and hard-won-gotcha mistakes.

**Architecture:** Pure documentation — no code changes. The parent file is the self-contained dossier (what the package is → architecture map → first hour → rulings → gotchas → pointers); the companion covers mixin conventions + sheet-local gotchas and points back for global rules. Every fact is verified against the repo during the fact-pass steps, never asserted from memory.

**Tech Stack:** Markdown; agents.md open format (overview → commands → workflow → testing → rules); repo conventions (IDEAS.md voice, conventional commits).

**Spec:** docs/superpowers/specs/2026-11-15-exporter-agents-md-design.md

## Global Constraints

- New files only: `deps/exporter/AGENTS.md` and `deps/exporter/worksheets/AGENTS.md`. No upstream-owned file is created, modified, or made to reference AGENTS.md (fork merge constraint).
- English prose, IDEAS.md voice (dated, direct, names hook points). No French.
- Section skeleton follows the agents.md open format.
- Soft length targets: parent ~150–250 lines, companion ~60–100.
- Rulings carry their dates; gotchas carry their commit reference where one exists (`2020ec3`, wave-6 fix).
- Every module one-liner and every command written into the files is verified against the repo first (fact-pass steps), not copied from memory.

## Review Focus

1. Wrong architecture claims (the "`sheets.py` is a sheet picker" class of misread) — a reader expects each module one-liner to match the module's actual docstring. → Pinned by Task 1 Steps 1+4.
2. Stale or broken commands — a reader expects every documented command (pytest collect, ruff, pylint, pyright) to run successfully today. → Task 1 Step 2.
3. Fork-constraint violation — no upstream/`.py` file may reference AGENTS.md. → Task 2 Step 3.
4. Design-history duplication — the files must point at `docs/superpowers/`, not retell it. → Task 2 Step 2.
5. French leakage — the only `SIERREUR` mention allowed is inside the formula ruling itself; all shown formula syntax is English (`IFERROR`, `,`). → Task 2 Step 3.

---

### Task 1: deps/exporter/AGENTS.md (parent dossier)

**Files:**
- Create: `deps/exporter/AGENTS.md`
- Read (fact pass): module docstrings under `deps/exporter/`, `pyproject.toml`, `tests/exporter/`

**Interfaces:**
- Produces: the parent rules file; Task 2's companion links back to it as `../AGENTS.md`.

- [ ] **Step 1: Fact pass — verify module one-liners against docstrings**

```bash
cd deps/exporter && for f in *.py worksheets/*.py; do echo "== $f"; awk '/^"""/{c++; if(c==2) exit} c>=1' $f | sed -n '2,6p'; done
```

Each one-liner written in Step 3 must be consistent with this output. The verified facts from the 2026-11-15 docstring pass:

- `generator.py` — composition root: `ModernReportGenerator` composes the worksheet mixins + `reportGenerator` + QObject; provides workbook/theme/progress plumbing and selection state; adapted from upstream `uic_stdExporter.py` (noonchen).
- `contract.py` — `ReportHost` typing-only host contract (PEP 526 annotations only, `__slots__ = ()`); the mixin classes' view of what the host guarantees. **Read first.**
- `worksheets/*.py` — one mixin per sheet type (bin_chart, dut_summary, file_info, gdr_dtr, histogram, test_statistics, trend_chart, wafer_map), each exposing a `write*` entry point; composed into `ModernReportGenerator`.
- `writer.py` — workbook-level assembly/ownership (~750 lines, largest non-chart module).
- `sheets.py` — plain-table spec builders (`plain_table_spec` → `render_plain_table`) + scaled image embedding (`insert_report_image`, Pillow probes). **Not a sheet picker — say so explicitly.**
- `charts.py` — native chart builders (~1700 lines: bin donuts, trend charts, histograms).
- `theme.py` — theme/styling. `helpers.py` — shared helpers.
- `dedup.py` — "Hide re-measured DUTs" keeps the latest measurement per serial number; returns a drop-set of `(fid, dutIndex)` pairs; touches no Rust core or GUI.
- `wafer_merge.py` — with re-measure-hide on, a WAFER_ID measured in 2+ loaded files gets one merged worksheet (die-by-die fusion, latest file's result superseding older).
- `worksheets/dut_summary.py` header comment says "matrix.py" — stale, harmless; do not propagate.

- [ ] **Step 2: Fact pass — verify commands and the test count**

```bash
uv run pytest tests/exporter --collect-only -q | tail -1   # live test count
uv run ruff check deps/exporter                            # expect clean
uv run pylint deps/exporter                                # expect 10.00/10
uv run pyright                                             # expect 0 errors
```

The file states the count with an as-of date ("NNN tests as of 2026-11-16").

- [ ] **Step 3: Write `deps/exporter/AGENTS.md`**

Sections in order:

1. **What this is** — fork-owned xlsxwriter report exporter for STDF Viewer; native charts + live formulas (cells recalculate when the user edits); replaces the upstream report library; adapted from `uic_stdExporter.py`. One line on audience (LLM agents).
2. **Architecture map** — the Step 1 one-liners, presented as the composition flow (generator composes the mixins); read-`contract.py`-first callout.
3. **First hour** — read `contract.py`; tests in `tests/exporter/` (per-module test files + `conftest.py`); the three gate commands from Step 2, "all must stay clean" (current state from Step 2).
4. **Settled rulings (dated)** — camelCase kept (upstream API mirror; ruff N unselected, pylint C0103 disabled — 2026-11-14); `ReportHost` stays annotation-only (2026-11-15); fork merge constraint (new files + tiny hooks, never rewrite upstream files in place); formula strings in English (`IFERROR`, `,`) regardless of the maintainer's French Excel (`SIERREUR`, `;`) — files are *written* in English formula syntax.
5. **Gotchas** — xlsxwriter writes the creation index into `localSheetId`: renumber `sheet.index` after any `worksheets_objs` surgery (corrupted-xlsx bug, wave 6); `merge_range()` needs a ≥2-cell span — single cells use `write()` (KPI-card sites guarded, commit `2020ec3`); cross-mixin couplings exist only in convention — `TestStatistics → DUTSummary._resolve_remeasure_drops`, both → `reportGenerator._fid_ranks` — declared in `contract.py`.
6. **Pointers** — `docs/superpowers/specs/` + `docs/superpowers/plans/` (design history: read, don't duplicate), `deps/exporter/IDEAS.md` (backlog), `pyproject.toml` (gate config).

- [ ] **Step 4: Verify the file**

- `wc -l deps/exporter/AGENTS.md` → within ~150–250 (soft target).
- Every path mentioned resolves: `test -e <path>` for each.
- Architecture check: re-diff the one-liners against the Step 1 docstring output; fix any mismatch (this pins Review Focus #1).

---

### Task 2: worksheets/AGENTS.md (companion) + commit

**Files:**
- Create: `deps/exporter/worksheets/AGENTS.md`
- Commit also: Task 1's parent file, the spec doc, this plan.

**Interfaces:**
- Consumes: `deps/exporter/AGENTS.md` from Task 1 (link target).
- Produces: the complete doc pair, committed.

- [ ] **Step 1: Write `deps/exporter/worksheets/AGENTS.md`**

Sections:

1. **Scope** — conventions for the sheet-builder mixins in this directory; global rules live in `../AGENTS.md`.
2. **Mixin conventions** — one module per sheet type; one public `write*` entry point (pylint `min-public-methods=1`); host attributes come from `ReportHost` (declare new ones in `contract.py`, PEP 526 annotations only); styling through `theme.py`; charts built through `charts.py` builders; live formulas sourced from table ranges so they recalculate on user edits.
3. **Sheet-local gotchas** — `sheet.index`/`localSheetId` renumbering (the `worksheets_objs` surgery happens here); 1-cell `merge_range` → `write()`; cross-mixin calls only via `ReportHost`-declared members.
4. **Adding a sheet type** — new module here → mixin composed into `ModernReportGenerator` (`generator.py`) → inherit `ReportHost` → declare host members in `contract.py` → new test file in `tests/exporter/` named after the module.
5. **Pointer back** to `../AGENTS.md`.

- [ ] **Step 2: Verify the pair**

- `wc -l deps/exporter/worksheets/AGENTS.md` → within ~60–100 (soft target).
- `test -e deps/exporter/AGENTS.md` (link target exists).
- No design-history retelling: no per-feature multi-paragraph narratives — pointers only (pins Review Focus #4).

- [ ] **Step 3: Fork-constraint + French grep**

```bash
grep -rln "AGENTS.md" --include="*.py" . | grep -v ".venv"   # expect empty — no .py file references it
grep -n "SIERREUR" deps/exporter/AGENTS.md deps/exporter/worksheets/AGENTS.md
# expect: exactly one hit, inside the formula ruling in the parent
```

- [ ] **Step 4: Commit (single commit — both files + docs)**

```bash
git add deps/exporter/AGENTS.md deps/exporter/worksheets/AGENTS.md \
        docs/superpowers/specs/2026-11-15-exporter-agents-md-design.md \
        docs/superpowers/plans/2026-11-16-exporter-agents-md.md
git commit -m "docs(exporter): add AGENTS.md onboarding for agents (parent + worksheets)"
```

Only these four paths — the working tree also carries unrelated untracked files (`.DS_Store`, `.superpowers/`, `.zcode/`, modified `cspell.json`) that must not be committed.
