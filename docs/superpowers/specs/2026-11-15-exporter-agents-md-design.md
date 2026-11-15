# AGENTS.md files for the exporter — design

*2026-11-15 · small change · two new markdown files, no code*

## Intent

Two new AGENTS.md files so a fresh LLM agent working in `deps/exporter/` is
functional in the first hour and stops repeating settled-decision and
hard-won-gotcha mistakes. Today none of that knowledge is in-repo: the
rulings and xlsxwriter traps live only in `docs/superpowers/` history and
the maintainer's session memory — invisible to any agent (or human) that
starts cold. Both jobs were explicitly requested: onboarding AND prevention.

## Goals

- **Orientation** — what the package is, the corrected architecture map,
  first-hour workflow (read `contract.py` first; run tests; run the three
  lint gates).
- **Rules** — settled rulings with dates, phrased as do/don't: camelCase
  kept (upstream API mirror, 2026-11-14), ReportHost annotation-only
  contract, fork merge constraint (new files + tiny hooks), formula strings
  in English.
- **Gotchas** — the traps that already cost fix waves: xlsxwriter
  `sheet.index`/`localSheetId` renumbering after worksheets_objs surgery,
  1-cell `merge_range` guard, cross-mixin couplings that exist only in
  convention.
- **Pointers** — `docs/superpowers/` specs+plans for design history,
  `IDEAS.md` for the backlog.

## Non-goals

- No duplication of design history — AGENTS.md inlines *rules*, points at
  `docs/superpowers/` for the *why*.
- No root-level AGENTS.md this round, no CI wiring, no lint-gate changes
  (.md files aren't gated).
- Human-facing docs (README) untouched.

## Anti-goals (failure modes the files must prevent)

- Renaming camelCase identifiers "for style" → breaks the upstream API
  mirror.
- "Simplifying" `contract.py` into a runtime base class → breaks the
  annotation-only mixin design; pyright/pylint regress.
- Editing upstream-owned files in place → breaks clean upstream pulls.
- Writing French formula strings or assuming 1-row merges work → wrong
  formulas / corrupted workbooks.

## Constraints

- Fork merge constraint: AGENTS.md are new files → safe; they must not be
  referenced from any upstream-owned file.
- English, IDEAS.md voice; structure follows the agents.md open format the
  user's create-agentsmd skill uses (overview → setup → workflow → testing
  → rules).

## Approaches

1. **Self-contained dossier + thin sheet-builder companion (chosen)** —
   `deps/exporter/AGENTS.md` carries everything global;
   `worksheets/AGENTS.md` covers mixin conventions + sheet-local gotchas
   and points back for global rules. Matches how the knowledge actually
   splits; each file readable in one sitting.
2. *Pointer-only stubs* — rejected: the user explicitly wanted both jobs
   done in-file.
3. *One root AGENTS.md covering the whole repo* — rejected this round:
   scope creep; the exporter is the active work area.

## Decided

- **Locations:** `deps/exporter/AGENTS.md` + `deps/exporter/worksheets/AGENTS.md`
  (new files; no upstream file touched).
- **Parent outline:** ① what the package is (fork-owned xlsxwriter report
  exporter, native charts + live formulas, replaces upstream report lib)
  ② architecture map — `generator.py` = composition root
  (`ModernReportGenerator` composes the worksheet mixins + `reportGenerator`
  + QObject); `worksheets/*.py` = one mixin per sheet type exposing a
  `write*` entry point; `contract.py` = ReportHost typing-only host contract
  (read first); `writer.py` = workbook assembly; `sheets.py` = plain-table
  specs + image embedding (**not** a sheet picker — common misread);
  `charts.py` = native chart builders; `theme.py` styling; `helpers.py`
  shared; `dedup.py` / `wafer_merge.py` preprocessing ③ first hour —
  `tests/exporter/` (per-module test files + conftest;
  `uv run pytest tests/exporter`), gates `uv run ruff check deps/exporter`,
  `uv run pylint deps/exporter`, `uv run pyright` — all must stay clean
  ④ rulings with dates ⑤ gotchas ⑥ pointers.
- **Companion outline:** mixin conventions (write* entry point, host
  attributes guaranteed by ReportHost, styling via `theme.py`, charts via
  `charts.py` builders, live formulas sourced from table ranges);
  sheet-local gotchas (`sheet.index`/`localSheetId` renumbering, 1-cell
  `merge_range`, cross-mixin calls `TestStatistics →
  DUTSummary._resolve_remeasure_drops` and both → `reportGenerator._fid_ranks`);
  pointer back to the parent for global rules.
- One-liner descriptions of every module verified against the actual code
  while writing (docstring pass), not from memory.

## Left to the builder

- Final wording/length (target ~150-250 lines parent, ~60-100 companion).
- Commit placement: default is to commit alongside the next exporter work
  on the current branch, not as a lone commit (user can override).
