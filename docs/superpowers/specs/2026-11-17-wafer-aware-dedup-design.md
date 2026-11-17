# Wafer-aware re-measure dedup — Design (small change)

- Date: 2026-11-17
- Status: implemented (see docs/superpowers/plans/2026-11-17-wafer-aware-dedup.md)
- Bug: with multiple wafers loaded, the "Hide re-measured DUTs" filter
  dedups by bare Part ID across all files; wafers reuse part IDs
  (per-die serials), so every wafer but the last is hidden as
  "re-measured".

## Intent and why

The re-measure filter (2026-11-10 design, §7.2) keeps only the latest
measurement per serial. Serial equality is only meaningful *within one
wafer* (or one wafer-less lot): the same part ID on two different
wafers is two different physical parts. The dedup must therefore match
on wafer identity first, Part ID second.

## Approach (picked in conversation, 2026-11-17)

**Wafer-aware matching key, Python-only; stat-only path repointed to
the full DUT-info fetch.** Rejected alternative: extending the Rust
partial fetch with the Wafer_Info join — touches upstream-owned
`rust_stdf_helper`, needs an extension rebuild, and risks the GUI Data
tab that consumes the same fetch. Both prior designs (2026-11-10,
2026-11-13) ruled out Rust changes.

## Decisions

1. **Matching key** — `compute_remeasure_drops` groups occurrences by
   `(WAFER_ID, Part ID)` instead of bare Part ID:
   - `WAFER_ID` is already in the full fetch's 11-slot dutInfo tuple at
     index 8 (Rust joins `Wafer_Info`); the dedup reads one more slot.
   - Rust returns `None` for a missing wafer id; `None` and `""` are
     the same bucket (no-wafer).
   - **No-wafer DUTs keep today's behavior**: they dedup by bare Part
     ID across files (cross-file lot retests still dedup). A wafer DUT
     and a no-wafer DUT never match each other.
   - Empty Part IDs still never match (unchanged).
   - Ordering (`(rank(fid), DUTIndex)`, `fid_ranks` override) and the
     "Superseded…" flag rule are unchanged.
2. **Stat-only path (Test Statistics without either DUT sheet)** —
   `getDutInfoForReport` (`STDF-Viewer.py:1539`, fork hook zone)
   repoints from `getPartialDUTInfoOnCondition` to
   `getDutIndexDictFromHeadSite` + `getFullDUTInfoFromDutArray` per
   fid — the same rows the heavy DUT-block fetch uses, wafer id and
   superseded DUTs included. Same source ⇒ identical drop-sets on both
   paths by construction (the 2026-11-10 §7.2 parity guarantee holds).
   Cost: a heavier query on the stat-only path, paid only when the
   filter is ON and no DUT sheet is selected.
3. **`dut_info_from_partial` becomes dead code** (its only consumer was
   the stat-only fallback) — removed from `dedup.py`, along with its
   tests in `tests/exporter/test_dedup.py`. `_resolve_remeasure_drops`
   passes the fetched dict straight to `compute_remeasure_drops`.

## Non-goals

- No Rust changes; no GUI changes; no change when the filter is OFF.
- No wafer-merge changes (already wafer-aware by WAFER_ID).
- No fuzzy wafer matching: exact string equality on WAFER_ID, same as
  Part ID.

## Tests

- `tests/exporter/test_dedup.py`: same Part ID on two different
  WAFER_IDs → neither drops; same Part ID, same WAFER_ID across fids →
  older drops (both rank modes); no-wafer DUTs keep bare-Part-ID dedup;
  None vs "" wafer id bucket together; wafer DUT vs no-wafer DUT never
  match; `dut_info_from_partial` tests removed.
- Stat-only fallback tests: drop-set equals the DUT-Summary-path
  drop-set for the same selection.

## Gates

`uv run pytest tests/exporter`, `ruff check deps/exporter`,
`pylint deps/exporter`, `pyright` — all must stay clean.
