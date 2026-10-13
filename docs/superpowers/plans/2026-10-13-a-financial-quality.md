# A financial quality repair

**Goal:** Repair shared A generation using original concepts and official standalone periods; deliver a separate verified AMD candidate.

**Scope:** SOURCE_HANDOFF.md (2026-10-13). No database relocation/config changes, source rewrites, formal AMD A/B/C replacement, or removal of C corrections. Preserve pre-existing skill edits.

**Architecture:** Read existing cached statements. Select exact consolidated liability/cost totals and expose equity-method/continuing/discontinued concepts. Nonadditive shares/EPS must remain missing unless an official standalone source is available. An optional SEC source pack provides hash-verified quarterly GAAP earnings tables, matched by period end with provenance; no AMD-specific rule or values.

**Tech:** Existing Python/pandas/lxml/openpyxl pipeline, offline pytest, read-only cached audit.

- [x] Reproduce AMD cached pipeline and preserve before snapshots for AMD/AAPL/NVDA/MSFT/MRVL.
- [x] Add failing tests for nonadditive Q4, total-vs-component cost, liabilities/equity scope and bridge concepts.
- [x] Repair deterministic selection; block stale/fuzzy liability overrides. Run targeted regression; commit this unit only (f2745db).
- [x] Test and implement generic optional source-pack quarterly supplement, including hashes, issuer, quarter-vs-year, units, conflicts and missing values. Independent review findings reproduced and fixed.
- [x] Generate candidate with production writer. Independently verify 32 periods/389 cells, peer snapshots, 1,880 offline tests, native Excel 81 headers/0 errors and saved numeric values.
- [x] Update ARCHITECTURE/CLI/PITFALLS/CHANGELOG/TODO and evidence with exact coverage, remaining gaps and C handoff; only task files enter commits.
