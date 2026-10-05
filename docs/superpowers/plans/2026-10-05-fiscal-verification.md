# Fiscal verification and CF safety

Authorized scope: improve the measurement first, measure all cached companies in independent processes, verify affected periods, repair only evidence-backed period/CF behavior, and compare all outputs before/after. Raw database files must not be modified by the audit.

1. Add anomaly detection for unclassified dates, collisions, duplicate ends, fiscal ordering and CF fallback details. Tests must fail before implementation.
2. Run the actual fetch pipeline against cached input adapters (no network, no writes, no AI diagnosis), 80 quarterly/20 annual filings; checkpoint results per company. No synthetic second financial pipeline.
3. Prefer validated DEI document period/year/quarter from the already cached cover table, matching the exact selected period. Never apply filing focus to comparative columns. Preserve old behavior when authority is absent; report uncertainty.
4. Measure and review changes before adopting a correction. Update Python and Excel together; labels are arithmetic keys. CF fallback duration values may not masquerade as standalone; preserve balances, report gaps, guard derived values.
5. Regression tests, full non-slow suite, cached all-company comparison, representative workbook cell comparison, fresh-context code review, docs/ledger, local integration and push authorized work.

Decisions and evidence will be recorded in docs/superpowers/fiscal-verification-results.md. Deferred: G13(a) original facts recovery and product warning redesign A are separate scopes until period evidence is reliable.
