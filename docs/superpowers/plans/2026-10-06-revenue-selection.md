# Revenue selection correction

Scope authorized by CTH: fix the template's Revenue selection using saved original concepts and independently verified totals; do not rewrite filing JSON. Source preservation boundary is in ARCHITECTURE.md. KR parser recovery is a separate task.

1. Baseline master 77e7805; isolated fix/revenue-selection worktree. Record original behavior failure, then implement total selection and preserve components in overflow.
2. Run full cached official-metadata pipeline on all 215 companies with the same adapter and listing metadata as the preserved financial baseline; disable source writes/network, AI diagnosis and companyfacts shares. Preserve rejected intermediate candidates separately.
3. Verify original SEC USD, no-dimension, exact duration examples: MAR, HLT, AMZN, MSFT, AFL, C and COST. Check calculation parent/subtotal distinctions instead of choosing maximum amounts.
4. Recalculate new owned Excel workbooks and check data/ratio export consistency. Run non-slow tests and reviewer; fix important findings before integration.
5. Verify all formal DB file hashes unchanged, update docs/TODO, commit each completed step, integrate and push only after merged tests pass.

Checkpoints: b76b97f RED tests; 867d1a0 basic total selection; bb28678 override/rebuild protection; d682ac6 calculation subtotal and missing-value correction; 575662c same-named overflow RED; 67e0a46 source template slot identity. Earlier revenue-candidate-all and revenue-final-all are intermediate only; the accepted all-company run uses 67e0a46 throughout.

Review rulings: Important override bypass fixed after three RED cases; recognized totals beat both inferred concept_override and structural_absence, custom legacy overrides remain when no recognized total exists. Full data exposed parent/child totals in COST/ADP/CAT; parent_concept establishes subtotal relation, while a parent explicitly including Other Income remains inconclusive. Synthetic cycle-plus-duplicate metadata edge was rated Minor with no cached example; document as further defensive work, not an observed source error or acceptance failure.

Deferred: original facts recovery/KR, custom totals with no recognized authority, scope decisions for ambiguous Revenue plus Other Income, GUI preview and specialized industry templates. Do not claim every bank/insurance company or every source cell has been certified by the selected source examples.

Late acceptance ruling: name-only merge classification promoted overflow named Revenue/Cash into fixed template rows and shifted BS/CF. Reviewer rated this Important and blocking; fixed by source slot boundaries (NG has zero template slots), then rerun all 215 from formal DB and rebuild/recalculate all seven workbooks. Preserve earlier full runs as rejected intermediate evidence. Fixed-row verifier reads source constants with AST, avoiding edgartools cache initialization. Existing excluding-tax NG routing and six partial-date override header cells remain explicitly deferred/inconclusive.
