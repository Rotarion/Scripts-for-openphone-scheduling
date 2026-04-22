# Next Sprint Recommendations

1. Add a real smoke-run harness that can execute AHK scripts headlessly or semi-headlessly against a staging browser session.
2. Replace clipboard-based DevTools execution with a narrower bridge if the target browser environment allows it.
3. Externalize the remaining micro-delays inside clipboard and DevTools helpers once live replay data exists.
4. Add fixture coverage for more raw lead variants, especially malformed batch rows and multi-line labeled leads.
5. Introduce a small validation tool that diff-checks legacy vs refactored output for `BuildMessage`, `BuildFollowupQueue`, and `BuildBatchLeadRecord`.
6. Decide whether `run_state.json` should become authoritative for resume-state persistence across script restarts.
7. Audit the QUO tag-selector asset against current DOM snapshots to confirm the `PLUS_FALLBACK` path is still safe.
8. Retire or archive the legacy `Final_V5.*.ahk` chain once production parity is confirmed.
