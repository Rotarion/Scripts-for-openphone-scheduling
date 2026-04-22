# Compatibility Note

## Preserved Behavior

- All 25 legacy hotkeys were preserved and re-registered.
- Current live settings were migrated from `..\time_rotation.ini` into `config/settings.ini` without changing values.
- The active tag-selector behavior still comes from the legacy `js for tag selector.js` implementation, now renamed to `assets/js/tag_selector.js`.
- `participant_input_focus.js` is still used before QUO phone pasting when available.
- Stable and fast batch modes still differ in scheduling strategy.
- Pricing values and the live single-car modern-year cutoff remain unchanged.
- The batch vehicle filter remains `MinVehicles=0` and `MaxVehicles=1`.
- Rotation offset still persists across runs.
- Batch logging still writes CSV rows with the same columns.
- Browser automation still prefers Chrome, then falls back to Edge.

## Intentional Structural Changes

- Rotation offset persistence now lives in `config/settings.ini` instead of `..\time_rotation.ini`.
- Runtime observability is mirrored into `logs/run_state.json`.
- Stable and fast batch logic now share one engine with explicit mode flags instead of duplicated functions.
- The refactor is isolated in a new folder so the old monolith remains untouched for comparison.

## Deferred For Safety

- QUO and CRM UI selectors were wrapped rather than deeply redesigned.
- Clipboard-driven DevTools execution was preserved as-is because it is fragile but business-critical.
- Form-fill field ordering and tab counts were preserved, not re-inferred.
