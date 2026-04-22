# Migration Plan

## Stage Execution

1. Reverse-engineered `..\Final_V5.5.ahk` and mapped hotkeys, parsing, scheduling, browser automation, JS assets, config, and logging.
2. Created a new root `Final_V5.5_refactored` so the legacy monolith remains available for side-by-side validation.
3. Extracted pure logic first into `domain/`.
4. Extracted orchestration into `workflows/`.
5. Wrapped browser, clipboard, DevTools, QUO, CRM, and tag behavior inside `adapters/`.
6. Re-registered all legacy hotkeys in `hotkeys/`.
7. Split config and assets into `config/`, `assets/js/`, and `logs/`.
8. Added lean test scripts in `tests/`.

## What Moved Where

| Legacy area | New file(s) |
| --- | --- |
| pricing and vehicle-year rules | `domain/pricing_rules.ahk` |
| follow-up queue and quote message text | `domain/message_templates.ahk`, `config/templates.ini` |
| business-day rules, holidays, rotation offset | `domain/date_rules.ahk`, `config/holidays_2026.ini`, `config/settings.ini` |
| raw/labeled/grid/batch parsing | `domain/lead_parser.ahk`, `domain/lead_normalizer.ahk`, `domain/batch_rules.ahk` |
| batch orchestration | `workflows/batch_run.ahk` |
| single lead create + tag | `workflows/single_lead_create.ahk` |
| follow-up scheduling workflow | `workflows/message_schedule.ahk` |
| prospect form fill workflow | `workflows/prospect_fill.ahk` |
| CRM activity workflow | `workflows/crm_activity.ahk` |
| clipboard restoration and stop-safe sleeps | `adapters/clipboard_adapter.ahk` |
| browser activation | `adapters/browser_focus_adapter.ahk` |
| DevTools JS bridge and CRM iframe helpers | `adapters/devtools_bridge.ahk` |
| Quo scheduling/new-chat interaction | `adapters/quo_adapter.ahk` |
| tag selector JS loader and tag application | `adapters/tag_selector_adapter.ahk`, `assets/js/tag_selector.js` |
| form fill and CRM appointment key sequences | `adapters/crm_adapter.ahk` |

## Consolidation Decisions

- Chosen authoritative tag selector asset: `assets/js/tag_selector.js`, normalized from legacy `js for tag selector.js`.
- Preserved `participant_input_focus.js` as a separate asset because the batch/new-chat flow still depends on it.
- Dropped `tag_activation.js` from the refactored tree because V5.5 no longer uses it.
- Consolidated fast and slow batch flows into one engine with a `mode` flag.
- Kept timing differences between stable and fast modes as explicit config-backed values instead of flattening them.
