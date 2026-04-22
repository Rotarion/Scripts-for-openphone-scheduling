# Responsibility Map

## Legacy Script Baseline

Active monolith analyzed: `..\Final_V5.5.ahk`

The legacy script mixed the following concerns in one file:

| Concern | Legacy evidence | Notes |
| --- | --- | --- |
| Global config and runtime state | lines 10-74 | INI reads, holidays, pricing, pacing, batch state, stop flags |
| Hotkey entrypoints | lines 116-964 | 25 hotkeys plus F8 spam loop |
| Persistence and logging | lines 982-1055 | INI defaults, batch CSV logging, rotation offset |
| Clipboard and input helpers | lines 1066-1179 | Clipboard mutation, paste helpers, tabbing, dropdown selection |
| UI picker flow | lines 1203-1337 | Follow-up batch picker and batch message selection |
| DevTools/browser bridge | lines 1338-1558 | Clipboard-driven JS execution, browser focus, Quo/CRM DOM helpers |
| Lead parsing and normalization | lines 1579-2584 | Raw lead, labeled lead, batch grid, address, DOB, phone, email, gender normalization |
| Tag selector bridge | lines 2585-2718 | Active JS selector loader plus tag application workflow |
| Date and scheduling rules | lines 2719-2895 | Holidays, business-day math, initial quote datetime, message scheduling |
| Message and pricing rules | lines 2896-3025 | Sales message composition, price resolution, follow-up queue generation |
| Batch lead shaping | lines 3026-3242 | Batch row parsing, vehicle extraction, holder creation |
| Batch orchestration | lines 3243-3546 | Stable and fast batch flows plus builder/follow-up scheduling |
| Prospect form fill | lines 3625-3746 | Edge and National General keystroke automations |
| Safety/debug utilities | lines 116-164, 453-517, 3747-3758 | stop handling, parsers preview, spam loop |

## Hotkey Surface

| Hotkey | Legacy role | New module |
| --- | --- | --- |
| `Esc` | global stop | `hotkeys/debug_hotkeys.ahk` |
| `F1` | exit | `hotkeys/debug_hotkeys.ahk` |
| `^!g` | tag selector debug | `hotkeys/debug_hotkeys.ahk` |
| `^!r` | reset rotation offset | `hotkeys/schedule_hotkeys.ahk` |
| `^!c` | active window controls debug | `hotkeys/debug_hotkeys.ahk` |
| `^!\`` | update agent config | `hotkeys/debug_hotkeys.ahk` |
| `^!u` | quick single lead create + tag | `hotkeys/lead_hotkeys.ahk` |
| `^!1` | ad-hoc quote message paste | `hotkeys/lead_hotkeys.ahk` |
| `^!6` | fast follow-up scheduling | `hotkeys/schedule_hotkeys.ahk` |
| `^!7` | stable follow-up scheduling | `hotkeys/schedule_hotkeys.ahk` |
| `^!8` | follow-up batch picker | `hotkeys/schedule_hotkeys.ahk` |
| `^!d` | follow-up preview | `hotkeys/schedule_hotkeys.ahk` |
| `^!0` | National General form fill | `hotkeys/crm_hotkeys.ahk` |
| `^!9` | Edge new prospect form fill | `hotkeys/crm_hotkeys.ahk` |
| `^!m` | DOB/ZIP debug | `hotkeys/debug_hotkeys.ahk` |
| `^!p` | parsed prospect preview | `hotkeys/debug_hotkeys.ahk` |
| `^!]` | raw labeled map preview | `hotkeys/debug_hotkeys.ahk` |
| `^!l` | batch holder preview | `hotkeys/debug_hotkeys.ahk` |
| `^!b` | stable batch run | `hotkeys/lead_hotkeys.ahk` |
| `^!n` | fast batch run | `hotkeys/lead_hotkeys.ahk` |
| `^!t` | last configured follow-up CRM preset | `hotkeys/crm_hotkeys.ahk` |
| `^!y` | tomorrow morning CRM preset | `hotkeys/crm_hotkeys.ahk` |
| `^!k` | CRM attempted-contact + appointment | `hotkeys/crm_hotkeys.ahk` |
| `^!j` | CRM quote-call + appointment | `hotkeys/crm_hotkeys.ahk` |
| `F8` | spam loop toggle | `hotkeys/debug_hotkeys.ahk` |

## Dependency Map

The new dependency direction is:

`hotkeys -> workflows -> domain + adapters`

`workflows -> domain`

`workflows -> adapters`

`adapters` do not call `workflows`

`domain` does not call `adapters`

Concrete module roles:

| Module | Primary dependencies |
| --- | --- |
| `workflows/batch_run.ahk` | `domain/batch_rules.ahk`, `domain/message_templates.ahk`, `domain/date_rules.ahk`, `adapters/quo_adapter.ahk`, `adapters/tag_selector_adapter.ahk` |
| `workflows/message_schedule.ahk` | `domain/message_templates.ahk`, `domain/date_rules.ahk`, `adapters/quo_adapter.ahk` |
| `workflows/prospect_fill.ahk` | `domain/lead_parser.ahk`, `adapters/browser_focus_adapter.ahk`, `adapters/crm_adapter.ahk` |
| `workflows/crm_activity.ahk` | `domain/date_rules.ahk`, `domain/message_templates.ahk`, `adapters/crm_adapter.ahk` |
| `adapters/tag_selector_adapter.ahk` | `adapters/devtools_bridge.ahk`, `adapters/quo_adapter.ahk`, `adapters/clipboard_adapter.ahk` |
| `domain/lead_parser.ahk` | `domain/lead_normalizer.ahk`, `domain/batch_rules.ahk` |

## Extraction Map

| Legacy responsibility | New destination |
| --- | --- |
| INI-backed settings and runtime bootstrap | `main.ahk`, `config/*.ini`, `logs/run_state.json` |
| Pure lead cleanup, names, address, DOB, phone, email | `domain/lead_normalizer.ahk` |
| Lead parsing from labeled/raw/grid input | `domain/lead_parser.ahk` |
| Batch row parsing and vehicle/tag holder shaping | `domain/batch_rules.ahk` |
| Pricing resolution | `domain/pricing_rules.ahk` |
| Quote/follow-up message composition | `domain/message_templates.ahk` |
| Business-day and rotation math | `domain/date_rules.ahk` |
| Clipboard pacing and low-level paste helpers | `adapters/clipboard_adapter.ahk` |
| Browser activation | `adapters/browser_focus_adapter.ahk` |
| DevTools JS execution and CRM DOM helpers | `adapters/devtools_bridge.ahk` |
| Quo composer/new-chat/schedule primitives | `adapters/quo_adapter.ahk` |
| CRM form fill and appointment UI sequences | `adapters/crm_adapter.ahk` |
| Active tag selector flow | `adapters/tag_selector_adapter.ahk` + `assets/js/tag_selector.js` |

## Dead Or Duplicate Candidates

Evidence-based candidates from `Final_V5.5.ahk`:

| Candidate | Evidence | Disposition |
| --- | --- | --- |
| `SendSelectedBatch(picker)` | defined at line 1203, no active caller in V5.5 | removed from refactor, replaced by active `SendSelectedBatchV2` path |
| `ClickAddNewAppointmentJS()` | defined at line 1477, no active caller in V5.5 | removed from refactor |
| `tag_activation.js` | variable declared at line 45, no later reads in V5.5 | not migrated into new assets tree |
| Stable vs fast batch engines | duplicated logic at lines 3293-3387 and 3439-3530 | consolidated into one batch engine with mode flag |
| Stable vs fast builder scheduling | duplicated logic at lines 3243-3287 and 3388-3433 | consolidated into one scheduling workflow with mode flag |
| Multiple historical monoliths | `Final V5.ahk`, `Final_V5.2.ahk`, `Final_V5.3.ahk`, `Final_V5.4.ahk`, `Final_V5.5.ahk` | left untouched for comparison, not imported into refactor |
