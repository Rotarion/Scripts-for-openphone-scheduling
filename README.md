# OpenPhone Scheduling Scripts

AutoHotkey v2 automation for scheduling insurance quotes and follow-up messages via OpenPhone (QUO), with CRM and form-fill support.

## Requirements

- AutoHotkey v2.0
- Chrome or Edge (for DevTools-based tag selection)

## Quick Start

Run `main.ahk` with AutoHotkey v2. All config is in `config/`.

## Key Hotkeys

| Hotkey | Action |
|--------|--------|
| `Ctrl+Alt+U` | Quick single lead create + tag |
| `Ctrl+Alt+B` | Stable batch run (typed input) |
| `Ctrl+Alt+N` | Fast batch run (pasted input) |
| `Ctrl+Alt+6` | Fast message schedule |
| `Ctrl+Alt+7` | Stable message schedule |
| `Ctrl+Alt+8` | Batch picker schedule |
| `Ctrl+Alt+0` | Fill National General form |
| `Ctrl+Alt+9` | Fill Edge prospect form |
| `Esc` | Stop current operation |
| `F1` | Exit script |

## Project Structure

```
main.ahk              — entry point; loads config, includes all modules
config/               — INI files (settings, timings, templates, holidays)
domain/               — pure business logic (parsing, pricing, dates, messages)
adapters/             — external system integration (clipboard, browser, CRM, QUO)
workflows/            — multi-step orchestration (batch, schedule, form-fill, CRM)
hotkeys/              — keyboard shortcut definitions
assets/js/            — JavaScript helpers executed via DevTools bridge
tests/                — dry-run and fixture tests
logs/                 — batch_lead_log.csv (audit trail); run_state.json (runtime, not committed)
docs/                 — design notes, migration plan, risk log
archive/              — historical monolith versions (reference only)
```

## Configuration

Edit `config/settings.ini` to set your agent name, email, tag symbol, schedule days, pricing, and vehicle filter bounds.

Edit `config/timings.ini` to tune UI interaction delays without touching code.

Edit `config/templates.ini` to update message copy (Spanish/English).

## Logs

`logs/batch_lead_log.csv` is an append-only audit trail of every lead processed. Do not delete it between sessions.
