# Risk Log

## Not Fully Verified In-Environment

| Risk | Why it remains |
| --- | --- |
| AutoHotkey runtime execution | AutoHotkey executable was not available in this terminal session, so syntax/runtime validation could not be executed locally |
| Live browser focus timing | Chrome/Edge window timing can only be verified against the real target tabs |
| Slate/Quo composer behavior | DevTools and Slate focus behavior depends on live DOM state |
| CRM iframe selectors | The iframe-based CRM actions were preserved but not replayed against a live session here |
| Non-ASCII message rendering | The refactor preserves Spanish copy and symbols, but actual rendering should be checked under the production AHK file encoding |
| Resume-state mirror | `logs/run_state.json` mirrors runtime state, but the batch raw clipboard comparison remains memory-backed for safety |

## High-Risk Areas Kept Conservative

- Clipboard restore semantics were preserved in the DevTools bridge.
- Destructive keystrokes such as backspace, enter, paste, and tab were not redesigned; they were wrapped.
- Participant-input focus and tag-selection JS were kept as separate assets instead of being fused or simplified.
- CRM note codes `txt` and `qt` were preserved as template values.
