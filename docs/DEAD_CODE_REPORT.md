# Dead Code Report

## Removed From Refactored Tree

| Item | Basis |
| --- | --- |
| `SendSelectedBatch(picker)` | Legacy V5.5 defines it, but the active picker button calls `SendSelectedBatchV2` |
| `ClickAddNewAppointmentJS()` | Defined in V5.5 with no active callers |
| `tag_activation.js` | Legacy V5.5 declares the path but never reads the file |

## Left Outside The Refactored Tree

| Item | Basis |
| --- | --- |
| Historical monoliths `Final V5.ahk` through `Final_V5.4.ahk` | Useful for forensic diffing, but not imported into the new runtime |
| `Automated copy fixed V2.0.ahk` | Legacy branch artifact, not used by V5.5 |

## Duplicate Logic Consolidated

| Duplicate pair | Consolidation |
| --- | --- |
| `RunBatchLeadFlow` and `RunBatchLeadFlowFast` | merged into `RunBatchLeadFlow(lead, mode)` |
| `ScheduleBuilderForLead` and `ScheduleBuilderForLeadFast` | merged into `ScheduleBuilderForLead(lead, offset, mode)` |
| `ScheduleRegularFollowupsForLead` and `ScheduleRegularFollowupsForLeadFast` | merged into `ScheduleRegularFollowupsForLead(lead, offset, mode)` |

## Quarantined By Omission

The refactored asset tree only carries the active JS assets:

- `assets/js/tag_selector.js`
- `assets/js/participant_input_focus.js`

The older `tag_activation.js` implementation remains only in the legacy workspace, which keeps it available for manual comparison without letting it stay on the active path.
