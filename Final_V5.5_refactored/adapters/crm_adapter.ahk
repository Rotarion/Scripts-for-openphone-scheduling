FillNewProspectForm(fields) {
    global FORM_FIELD_DELAY, FORM_TAB_DELAY, FORM_PASTE_DELAY, FORM_PASTE_TAB_DELAY, FORM_CITY_TAB_DELAY

    FastType(fields["FIRST_NAME"])
    Sleep FORM_FIELD_DELAY
    Send "{Tab}"
    Sleep FORM_TAB_DELAY

    Send "{Tab}"
    Sleep FORM_TAB_DELAY

    FastType(fields["LAST_NAME"])
    Sleep FORM_FIELD_DELAY
    Send "{Tab}"
    Sleep FORM_TAB_DELAY

    Send "{Tab}"
    Sleep FORM_TAB_DELAY

    PasteField(fields["DOB"])
    Sleep FORM_PASTE_DELAY
    Send "{Tab}"
    Sleep FORM_PASTE_TAB_DELAY

    SelectDropdownValue(fields["GENDER"])

    FastType(fields["ADDRESS_1"])
    Sleep FORM_FIELD_DELAY
    Send "{Tab}"
    Sleep FORM_TAB_DELAY

    FastType(fields["APT_SUITE"])
    Sleep FORM_FIELD_DELAY
    Send "{Tab}"
    Sleep FORM_TAB_DELAY

    FastType(fields["BUILDING"])
    Sleep FORM_FIELD_DELAY
    Send "{Tab}"
    Sleep FORM_TAB_DELAY

    FastType(fields["RR_NUMBER"])
    Sleep FORM_FIELD_DELAY
    Send "{Tab}"
    Sleep FORM_TAB_DELAY

    FastType(fields["LOT_NUMBER"])
    Sleep FORM_FIELD_DELAY
    Send "{Tab}"
    Sleep FORM_TAB_DELAY

    FastType(fields["CITY"])
    Sleep FORM_FIELD_DELAY
    Send "{Tab}"
    Sleep FORM_CITY_TAB_DELAY

    SelectDropdownValue(fields["STATE"])

    FastType(fields["ZIP"])
    Sleep FORM_FIELD_DELAY
    Send "{Tab}"
    Sleep FORM_TAB_DELAY

    FastType(fields["PHONE"])
}

FillNationalGeneralForm(fields) {
    global FORM_FIELD_DELAY, FORM_TAB_DELAY, FORM_PASTE_DELAY, FORM_PASTE_TAB_DELAY, FORM_CITY_TAB_DELAY

    FastType(fields["FIRST_NAME"])
    Sleep FORM_FIELD_DELAY
    Send "{Tab}"
    Sleep FORM_TAB_DELAY

    Send "{Tab}"
    Sleep FORM_TAB_DELAY

    FastType(fields["LAST_NAME"])
    Sleep FORM_FIELD_DELAY
    Send "{Tab}"
    Sleep FORM_TAB_DELAY

    Loop 5 {
        Send "{Tab}"
        Sleep FORM_TAB_DELAY
    }

    PasteField(fields["DOB"])
    Sleep FORM_PASTE_DELAY
    Send "{Tab}"
    Sleep FORM_PASTE_TAB_DELAY

    Loop 3 {
        Send "{Tab}"
        Sleep FORM_TAB_DELAY
    }

    FastType(fields["ADDRESS_1"])
    Sleep FORM_FIELD_DELAY
    Send "{Tab}"
    Sleep FORM_TAB_DELAY

    FastType(fields["APT_SUITE"])
    Sleep FORM_FIELD_DELAY
    Send "{Tab}"
    Sleep FORM_TAB_DELAY

    FastType(fields["CITY"])
    Sleep FORM_FIELD_DELAY
    Send "{Tab}"
    Sleep FORM_CITY_TAB_DELAY

    SelectDropdownValue(fields["STATE"])

    FastType(fields["ZIP"])
    Sleep FORM_FIELD_DELAY
    Send "{Tab}"
    Sleep FORM_TAB_DELAY
}

CrmApplyAppointmentPreset(dtText, postDateKeys) {
    global CRM_KEYSTEP_DELAY, CRM_MEDIUM_DELAY

    if !SetClip(dtText) {
        MsgBox("Clipboard failed (date).")
        return false
    }

    Sleep 80
    SendEvent "^v"
    Sleep 220

    if !SendTabs(6)
        return false
    Sleep CRM_KEYSTEP_DELAY
    for _, key in postDateKeys {
        SendEvent key
        Sleep CRM_KEYSTEP_DELAY
    }
    if !SendTabs(3)
        return false
    Sleep CRM_KEYSTEP_DELAY
    SendEvent "c"
    Sleep CRM_MEDIUM_DELAY
    return true
}

CrmRunAttemptedContactAppointment(noteText, dtText) {
    global CRM_ACTION_FOCUS_DELAY, CRM_KEYSTEP_DELAY, CRM_SHORT_DELAY, CRM_MEDIUM_DELAY
    global CRM_SAVE_HISTORY_DELAY, CRM_ADD_APPOINTMENT_DELAY, CRM_FOCUS_DATE_DELAY, CRM_FINAL_SAVE_DELAY

    JS_FocusActionDropdown()
    Sleep CRM_ACTION_FOCUS_DELAY

    SendEvent "l"
    Sleep CRM_KEYSTEP_DELAY
    SendEvent "{Tab}"
    Sleep CRM_KEYSTEP_DELAY
    SendEvent "{Tab}"
    Sleep CRM_MEDIUM_DELAY
    SendEvent "1"
    Sleep CRM_MEDIUM_DELAY

    if !SendTabs(9)
        return false
    Sleep CRM_SHORT_DELAY

    if !SetClip(noteText) {
        MsgBox("Clipboard failed (note text).")
        return false
    }
    Sleep 80
    SendEvent "^v"
    Sleep CRM_MEDIUM_DELAY

    JS_SaveHistoryNote()
    Sleep CRM_SAVE_HISTORY_DELAY

    JS_AddNewAppointment()
    Sleep CRM_ADD_APPOINTMENT_DELAY

    JS_FocusDateTimeField()
    Sleep CRM_FOCUS_DATE_DELAY

    if !CrmApplyAppointmentPreset(dtText, ["e"])
        return false

    JS_SaveAppointment()
    Sleep CRM_FINAL_SAVE_DELAY
    return true
}

CrmRunQuoteCallAppointment(noteText, dtText) {
    global CRM_ACTION_FOCUS_DELAY, CRM_KEYSTEP_DELAY, CRM_MEDIUM_DELAY
    global CRM_QUOTE_SHIFT_TAB_DELAY, CRM_SAVE_HISTORY_DELAY, CRM_ADD_APPOINTMENT_DELAY
    global CRM_FOCUS_DATE_DELAY, CRM_FINAL_SAVE_DELAY

    JS_FocusActionDropdown()
    Sleep CRM_ACTION_FOCUS_DELAY

    SendEvent "l"
    Sleep CRM_KEYSTEP_DELAY
    SendEvent "{Tab}"
    Sleep CRM_KEYSTEP_DELAY
    SendEvent "q"
    Sleep CRM_KEYSTEP_DELAY
    SendEvent "+{Tab}"
    Sleep CRM_QUOTE_SHIFT_TAB_DELAY
    SendEvent "{Tab}"
    Sleep CRM_MEDIUM_DELAY
    SendEvent "{Tab}"
    Sleep CRM_MEDIUM_DELAY
    SendEvent "3"
    Sleep CRM_MEDIUM_DELAY

    if !SendTabs(9)
        return false
    Sleep CRM_MEDIUM_DELAY

    if !SetClip(noteText) {
        MsgBox("Clipboard failed (note text).")
        return false
    }
    Sleep 80
    SendEvent "^v"
    Sleep CRM_MEDIUM_DELAY

    JS_SaveHistoryNote()
    Sleep CRM_SAVE_HISTORY_DELAY

    JS_AddNewAppointment()
    Sleep CRM_ADD_APPOINTMENT_DELAY

    JS_FocusDateTimeField()
    Sleep CRM_FOCUS_DATE_DELAY

    if !CrmApplyAppointmentPreset(dtText, ["p", "p"])
        return false

    JS_SaveAppointment()
    Sleep CRM_FINAL_SAVE_DELAY
    return true
}
