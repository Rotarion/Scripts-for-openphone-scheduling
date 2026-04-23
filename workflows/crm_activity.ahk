BuildLastConfiguredFollowupDateText() {
    global configDays, holidays

    lastDay := GetMaxConfiguredDay(configDays)
    lastDate := BusinessDateForDay(lastDay, holidays)
    return lastDate . " 3:00 PM"
}

BuildTomorrowMorningDateText() {
    global holidays

    tomorrow := BusinessDateForDay(1, holidays)
    return tomorrow . " 10:00 AM"
}

RunPasteLastConfiguredDatePreset() {
    dtText := BuildLastConfiguredFollowupDateText()
    BeginAutomationRun()
    CrmApplyAppointmentPreset(dtText, ["e"])
}

RunPasteTomorrowPhonePreset() {
    dtText := BuildTomorrowMorningDateText()
    BeginAutomationRun()
    CrmApplyAppointmentPreset(dtText, ["P", "P"])
}

RunCrmAttemptedContactWorkflow() {
    noteText := TemplateRead("CrmNotes", "AttemptedContact", "txt")
    dtText := BuildLastConfiguredFollowupDateText()

    if !FocusWorkBrowser() {
        MsgBox("Browser not detected.")
        return
    }

    BeginAutomationRun()
    CrmRunAttemptedContactAppointment(noteText, dtText)
}

RunCrmQuoteCallWorkflow() {
    noteText := TemplateRead("CrmNotes", "QuoteCall", "qt")
    dtText := BuildTomorrowMorningDateText()

    if !FocusWorkBrowser() {
        MsgBox("Browser not detected.")
        return
    }

    BeginAutomationRun()
    CrmRunQuoteCallAppointment(noteText, dtText)
}
