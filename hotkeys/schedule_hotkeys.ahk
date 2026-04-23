^!r:: {
    ResetRotationOffset()
    MsgBox("Rotation index reset to 0.")
}

^!6::ScheduleLeadFollowupsByClipboard(false)
^!7::ScheduleLeadFollowupsByClipboard(true)
^!8::ShowFollowupBatchPickerFromClipboard()

^!d:: {
    global configDays, holidays

    if (configDays.Length != 4) {
        MsgBox("configDays has " configDays.Length " items. Check [Schedule] Days= in settings.ini.", "Follow-up Preview")
        return
    }

    Show := (n) => BusinessDateForDay(n, holidays)
    msg := "Config days: " configDays[1] ", " configDays[2] ", " configDays[3] ", " configDays[4] "`n`n"
    msg .= "Resolved dates (business days from today):`n"
    msg .= "A: day " configDays[1] " -> " Show(configDays[1]) "`n"
    msg .= "B: day " configDays[2] " -> " Show(configDays[2]) "`n"
    msg .= "C: day " configDays[3] " -> " Show(configDays[3]) "`n"
    msg .= "D: day " configDays[4] " -> " Show(configDays[4])
    MsgBox(msg, "Follow-up Preview")
}
