FocusSlateComposer() {
    return RunDevToolsJS("(() => { let el = document.querySelector('[data-slate-editor=`"true`"]'); if (!el) return 'NO_COMPOSER'; el.focus(); return 'OK_COMPOSER'; })()")
}

FocusSlateComposerReady() {
    if StopRequested()
        return false
    result := RunDevToolsJSGetResult("copy(String((() => { const editors = Array.from(document.querySelectorAll('[data-slate-editor=true]')); const el = editors.find(node => node.offsetParent !== null); if (!el) return 'NO_COMPOSER'; el.focus(); el.click(); return 'OK_COMPOSER'; })()));")
    return (result = "OK_COMPOSER")
}

EnsureParticipantInputReady() {
    if StopRequested()
        return false
    result := RunParticipantInputFocus()
    if (result = "OK_PARTICIPANT_INPUT")
        return true

    if !FocusWorkBrowser()
        return false

    ToolTip("Click on the To: field, then release to continue.")
    KeyWait "LButton"
    KeyWait "LButton", "D"
    KeyWait "LButton"
    ToolTip()
    return SafeSleep(150)
}

EnsureQuoComposerReady() {
    if StopRequested()
        return false
    if !FocusWorkBrowser()
        return false

    if FocusSlateComposerReady() {
        Sleep 150
        return true
    }

    ToolTip("Click in the Quo chat box, then release to start.")
    KeyWait "LButton"
    KeyWait "LButton", "D"
    KeyWait "LButton"
    ToolTip()
    return SafeSleep(150)
}

QuoPrimeNewConversation(phone, mode := "stable") {
    global BATCH_AFTER_ALTN, BATCH_AFTER_PHONE, BATCH_AFTER_TAB
    global BATCH_POST_PARTICIPANT_READY_STABLE, BATCH_POST_PARTICIPANT_READY_FAST
    global BATCH_AFTER_PARTICIPANT_TO_COMPOSER

    if StopRequested()
        return false
    Send "!n"
    if !SafeSleep(BATCH_AFTER_ALTN)
        return false

    if !EnsureParticipantInputReady()
        return false

    postReadyDelay := (mode = "stable") ? BATCH_POST_PARTICIPANT_READY_STABLE : BATCH_POST_PARTICIPANT_READY_FAST
    if !SafeSleep(postReadyDelay)
        return false

    if !PasteValue(phone)
        return false
    if !SafeSleep(BATCH_AFTER_PHONE)
        return false

    if StopRequested()
        return false
    Send "{Tab}"
    if !SafeSleep(BATCH_AFTER_TAB)
        return false
    if StopRequested()
        return false
    Send "{Tab}"
    if !SafeSleep(BATCH_AFTER_PARTICIPANT_TO_COMPOSER)
        return false

    FocusSlateComposer()
    return true
}

QuoSelectLeadHolder(holderName) {
    global batchTabsChatToName, BATCH_AFTER_ENTER, BATCH_AFTER_NAME_PICK

    if !SafeSleep(400)
        return false
    if !SendTabs(batchTabsChatToName)
        return false
    if !SafeSleep(200)
        return false

    if StopRequested()
        return false
    Send "{Enter}"
    if !SafeSleep(BATCH_AFTER_ENTER)
        return false

    if StopRequested()
        return false
    Send "^a"
    if !SafeSleep(80)
        return false
    if !PasteValue(holderName)
        return false
    if !SafeSleep(BATCH_AFTER_NAME_PICK)
        return false

    if StopRequested()
        return false
    Send "{Enter}"
    if !SafeSleep(BATCH_AFTER_ENTER)
        return false

    return true
}

QuoScheduleCurrentMessage(msgText, dateMDY, time12) {
    if StopRequested()
        return false
    if !SetClip(msgText) {
        if StopRequested()
            return false
        MsgBox("Clipboard failed to set message text.")
        return false
    }
    if !SafeSleep(100)
        return false
    if StopRequested()
        return false
    Send "^v"
    if StopRequested()
        return false
    Send "^!{Enter}"
    if !SafeSleep(300)
        return false

    if !SetClip(dateMDY " " time12) {
        if StopRequested()
            return false
        MsgBox("Clipboard failed to set date/time.")
        return false
    }
    if !SafeSleep(100)
        return false
    if StopRequested()
        return false
    Send "^v"
    if !SafeSleep(300)
        return false
    if StopRequested()
        return false
    Send "{Enter}"
    if !SafeSleep(200)
        return false
    return true
}

QuoScheduleCurrentMessageTyped(msgText, dateMDY, time12) {
    global SLOW_ACTIVATE_DELAY, SLOW_AFTER_MSG, SLOW_AFTER_SCHED, SLOW_AFTER_DT_PASTE, SLOW_AFTER_ENTER

    if StopRequested()
        return false
    if !FocusWorkBrowser() {
        MsgBox("Browser not found/active. Open the chat window first.")
        return false
    }
    if !SafeSleep(SLOW_ACTIVATE_DELAY)
        return false

    if !SetClip(msgText)
        return false
    if !SafeSleep(50)
        return false
    if StopRequested()
        return false
    Send "^v"
    if !SafeSleep(SLOW_AFTER_MSG)
        return false

    if StopRequested()
        return false
    Send "^!{Enter}"
    if !SafeSleep(SLOW_AFTER_SCHED)
        return false
    if StopRequested()
        return false
    SendText dateMDY . " " . time12
    if !SafeSleep(SLOW_AFTER_DT_PASTE)
        return false
    if StopRequested()
        return false
    Send "{Enter}"
    if !SafeSleep(SLOW_AFTER_ENTER)
        return false
    return true
}
