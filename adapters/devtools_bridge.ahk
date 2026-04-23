RunDevToolsJS(jsCode) {
    if StopRequested()
        return false
    if !FocusWorkBrowser()
        return false

    savedClip := ClipboardAll()

    try {
        A_Clipboard := ""
        if !SafeSleep(30)
            return false
        if StopRequested()
            return false
        A_Clipboard := jsCode
        if !WaitForClip(1000)
            return false

        if StopRequested()
            return false
        Send "^+j"
        if !SafeSleep(500)
            return false

        if StopRequested()
            return false
        Send "^a"
        if !SafeSleep(80)
            return false
        if StopRequested()
            return false
        Send "^v"
        if !SafeSleep(120)
            return false
        if StopRequested()
            return false
        Send "{Enter}"
        if !SafeSleep(180)
            return false

        if StopRequested()
            return false
        Send "^+j"
        if !SafeSleep(180)
            return false
        return true
    } finally {
        A_Clipboard := savedClip
    }
}

RunDevToolsJSGetResult(jsCode) {
    if StopRequested()
        return ""
    if !FocusWorkBrowser()
        return ""

    savedClip := ClipboardAll()

    try {
        A_Clipboard := ""
        if !SafeSleep(30)
            return ""
        if StopRequested()
            return ""
        A_Clipboard := jsCode
        if !WaitForClip(1000)
            return ""

        sentCode := A_Clipboard

        if StopRequested()
            return ""
        Send "^+j"
        if !SafeSleep(500)
            return ""

        if StopRequested()
            return ""
        Send "^a"
        if !SafeSleep(80)
            return ""
        if StopRequested()
            return ""
        Send "^v"
        if !SafeSleep(120)
            return ""
        if StopRequested()
            return ""
        Send "{Enter}"
        if !SafeSleep(300)
            return ""

        result := ""
        Loop 20 {
            if !SafeSleep(100)
                return ""
            if (A_Clipboard != sentCode && Trim(A_Clipboard) != "") {
                result := Trim(A_Clipboard)
                break
            }
        }

        if StopRequested()
            return ""
        Send "^+j"
        if !SafeSleep(220)
            return ""
        FocusWorkBrowser()
        if !SafeSleep(150)
            return ""

        return result
    } finally {
        A_Clipboard := savedClip
    }
}

BuildParticipantInputFocusJS() {
    global participantInputJsFile
    if !FileExist(participantInputJsFile)
        return ""

    return Trim(FileRead(participantInputJsFile, "UTF-8"))
}

RunParticipantInputFocus() {
    js := BuildParticipantInputFocusJS()
    if (js = "")
        return ""
    return RunDevToolsJSGetResult(js)
}

JS_FocusActionDropdown() {
    return RunDevToolsJS("(() => { let d = document.querySelectorAll('iframe')[0]?.contentDocument; let el = d?.getElementById('ctl00_ContentPlaceHolder1_DDLogType_Input'); if (!el) return 'NO_ACTION'; el.focus(); ['mouseover','mousedown','mouseup','click'].forEach(t => el.dispatchEvent(new MouseEvent(t,{bubbles:true,cancelable:true,view:d.defaultView}))); return 'OK_ACTION'; })()")
}

JS_SaveHistoryNote() {
    return RunDevToolsJS("(() => { let d = document.querySelectorAll('iframe')[0]?.contentDocument; let el = d?.getElementById('ctl00_ContentPlaceHolder1_btnUpdate_input'); if (!el) return 'NO_SAVE'; el.focus(); ['mouseover','mousedown','mouseup','click'].forEach(t => el.dispatchEvent(new MouseEvent(t,{bubbles:true,cancelable:true,view:d.defaultView}))); return 'OK_SAVE'; })()")
}

JS_AddNewAppointment() {
    return RunDevToolsJS("(() => { let d = document.querySelectorAll('iframe')[0]?.contentDocument; if (!d) return 'NO_FRAME1'; if (typeof d.defaultView.AppointmentInserting === 'function') { d.defaultView.AppointmentInserting(); return 'OK_FUNC'; } let el = d.querySelector('a.js-Lead-Log-Add-New-Appointment'); if (!el) return 'NO_APPT'; el.focus(); ['mouseover','mousedown','mouseup','click'].forEach(t => el.dispatchEvent(new MouseEvent(t,{bubbles:true,cancelable:true,view:d.defaultView}))); return 'OK_APPT'; })()")
}

JS_FocusDateTimeField() {
    return RunDevToolsJS("(() => { let d1 = document.querySelectorAll('iframe')[0]?.contentDocument; let d2 = d1?.querySelectorAll('iframe')[0]?.contentDocument; let el = d2?.getElementById('ctl00_ContentPlaceHolder1_RadDateTimePicker1_dateInput'); if (!el) return 'NO_TIME'; el.focus(); el.select(); return 'OK_TIME'; })()")
}

JS_SaveAppointment() {
    return RunDevToolsJS("(() => { let d1 = document.querySelectorAll('iframe')[0]?.contentDocument; let d2 = d1?.querySelectorAll('iframe')[0]?.contentDocument; let el = d2?.getElementById('ctl00_ContentPlaceHolder1_lnkSave_input'); if (!el) return 'NO_FINAL'; el.focus(); ['mouseover','mousedown','mouseup','click'].forEach(t => el.dispatchEvent(new MouseEvent(t,{bubbles:true,cancelable:true,view:d2.defaultView}))); return 'OK_FINAL'; })()")
}
