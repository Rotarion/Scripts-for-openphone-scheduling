ScheduleFollowupMessages(msgs, useTyped := false) {
    for m in msgs {
        if StopRequested()
            return false
        ok := useTyped
            ? QuoScheduleCurrentMessageTyped(m["text"], m["date"], m["time"])
            : QuoScheduleCurrentMessage(m["text"], m["date"], m["time"])
        if !ok
            return false
    }
    return true
}

ScheduleBuilderForLead(lead, offset, mode := "stable") {
    if StopRequested()
        return false

    dt := GetInitialQuoteDateTime(offset)
    msg := BuildMessage(lead["FULL_NAME"], lead["VEHICLE_COUNT"], lead["VEHICLES"], true)
    useTyped := (mode = "stable")

    if !FocusWorkBrowser() {
        MsgBox("Browser not found/active.")
        return false
    }
    if !SafeSleep(200)
        return false

    return useTyped
        ? QuoScheduleCurrentMessageTyped(msg, dt["date"], dt["time"])
        : QuoScheduleCurrentMessage(msg, dt["date"], dt["time"])
}

ScheduleRegularFollowupsForLead(lead, offset, mode := "stable") {
    msgs := BuildFollowupQueue(lead["FULL_NAME"], offset)
    return ScheduleFollowupMessages(msgs, mode = "stable")
}

ScheduleLeadFollowupsByClipboard(useTyped := false) {
    if !ClipWait(1) {
        MsgBox("Clipboard empty. Copy the lead's name first.")
        return
    }

    lead := CleanName(A_Clipboard)
    if (lead = "" || StrLen(lead) > 30) {
        MsgBox("Please copy just the lead's name (<= 30 chars) and try again.")
        return
    }

    BeginAutomationRun()
    idx := NextRotationOffset()
    msgs := BuildFollowupQueue(lead, idx)

    if !ScheduleFollowupMessages(msgs, useTyped) {
        if StopRequested()
            return
        MsgBox(useTyped
            ? "Failed scheduling one of the messages (typed mode). Stopping."
            : "Failed scheduling one of the messages. Stopping.")
    }
}

ShowFollowupBatchPickerFromClipboard() {
    global configDays

    if !ClipWait(1) {
        MsgBox("Clipboard empty. Copy the lead's name first.")
        return
    }

    lead := CleanName(A_Clipboard)
    if (lead = "" || StrLen(lead) > 30) {
        MsgBox("Please copy just the lead's name (<= 30 chars) and try again.")
        return
    }

    offset := NextRotationOffset()
    fullMsgs := BuildFollowupQueue(lead, offset)

    dA := configDays[1], dB := configDays[2], dC := configDays[3], dD := configDays[4]

    picker := Gui("+AlwaysOnTop", "Select Follow-Up Batch")
    picker.SetFont("s10")
    picker.Add("Text",, "Selecciona los bloques (día de envío entre paréntesis):")
    cbD1 := picker.Add("CheckBox",, "Bloque A (día " dA ")"), cbD1.Value := 0
    cbD2 := picker.Add("CheckBox",, "Bloque B (día " dB ")"), cbD2.Value := 0
    cbD4 := picker.Add("CheckBox",, "Bloque C (día " dC ")"), cbD4.Value := 0
    cbD5 := picker.Add("CheckBox",, "Bloque D (día " dD ")"), cbD5.Value := 0

    picker.Add("Text", "xm y+10", "Modo de envío")
    ddMode := picker.Add("DropDownList", "w220", ["Pegar (rápido)", "Escritura estable (Chrome)"])
    ddMode.Choose(1)

    picker.Add("Text", "xm y+10", "")
    btnStart := picker.Add("Button", "w120", "Iniciar")
    btnCancel := picker.Add("Button", "x+10 w90", "Cancelar")

    picker.cbD1 := cbD1
    picker.cbD2 := cbD2
    picker.cbD4 := cbD4
    picker.cbD5 := cbD5
    picker.ddMode := ddMode
    picker.fullMsgs := fullMsgs

    btnStart.OnEvent("Click", (*) => SendSelectedBatchV2(picker))
    btnCancel.OnEvent("Click", (*) => picker.Destroy())
    picker.Show()
}

SendSelectedBatchV2(picker) {
    BeginAutomationRun()

    useD1 := picker.cbD1.Value
    useD2 := picker.cbD2.Value
    useD4 := picker.cbD4.Value
    useD5 := picker.cbD5.Value
    modeText := picker.ddMode.Text
    fullMsgs := picker.fullMsgs

    selectedDays := []
    if (useD1)
        selectedDays.Push(fullMsgs[1]["day"])
    if (useD2)
        selectedDays.Push(fullMsgs[5]["day"])
    if (useD4)
        selectedDays.Push(fullMsgs[7]["day"])
    if (useD5)
        selectedDays.Push(fullMsgs[9]["day"])

    if (selectedDays.Length = 0) {
        MsgBox("Select at least one block (A, B, C, or D).")
        return
    }

    toSend := []
    for m in fullMsgs {
        if (m.Has("day") && ArrContains(selectedDays, m["day"]))
            toSend.Push(m)
    }

    SortMessagesByDaySeq(toSend)

    try picker.Opt("-AlwaysOnTop")
    try picker.Hide()

    if !EnsureQuoComposerReady() {
        if StopRequested()
            return
        try picker.Show()
        try picker.Opt("+AlwaysOnTop")
        MsgBox("Could not focus the Quo chat box. Open the Quo tab and try again.")
        return
    }

    useTyped := InStr(modeText, "Stable") || InStr(modeText, "estable")
    try picker.Destroy()

    if !ScheduleFollowupMessages(toSend, useTyped) {
        if StopRequested()
            return
        MsgBox(useTyped
            ? "Failed scheduling one of the messages in stable mode. Stopping."
            : "Failed scheduling one of the messages in fast mode. Stopping.")
        return
    }

    TrayTip("AHK", "Scheduled messages: " . toSend.Length, 1)
}
