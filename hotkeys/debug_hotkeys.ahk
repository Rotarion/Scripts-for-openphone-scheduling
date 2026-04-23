Esc:: {
    global StopFlag, running
    StopFlag := true
    running := false
    SetTimer(SpamLoop, 0)
    PersistRunState("stop-requested")
    ToolTip("STOPPED")
    SetTimer(ClearStopToolTip, -700)
}

F1::ExitApp

^!g:: {
    global tagSymbol

    BeginAutomationRun()
    result := RunQuoTagSelector()
    if (result = "") {
        MsgBox("Selector JS returned blank result.")
        return
    }

    finalStatus := HandleQuoTagSelectorResult(result, tagSymbol)
    MsgBox(
        "Selector result: " result "`n"
        . "Final status: " finalStatus "`n"
        . "Tag used: " tagSymbol,
        "Quo Tag Selector Test"
    )
}

^!c:: {
    hwnd := WinGetID("A")
    controls := WinGetControls("ahk_id " hwnd)

    out := ""
    for c in controls
        out .= c "`n"

    MsgBox(out = "" ? "No controls found." : out)
}

^!`:: {
    global agentName, agentEmail, tagSymbol

    ib1 := InputBox("Escribe el nombre del agente:", "Cambiar nombre del agente", "w360 h130", agentName)
    if (ib1.Result != "OK")
        return
    newName := Trim(ib1.Value)
    if (newName = "" || StrLen(newName) > 60) {
        MsgBox("Nombre inválido. Intenta de nuevo.")
        return
    }

    ib2 := InputBox("Escribe el correo del agente:", "Cambiar correo del agente", "w420 h130", agentEmail)
    if (ib2.Result != "OK")
        return
    newEmail := Trim(ib2.Value)
    if (newEmail = "" || StrLen(newEmail) > 100 || !RegExMatch(newEmail, "^[^@\s]+@[^@\s]+\.[^@\s]+$")) {
        MsgBox("Correo inválido. Intenta de nuevo.")
        return
    }

    ib3 := InputBox("Escribe el símbolo del tag:", "Cambiar símbolo del tag", "w260 h130", tagSymbol)
    if (ib3.Result != "OK")
        return
    newTagSymbol := Trim(ib3.Value)
    if (newTagSymbol = "" || StrLen(newTagSymbol) > 3) {
        MsgBox("Símbolo inválido. Intenta de nuevo.")
        return
    }

    UpdateAgentConfiguration(newName, newEmail, newTagSymbol)
    TrayTip("AHK", "Agente actualizado a:`n" newName "`n" newEmail "`nSímbolo: " newTagSymbol, 1)
}

^!m:: {
    raw := Trim(A_Clipboard)
    if (raw = "") {
        MsgBox("Clipboard empty.")
        return
    }

    data := ParseLabeledLeadRaw(raw)
    dobRaw := data.Has("Date of Birth") ? data["Date of Birth"] : ""
    zipRaw := data.Has("Zip Code") ? data["Zip Code"] : ""

    MsgBox(
        "DOB RAW:`n" dobRaw
        . "`n`nDOB NORMALIZED:`n" NormalizeDOB(dobRaw)
        . "`n`nZIP RAW:`n" zipRaw
        . "`n`nZIP NORMALIZED:`n" NormalizeZip(zipRaw),
        "DOB / ZIP Debug"
    )
}

^!p:: {
    raw := Trim(A_Clipboard)
    fields := NormalizeProspectInput(raw)
    msg := ""
    for k, v in fields
        msg .= k ": " v "`n"
    MsgBox(msg, "Parsed Prospect")
}

^!]:: {
    raw := Trim(A_Clipboard)
    if (raw = "") {
        MsgBox("Clipboard empty.")
        return
    }
    data := ParseLabeledLeadRaw(raw)
    msg := ""
    for k, v in data
        msg .= "[" k "] = " v "`n"
    MsgBox(msg, "Labeled Lead Raw Map")
}

^!l:: {
    global batchLeadHolder
    raw := Trim(A_Clipboard)
    if (raw = "") {
        MsgBox("Clipboard empty.")
        return
    }

    batchLeadHolder := BuildBatchLeadHolder(raw)

    msg := "Leads found: " batchLeadHolder.Length "`n`n"
    for i, lead in batchLeadHolder {
        previewPrice := ResolveQuotePrice(lead["VEHICLE_COUNT"], lead["VEHICLES"], true)
        msg .= i ". " lead["FULL_NAME"]
            . " | Phone: " lead["PHONE"]
            . " | Cars: " lead["VEHICLE_COUNT"]
            . " | Price: " previewPrice
            . "`n"
    }
    MsgBox(msg, "Batch Lead Holder Preview")
}

F8:: {
    global running
    running := !running

    if running {
        BeginAutomationRun()
        ToolTip("RUNNING (F8 to stop)")
        SetTimer(SpamLoop, 200)
    } else {
        SetTimer(SpamLoop, 0)
        PersistRunState("spamloop-stopped")
        ToolTip("STOPPED")
        Sleep 800
        ToolTip()
    }
}

SpamLoop() {
    global running
    if !running || StopRequested()
        return

    Click
    if !SafeSleep(20)
        return
    if StopRequested()
        return
    Send "{Enter}"
}
