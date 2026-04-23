#Requires AutoHotkey v2.0

#Requires AutoHotkey v2.0
#SingleInstance Force
SendMode "Input"

; ===================== HEADER / DIRECTIVES =====================

; ===================== CONFIGURATION =====================
iniFile   := A_ScriptDir "\time_rotation.ini"
holidays := ["01/01/2026","05/25/2026","06/19/2026","07/03/2026","07/04/2026","09/07/2026","11/26/2026","12/25/2026"]  ; MM/dd/yyyy

EnsureIniDefaults()

; Agent persisted in the same INI
agentName  := IniRead(iniFile, "Agent", "Name", "Pablo Cabrera")
agentEmail := IniRead(iniFile, "Agent", "Email", "pablocabrera@allstate.com")
tagSymbol  := Trim(IniRead(iniFile, "Agent", "TagSymbol", "+"))
if (tagSymbol = "")
    tagSymbol := "+"

; Schedule days (read from INI)
daysStr := IniRead(iniFile, "Schedule", "Days", "1,2,4,5")
configDays := ParseDays(daysStr)
if (configDays.Length != 4) {
    configDays := [1, 2, 4, 5]
    IniWrite("1,2,4,5", iniFile, "Schedule", "Days")
}

batchMinVehicles := Integer(IniRead(iniFile, "Batch", "MinVehicles", "0"))
batchMaxVehicles := Integer(IniRead(iniFile, "Batch", "MaxVehicles", "99"))
dobDefaultDay    := Integer(IniRead(iniFile, "DOB", "DefaultDay", "16"))
batchTabsChatToName := Integer(IniRead(iniFile, "UI", "BatchTabsChatToName", "7"))
priceOldCar := Integer(IniRead(iniFile, "Pricing", "OldCar", "98"))
priceOneCar := Integer(IniRead(iniFile, "Pricing", "OneCar", "117"))
priceOneCar2020Plus := Integer(IniRead(iniFile, "Pricing", "OneCar2020Plus", "167"))
priceTwoCars := Integer(IniRead(iniFile, "Pricing", "TwoCars", "176"))
priceThreeCars := Integer(IniRead(iniFile, "Pricing", "ThreeCars", "284"))
priceFourCars := Integer(IniRead(iniFile, "Pricing", "FourCars", "397"))
priceFiveCars := Integer(IniRead(iniFile, "Pricing", "FiveCars", "397"))
singleCarModernYearCutoff := Integer(IniRead(iniFile, "Pricing", "SingleCarModernYearCutoff", "2020"))

; Log file creator
batchLogFile := A_ScriptDir "\batch_lead_log.csv"
tagActivationJsFile := A_ScriptDir "\tag_activation.js"
participantInputJsFile := A_ScriptDir "\participant_input_focus.js"

; Typing-mode pacing for ^!7 (stable)
SLOW_ACTIVATE_DELAY := 250
SLOW_AFTER_MSG      := 550
SLOW_AFTER_SCHED    := 650
SLOW_AFTER_DT_PASTE := 650
SLOW_AFTER_ENTER    := 300

; Batch holder + pacing
batchLeadHolder := []

BATCH_AFTER_ALTN        := 5000
BATCH_AFTER_PHONE       := 650
BATCH_AFTER_TAB         := 150
BATCH_AFTER_SCHEDULE    := 600
BATCH_AFTER_ENTER       := 150
BATCH_AFTER_NAME_PICK   := 250
BATCH_AFTER_TAG_PICK    := 250
BATCH_BEFORE_TAG_PASTE  := 500
BATCH_AFTER_TAG_PASTE   := 700

; ===================== GLOBAL STATE =====================

running := false
StopFlag := false
batchResumeIndex := 1
batchResumeMode := ""
batchResumeRaw := ""

BeginAutomationRun() {
    global StopFlag
    StopFlag := false
}

StopRequested() {
    global StopFlag
    return StopFlag
}

SafeSleep(ms) {
    if (ms <= 0)
        return !StopRequested()

    endTick := A_TickCount + ms
    while (A_TickCount < endTick) {
        if StopRequested()
            return false
        Sleep Min(25, endTick - A_TickCount)
    }
    return !StopRequested()
}

WaitForClip(timeoutMs := 1000) {
    endTick := A_TickCount + timeoutMs
    while (A_TickCount < endTick) {
        if StopRequested()
            return false
        if ClipWait(0.05)
            return true
    }
    return false
}

ClearStopToolTip() {
    ToolTip()
}

; ===================== HOTKEY ENTRY POINTS =====================

Esc:: {
    global StopFlag, running
    StopFlag := true
    running := false
    SetTimer(SpamLoop, 0)
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

^!r:: {
    global iniFile
    IniWrite(0, iniFile, "Times", "Offset")
    MsgBox("Rotation index reset to 0.")
}

^!c:: {
    hwnd := WinGetID("A")  ; "A" = active window
    controls := WinGetControls("ahk_id " hwnd)

    out := ""
    for c in controls
        out .= c "`n"

    MsgBox out = "" ? "No controls found." : out
}

^!`:: {
    global agentName, agentEmail, tagSymbol, iniFile

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

    IniWrite(newName,      iniFile, "Agent", "Name")
    IniWrite(newEmail,     iniFile, "Agent", "Email")
    IniWrite(newTagSymbol, iniFile, "Agent", "TagSymbol")

    agentName  := newName
    agentEmail := newEmail
    tagSymbol  := newTagSymbol

    TrayTip("AHK", "Agente actualizado a:`n" newName "`n" newEmail "`nSímbolo: " newTagSymbol, 1)
}

^!u:: {
    if !ClipWait(1) {
        MsgBox("Clipboard empty. Copy one lead row first.")
        return
    }

    raw := Trim(A_Clipboard)
    if (raw = "") {
        MsgBox("Clipboard empty. Copy one lead row first.")
        return
    }

    lead := BuildBatchLeadRecord(raw)

    if (lead["PHONE"] = "" || lead["FULL_NAME"] = "") {
        MsgBox("Could not parse phone/name from clipboard.")
        return
    }

    BeginAutomationRun()
    result := RunQuickLeadCreateAndTag(lead)

    if StopRequested()
        return
    if (result != "OK")
        MsgBox(result)
}

^!1:: {
    global priceOldCar, priceOneCar, priceTwoCars, priceThreeCars, priceFourCars, priceFiveCars

    if !ClipWait(1) {
        MsgBox("Clipboard empty. Copy the lead's name first.")
        return
    }
    leadName := CleanName(A_Clipboard)
    if (leadName = "" || StrLen(leadName) > 30) {
        MsgBox("Please copy just the lead's name (<= 30 chars).")
        return
    }

    promptText :=
        "Escribe 0, 1, 2, 3, 4 o 5:`n"
        . "0 = auto muy antiguo (" FormatMonthlyPrice(priceOldCar) ", sin 'FULL COVERAGE')`n"
        . "1 = su carro (" FormatMonthlyPrice(priceOneCar) ", FULL COVERAGE)`n"
        . "2 = sus carros (" FormatMonthlyPrice(priceTwoCars) ", FULL COVERAGE)`n"
        . "3 = sus carros (" FormatMonthlyPrice(priceThreeCars) ", FULL COVERAGE)`n"
        . "4 = sus carros (" FormatMonthlyPrice(priceFourCars) ", FULL COVERAGE)`n"
        . "5 = sus carros (" FormatMonthlyPrice(priceFiveCars) ", FULL COVERAGE, 100/300)"

    ib := InputBox(promptText, "Número de vehículos", "w420 h280")
    if (ib.Result != "OK")
        return

    choice := Trim(ib.Value)
    if !(choice ~= "^(0|1|2|3|4|5)$") {
        MsgBox("Opción inválida. Usa 0, 1, 2, 3, 4 o 5.")
        return
    }

    carCount := Integer(choice)
    msg := BuildMessage(leadName, carCount)

    A_Clipboard := ""
    Sleep 50
    A_Clipboard := msg
    if !ClipWait(1) {
        MsgBox("No se pudo copiar el mensaje al portapapeles.")
        return
    }
    Sleep 100
    Send "^v"
    Sleep 300

    A_Clipboard := ""
    Sleep 50
    A_Clipboard := ProperCase(leadName)
    ClipWait(1)
}

^!6:: {
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

    if !ScheduleFollowupMessages(msgs, false) {
        if StopRequested()
            return
        MsgBox("Failed scheduling one of the messages. Stopping.")
        return
    }
}

^!7:: {
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

    if !ScheduleFollowupMessages(msgs, true) {
        if StopRequested()
            return
        MsgBox("Failed scheduling one of the messages (typed mode). Stopping.")
        return
    }
}

^!8:: {
    global iniFile, holidays, configDays

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
    ddMode := picker.Add("DropDownList", "w220", ["Pegar (rápido)","Escritura estable (Chrome)"])
    ddMode.Choose(1)

    picker.Add("Text", "xm y+10", "")
    btnStart  := picker.Add("Button", "w120", "Iniciar")
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

^!d:: {
    global configDays, holidays
    if (configDays.Length != 4) {
        MsgBox("configDays has " configDays.Length " items. Check [Schedule] Days= in the INI.", "Follow-up Preview")
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
^!0:: {
    raw := Trim(A_Clipboard)
    if (raw = "") {
        MsgBox("Clipboard empty. Copy the raw lead first.")
        return
    }

    fields := NormalizeProspectInput(raw)

    if !FocusWorkBrowser() {
        MsgBox("Browser not found. Open National General first.")
        return
    }

    BeginAutomationRun()

   ; ToolTip("Click FIRST NAME now")
   ; KeyWait "LButton", "D"
   ; Sleep 120
   ; KeyWait "LButton"
   ; ToolTip()


    ToolTip("Click FIRST NAME now (2s)")
    Sleep 500
    ToolTip()


    FillNationalGeneralForm(fields)
}


^!9:: {
    raw := Trim(A_Clipboard)
    if (raw = "") {
        MsgBox("Clipboard empty. Copy the raw lead or FORMMAP block first.")
        return
    }

    fields := NormalizeProspectInput(raw)

    if !FocusEdge() {
        MsgBox("Microsoft Edge not found. Open the target prospect page first.")
        return
    }

    BeginAutomationRun()

    ToolTip("Click FIRST NAME now (2s)")
    Sleep 500
    ToolTip()

    FillNewProspectForm(fields)

    email := Trim(fields["EMAIL"])
    if IsEmailToken(email) && !SetClip(email)
        MsgBox("Failed to copy the lead email to the clipboard.")
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

^!b:: {
    global batchLeadHolder, batchLogFile
    global batchResumeIndex, batchResumeMode, batchResumeRaw

    raw := Trim(A_Clipboard)
    if (raw = "") {
        MsgBox("Clipboard empty. Copy the batch first.")
        return
    }

    batchLeadHolder := BuildBatchLeadHolder(raw)

    if (batchLeadHolder.Length = 0) {
        MsgBox("No lead rows detected in clipboard.")
        return
    }

    if !FocusWorkBrowser() {
        MsgBox("Browser not detected. Open the CRM page first.")
        return
    }

    BeginAutomationRun()
    if !SafeSleep(300)
        return

    log := []
    okCount := 0
    failCount := 0
    pausedAt := 0
    stopped := false

    try {
        EnsureBatchLogHeader()
    } catch Error as err {
        MsgBox("Cannot write to batch log file:`n" batchLogFile "`n`nClose Excel or any app using it, then try again.`n`nDetails: " err.Message)
        return
    }

    startIndex := 1
    if (batchResumeMode = "B" && batchResumeRaw = raw && batchResumeIndex >= 1 && batchResumeIndex <= batchLeadHolder.Length)
        startIndex := batchResumeIndex
    else {
        batchResumeIndex := 1
        batchResumeMode := ""
        batchResumeRaw := ""
    }

    Loop batchLeadHolder.Length - startIndex + 1 {
        if StopRequested() {
            stopped := true
            break
        }
        i := startIndex + A_Index - 1
        lead := batchLeadHolder[i]
        status := RunBatchLeadFlow(lead)

        try {
            AppendBatchLog(lead, status)
        } catch Error as err {
            MsgBox("Batch ran, but logging failed for:`n" lead["FULL_NAME"] "`n`nClose the CSV if it's open.`n`nDetails: " err.Message)
            return
        }

        log.Push(i . ". " . lead["FULL_NAME"] . " -> " . status)

        if (status = "OK")
            okCount += 1
        else
            failCount += 1

        if (status = "STOPPED") {
            stopped := true
            break
        }

        if (SubStr(status, 1, 8) = "PAUSED -") {
            pausedAt := i
            if (i < batchLeadHolder.Length) {
                batchResumeIndex := i + 1
                batchResumeMode := "B"
                batchResumeRaw := raw
            } else {
                batchResumeIndex := 1
                batchResumeMode := ""
                batchResumeRaw := ""
            }
            break
        }
    }

    msg := stopped
        ? "Batch stopped by ESC.`n`n"
        : pausedAt
        ? "Batch paused at lead " pausedAt ".`nRe-run Ctrl+Alt+B to continue with the next lead.`n`n"
        : "Batch complete.`n`n"

    msg .= "Success: " okCount "`n"
        . "Failed/Skipped: " failCount "`n`n"

    for _, line in log
        msg .= line "`n"

    MsgBox(msg, pausedAt ? "Batch Run Log (Paused)" : "Batch Run Log")
}

^!n:: {
    global batchLeadHolder, batchLogFile
    global batchResumeIndex, batchResumeMode, batchResumeRaw

    raw := Trim(A_Clipboard)
    if (raw = "") {
        MsgBox("Clipboard empty. Copy the batch first.")
        return
    }

    batchLeadHolder := BuildBatchLeadHolder(raw)

    if (batchLeadHolder.Length = 0) {
        MsgBox("No lead rows detected in clipboard.")
        return
    }

    if !FocusWorkBrowser() {
        MsgBox("Browser not detected. Open the CRM page first.")
        return
    }

    BeginAutomationRun()
    if !SafeSleep(300)
        return

    log := []
    okCount := 0
    failCount := 0
    pausedAt := 0
    stopped := false

    try {
        EnsureBatchLogHeader()
    } catch Error as err {
        MsgBox("Cannot write to batch log file:`n" batchLogFile "`n`nClose Excel or any app using it, then try again.`n`nDetails: " err.Message)
        return
    }

    startIndex := 1
    if (batchResumeMode = "N" && batchResumeRaw = raw && batchResumeIndex >= 1 && batchResumeIndex <= batchLeadHolder.Length)
        startIndex := batchResumeIndex
    else {
        batchResumeIndex := 1
        batchResumeMode := ""
        batchResumeRaw := ""
    }

    Loop batchLeadHolder.Length - startIndex + 1 {
        if StopRequested() {
            stopped := true
            break
        }
        i := startIndex + A_Index - 1
        lead := batchLeadHolder[i]
        status := RunBatchLeadFlowFast(lead)

        try {
            AppendBatchLog(lead, status)
        } catch Error as err {
            MsgBox("Batch ran, but logging failed for:`n" lead["FULL_NAME"] "`n`nClose the CSV if it's open.`n`nDetails: " err.Message)
            return
        }

        log.Push(i . ". " . lead["FULL_NAME"] . " -> " . status)

        if (status = "OK")
            okCount += 1
        else
            failCount += 1

        if (status = "STOPPED") {
            stopped := true
            break
        }

        if (SubStr(status, 1, 8) = "PAUSED -") {
            pausedAt := i
            if (i < batchLeadHolder.Length) {
                batchResumeIndex := i + 1
                batchResumeMode := "N"
                batchResumeRaw := raw
            } else {
                batchResumeIndex := 1
                batchResumeMode := ""
                batchResumeRaw := ""
            }
            break
        }
    }

    msg := stopped
        ? "Batch stopped by ESC.`n`n"
        : pausedAt
        ? "Batch paused at lead " pausedAt ".`nRe-run Ctrl+Alt+N to continue with the next lead.`n`n"
        : "Batch complete (FAST).`n`n"

    msg .= "Success: " okCount "`n"
        . "Failed/Skipped: " failCount "`n`n"

    for _, line in log
        msg .= line "`n"

    MsgBox(msg, pausedAt ? "Batch Run Log (Fast Paused)" : "Batch Run Log (Fast)")
}

^!t:: {
    global configDays, holidays

    if (configDays.Length = 0) {
        MsgBox("configDays is empty.")
        return
    }

    lastDay := 0
    for _, d in configDays
        if (d > lastDay)
            lastDay := d

    lastDate := BusinessDateForDay(lastDay, holidays)
    dtText := lastDate . " 3:00 PM"

    BeginAutomationRun()
    if !SetClip(dtText) {
        MsgBox("Failed to set clipboard.")
        return
    }

    Sleep 80
    Send "^v"
    Sleep 120

    SendTabs(6)
    Sleep 120
    Send "e"
    Sleep 120

    SendTabs(3)
    Sleep 120
    Send "c"
}
^!y:: {
    global holidays

    tomorrow := BusinessDateForDay(1, holidays)
    dtText := tomorrow . " 10:00 AM"

    BeginAutomationRun()
    if !SetClip(dtText) {
        MsgBox("Failed to set clipboard.")
        return
    }

    Sleep 80
    Send "^v"
    Sleep 120

    SendTabs(6)
    Sleep 120
    Send "P"
    Sleep 120
    Send "P"
    Sleep 120

    SendTabs(3)
    Sleep 120
    Send "c"
}

^!k:: {
    global configDays, holidays

    noteText := "txt"

    if (configDays.Length = 0) {
        MsgBox("configDays is empty.")
        return
    }

    lastDay := 0
    for _, d in configDays
        if (d > lastDay)
            lastDay := d

    lastDate := BusinessDateForDay(lastDay, holidays)
    dtText := lastDate . " 3:00 PM"

    if !FocusWorkBrowser() {
        MsgBox("Browser not detected.")
        return
    }

    BeginAutomationRun()

    ; Focus Action dropdown
    JS_FocusActionDropdown()
    Sleep 500

    ; Lead Update / Attempted Contact
    SendEvent "l"
    Sleep 150
    SendEvent "{Tab}"
    Sleep 150
    SendEvent "{Tab}"
    Sleep 250
    SendEvent "1"
    Sleep 250

    ; Note field
    SendTabs(9)
    Sleep 200

    if !SetClip(noteText) {
        MsgBox("Clipboard failed (note text).")
        return
    }

    Sleep 80
    SendEvent "^v"
    Sleep 250

    ; Save History Note
    JS_SaveHistoryNote()
    Sleep 800

    ; New Appointment
    JS_AddNewAppointment()
    Sleep 800

    ; Focus time field
    JS_FocusDateTimeField()
    Sleep 300

    if !SetClip(dtText) {
        MsgBox("Clipboard failed (date).")
        return
    }

    Sleep 80
    SendEvent "^v"
    Sleep 220

    SendTabs(6)
    Sleep 150
    SendEvent "e"
    Sleep 150
    SendTabs(3)
    Sleep 150
    SendEvent "c"
    Sleep 250

    ; Final Add New Appointment
    JS_SaveAppointment()
    Sleep 400
}

^!j:: {
    global holidays

    noteText := "qt"

    tomorrow := BusinessDateForDay(1, holidays)
    dtText := tomorrow . " 10:00 AM"

    if !FocusWorkBrowser() {
        MsgBox("Browser not detected.")
        return
    }

    BeginAutomationRun()

    ; Focus Action dropdown
    JS_FocusActionDropdown()
    Sleep 500

    ; Lead Update / Phone Call (l, Tab, q, Tab, 3)
    SendEvent "l"
    Sleep 150
    SendEvent "{Tab}"
    Sleep 150
    SendEvent "q"
    Sleep 150
    SendEvent "+{Tab}"
    Sleep 3050
    SendEvent "{Tab}"
    Sleep 250
    SendEvent "{Tab}"
    Sleep 250
    SendEvent "3"
    Sleep 250

    ; Note field
    SendTabs(9)
    Sleep 200

    if !SetClip(noteText) {
        MsgBox("Clipboard failed (note text).")
        return
    }

    Sleep 80
    SendEvent "^v"
    Sleep 250

    ; Save History Note
    JS_SaveHistoryNote()
    Sleep 800

    ; New Appointment
    JS_AddNewAppointment()
    Sleep 800

    ; Focus time field
    JS_FocusDateTimeField()
    Sleep 300

    if !SetClip(dtText) {
        MsgBox("Clipboard failed (date).")
        return
    }

    Sleep 80
    SendEvent "^v"
    Sleep 220

    SendTabs(6)
    Sleep 150
    SendEvent "p"
    Sleep 150
    SendEvent "p"
    Sleep 150
    SendTabs(3)
    Sleep 150
    SendEvent "c"
    Sleep 250

    ; Final Add New Appointment
    JS_SaveAppointment()
    Sleep 400
}

F8:: {  ; Start/Stop toggle
    global running
    running := !running

    if running {
        BeginAutomationRun()
        ToolTip("RUNNING (F8 to stop)")
        SetTimer(SpamLoop, 200)  ; adjust speed (ms)
    } else {
        SetTimer(SpamLoop, 0)
        ToolTip("STOPPED")
        Sleep 800
        ToolTip()
    }
}

; ===================== PERSISTENCE / INI HELPERS =====================

EnsureIniDefaults() {
    global iniFile

    defaults := Map(
        "Agent.Name", "Pablo Cabrera",
        "Agent.Email", "pablocabrera@allstate.com",
        "Agent.TagSymbol", "+",
        "Schedule.Days", "1,2,4,5",
        "Times.Offset", "0",
        "Batch.MinVehicles", "0",
        "Batch.MaxVehicles", "99",
        "DOB.DefaultDay", "16",
        "UI.BatchTabsChatToName", "7",
        "Pricing.OldCar", "98",
        "Pricing.OneCar", "117",
        "Pricing.OneCar2020Plus", "167",
        "Pricing.TwoCars", "176",
        "Pricing.ThreeCars", "284",
        "Pricing.FourCars", "397",
        "Pricing.FiveCars", "397",
        "Pricing.SingleCarModernYearCutoff", "2020"
    )

    for key, val in defaults {
        parts := StrSplit(key, ".")
        section := parts[1]
        setting := parts[2]

        existing := IniRead(iniFile, section, setting, "")
        if (existing = "")
            IniWrite(val, iniFile, section, setting)
    }
}

CsvEscape(value) {
    text := value ?? ""
    text := StrReplace(text, '"', '""')
    return '"' . text . '"'
}

EnsureBatchLogHeader() {
    global batchLogFile
    if !FileExist(batchLogFile) {
        FileAppend("Timestamp,LeadName,Phone,CarCount,Status`n", batchLogFile, "UTF-8")
    }
}

AppendBatchLog(lead, status) {
    global batchLogFile

    timestamp := FormatTime(A_Now, "yyyy-MM-dd HH:mm:ss")
    line := CsvEscape(timestamp) ","
        . CsvEscape(lead["FULL_NAME"]) ","
        . CsvEscape(lead["PHONE"]) ","
        . CsvEscape(lead["VEHICLE_COUNT"]) ","
        . CsvEscape(status) "`n"

    FileAppend(line, batchLogFile, "UTF-8")
}

ParseDays(str) {
    parts := StrSplit(str, ",")
    out := []
    for p in parts {
        p := Trim(p)
        if !(p ~= "^\d+$")
            continue
        v := Integer(p)
        if (v > 0)
            out.Push(v)
    }
    return out
}

NextRotationOffset() {
    global iniFile
    idx := IniRead(iniFile, "Times", "Offset", 0)
    idx := Mod(idx + 1, 60)
    IniWrite(idx, iniFile, "Times", "Offset")
    return idx
}

; ===================== LOW-LEVEL INPUT / AUTOMATION HELPERS =====================

SetClip(text) {
    if StopRequested()
        return false
    A_Clipboard := ""
    if !SafeSleep(30)
        return false
    if StopRequested()
        return false
    A_Clipboard := text
    return WaitForClip(1000)
}

PasteValue(text) {
    text := text ?? ""
    if (text = "" || StopRequested())
        return false
    if !SetClip(text)
        return false
    if !SafeSleep(60)
        return false
    if StopRequested()
        return false
    Send "{Backspace}"
    if !SafeSleep(60)
        return false
    if StopRequested()
        return false
    Send "^v"
    if !SafeSleep(90)
        return false
    return true
}

PasteValueRaw(text) {
    text := text ?? ""
    if (text = "" || StopRequested())
        return false
    if !SetClip(text)
        return false
    if !SafeSleep(60)
        return false
    if StopRequested()
        return false
    Send "^v"
    if !SafeSleep(90)
        return false
    return true
}

SendTabs(count) {
    Loop count {
        if StopRequested()
            return false
        Send "{Tab}"
        if !SafeSleep(50)
            return false
    }
    return true
}

FastType(value) {
    value := value ?? ""
    if (value = "")
        return true
    SendText value
    return true
}

PasteField(value) {
    value := value ?? ""
    Send "^a"
    Sleep 60

    if (value = "")
        return true

    return PasteValue(value)
}

PasteFieldSafe(value) {
    value := value ?? ""
    if (value = "")
        return true

    if !SetClip(value)
        return false

    Sleep 60
    Send "^v"
    Sleep 100
    return true
}

SelectDropdownValue(value) {
    value := Trim(value)
    if (value = "") {
        Send "{Tab}"
        Sleep 90
        return
    }

    SendText value
    Sleep 120
    Send "{Tab}"
    Sleep 100
}

SortMessagesByDaySeq(arr) {
    if (arr.Length <= 1)
        return arr

    Loop arr.Length - 1 {
        i := A_Index + 1
        current := arr[i]
        j := i - 1

        while (j >= 1) {
            left := arr[j]

            shouldMove := (left["day"] > current["day"])
                || ((left["day"] = current["day"]) && (left["seq"] > current["seq"]))

            if !shouldMove
                break

            arr[j + 1] := arr[j]
            j -= 1
        }

        arr[j + 1] := current
    }

    return arr
}

; ===================== UI / USER INTERACTION HELPERS =====================

SendSelectedBatch(picker) {
    global holidays

    try picker.Opt("-AlwaysOnTop")
    try picker.Hide()

    if !FocusWorkBrowser() {
        MsgBox("Browser not detected.")
        return
    }

    ToolTip("Click en el cuadro de Quo para comenzar")
    KeyWait "LButton"
    KeyWait "LButton", "D"
    Sleep 150
    ToolTip()

    useD1 := picker.cbD1.Value
    useD2 := picker.cbD2.Value
    useD4 := picker.cbD4.Value
    useD5 := picker.cbD5.Value
    modeText := picker.ddMode.Text
    fullMsgs := picker.fullMsgs

    selectedDays := []
    if (useD1) selectedDays.Push(fullMsgs[1]["day"])
    if (useD2) selectedDays.Push(fullMsgs[5]["day"])
    if (useD4) selectedDays.Push(fullMsgs[7]["day"])
    if (useD5) selectedDays.Push(fullMsgs[9]["day"])

    if (selectedDays.Length = 0) {
        MsgBox("Selecciona al menos un bloque (A, B, C o D).")
        return
    }

    baseDay := 999
    for d in selectedDays
        if (d < baseDay)
            baseDay := d

    todayYMD := FormatTime(A_Now, "yyyyMMdd")
    baseYMD  := NextBusinessDateYYYYMMDD(todayYMD, holidays)

    toSend := []
    for m in fullMsgs {
        if (m.Has("day") && ArrContains(selectedDays, m["day"])) {
            daysAfter := m["day"] - baseDay
            targetYMD := AddBusinessDays(baseYMD, daysAfter, holidays)
            m["date"] := FormatTime(targetYMD . "000000", "MM/dd/yyyy")
            toSend.Push(m)
        }
    }

    SortMessagesByDaySeq(toSend)

    useTyped := (modeText = "Escritura estable (Chrome)")
    for m in toSend {
        FocusWorkBrowser()
        Sleep 120
        ok := useTyped
            ? ScheduleMessageTyped(m["text"], m["date"], m["time"])
            : ScheduleMessage(m["text"], m["date"], m["time"])
        if !ok {
            MsgBox("Falló el envío de uno de los mensajes. Se detiene el proceso.")
            return
        }
        Sleep 250
    }

    try picker.Destroy()
    TrayTip("AHK", "Mensajes programados: " . toSend.Length, 1)
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

; ===================== BROWSER / DOM HELPERS =====================

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
    originalClip := A_Clipboard

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

JS_FocusActionDropdown() {
    RunDevToolsJS("(() => { let d = document.querySelectorAll('iframe')[0]?.contentDocument; let el = d?.getElementById('ctl00_ContentPlaceHolder1_DDLogType_Input'); if (!el) return 'NO_ACTION'; el.focus(); ['mouseover','mousedown','mouseup','click'].forEach(t => el.dispatchEvent(new MouseEvent(t,{bubbles:true,cancelable:true,view:d.defaultView}))); return 'OK_ACTION'; })()")
}

JS_SaveHistoryNote() {
    RunDevToolsJS("(() => { let d = document.querySelectorAll('iframe')[0]?.contentDocument; let el = d?.getElementById('ctl00_ContentPlaceHolder1_btnUpdate_input'); if (!el) return 'NO_SAVE'; el.focus(); ['mouseover','mousedown','mouseup','click'].forEach(t => el.dispatchEvent(new MouseEvent(t,{bubbles:true,cancelable:true,view:d.defaultView}))); return 'OK_SAVE'; })()")
}

JS_AddNewAppointment() {
    RunDevToolsJS("(() => { let d = document.querySelectorAll('iframe')[0]?.contentDocument; if (!d) return 'NO_FRAME1'; if (typeof d.defaultView.AppointmentInserting === 'function') { d.defaultView.AppointmentInserting(); return 'OK_FUNC'; } let el = d.querySelector('a.js-Lead-Log-Add-New-Appointment'); if (!el) return 'NO_APPT'; el.focus(); ['mouseover','mousedown','mouseup','click'].forEach(t => el.dispatchEvent(new MouseEvent(t,{bubbles:true,cancelable:true,view:d.defaultView}))); return 'OK_APPT'; })()")
}

JS_FocusDateTimeField() {
    RunDevToolsJS("(() => { let d1 = document.querySelectorAll('iframe')[0]?.contentDocument; let d2 = d1?.querySelectorAll('iframe')[0]?.contentDocument; let el = d2?.getElementById('ctl00_ContentPlaceHolder1_RadDateTimePicker1_dateInput'); if (!el) return 'NO_TIME'; el.focus(); el.select(); return 'OK_TIME'; })()")
}

JS_SaveAppointment() {
    RunDevToolsJS("(() => { let d1 = document.querySelectorAll('iframe')[0]?.contentDocument; let d2 = d1?.querySelectorAll('iframe')[0]?.contentDocument; let el = d2?.getElementById('ctl00_ContentPlaceHolder1_lnkSave_input'); if (!el) return 'NO_FINAL'; el.focus(); ['mouseover','mousedown','mouseup','click'].forEach(t => el.dispatchEvent(new MouseEvent(t,{bubbles:true,cancelable:true,view:d2.defaultView}))); return 'OK_FINAL'; })()")
}

ClickAddNewAppointmentJS() {
    Send "^+j"
    Sleep 600
    SendText "document.getElementById('ctl00_ContentPlaceHolder1_lnkSave_input')?.click()"
    Sleep 150
    Send "{Enter}"
    Sleep 300
    Send "^+j"
    Sleep 400
}

FocusEdge() {
    if WinExist("ahk_exe msedge.exe") {
        WinActivate
        WinWaitActive "ahk_exe msedge.exe",, 2
        return true
    }
    return false
}

FocusChrome() {
    if WinExist("ahk_exe chrome.exe") {
        WinActivate
        WinWaitActive "ahk_exe chrome.exe",, 2
        return true
    }
    return false
}

FocusWorkBrowser() {
    if FocusChrome()
        return true
    if FocusEdge()
        return true
    return false
}

FocusSlateComposer() {
    return RunDevToolsJS("(() => { let el = document.querySelector('[data-slate-editor=`"true`"]'); if (!el) return 'NO_COMPOSER'; el.focus(); return 'OK_COMPOSER'; })()")
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

FocusSlateComposerReady() {
    if StopRequested()
        return false
    result := RunDevToolsJSGetResult("copy(String((() => { const editors = Array.from(document.querySelectorAll('[data-slate-editor=true]')); const el = editors.find(node => node.offsetParent !== null); if (!el) return 'NO_COMPOSER'; el.focus(); el.click(); return 'OK_COMPOSER'; })()));")
    return (result = "OK_COMPOSER")
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

; ===================== TEXT / ARRAY SUPPORT HELPERS =====================

CleanCityCol(city) {
    city := Trim(city)
    city := RegExReplace(city, ",\s*[A-Z]{2}$", "")
    city := RegExReplace(city, "[,\s]+$", "")
    return ProperCasePhrase(city)
}

GetMaxConfiguredDay(daysArr) {
    maxDay := 0
    for _, d in daysArr
        if (d > maxDay)
            maxDay := d
    return maxDay
}

CleanName(str) {
    str := Trim(str)
    str := StrReplace(str, "`r")
    str := StrReplace(str, "`n")
    return str
}

ProperCase(str) {
    s := StrLower(Trim(str))
    parts := StrSplit(s, A_Space)
    out := ""
    for _, p in parts {
        if (p = "")
            continue
        out .= (out != "" ? " " : "") . StrUpper(SubStr(p, 1, 1)) . SubStr(p, 2)
    }
    return out
}

ExtractFirstName(str) {
    parts := StrSplit(Trim(str), " ")
    return (parts.Length >= 1) ? parts[1] : str
}

ProperCasePhrase(str) {
    text := Trim(RegExReplace(str, "\s+", " "))
    if (text = "")
        return ""
    parts := StrSplit(text, " ")
    out := ""
    for _, part in parts {
        if (part = "")
            continue
        out .= (out != "" ? " " : "") . ProperCaseWord(part)
    }
    return out
}

ProperCaseWord(word) {
    clean := Trim(word)
    if (clean = "")
        return ""

    if RegExMatch(clean, "^(?:N|S|E|W|NE|NW|SE|SW)$")
        return StrUpper(clean)

    if RegExMatch(clean, "^\d+[A-Za-z]{2}$", &m) {
        prefix := RegExReplace(clean, "[A-Za-z]{2}$", "")
        suffix := RegExMatch(clean, "[A-Za-z]{2}$", &sfx) ? StrLower(sfx[0]) : ""
        return prefix . suffix
    }

    parts := StrSplit(StrLower(clean), "-")
    out := ""
    for i, piece in parts
        out .= (i > 1 ? "-" : "") . StrUpper(SubStr(piece, 1, 1)) . SubStr(piece, 2)
    return out
}

JoinArray(arr, delim := "") {
    out := ""
    for i, item in arr
        out .= (i > 1 ? delim : "") . item
    return out
}

ArrContains(arr, val) {
    for v in arr
        if (v = val)
            return true
    return false
}

; ===================== PROSPECT FIELD BUILDERS =====================

NewProspectFields() {
    return Map(
        "FIRST_NAME", "",
        "LAST_NAME", "",
        "DOB", "",
        "GENDER", "N",
        "ADDRESS_1", "",
        "APT_SUITE", "",
        "BUILDING", "",
        "RR_NUMBER", "",
        "LOT_NUMBER", "",
        "CITY", "",
        "STATE", "",
        "ZIP", "",
        "PHONE", "",
        "EMAIL", "",
        "RAW_NAME", ""
    )
}

; ===================== LEAD PARSING / NORMALIZATION HELPERS =====================

IsLabeledLeadFormat(raw) {
    return RegExMatch(raw, "m)^\s*Name:\s*$")
        || RegExMatch(raw, "m)^\s*Address Line 1:\s*$")
        || RegExMatch(raw, "m)^\s*Date of Birth::\s*$")
        || RegExMatch(raw, "m)^\s*First Name::\s*$")
}

NormalizeProspectInput(raw) {
    if RegExMatch(raw, "m)^\s*FORMMAP:\s*CREATE_NEW_PROSPECT_V1\s*$") || RegExMatch(raw, "m)^\s*[A-Z_]+=")
        return ParseProspectFieldBlock(raw)

    if IsLabeledLeadFormat(raw)
        return ParseLabeledLeadToProspect(raw)

    if IsLikelyBatchGridInput(raw)
        return ParseBatchGridRow(ExtractFirstBatchGridRow(raw))

    if RegExMatch(raw, "i)PERSONAL\s+LEAD")
        return ParseBatchCRMToProspect(raw)

    return NormalizeRawLeadToProspect(raw)
}

IsLikelyBatchGridInput(raw) {
    firstRow := ExtractFirstBatchGridRow(raw)
    if (firstRow = "")
        return false

    cols := StrSplit(firstRow, "`t")
    return cols.Length >= 11
        && RegExMatch(Trim(cols[1]), "i)PERSONAL\s+LEAD")
        && IsTimestampToken(cols[4])
        && NormalizeZip(cols[8]) != ""
        && NormalizePhone(cols[9]) != ""
}

ExtractFirstBatchGridRow(raw) {
    clean := StrReplace(raw, "`r", "")
    for _, line in StrSplit(clean, "`n") {
        line := Trim(line)
        if (line = "")
            continue
        if InStr(line, "`t") && RegExMatch(line, "i)PERSONAL\s+LEAD")
            return line
    }
    return ""
}

ParseBatchCRMToProspect(raw) {
    if IsLikelyBatchGridInput(raw)
        return ParseBatchGridRow(ExtractFirstBatchGridRow(raw))

    fields := NewProspectFields()

    ; If multiple leads, take the first one
    rows := ParseBatchLeadRows(raw)
    lead := (rows.Length > 0) ? rows[1] : raw

    ; Name
    batchName := ExtractBatchName(lead)
    if (batchName != "")
        ApplyLeadName(fields, batchName)

    ; Phone
    fields["PHONE"] := ExtractBatchPhone(lead)
    fields["EMAIL"] := ExtractFirstEmail(lead)

    ; Gender
    if RegExMatch(lead, "i)\b(Male|Female)\b", &mg)
        fields["GENDER"] := NormalizeGender(mg[1])

    ; DOB: text between phone/email area and Male/Female
    dobText := ""
    if RegExMatch(lead, "i)\(\d{3}\)\s*\d{3}-\d{4}\s*(?:[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}\s*)?(.+?)(?=\s*(?:Male|Female)\b)", &md)
        dobText := Trim(md[1])
    if (dobText != "")
        fields["DOB"] := NormalizeDOB(dobText)

    ; Location: block between timestamp and phone
    locationBlock := ""
    if RegExMatch(lead, "i)\d{1,2}/\d{1,2}/\d{4}\s+\d{1,2}:\d{2}:\d{2}\s*(?:AM|PM)\s*(.+?)(?=\(\d{3}\))", &ml)
        locationBlock := Trim(ml[1])

    if (locationBlock != "")
    ParseBatchLocationBlock(locationBlock, fields)

    NormalizeAddressMap(fields)
    fields["DOB"]    := NormalizeDOB(fields["DOB"])
    fields["GENDER"] := NormalizeGender(fields["GENDER"])
    fields["STATE"]  := NormalizeState(fields["STATE"])
    fields["ZIP"]    := NormalizeZip(fields["ZIP"])
    fields["PHONE"]  := NormalizePhone(fields["PHONE"])

    return fields
}

ParseBatchLocationBlock(locBlock, fields) {
    stateFullNames := "Alabama|Alaska|Arizona|Arkansas|California|Colorado|Connecticut|Delaware|Florida|Georgia|Hawaii|Idaho|Illinois|Indiana|Iowa|Kansas|Kentucky|Louisiana|Maine|Maryland|Massachusetts|Michigan|Minnesota|Mississippi|Missouri|Montana|Nebraska|Nevada|New Hampshire|New Jersey|New Mexico|New York|North Carolina|North Dakota|Ohio|Oklahoma|Oregon|Pennsylvania|Rhode Island|South Carolina|South Dakota|Tennessee|Texas|Utah|Vermont|Virginia|Washington|West Virginia|Wisconsin|Wyoming"
    pattern := "i)(" . stateFullNames . ")\s*(\d{5})"

    pos := 1
    lastState := ""
    lastZip := ""
    lastPos := 0

    while RegExMatch(locBlock, pattern, &ms, pos) {
        lastState := ms[1]
        lastZip := ms[2]
        lastPos := ms.Pos
        pos := ms.Pos + ms.Len
    }

    if (lastPos = 0)
        return

    fields["STATE"] := NormalizeState(lastState)
    fields["ZIP"]   := lastZip

    beforeState := Trim(SubStr(locBlock, 1, lastPos - 1))
    beforeState := RegExReplace(beforeState, "[,\s]+$", "")

    stateAbbrs := "AL|AK|AZ|AR|CA|CO|CT|DC|DE|FL|GA|HI|IA|ID|IL|IN|KS|KY|LA|MA|MD|ME|MI|MN|MO|MS|MT|NC|ND|NE|NH|NJ|NM|NV|NY|OH|OK|OR|PA|RI|SC|SD|TN|TX|UT|VA|VT|WA|WI|WV|WY"

    if RegExMatch(beforeState, "i)^(.*?)\s*,\s*(" . stateAbbrs . ")\s*(?:\d{5}(?:-\d{4})?)?\s*(.*)$", &mc) {
        addrPart := Trim(mc[1])
        cityPart := Trim(mc[3])

        if (cityPart = "") {
            streetOnly := ""
            cityFromAddr := ""
            SplitStreetAndTrailingCity(addrPart, &streetOnly, &cityFromAddr)
            if (cityFromAddr != "") {
                addrPart := streetOnly
                cityPart := cityFromAddr
            }
        }

        if (cityPart != "") {
            cityPart := RegExReplace(cityPart, "^[,\s]+|[,\s]+$", "")
            fields["CITY"] := ProperCasePhrase(cityPart)

            if (fields["CITY"] != "" && RegExMatch(addrPart, "i)\s+" . cityPart . "$"))
                addrPart := Trim(RegExReplace(addrPart, "i)\s+" . cityPart . "$", ""))
        }

        SetAddressFields(fields, addrPart)
    } else {
        streetOnly := ""
        cityFromAddr := ""
        SplitStreetAndTrailingCity(beforeState, &streetOnly, &cityFromAddr)

        if (cityFromAddr != "")
            fields["CITY"] := ProperCasePhrase(cityFromAddr)

        SetAddressFields(fields, streetOnly != "" ? streetOnly : beforeState)
    }
}

NormalizeLeadLabel(label) {
    label := Trim(label)
    label := RegExReplace(label, "\s+", " ")
    label := RegExReplace(label, ":+$", "")
    return Trim(label)
}

StoreLabeledField(data, label, value) {
    label := NormalizeLeadLabel(label)
    value := Trim(value)

    if (label = "" || value = "")
        return

    if RegExMatch(label, "i)^Open the calendar popup\.?$")
        return

    if !data.Has(label) || data[label] = ""
        data[label] := value
}

ParseLabeledLeadToProspect(raw) {
    fields := NewProspectFields()
    lines := StrSplit(StrReplace(raw, "`r", ""), "`n")

    data := Map()
    currentLabel := ""
    currentValue := ""

    FlushPair := (*) => (currentLabel != "" ? StoreLabeledField(data, currentLabel, currentValue) : 0)

    for _, line in lines {
        line := Trim(line)
        if (line = "")
            continue

        if RegExMatch(line, "^([^:]+:?):\s*$", &m) {
            FlushPair()
            currentLabel := Trim(m[1], " :")
            currentValue := ""
            continue
        }

        if RegExMatch(line, "^([^:]+:?):\s*(.+)$", &m) {
            FlushPair()
            currentLabel := Trim(m[1], " :")
            currentValue := Trim(m[2])
            FlushPair()
            currentLabel := ""
            currentValue := ""
            continue
        }

        if (currentLabel != "" && currentValue = "")
            currentValue := line
        else if (currentLabel != "")
            currentValue .= " " line
    }

    FlushPair()

    if data.Has("First Name")
        fields["FIRST_NAME"] := ProperCasePhrase(data["First Name"])
    else if data.Has("Name")
        ApplyLeadName(fields, RegExReplace(data["Name"], "i)^PERSONAL LEAD\s*-\s*"))

    if data.Has("Last Name")
        fields["LAST_NAME"] := ProperCasePhrase(data["Last Name"])

    if ((fields["FIRST_NAME"] = "" || fields["LAST_NAME"] = "") && data.Has("Contact"))
        ApplyLeadName(fields, data["Contact"])

    if data.Has("Date of Birth")
        fields["DOB"] := NormalizeDOB(data["Date of Birth"])

    if data.Has("Gender")
        fields["GENDER"] := NormalizeGender(data["Gender"])

    if data.Has("Address Line 1")
        SetAddressFields(fields, data["Address Line 1"])

    if data.Has("City")
        fields["CITY"] := data["City"]

    if data.Has("State")
        fields["STATE"] := data["State"]

    if data.Has("Zip Code")
        fields["ZIP"] := data["Zip Code"]

    if data.Has("Phone")
        fields["PHONE"] := data["Phone"]

    fields["EMAIL"] := ExtractFirstEmail(raw)

    NormalizeAddressMap(fields)
    fields["DOB"]    := NormalizeDOB(fields["DOB"])
    fields["GENDER"] := NormalizeGender(fields["GENDER"])
    fields["STATE"]  := NormalizeState(fields["STATE"])
    fields["ZIP"]    := NormalizeZip(fields["ZIP"])
    fields["PHONE"]  := NormalizePhone(fields["PHONE"])

    return fields
}

ParseLabeledLeadRaw(raw) {
    lines := StrSplit(StrReplace(raw, "`r", ""), "`n")
    data := Map()
    currentLabel := ""
    currentValue := ""

    FlushPair := (*) => (currentLabel != "" ? StoreLabeledField(data, currentLabel, currentValue) : 0)

    for _, line in lines {
        line := Trim(line)
        if (line = "")
            continue

        if RegExMatch(line, "^(.+?):\s*$", &m) {
            FlushPair()
            currentLabel := m[1]
            currentValue := ""
            continue
        }

        if RegExMatch(line, "^(.+?):\s*(.+)$", &m) {
            FlushPair()
            currentLabel := m[1]
            currentValue := m[2]
            FlushPair()
            currentLabel := ""
            currentValue := ""
            continue
        }

        if (currentLabel != "" && currentValue = "")
            currentValue := line
        else if (currentLabel != "")
            currentValue .= " " line
    }

    FlushPair()
    return data
}

ParseProspectFieldBlock(raw) {
    fields := NewProspectFields()
    clean := StrReplace(raw, "`r", "")

    for line in StrSplit(clean, "`n") {
        line := Trim(line)
        if !RegExMatch(line, "^([A-Z_]+)=(.*)$", &m)
            continue
        key := m[1]
        if fields.Has(key)
            fields[key] := Trim(m[2])
    }

    fields["DOB"]    := NormalizeDOB(fields["DOB"])
    fields["GENDER"] := NormalizeGender(fields["GENDER"])
    fields["STATE"]  := NormalizeState(fields["STATE"])
    fields["ZIP"]    := NormalizeZip(fields["ZIP"])
    fields["PHONE"]  := NormalizePhone(fields["PHONE"])
    NormalizeAddressMap(fields)
    return fields
}

NormalizeRawLeadToProspect(raw) {
    fields := NewProspectFields()
    tokens := TokenizeLead(raw)

    if (tokens.Length = 0)
        return fields

    leadName := FindLeadName(tokens)
    if (leadName != "")
        ApplyLeadName(fields, leadName)

    fields["EMAIL"] := ExtractFirstEmail(raw)

    zipIdx := 0
    for i, token in tokens {
        if IsTimestampToken(token) || IsPhoneToken(token) || IsEmailToken(token)
            continue

        zip := NormalizeZip(token)
        if (zip = "")
            continue

        zipIdx := i
        fields["ZIP"] := zip
        break
    }

    if (zipIdx >= 2) {
        fields["STATE"] := NormalizeState(tokens[zipIdx - 1])
        if (zipIdx >= 3)
            fields["CITY"] := ProperCasePhrase(tokens[zipIdx - 2])
        addrIdx := FindAddressIndex(tokens, zipIdx - 2)
        if (addrIdx)
            SetAddressFields(fields, tokens[addrIdx])
    }

    for _, token in tokens {
        if (fields["PHONE"] = "" && !IsTimestampToken(token) && !IsEmailToken(token)) {
            phone := NormalizePhone(token)
            if (phone != "")
                fields["PHONE"] := phone
        }

        if !IsTimestampToken(token) {
            dob := NormalizeDOB(token)
            if (dob != "" && IsBetterDOBCandidate(dob, fields["DOB"]))
                fields["DOB"] := dob
        }

        if (fields["GENDER"] = "N") {
            gender := NormalizeGender(token)
            if (gender != "N" || RegExMatch(token, "i)^(?:male|female|m|f|non[- ]?binary|nonbinary|not specified|x)$"))
                fields["GENDER"] := gender
        }
    }

    if (fields["CITY"] = "") {
        for i, token in tokens {
            if (NormalizeState(token) != "" && i > 1) {
                fields["CITY"] := ProperCasePhrase(tokens[i - 1])
                break
            }
        }
    }

    NormalizeAddressMap(fields)
    fields["GENDER"] := NormalizeGender(fields["GENDER"])
    return fields
}

TokenizeLead(raw) {
    text := Trim(StrReplace(raw, "`r", ""))
    if (text = "")
        return []

    rawTokens := InStr(text, "`t")
        ? StrSplit(text, "`t")
        : StrSplit(RegExReplace(text, "\s{2,}", "`n"), "`n")

    tokens := []
    for _, token in rawTokens {
        token := Trim(token)
        if (token != "")
            tokens.Push(token)
    }
    return tokens
}

FindLeadName(tokens) {
    for _, token in tokens {
        if RegExMatch(token, "i)personal lead\s*-\s*(.+)$", &m)
            return Trim(m[1])
    }

    for _, token in tokens {
        if IsTimestampToken(token) || IsPhoneToken(token) || IsEmailToken(token)
            continue
        if RegExMatch(token, "\d")
            continue
        if RegExMatch(token, "i)\b(?:folder|new|personal)\b")
            continue
        return token
    }
    return ""
}

ApplyLeadName(fields, fullName) {
    clean := RegExReplace(Trim(fullName), "\s+", " ")
    parts := StrSplit(clean, " ")
    if (parts.Length = 0)
        return
    fields["FIRST_NAME"] := ProperCasePhrase(parts[1])
    if (parts.Length >= 2)
        fields["LAST_NAME"] := ProperCasePhrase(parts[parts.Length])
}

FindAddressIndex(tokens, beforeIdx) {
    Loop beforeIdx {
        i := beforeIdx - A_Index + 1
        token := tokens[i]
        if IsAddressToken(token)
            return i
    }
    return 0
}

IsAddressToken(token) {
    return RegExMatch(token, "\d")
        && !IsTimestampToken(token)
        && !IsPhoneToken(token)
        && !IsEmailToken(token)
        && (NormalizeDOB(token) = "")
}

SetAddressFields(fields, rawAddress) {
    street := ""
    unit := ""
    SplitAddressAndUnit(rawAddress, &street, &unit)
    fields["ADDRESS_1"] := street
    if (fields["APT_SUITE"] = "")
        fields["APT_SUITE"] := unit
}

SplitAddressAndUnit(rawAddress, &street, &unit) {
    text := Trim(RegExReplace(rawAddress, "\s+", " "))
    street := text
    unit := ""

    if RegExMatch(text, "i)^(.*?)(?:\s+(?:apt|apartment|apart|unit|suite|ste)\.?\s*|\s+#\s*)([A-Za-z0-9\-]+(?:\s+[A-Za-z0-9\-]+)*)$", &m) {
        street := Trim(m[1], " ,")
        unit := Trim(m[2], " ,")
    }
}

SplitStreetAndTrailingCity(text, &street, &city) {
    street := Trim(text, " ,")
    city := ""

    suffixes := "Street|St|Avenue|Ave|Road|Rd|Drive|Dr|Boulevard|Blvd|Lane|Ln|Court|Ct|Circle|Cir|Way|Terrace|Ter|Trail|Trl|Parkway|Pkwy|Place|Pl|Highway|Hwy|Loop"

    if RegExMatch(
        street,
        "i)^(.*?\b(?:" . suffixes . ")\b(?:\s+(?:#|apt|apartment|apart|unit|suite|ste)\.?\s*[A-Za-z0-9\-]+)?)\s+([A-Za-z]+(?:\s+[A-Za-z]+){0,2})$",
        &m
    ) {
        street := Trim(m[1], " ,")
        city := Trim(m[2], " ,")
    }
}

NormalizeAddressMap(fields) {
    address1 := fields["ADDRESS_1"]
    city     := fields["CITY"]
    state    := fields["STATE"]
    zip      := fields["ZIP"]
    aptSuite := fields["APT_SUITE"]

    if (address1 != "") {
        ExtractAddressTail(&address1, &city, &state, &zip, &aptSuite)
        street := ""
        unit := ""
        SplitAddressAndUnit(address1, &street, &unit)
        address1 := street
        if (aptSuite = "")
            aptSuite := unit
    }

    NormalizeCityStateZipFields(&city, &state, &zip)

    if (state != "")
        address1 := Trim(RegExReplace(address1, "i)\s+,?\s*" state "\s*$", ""), " ,")

    address1 := StripTrailingStateAbbrev(address1)

    fields["ADDRESS_1"] := address1
    fields["CITY"]      := NormalizeCity(city)
    fields["STATE"]     := NormalizeState(state)
    fields["ZIP"]       := NormalizeZip(zip)
    fields["APT_SUITE"] := aptSuite
    fields["PHONE"]     := NormalizePhone(fields["PHONE"])
}

NormalizeCity(city) {
    city := Trim(city)
    if (city = "")
        return ""

    city := RegExReplace(city, "[,\.]+\s*$", "")
    city := RegExReplace(city, "\b\d{5}(?:-\d{4})?\b", "")
    city := RegExReplace(city, "i)\b(AL|AK|AZ|AR|CA|CO|CT|DC|DE|FL|GA|HI|IA|ID|IL|IN|KS|KY|LA|MA|MD|ME|MI|MN|MO|MS|MT|NC|ND|NE|NH|NJ|NM|NV|NY|OH|OK|OR|PA|RI|SC|SD|TN|TX|UT|VA|VT|WA|WI|WV|WY)\b\s*$", "")
    city := RegExReplace(city, "i)\b(alabama|alaska|arizona|arkansas|california|colorado|connecticut|delaware|district of columbia|florida|georgia|hawaii|idaho|illinois|indiana|iowa|kansas|kentucky|louisiana|maine|maryland|massachusetts|michigan|minnesota|mississippi|missouri|montana|nebraska|nevada|new hampshire|new jersey|new mexico|new york|north carolina|north dakota|ohio|oklahoma|oregon|pennsylvania|rhode island|south carolina|south dakota|tennessee|texas|utah|vermont|virginia|washington|west virginia|wisconsin|wyoming)\b\s*$", "")
    city := RegExReplace(city, "\s+", " ")
    city := Trim(city, " ,.-")
    return ProperCasePhrase(city)
}

NormalizeCityStateZipFields(&city, &state, &zip) {
    raw := Trim(city)
    if (raw = "")
        return

    if (zip = "" && RegExMatch(raw, "\b(\d{5})(?:-\d{4})?\b", &m))
        zip := m[1]

    if (state = "") {
        if RegExMatch(raw, "i)\b(AL|AK|AZ|AR|CA|CO|CT|DC|DE|FL|GA|HI|IA|ID|IL|IN|KS|KY|LA|MA|MD|ME|MI|MN|MO|MS|MT|NC|ND|NE|NH|NJ|NM|NV|NY|OH|OK|OR|PA|RI|SC|SD|TN|TX|UT|VA|VT|WA|WI|WV|WY)\b", &m2)
            state := StrUpper(m2[1])
        else {
            words := StrSplit(raw, ",")
            for _, w in words {
                st := NormalizeState(w)
                if (st != "") {
                    state := st
                    break
                }
            }
            if (state = "") {
                st := NormalizeState(raw)
                if (st != "")
                    state := st
            }
        }
    }

    city := NormalizeCity(raw)
}

ExtractAddressTail(&address1, &city, &state, &zip, &aptSuite) {
    text := Trim(address1)
    if (text = "")
        return

    if (zip = "" && RegExMatch(text, "\b(\d{5})(?:-\d{4})?\b", &mz))
        zip := mz[1]

    if (state = "") {
        if RegExMatch(text, "i)\b(AL|AK|AZ|AR|CA|CO|CT|DC|DE|FL|GA|HI|IA|ID|IL|IN|KS|KY|LA|MA|MD|ME|MI|MN|MO|MS|MT|NC|ND|NE|NH|NJ|NM|NV|NY|OH|OK|OR|PA|RI|SC|SD|TN|TX|UT|VA|VT|WA|WI|WV|WY)\b", &ms)
            state := StrUpper(ms[1])
        else {
            st := NormalizeState(text)
            if (st != "")
                state := st
        }
    }

    if (city = "") {
        if RegExMatch(text, "i),?\s*([A-Za-z]+(?:\s+[A-Za-z]+){0,2})\s*,?\s*(?:AL|AK|AZ|AR|CA|CO|CT|DC|DE|FL|GA|HI|IA|ID|IL|IN|KS|KY|LA|MA|MD|ME|MI|MN|MO|MS|MT|NC|ND|NE|NH|NJ|NM|NV|NY|OH|OK|OR|PA|RI|SC|SD|TN|TX|UT|VA|VT|WA|WI|WV|WY)\s*\d{5}(?:-\d{4})?$", &mc)
            city := mc[1]
    }

    text := RegExReplace(
        text,
        "i),?\s*[A-Za-z]+(?:\s+[A-Za-z]+){0,2}\s*,?\s*(AL|AK|AZ|AR|CA|CO|CT|DC|DE|FL|GA|HI|IA|ID|IL|IN|KS|KY|LA|MA|MD|ME|MI|MN|MO|MS|MT|NC|ND|NE|NH|NJ|NM|NV|NY|OH|OK|OR|PA|RI|SC|SD|TN|TX|UT|VA|VT|WA|WI|WV|WY)\s*\d{5}(?:-\d{4})?$",
        ""
    )
    text := Trim(RegExReplace(text, "\s+", " "), " ,")

    street := ""
    unit := ""
    SplitAddressAndUnit(text, &street, &unit)
    address1 := street
    if (aptSuite = "")
        aptSuite := unit
}

StripTrailingStateAbbrev(text) {
    clean := Trim(text)
    if (clean = "")
        return ""
    clean := RegExReplace(clean, "i)[,\s]+\b(AL|AK|AZ|AR|CA|CO|CT|DC|DE|FL|GA|HI|IA|ID|IL|IN|KS|KY|LA|MA|MD|ME|MI|MN|MO|MS|MT|NC|ND|NE|NH|NJ|NM|NV|NY|OH|OK|OR|PA|RI|SC|SD|TN|TX|UT|VA|VT|WA|WI|WV|WY)\b\s*$", "")
    return Trim(clean, " ,")
}

NormalizeState(state) {
    static states := Map(
        "alabama", "AL", "alaska", "AK", "arizona", "AZ", "arkansas", "AR",
        "california", "CA", "colorado", "CO", "connecticut", "CT", "delaware", "DE",
        "district of columbia", "DC", "florida", "FL", "georgia", "GA", "hawaii", "HI",
        "idaho", "ID", "illinois", "IL", "indiana", "IN", "iowa", "IA", "kansas", "KS",
        "kentucky", "KY", "louisiana", "LA", "maine", "ME", "maryland", "MD",
        "massachusetts", "MA", "michigan", "MI", "minnesota", "MN", "mississippi", "MS",
        "missouri", "MO", "montana", "MT", "nebraska", "NE", "nevada", "NV",
        "new hampshire", "NH", "new jersey", "NJ", "new mexico", "NM", "new york", "NY",
        "north carolina", "NC", "north dakota", "ND", "ohio", "OH", "oklahoma", "OK",
        "oregon", "OR", "pennsylvania", "PA", "rhode island", "RI", "south carolina", "SC",
        "south dakota", "SD", "tennessee", "TN", "texas", "TX", "utah", "UT",
        "vermont", "VT", "virginia", "VA", "washington", "WA", "west virginia", "WV",
        "wisconsin", "WI", "wyoming", "WY"
    )

    clean := StrLower(Trim(RegExReplace(state, "\.", "")))
    if (clean = "")
        return ""

    if (StrLen(clean) = 2) {
        abbr := StrUpper(clean)
        for _, val in states
            if (val = abbr)
                return abbr
    }

    return states.Has(clean) ? states[clean] : ""
}

NormalizeZip(zip) {
    text := Trim(zip)
    if (text = "")
        return ""
    if RegExMatch(text, "\b(\d{5})(?:-\d{4})?\b", &m)
        return m[1]
    return ""
}

IsPhoneToken(token) {
    return NormalizePhone(token) != ""
}

NormalizePhone(phone) {
    text := Trim(phone)
    if (text = "")
        return ""

    if !RegExMatch(text, "^\+?1?[\s\-\(\)\.]*\d{3}[\s\-\)\.]*\d{3}[\s\-\.]*\d{4}$")
        return ""

    digits := RegExReplace(text, "\D")
    if (StrLen(digits) = 11 && SubStr(digits, 1, 1) = "1")
        digits := SubStr(digits, 2)

    return (StrLen(digits) = 10) ? digits : ""
}

IsEmailToken(token) {
    token := Trim(token)
    return RegExMatch(token, "i)^[A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,}$")
}

ExtractFirstEmail(text) {
    text := Trim(text ?? "")
    if (text = "")
        return ""
    if RegExMatch(text, "i)([A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,})", &m)
        return Trim(m[1])
    return ""
}

IsTimestampToken(token) {
    return RegExMatch(Trim(token), "^\d{1,2}/\d{1,2}/\d{4}\s+\d{1,2}:\d{2}:\d{2}\s*(?:AM|PM)$")
}

NormalizeGender(gender) {
    clean := StrLower(Trim(gender))
    if (clean = "")
        return "N"
    if RegExMatch(clean, "^(?:male|m)$")
        return "M"
    if RegExMatch(clean, "^(?:female|f)$")
        return "F"
    if RegExMatch(clean, "^(?:non[- ]?binary|nonbinary|n)$")
        return "N"
    if RegExMatch(clean, "^(?:not specified|x)$")
        return "X"
    return "N"
}

IsBetterDOBCandidate(newDob, currentDob) {
    if (newDob = "")
        return false
    if (currentDob = "")
        return true

    if RegExMatch(currentDob, "^\d{2}/16/\d{4}$") && !RegExMatch(newDob, "^\d{2}/16/\d{4}$")
        return true

    return false
}

NormalizeDOB(dob) {
    global dobDefaultDay
    text := Trim(dob)
    if (text = "")
        return ""

    if RegExMatch(text, "\(([^()]*)\)", &mp) {
        inner := NormalizeDOB(mp[1])
        if (inner != "")
            return inner
    }

    work := StrLower(text)
    work := NormalizeMonthWords(work)

    work := RegExReplace(work, "i)\bborn\s+in\b", " ")
    work := RegExReplace(work, "i)\bborn\b", " ")
    work := RegExReplace(work, "i)\bnacid[oa]\s+en\b", " ")
    work := RegExReplace(work, "i)\bnaci[oó]\s+en\b", " ")
    work := RegExReplace(work, "i)\bconfirm\b", " ")

    work := RegExReplace(work, "i)^\s*age\s*:?\s*\d{1,3}\s*,?\s*", "")
    work := RegExReplace(work, "i)^\s*\d{1,3}\s*(?:años|anos)\s*,?\s*", "")
    work := RegExReplace(work, "i)^\s*\d{1,3}\s*,\s*", "")
    work := RegExReplace(work, "i)\s*/\s*\d{1,3}\s*(?:años|anos)\s*$", "")
    work := RegExReplace(work, "i)\s+\d{1,3}\s*(?:años|anos)\s*$", "")

    work := RegExReplace(work, "[\(\)]", " ")
    work := RegExReplace(work, "[\.,;:]+", " ")
    work := RegExReplace(work, "\s+", " ")
    work := Trim(work)

    if RegExMatch(work, "i)^(\d{1,2})[-\s](jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec|january|february|march|april|june|july|august|september|october|november|december)[-\s](\d{2,4})$", &m) {
        month := MonthNumber(m[2])
        if month
            return FormatDateString(month, Integer(m[1]), NormalizeYear(m[3]))
    }

    if RegExMatch(work, "i)^(\d{1,2})\s+(jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec|january|february|march|april|june|july|august|september|october|november|december)\s+(\d{4})$", &m) {
        month := MonthNumber(m[2])
        if month
            return FormatDateString(month, Integer(m[1]), Integer(m[3]))
    }

    if RegExMatch(work, "i)^(jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec|january|february|march|april|june|july|august|september|october|november|december)\s+(\d{1,2})\s+(\d{4})$", &m) {
        month := MonthNumber(m[1])
        if month
            return FormatDateString(month, Integer(m[2]), Integer(m[3]))
    }

    if RegExMatch(work, "i)^(jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec|january|february|march|april|june|july|august|september|october|november|december)[-\s](\d{2,4})$", &m) {
        month := MonthNumber(m[1])
        if month
            return FormatDateString(month, dobDefaultDay, NormalizeYear(m[2]))
    }

    if RegExMatch(work, "i)^(jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec|january|february|march|april|june|july|august|september|october|november|december)\s+(\d{4})$", &m) {
        month := MonthNumber(m[1])
        if month
            return FormatDateString(month, dobDefaultDay, Integer(m[2]))
    }

    if RegExMatch(work, "i)^(jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec|january|february|march|april|june|july|august|september|october|november|december)/(\d{1,2})/(\d{2,4})$", &m) {
        month := MonthNumber(m[1])
        if month
            return FormatDateString(month, Integer(m[2]), NormalizeYear(m[3]))
    }

    if RegExMatch(work, "i)^(jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec|january|february|march|april|june|july|august|september|october|november|december)\s+(\d{1,2})/(\d{2,4})$", &m) {
        month := MonthNumber(m[1])
        if month
            return FormatDateString(month, Integer(m[2]), NormalizeYear(m[3]))
    }

    if RegExMatch(work, "^(\d{1,2})-(\d{1,2})-(\d{2,4})$", &m) {
        p1 := Integer(m[1])
        p2 := Integer(m[2])
        yr := NormalizeYear(m[3])

        if (p1 > 12 && p2 <= 12)
            return FormatDateString(p2, p1, yr)

        return FormatDateString(p1, p2, yr)
    }

    if RegExMatch(work, "^(\d{1,2})/(\d{1,2})/(\d{2,4})$", &m) {
        p1 := Integer(m[1])
        p2 := Integer(m[2])
        yr := NormalizeYear(m[3])

        if (p1 > 12 && p2 <= 12)
            return FormatDateString(p2, p1, yr)

        return FormatDateString(p1, p2, yr)
    }

    if RegExMatch(work, "^(\d{4})-(\d{1,2})-(\d{1,2})$", &m)
        return FormatDateString(Integer(m[2]), Integer(m[3]), Integer(m[1]))

    if RegExMatch(work, "^(\d{1,2})/(\d{4})$", &m)
        return FormatDateString(Integer(m[1]), dobDefaultDay, Integer(m[2]))

    return ""
}

NormalizeMonthWords(text) {
    static monthMap := Map(
        "enero", "january",
        "ene", "jan",
        "febrero", "february",
        "feb", "feb",
        "marzo", "march",
        "mar", "mar",
        "abril", "april",
        "abr", "apr",
        "mayo", "may",
        "junio", "june",
        "jun", "jun",
        "julio", "july",
        "jul", "jul",
        "agosto", "august",
        "ago", "aug",
        "septiembre", "september",
        "setiembre", "september",
        "sep", "sep",
        "set", "sep",
        "octubre", "october",
        "oct", "oct",
        "noviembre", "november",
        "nov", "nov",
        "diciembre", "december",
        "dic", "dec"
    )

    for spanish, english in monthMap
        text := RegExReplace(text, "i)\b" spanish "\b", english)

    text := RegExReplace(text, "i)\bde\b", " ")
    text := RegExReplace(text, "\s+", " ")
    return Trim(text)
}

NormalizeYear(yearText) {
    yr := Integer(yearText)
    return (StrLen(yearText) = 2) ? ((yr <= 29) ? 2000 + yr : 1900 + yr) : yr
}

MonthNumber(name) {
    static months := Map(
        "jan", 1, "january", 1,
        "feb", 2, "february", 2,
        "mar", 3, "march", 3,
        "apr", 4, "april", 4,
        "may", 5,
        "jun", 6, "june", 6,
        "jul", 7, "july", 7,
        "aug", 8, "august", 8,
        "sep", 9, "sept", 9, "september", 9,
        "oct", 10, "october", 10,
        "nov", 11, "november", 11,
        "dec", 12, "december", 12
    )
    key := StrLower(Trim(name))
    return months.Has(key) ? months[key] : 0
}

FormatDateString(month, day, year) {
    return Format("{:02}/{:02}/{:04}", month, day, year)
}

; ===================== QUO TAG SELECTOR / TEST HELPERS =====================

tagSelectorJsFile := A_ScriptDir "\js for tag selector.js"

QuoBuildTagSelectorJS() {
    global tagSelectorJsFile
    if !FileExist(tagSelectorJsFile) {
        MsgBox("Missing file: " tagSelectorJsFile "`n`nPlace 'js for tag selector.js' next to the script.")
        return ""
    }

    js := Trim(FileRead(tagSelectorJsFile, "UTF-8"))
    if (js = "")
        return ""

    if RegExMatch(js, "i)^\s*copy\s*\(")
        return js

    return "copy(String(" . js . "))"
}

RunQuoTagSelector() {
    js := QuoBuildTagSelectorJS()
    if (js = "")
        return ""
    return RunDevToolsJSGetResult(js)
}

QuoDeleteAddTagConfirm(tagText) {
    global BATCH_BEFORE_TAG_PASTE, BATCH_AFTER_TAG_PASTE, BATCH_AFTER_ENTER

    if StopRequested()
        return false
    if !FocusWorkBrowser()
        return false

    if !SafeSleep(120)
        return false
    if StopRequested()
        return false
    Send "{Backspace}"
    if !SafeSleep(120)
        return false

    if !SafeSleep(BATCH_BEFORE_TAG_PASTE)
        return false
    if !PasteValueRaw(tagText)
        return false
    if !SafeSleep(BATCH_AFTER_TAG_PASTE)
        return false

    if StopRequested()
        return false
    Send "{Enter}"
    if !SafeSleep(BATCH_AFTER_ENTER)
        return false
    if StopRequested()
        return false
    Send "{Enter}"
    if !SafeSleep(BATCH_AFTER_ENTER)
        return false

    FocusSlateComposer()
    if !SafeSleep(200)
        return false

    return true
}

QuoPrepareFieldThenAddTag(tagText) {
    if StopRequested()
        return false
    if !FocusWorkBrowser()
        return false

    if StopRequested()
        return false
    Send "{Tab}"
    if !SafeSleep(220)
        return false
    if StopRequested()
        return false
    Send "{Tab}"
    if !SafeSleep(220)
        return false
    if StopRequested()
        return false
    Send "{Enter}"
    if !SafeSleep(300)
        return false

    return QuoDeleteAddTagConfirm(tagText)
}

HandleQuoTagSelectorResult(result, tagText) {
    if !FocusWorkBrowser()
        return "FAILED - Browser lost focus"

    if (result = "PLUS_FALLBACK") {
        if !QuoDeleteAddTagConfirm(tagText)
            return "FAILED - PLUS_FALLBACK sequence failed"
        return "OK"
    }

    if (result = "HITTEST_TARGET" || result = "STRUCTURE_TARGET" || result = "ANCESTOR_TARGET" || result = "LOCAL_TARGET") {
        if !QuoPrepareFieldThenAddTag(tagText)
            return "FAILED - " result " sequence failed"
        return "OK"
    }

    if (result = "NO_TARGET")
        return "PAUSED - No tag target"

    if (result = "")
        return "FAILED - Blank selector result"

    return "FAILED - Unhandled selector result: " result
}

ApplyQuoTag(tagText) {
    tagText := Trim(tagText)
    if (tagText = "")
        return "FAILED - Blank tag value"
    if StopRequested()
        return "STOPPED"

    result := RunQuoTagSelector()
    if StopRequested()
        return "STOPPED"
    if (result = "")
        return "FAILED - Blank selector result"

    return HandleQuoTagSelectorResult(result, tagText)
}
; ===================== DATE / BUSINESS-DAY HELPERS =====================

IsHoliday(mmddyyyy, holidaysArr) {
    for h in holidaysArr
        if (h = mmddyyyy)
            return true
    return false
}

NextBusinessDateYYYYMMDD(startYYYYMMDD, holidaysArr) {
    ts := startYYYYMMDD
    if (StrLen(ts) = 8)
        ts := ts . "000000"
    else if (StrLen(ts) != 14)
        ts := FormatTime(A_Now, "yyyyMMddHHmmss")

    Loop {
        ts := DateAdd(ts, 1, "D")
        ymd := FormatTime(ts, "yyyyMMdd")
        wday := FormatTime(ts, "WDay")
        mmddyyyy := FormatTime(ts, "MM/dd/yyyy")
        if (wday != 1 && wday != 7 && !IsHoliday(mmddyyyy, holidaysArr))
            return ymd
    }
}

BusinessDateForDay(dayIndex, holidaysArr) {
    if (dayIndex <= 0)
        dayIndex := 1
    baseYMD := FormatTime(A_Now, "yyyyMMdd")
    ymd := baseYMD
    Loop dayIndex
        ymd := NextBusinessDateYYYYMMDD(ymd, holidaysArr)
    return FormatTime(ymd . "000000", "MM/dd/yyyy")
}

AddBusinessDays(startYYYYMMDD, k, holidaysArr) {
    ymd := startYYYYMMDD
    Loop k
        ymd := NextBusinessDateYYYYMMDD(ymd, holidaysArr)
    return ymd
}

BuildBusinessDates(n, holidaysArr) {
    arr := []
    last := FormatTime(A_Now, "yyyyMMdd")
    Loop n {
        last := NextBusinessDateYYYYMMDD(last, holidaysArr)
        arr.Push(FormatTime(last . "000000", "MM/dd/yyyy"))
    }
    return arr
}

Pad2(n) => Format("{:02}", n)

TimeWithOffset(h, m, s, offsetMin) {
    dt := FormatTime(A_Now, "yyyyMMdd") . Pad2(h) . Pad2(m) . Pad2(s)
    dt := DateAdd(dt, offsetMin, "M")
    return FormatTime(dt, "hh:mm:ss tt")
}

GetInitialQuoteDateTime(offset) {
    global holidays
    todayDate := FormatTime(A_Now, "MM/dd/yyyy")
    todayYMD  := FormatTime(A_Now, "yyyyMMdd")
    nowHHMMSS := Integer(FormatTime(A_Now, "HHmmss"))

    noonTime := TimeWithOffset(12, 0, 0, offset)
    dt12 := todayYMD . "120000"
    dt12 := DateAdd(dt12, offset, "M")
    noonHHMMSS := Integer(FormatTime(dt12, "HHmmss"))

    if (nowHHMMSS <= noonHHMMSS)
        return Map("date", todayDate, "time", noonTime)

    if (nowHHMMSS <= 175500)
        return Map("date", todayDate, "time", "05:55:00 PM")

    nextBizYMD  := NextBusinessDateYYYYMMDD(todayYMD, holidays)
    nextBizDate := FormatTime(nextBizYMD . "000000", "MM/dd/yyyy")
    return Map("date", nextBizDate, "time", noonTime)
}

; ===================== PRIMARY SCHEDULING / MESSAGE WORKFLOWS =====================

ScheduleMessage(msgText, dateMDY, time12) {
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

ScheduleMessageTyped(msgText, dateMDY, time12) {
    global SLOW_ACTIVATE_DELAY, SLOW_AFTER_MSG, SLOW_AFTER_SCHED, SLOW_AFTER_DT_PASTE, SLOW_AFTER_ENTER
    if StopRequested()
        return false
    if !FocusWorkBrowser() {
        MsgBox("Browser not found/active. Open the chat window first.")
        return false
    }
    if !SafeSleep(SLOW_ACTIVATE_DELAY)
        return false

    ; Clipboard paste to avoid Slate emoji conversion from SendText
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

ScheduleFollowupMessages(msgs, useTyped := false) {
    for m in msgs {
        if StopRequested()
            return false
        ok := useTyped
            ? ScheduleMessageTyped(m["text"], m["date"], m["time"])
            : ScheduleMessage(m["text"], m["date"], m["time"])
        if !ok
            return false
    }
    return true
}

BuildMessage(leadName, carCount, vehicles := "", useBatchSingleCarPricing := false) {
    global agentName, agentEmail
    greeting := (A_Hour < 12) ? "Buenos días" : "Buenas tardes"

    leadName := ProperCase(leadName)
    firstName := ExtractFirstName(leadName)

    vehLine := (carCount >= 2)
        ? "Hicimos la cotización para el seguro de sus carros.`n"
        : "Hicimos la cotización para el seguro de su carro.`n"

    ; Append vehicle names if provided
    if (IsObject(vehicles) && vehicles.Length > 0) {
        for _, v in vehicles
            vehLine .= v . "`n"
    }

    coverageLine := ""
    if (carCount = 5)
        coverageLine := "`n`nBodily Injury $100k per person $300k per ocurrence`nProperty Damage $100k per ocurrence"

    price := ResolveQuotePrice(carCount, vehicles, useBatchSingleCarPricing)
    coverageSuffix := " al mes."

    return Format("
    (
    {1} {2},

    {3}{4}

    Actualmente tenemos opciones con ALLSTATE desde {5}{6}

    Muchos clientes en su misma situación han logrado ahorrar cambiándose con nosotros.

    Si quiere, en una llamada rápida de 2-3 minutos podemos revisar si realmente le conviene o no.

    ¿Le parece bien si lo revisamos juntos?

    📞 (561) 220-7073

    {7}
    Agente de Seguros – Allstate
    Direct Line: (561) 220-7073
    Office Line: (754) 236-8009
    {8}
    Reply STOP to unsubscribe
    )"
    , greeting, firstName
    , vehLine, coverageLine
    , price, coverageSuffix
    , agentName, agentEmail)
}

ResolveQuotePrice(carCount, vehicles := "", useBatchSingleCarPricing := false) {
    global priceOldCar, priceOneCar, priceOneCar2020Plus
    global priceTwoCars, priceThreeCars, priceFourCars, priceFiveCars
    global singleCarModernYearCutoff

    if (useBatchSingleCarPricing && carCount = 1) {
        vehicleYear := ExtractVehicleYearFromList(vehicles)
        if (vehicleYear >= singleCarModernYearCutoff)
            return FormatMonthlyPrice(priceOneCar2020Plus)
        return FormatMonthlyPrice(priceOneCar)
    }

    prices := Map(0, priceOldCar, 1, priceOneCar, 2, priceTwoCars, 3, priceThreeCars, 4, priceFourCars, 5, priceFiveCars)
    return prices.Has(carCount) ? FormatMonthlyPrice(prices[carCount]) : "$127"
}

ExtractVehicleYearFromList(vehicles) {
    if !(IsObject(vehicles) && vehicles.Length >= 1)
        return 0
    return ExtractVehicleYear(vehicles[1])
}

ExtractVehicleYear(vehicleText) {
    if RegExMatch(vehicleText, "i)\b((19|20)\d{2})\b", &m)
        return Integer(m[1])
    return 0
}

FormatMonthlyPrice(amount) {
    return "$" . Integer(amount)
}

BuildFollowupQueue(leadName, offset) {
    global agentName, configDays, holidays
    leadName := ExtractFirstName(ProperCase(leadName))

    dA := configDays[1]
    dB := configDays[2]
    dC := configDays[3]
    dD := configDays[4]

    DA_date := BusinessDateForDay(dA, holidays)
    DB_date := BusinessDateForDay(dB, holidays)
    DC_date := BusinessDateForDay(dC, holidays)
    DD_date := BusinessDateForDay(dD, holidays)

    t1_1 := TimeWithOffset(9, 30, 30, offset)
    t1_2 := TimeWithOffset(9, 31, 10, offset)
    t1_3 := TimeWithOffset(9, 31, 30, offset)
    t1_4 := TimeWithOffset(10, 45, 0, offset)
    t2_1 := TimeWithOffset(16, 0, 10, offset)
    t2_2 := TimeWithOffset(16, 1, 10, offset)
    t4_1 := TimeWithOffset(16, 30, 0, offset)
    t4_2 := TimeWithOffset(16, 31, 0, offset)
    t5_1 := TimeWithOffset(12, 0, 0, offset)
    t5_2 := TimeWithOffset(12, 1, 0, offset)

    msgs := [
        Map("day", dA, "seq", 1, "text", "Buenos días, " . leadName . ".", "date", DA_date, "time", t1_1),
        Map("day", dA, "seq", 2, "text", "Soy " . agentName . " de Allstate. Ya le preparé la cotización de su auto.", "date", DA_date, "time", t1_2),
        Map("day", dA, "seq", 3, "text", "En muchos casos logramos bajar el pago mensual sin quitar coberturas. Si gusta, se la resumo en 2 minutos por aquí.", "date", DA_date, "time", t1_3),
        Map("day", dA, "seq", 4, "text", "Si me responde “Sí”, se la envío ahora mismo.", "date", DA_date, "time", t1_4),

        Map("day", dB, "seq", 1, "text", "Hola, " . leadName . ".", "date", DB_date, "time", t2_1),
        Map("day", dB, "seq", 2, "text", "Hoy intenté comunicarme con usted porque todavía puedo revisar si califica a descuentos disponibles. Si me responde “Revisar”, yo me encargo de validar todo por usted.", "date", DB_date, "time", t2_2),

        Map("day", dC, "seq", 1, "text", "Buenas tardes, " . leadName . ".", "date", DC_date, "time", t4_1),
        Map("day", dC, "seq", 2, "text", "Esta semana hemos ayudado a varios conductores a comparar su póliza actual con Allstate y en muchos casos encontraron una mejor opción. Si me responde “Comparar”, reviso su caso y le digo honestamente si le conviene o no. Reply STOP to unsubscribe", "date", DC_date, "time", t4_2),

        Map("day", dD, "seq", 1, "text", leadName . ", sigo teniendo su cotización disponible, pero normalmente cierro los pendientes cuando no recibo respuesta.", "date", DD_date, "time", t5_1),
        Map("day", dD, "seq", 2, "text", "Si todavía quiere revisarla, respóndame “Continuar” y le envío el resumen por aquí.", "date", DD_date, "time", t5_2)
    ]
    return msgs
}

; ===================== BATCH LEAD WORKFLOWS =====================

ParseBatchGridRow(raw) {
    fields := NewProspectFields()
    cols := StrSplit(raw, "`t")
    if (cols.Length < 11)
        return fields

    fields["RAW_NAME"]  := Trim(cols[1])
    fields["ADDRESS_1"] := Trim(cols[5])
    fields["CITY"]      := NormalizeCity(cols[6])
    fields["STATE"]     := NormalizeState(cols[7])
    fields["ZIP"]       := NormalizeZip(cols[8])
    fields["PHONE"]     := NormalizePhone(cols[9])

    idx := 10
    if (cols.Length >= idx && IsEmailToken(cols[idx])) {
        fields["EMAIL"] := Trim(cols[idx])
        idx += 1
    }

    if (cols.Length >= idx)
        fields["DOB"] := NormalizeDOB(cols[idx])
    if (cols.Length >= idx + 1)
        fields["GENDER"] := NormalizeGender(cols[idx + 1])

    nameText := ExtractBatchName(cols[1])
    if (nameText != "")
        ApplyLeadName(fields, nameText)

    NormalizeAddressMap(fields)
    fields["DOB"]    := NormalizeDOB(fields["DOB"])
    fields["GENDER"] := NormalizeGender(fields["GENDER"])
    fields["STATE"]  := NormalizeState(fields["STATE"])
    fields["ZIP"]    := NormalizeZip(fields["ZIP"])
    fields["PHONE"]  := NormalizePhone(fields["PHONE"])
    return fields
}

ParseBatchLeadRows(raw) {
    rows := []
    clean := StrReplace(raw, "`r", "")

    split := RegExReplace(clean, "i)(?=(?:DUPLICATED\s+(?:OPPORTUNITY\s+)?)?PERSONAL\s+LEAD)", "`n")

    for _, chunk in StrSplit(split, "`n") {
        chunk := Trim(chunk)
        if (chunk = "")
            continue
        if RegExMatch(chunk, "i)PERSONAL\s+LEAD")
            rows.Push(chunk)
    }
    return rows
}

ExtractBatchName(raw) {
    prefixPattern := "i)^\s*(?:DUPLICATED\s+)?(?:OPPORTUNITY\s+)?PERSONAL\s+LEAD\s*-\s*"

    if IsLabeledLeadFormat(raw) {
        data := ParseLabeledLeadRaw(raw)

        if data.Has("Name") {
            name := CleanBatchNameCandidate(RegExReplace(data["Name"], prefixPattern, ""))
            if (name != "")
                return name
        }

        if data.Has("Contact") {
            name := CleanBatchNameCandidate(RegExReplace(data["Contact"], prefixPattern, ""))
            if (name != "")
                return name
        }

        first := data.Has("First Name") ? Trim(data["First Name"]) : ""
        last  := data.Has("Last Name")  ? Trim(data["Last Name"])  : ""
        name := CleanBatchNameCandidate(first . " " . last)
        if (name != "")
            return name
    }

    raw := StripGridActionText(raw)

    if RegExMatch(raw, prefixPattern . "(.*)", &m)
        return CleanBatchNameCandidate(m[1])

    return CleanBatchNameCandidate(raw)
}

CleanBatchNameCandidate(name) {
    name := Trim(StrReplace(StrReplace(name, "`r", ""), "`n", " "))
    if (name = "")
        return ""

    name := RegExReplace(name, "[\t\r\n].*$", "")
    name := RegExReplace(name, "i)\s{2,}(?=(?:new\s+(?:webform\s+folder|skyline\s+leads)\s*-\s*personal|(?:[A-Za-z]+\s+){0,3}(?:folder|leads?|source|status)\b(?:\s*-\s*[A-Za-z]+)?|\d+\s*-\s*(?:new|open|working|quoted|pending|closed|sold|contacted)|move\s+to\s+recycle\s+bin|recycle\s+bin|\d{1,2}/\d{1,2}/\d{2,4}\b)).*$", "")
    name := RegExReplace(name, "i)\s+(?:new\s+webform\s+folder\s*-\s*personal|new\s+skyline\s+leads\s*-\s*personal|move\s+to\s+recycle\s+bin|recycle\s+bin)\b.*$", "")
    name := RegExReplace(name, "i)\s+\d+\s*-\s*(?:new|open|working|quoted|pending|closed|sold|contacted)\b.*$", "")
    name := RegExReplace(name, "i)\s+(?:[A-Za-z]+\s+){0,3}(?:folder|leads?|source|status)\b(?:\s*-\s*[A-Za-z]+)?(?:\s+.*)?$", "")
    name := RegExReplace(name, "i)\s+\d{1,2}/\d{1,2}/\d{2,4}\s+\d{1,2}:\d{2}(?::\d{2})?\s*(?:AM|PM)?\b.*$", "")
    name := RegExReplace(name, "i)^\s*(?:DUPLICATED\s+)?(?:OPPORTUNITY\s+)?PERSONAL\s+LEAD\s*-\s*", "")
    name := RegExReplace(name, "\s{2,}", " ")
    return Trim(name, " -:`t`r`n")
}

ExtractBatchPhone(raw) {
    if RegExMatch(raw, "\((\d{3})\)\s*(\d{3})-(\d{4})", &m)
        return m[1] . m[2] . m[3]
    return ""
}

ExtractVehicleList(rawLead) {
    vehicles := []

    rawLead := StripGridActionText(rawLead)

    vehicleText := ""
    if RegExMatch(rawLead, "i)(?:Male|Female)\s*(.*)", &m)
        vehicleText := Trim(m[1])

    if (vehicleText = "")
        return vehicles

    vehicleText := StripGridActionText(vehicleText)
    split := RegExReplace(vehicleText, "i)((19|20)\d{2})(\s*/?[\s]*[A-Za-z])", "`n$1$3")

    for _, part in StrSplit(split, "`n") {
        part := SanitizeVehicleLine(part)
        if (part = "")
            continue
        if RegExMatch(part, "i)^(19|20)\d{2}\s*/?[\s]*[A-Za-z]")
            vehicles.Push(ProperCasePhrase(part))
    }

    return vehicles
}

BuildBatchLeadRecord(rawLead) {
    global tagSymbol

    rawLead := StripGridActionText(rawLead)
    nameSource := rawLead

    if InStr(rawLead, "`t") {
        cols := StrSplit(rawLead, "`t")
        if (cols.Length >= 1) {
            firstCol := Trim(cols[1])
            if (firstCol != "")
                nameSource := firstCol
        }
    }

    batchName  := ExtractBatchName(nameSource)
    if (batchName = "" && nameSource != rawLead)
        batchName := ExtractBatchName(rawLead)
    batchPhone := ExtractBatchPhone(rawLead)
    vehicles   := ExtractVehicleList(rawLead)

    fullName := ProperCase(batchName)
    firstName := ExtractFirstName(fullName)

    parts := StrSplit(Trim(fullName), " ")
    lastName  := (parts.Length >= 2) ? parts[parts.Length] : ""

    tagValue   := Trim(tagSymbol)
    if (tagValue = "")
        tagValue := "+"

    holderName := tagValue . " " . fullName

    return Map(
        "RAW", rawLead,
        "FIELDS", Map("FIRST_NAME", firstName, "LAST_NAME", lastName, "PHONE", batchPhone),
        "FULL_NAME", fullName,
        "PHONE", batchPhone,
        "VEHICLES", vehicles,
        "VEHICLE_COUNT", vehicles.Length,
        "HOLDER_NAME", holderName,
        "TAG_VALUE", tagValue
    )
}

BuildBatchLeadHolder(raw) {
    global batchMinVehicles, batchMaxVehicles

    holder := []
    rows := ParseBatchLeadRows(raw)

    for _, row in rows {
        lead := BuildBatchLeadRecord(row)
        vc := lead["VEHICLE_COUNT"]

        if (vc < batchMinVehicles || vc > batchMaxVehicles)
            continue

        holder.Push(lead)
    }
    return holder
}

ScheduleBuilderForLead(lead, offset) {
    global SLOW_AFTER_SCHED, SLOW_AFTER_DT_PASTE, SLOW_AFTER_ENTER

    if StopRequested()
        return false
    dt := GetInitialQuoteDateTime(offset)
    msg := BuildMessage(lead["FULL_NAME"], lead["VEHICLE_COUNT"], lead["VEHICLES"], true)

    if !FocusWorkBrowser() {
        MsgBox("Browser not found/active.")
        return false
    }
    if !SafeSleep(200)
        return false

    if !SetClip(msg)
        return false
    if !SafeSleep(80)
        return false
    if StopRequested()
        return false
    Send "^v"
    if !SafeSleep(400)
        return false

    if StopRequested()
        return false
    Send "^!{Enter}"
    if !SafeSleep(SLOW_AFTER_SCHED)
        return false

    if StopRequested()
        return false
    SendText dt["date"] . " " . dt["time"]
    if !SafeSleep(SLOW_AFTER_DT_PASTE)
        return false

    if StopRequested()
        return false
    Send "{Enter}"
    if !SafeSleep(SLOW_AFTER_ENTER)
        return false
    return true
}

ScheduleRegularFollowupsForLead(lead, offset) {
    msgs := BuildFollowupQueue(lead["FULL_NAME"], offset)
    return ScheduleFollowupMessages(msgs, true)
}

RunBatchLeadFlow(lead) {
    global BATCH_AFTER_ALTN, BATCH_AFTER_PHONE, BATCH_AFTER_TAB
    global BATCH_AFTER_SCHEDULE, BATCH_AFTER_ENTER, BATCH_AFTER_NAME_PICK, BATCH_AFTER_TAG_PICK
    global BATCH_BEFORE_TAG_PASTE, BATCH_AFTER_TAG_PASTE, batchTabsChatToName

    if StopRequested()
        return "STOPPED"
    if !FocusWorkBrowser()
        return "FAILED - Browser lost focus"

    if !SafeSleep(150)
        return "STOPPED"

    if (lead["PHONE"] = "" || lead["FULL_NAME"] = "")
        return "SKIPPED - Missing phone or name"

    offset := NextRotationOffset()

    if StopRequested()
        return "STOPPED"
    Send "!n"
    if !SafeSleep(BATCH_AFTER_ALTN)
        return "STOPPED"

    if !EnsureParticipantInputReady()
        return StopRequested() ? "STOPPED" : "FAILED - Could not focus To field"
    if !SafeSleep(150)
        return "STOPPED"

    if !PasteValue(lead["PHONE"])
        return StopRequested() ? "STOPPED" : "FAILED - Could not paste phone"
    if !SafeSleep(BATCH_AFTER_PHONE)
        return "STOPPED"

    if StopRequested()
        return "STOPPED"
    Send "{Tab}"
    if !SafeSleep(BATCH_AFTER_TAB)
        return "STOPPED"
    if StopRequested()
        return "STOPPED"
    Send "{Tab}"
    if !SafeSleep(1000)
        return "STOPPED"
    FocusSlateComposer()

    if !ScheduleBuilderForLead(lead, offset)
        return StopRequested() ? "STOPPED" : "FAILED - Builder scheduling failed"
    if !SafeSleep(BATCH_AFTER_SCHEDULE)
        return "STOPPED"

    if !ScheduleRegularFollowupsForLead(lead, offset)
        return StopRequested() ? "STOPPED" : "FAILED - Follow-up scheduling failed"
    if !SafeSleep(BATCH_AFTER_SCHEDULE)
        return "STOPPED"

    ; move to the lead name field and select the created lead
    if !SafeSleep(400)
        return "STOPPED"
    if !SendTabs(batchTabsChatToName)
        return "STOPPED"
    if !SafeSleep(200)
        return "STOPPED"
    if StopRequested()
        return "STOPPED"
    Send "{Enter}"
    if !SafeSleep(BATCH_AFTER_ENTER)
        return "STOPPED"

    if StopRequested()
        return "STOPPED"
    Send "^a"
    if !SafeSleep(80)
        return "STOPPED"
    if !PasteValue(lead["HOLDER_NAME"])
        return StopRequested() ? "STOPPED" : "FAILED - Could not paste holder into name field"
    if !SafeSleep(BATCH_AFTER_NAME_PICK)
        return "STOPPED"

    if StopRequested()
        return "STOPPED"
    Send "{Enter}"
    if !SafeSleep(BATCH_AFTER_ENTER)
        return "STOPPED"

    if !SafeSleep(300)
        return "STOPPED"
		
    tagStatus := ApplyQuoTag(lead["TAG_VALUE"])
    if (tagStatus != "OK")
        return tagStatus

    return "OK"
}

ScheduleBuilderForLeadFast(lead, offset) {
    if StopRequested()
        return false
    dt := GetInitialQuoteDateTime(offset)
    msg := BuildMessage(lead["FULL_NAME"], lead["VEHICLE_COUNT"], lead["VEHICLES"], true)

    if !FocusWorkBrowser() {
        MsgBox("Browser not found/active.")
        return false
    }
    if !SafeSleep(200)
        return false

    if !SetClip(msg)
        return false
    if !SafeSleep(80)
        return false
    if StopRequested()
        return false
    Send "^v"
    if !SafeSleep(400)
        return false

    if StopRequested()
        return false
    Send "^!{Enter}"
    if !SafeSleep(300)
        return false

    if !SetClip(dt["date"] . " " . dt["time"])
        return false
    if !SafeSleep(80)
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

ScheduleRegularFollowupsForLeadFast(lead, offset) {
    msgs := BuildFollowupQueue(lead["FULL_NAME"], offset)
    return ScheduleFollowupMessages(msgs, false)
}

RunBatchLeadFlowFast(lead) {
    global BATCH_AFTER_ALTN, BATCH_AFTER_PHONE, BATCH_AFTER_TAB
    global BATCH_AFTER_SCHEDULE, BATCH_AFTER_ENTER, BATCH_AFTER_NAME_PICK, BATCH_AFTER_TAG_PICK
    global BATCH_BEFORE_TAG_PASTE, BATCH_AFTER_TAG_PASTE, batchTabsChatToName

    if StopRequested()
        return "STOPPED"
    if !FocusWorkBrowser()
        return "FAILED - Browser lost focus"

    if !SafeSleep(150)
        return "STOPPED"

    if (lead["PHONE"] = "" || lead["FULL_NAME"] = "")
        return "SKIPPED - Missing phone or name"

    offset := NextRotationOffset()

    if StopRequested()
        return "STOPPED"
    Send "!n"
    if !SafeSleep(BATCH_AFTER_ALTN)
        return "STOPPED"

    if !EnsureParticipantInputReady()
        return StopRequested() ? "STOPPED" : "FAILED - Could not focus To field"

    if !PasteValue(lead["PHONE"])
        return StopRequested() ? "STOPPED" : "FAILED - Could not paste phone"
    if !SafeSleep(BATCH_AFTER_PHONE)
        return "STOPPED"

    if StopRequested()
        return "STOPPED"
    Send "{Tab}"
    if !SafeSleep(BATCH_AFTER_TAB)
        return "STOPPED"
    if StopRequested()
        return "STOPPED"
    Send "{Tab}"
    if !SafeSleep(1000)
        return "STOPPED"
    FocusSlateComposer()

    if !ScheduleBuilderForLeadFast(lead, offset)
        return StopRequested() ? "STOPPED" : "FAILED - Builder scheduling failed"
    if !SafeSleep(BATCH_AFTER_SCHEDULE)
        return "STOPPED"

    if !ScheduleRegularFollowupsForLeadFast(lead, offset)
        return StopRequested() ? "STOPPED" : "FAILED - Follow-up scheduling failed"
    if !SafeSleep(BATCH_AFTER_SCHEDULE)
        return "STOPPED"

    if !SafeSleep(400)
        return "STOPPED"
    if !SendTabs(batchTabsChatToName)
        return "STOPPED"
    if !SafeSleep(200)
        return "STOPPED"
    if StopRequested()
        return "STOPPED"
    Send "{Enter}"
    if !SafeSleep(BATCH_AFTER_ENTER)
        return "STOPPED"

    if StopRequested()
        return "STOPPED"
    Send "^a"
    if !SafeSleep(80)
        return "STOPPED"
    if !PasteValue(lead["HOLDER_NAME"])
        return StopRequested() ? "STOPPED" : "FAILED - Could not paste holder into name field"
    if !SafeSleep(BATCH_AFTER_NAME_PICK)
        return "STOPPED"

    if StopRequested()
        return "STOPPED"
    Send "{Enter}"
    if !SafeSleep(BATCH_AFTER_ENTER)
        return "STOPPED"

    if !SafeSleep(300)
        return "STOPPED"

    tagStatus := ApplyQuoTag(lead["TAG_VALUE"])
    if (tagStatus != "OK")
        return tagStatus

    return "OK"
}

SanitizeVehicleLine(text) {
    text := StripGridActionText(text)
    text := RegExReplace(text, "\s*/\s*", " ")   ; <-- ADD THIS LINE
    text := RegExReplace(text, "\s+", " ")        ; <-- collapse extra spaces
    text := RegExReplace(text, "[\t ]+$", "")
    return Trim(text)
}

StripGridActionText(text) {
    text := Trim(text)
    text := RegExReplace(text, "i)\bMove\s+To\s+Recycle\s+Bin\b", "")
    text := RegExReplace(text, "i)\bRecycle\s+Bin\b", "")
    text := RegExReplace(text, "\s{2,}", " ")
    return Trim(text, " `t-")
}

RunQuickLeadCreateAndTag(lead) {
    global BATCH_AFTER_ALTN, BATCH_AFTER_PHONE, BATCH_AFTER_TAB
    global BATCH_AFTER_ENTER, BATCH_AFTER_NAME_PICK, batchTabsChatToName

    if StopRequested()
        return "STOPPED"
    if !FocusWorkBrowser()
        return "FAILED - Browser lost focus"

    if !SafeSleep(150)
        return "STOPPED"

    if (lead["PHONE"] = "" || lead["FULL_NAME"] = "")
        return "FAILED - Missing phone or name"

    if StopRequested()
        return "STOPPED"
    Send "!n"
    if !SafeSleep(BATCH_AFTER_ALTN)
        return "STOPPED"

    if !EnsureParticipantInputReady()
        return StopRequested() ? "STOPPED" : "FAILED - Could not focus To field"

    if !PasteValue(lead["PHONE"])
        return StopRequested() ? "STOPPED" : "FAILED - Could not paste phone"
    if !SafeSleep(BATCH_AFTER_PHONE)
        return "STOPPED"

    if StopRequested()
        return "STOPPED"
    Send "{Tab}"
    if !SafeSleep(BATCH_AFTER_TAB)
        return "STOPPED"
    if StopRequested()
        return "STOPPED"
    Send "{Tab}"
    if !SafeSleep(1000)
        return "STOPPED"
    FocusSlateComposer()
    if !SafeSleep(250)
        return "STOPPED"

    if !SendTabs(batchTabsChatToName)
        return "STOPPED"
    if !SafeSleep(200)
        return "STOPPED"
    if StopRequested()
        return "STOPPED"
    Send "{Enter}"
    if !SafeSleep(BATCH_AFTER_ENTER)
        return "STOPPED"

    if StopRequested()
        return "STOPPED"
    Send "^a"
    if !SafeSleep(80)
        return "STOPPED"
    if !PasteValue(lead["HOLDER_NAME"])
        return StopRequested() ? "STOPPED" : "FAILED - Could not paste holder into name field"
    if !SafeSleep(BATCH_AFTER_NAME_PICK)
        return "STOPPED"

    if StopRequested()
        return "STOPPED"
    Send "{Enter}"
    if !SafeSleep(BATCH_AFTER_ENTER)
        return "STOPPED"

    tagStatus := ApplyQuoTag(lead["TAG_VALUE"])
    if (tagStatus != "OK")
        return tagStatus

    return "OK"
}

; ===================== FORM-FILL WORKFLOWS =====================

FillNewProspectForm(fields) {
    FastType(fields["FIRST_NAME"])
    Sleep 30
    Send "{Tab}"
    Sleep 30

    Send "{Tab}"
    Sleep 30

    FastType(fields["LAST_NAME"])
    Sleep 30
    Send "{Tab}"
    Sleep 30

    Send "{Tab}"
    Sleep 30

    PasteField(fields["DOB"])
    Sleep 120
    Send "{Tab}"
    Sleep 80

    SelectDropdownValue(fields["GENDER"])

    FastType(fields["ADDRESS_1"])
    Sleep 30
    Send "{Tab}"
    Sleep 30

    FastType(fields["APT_SUITE"])
    Sleep 30
    Send "{Tab}"
    Sleep 30

    FastType(fields["BUILDING"])
    Sleep 30
    Send "{Tab}"
    Sleep 30

    FastType(fields["RR_NUMBER"])
    Sleep 30
    Send "{Tab}"
    Sleep 30

    FastType(fields["LOT_NUMBER"])
    Sleep 30
    Send "{Tab}"
    Sleep 30

    FastType(fields["CITY"])
    Sleep 30
    Send "{Tab}"
    Sleep 50

    SelectDropdownValue(fields["STATE"])

    FastType(fields["ZIP"])
    Sleep 30
    Send "{Tab}"
    Sleep 30

    FastType(fields["PHONE"])
}



FillNationalGeneralForm(fields) {
FastType(fields["FIRST_NAME"])
    Sleep 30
    Send "{Tab}"
    Sleep 30

    Send "{Tab}"
    Sleep 30

    FastType(fields["LAST_NAME"])
    Sleep 30
    Send "{Tab}"
    Sleep 30
	
    Loop 5{
    Send "{Tab}"
    Sleep 30
    }

    PasteField(fields["DOB"])
    Sleep 120
    Send "{Tab}"
    Sleep 80

    Loop 3 {
    Send "{Tab}" 
    Sleep 30
    }

    FastType(fields["ADDRESS_1"])
    Sleep 30
    Send "{Tab}"
    Sleep 30

    FastType(fields["APT_SUITE"])
    Sleep 30
    Send "{Tab}"
    Sleep 30

    FastType(fields["CITY"])
    Sleep 30
    Send "{Tab}"
    Sleep 50

    SelectDropdownValue(fields["STATE"])

    FastType(fields["ZIP"])
    Sleep 30
    Send "{Tab}"
    Sleep 30

}


; ===================== OPTIONAL DEBUG / TEST UTILITIES =====================

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
