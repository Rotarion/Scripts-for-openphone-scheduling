#Requires AutoHotkey v2.0
#SingleInstance Force
SendMode "Input"

; ===================== CONFIG =====================
iniFile   := A_ScriptDir "\time_rotation.ini"
holidays  := ["01/01/2026", "07/04/2025", "11/27/2025", "12/25/2025"]  ; MM/dd/yyyy

; Agent persisted in the same INI
agentName  := IniRead(iniFile, "Agent", "Name",  "Pablo Cabrera")
agentEmail := IniRead(iniFile, "Agent", "Email", "pablocabrera@allstate.com")
; Schedule days (read from INI)
daysStr := IniRead(iniFile, "Schedule", "Days", "1,2,4,5")
configDays := ParseDays(daysStr)
if (configDays.Length != 4) {
    ; fallback to default if malformed
    configDays := [1, 2, 4, 5]
}


; Typing-mode pacing for ^!7 (stable)
SLOW_ACTIVATE_DELAY := 200
SLOW_AFTER_MSG      := 300
SLOW_AFTER_SCHED    := 900
SLOW_AFTER_DT_PASTE := 900
SLOW_AFTER_ENTER    := 300
; ==================================================

; ================== SYSTEM HOTKEYS =================
Esc::ExitApp

^!r:: {  ; Reset rotation index
    IniWrite(0, iniFile, "Times", "Offset")
    MsgBox("Rotation index reset to 0.")
}

; Change agent name + email (Ctrl+Alt+`)
^!`:: {
    global agentName, agentEmail, iniFile
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

    IniWrite(newName,  iniFile, "Agent", "Name")
    IniWrite(newEmail, iniFile, "Agent", "Email")
    agentName  := newName
    agentEmail := newEmail
    TrayTip("AHK", "Agente actualizado a:`n" newName "`n" newEmail, 1)
}

; ===================== HELPERS =====================

; Parse "1,2,4,5" -> [1,2,4,5] (ints, trimmed, >0)
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

; Return the MM/dd/yyyy for business day #N counting from today:
; N=1 => next business day, N=2 => the business day after that, etc.
BusinessDateForDay(dayIndex, holidaysArr) {
    if (dayIndex <= 0)
        dayIndex := 1
    baseYMD := FormatTime(A_Now, "yyyyMMdd")
    ymd := baseYMD
    Loop dayIndex {
        ymd := NextBusinessDateYYYYMMDD(ymd, holidaysArr)
    }
    return FormatTime(ymd . "000000", "MM/dd/yyyy")
}


; Array membership
ArrContains(arr, val) {
    for v in arr
        if (v = val)
            return true
    return false
}

; Add k business days to a YYYYMMDD date
AddBusinessDays(startYYYYMMDD, k, holidaysArr) {
    ymd := startYYYYMMDD
    Loop k
        ymd := NextBusinessDateYYYYMMDD(ymd, holidaysArr)
    return ymd
}

Pad2(n) => Format("{:02}", n)

SetClip(text) {
    A_Clipboard := ""
    Sleep 30
    A_Clipboard := text
    return ClipWait(1)
}

CleanName(str) {
    str := Trim(str)
    str := StrReplace(str, "`r")
    str := StrReplace(str, "`n")
    return str
}

; Proper-case a name: "JOHN DOE" -> "John Doe"
ProperCase(str) {
    s := StrLower(Trim(str))
    parts := StrSplit(s, A_Space)
    out := ""
    for i, p in parts {
        if (p = "")
            continue
        out .= (out != "" ? " " : "") . StrUpper(SubStr(p, 1, 1)) . SubStr(p, 2)
    }
    return out
}

IsHoliday(mmddyyyy, holidaysArr) {
    for h in holidaysArr
        if (h = mmddyyyy)
            return true
    return false
}

; Return YYYYMMDD for the next business day after startYYYYMMDD (or A_Now)
NextBusinessDateYYYYMMDD(startYYYYMMDD, holidaysArr) {
    ts := startYYYYMMDD
    if (StrLen(ts) = 8)
        ts := ts . "000000"
    else if (StrLen(ts) != 14)
        ts := FormatTime(A_Now, "yyyyMMddHHmmss")  ; fallback: now

    Loop {
        ts := DateAdd(ts, 1, "D")
        ymd := FormatTime(ts, "yyyyMMdd")
        wday := FormatTime(ts, "WDay")             ; 1=Sun ... 7=Sat
        mmddyyyy := FormatTime(ts, "MM/dd/yyyy")
        if (wday != 1 && wday != 7 && !IsHoliday(mmddyyyy, holidaysArr))
            return ymd
    }
}

; Build N business days as ["MM/dd/yyyy", ...]
BuildBusinessDates(n, holidaysArr) {
    arr := []
    last := FormatTime(A_Now, "yyyyMMdd")
    Loop n {
        last := NextBusinessDateYYYYMMDD(last, holidaysArr)
        arr.Push(FormatTime(last . "000000", "MM/dd/yyyy"))
    }
    return arr
}

; Time string builder: "hh:mm:ss tt" after adding minute rotation offset
TimeWithOffset(h, m, s, offsetMin) {
    dt := FormatTime(A_Now, "yyyyMMdd") . Pad2(h) . Pad2(m) . Pad2(s)
    dt := DateAdd(dt, offsetMin, "M")
    return FormatTime(dt, "hh:mm:ss tt")
}

; Focus helper for typed mode
FocusChrome() {
    if WinExist("ahk_exe chrome.exe") {
        WinActivate
        WinWaitActive "ahk_exe chrome.exe",, 2
        return true
    }
    return false
}

; ===================== SCHEDULERS =====================
; Clipboard-based (fast)
ScheduleMessage(msgText, dateMDY, time12) {
    A_Clipboard := ""
    Sleep 50
    A_Clipboard := msgText
    if !ClipWait(1) {
        MsgBox("Clipboard failed to set message text.")
        return false
    }
    Sleep 100
    Send "^v"
    Send "^!{Enter}"
    Sleep 300

    A_Clipboard := ""
    Sleep 50
    A_Clipboard := dateMDY " " time12
    if !ClipWait(1) {
        MsgBox("Clipboard failed to set date/time.")
        return false
    }
    Sleep 100
    Send "^v"
    Sleep 300
    Send "{Enter}"
    Sleep 200
    return true
}

; Typed mode (slow & stable)
ScheduleMessageTyped(msgText, dateMDY, time12) {
    global SLOW_ACTIVATE_DELAY, SLOW_AFTER_MSG, SLOW_AFTER_SCHED, SLOW_AFTER_DT_PASTE, SLOW_AFTER_ENTER
    if !FocusChrome() {
        MsgBox("Chrome not found/active. Open the chat window first.")
        return false
    }
    Sleep SLOW_ACTIVATE_DELAY
    SendText msgText
    Sleep SLOW_AFTER_MSG
    Send "^!{Enter}"
    Sleep SLOW_AFTER_SCHED
    SendText dateMDY . " " . time12
    Sleep SLOW_AFTER_DT_PASTE
    Send "{Enter}"
    Sleep SLOW_AFTER_ENTER
    return true
}

; ===================== CAR QUOTE BUILDER =====================
BuildMessage(leadName, carCount) {
    global agentName, agentEmail
    greeting := (A_Hour < 12) ? "Buenos días" : "Buenas tardes"

    leadName := ProperCase(leadName)

    vehLine := (carCount >= 2)
        ? "Hicimos la cotización para el seguro de sus carros.`n"
        : "Hicimos la cotización para el seguro de su carro.`n"

    ; If carCount=5, add the BI/PD line right after "sus carros", with one blank line after it
    coverageLine := ""
    if (carCount = 5) {
        coverageLine := "`n`nBodily Injury $100k per person $300k per ocurrence`nProperty Damage $100k per ocurrence"   ;
    }

    ; Price map (includes 4 cars = $397, and 5 uses same $397 per your note)
    prices := Map(0, "$98", 1, "$117", 2, "$176", 3, "$284", 4, "$397", 5, "$397")
    price  := prices.Has(carCount) ? prices[carCount] : "$127"

    coverageSuffix := (carCount = 0) ? " al mes." : " al mes."

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
    , greeting, leadName
    , vehLine, coverageLine
    , price, coverageSuffix
    , agentName, agentEmail)
}

^!1:: { ; Ctrl+Alt+1 -> build car message (leaves name in clipboard)
    if !ClipWait(1) {
        MsgBox("Clipboard empty. Copy the lead's name first.")
        return
    }
    leadName := CleanName(A_Clipboard)
    if (leadName = "" || StrLen(leadName) > 30) {
        MsgBox("Please copy just the lead's name (<= 30 chars).")
        return
    }

    ib := InputBox(
        "Escribe 0, 1, 2, 3, 4 o 5:" . "`n"
        . "0 = auto muy antiguo ($98, sin 'FULL COVERAGE')" . "`n"
        . "1 = su carro ($117, FULL COVERAGE)" . "`n"
        . "2 = sus carros ($176, FULL COVERAGE)" . "`n"
        . "3 = sus carros ($284, FULL COVERAGE)" . "`n"
        . "4 = sus carros ($397, FULL COVERAGE)" . "`n"
        . "5 = sus carros ($397, FULL COVERAGE, 100/300)",
        "Número de vehículos",
        "w420 h280"
    )
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

; =================== FOLLOW-UP QUEUE BUILDER ===================
BuildFollowupQueue(leadName, offset) {
    global agentName, configDays, holidays
    leadName := ProperCase(leadName)

    ; Map “blocks” to configurable calendar days
    ; Block A = formerly Day 1 (4 msgs)
    ; Block B = formerly Day 2 (2 msgs)
    ; Block C = formerly Day 4 (2 msgs)
    ; Block D = formerly Day 5 (2 msgs)
    dA := configDays[1]
    dB := configDays[2]
    dC := configDays[3]
    dD := configDays[4]

    ; Resolve actual dates from business-day numbers
    DA_date := BusinessDateForDay(dA, holidays)
    DB_date := BusinessDateForDay(dB, holidays)
    DC_date := BusinessDateForDay(dC, holidays)
    DD_date := BusinessDateForDay(dD, holidays)

    ; Times (already staggered per day)
    t1_1 := TimeWithOffset( 9, 30, 30, offset)
    t1_2 := TimeWithOffset( 9, 31, 10, offset)
    t1_3 := TimeWithOffset( 9, 31, 30, offset)
    t1_4 := TimeWithOffset(10, 45,  0, offset)
    t2_1 := TimeWithOffset(16,  0, 10, offset)
    t2_2 := TimeWithOffset(16,  1, 10, offset)
    t4_1 := TimeWithOffset(16, 30,  0, offset)
    t4_2 := TimeWithOffset(16, 31,  0, offset)
    t5_1 := TimeWithOffset(12,  0,  0, offset)
    t5_2 := TimeWithOffset(12,  1,  0, offset)
    t6_1 := TimeWithOffset(6,  30,  0, offset)

msgs := [
    ; -------- Block A (config day dA) --------
    Map("day", dA, "seq", 1, "text", "Buenos días, " . leadName . ".", "date", DA_date, "time", t1_1),
    Map("day", dA, "seq", 2, "text", "Soy " . agentName . " de Allstate. Ya le preparé la cotización de su auto.", "date", DA_date, "time", t1_2),
    Map("day", dA, "seq", 3, "text", "En muchos casos logramos bajar el pago mensual sin quitar coberturas. Si gusta, se la resumo en 2 minutos por aquí.", "date", DA_date, "time", t1_3),
    Map("day", dA, "seq", 4, "text", "Si me responde “Sí”, se la envío ahora mismo.", "date", DA_date, "time", t1_4),

    ; -------- Block B (config day dB) --------
    Map("day", dB, "seq", 1, "text", "Hola, " . leadName . ".", "date", DB_date, "time", t2_1),
    Map("day", dB, "seq", 2, "text", "Hoy intenté comunicarme con usted porque todavía puedo revisar si califica a descuentos disponibles. Si me responde “Revisar”, yo me encargo de validar todo por usted.", "date", DB_date, "time", t2_2),

    ; -------- Block C (config day dC) --------
    Map("day", dC, "seq", 1, "text", "Buenas tardes, " . leadName . ".", "date", DC_date, "time", t4_1),
    Map("day", dC, "seq", 2, "text", "Esta semana hemos ayudado a varios conductores a comparar su póliza actual con Allstate y en muchos casos encontraron una mejor opción. Si me responde “Comparar”, reviso su caso y le digo honestamente si le conviene o no. Reply STOP to unsubscribe", "date", DC_date, "time", t4_2),

    ; -------- Block D (config day dD) --------
    Map("day", dD, "seq", 1, "text", leadName . ", sigo teniendo su cotización disponible, pero normalmente cierro los pendientes cuando no recibo respuesta.", "date", DD_date, "time", t5_1),
    Map("day", dD, "seq", 2, "text", "Si todavía quiere revisarla, respóndame “Continuar” y le envío el resumen por aquí.", "date", DD_date, "time", t5_2)
]
    return msgs
}


; =================== 6-DAY FOLLOW-UP (PASTE) ===================
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

    idx := IniRead(iniFile, "Times", "Offset", 0)
    idx := Mod(idx + 1, 60)
    IniWrite(idx, iniFile, "Times", "Offset")
    offset := idx

    ; Build queue with INI-based days
    msgs := BuildFollowupQueue(lead, offset)

    for m in msgs {
        ok := ScheduleMessage(m["text"], m["date"], m["time"])
        if !ok {
            MsgBox("Failed scheduling one of the messages. Stopping.")
            return
        }
    }
}

; =================== 6-DAY FOLLOW-UP (TYPE) ===================
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

    idx := IniRead(iniFile, "Times", "Offset", 0)
    idx := Mod(idx + 1, 60)
    IniWrite(idx, iniFile, "Times", "Offset")
    offset := idx

    ; Build queue with INI-based days
    msgs := BuildFollowupQueue(lead, offset)


    for m in msgs {
        ok := ScheduleMessageTyped(m["text"], m["date"], m["time"])
        if !ok {
            MsgBox("Failed scheduling one of the messages (typed mode). Stopping.")
            return
        }
    }
}

; =================== DAY PICKER (no 'start from message #') ===================
^!8:: {  ; earliest selected block = tomorrow; others keep spacing
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

    idx := IniRead(iniFile, "Times", "Offset", 0)
    idx := Mod(idx + 1, 60)
    IniWrite(idx, iniFile, "Times", "Offset")
    offset := idx

    fullMsgs := BuildFollowupQueue(lead, offset)

    ; Dynamic labels based on INI-configured days
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

    ; Stash controls on the picker to read them later inside SendSelectedBatch
    picker.cbD1  := cbD1
    picker.cbD2  := cbD2
    picker.cbD4  := cbD4
    picker.cbD5  := cbD5
    picker.ddMode := ddMode
    picker.fullMsgs := fullMsgs

    btnStart.OnEvent("Click", (*) => SendSelectedBatch(picker))
    btnCancel.OnEvent("Click", (*) => picker.Destroy())

    picker.Show()
}

; -------- Runs selected block(s) with relative spacing (earliest = tomorrow) --------
; Reads checkbox values directly from the picker (no stale captures)
SendSelectedBatch(picker) {
    global holidays

    try picker.Opt("-AlwaysOnTop")
    try picker.Hide()
    FocusChrome()
    Sleep 200

    ; Read current UI state
    useD1 := picker.cbD1.Value
    useD2 := picker.cbD2.Value
    useD4 := picker.cbD4.Value
    useD5 := picker.cbD5.Value
    modeText := picker.ddMode.Text
    fullMsgs := picker.fullMsgs

    selectedDays := []
    if (useD1) selectedDays.Push( fullMsgs[1]["day"] )  ; Block A day number
    if (useD2) selectedDays.Push( fullMsgs[5]["day"] )  ; Block B day number
    if (useD4) selectedDays.Push( fullMsgs[7]["day"] )  ; Block C day number
    if (useD5) selectedDays.Push( fullMsgs[9]["day"] )  ; Block D day number

    if (selectedDays.Length = 0) {
        MsgBox("Selecciona al menos un bloque (A, B, C o D).")
        return
    }

    ; earliest selected block becomes base → tomorrow (next business day)
    baseDay := 999
    for d in selectedDays
        if (d < baseDay)
            baseDay := d

    todayYMD := FormatTime(A_Now, "yyyyMMdd")
    baseYMD  := NextBusinessDateYYYYMMDD(todayYMD, holidays)  ; tomorrow

    ; Remap dates keeping original gaps and sort by (day, seq)
    toSend := []
    for m in fullMsgs {
        if (m.Has("day") && ArrContains(selectedDays, m["day"])) {
            daysAfter := m["day"] - baseDay
            targetYMD := AddBusinessDays(baseYMD, daysAfter, holidays)
            m["date"] := FormatTime(targetYMD . "000000", "MM/dd/yyyy")
            toSend.Push(m)
        }
    }
    toSend.Sort((a, b) => (a["day"] = b["day"]) ? (a["seq"] - b["seq"]) : (a["day"] - b["day"]))

    useTyped := (modeText = "Escritura estable (Chrome)")
    for m in toSend {
        FocusChrome()
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


