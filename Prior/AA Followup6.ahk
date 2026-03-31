 ^!6:: ; Ctrl+Alt+6- Full 6-Day Follow-Up Messaging with Simplified Time Rotation
{
    ; Validate clipboard and clean up name input
    ClipWait, 1
    StringReplace, name, Clipboard, `r`n,, All
    StringReplace, name, name, `n,, All
    if (StrLen(name) > 30)
    {
        MsgBox, Clipboard seems too long. Please copy just the lead's name and try again.
        return
    }

    ; Load rotation index (0-59)
    iniFile := A_ScriptDir . "\\time_rotation.ini"
    IniRead, rotationIndex, %iniFile%, Times, Offset, 0
    rotationIndex := Mod(rotationIndex + 1, 60)  ; 60-minute window
    IniWrite, %rotationIndex%, %iniFile%, Times, Offset

    GetTimeBlockOffset(startHour, startMin, offset) {
        FormatTime, today, , yyyyMMdd
        hour := SubStr("0" . startHour, -1)
        min := SubStr("0" . startMin, -1)
        timeStr := today . hour . min . "00"
        FormatTime, full, %timeStr%, yyyyMMddHHmmss
        EnvAdd, full, %offset%, Minutes
        FormatTime, result, %full%, hh:mm tt
        return result
    }

    holidays := ["01/01/2025", "07/04/2025", "11/27/2025", "12/25/2025"]

    IsHoliday(dateStr) {
        global holidays
        for index, holiday in holidays {
            if (holiday == dateStr)
                return true
        }
        return false
    }

    GetNextDate(startDate) {
        global holidays
        Loop {
            EnvAdd, startDate, 1, Days
            FormatTime, wday, %startDate%, WDay
            FormatTime, checkDate, %startDate%, MM/dd/yyyy
            if (wday != 1 && wday != 7 && !IsHoliday(checkDate)) {
                return startDate
            }
        }
    }

    ; Generate valid message days using datetime values
    today := A_Now
    day2 := GetNextDate(today)
    day3 := GetNextDate(day2)
    day4 := GetNextDate(day3)
    day5 := GetNextDate(day4)
    day6 := GetNextDate(day5)
    day7 := GetNextDate(day6)
    day8 := GetNextDate(day7) 


 

    ; Format for scheduling display
    FormatTime, d2, %day2%, MM/dd/yyyy
    FormatTime, d3, %day3%, MM/dd/yyyy
    FormatTime, d4, %day4%, MM/dd/yyyy
    FormatTime, d5, %day5%, MM/dd/yyyy
    FormatTime, d6, %day6%, MM/dd/yyyy
    FormatTime, d7, %day7%, MM/dd/yyyy
    FormatTime, d8, %day8%, MM/dd/yyyy 
    FormatTime, d9, %day9%, MM/dd/yyyy
    FormatTime, d10, %day10%, MM/dd/yyyy


    offset := rotationIndex
    messages := []

    ; DAY 2 - morning block
    t1 := GetTimeBlockOffset(9, 30, offset)
    t2 := GetTimeBlockOffset(9, 31, offset)
    t3 := GetTimeBlockOffset(9, 32, offset)
    t4 := GetTimeBlockOffset(9, 40, offset)
    messages.Push({text: "Muy Buenos días " . name, date: d2, time: t1})
    messages.Push({text: "Le saluda Patricia Lucena, desde Allstate Ins.", date: d2, time: t2})
    messages.Push({text: "Hace poco le compartí una cotización para asegurar su vehículo. ¿Tuvo oportunidad de revisarla? Me encantaría saber si le estamos ofreciendo el mejor precio posible.", date: d2, time: t3})
    messages.Push({text: "¿Le gustaría que se la reenvíe por aquí o prefiere por correo?", date: d2, time: t4})


    ; DAY 4 - afternoon block
    t1 := GetTimeBlockOffset(15, 00, offset)
    t2 := GetTimeBlockOffset(15, 01, offset)
    messages.Push({text: "Hola de nuevo, " . name . " 👋", date: d4, time: t1})
    messages.Push({text: "Le llamé hoy para ver si tenía alguna duda sobre la cotización que le envié.", date: d4, time: t2})
    messages.Push({text: "Si desea, con gusto puedo revisarla de nuevo o incluir otro vehículo para brindarle una mejor propuesta.", date: d4, time: t2})

    ; DAY 6 - morning block
    t1 := GetTimeBlockOffset(10, 30, offset)
    t2 := GetTimeBlockOffset(10, 31, offset)
    messages.Push({text: "Otros clientes con situaciones similares, están pagando lo mismo o incluso menos por su/s vehículo/s.", date: d6, time: t1})
    messages.Push({text: "¿Le gustaría que lo revisemos juntos para ver si usted también califica? Puede ser una excelente oportunidad para ahorrar.", date: d6, time: t2})

    ; DAY 9 - fixed noon block
    messages.Push({text: name . " solo quería confirmar si todavía está interesado en su cotización,de no ser asi, dejemelo saber para asi cerrar la propuesta esta semana para dejar espacio a nuevos casos.", date: d8, time: "14:00 PM"})
    messages.Push({text: "Si en algún momento desea retomarla, estaré aquí para ayudarle sin compromiso.", date: d8, time: "14:01 PM"})
    messages.Push({text: "Gracias por su tiempo.", date: d8, time: "14:02 PM"})

   ; Loop to send/schedule
    for index, msg in messages {
        WinActivate, ahk_exe chrome.exe
        Sleep, 300
        Send, % msg.text
        Sleep, 300
        Send, ^!{Enter}
        Sleep, 400
        Send, % msg.date . " " . msg.time
        Sleep, 200
        Send, {Enter}
        Sleep, 500
    }
}
return

; Kill switch
-::ExitApp

; Reset rotation index manually
^!r::
IniWrite, 0, %A_ScriptDir%\time_rotation.ini, Times, Offset
MsgBox, Rotation index reset to 0.
return
