^!6:: ; Ctrl+Alt+6 - Full 6-Day Follow-Up Messaging with Simplified Time Rotation
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

    GetTimeBlockOffset(startHour, startMin, startSecond, offset) {
    FormatTime, today, , yyyyMMdd
    hour := SubStr("0" . startHour, -1)
    min := SubStr("0" . startMin, -1)
    sec := SubStr("0" . startSecond, -1)
    timeStr := today . hour . min . sec
    ; Build DateTime: yyyyMMddHHmmss
    datetime := today . hour . min . sec
    EnvAdd, datetime, %offset%, Minutes
    FormatTime, result, %datetime%, hh:mm:ss tt
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
lastDate := today
Loop, 10
{
	; Get the next valid business day
	nextDate := GetNextDate(lastDate)
	lastDate := nextDate

	; Format it as MM/dd/yyyy
	FormatTime, formattedDate, %nextDate%, MM/dd/yyyy

	; Create variable name dynamically (d1, d2, ..., d10)
	varName := "d" . A_Index

	; Assign the formatted date to the dynamically named variable
	%varName% := formattedDate
}


    offset := rotationIndex
    messages := []

    ; DAY 1 - next morning block
    t1 := GetTimeBlockOffset(9, 30, 30, offset)
    t2 := GetTimeBlockOffset(9, 31, 10, offset)
    t3 := GetTimeBlockOffset(9, 31, 30, offset)
    t4 := GetTimeBlockOffset(11, 45, 00, offset)
    messages.Push({text: "Buenos días " . name, date: d1, time: t1})
    messages.Push({text: "Le escribe Pablo Cabrera, desde Allstate Ins.", date: d1, time: t2})
    messages.Push({text: "Ayer le envié una cotización para el seguro de su vehículo, me gustaría saber si le estamos ofreciendo mejor precio.", date: d1, time: t3})
    messages.Push({text: "¿Le gustaría que se la reenvíe por aquí o prefiere por correo?", date: d1, time: t4})

    ; DAY 2 - afternoon block
    t1 := GetTimeBlockOffset(16, 00, 10, offset)
    t2 := GetTimeBlockOffset(16, 01, 10, offset)
    t2 := GetTimeBlockOffset(16, 01, 30, offset)
    messages.Push({text: "Hola " . name, date: d2, time: t1})
    messages.Push({text: "Le llame hoy para ver si tenía alguna duda sobre la cotización.", date: d2, time: t2})
    messages.Push({text: "Podemos revisar juntos para ver si califica a un descuento adicional. ¿Quiere que lo revise rápido?.", date: d2, time: t3})

    ; DAY 3 - morning block
    t1 := GetTimeBlockOffset(9, 30, 00, offset)
    t2 := GetTimeBlockOffset(9, 31, 00, offset)
    messages.Push({text: "Otros clientes con situaciones similares, están pagando lo mismo o incluso menos por su/s vehículo/s.", date: d3, time: t1})
    messages.Push({text: "¿Quiere que revisemos juntos para ver si califica?", date: d3, time: t2})

    ; DAY 4 - fixed noon block
    t1 := GetTimeBlockOffset(16, 45, 00, offset)
    t2 := GetTimeBlockOffset(16, 46, 00, offset)
    messages.Push({text: "Buenas tardes " . name, date: d4, time: t1})
    messages.Push({text: "Muchos clientes han estado mejorando sus precios.", date: d4, time: t2})
    messages.Push({text: "Si desea, puedo comparar su póliza actual con lo que Allstate está ofreciendo.", date: d4, time: t2})

    ; DAY 5 - fixed noon block
    t1 := GetTimeBlockOffset(12, 00, 00, offset)
    t2 := GetTimeBlockOffset(12, 01, 00, offset)
    messages.Push({text: name . " solo quería confirmar si aún está interesado. Si no recibo respuesta, cerraré esta cotización.", date: d5, time: t1})
    messages.Push({text: "Si en algún momento desea retomarla, estaré aquí para ayudarle sin compromiso.", date: d5, time: t2})
    messages.Push({text: "Gracias por su tiempo.", date: d5, time: t2})

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
