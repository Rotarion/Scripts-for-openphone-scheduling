^!4:: ; Ctrl+Alt+4 - Spanish Follow-Up, Pasting Messages
{
    ; Clipboard cleanup
    ClipWait, 1
    StringReplace, name, Clipboard, `r`n,, All
    StringReplace, name, name, `n,, All
    if (StrLen(name) > 30)
    {
        MsgBox, El nombre copiado es muy largo. Copia solo el nombre del cliente.
        return
    }

    ; Rotation offset
    iniFile := A_ScriptDir . "\\time_rotation2.ini"
    IniRead, rotationIndex, %iniFile%, Times, Offset, 0
    rotationIndex := Mod(rotationIndex + 1, 60)
    IniWrite, %rotationIndex%, %iniFile%, Times, Offset

    ; Functions
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

    ; Día 1 – 9:30 AM
t1 := GetTimeBlockOffset(9, 30, 30, offset)
t2 := GetTimeBlockOffset(9, 31, 10, offset)
t3 := GetTimeBlockOffset(9, 31, 30, offset)
messages.Push({text: "Hola, disculpa el mensaje inesperado.", date: d1, time: t1})
messages.Push({text: "Te conocimos por tu póliza comercial hace un tiempo.", date: d1, time: t2})
messages.Push({text: "Solo quería saber… ¿también manejas algún auto personal que quieras cotizar?", date: d1, time: t3})

; Día 3 – 12:30 PM
t1 := GetTimeBlockOffset(12, 30, 30, offset)
t2 := GetTimeBlockOffset(12, 31, 10, offset)
t3 := GetTimeBlockOffset(12, 31, 30, offset)
messages.Push({text: "Muchos de nuestros clientes comerciales también protegen sus autos personales con nosotros.", date: d3, time: t1})
messages.Push({text: "Así mantienen todo en un solo lugar y reciben mejores beneficios.", date: d3, time: t2})
messages.Push({text: "¿Tienes algún auto tuyo o de casa que quieres que coticemos?", date: d3, time: t3})

; Día 4 – 4:00 PM
t1 := GetTimeBlockOffset(16, 0, 30, offset)
t2 := GetTimeBlockOffset(16, 1, 10, offset)
t3 := GetTimeBlockOffset(16, 1, 30, offset)
messages.Push({text: "Un cliente que tenía su póliza comercial con nosotros agregó 2 autos personales y bajó costos.", date: d4, time: t1})
messages.Push({text: "Muchas veces, se paga menos al juntar coberturas.", date: d4, time: t2})
messages.Push({text: "¿Te gustaría que te cotice algo similar?", date: d4, time: t3})

; Día 6 – 11:00 AM
t1 := GetTimeBlockOffset(11, 0, 30, offset)
t2 := GetTimeBlockOffset(11, 1, 10, offset)
t3 := GetTimeBlockOffset(11, 1, 30, offset)
messages.Push({text: "Puedo enviarte una cotización sin compromiso si me das datos de tu auto personal.", date: d6, time: t1})
messages.Push({text: "Solo marca, modelo, año y código postal.", date: d6, time: t2})
messages.Push({text: "¿Te parece si lo revisamos juntos?", date: d6, time: t3})

; Día 7 – 9:00 AM
t1 := GetTimeBlockOffset(9, 0, 30, offset)
t2 := GetTimeBlockOffset(9, 1, 10, offset)
t3 := GetTimeBlockOffset(9, 1, 30, offset)
messages.Push({text: "Estoy cerrando cotizaciones familiares esta semana.", date: d7, time: t1})
messages.Push({text: "Si tienes un auto personal que quieres revisar, dime ‘sí’ y te escribo lo que necesito.", date: d7, time: t2})

; Día 8 – 10:00 AM
t1 := GetTimeBlockOffset(10, 0, 30, offset)
t2 := GetTimeBlockOffset(10, 1, 10, offset)
t3 := GetTimeBlockOffset(10, 1, 30, offset)
messages.Push({text: "Último intento: si tienes un vehículo personal sin asegurar o que podríamos revisar, avísame hoy.", date: d8, time: t1})
messages.Push({text: "Te lo dejo todo por escrito y sin compromiso.", date: d8, time: t2})

; Día 10 – 11:45 AM
t1 := GetTimeBlockOffset(11, 45, 30, offset)
t2 := GetTimeBlockOffset(11, 46, 10, offset)
t3 := GetTimeBlockOffset(11, 46, 30, offset)
messages.Push({text: "No tuve respuesta estos días, así que cerraré tu expediente.", date: d10, time: t1})
messages.Push({text: "Si en el futuro quieres revisar tus seguros personales, puedes escribirme directo.", date: d10, time: t2})
messages.Push({text: "Gracias por tu tiempo. ¡Que tengas un excelente mes!", date: d10, time: t3})


     ; Loop to send/schedule
     for index, msg in messages {
        Clipboard := msg.text
	ClipWait, 1
        Send ^v
        Send ^!{Enter}
	Sleep, 400
        Clipboard := msg.date . " " . msg.time
	ClipWait, 1
        Sleep, 100
        Send ^v
        Sleep, 100
        Send {Enter}
        Sleep, 100
    }
}
return

; Emergency kill switch
-::ExitApp
