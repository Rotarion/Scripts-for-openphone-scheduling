^!6:: ; Ctrl+Alt+6 - Full 6-Day Follow-Up Messaging with Simplified Time Rotation
{
    ; Validate clipboard and clean up name input
    ClipWait, 1
    StringReplace, name, Clipboard, `r`n,, All
    StringReplace, name, name, `n,, All
   name := Trim(name)
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
Loop, 12
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

    ; DAY 1 – follow-up after quote sent
    t1 := GetTimeBlockOffset(9, 30, 00, offset)
    t2 := GetTimeBlockOffset(9, 31, 10, offset)
    t3 := GetTimeBlockOffset(9, 32, 30, offset)
    t4 := GetTimeBlockOffset(10, 34, 00, offset)
    messages.Push({text: "Buenos días " . name, date: d1, time: t1})
    messages.Push({text: "Soy Pablo Cabrera de Allstate. Solo quería confirmar que recibió la cotización que le envié.", date: d1, time: t2})
    messages.Push({text: "Estoy a sus órdenes si desea revisar los detalles o tiene dudas.", date: d1, time: t3})
    messages.Push({text: "¿Le gustaría que se la reenvíe por aquí o prefiere por correo?", date: d1, time: t4})

    ; DAY 2 – follow-up after missed call
    t1 := GetTimeBlockOffset(10, 00, 00, offset)
    t2 := GetTimeBlockOffset(10, 01, 10, offset)
    t3 := GetTimeBlockOffset(10, 02, 30, offset)
    t4 := GetTimeBlockOffset(10, 03, 40, offset)
    messages.Push({text: "Hola " . name, date: d2, time: t1})
    messages.Push({text: "le llamé hoy más temprano sobre su cotización.Lamento no haber podido hablar con usted.", date: d2, time: t2})
    messages.Push({text: "Estoy disponible para asesorarle cuando guste.", date: d2, time: t3})
    messages.Push({text: "¿Hay algún horario en que prefiere que le vuelva a llamar?", date: d2, time: t4})

    ; DAY 3 – others are saving
    t1 := GetTimeBlockOffset(11, 30, 00, offset)
    t2 := GetTimeBlockOffset(11, 31, 15, offset)
    t3 := GetTimeBlockOffset(11, 32, 30, offset)
    messages.Push({text: "Muchos clientes similares han logrado ahorrar al cambiarse a Allstate.", date: d3, time: t1})
    messages.Push({text: "Protegen a su familia y cuidan su presupuesto sin perder cobertura.", date: d3, time: t2})
    messages.Push({text: "¿Le gustaría conocer cómo podría beneficiarse usted también?", date: d3, time: t3})

    ; DAY 4 – second call attempt
    t1 := GetTimeBlockOffset(12, 00, 00, offset)
    t2 := GetTimeBlockOffset(12, 01, 15, offset)
    t3 := GetTimeBlockOffset(12, 02, 30, offset)
    t4 := GetTimeBlockOffset(12, 03, 30, offset)
    messages.Push({text: "Hola " . name, date: d4, time: t1})
    messages.Push({text: "intenté comunicarme nuevamente hoy. Sé que su tiempo es valioso y no quiero incomodarle.", date: d4, time: t2})
    messages.Push({text: "Si prefiere seguir por mensaje, podemos continuar por aquí.", date: d4, time: t3})
    messages.Push({text: "¿Le gustaría que sigamos conversando por este medio?", date: d4, time: t4})

    ; DAY 5 – quote still active
    t1 := GetTimeBlockOffset(9, 45, 00, offset)
    t2 := GetTimeBlockOffset(9, 46, 20, offset)
    t3 := GetTimeBlockOffset(9, 47, 45, offset)
    messages.Push({text: "Buenas tardes " . name, date: d5, time: t1})
    messages.Push({text: "su cotización de Allstate sigue activa.Aún puede aprovechar esta oportunidad de ahorro y protección.", date: d5, time: t2})
    messages.Push({text: "¿Hay alguna inquietud que le impida avanzar? Estoy para ayudarle.", date: d5, time: t3})

    ; DAY 6 – value and benefits
    t1 := GetTimeBlockOffset(14, 00, 00, offset)
    t2 := GetTimeBlockOffset(14, 01, 15, offset)
    t3 := GetTimeBlockOffset(14, 02, 30, offset)
    messages.Push({text: "Con Allstate obtiene más que buen precio: tranquilidad y servicio confiable.", date: d6, time: t1})
    messages.Push({text: "Protegemos autos y familias con rapidez en reclamos y respaldo sólido.", date: d6, time: t2})
    messages.Push({text: "¿Le gustaría conocer más beneficios antes de tomar una decisión?", date: d6, time: t3})

    ; DAY 7 – family security
    t1 := GetTimeBlockOffset(10, 30, 00, offset)
    t2 := GetTimeBlockOffset(10, 31, 30, offset)
    t3 := GetTimeBlockOffset(10, 33, 00, offset)
    t4 := GetTimeBlockOffset(10, 34, 30, offset)
    messages.Push({text: "Allstate lleva casi 90 años protegiendo familias como la suya.", date: d7, time: t1})
    messages.Push({text: "Su tranquilidad y la seguridad de sus seres queridos son nuestra prioridad.", date: d7, time: t2})
    messages.Push({text: "¿Le gustaría sentir ese respaldo para su auto y su familia?", date: d7, time: t3})

    ; DAY 8 – easy process
    t1 := GetTimeBlockOffset(15, 00, 00, offset)
    t2 := GetTimeBlockOffset(15, 01, 15, offset)
    t3 := GetTimeBlockOffset(15, 02, 30, offset)
    messages.Push({text: "Obtener su póliza es muy sencillo.", date: d8, time: t1})
    messages.Push({text: "Puedo ayudarle en minutos por teléfono o mensaje, sin complicaciones.", date: d8, time: t2})
    messages.Push({text: "Si desea avanzar, solo avíseme y me encargo de todo.", date: d8, time: t3})

    ; DAY 9 – personal care
    t1 := GetTimeBlockOffset(13, 00, 00, offset)
    t2 := GetTimeBlockOffset(13, 01, 20, offset)
    t3 := GetTimeBlockOffset(13, 02, 40, offset)
    messages.Push({text: "Como su agente, estoy aquí para brindarle atención personal en español.", date: d9, time: t1})
    messages.Push({text: "En Allstate tratamos a nuestros clientes como familia.", date: d9, time: t2})
    messages.Push({text: "¿Le parece si damos el siguiente paso juntos?", date: d9, time: t3})

    ; DAY 10 – final reminder
    t1 := GetTimeBlockOffset(10, 00, 00, offset)
    t2 := GetTimeBlockOffset(10, 01, 30, offset)
    t3 := GetTimeBlockOffset(10, 02, 45, offset)
    messages.Push({text: "Hola " . name, date: d10, time: t1})
    messages.Push({text: "su cotización está por expirar muy pronto. No quiero que pierda esta oportunidad de proteger su auto y su familia.", date: d10, time: t2})
    messages.Push({text: "Si desea aprovecharla, estoy disponible para ayudarle ahora mismo.", date: d10, time: t3})

    ; DAY 11 – last opportunity
    t1 := GetTimeBlockOffset(9, 00, 00, offset)
    t2 := GetTimeBlockOffset(9, 01, 15, offset)
    t3 := GetTimeBlockOffset(9, 02, 30, offset)
    t4 := GetTimeBlockOffset(9, 03, 45, offset)
    messages.Push({text: "Hola " . name . ", ", date: d11, time: t1})
    messages.Push({text: "le agradezco su tiempo durante este seguimiento. Esta será mi última comunicación sobre su cotización.", date: d11, time: t2})
    messages.Push({text: "Si más adelante desea proteger su auto, estaré aquí para servirle.", date: d11, time: t3})
    messages.Push({text: "Gracias por considerar a Allstate.", date: d11, time: t4})


   for index, msg in messages {
    Clipboard := ""
    Sleep, 50
    Clipboard := msg.text
    ClipWait, 1
    Sleep, 100
    Send ^v
    Send ^!{Enter}
    Sleep, 400

    Clipboard := ""
    Sleep, 50
    Clipboard := msg.date . " " . msg.time
    ClipWait, 1
    Sleep, 100
    Send ^v
    Sleep, 100
    Send {Enter}
    Sleep, 200
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