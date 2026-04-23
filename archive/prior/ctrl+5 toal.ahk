^!5:: ; Ctrl+Alt+5 - Spanish Follow-Up, Pasting Messages
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
    nextDate := GetNextDate(lastDate)
    lastDate := nextDate  ; Update for next iteration

    FormatTime, formattedDate, %nextDate%, MM/dd/yyyy

    ; Dynamically assign to d1, d2, ..., d10
    varName := "d" . A_Index
    %varName% := formattedDate
}

    offset := rotationIndex
    messages := []

; Day 1 – 9:30 AM
t1 := GetTimeBlockOffset(9, 30, 30, offset)
t2 := GetTimeBlockOffset(9, 31, 10, offset)
t3 := GetTimeBlockOffset(9, 31, 30, offset)
messages.Push({text: "Hi there, hope you're doing well.", date: d1, time: t1})
messages.Push({text: "We worked with you before on a commercial auto policy.", date: d1, time: t2})
messages.Push({text: "Do you also have a personal car I could quote for you? Just reply YES and I’ll take care of it.", date: d1, time: t3})

; Day 3 – 12:30 PM
t1 := GetTimeBlockOffset(12, 30, 30, offset)
t2 := GetTimeBlockOffset(12, 31, 10, offset)
t3 := GetTimeBlockOffset(12, 31, 30, offset)
messages.Push({text: "Many of our business clients also insure their personal vehicles with us.", date: d3, time: t1})
messages.Push({text: "It keeps everything in one place and often leads to better pricing.", date: d3, time: t2})
messages.Push({text: "Want to see what you'd qualify for? Just say “Go” and I’ll run it.", date: d3, time: t3})

; Day 4 – 4:00 PM
t1 := GetTimeBlockOffset(16, 0, 30, offset)
t2 := GetTimeBlockOffset(16, 1, 10, offset)
t3 := GetTimeBlockOffset(16, 1, 30, offset)
messages.Push({text: "One client added two family cars and saved over $600/year.", date: d4, time: t1})
messages.Push({text: "It only took him 2 minutes and we handled everything for him.", date: d4, time: t2})
messages.Push({text: "Can I check if that kind of deal works for you too?", date: d4, time: t3})

; Day 6 – 11:00 AM
t1 := GetTimeBlockOffset(11, 0, 30, offset)
t2 := GetTimeBlockOffset(11, 1, 10, offset)
t3 := GetTimeBlockOffset(11, 1, 30, offset)
messages.Push({text: "I can send you a quick personal quote—no commitment needed.", date: d6, time: t1})
messages.Push({text: "Just reply with year, make, model, and zip code.", date: d6, time: t2})
messages.Push({text: "Send me the car info now and I’ll text back the numbers.", date: d6, time: t3})

; Day 7 – 9:00 AM
t1 := GetTimeBlockOffset(9, 0, 30, offset)
t2 := GetTimeBlockOffset(9, 1, 10, offset)
t3 := GetTimeBlockOffset(9, 1, 30, offset)
messages.Push({text: "I’m wrapping up personal quotes for the week.", date: d7, time: t1})
messages.Push({text: "If you'd like your vehicle included before I close it out...", date: d7, time: t2})
messages.Push({text: "Text “Add Me” and I’ll include you right now.", date: d7, time: t3})

; Day 8 – 10:00 AM
t1 := GetTimeBlockOffset(10, 0, 30, offset)
t2 := GetTimeBlockOffset(10, 1, 10, offset)
t3 := GetTimeBlockOffset(10, 1, 30, offset)
messages.Push({text: "Final reminder—I won’t message again after today.", date: d8, time: t1})
messages.Push({text: "If your car isn’t insured yet, I’d really suggest checking.", date: d8, time: t2})
messages.Push({text: "Would you like me to handle it before I close your file?", date: d8, time: t3})

; Day 10 – 11:45 AM
t1 := GetTimeBlockOffset(11, 45, 30, offset)
t2 := GetTimeBlockOffset(11, 46, 10, offset)
t3 := GetTimeBlockOffset(11, 46, 30, offset)
messages.Push({text: "Haven’t heard back, so I’ll go ahead and close your file.", date: d10, time: t1})
messages.Push({text: "If you ever want to review your personal insurance, you’ve got my number.", date: d10, time: t2})
messages.Push({text: "Just reply “QUOTE” anytime and I’ll jump back in.", date: d10, time: t3})



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
