BeginAutomationRun() {
    global StopFlag
    StopFlag := false
    PersistRunState("automation-begin")
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
