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

LoadHolidayList(path) {
    datesText := IniRead(path, "Holidays", "Dates", "")
    dates := []
    for item in StrSplit(datesText, ",") {
        clean := Trim(item)
        if (clean != "")
            dates.Push(clean)
    }
    return dates
}

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
    todayYMD := FormatTime(A_Now, "yyyyMMdd")
    nowHHMMSS := Integer(FormatTime(A_Now, "HHmmss"))

    noonTime := TimeWithOffset(12, 0, 0, offset)
    dt12 := todayYMD . "120000"
    dt12 := DateAdd(dt12, offset, "M")
    noonHHMMSS := Integer(FormatTime(dt12, "HHmmss"))

    if (nowHHMMSS <= noonHHMMSS)
        return Map("date", todayDate, "time", noonTime)

    if (nowHHMMSS <= 175500)
        return Map("date", todayDate, "time", "05:55:00 PM")

    nextBizYMD := NextBusinessDateYYYYMMDD(todayYMD, holidays)
    nextBizDate := FormatTime(nextBizYMD . "000000", "MM/dd/yyyy")
    return Map("date", nextBizDate, "time", noonTime)
}

NextRotationOffset() {
    global settingsFile

    idx := Integer(IniRead(settingsFile, "Times", "Offset", "0"))
    idx := Mod(idx + 1, 60)
    IniWrite(idx, settingsFile, "Times", "Offset")
    PersistRunState("rotation-advanced")
    return idx
}
