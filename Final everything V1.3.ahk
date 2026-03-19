#Requires AutoHotkey v2.0
#SingleInstance Force
SendMode "Input"

; ===================== CONFIG =====================
iniFile   := A_ScriptDir "\time_rotation.ini"
holidays := ["01/01/2026","05/25/2026","06/19/2026","07/03/2026","07/04/2026","09/07/2026","11/26/2026","12/25/2026"]  ; MM/dd/yyyy

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

JoinArray(arr, delim := "") {
    out := ""
    for i, item in arr
        out .= (i > 1 ? delim : "") . item
    return out
}

FocusEdge() {
    if WinExist("ahk_exe msedge.exe") {
        WinActivate
        WinWaitActive "ahk_exe msedge.exe",, 2
        return true
    }
    return false
}

GetProspectSelectorMap() {
    return Map(
        "FIRST_NAME", "ConsumerData.People[0].Name.GivenName",
        "LAST_NAME", "ConsumerData.People[0].Name.Surname",
        "DOB", "ConsumerData.People[0].Personal.BirthDt",
        "GENDER", "ConsumerData.People[0].Personal.GenderCd.SrcCd",
        "ADDRESS_1", "ConsumerData.Assets.Properties[0].Addr.Addr1",
        "APT_SUITE", "ConsumerData.Assets.Properties[0].Addr.AdditionalAddressFields.Apartment",
        "BUILDING", "ConsumerData.Assets.Properties[0].Addr.AdditionalAddressFields.Building",
        "RR_NUMBER", "ConsumerData.Assets.Properties[0].Addr.AdditionalAddressFields.RuralRoute",
        "LOT_NUMBER", "ConsumerData.Assets.Properties[0].Addr.AdditionalAddressFields.LotNumber",
        "CITY", "ConsumerData.Assets.Properties[0].Addr.City",
        "STATE", "ConsumerData.Assets.Properties[0].Addr.StateProvCd.SrcCd",
        "ZIP", "ConsumerData.Assets.Properties[0].Addr.PostalCode",
        "PHONE", "ConsumerData.People[0].Communications.PhoneNumber"
    )
}

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
        "PHONE", ""
    )
}

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

    return NormalizeRawLeadToProspect(raw)
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

    ; ignore junk helper lines
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

    FlushPair := (*) => (
        currentLabel != "" ? StoreLabeledField(data, currentLabel, currentValue) : 0
    )

    for _, line in lines {
        line := Trim(line)
        if (line = "")
            continue

        ; Match label-only line like "Name:" or "First Name::"
        if RegExMatch(line, "^([^:]+:?):\s*$", &m) {
            FlushPair()
            currentLabel := Trim(m[1], " :")
            currentValue := ""
            continue
        }

        ; Match inline label:value on one line
        if RegExMatch(line, "^([^:]+:?):\s*(.+)$", &m) {
            FlushPair()
            currentLabel := Trim(m[1], " :")
            currentValue := Trim(m[2])
            FlushPair()
            currentLabel := ""
            currentValue := ""
            continue
        }

        ; Otherwise treat as value for previous label
        if (currentLabel != "" && currentValue = "") {
            currentValue := line
        } else if (currentLabel != "") {
            currentValue .= " " line
        }
    }

    FlushPair()

    ; Priority mapping
    if data.Has("First Name")
        fields["FIRST_NAME"] := ProperCasePhrase(data["First Name"])
    else if data.Has("Name")
        ApplyLeadName(fields, RegExReplace(data["Name"], "i)^PERSONAL LEAD\s*-\s*"))

    if data.Has("Last Name")
        fields["LAST_NAME"] := ProperCasePhrase(data["Last Name"])

    ; Fallback from Contact if needed
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

    NormalizeAddressMap(fields)
    fields["DOB"] := NormalizeDOB(fields["DOB"])
    fields["GENDER"] := NormalizeGender(fields["GENDER"])
    fields["STATE"] := NormalizeState(fields["STATE"])
    fields["ZIP"] := NormalizeZip(fields["ZIP"])
    fields["PHONE"] := NormalizePhone(fields["PHONE"])

    return fields
}

ParseLabeledLeadRaw(raw) {
    lines := StrSplit(StrReplace(raw, "`r", ""), "`n")
    data := Map()
    currentLabel := ""
    currentValue := ""

    FlushPair := (*) => (
        currentLabel != "" ? StoreLabeledField(data, currentLabel, currentValue) : 0
    )

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

    fields["DOB"] := NormalizeDOB(fields["DOB"])
    fields["GENDER"] := NormalizeGender(fields["GENDER"])
    fields["STATE"] := NormalizeState(fields["STATE"])
    fields["ZIP"] := NormalizeZip(fields["ZIP"])
    fields["PHONE"] := NormalizePhone(fields["PHONE"])
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

zipIdx := 0
for i, token in tokens {
    if IsTimestampToken(token)
        continue
    if IsPhoneToken(token)
        continue
    if IsEmailToken(token)
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

ProperCasePhrase(str) {
    text := Trim(RegExReplace(str, "\s+", " "))
    if (text = "")
        return ""

    parts := StrSplit(text, " ")
    out := ""
    for i, part in parts {
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
    for i, piece in parts {
        out .= (i > 1 ? "-" : "") . StrUpper(SubStr(piece, 1, 1)) . SubStr(piece, 2)
    }
    return out
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

NormalizeAddressMap(fields) {
    address1 := fields["ADDRESS_1"]
    city     := fields["CITY"]
    state    := fields["STATE"]
    zip      := fields["ZIP"]
    aptSuite := fields["APT_SUITE"]

    ; If address line contains extra city/state/zip junk, strip it out first
    if (address1 != "") {
        ExtractAddressTail(&address1, &city, &state, &zip, &aptSuite)
        SetAddressFields(fields, address1)
    }

    ; If city itself contains state/zip, split it
    NormalizeCityStateZipFields(&city, &state, &zip)

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

    ; remove trailing commas/periods
    city := RegExReplace(city, "[,\.]+\s*$", "")

    ; if city contains zip, remove it
    city := RegExReplace(city, "\b\d{5}(?:-\d{4})?\b", "")

    ; if city contains state abbreviation at the end, remove it
    city := RegExReplace(city, "i)\b(AL|AK|AZ|AR|CA|CO|CT|DC|DE|FL|GA|HI|IA|ID|IL|IN|KS|KY|LA|MA|MD|ME|MI|MN|MO|MS|MT|NC|ND|NE|NH|NJ|NM|NV|NY|OH|OK|OR|PA|RI|SC|SD|TN|TX|UT|VA|VT|WA|WI|WV|WY)\b\s*$", "")

    ; if city contains full state name at the end, remove it
    city := RegExReplace(city, "i)\b(alabama|alaska|arizona|arkansas|california|colorado|connecticut|delaware|district of columbia|florida|georgia|hawaii|idaho|illinois|indiana|iowa|kansas|kentucky|louisiana|maine|maryland|massachusetts|michigan|minnesota|mississippi|missouri|montana|nebraska|nevada|new hampshire|new jersey|new mexico|new york|north carolina|north dakota|ohio|oklahoma|oregon|pennsylvania|rhode island|south carolina|south dakota|tennessee|texas|utah|vermont|virginia|washington|west virginia|wisconsin|wyoming)\b\s*$", "")

    city := RegExReplace(city, "\s+", " ")
    city := Trim(city, " ,.-")
    return ProperCasePhrase(city)
}

NormalizeCityStateZipFields(&city, &state, &zip) {
    raw := Trim(city)
    if (raw = "")
        return

    ; capture ZIP from city field if present
    if (zip = "" && RegExMatch(raw, "\b(\d{5})(?:-\d{4})?\b", &m))
        zip := m[1]

    ; capture state from city field if present
    if (state = "") {
        if RegExMatch(raw, "i)\b(AL|AK|AZ|AR|CA|CO|CT|DC|DE|FL|GA|HI|IA|ID|IL|IN|KS|KY|LA|MA|MD|ME|MI|MN|MO|MS|MT|NC|ND|NE|NH|NJ|NM|NV|NY|OH|OK|OR|PA|RI|SC|SD|TN|TX|UT|VA|VT|WA|WI|WV|WY)\b", &m2)
            state := StrUpper(m2[1])
        else {
            ; try full state name
            words := StrSplit(raw, ",")
            for _, w in words {
                st := NormalizeState(w)
                if (st != "") {
                    state := st
                    break
                }
            }
            if (state = "") {
                ; fallback on whole string
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

    ; pull ZIP from address line if present
    if (zip = "" && RegExMatch(text, "\b(\d{5})(?:-\d{4})?\b", &mz))
        zip := mz[1]

    ; pull state from address line if present
    if (state = "") {
        if RegExMatch(text, "i)\b(AL|AK|AZ|AR|CA|CO|CT|DC|DE|FL|GA|HI|IA|ID|IL|IN|KS|KY|LA|MA|MD|ME|MI|MN|MO|MS|MT|NC|ND|NE|NH|NJ|NM|NV|NY|OH|OK|OR|PA|RI|SC|SD|TN|TX|UT|VA|VT|WA|WI|WV|WY)\b", &ms)
            state := StrUpper(ms[1])
        else {
            st := NormalizeState(text)
            if (st != "")
                state := st
        }
    }

    ; if city is blank, try to extract city from "... Miami, FL 33181"
    if (city = "") {
        if RegExMatch(text, "i),?\s*([A-Za-z]+(?:\s+[A-Za-z]+){0,2})\s*,?\s*(?:AL|AK|AZ|AR|CA|CO|CT|DC|DE|FL|GA|HI|IA|ID|IL|IN|KS|KY|LA|MA|MD|ME|MI|MN|MO|MS|MT|NC|ND|NE|NH|NJ|NM|NV|NY|OH|OK|OR|PA|RI|SC|SD|TN|TX|UT|VA|VT|WA|WI|WV|WY)\s*\d{5}(?:-\d{4})?$", &mc)
            city := mc[1]
    }

    ; strip only a TRAILING city/state/zip tail, not the house number
    text := RegExReplace(
        text,
        "i),?\s*[A-Za-z]+(?:\s+[A-Za-z]+){0,2}\s*,?\s*(AL|AK|AZ|AR|CA|CO|CT|DC|DE|FL|GA|HI|IA|ID|IL|IN|KS|KY|LA|MA|MD|ME|MI|MN|MO|MS|MT|NC|ND|NE|NH|NJ|NM|NV|NY|OH|OK|OR|PA|RI|SC|SD|TN|TX|UT|VA|VT|WA|WI|WV|WY)\s*\d{5}(?:-\d{4})?$",
        ""
    )
    text := Trim(RegExReplace(text, "\s+", " "), " ,")

    ; re-split street/unit after stripping tail
    street := ""
    unit := ""
    SplitAddressAndUnit(text, &street, &unit)
    address1 := street
    if (aptSuite = "")
        aptSuite := unit
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
        for _, val in states {
            if (val = abbr)
                return abbr
        }
    }

    return states.Has(clean) ? states[clean] : ""
}

NormalizeZip(zip) {
    text := Trim(zip)
    if (text = "")
        return ""

    ; Accept only standalone ZIP or ZIP+4 style text
    if RegExMatch(text, "^\d{5}(?:-\d{4})?$", &m)
        return SubStr(m[0], 1, 5)

    return ""
}

IsPhoneToken(token) {
    return NormalizePhone(token) != ""
}

NormalizePhone(phone) {
    text := Trim(phone)
    if (text = "")
        return ""

    ; Only treat it as phone if the token looks phone-like
    if !RegExMatch(text, "^\+?1?[\s\-\(\)\.]*\d{3}[\s\-\)\.]*\d{3}[\s\-\.]*\d{4}$")
        return ""

    digits := RegExReplace(text, "\D")
    if (StrLen(digits) = 11 && SubStr(digits, 1, 1) = "1")
        digits := SubStr(digits, 2)

    return (StrLen(digits) = 10) ? digits : ""
}

IsEmailToken(token) {
    return InStr(token, "@") && RegExMatch(token, "\.")
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

    ; Prefer full dates over default-day month/year dates.
    ; Since your default day is 16, if current uses 16 and new does not, prefer new.
    if RegExMatch(currentDob, "^\d{2}/16/\d{4}$") && !RegExMatch(newDob, "^\d{2}/16/\d{4}$")
        return true

    return false
}

NormalizeDOB(dob) {
    text := Trim(dob)
    if (text = "")
        return ""

    work := StrLower(text)
    work := RegExReplace(work, "\s+", " ")
    work := Trim(work)

    ; remove parentheses and commas
    work := RegExReplace(work, "[\(\)]", " ")
    work := RegExReplace(work, ",", " ")
    work := RegExReplace(work, "\s+", " ")
    work := Trim(work)

    ; remove leading age patterns
    work := RegExReplace(work, "i)^\s*age\s*\d{1,3}\s*", "")
    work := RegExReplace(work, "i)^\s*\d{1,3}\s*(?:años|anos)\s*", "")
    work := RegExReplace(work, "i)^\s*\d{1,3}\s*,\s*", "")

    ; remove filler phrases
    work := RegExReplace(work, "i)\bnacida?\s+en\b", " ")
    work := RegExReplace(work, "i)\bnacido\s+en\b", " ")
    work := RegExReplace(work, "i)\bconfirm\b", " ")

    ; HARD CUT: remove slash + trailing age text completely
    ; examples:
    ; "mayo de 1965/ 60 AÑOS"
    ; "3 junio de 1979 / 46 AÑOS"
    work := RegExReplace(work, "i)\s*/\s*\d{1,3}.*$", "")

    ; remove trailing age if no slash
    work := RegExReplace(work, "i)\s+\d{1,3}\s*(?:años|anos)$", "")

    ; normalize Spanish months + remove "de"
    work := NormalizeMonthWords(work)

    work := RegExReplace(work, "\s+", " ")
    work := Trim(work)

    ; dd-mon-yy or dd-mon-yyyy
    if RegExMatch(work, "i)^(\d{1,2})[-\s](jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[-\s](\d{2,4})$", &m) {
        month := MonthNumber(m[2])
        if month
            return FormatDateString(month, Integer(m[1]), NormalizeYear(m[3]))
    }

    ; dd month yyyy
    if RegExMatch(work, "i)^(\d{1,2})\s+(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec|january|february|march|april|june|july|august|september|october|november|december)\s+(\d{4})$", &m) {
        month := MonthNumber(m[2])
        if month
            return FormatDateString(month, Integer(m[1]), Integer(m[3]))
    }

    ; month dd yyyy
    if RegExMatch(work, "i)^(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec|january|february|march|april|june|july|august|september|october|november|december)\s+(\d{1,2})\s+(\d{4})$", &m) {
        month := MonthNumber(m[1])
        if month
            return FormatDateString(month, Integer(m[2]), Integer(m[3]))
    }

    ; month yyyy -> default day 16
    if RegExMatch(work, "i)^(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec|january|february|march|april|june|july|august|september|october|november|december)\s+(\d{4})$", &m) {
        month := MonthNumber(m[1])
        if month
            return FormatDateString(month, 16, Integer(m[2]))
    }

    ; mm/dd/yyyy
    if RegExMatch(work, "^(\d{1,2})/(\d{1,2})/(\d{2,4})$", &m)
        return FormatDateString(Integer(m[1]), Integer(m[2]), NormalizeYear(m[3]))

    ; yyyy-mm-dd
    if RegExMatch(work, "^(\d{4})-(\d{1,2})-(\d{1,2})$", &m)
        return FormatDateString(Integer(m[2]), Integer(m[3]), Integer(m[1]))

    ; mm/yyyy
    if RegExMatch(work, "^(\d{1,2})/(\d{4})$", &m)
        return FormatDateString(Integer(m[1]), 16, Integer(m[2]))

    return ""
}

NormalizeMonthWords(text) {
    static monthMap := Map(
        "enero", "january",
        "febrero", "february",
        "marzo", "march",
        "abril", "april",
        "mayo", "may",
        "junio", "june",
        "julio", "july",
        "agosto", "august",
        "septiembre", "september",
        "setiembre", "september",
        "octubre", "october",
        "noviembre", "november",
        "diciembre", "december"
    )

    for spanish, english in monthMap
        text := RegExReplace(text, "i)\b" spanish "\b", english)

    ; remove leftover filler words after month replacement
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

JsQuote(str) {
    text := StrReplace(str, "\", "\\")
    text := StrReplace(text, "'", "\'")
    text := StrReplace(text, "`r", "")
    text := StrReplace(text, "`n", "\n")
    return "'" . text . "'"
}

BuildProspectBookmarklet(fields) {
    selectors := GetProspectSelectorMap()
    dataParts := []

    for fieldName, selector in selectors {
        value := fields.Has(fieldName) ? fields[fieldName] : ""
        dataParts.Push(JsQuote(selector) . ":" . JsQuote(value))
    }

    script := []
    script.Push("(function(){")
    script.Push("const roots=[document];")
    script.Push("for(const frame of Array.from(document.querySelectorAll('iframe'))){try{if(frame.contentDocument)roots.push(frame.contentDocument);}catch(e){}}")
    script.Push("const pool=roots.flatMap((doc)=>Array.from(doc.querySelectorAll('input,textarea,select')));")
    script.Push("const find=(key)=>{for(const doc of roots){const byId=doc.getElementById(key);if(byId)return byId;}return pool.find((el)=>el.id===key||el.name===key||el.getAttribute('formcontrolname')===key||el.getAttribute('data-form-id')===key||el.getAttribute('data-testid')===key||el.getAttribute('aria-label')===key);};")
    script.Push("const fire=(el,type)=>el.dispatchEvent(new Event(type,{bubbles:true}));")
    script.Push("const setValue=(el,val)=>{if(!el)return false;el.focus();const tag=(el.tagName||'').toLowerCase();if(tag==='select'){const match=Array.from(el.options||[]).find((opt)=>opt.value===val||opt.text.trim().toLowerCase()===String(val).trim().toLowerCase());el.value=match?match.value:val;}else{el.value=val;}fire(el,'input');fire(el,'change');fire(el,'blur');return true;};")
    script.Push("const data={" . JoinArray(dataParts, ",") . "};")
    script.Push("const missing=[];for(const [key,val] of Object.entries(data)){if(!setValue(find(key),val == null ? '' : String(val)))missing.push(key);}")
    script.Push("if(missing.length){alert('Prospect paste complete. Missing selectors:\\n'+missing.join('\\n'));}else{console.log('Prospect paste complete.');}")
    script.Push("})();")

    return "javascript:" . JoinArray(script)
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

FastType(value) {
    value := value ?? ""
    if (value = "")
        return true
    SendText value
    return true
}

SelectDropdownValue(value) {
    value := Trim(value)
    if (value = "") {
        Send "{Tab}"
        Sleep 90
        return
    }

    ; For standard browser selects, typing the value is usually enough.
    SendText value
    Sleep 120
    Send "{Tab}"
    Sleep 100
}

FillNewProspectForm(fields) {
    ; Assumes cursor is already inside First Name.

    FastType(fields["FIRST_NAME"])
    Sleep 30
    Send "{Tab}"
    Sleep 30

    ; MI blank
    Send "{Tab}"
    Sleep 30

    FastType(fields["LAST_NAME"])
    Sleep 30
    Send "{Tab}"
    Sleep 30

    ; Suffix blank
    Send "{Tab}"
    Sleep 30

    FastType(fields["DOB"])
    Sleep 40
    Send "{Tab}"
    Sleep 60

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

; =================== EDGE PROSPECT FILLER ===================
; Copy either a raw CRM lead row or the FORMMAP block, then press Ctrl+Alt+9
; while the target prospect page is open in Microsoft Edge.
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

    ToolTip("Click FIRST NAME now (2s)")
    Sleep 2000
    ToolTip()

    FillNewProspectForm(fields)
}
^!0:: {
    raw := Trim(A_Clipboard)
    if (raw = "") {
        MsgBox("Clipboard empty.")
        return
    }
    fields := NormalizeProspectInput(raw)
    bookmarklet := BuildProspectBookmarklet(fields)
    MsgBox(SubStr(bookmarklet, 1, 300))
}
^!m:: {
    raw := Trim(A_Clipboard)
    result := NormalizeDOB(raw)
    MsgBox("RAW:`n" raw "`n`nDOB:`n" result)
}
^!p:: {
    raw := Trim(A_Clipboard)
    fields := NormalizeProspectInput(raw)
    msg := ""
    for k, v in fields
        msg .= k ": " v "`n"
    MsgBox(msg, "Parsed Prospect")
}
^!O:: {
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
