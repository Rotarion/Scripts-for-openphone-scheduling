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
    city := fields["CITY"]
    state := fields["STATE"]
    zip := fields["ZIP"]
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
    fields["CITY"] := NormalizeCity(city)
    fields["STATE"] := NormalizeState(state)
    fields["ZIP"] := NormalizeZip(zip)
    fields["APT_SUITE"] := aptSuite
    fields["PHONE"] := NormalizePhone(fields["PHONE"])
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
