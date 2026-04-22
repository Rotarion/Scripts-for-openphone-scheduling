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

    if IsLikelyBatchGridInput(raw)
        return ParseBatchGridRow(ExtractFirstBatchGridRow(raw))

    if RegExMatch(raw, "i)PERSONAL\s+LEAD")
        return ParseBatchCRMToProspect(raw)

    return NormalizeRawLeadToProspect(raw)
}

IsLikelyBatchGridInput(raw) {
    firstRow := ExtractFirstBatchGridRow(raw)
    if (firstRow = "")
        return false

    cols := StrSplit(firstRow, "`t")
    return cols.Length >= 11
        && RegExMatch(Trim(cols[1]), "i)PERSONAL\s+LEAD")
        && IsTimestampToken(cols[4])
        && NormalizeZip(cols[8]) != ""
        && NormalizePhone(cols[9]) != ""
}

ExtractFirstBatchGridRow(raw) {
    clean := StrReplace(raw, "`r", "")
    for _, line in StrSplit(clean, "`n") {
        line := Trim(line)
        if (line = "")
            continue
        if InStr(line, "`t") && RegExMatch(line, "i)PERSONAL\s+LEAD")
            return line
    }
    return ""
}

ParseBatchCRMToProspect(raw) {
    if IsLikelyBatchGridInput(raw)
        return ParseBatchGridRow(ExtractFirstBatchGridRow(raw))

    fields := NewProspectFields()

    rows := ParseBatchLeadRows(raw)
    lead := (rows.Length > 0) ? rows[1] : raw

    batchName := ExtractBatchName(lead)
    if (batchName != "")
        ApplyLeadName(fields, batchName)

    fields["PHONE"] := ExtractBatchPhone(lead)
    fields["EMAIL"] := ExtractFirstEmail(lead)

    if RegExMatch(lead, "i)\b(Male|Female)\b", &mg)
        fields["GENDER"] := NormalizeGender(mg[1])

    dobText := ""
    if RegExMatch(lead, "i)\(\d{3}\)\s*\d{3}-\d{4}\s*(?:[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}\s*)?(.+?)(?=\s*(?:Male|Female)\b)", &md)
        dobText := Trim(md[1])
    if (dobText != "")
        fields["DOB"] := NormalizeDOB(dobText)

    locationBlock := ""
    if RegExMatch(lead, "i)\d{1,2}/\d{1,2}/\d{4}\s+\d{1,2}:\d{2}:\d{2}\s*(?:AM|PM)\s*(.+?)(?=\(\d{3}\))", &ml)
        locationBlock := Trim(ml[1])

    if (locationBlock != "")
        ParseBatchLocationBlock(locationBlock, fields)

    NormalizeAddressMap(fields)
    fields["DOB"] := NormalizeDOB(fields["DOB"])
    fields["GENDER"] := NormalizeGender(fields["GENDER"])
    fields["STATE"] := NormalizeState(fields["STATE"])
    fields["ZIP"] := NormalizeZip(fields["ZIP"])
    fields["PHONE"] := NormalizePhone(fields["PHONE"])

    return fields
}

ParseBatchLocationBlock(locBlock, fields) {
    stateFullNames := "Alabama|Alaska|Arizona|Arkansas|California|Colorado|Connecticut|Delaware|Florida|Georgia|Hawaii|Idaho|Illinois|Indiana|Iowa|Kansas|Kentucky|Louisiana|Maine|Maryland|Massachusetts|Michigan|Minnesota|Mississippi|Missouri|Montana|Nebraska|Nevada|New Hampshire|New Jersey|New Mexico|New York|North Carolina|North Dakota|Ohio|Oklahoma|Oregon|Pennsylvania|Rhode Island|South Carolina|South Dakota|Tennessee|Texas|Utah|Vermont|Virginia|Washington|West Virginia|Wisconsin|Wyoming"
    pattern := "i)(" . stateFullNames . ")\s*(\d{5})"

    pos := 1
    lastState := ""
    lastZip := ""
    lastPos := 0

    while RegExMatch(locBlock, pattern, &ms, pos) {
        lastState := ms[1]
        lastZip := ms[2]
        lastPos := ms.Pos
        pos := ms.Pos + ms.Len
    }

    if (lastPos = 0)
        return

    fields["STATE"] := NormalizeState(lastState)
    fields["ZIP"] := lastZip

    beforeState := Trim(SubStr(locBlock, 1, lastPos - 1))
    beforeState := RegExReplace(beforeState, "[,\s]+$", "")

    stateAbbrs := "AL|AK|AZ|AR|CA|CO|CT|DC|DE|FL|GA|HI|IA|ID|IL|IN|KS|KY|LA|MA|MD|ME|MI|MN|MO|MS|MT|NC|ND|NE|NH|NJ|NM|NV|NY|OH|OK|OR|PA|RI|SC|SD|TN|TX|UT|VA|VT|WA|WI|WV|WY"

    if RegExMatch(beforeState, "i)^(.*?)\s*,\s*(" . stateAbbrs . ")\s*(?:\d{5}(?:-\d{4})?)?\s*(.*)$", &mc) {
        addrPart := Trim(mc[1])
        cityPart := Trim(mc[3])

        if (cityPart = "") {
            streetOnly := ""
            cityFromAddr := ""
            SplitStreetAndTrailingCity(addrPart, &streetOnly, &cityFromAddr)
            if (cityFromAddr != "") {
                addrPart := streetOnly
                cityPart := cityFromAddr
            }
        }

        if (cityPart != "") {
            cityPart := RegExReplace(cityPart, "^[,\s]+|[,\s]+$", "")
            fields["CITY"] := ProperCasePhrase(cityPart)

            if (fields["CITY"] != "" && RegExMatch(addrPart, "i)\s+" . cityPart . "$"))
                addrPart := Trim(RegExReplace(addrPart, "i)\s+" . cityPart . "$", ""))
        }

        SetAddressFields(fields, addrPart)
    } else {
        streetOnly := ""
        cityFromAddr := ""
        SplitStreetAndTrailingCity(beforeState, &streetOnly, &cityFromAddr)

        if (cityFromAddr != "")
            fields["CITY"] := ProperCasePhrase(cityFromAddr)

        SetAddressFields(fields, streetOnly != "" ? streetOnly : beforeState)
    }
}

ParseLabeledLeadToProspect(raw) {
    fields := NewProspectFields()
    lines := StrSplit(StrReplace(raw, "`r", ""), "`n")

    data := Map()
    currentLabel := ""
    currentValue := ""
    FlushPair := (*) => (currentLabel != "" ? StoreLabeledField(data, currentLabel, currentValue) : 0)

    for _, line in lines {
        line := Trim(line)
        if (line = "")
            continue

        if RegExMatch(line, "^([^:]+:?):\s*$", &m) {
            FlushPair()
            currentLabel := Trim(m[1], " :")
            currentValue := ""
            continue
        }

        if RegExMatch(line, "^([^:]+:?):\s*(.+)$", &m) {
            FlushPair()
            currentLabel := Trim(m[1], " :")
            currentValue := Trim(m[2])
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

    if data.Has("First Name")
        fields["FIRST_NAME"] := ProperCasePhrase(data["First Name"])
    else if data.Has("Name")
        ApplyLeadName(fields, RegExReplace(data["Name"], "i)^PERSONAL LEAD\s*-\s*"))

    if data.Has("Last Name")
        fields["LAST_NAME"] := ProperCasePhrase(data["Last Name"])

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

    fields["EMAIL"] := ExtractFirstEmail(raw)

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
    FlushPair := (*) => (currentLabel != "" ? StoreLabeledField(data, currentLabel, currentValue) : 0)

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

    fields["EMAIL"] := ExtractFirstEmail(raw)

    zipIdx := 0
    for i, token in tokens {
        if IsTimestampToken(token) || IsPhoneToken(token) || IsEmailToken(token)
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
        if addrIdx
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

ParseBatchGridRow(raw) {
    fields := NewProspectFields()
    cols := StrSplit(raw, "`t")
    if (cols.Length < 11)
        return fields

    fields["RAW_NAME"] := Trim(cols[1])
    fields["ADDRESS_1"] := Trim(cols[5])
    fields["CITY"] := NormalizeCity(cols[6])
    fields["STATE"] := NormalizeState(cols[7])
    fields["ZIP"] := NormalizeZip(cols[8])
    fields["PHONE"] := NormalizePhone(cols[9])

    idx := 10
    if (cols.Length >= idx && IsEmailToken(cols[idx])) {
        fields["EMAIL"] := Trim(cols[idx])
        idx += 1
    }

    if (cols.Length >= idx)
        fields["DOB"] := NormalizeDOB(cols[idx])
    if (cols.Length >= idx + 1)
        fields["GENDER"] := NormalizeGender(cols[idx + 1])

    nameText := ExtractBatchName(cols[1])
    if (nameText != "")
        ApplyLeadName(fields, nameText)

    NormalizeAddressMap(fields)
    fields["DOB"] := NormalizeDOB(fields["DOB"])
    fields["GENDER"] := NormalizeGender(fields["GENDER"])
    fields["STATE"] := NormalizeState(fields["STATE"])
    fields["ZIP"] := NormalizeZip(fields["ZIP"])
    fields["PHONE"] := NormalizePhone(fields["PHONE"])
    return fields
}
