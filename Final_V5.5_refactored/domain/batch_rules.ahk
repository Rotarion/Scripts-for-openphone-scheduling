ParseBatchLeadRows(raw) {
    rows := []
    clean := StrReplace(raw, "`r", "")
    split := RegExReplace(clean, "i)(?=(?:DUPLICATED\s+(?:OPPORTUNITY\s+)?)?PERSONAL\s+LEAD)", "`n")

    for _, chunk in StrSplit(split, "`n") {
        chunk := Trim(chunk)
        if (chunk = "")
            continue
        if RegExMatch(chunk, "i)PERSONAL\s+LEAD")
            rows.Push(chunk)
    }
    return rows
}

ExtractLikelyLeadNameField(raw) {
    raw := StripGridActionText(raw)
    raw := Trim(StrReplace(StrReplace(raw, "`r", ""), "`n", " "))
    if (raw = "")
        return ""

    if RegExMatch(raw, "i)^\s*((?:DUPLICATED\s+)?(?:OPPORTUNITY\s+)?PERSONAL\s+LEAD\s*-\s*.*?)(?:\t|\s{2,}(?=\S)|$)", &m)
        return Trim(m[1])

    if RegExMatch(raw, "i)((?:DUPLICATED\s+)?(?:OPPORTUNITY\s+)?PERSONAL\s+LEAD\s*-\s*.*?)(?:\t|\s{2,}(?=\S)|$)", &m)
        return Trim(m[1])

    return raw
}

ExtractBatchName(raw) {
    prefixPattern := "i)^\s*(?:DUPLICATED\s+)?(?:OPPORTUNITY\s+)?PERSONAL\s+LEAD\s*-\s*"

    if IsLabeledLeadFormat(raw) {
        data := ParseLabeledLeadRaw(raw)

        if data.Has("Name") {
            name := CleanBatchNameCandidate(RegExReplace(data["Name"], prefixPattern, ""))
            if (name != "")
                return name
        }

        if data.Has("Contact") {
            name := CleanBatchNameCandidate(RegExReplace(data["Contact"], prefixPattern, ""))
            if (name != "")
                return name
        }

        first := data.Has("First Name") ? Trim(data["First Name"]) : ""
        last := data.Has("Last Name") ? Trim(data["Last Name"]) : ""
        name := CleanBatchNameCandidate(first . " " . last)
        if (name != "")
            return name
    }

    raw := ExtractLikelyLeadNameField(raw)
    raw := StripGridActionText(raw)
    raw := Trim(raw)

    if RegExMatch(raw, prefixPattern . "([^\t\r\n]+)", &m)
        return CleanBatchNameCandidate(m[1])

    return CleanBatchNameCandidate(raw)
}

CleanBatchNameCandidate(name) {
    name := Trim(StrReplace(StrReplace(name, "`r", ""), "`n", " "))
    if (name = "")
        return ""

    name := RegExReplace(name, "[\t\r\n].*$", "")
    name := RegExReplace(name, "i)\s{2,}(?=(?:new\s+(?:webform\s+folder|skyline\s+leads)\s*-\s*personal|(?:[A-Za-z]+\s+){0,3}(?:folder|leads?|source|status)\b(?:\s*-\s*[A-Za-z]+)?|\d+\s*-\s*(?:new|open|working|quoted|pending|closed|sold|contacted)|move\s+to\s+recycle\s+bin|recycle\s+bin|\d{1,2}/\d{1,2}/\d{2,4}\b)).*$", "")
    name := RegExReplace(name, "i)\s+(?:new\s+webform\s+folder\s*-\s*personal|new\s+skyline\s+leads\s*-\s*personal|move\s+to\s+recycle\s+bin|recycle\s+bin)\b.*$", "")
    name := RegExReplace(name, "i)\s+\d+\s*-\s*(?:new|open|working|quoted|pending|closed|sold|contacted)\b.*$", "")
    name := RegExReplace(name, "i)\s+(?:[A-Za-z]+\s+){0,3}(?:folder|leads?|source|status)\b(?:\s*-\s*[A-Za-z]+)?(?:\s+.*)?$", "")
    name := RegExReplace(name, "i)\s+\d{1,2}/\d{1,2}/\d{2,4}\s+\d{1,2}:\d{2}(?::\d{2})?\s*(?:AM|PM)?\b.*$", "")
    name := RegExReplace(name, "i)^\s*(?:DUPLICATED\s+)?(?:OPPORTUNITY\s+)?PERSONAL\s+LEAD\s*-\s*", "")
    name := RegExReplace(name, "\s{2,}", " ")
    return Trim(name, " -:`t`r`n")
}

ExtractBatchPhone(raw) {
    if RegExMatch(raw, "\((\d{3})\)\s*(\d{3})-(\d{4})", &m)
        return m[1] . m[2] . m[3]
    return ""
}

ExtractVehicleList(rawLead) {
    vehicles := []

    rawLead := StripGridActionText(rawLead)

    vehicleText := ""
    if RegExMatch(rawLead, "i)(?:Male|Female)\s*(.*)", &m)
        vehicleText := Trim(m[1])

    if (vehicleText = "")
        return vehicles

    vehicleText := StripGridActionText(vehicleText)
    split := RegExReplace(vehicleText, "i)((19|20)\d{2})(\s*/?[\s]*[A-Za-z])", "`n$1$3")

    for _, part in StrSplit(split, "`n") {
        part := SanitizeVehicleLine(part)
        if (part = "")
            continue
        if RegExMatch(part, "i)^(19|20)\d{2}\s*/?[\s]*[A-Za-z]")
            vehicles.Push(ProperCasePhrase(part))
    }

    return vehicles
}

BuildBatchLeadRecord(rawLead) {
    global tagSymbol

    rawLead := StripGridActionText(rawLead)
    nameSource := ""

    if InStr(rawLead, "`t") {
        cols := StrSplit(rawLead, "`t")
        if (cols.Length >= 1) {
            firstCol := Trim(cols[1])
            if (firstCol != "")
                nameSource := firstCol
        }
    }

    if (nameSource = "")
        nameSource := ExtractLikelyLeadNameField(rawLead)

    batchName := ExtractBatchName(nameSource)
    batchPhone := ExtractBatchPhone(rawLead)
    vehicles := ExtractVehicleList(rawLead)

    fullName := ProperCase(batchName)
    firstName := ExtractFirstName(fullName)

    parts := StrSplit(Trim(fullName), " ")
    lastName := (parts.Length >= 2) ? parts[parts.Length] : ""

    tagValue := Trim(tagSymbol)
    if (tagValue = "")
        tagValue := "+"

    holderName := tagValue . " " . fullName

    return Map(
        "RAW", rawLead,
        "FIELDS", Map("FIRST_NAME", firstName, "LAST_NAME", lastName, "PHONE", batchPhone),
        "FULL_NAME", fullName,
        "PHONE", batchPhone,
        "VEHICLES", vehicles,
        "VEHICLE_COUNT", vehicles.Length,
        "HOLDER_NAME", holderName,
        "TAG_VALUE", tagValue
    )
}

BuildBatchLeadHolder(raw) {
    global batchMinVehicles, batchMaxVehicles

    holder := []
    rows := ParseBatchLeadRows(raw)

    for _, row in rows {
        lead := BuildBatchLeadRecord(row)
        vc := lead["VEHICLE_COUNT"]

        if (lead["FULL_NAME"] = "" || lead["PHONE"] = "")
            continue
        if (vc < batchMinVehicles || vc > batchMaxVehicles)
            continue

        holder.Push(lead)
    }
    return holder
}

SanitizeVehicleLine(text) {
    text := StripGridActionText(text)
    text := RegExReplace(text, "\s*/\s*", " ")
    text := RegExReplace(text, "\s+", " ")
    text := RegExReplace(text, "[\t ]+$", "")
    return Trim(text)
}

StripGridActionText(text) {
    text := Trim(text)
    text := RegExReplace(text, "i)\bMove\s+To\s+Recycle\s+Bin\b", "")
    text := RegExReplace(text, "i)\bRecycle\s+Bin\b", "")
    text := RegExReplace(text, "\s{2,}", " ")
    return Trim(text, " `t-")
}
