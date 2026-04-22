TemplateRead(section, key, defaultValue := "") {
    global templatesFile
    value := IniRead(templatesFile, section, key, defaultValue)
    return DecodeTemplateValue(value)
}

DecodeTemplateValue(value) {
    text := value ?? ""
    text := StrReplace(text, "\n", "`n")
    return text
}

ExpandTemplate(text, tokens) {
    output := text
    for key, value in tokens
        output := StrReplace(output, "{" key "}", value)
    return output
}

BuildMessage(leadName, carCount, vehicles := "", useBatchSingleCarPricing := false) {
    global agentName, agentEmail

    greeting := (A_Hour < 12)
        ? TemplateRead("QuoteMessage", "GreetingMorning", "Buenos días")
        : TemplateRead("QuoteMessage", "GreetingAfternoon", "Buenas tardes")

    leadName := ProperCase(leadName)
    firstName := ExtractFirstName(leadName)

    vehLine := (carCount >= 2)
        ? TemplateRead("QuoteMessage", "VehicleLineMultiple", "Hicimos la cotización para el seguro de sus carros.")
        : TemplateRead("QuoteMessage", "VehicleLineSingle", "Hicimos la cotización para el seguro de su carro.")

    vehBlock := vehLine . "`n"
    if (IsObject(vehicles) && vehicles.Length > 0) {
        for _, v in vehicles
            vehBlock .= v . "`n"
    }

    coverageLine := (carCount = 5)
        ? TemplateRead("QuoteMessage", "CoverageLineFive", "Bodily Injury $100k per person $300k per ocurrence\nProperty Damage $100k per ocurrence")
        : ""

    price := ResolveQuotePrice(carCount, vehicles, useBatchSingleCarPricing)
    priceLine := ExpandTemplate(
        TemplateRead("QuoteMessage", "PriceLine", "Actualmente tenemos opciones con ALLSTATE desde {PRICE} al mes."),
        Map("PRICE", price)
    )

    parts := [
        greeting . " " . firstName . ",",
        vehBlock . (coverageLine != "" ? "`n`n" . coverageLine : ""),
        priceLine,
        TemplateRead("QuoteMessage", "SavingsLine", "Muchos clientes en su misma situación han logrado ahorrar cambiándose con nosotros."),
        TemplateRead("QuoteMessage", "CallInviteLine", "Si quiere, en una llamada rápida de 2-3 minutos podemos revisar si realmente le conviene o no."),
        TemplateRead("QuoteMessage", "CloseQuestionLine", "¿Le parece bien si lo revisamos juntos?"),
        TemplateRead("QuoteMessage", "PhoneLine", "📞 (561) 220-7073"),
        agentName,
        TemplateRead("QuoteMessage", "SignatureTitle", "Agente de Seguros - Allstate"),
        TemplateRead("QuoteMessage", "DirectLine", "Direct Line: (561) 220-7073"),
        TemplateRead("QuoteMessage", "OfficeLine", "Office Line: (754) 236-8009"),
        agentEmail,
        TemplateRead("QuoteMessage", "UnsubscribeLine", "Reply STOP to unsubscribe")
    ]

    return JoinArray(parts, "`n`n")
}

BuildFollowupQueue(leadName, offset) {
    global agentName, configDays, holidays

    leadName := ExtractFirstName(ProperCase(leadName))

    dA := configDays[1]
    dB := configDays[2]
    dC := configDays[3]
    dD := configDays[4]

    dADate := BusinessDateForDay(dA, holidays)
    dBDate := BusinessDateForDay(dB, holidays)
    dCDate := BusinessDateForDay(dC, holidays)
    dDDate := BusinessDateForDay(dD, holidays)

    t1_1 := TimeWithOffset(9, 30, 30, offset)
    t1_2 := TimeWithOffset(9, 31, 10, offset)
    t1_3 := TimeWithOffset(9, 31, 30, offset)
    t1_4 := TimeWithOffset(10, 45, 0, offset)
    t2_1 := TimeWithOffset(16, 0, 10, offset)
    t2_2 := TimeWithOffset(16, 1, 10, offset)
    t4_1 := TimeWithOffset(16, 30, 0, offset)
    t4_2 := TimeWithOffset(16, 31, 0, offset)
    t5_1 := TimeWithOffset(12, 0, 0, offset)
    t5_2 := TimeWithOffset(12, 1, 0, offset)

    tokens := Map("FIRST_NAME", leadName, "AGENT_NAME", agentName)
    msgs := [
        Map("day", dA, "seq", 1, "text", ExpandTemplate(TemplateRead("Followups", "A1", "Buenos días, {FIRST_NAME}."), tokens), "date", dADate, "time", t1_1),
        Map("day", dA, "seq", 2, "text", ExpandTemplate(TemplateRead("Followups", "A2", "Soy {AGENT_NAME} de Allstate. Ya le preparé la cotización de su auto."), tokens), "date", dADate, "time", t1_2),
        Map("day", dA, "seq", 3, "text", ExpandTemplate(TemplateRead("Followups", "A3", "En muchos casos logramos bajar el pago mensual sin quitar coberturas. Si gusta, se la resumo en 2 minutos por aquí."), tokens), "date", dADate, "time", t1_3),
        Map("day", dA, "seq", 4, "text", ExpandTemplate(TemplateRead("Followups", "A4", "Si me responde “Sí”, se la envío ahora mismo."), tokens), "date", dADate, "time", t1_4),

        Map("day", dB, "seq", 1, "text", ExpandTemplate(TemplateRead("Followups", "B1", "Hola, {FIRST_NAME}."), tokens), "date", dBDate, "time", t2_1),
        Map("day", dB, "seq", 2, "text", ExpandTemplate(TemplateRead("Followups", "B2", "Hoy intenté comunicarme con usted porque todavía puedo revisar si califica a descuentos disponibles. Si me responde “Revisar”, yo me encargo de validar todo por usted."), tokens), "date", dBDate, "time", t2_2),

        Map("day", dC, "seq", 1, "text", ExpandTemplate(TemplateRead("Followups", "C1", "Buenas tardes, {FIRST_NAME}."), tokens), "date", dCDate, "time", t4_1),
        Map("day", dC, "seq", 2, "text", ExpandTemplate(TemplateRead("Followups", "C2", "Esta semana hemos ayudado a varios conductores a comparar su póliza actual con Allstate y en muchos casos encontraron una mejor opción. Si me responde “Comparar”, reviso su caso y le digo honestamente si le conviene o no. Reply STOP to unsubscribe"), tokens), "date", dCDate, "time", t4_2),

        Map("day", dD, "seq", 1, "text", ExpandTemplate(TemplateRead("Followups", "D1", "{FIRST_NAME}, sigo teniendo su cotización disponible, pero normalmente cierro los pendientes cuando no recibo respuesta."), tokens), "date", dDDate, "time", t5_1),
        Map("day", dD, "seq", 2, "text", ExpandTemplate(TemplateRead("Followups", "D2", "Si todavía quiere revisarla, respóndame “Continuar” y le envío el resumen por aquí."), tokens), "date", dDDate, "time", t5_2)
    ]

    return msgs
}
