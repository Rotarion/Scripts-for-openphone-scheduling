^!u:: {
    if !ClipWait(1) {
        MsgBox("Clipboard empty. Copy one lead row first.")
        return
    }

    raw := Trim(A_Clipboard)
    if (raw = "") {
        MsgBox("Clipboard empty. Copy one lead row first.")
        return
    }

    lead := BuildBatchLeadRecord(raw)
    if (lead["PHONE"] = "" || lead["FULL_NAME"] = "") {
        MsgBox("Could not parse phone/name from clipboard.")
        return
    }

    BeginAutomationRun()
    result := RunQuickLeadCreateAndTag(lead)
    if StopRequested()
        return
    if (result != "OK")
        MsgBox(result)
}

^!1:: {
    global priceOldCar, priceOneCar, priceTwoCars, priceThreeCars, priceFourCars, priceFiveCars

    if !ClipWait(1) {
        MsgBox("Clipboard empty. Copy the lead's name first.")
        return
    }

    leadName := CleanName(A_Clipboard)
    if (leadName = "" || StrLen(leadName) > 30) {
        MsgBox("Please copy just the lead's name (<= 30 chars).")
        return
    }

    promptText :=
        "Escribe 0, 1, 2, 3, 4 o 5:`n"
        . "0 = auto muy antiguo (" FormatMonthlyPrice(priceOldCar) ", sin 'FULL COVERAGE')`n"
        . "1 = su carro (" FormatMonthlyPrice(priceOneCar) ", FULL COVERAGE)`n"
        . "2 = sus carros (" FormatMonthlyPrice(priceTwoCars) ", FULL COVERAGE)`n"
        . "3 = sus carros (" FormatMonthlyPrice(priceThreeCars) ", FULL COVERAGE)`n"
        . "4 = sus carros (" FormatMonthlyPrice(priceFourCars) ", FULL COVERAGE)`n"
        . "5 = sus carros (" FormatMonthlyPrice(priceFiveCars) ", FULL COVERAGE, 100/300)"

    ib := InputBox(promptText, "Número de vehículos", "w420 h280")
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

^!b::RunBatchFromClipboard("stable")
^!n::RunBatchFromClipboard("fast")
