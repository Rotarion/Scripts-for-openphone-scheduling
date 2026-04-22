#Requires AutoHotkey v2.0

global templatesFile := A_ScriptDir "\..\config\templates.ini"
global agentName := "Pablo Cabrera"
global agentEmail := "pablocabrera@allstate.com"
global configDays := [2, 4, 7, 9]
global holidays := ["05/25/2026"]
global priceOldCar := 98
global priceOneCar := 117
global priceOneCar2020Plus := 167
global priceTwoCars := 176
global priceThreeCars := 284
global priceFourCars := 397
global priceFiveCars := 397
global singleCarModernYearCutoff := 2017

#Include ..\domain\lead_normalizer.ahk
#Include ..\domain\pricing_rules.ahk
#Include ..\domain\date_rules.ahk
#Include ..\domain\message_templates.ahk

message := BuildMessage("juan perez", 1, ["2020 Toyota Camry"], true)
AssertTrue(InStr(message, "Juan"), "Message should proper-case the lead name")
AssertTrue(InStr(message, "$167"), "Message should include the 2020+ single-car price")
AssertTrue(InStr(message, agentName), "Message should include the configured agent name")
AssertTrue(InStr(message, agentEmail), "Message should include the configured agent email")

queue := BuildFollowupQueue("juan perez", 0)
AssertEqual(queue.Length, 10, "Follow-up queue should still create ten scheduled messages")
AssertEqual(queue[1]["day"], 2, "First follow-up block should use configured day A")
AssertEqual(queue[10]["day"], 9, "Last follow-up block should use configured day D")

MsgBox("message_tests passed")

AssertEqual(actual, expected, message) {
    if (actual != expected)
        throw Error(message . "`nExpected: " expected "`nActual: " actual)
}

AssertTrue(condition, message) {
    if !condition
        throw Error(message)
}
