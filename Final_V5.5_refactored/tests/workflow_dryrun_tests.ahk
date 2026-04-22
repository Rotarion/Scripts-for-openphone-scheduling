#Requires AutoHotkey v2.0

global templatesFile := A_ScriptDir "\..\config\templates.ini"
global holidays := ["05/25/2026"]
global configDays := [2, 4, 7, 9]
global priceOldCar := 98
global priceOneCar := 117
global priceOneCar2020Plus := 167
global priceTwoCars := 176
global priceThreeCars := 284
global priceFourCars := 397
global priceFiveCars := 397
global singleCarModernYearCutoff := 2017
global tagSymbol := "+"

#Include ..\domain\lead_normalizer.ahk
#Include ..\domain\lead_parser.ahk
#Include ..\domain\pricing_rules.ahk
#Include ..\domain\date_rules.ahk
#Include ..\domain\message_templates.ahk
#Include ..\domain\batch_rules.ahk
#Include ..\workflows\batch_run.ahk

lead := BuildBatchLeadRecord("PERSONAL LEAD - JOHN SMITH 12/01/2026 10:00:00 AM 123 Main St Miami FL 33101 (561) 555-1212 john@example.com Jan 1980 Male 2020 Toyota Camry")
plan := TraceBatchLeadPlan(lead, "stable")

AssertEqual(plan["mode"], "stable", "Dry run should preserve requested mode")
AssertEqual(plan["followupCount"], 10, "Dry run should account for all follow-up messages")
AssertEqual(plan["steps"].Length, 8, "Dry run should trace the full batch flow")
AssertTrue(InStr(plan["steps"][8], "+"), "Dry run should include tag application in the final step")

MsgBox("workflow_dryrun_tests passed")

AssertEqual(actual, expected, message) {
    if (actual != expected)
        throw Error(message . "`nExpected: " expected "`nActual: " actual)
}

AssertTrue(condition, message) {
    if !condition
        throw Error(message)
}
