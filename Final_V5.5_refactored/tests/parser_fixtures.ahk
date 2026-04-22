#Requires AutoHotkey v2.0

global dobDefaultDay := 16
global batchMinVehicles := 0
global batchMaxVehicles := 99
global tagSymbol := "+"

#Include ..\domain\lead_normalizer.ahk
#Include ..\domain\lead_parser.ahk
#Include ..\domain\batch_rules.ahk

labeledRaw :=
(
Name:
Maria Gomez
Date of Birth:
Jan 1985
Gender:
Female
Address Line 1:
123 Main St Apt 4B
City:
Miami
State:
FL
Zip Code:
33101
Phone:
(305) 555-1212
Email:
maria@example.com
)

fields := ParseLabeledLeadToProspect(labeledRaw)
AssertEqual(fields["FIRST_NAME"], "Maria", "First name should come from labeled lead")
AssertEqual(fields["LAST_NAME"], "Gomez", "Last name should come from labeled lead")
AssertEqual(fields["DOB"], "01/16/1985", "Month-only DOB should use default day")
AssertEqual(fields["PHONE"], "3055551212", "Phone should normalize to digits")
AssertEqual(fields["APT_SUITE"], "4B", "Apartment should split out of address")

batchRaw := "PERSONAL LEAD - JOHN SMITH 12/01/2026 10:00:00 AM 123 Main St Miami FL 33101 (561) 555-1212 john@example.com Jan 1980 Male 2020 Toyota Camry"
lead := BuildBatchLeadRecord(batchRaw)
AssertEqual(lead["FULL_NAME"], "John Smith", "Batch lead name should be normalized")
AssertEqual(lead["PHONE"], "5615551212", "Batch phone should be parsed")
AssertEqual(lead["VEHICLE_COUNT"], 1, "Vehicle count should detect one car")

multiBatch :=
(
PERSONAL LEAD - JOHN SMITH
PERSONAL LEAD - JANE DOE
)

rows := ParseBatchLeadRows(multiBatch)
AssertEqual(rows.Length, 2, "Batch rows should split on PERSONAL LEAD markers")

MsgBox("parser_fixtures passed")

AssertEqual(actual, expected, message) {
    if (actual != expected)
        throw Error(message . "`nExpected: " expected "`nActual: " actual)
}
