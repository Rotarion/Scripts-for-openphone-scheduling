#Requires AutoHotkey v2.0

#Include ..\domain\date_rules.ahk

holidays := ["05/25/2026"]
days := ParseDays("2,4,7,9")
AssertEqual(days.Length, 4, "Day parser should keep four schedule days")
AssertEqual(days[3], 7, "Day parser should preserve order")
AssertEqual(NextBusinessDateYYYYMMDD("20260522", holidays), "20260526", "Friday before holiday weekend should skip to Tuesday")
AssertEqual(AddBusinessDays("20260522", 2, holidays), "20260527", "Two business days should skip weekend and holiday")

MsgBox("date_tests passed")

AssertEqual(actual, expected, message) {
    if (actual != expected)
        throw Error(message . "`nExpected: " expected "`nActual: " actual)
}
