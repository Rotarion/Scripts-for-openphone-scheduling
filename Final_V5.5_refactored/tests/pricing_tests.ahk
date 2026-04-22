#Requires AutoHotkey v2.0

global priceOldCar := 98
global priceOneCar := 117
global priceOneCar2020Plus := 167
global priceTwoCars := 176
global priceThreeCars := 284
global priceFourCars := 397
global priceFiveCars := 397
global singleCarModernYearCutoff := 2017

#Include ..\domain\pricing_rules.ahk

AssertEqual(ResolveQuotePrice(0), "$98", "Old car price should stay intact")
AssertEqual(ResolveQuotePrice(1, ["2020 Toyota Camry"], true), "$167", "Modern single-car batch price should use 2020+ tier")
AssertEqual(ResolveQuotePrice(1, ["2010 Toyota Camry"], true), "$117", "Older single-car batch price should use baseline tier")
AssertEqual(ResolveQuotePrice(3), "$284", "Three-car price should stay intact")

MsgBox("pricing_tests passed")

AssertEqual(actual, expected, message) {
    if (actual != expected)
        throw Error(message . "`nExpected: " expected "`nActual: " actual)
}
