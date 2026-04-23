ResolveQuotePrice(carCount, vehicles := "", useBatchSingleCarPricing := false) {
    global priceOldCar, priceOneCar, priceOneCar2020Plus
    global priceTwoCars, priceThreeCars, priceFourCars, priceFiveCars
    global singleCarModernYearCutoff

    if (useBatchSingleCarPricing && carCount = 1) {
        vehicleYear := ExtractVehicleYearFromList(vehicles)
        if (vehicleYear >= singleCarModernYearCutoff)
            return FormatMonthlyPrice(priceOneCar2020Plus)
        return FormatMonthlyPrice(priceOneCar)
    }

    prices := Map(0, priceOldCar, 1, priceOneCar, 2, priceTwoCars, 3, priceThreeCars, 4, priceFourCars, 5, priceFiveCars)
    return prices.Has(carCount) ? FormatMonthlyPrice(prices[carCount]) : "$127"
}

ExtractVehicleYearFromList(vehicles) {
    if !(IsObject(vehicles) && vehicles.Length >= 1)
        return 0
    return ExtractVehicleYear(vehicles[1])
}

ExtractVehicleYear(vehicleText) {
    if RegExMatch(vehicleText, "i)\b((19|20)\d{2})\b", &m)
        return Integer(m[1])
    return 0
}

FormatMonthlyPrice(amount) {
    return "$" . Integer(amount)
}
