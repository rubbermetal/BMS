' VBScript for HVAC control system

Class ControlField
    Public value
End Class

Dim sleepDuration, controlField, statusField

' Time to sleep between setting changes
sleepDuration = 2000

' Function to calculate the vapor pressure of water
Function calculateVaporPressure(temperature, pressure)
    calculateVaporPressure = 0.61094 * Exp(17.625 * temperature / (temperature + 243.04)) * (pressure / 101.325)
End Function

' Function to calculate the saturation vapor pressure of water
Function calculateSaturationVaporPressure(temperature)
    calculateSaturationVaporPressure = 0.61094 * Exp(17.625 * temperature / (temperature + 243.04))
End Function

' Function to calculate the wet bulb temperature
Function calculateWetBulbTemperature(dryBulbTemp, relativeHumidity, pressure)
    calculateWetBulbTemperature = wetBulbTemperature(dryBulbTemp, relativeHumidity, pressure)
End Function

' Function to calculate wet bulb temperature using Newton-Raphson iteration
Function wetBulbTemperature(dryBulbTemp, relativeHumidity, pressure)
    Dim normalHumidityRatio, resultTemp, newHumidityRatio, newHumidityRatio2, dwBydTwb

    normalHumidityRatio = calculateHumidityRatio(dryBulbTemp, relativeHumidity, pressure)
    resultTemp = dryBulbTemp

    ' Iteratively solve for wet bulb temperature using Newton-Raphson method
    newHumidityRatio = calculateHumidityRatioFromWetBulb(dryBulbTemp, resultTemp, pressure)

    Do While Abs((newHumidityRatio - normalHumidityRatio) / normalHumidityRatio) > 0.00001
        newHumidityRatio2 = calculateHumidityRatioFromWetBulb(dryBulbTemp, resultTemp - 0.001, pressure)
        dwBydTwb = (newHumidityRatio - newHumidityRatio2) / 0.001
        resultTemp = resultTemp - (newHumidityRatio - normalHumidityRatio) / dwBydTwb
        newHumidityRatio = calculateHumidityRatioFromWetBulb(dryBulbTemp, resultTemp, pressure)
    Loop

    wetBulbTemperature = resultTemp
End Function

' Function to calculate humidity ratio given dry bulb and wet bulb temperatures
Function calculateHumidityRatioFromWetBulb(dryBulbTemp, wetBulbTemp, pressure)
    Dim saturationPressure, humidityRatio

    saturationPressure = calculateSaturationPressure(wetBulbTemp)
    humidityRatio = 0.62198 * saturationPressure / (pressure - saturationPressure)

    If dryBulbTemp >= 0 Then
        calculateHumidityRatioFromWetBulb = (((2501 - 2.326 * wetBulbTemp) * humidityRatio - 1.006 * (dryBulbTemp - wetBulbTemp)) / (2501 + 1.86 * dryBulbTemp - 4.186 * wetBulbTemp))
    Else
        calculateHumidityRatioFromWetBulb = (((2830 - 0.24 * wetBulbTemp) * humidityRatio - 1.006 * (dryBulbTemp - wetBulbTemp)) / (2830 + 1.86 * dryBulbTemp - 2.1 * wetBulbTemp))
    End If
End Function

' Function to calculate humidity ratio given dry bulb temperature and relative humidity
Function calculateHumidityRatio(dryBulbTemp, relativeHumidity, pressure)
    Dim saturationPressure

    saturationPressure = calculateSaturationPressure(dryBulbTemp)
    calculateHumidityRatio = 0.62198 * relativeHumidity * saturationPressure / (pressure - relativeHumidity * saturationPressure)
End Function

' Function to compute partial vapor pressure
Function calculatePartialVaporPressure(pressure, humidityRatio)
    calculatePartialVaporPressure = pressure * humidityRatio / (0.62198 + humidityRatio)
End Function

' Function to calculate saturation pressure given dry bulb temperature
Function calculateSaturationPressure(dryBulbTemp)
    Dim C1, C2, C3, C4, C5, C6, C7, C8, C9, C10, C11, C12, C13, kelvinTemp

    ' Coefficients for the equations
    C1 = -5674.5359
    C2 = 6.3925247
    C3 = -0.009677843
    C4 = 0.00000062215701
    C5 = 2.0747825E-09
    C6 = -9.484024E-13
    C7 = 4.1635019
    C8 = -5800.2206
    C9 = 1.3914993
    C10 = -0.048640239
    C11 = 0.000041764768
    C12 = -0.000000014452093
    C13 = 6.5459673

    kelvinTemp = dryBulbTemp + 273.15

    If kelvinTemp <= 273.15 Then
        calculateSaturationPressure = Exp(C1 / kelvinTemp + C2 + C3 * kelvinTemp + C4 * kelvinTemp ^ 2 + C5 * kelvinTemp ^ 3 + C6 * kelvinTemp ^ 4 + C7 * Log(kelvinTemp)) / 1000
    Else
        calculateSaturationPressure = Exp(C8 / kelvinTemp + C9 + C10 * kelvinTemp + C11 * kelvinTemp ^ 2 + C12 * kelvinTemp ^ 3 + C13 * Log(kelvinTemp)) / 1000
    End If
End Function

' Function to calculate relative humidity given dry bulb and wet bulb temperatures
Function calculateRelativeHumidity(dryBulbTemp, wetBulbTemp, pressure)
    Dim humidityRatio, result

    humidityRatio = calculateHumidityRatioFromWetBulb(dryBulbTemp, wetBulbTemp, pressure)
    result = calculatePartialVaporPressure(pressure, humidityRatio) / calculateSaturationPressure(dryBulbTemp)

    calculateRelativeHumidity = result
End Function

' Function to calculate relative humidity given dry bulb temperature and humidity ratio
Function calculateRelativeHumidityFromRatio(dryBulbTemp, humidityRatio, pressure)
    Dim partialVaporPressure, saturationPressure, result

    partialVaporPressure = calculatePartialVaporPressure(pressure, humidityRatio)
    saturationPressure = calculateSaturationPressure(dryBulbTemp)
    result = partialVaporPressure / saturationPressure

    calculateRelativeHumidityFromRatio = result
End Function

' Function to calculate enthalpy of moist air given dry bulb temperature and humidity ratio
Function calculateEnthalpy(dryBulbTemp, humidityRatio)
    calculateEnthalpy = 1.006 * dryBulbTemp + humidityRatio * (2501 + 1.86 * dryBulbTemp)
End Function

' Function to calculate dry bulb temperature from enthalpy and humidity ratio
Function calculateDryBulbTemperature(enthalpy, humidityRatio)
    calculateDryBulbTemperature = (enthalpy - (2501 * humidityRatio)) / (1.006 + (1.86 * humidityRatio))
End Function

' Function to compute the dew point temperature
Function calculateDewPoint(pressure, humidityRatio)
    Dim C14, C15, C16, C17, C18, partialVaporPressure, alpha, dewPoint1, dewPoint2, result

    C14 = 6.54
    C15 = 14.526
    C16 = 0.7389
    C17 = 0.09486
    C18 = 0.4569

    partialVaporPressure = calculatePartialVaporPressure(pressure, humidityRatio)
    alpha = Log(partialVaporPressure)
    dewPoint1 = C14 + C15 * alpha + C16 * alpha ^ 2 + C17 * alpha ^ 3 + C18 * partialVaporPressure ^ 0.1984
    dewPoint2 = 6.09 + 12.608 * alpha + 0.4959 * alpha ^ 2

    If dewPoint1 >= 0 Then
        result = dewPoint1
    Else
        result = dewPoint2
    End If

    calculateDewPoint = result
End Function

' Function to compute the dry air density given pressure, temperature, and humidity ratio
Function calculateDryAirDensity(pressure, dryBulbTemp, humidityRatio)
    Dim gasConstantDryAir, result

    gasConstantDryAir = 287.055
    result = 1000 * pressure / (gasConstantDryAir * (273.15 + dryBulbTemp) * (1 + 1.6078 * humidityRatio))

    calculateDryAirDensity = result
End Function

' Function to convert absolute humidity to relative humidity
Function convertAbsoluteToRelativeHumidity(dryBulbTemp, absoluteHumidity, pressure)
    If IsNumeric(absoluteHumidity) Then
        Dim saturationVaporPressure, saturationAbsoluteHumidity

        saturationVaporPressure = calculateSaturationVaporPressure(dryBulbTemp)
        saturationAbsoluteHumidity = 0.622 * saturationVaporPressure / (pressure - saturationVaporPressure)

        convertAbsoluteToRelativeHumidity = absoluteHumidity / saturationAbsoluteHumidity * 100
    End If
End Function

' Function to calculate total cooling
Function calculateTotalCooling(dryBulbTemp, humidityRatio, pressure)
    Dim initialEnthalpy, finalEnthalpy, finalDryBulbTemp, finalHumidityRatio

    initialEnthalpy = calculateEnthalpy(dryBulbTemp, humidityRatio)

    ' Assuming the required final state values are provided
    ' Replace the placeholders with actual values
    finalDryBulbTemp = U18 ' Placeholder value
    finalHumidityRatio = U20 ' Placeholder value

    finalEnthalpy = calculateEnthalpy(finalDryBulbTemp, finalHumidityRatio)
    calculateTotalCooling = finalEnthalpy - initialEnthalpy
End Function

' Psychrometric function to calculate various properties
Function calculatePsychrometricProperties(pressure, inputType1, inputValue1, inputType2, inputValue2, outputType, Optional unitType As String = "Imp")
    Dim dryBulbTemp, wetBulbTemp, dewPoint, relativeHumidity, humidityRatio, enthalpy, outputValue

    If inputType1 <> "h" And inputType1 <> "W" And inputType1 <> "Tdb" Then
        calculatePsychrometricProperties = "NAN"
        Exit Function
    End If

    If inputType1 = inputType2 Then
        calculatePsychrometricProperties = "NAN"
        Exit Function
    End If

    ' Convert to SI units if necessary
    If unitType = "SI" Then
        pressure = pressure / 1000

        If inputType1 = "Tdb" Then
            dryBulbTemp = inputValue1
        ElseIf inputType1 = "W" Then
            humidityRatio = inputValue1
        ElseIf inputType1 = "h" Then
            enthalpy = inputValue1
        End If

        If inputType2 = "Tdb" Then
            dryBulbTemp = inputValue2
        ElseIf inputType2 = "Twb" Then
            wetBulbTemp = inputValue2
        ElseIf inputType2 = "DP" Then
            dewPoint = inputValue2
        ElseIf inputType2 = "RH" Then
            relativeHumidity = inputValue2
        ElseIf inputType2 = "W" Then
            humidityRatio = inputValue2
        ElseIf inputType2 = "h" Then
            enthalpy = inputValue2
        End If
    Else
        ' Convert to SI units
        pressure = (pressure * 4.4482216152605) / (0.0254 ^ 2 * 1000)

        If inputType1 = "Tdb" Then
            dryBulbTemp = (inputValue1 - 32) / 1.8
        ElseIf inputType1 = "W" Then
            humidityRatio = inputValue1
        ElseIf inputType1 = "h" Then
            enthalpy = ((inputValue1 * 1.055056) / 0.45359237) - 17.884444444
        End If

        If inputType2 = "Tdb" Then
            dryBulbTemp = (inputValue2 - 32) / 1.8
        ElseIf inputType2 = "Twb" Then
            wetBulbTemp = (inputValue2 - 32) / 1.8
        ElseIf inputType2 = "DP" Then
            dewPoint = (inputValue2 - 32) / 1.8
        ElseIf inputType2 = "RH" Then
            relativeHumidity = inputValue2
        ElseIf inputType2 = "W" Then
            humidityRatio = inputValue2
        ElseIf inputType2 = "h" Then
            enthalpy = ((inputValue2 * 1.055056) / 0.45359237) - 17.884444444
        End If
    End If

    If (inputType1 = "h" And inputType2 = "W") Or (inputType1 = "W" And inputType2 = "h") Then
        dryBulbTemp = calculateDryBulbTemperature(enthalpy, humidityRatio)
    End If

    ' Determine output based on requested output type
    If outputType = "RH" Or outputType = "Twb" Then
        If inputType2 = "Twb" Then
            relativeHumidity = calculateRelativeHumidity(dryBulbTemp, wetBulbTemp, pressure)
        ElseIf inputType2 = "DP" Then
            relativeHumidity = calculateSaturationPressure(dewPoint) / calculateSaturationPressure(dryBulbTemp)
        ElseIf inputType2 = "W" Then
            relativeHumidity = calculatePartialVaporPressure(pressure, humidityRatio) / calculateSaturationPressure(dryBulbTemp)
        ElseIf inputType2 = "h" Then
            humidityRatio = (1.006 * dryBulbTemp - enthalpy) / (-(2501 + 1.86 * dryBulbTemp))
            relativeHumidity = calculatePartialVaporPressure(pressure, humidityRatio) / calculateSaturationPressure(dryBulbTemp)
        End If
    Else
        If inputType1 <> "W" Then
            If inputType2 = "Twb" Then
                humidityRatio = calculateHumidityRatioFromWetBulb(dryBulbTemp, wetBulbTemp, pressure)
            ElseIf inputType2 = "DP" Then
                humidityRatio = 0.621945 * calculateSaturationPressure(dewPoint) / (pressure - calculateSaturationPressure(dewPoint))
            ElseIf inputType2 = "RH" Then
                humidityRatio = calculateHumidityRatio(dryBulbTemp, relativeHumidity, pressure)
            ElseIf inputType2 = "h" Then
                humidityRatio = (1.006 * dryBulbTemp - enthalpy) / (-(2501 + 1.86 * dryBulbTemp))
            End If
        End If
    End If

    ' Assign the final output value based on output type
    Select Case outputType
        Case "Tdb"
            outputValue = dryBulbTemp
        Case "Twb"
            outputValue = wetBulbTemperature(dryBulbTemp, relativeHumidity, pressure)
        Case "DP"
            outputValue = calculateDewPoint(pressure, humidityRatio)
        Case "RH"
            outputValue = relativeHumidity
        Case "W"
            outputValue = humidityRatio
        Case "WVP"
            outputValue = calculatePartialVaporPressure(pressure, humidityRatio) * 1000
        Case "DSat"
            outputValue = humidityRatio / calculateHumidityRatio(dryBulbTemp, 1, pressure)
        Case "h"
            outputValue = calculateEnthalpy(dryBulbTemp, humidityRatio)
        Case "s"
            ' No equation for entropy; induce an error
            outputValue = 5 / 0
        Case "SV"
            outputValue = 1 / (calculateDryAirDensity(pressure, dryBulbTemp, humidityRatio))
        Case "MAD"
            outputValue = calculateDryAirDensity(pressure, dryBulbTemp, humidityRatio) * (1 + humidityRatio)
    End Select

    ' Convert to Imperial units if required
    If unitType = "Imp" Then
        Select Case outputType
            Case "Tdb", "Twb", "DP"
                outputValue = 1.8 * outputValue + 32
            Case "WVP"
                outputValue = outputValue * 0.0254 ^ 2 / 4.448230531
            Case "h"
                outputValue = (outputValue + 17.88444444444) * 0.45359237 / 1.055056
            Case "SV"
                outputValue = outputValue * 0.45359265 / ((12 * 0.0254) ^ 3)
            Case "MAD"
                outputValue = outputValue * (12 * 0.0254) ^ 3 / 0.45359265
        End Select
    End If

    calculatePsychrometricProperties = outputValue
End Function

' Custom logarithm function using series expansion
Function customLog(x)
    Dim result, y, power, term, n, maxIterations

    maxIterations = 100
    y = (x - 1) / (x + 1)
    result = 0
    power = y
    term = 2 * power / 1

    For n = 1 To maxIterations
        result = result + term
        power = power * y * y
        term = 2 * power / (2 * n + 1)
    Next

    customLog = result
End Function

' Function to round a number to one decimal place
Function roundToOneDecimal(number)
    roundToOneDecimal = Int(number * 10 + 0.5) / 10
End Function

' Function to calculate dew point
Function calculateDewPointTemperature(dryBulbTemp, humidity)
    Dim tempCelsius
    tempCelsius = (dryBulbTemp - 32) * 5 / 9
    calculateDewPointTemperature = roundToOneDecimal(((tempCelsius - ((100 - humidity) / 5)) * 1.8) + 32)
End Function

' Function to round up a number
Function roundUp(number)
    roundedValue = Int(number + 0.999999999)
    roundUp = CStr(roundedValue) & ".0"
End Function

' Function to format a number with one decimal place
Function formatNumberWithDecimal(number)
    formatNumberWithDecimal = FormatNumber(number, 1)
End Function

' Sleep function placeholder (uncomment WScript.Sleep if required)
Sub sleep(milliseconds)
    'WScript.Sleep(milliseconds)
End Sub

' Function to change control settings
' statusField: the override checkbox
' controlField: the control to be adjusted
' setpoint: the new setpoint value
' mode: "manual" or "auto" mode
Sub changeControlSettings(statusField, controlField, setpoint, mode)
    Dim formattedSetpoint
    formattedSetpoint = formatNumberWithDecimal(setpoint)

    If mode = "manual" Then
        If statusField.value <> True Then
            statusField.value = True
            If controlField.value <> formattedSetpoint Then
                controlField.value = formattedSetpoint
                controlField.linecolor = RGB(0, 255, 0)
                sleep(sleepDuration)
            End If
        Else
            If controlField.value <> formattedSetpoint Then
                controlField.value = formattedSetpoint
                controlField.linecolor = RGB(0, 255, 0)
                'sleep(sleepDuration)
            End If
        End If
    Else
        If statusField.value = True Then
            statusField.value = False
            sleep(sleepDuration)
        End If
    End If
End Sub

' Cooling scenario A
Sub executeCoolingScenario(raDamper1, raDamper2, oaDamper1, oaDamper2, blowerControl1, blowerControl2, hv4Damper, rhSp1, rhSp2, clgSp1, clgSp2, htgSp1, htgSp2)
    changeControlSettings ahu1RAdamperMode, ahu1RAdamper, raDamper1, "manual"
    changeControlSettings ahu2RAdamperMode, ahu2RAdamper, raDamper2, "manual"
    changeControlSettings ahu1OAdamperMode, ahu1OAdamper, oaDamper1, "manual"
    changeControlSettings ahu2OAdamperMode, ahu2OAdamper, oaDamper2, "manual"
    changeControlSettings ahu1BlowerMode, ahu1BlowerControl, blowerControl1, "manual"
    changeControlSettings ahu2BlowerMode, ahu2BlowerControl, blowerControl2, "manual"
    changeControlSettings hv4DamperMode, hv4MAdamper, hv4Damper, "manual"
    changeControlSettings ahu1RhSpMode, ahu1RhSp, rhSp1, "manual"
    changeControlSettings ahu2RhSpMode, ahu2RhSp, rhSp2, "manual"
    changeControlSettings ahu1ClgSpMode, ahu1ClgSp, clgSp1, "manual"
    changeControlSettings ahu2ClgSpMode, ahu2ClgSp, clgSp2, "manual"
    changeControlSettings ahu1HtgSpMode, ahu1HtgSp, htgSp1, "manual"
    changeControlSettings ahu2HtgSpMode, ahu2HtgSp, htgSp2, "manual"
End Sub
