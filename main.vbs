Sub alphanum30_OnPeriodicUpdate(  )
Dim IsSelected
Dim currentTime, currentHour, amOrPm, shift
Dim dp, dp1, dp2
Dim avgRaTemp
Dim avgRaRh, a, b, alpha

currentTime = Time
FormatDateTime currentTime, vbLongTime

currentHour = Hour(currentTime)
currentMinutes = Minute(currentTime)
amOrPm =currentTime
timeNow.value = currentTime
IsSelected = copilot.value
If (currentHour >= 7) And (currentHour <15) Then
	shift = 2
ElseIf (currentHour >= 15) And (currentHour < 23) Then
	shift = 3
Else
	shift = 1
End If
currentShift.value = shift
' Calulate dew point
' Magnus Constants
a = 17.27
b = 237.7
avgRaTemp = (((CDbl(Abs(ahu2RAtemp.value)) + CDbl(Abs(ahu1RAtemp.value))) / 2) - 32)  *  5 / 9
avgRaRh = (CDbl(ahu1RH.value) + CDbl(ahu2RH.value)) / 2
dp =  RoundToOneDecimal(((avgRaTemp  -  ((100 - avgRaRh) / 5)) * 1.8) + 32) 
DewPoint.value = dp 
If IsSelected = True Then 
	' Co-pilot has been selected.
	If shift = 2 Or coolDown.value = True Then
		' Scenario controls various dampers, blower controls, and setpoints based on the outside temperature.
		' Execute 2nd shift schedule
		' Scenario controls various dampers, blower controls, and setpoints based on the outside temperature.
		
		If outsideTemp.value > 85 Then
		    ' If outside temperature is greater than 85
		    ' Set Scenario parameters for high cooling
		    Scenario 0, 0, 15, 15, 88, 88, 30, 58, 58, 55, 55, 53, 53

		ElseIf outsideTemp.value > 80 And outsideTemp.value < 85 Then
		    ' If outside temperature is between 80 and 85
		    ' Set Scenario parameters for moderate cooling
		    Scenario 0, 0, 30, 30, 92, 92, 50, 55, 55, 55, 55, 53, 53

		ElseIf outsideTemp.value > 72 And outsideTemp.value < 80 Then
		    ' If outside temperature is between 72 and 80
		    ' Set Scenario parameters for mild cooling
		    Scenario 0, 0, 30, 30, 98, 98, 75, 50, 50, 55, 55, 53, 53

		ElseIf outsideTemp.value < 72 And outsideTemp.value > 50 Then
		    ' If outside temperature is between 50 and 72
		    ' Set Scenario parameters for neutral temperature conditions
		    Scenario 30, 30, 30, 30, 98, 98, 100, 55, 55, 55, 55, 53, 53

		ElseIf outsideTemp.value < 50 And outsideTemp.value > 42 Then
		    ' If outside temperature is between 42 and 50
		    ' Set Scenario parameters for moderate heating
		    Scenario 100, 100, 100, 100, 82, 82, 50, 50, 50, 55, 55, 53, 53

		ElseIf outsideTemp.value < 42 And outsideTemp.value > 25 Then
		' If outside temperature is between 25 and 42
		' Set Scenario parameters for moderate heating
		Scenario 30, 30, 30, 30, 90, 90, 75, 35, 35, 55, 55, 53, 53

		ElseIf outsideTemp.value < 25 Then
		    ' If outside temperature is less than 25
		    ' Set Scenario parameters for high heating
		    Scenario 0, 0, 30, 30, 85, 85, 20, 20, 20, 53, 53

		Else
		    ' If none of the above conditions are met
		    ' Ensure hv4MAdamper is set to 50
		    If hv4MAdamper.value <> 50 Then
			hv4MAdamper.value = 50
		    End If
		End If
	ElseIf shift = 3 Then 
		
		' Execute 3rd shift schedule
		' Scenario controls various dampers, blower controls, and setpoints based on the outside temperature.
		
		If outsideTemp.value > 85 Then
		' If outside temperature is greater than 85
		' Set Scenario parameters for high cooling
		Scenario 0, 0, 15, 15, 88, 88, 30, 58, 58, 55, 55, 53, 53

		ElseIf outsideTemp.value > 80 And outsideTemp.value < 85 Then
		' If outside temperature is between 80 and 85
		' Set Scenario parameters for moderate cooling
		Scenario 0, 0, 30, 30, 92, 92, 50, 55, 55, 55, 55, 53, 53

		ElseIf outsideTemp.value > 72 And outsideTemp.value < 80 Then
		' If outside temperature is between 72 and 80
		' Set Scenario parameters for mild cooling
		Scenario 0, 0, 30, 30, 98, 98, 75, 50, 50, 55, 55, 53, 53

		ElseIf outsideTemp.value < 72 And outsideTemp.value > 50 Then
		' If outside temperature is between 50 and 72
		' Set Scenario parameters for neutral temperature conditions
		Scenario 30, 30, 30, 30, 98, 98, 100, 55, 55, 55, 55, 53, 53

		ElseIf outsideTemp.value < 50 And outsideTemp.value > 42 Then
		' If outside temperature is between 42 and 50
		' Set Scenario parameters for moderate heating
		Scenario 100, 100, 100, 100, 82, 82, 50, 50, 50, 55, 55, 53, 53

		ElseIf outsideTemp.value < 42 And outsideTemp.value > 25 Then
		' If outside temperature is between 25 and 42
		' Set Scenario parameters for moderate heating
		Scenario 30, 30, 30, 30, 90, 90, 75, 35, 35, 55, 55, 53, 53

		ElseIf outsideTemp.value < 25 Then
		' If outside temperature is less than 25
		' Set Scenario parameters for high heating
		Scenario 0, 0, 10, 10, 85, 85, 20, 20, 20, 53, 53

		Else
		' If none of the above conditions are met
		' Ensure hv4MAdamper is set to 50
			If hv4MAdamper.value <> 50 Then
				hv4MAdamper.value = 50
			End If
		End If
	Else

		' Execute 1st shift schedule
		' Scenario controls various dampers, blower controls, and setpoints based on the outside temperature.

		If outsideTemp.value > 85 Then
		' If outside temperature is greater than 85
		' Set Scenario parameters for high cooling
		Scenario 0, 0, 0, 0, 92, 92, 30, 58, 58, 55, 55, 53, 53

		ElseIf outsideTemp.value > 80 And outsideTemp.value < 85 Then
		' If outside temperature is between 80 and 85
		' Set Scenario parameters for moderate cooling
		Scenario 0, 0, 30, 30, 92, 92, 50, 55, 55, 57, 57, 53, 53

		ElseIf outsideTemp.value > 72 And outsideTemp.value < 80 Then
		' If outside temperature is between 72 and 80
		' Set Scenario parameters for mild cooling
		Scenario 0, 0, 30, 30, 98, 98, 75, 50, 50, 57, 57, 53, 53

		ElseIf (outsideTemp.value < 72 And outsideTemp.value > 50) Or outsideTemp.value < ahu1RAtemp.value Then
		' If outside temperature is between 50 and 72 or less than the return air temperature
		' Set Scenario parameters for neutral or slightly cooling conditions
		Scenario 30, 30, 30, 30, 98, 98, 100, 65, 65, 57, 57, 53, 53

		ElseIf outsideTemp.value < 50 And outsideTemp.value > 42 Then
		' If outside temperature is between 42 and 50
		' Set Scenario parameters for moderate heating
		Scenario 100, 100, 100, 100, 82, 82, 50, 50, 50, 55, 55, 53, 53

		ElseIf outsideTemp.value < 42 And outsideTemp.value > 25 Then
		' If outside temperature is between 25 and 42
		' Set Scenario parameters for moderate heating
		Scenario 30, 30, 30, 30, 90, 90, 75, 35, 35, 55, 55, 53, 53

		ElseIf outsideTemp.value < 25 Then
		' If outside temperature is less than 25
		' Set Scenario parameters for high heating
		Scenario 0, 0, 10, 10, 85, 85, 20, 20, 20, 53, 53

		Else
		' If none of the above conditions are met
		' Ensure hv4MAdamper is set to 50
			If hv4MAdamper.value <> 50 Then
				hv4MAdamper.value = 50
			End If
		End If
	End If
End If
End Sub
