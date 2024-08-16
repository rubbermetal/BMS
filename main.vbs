Sub alphanum30_OnPeriodicUpdate(  )
Dim IsSelected, currentTime, currentHour, amOrPm, shift, dp, dp1, dp2, avgRaTemp, avgRaRh, a, b, alpha

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
	' This block is for occupied settings
	If shift = 2 Or coolDown.value = True Then
		' Co-pilot has been selected.
		' execute 2nd shift schedule
		If outsideTemp.value > 98 Then
				  ' AHU1 RA Damper
			ScenarioA 0, _
				  ' AHU2 RA Damper
				  0, _
				  ' AHU1 OA Damper
				 0, _
				  ' AHU2 OA Damper
				 0, _
				  ' AHU1 Blower Speed
				 85, _
				  ' AHU2 Blower Speed
				 85, _
				  ' HV4 Damper
				 20, _
				 ' AHU1 RH Setpoint
				 60, _
				 ' AHU2 RH Setpoint
				 60, _
				 ' AHU1 Cooling DA Setpoint
				 55, _
				 ' AHU2 Cooling DA Setpoint
				 55, _
				 ' AHU1 Heating DA Setpoint
				 53, _
				 ' AHU2 Heating DA Setpoint
				 53
		ElseIf outsideTemp.value > 92 And outsideTemp.value < 98 Then
				  ' AHU1 RA Damper
			ScenarioA 0, _
				  ' AHU2 RA Damper
				  0, _
				  ' AHU1 OA Damper
				 15, _
				  ' AHU2 OA Damper
				 15, _
				  ' AHU1 Blower Speed
				 88, _
				  ' AHU2 Blower Speed
				 88, _
				  ' HV4 Damper
				 30, _
				 ' AHU1 RH Setpoint
				 55, _
				 ' AHU2 RH Setpoint
				 55, _
				 ' AHU1 Cooling DA Setpoint
				 55, _
				 ' AHU2 Cooling DA Setpoint
				 55, _
				 ' AHU1 Heating DA Setpoint
				 53, _
				 ' AHU2 Heating DA Setpoint
				 53
		ElseIf outsideTemp.value > 85 And outsideTemp.value < 92 Then
				  ' AHU1 RA Damper
			ScenarioA 0, _
				  ' AHU2 RA Damper
				  0, _
				  ' AHU1 OA Damper
				 25, _
				  ' AHU2 OA Damper
				 25, _
				  ' AHU1 Blower Speed
				 88, _
				  ' AHU2 Blower Speed
				 88, _
				  ' HV4 Damper
				 30, _
				 ' AHU1 RH Setpoint
				 55, _
				 ' AHU2 RH Setpoint
				 55, _
				 ' AHU1 Cooling DA Setpoint
				 55, _
				 ' AHU2 Cooling DA Setpoint
				 55, _
				 ' AHU1 Heating DA Setpoint
				 53, _
				 ' AHU2 Heating DA Setpoint
				 53
		' Economizer Mode
		ElseIf outsideTemp.value > 80 And outsideTemp.value < 85 Then
				  ' AHU1 RA Damper
			ScenarioA 0, _
				  ' AHU2 RA Damper
				  0, _
				  ' AHU1 OA Damper
				 35, _
				  ' AHU2 OA Damper
				 35, _
				  ' AHU1 Blower Speed
				 92, _
				  ' AHU2 Blower Speed
				 92, _
				  ' HV4 Damper
				 50, _
				 ' AHU1 RH Setpoint
				 55, _
				 ' AHU2 RH Setpoint
				 55, _
				 ' AHU1 Cooling DA Setpoint
				 55, _
				 ' AHU2 Cooling DA Setpoint
				 55, _
				 ' AHU1 Heating DA Setpoint
				 53, _
				 ' AHU2 Heating DA Setpoint
				 53
		' Economizer Mode
		ElseIf outsideTemp.value > 72 And outsideTemp.value < 80 Then
				  ' AHU1 RA Damper
			ScenarioA 0, _
				  ' AHU2 RA Damper
				  0, _
				  ' AHU1 OA Damper
				 35, _
				  ' AHU2 OA Damper
				 35, _
				  ' AHU1 Blower Speed
				 98, _
				  ' AHU2 Blower Speed
				 98, _
				  ' HV4 Damper
				 75, _
				 ' AHU1 RH Setpoint
				 55, _
				 ' AHU2 RH Setpoint
				 55, _
				 ' AHU1 Cooling DA Setpoint
				 55, _
				 ' AHU2 Cooling DA Setpoint
				 55, _
				 ' AHU1 Heating DA Setpoint
				 53, _
				 ' AHU2 Heating DA Setpoint
				 53
		' Purge Heat Via Ventialtion
		ElseIf (outsideTemp.value < 72 And outsideTemp.value > 50) Or outsideTemp.value <  ahu1RAtemp.value Then
				  ' AHU1 RA Damper
			ScenarioA 50, _
				  ' AHU2 RA Damper
				  50, _
				  ' AHU1 OA Damper
				 50, _
				  ' AHU2 OA Damper
				 50, _
				  ' AHU1 Blower Speed
				 98, _
				  ' AHU2 Blower Speed
				 98, _
				  ' HV4 Damper
				 100, _
				 ' AHU1 RH Setpoint
				 50, _
				 ' AHU2 RH Setpoint
				 50, _
				 ' AHU1 Cooling DA Setpoint
				 55, _
				 ' AHU2 Cooling DA Setpoint
				 55, _
				 ' AHU1 Heating DA Setpoint
				 53, _
				 ' AHU2 Heating DA Setpoint
				 53
		' 100% Outside Air
		ElseIf outsideTemp.value < 50 And outsideTemp.value > 42 Then
				  ' AHU1 RA Damper
			ScenarioA 100, _
				  ' AHU2 RA Damper
				  100, _
				  ' AHU1 OA Damper
				 100, _
				  ' AHU2 OA Damper
				 100, _
				  ' AHU1 Blower Speed
				 82, _
				  ' AHU2 Blower Speed
				 82, _
				  ' HV4 Damper
				 50, _
				 ' AHU1 RH Setpoint
				 50, _
				 ' AHU2 RH Setpoint
				 50, _
				 ' AHU1 Cooling DA Setpoint
				 55, _
				 ' AHU2 Cooling DA Setpoint
				 55, _
				 ' AHU1 Heating DA Setpoint
				 53, _
				 ' AHU2 Heating DA Setpoint
				 53
		' Purge Ventialtion
		ElseIf outsideTemp.value < 42 And outsideTemp.value > 15 Then				  
				  ' AHU1 RA Damper
			ScenarioA 50, _
				  ' AHU2 RA Damper
				  50, _
				  ' AHU1 OA Damper
				 50, _
				  ' AHU2 OA Damper
				 50, _
				  ' AHU1 Blower Speed
				 88, _
				  ' AHU2 Blower Speed
				 88, _
				  ' HV4 Damper
				 50, _
				 ' AHU1 RH Setpoint
				 35, _
				 ' AHU2 RH Setpoint
				 35, _
				 ' AHU1 Cooling DA Setpoint
				 55, _
				 ' AHU2 Cooling DA Setpoint
				 55, _
				 ' AHU1 Heating DA Setpoint
				 53, _
				 ' AHU2 Heating DA Setpoint
				 53
		ElseIf outsideTemp.value < 15 And outsideTemp.value > 0 Then
				  ' AHU1 RA Damper
			ScenarioA 0, _
				  ' AHU2 RA Damper
				  0, _
				  ' AHU1 OA Damper
				 30, _
				  ' AHU2 OA Damper
				 30, _
				  ' AHU1 Blower Speed
				 85, _
				  ' AHU2 Blower Speed
				 85, _
				  ' HV4 Damper
				 50, _
				 ' AHU1 RH Setpoint
				 25, _
				 ' AHU2 RH Setpoint
				 25, _
				 ' AHU1 Cooling DA Setpoint
				 55, _
				 ' AHU2 Cooling DA Setpoint
				 55, _
				 ' AHU1 Heating DA Setpoint
				 53, _
				 ' AHU2 Heating DA Setpoint
				 53
		' We are below 0 degrees F
		Else
				  ' AHU1 RA Damper
			ScenarioA 0, _
				  ' AHU2 RA Damper
				  0, _
				  ' AHU1 OA Damper
				 10, _
				  ' AHU2 OA Damper
				 10, _
				  ' AHU1 Blower Speed
				 85, _
				  ' AHU2 Blower Speed
				 85, _
				  ' HV4 Damper
				 50, _
				 ' AHU1 RH Setpoint
				 25, _
				 ' AHU2 RH Setpoint
				 25, _
				 ' AHU1 Cooling DA Setpoint
				 55, _
				 ' AHU2 Cooling DA Setpoint
				 55, _
				 ' AHU1 Heating DA Setpoint
				 53, _
				 ' AHU2 Heating DA Setpoint
				 53
		End If
	ElseIf shift = 3 Then 
		If outsideTemp.value > 85 Then
				  ' AHU1 RA Damper
			ScenarioA 0, _
				  ' AHU2 RA Damper
				  0, _
				  ' AHU1 OA Damper
				 15, _
				  ' AHU2 OA Damper
				 15, _
				  ' AHU1 Blower Speed
				 88, _
				  ' AHU2 Blower Speed
				 88, _
				  ' HV4 Damper
				 30, _
				 ' AHU1 RH Setpoint
				 55, _
				 ' AHU2 RH Setpoint
				 55, _
				 ' AHU1 Cooling DA Setpoint
				 55, _
				 ' AHU2 Cooling DA Setpoint
				 55, _
				 ' AHU1 Heating DA Setpoint
				 53, _
				 ' AHU2 Heating DA Setpoint
				 53
		ElseIf outsideTemp.value > 80 And outsideTemp.value < 85 Then
				  ' AHU1 RA Damper
			ScenarioA 0, _
				  ' AHU2 RA Damper
				  0, _
				  ' AHU1 OA Damper
				 30, _
				  ' AHU2 OA Damper
				 30, _
				  ' AHU1 Blower Speed
				 92, _
				  ' AHU2 Blower Speed
				 92, _
				  ' HV4 Damper
				 50, _
				 ' AHU1 RH Setpoint
				 55, _
				 ' AHU2 RH Setpoint
				 55, _
				 ' AHU1 Cooling DA Setpoint
				 55, _
				 ' AHU2 Cooling DA Setpoint
				 55, _
				 ' AHU1 Heating DA Setpoint
				 53, _
				 ' AHU2 Heating DA Setpoint
				 53
		ElseIf outsideTemp.value > 72 And outsideTemp.value < 80 Then
				  ' AHU1 RA Damper
			ScenarioA 0, _
				  ' AHU2 RA Damper
				  0, _
				  ' AHU1 OA Damper
				 30, _
				  ' AHU2 OA Damper
				 30, _
				  ' AHU1 Blower Speed
				 98, _
				  ' AHU2 Blower Speed
				 98, _
				  ' HV4 Damper
				 75, _
				 ' AHU1 RH Setpoint
				 55, _
				 ' AHU2 RH Setpoint
				 55, _
				 ' AHU1 Cooling DA Setpoint
				 55, _
				 ' AHU2 Cooling DA Setpoint
				 55, _
				 ' AHU1 Heating DA Setpoint
				 53, _
				 ' AHU2 Heating DA Setpoint
				 53
		ElseIf (outsideTemp.value < 72 And outsideTemp.value > 50) Or outsideTemp.value <  ahu1RAtemp.value Then
				  ' AHU1 RA Damper
			ScenarioA 30, _
				  ' AHU2 RA Damper
				  30, _
				  ' AHU1 OA Damper
				 30, _
				  ' AHU2 OA Damper
				 30, _
				  ' AHU1 Blower Speed
				 98, _
				  ' AHU2 Blower Speed
				 98, _
				  ' HV4 Damper
				 100, _
				 ' AHU1 RH Setpoint
				 50, _
				 ' AHU2 RH Setpoint
				 50, _
				 ' AHU1 Cooling DA Setpoint
				 55, _
				 ' AHU2 Cooling DA Setpoint
				 55, _
				 ' AHU1 Heating DA Setpoint
				 53, _
				 ' AHU2 Heating DA Setpoint
				 53
		ElseIf outsideTemp.value < 50 And outsideTemp.value > 42 Then
				  ' AHU1 RA Damper
			ScenarioA 100, _
				  ' AHU2 RA Damper
				  100, _
				  ' AHU1 OA Damper
				 100, _
				  ' AHU2 OA Damper
				 100, _
				  ' AHU1 Blower Speed
				 82, _
				  ' AHU2 Blower Speed
				 82, _
				  ' HV4 Damper
				 50, _
				 ' AHU1 RH Setpoint
				 50, _
				 ' AHU2 RH Setpoint
				 50, _
				 ' AHU1 Cooling DA Setpoint
				 55, _
				 ' AHU2 Cooling DA Setpoint
				 55, _
				 ' AHU1 Heating DA Setpoint
				 53, _
				 ' AHU2 Heating DA Setpoint
				 53
		ElseIf outsideTemp.value < 25 Then
				  ' AHU1 RA Damper
			ScenarioA 30, _
				  ' AHU2 RA Damper
				  30, _
				  ' AHU1 OA Damper
				 30, _
				  ' AHU2 OA Damper
				 30, _
				  ' AHU1 Blower Speed
				 85, _
				  ' AHU2 Blower Speed
				 85, _
				  ' HV4 Damper
				 50, _
				 ' AHU1 RH Setpoint
				 35, _
				 ' AHU2 RH Setpoint
				 35, _
				 ' AHU1 Cooling DA Setpoint
				 55, _
				 ' AHU2 Cooling DA Setpoint
				 55, _
				 ' AHU1 Heating DA Setpoint
				 53, _
				 ' AHU2 Heating DA Setpoint
				 53
		Else
			If hv4MAdamper.value <> 50 Then
				hv4MAdamper.value = 50
				
			End If
		End If
	Else
		' Shift 1
		If outsideTemp.value > 85 Then
				  ' AHU1 RA Damper
			ScenarioA 0, _
				  ' AHU2 RA Damper
				  0, _
				  ' AHU1 OA Damper
				 15, _
				  ' AHU2 OA Damper
				 15, _
				  ' AHU1 Blower Speed
				 88, _
				  ' AHU2 Blower Speed
				 88, _
				  ' HV4 Damper
				 30, _
				 ' AHU1 RH Setpoint
				 55, _
				 ' AHU2 RH Setpoint
				 55, _
				 ' AHU1 Cooling DA Setpoint
				 55, _
				 ' AHU2 Cooling DA Setpoint
				 55, _
				 ' AHU1 Heating DA Setpoint
				 53, _
				 ' AHU2 Heating DA Setpoint
				 53
		ElseIf outsideTemp.value > 80 And outsideTemp.value < 85 Then
				  ' AHU1 RA Damper
			ScenarioA 0, _
				  ' AHU2 RA Damper
				  0, _
				  ' AHU1 OA Damper
				 30, _
				  ' AHU2 OA Damper
				 30, _
				  ' AHU1 Blower Speed
				 92, _
				  ' AHU2 Blower Speed
				 92, _
				  ' HV4 Damper
				 50, _
				 ' AHU1 RH Setpoint
				 55, _
				 ' AHU2 RH Setpoint
				 55, _
				 ' AHU1 Cooling DA Setpoint
				 55, _
				 ' AHU2 Cooling DA Setpoint
				 55, _
				 ' AHU1 Heating DA Setpoint
				 53, _
				 ' AHU2 Heating DA Setpoint
				 53
		ElseIf outsideTemp.value > 72 And outsideTemp.value < 80 Then
				  ' AHU1 RA Damper
			ScenarioA 0, _
				  ' AHU2 RA Damper
				  0, _
				  ' AHU1 OA Damper
				 30, _
				  ' AHU2 OA Damper
				 30, _
				  ' AHU1 Blower Speed
				 98, _
				  ' AHU2 Blower Speed
				 98, _
				  ' HV4 Damper
				 75, _
				 ' AHU1 RH Setpoint
				 55, _
				 ' AHU2 RH Setpoint
				 55, _
				 ' AHU1 Cooling DA Setpoint
				 55, _
				 ' AHU2 Cooling DA Setpoint
				 55, _
				 ' AHU1 Heating DA Setpoint
				 53, _
				 ' AHU2 Heating DA Setpoint
				 53
		ElseIf (outsideTemp.value < 72 And outsideTemp.value > 50) Or outsideTemp.value <  ahu1RAtemp.value Then
				  ' AHU1 RA Damper
			ScenarioA 30, _
				  ' AHU2 RA Damper
				  30, _
				  ' AHU1 OA Damper
				 30, _
				  ' AHU2 OA Damper
				 30, _
				  ' AHU1 Blower Speed
				 98, _
				  ' AHU2 Blower Speed
				 98, _
				  ' HV4 Damper
				 100, _
				 ' AHU1 RH Setpoint
				 50, _
				 ' AHU2 RH Setpoint
				 50, _
				 ' AHU1 Cooling DA Setpoint
				 55, _
				 ' AHU2 Cooling DA Setpoint
				 55, _
				 ' AHU1 Heating DA Setpoint
				 53, _
				 ' AHU2 Heating DA Setpoint
				 53
		ElseIf outsideTemp.value < 50 And outsideTemp.value > 42 Then
				  ' AHU1 RA Damper
			ScenarioA 100, _
				  ' AHU2 RA Damper
				  100, _
				  ' AHU1 OA Damper
				 100, _
				  ' AHU2 OA Damper
				 100, _
				  ' AHU1 Blower Speed
				 82, _
				  ' AHU2 Blower Speed
				 82, _
				  ' HV4 Damper
				 50, _
				 ' AHU1 RH Setpoint
				 50, _
				 ' AHU2 RH Setpoint
				 50, _
				 ' AHU1 Cooling DA Setpoint
				 55, _
				 ' AHU2 Cooling DA Setpoint
				 55, _
				 ' AHU1 Heating DA Setpoint
				 53, _
				 ' AHU2 Heating DA Setpoint
				 53
		ElseIf outsideTemp.value < 25 Then
				  ' AHU1 RA Damper
			ScenarioA 30, _
				  ' AHU2 RA Damper
				  30, _
				  ' AHU1 OA Damper
				 30, _
				  ' AHU2 OA Damper
				 30, _
				  ' AHU1 Blower Speed
				 85, _
				  ' AHU2 Blower Speed
				 85, _
				  ' HV4 Damper
				 50, _
				 ' AHU1 RH Setpoint
				 35, _
				 ' AHU2 RH Setpoint
				 35, _
				 ' AHU1 Cooling DA Setpoint
				 55, _
				 ' AHU2 Cooling DA Setpoint
				 55, _
				 ' AHU1 Heating DA Setpoint
				 53, _
				 ' AHU2 Heating DA Setpoint
				 53
		Else
			If hv4MAdamper.value <> 50 Then
				hv4MAdamper.value = 50
				
			End If
		End If
	End If
End If
End Sub