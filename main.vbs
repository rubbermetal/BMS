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
		' execute 2nd shift schedule
		'ScenarioA(raDamper1,raDamper2,oaDamper1,oaDamper2,blowerControl1,blowerControl2,hv4Damper,rhSp1,rhSp2,clgSp1,clgSp2,htgSp1,htgSp2)
		If outsideTemp.value > 85 Then
			ScenarioA 0, 0, 15, 15, 88, 88, 30, 58, 58, 55, 55, 53, 53
		ElseIf outsideTemp.value > 80 And outsideTemp.value < 85 Then
			ScenarioA 0, 0, 30, 30, 92, 92, 50, 55, 55, 55, 55, 53, 53
		ElseIf outsideTemp.value > 72 And outsideTemp.value < 80 Then
			ScenarioA 0, 0, 30, 30, 98, 98, 75, 50, 50, 55, 55, 53, 53
		ElseIf outsideTemp.value < 72 And outsideTemp.value > 50 Then
			ScenarioA 30, 30, 30, 30, 98, 98, 100, 55, 55, 55, 55, 53, 53
		ElseIf outsideTemp.value < 50 And outsideTemp.value > 42 Then
			ScenarioA 100, 100, 100, 100, 82, 82, 50, 50, 50, 55, 55, 53, 53
		ElseIf outsideTemp.value < 25 Then
			ScenarioA 0, 0, 10, 10, 85, 85, 20, 20, 20, 53, 53
		Else
			If hv4MAdamper.value <> 50 Then
				hv4MAdamper.value = 50
				
			End If
		End If
	ElseIf shift = 3 Then 
		' execute 3rd shift schedule
		'ScenarioA(raDamper1,raDamper2,oaDamper1,oaDamper2,blowerControl1,blowerControl2,hv4Damper,rhSp1,rhSp2,clgSp1,clgSp2,htgSp1,htgSp2)
		If outsideTemp.value > 85 Then
			ScenarioA 0, 0, 15, 15, 88, 88, 30, 58, 58, 55, 55, 53, 53
		ElseIf outsideTemp.value > 80 And outsideTemp.value < 85 Then
			ScenarioA 0, 0, 30, 30, 92, 92, 50, 55, 55, 55, 55, 53, 53
		ElseIf outsideTemp.value > 72 And outsideTemp.value < 80 Then
			ScenarioA 0, 0, 30, 30, 98, 98, 75, 50, 50, 55, 55, 53, 53
		ElseIf outsideTemp.value < 72 And outsideTemp.value > 50 Then
			ScenarioA 30, 30, 30, 30, 98, 98, 100, 55, 55, 55, 55, 53, 53
		ElseIf outsideTemp.value < 50 And outsideTemp.value > 42 Then
			ScenarioA 100, 100, 100, 100, 82, 82, 50, 50, 50, 55, 55, 53, 53
		ElseIf outsideTemp.value < 25 Then
			ScenarioA 0, 0, 10, 10, 85, 85, 20, 20, 20, 53, 53
		Else
			If hv4MAdamper.value <> 50 Then
				hv4MAdamper.value = 50
				
			End If
		End If
	Else
		' Shift 1
		' ScenarioA(raDamper1,raDamper2,oaDamper1,oaDamper2,blowerControl1,blowerControl2,hv4Damper,rhSp1,rhSp2,clgSp1,clgSp2,htgSp1,htgSp2)
		If outsideTemp.value > 85 Then
			ScenarioA 0, 0, 0, 0, 92, 92, 30, 58, 58, 55, 55, 53, 53
		ElseIf outsideTemp.value > 80 And outsideTemp.value < 85 Then
			ScenarioA 0, 0, 30, 30, 92, 92, 50, 55, 55, 57, 57, 53, 53
		ElseIf outsideTemp.value > 72 And outsideTemp.value < 80 Then
			ScenarioA 0, 0, 30, 30, 98, 98, 75, 50, 50, 57, 57, 53, 53
		ElseIf (outsideTemp.value < 72 And outsideTemp.value > 50) Or outsideTemp.value <  ahu1RAtemp.value Then
			ScenarioA 30, 30, 30, 30, 98, 98, 100, 65, 65, 57, 57, 53, 53
		ElseIf outsideTemp.value < 50 And outsideTemp.value > 42 Then
			ScenarioA 100, 100, 100, 100, 82, 82, 50, 50, 50, 55, 55, 53, 53
		ElseIf outsideTemp.value < 25 Then
			ScenarioA 0, 0, 10, 10, 85, 85, 20, 20, 20, 53, 53
		Else
			If hv4MAdamper.value <> 50 Then
				hv4MAdamper.value = 50
				
			End If
		End If
	End If
End If
End Sub