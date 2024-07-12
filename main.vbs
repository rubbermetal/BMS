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
	' This block is for occupied settings
	If (shift = 2) Or (coolDown.value = True) Then
		' Co-pilot has been selected.
		' execute 2nd shift schedule
		' Begin HV-4 Co-pilot sequence
		If outsideTemp.value > 85 Then
			ScenarioA(0, 0, 15, 15, 88, 88, 30, 58, 58)
		ElseIf outsideTemp.value > 80 And outsideTemp.value < 85 Then
			ScenarioA(0, 0, 30, 30, 92, 92, 50, 55, 55)
		ElseIf outsideTemp.value > 72 And outsideTemp.value < 80 Then
			'ScenarioA(raDamper1,raDamper2,oaDamper1,oaDamper2,blowerControl1,blowerControl2,hv4Damper,rhSp1,rhSp2)
			ScenarioA(0, 0, 30, 30, 98, 98, 75, 50, 50)
		ElseIf outsideTemp.value < 72 And outsideTemp.value > 50 Then
			ScenarioA(30, 30, 30, 30, 98, 98, 100, 55, 55)
		ElseIf outsideTemp.value < 50 And outsideTemp.value > 42
			ScenarioA(100, 100, 100, 100, 82, 82, 50, 50, 50)
		ElseIf outsideTemp.value < 25 Then
			ScenarioA(0, 0, 10, 10, 85, 85, 20, 20, 20)	
		Else
			If hv4MAdamper.value <> 50 Then
				hv4MAdamper.value = 50
				
			End If
		End If
	ElseIf (shift = 3) Then 
		' execute 3rd shift schedule
		If hv4DamperMode.value <> True Then
			hv4DamperMode.value= True
			
		End If
	        If ahu1ClgSpMode.value <> True Then
			ahu1ClgSpMode.value = True
			
		End If
		 If ahu1RhSpMode.value <> True Then
			ahu1RhSpMode.value = True
			
		End If
		If ahu1RAdamperMode.value <> True Then
			ahu1RAdamperMode.value = True
			
		End If
		If ahu1OAdamperMode.value <> True Then
			ahu1OAdamperMode.value = True
			
		End If
		If ahu1BlowerMode.value <> True Then
			ahu1BlowerMode.value = True
			
		End If
		If ahu2ClgSpMode.value <> True Then
			ahu2ClgSpMode.value = True
			
		End If
		If ahu2RhSpMode.value <> True Then
			ahu2RhSpMode.value = True
			
		End If
		If ahu2OAdamperMode.value <> True Then
			ahu2OAdamperMode.value = True
			
		End If
		If ahu2RAdamperMode.value <> True Then
			ahu2RAdamperMode.value = True
			
		End If
		If ahu2BlowerMode.value <> True Then
			ahu2BlowerMode.value = True
			
		End If
		If ((outsideTemp.value > 60) And (outsideTemp.value < 80)) Then
			'Set back HV-4
			If hv4MAdamper.value <> 100 Then
				hv4MAdamper.value = 100
				
			End If
			' Set back AHU-1 based upon if the chiller is running
			If ahu1ClgEn.value = "Enable" Then
				' Increase AHU-1 DA set point
				' dp1 =  RoundUp(calculateDP(ahu1RAtemp.value, ahu1RH.value))
				dp1 = RoundUp(60)
				If (ahu1ClgSp.value <> dp1) Then
					ahu1ClgSp.value = dp1
					
				End If
				' Increase AHU-1  RH setpoint
				If ahu1RhSp.value <> 75 Then
					ahu1RhSp.value = 75
					
				End If
				If ahu1RAdamper.value <> 0 Then
					ahu1RAdamper.value = 0
					
				End If
				If ahu1OAdampers.value <> 25 Then
					ahu1OAdampers.value = 25
					
				End If
				If ahu1BlowerControl.value <> 85 Then
					ahu1BlowerControl.value = 85
					
				End If
			End If
			' Set back AHU-2 based upon if the chiller is running
			If ahu2ClgEn.value = "Enable" Then
				' Increase AHU-2 DA set point
				' dp2 =  RoundUp(calculateDP(ahu2RAtemp.value, ahu2RH.value))
				dp2 = RoundUp(60)
				
				If (ahu2ClgSp.value <> dp2)  Then
					ahu2ClgSp.value = dp2
					
				End If
				' Increase AHU-2  RH setpoint
				If ahu2RhSp.value <> 75 Then
					ahu2RhSp.value = 75
					
				End If
				If ahu2RAdamper.value <> 0 Then
					ahu2RAdamper.value = 0
					
				End If
				If ahu2OAdamper.value <> 25 Then
					ahu2OAdamper.value = 25
					
				End If
				If ahu2BlowerControl.value <> 85 Then
					ahu2BlowerControl.value = 85
					
				End If
			End If
		' Run the economizer
		ElseIf ((outsideTemp.value < 60) And (outsideTemp.value > 45)) Then
			' Set back HV-4
			If hv4MAdamper.value <> 100 Then
				hv4MAdamper.value = 100
				
			End If
			' Set back AHU-1 based upon if the chiller is running
			If ahu1RAdamper.value <> 0 Then
				ahu1RAdamper.value = 0
				
			End If
			If ahu1OAdampers.value <> 25 Then
				ahu1OAdampers.value = 25
				
			End If
			If ahu1BlowerControl.value <> 85 Then
				ahu1BlowerControl.value = 85
			End If
		End If
	Else
		' Execute 1st shift schedule
		If hv4DamperMode.value <> True Then
			hv4DamperMode.value= True
			
		End If
	        If ahu1ClgSpMode.value <> True Then
			ahu1ClgSpMode.value = True
			
		End If
		 If ahu1RhSpMode.value <> True Then
			ahu1RhSpMode.value = True
			
		End If
		If ahu1RAdamperMode.value <> True Then
			ahu1RAdamperMode.value = True
			
		End If
		If ahu1OAdamperMode.value <> True Then
			ahu1OAdamperMode.value = True
			
		End If
		If ahu1BlowerMode.value <> True Then
			ahu1BlowerMode.value = True
			
		End If
		If ahu2ClgSpMode.value <> True Then
			ahu2ClgSpMode.value = True
			
		End If
		If ahu2RhSpMode.value <> True Then
			ahu2RhSpMode.value = True
			
		End If
		If ahu2OAdamperMode.value <> True Then
			ahu2OAdamperMode.value = True
			
		End If
		If ahu2RAdamperMode.value <> True Then
			ahu2RAdamperMode.value = True
			
		End If
		If ahu2BlowerMode.value <> True Then
			ahu2BlowerMode.value = True
			
		End If

	End If
End If
End Sub
