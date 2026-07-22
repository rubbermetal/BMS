Sub  outsideTemp_OnLoad(  )
 Dim IsSelected
Dim currentTime, currentHour, amOrPm, shift
Dim dp, dp1, dp2
Dim avgRaTemp
Dim avgRaRh, a, b, alpha
Dim raDamper1,raDamper2,oaDamper1,oaDamper2,blowerControl1,blowerControl2,hv4Damper,rhSp1,rhSp2,clgSp1,clgSp2,htgSp1,htgSp2, speed1, speed2


FormatDateTime currentTime, vbLongTime
currentTime = Time
currentHour = Hour(currentTime)
currentMinutes = Minute(currentTime)
currentSeconds = Second(currentTime)
amOrPm =currentTime

enthMaintain.value = true
maEnthSp.value = 26.2

IsSelected = copilot.value
If (currentHour >= 7) And (currentHour <15) Then
	shift = 2
ElseIf (currentHour >= 15) And (currentHour < 23) Then
	shift = 3
Else
	shift = 1
End If
currentShift.value = shift
'Call AlarmStatus()
Call chillerMode()

avgRaTemp = (((CDbl(Abs(ahu2RAtemp.value)) + CDbl(Abs(ahu1RAtemp.value))) / 2))
avgRaRh = CDbl(ahu2RH.value)

DewPoint.value = calculateDP(avgRaTemp, avgRaRh )

Dim TempEnth, avgRH
	TempEnth = (CDbl(HtgDsch1.value) + CDbl(HtgDsch2.value)) / 2
	avgRH = CDbl(ahu2RH.value)   ' ahu2RH only - ahu1RH reads high (boiler-room humidity)
	maEnth.value = Round(enthalpyIP(TempEnth, avgRH), 2)

If IsSelected = True Then

		If outsideTemp.value >= 45 Then
			If Not dispatchCooling(outsideTemp.value) Then
				dispatchFreeCooling outsideTemp.value
			End If
		ElseIf outsideTemp.value =< 45 Then
			dispatchHeating outsideTemp.value
		Else
			If hv4MAdamper.value <> 60 Then
				hv4MAdamper.value = 60
			End If
		End If
End If
End Sub
