Sub outsideTemp_OnUpdate(  )
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

If IsSelected = True Then

		If outsideTemp.value => 92 Then
			Call condA()
		ElseIf outsideTemp.value => 91 And outsideTemp.value < 92 Then
			Call condA_1()
		ElseIf outsideTemp.value => 90 And outsideTemp.value < 91 Then
			Call condA_2
		ElseIf outsideTemp.value => 89 And outsideTemp.value < 90 Then
			Call condA_3
		ElseIf outsideTemp.value => 88 And outsideTemp.value < 89 Then
			Call condA_4
		ElseIf outsideTemp.value => 87 And outsideTemp.value < 88 Then
			Call condB()
		ElseIf outsideTemp.value => 86 And outsideTemp.value < 87 Then
			Call condB_1()
		ElseIf outsideTemp.value => 85 And outsideTemp.value < 86 Then
			Call condB_2()
		ElseIf outsideTemp.value => 84 And outsideTemp.value < 85 Then
			Call condB_3()
		ElseIf outsideTemp.value => 83 And outsideTemp.value < 84 Then
			Call condB_4()
		ElseIf outsideTemp.value => 80 And outsideTemp.value < 83 Then
			Call condC()
		ElseIf outsideTemp.value > 72 And outsideTemp.value < 80 Then
			Call condD()
		ElseIf outsideTemp.value =< 72 And outsideTemp.value > 55 Then
			If chillerStatus.value = "OFF" Then
				Call condE()
			Else
				Call condF()
			End If
		ElseIf outsideTemp.value =< 55 And outsideTemp.value > 50 Then
			If chillerStatus.value = "OFF" Then
				Call condG()
			Else
				Call condH()
			End If
		ElseIf outsideTemp.value =< 50 And outsideTemp.value >= 45 Then
			If chillerStatus.value = "OFF" Then
				Call condI()
			Else
				Call condJ()
			End If
		ElseIf outsideTemp.value =< 45 And outsideTemp.value > 0 Then
			If outsideTemp.value =< 45 And outsideTemp.value > 40 Then
				Call condK()
			ElseIf outsideTemp.value  =< 40 And outsideTemp.value > 35 Then
				Call condL()
			ElseIf outsideTemp.value  =< 35 And outsideTemp.value > 30 Then
				Call condM()
			ElseIf outsideTemp.value  =< 30 And outsideTemp.value > 25 Then
				Call condN()
			ElseIf outsideTemp.value =< 25 And outsideTemp.value > 20 Then
				Call condO()
			ElseIf outsideTemp.value =< 20 And outsideTemp.value > 15 Then
				Call condP()
			ElseIf outsideTemp.value =< 15 And outsideTemp.value > 10 Then
				Call condQ()
			ElseIf outsideTemp.value =< 10 And outsideTemp.value >= 0 Then
				Call condR()
			End If
			dataWindow.value = "Normal Heating Mode"
		ElseIf outsideTemp.value =< 0 Then
			Call condT()
		Else
			If hv4MAdamper.value <> 60 Then
				hv4MAdamper.value = 60
			End If
		End If
End If
End Sub
