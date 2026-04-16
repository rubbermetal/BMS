Dim sleepX, control, status, steamPressure
steamPressure = alphanum11.value - 1

Function timeFetch ( )
	
	currentTime = Time
	timeNow.value = currentTime

End Function

Function dp ( )
	a = 17.27
	b = 237.7
	avgRaTemp = (((CDbl(Abs(ahu2RAtemp.value)) + CDbl(Abs(ahu1RAtemp.value))) / 2) - 32)  *  5 / 9
	avgRaRh = (CDbl(ahu1RH.value) + CDbl(ahu2RH.value)) / 2
	avgRaRhFormatted = avgRaRh / 100

	dp =  RoundToOneDecimal(((avgRaTemp  -  ((100 - avgRaRh) / 5)) * 1.8) + 32) 
End Function
Function chillerMode(  )
	Dim ra, sa, diff
	ra = (CDbl(ahu1MA.value) + CDbl(ahu2MA.value)) / 2
	sa = (CDbl(ahu1DA.value) + CDbl(ahu2DA.value)) / 2
	diff = CDbl(ra - sa)
	If diff > 3 Then
		If chillerStatus.value <> "ON" Then
			chillerStatus.value = "ON"
			chillerStatus.linecolor = RGB(0,255,0)
		End If
	Else 
		If chillerStatus.value <> "OFF" Then
			chillerStatus.value = "OFF"
			chillerStatus.linecolor = RGB(255,0,0)
		End If
	End If
End Function
Function HumidityAdjust ( t  )
	If t >= 45 And t < 50 Then
		x = 40
	ElseIf t >= 40 And t < 45 Then
		x = 37.5
	ElseIf t >= 35 And t < 40 Then
		x = 35
	ElseIf t >= 30 And t < 40 Then
		x = 32.5
	ElseIf t >= 25 And t < 30 Then
		x = 30
	ElseIf t >= 20 And t < 25 Then
		x = 27.5
	ElseIf t >= 15 And t < 20 Then
		x = 25
	ElseIf t >= 10 And t < 15 Then
		x = 22.5
	ElseIf t >= 5 And t < 10 Then
		x = 17.5
	ElseIf t >= 0 And t < 5 Then
		x = 15
	ElseIf t >= -10 And t < 0 Then
		x = 15
	ElseIf t >= -20 And t < -10 Then
		x = 10
	ElseIf t >= -30 And t < -20 Then
		x = 5
	ElseIf t >= -40 And t < -30 Then
		x = 5
	Else 
		x = 45
	End If
	HumidityAdjust = x
End Function

Function AlarmStatus(  )
 	Dim player
	Dim IsPlaying
	
	
	' hv4Alarm.value = 3 is normal, no alarm
	If hv4Alarm.value <> 3  And alarmSilence.value <> True Then
		Set player = CreateObject("WMPlayer.OCX")
		player.URL = "C:\Program Files\Honeywell\client\abstract\hv4-blower-alarm.mp3"
		player.Controls.play
		MsgBox "HV4 is in alarm, blower is offline."
		Set player = Nothing
		alarmSilence.value = True
		Exit Function
	ElseIf ahu2Alarm.value <> 3 And alarmSilence.value <> True Then
		Set player = CreateObject("WMPlayer.OCX")
		player.URL = "C:\Program Files\Honeywell\client\abstract\ahu2-blower-alarm.mp3"
		player.Controls.play
		MsgBox "AHU-2 is in alarm, blower is offline."
		Set player = Nothing
		alarmSilence.value = True
		Exit Function
	ElseIf ahu1Alarm.value <> 3 And alarmSilence.value <> True Then
		Set player = CreateObject("WMPlayer.OCX")
		player.URL = "C:\Program Files\Honeywell\client\abstract\ahu1-blower-alarm.mp3"
		player.Controls.play
		MsgBox "AHU-1 is in alarm, blower is offline."
		Set player = Nothing
		alarmSilence.value = True
		Exit Function
	End If
End Function

Function dupControl( )
	Dim ahu2HtHi, ahu2HtLo, ahu2HtLoMode, ahu2HtHiMode, ahu1HtHi, ahu1HtLo, ahu1HtHiMode, ahu1HtLoMode

	ahu2HtHi = numPad(ahu2HtSpHi.value)
	ahu2HtLo = numPad(ahu2HtSpLo.value)
	ahu2HtHiMode = ahu2HtSpHiMode.value
	ahu2HtLoMode = ahu2HtSpLoMode.value

	ahu1HtHi = numPad(ahu1HtSpHi.value)
	ahu1HtLo = numPad(ahu1HtSpLo.value)
	ahu1HtHiMode = ahu1HtSpHiMode.value
	ahu1HtLoMode = ahu1HtSpLoMode.value

	' If ahu1 Heat Hi set point is in auto then put Ahu2 in auto
	If ahu1HtSpHiMode.value <> True And ahu2HtSpHiMode.value <> False Then
		ahu2HtSpHiMode.value = False
	End If
	' If ahu1 Heat Lo set point is in auto then put Ahu2 in auto
	If ahu1HtSpLoMode.value <> True And ahu2HtSpLoMode.value <> False Then
		ahu2HtSpLoMode.value = False
	End If
	' If ahu1 Heat Hi set point is in Manual then put Ahu2 in Manual
	If ahu1HtSpHiMode.value <> False And ahu2HtSpHiMode.value <> True Then
		ahu2HtSpHiMode.value = True
	End If
	' If ahu1 Heat Lo set point is in Manual then put Ahu2 in Manual
	If ahu1HtSpLoMode.value <> False And ahu2HtSpLoMode.value <> True Then
		ahu2HtSpLoMode.value = True
	End If
	' If ahu2 heat hi set pont doesn't match ahu1 then update it to match
	If ahu2HtHi <> ahu1HtHi And ahu2HtHiMode <> False Then
		ahu2HtSpHi.value = ahu1HtSpHi.value
	End If
	' If ahu2 heat lo set pont doesn't match ahu1 then update it to match
	If ahu2HtLo <> ahu1HtLo And ahu2HtLoMode <> False Then
		ahu2HtSpLo.value = ahu1HtSpLo.value
	End If
End Function

Function RoundToOneDecimal(number)
	RoundToOneDecimal = Int(number * 10 + 0.5) / 10
End Function

Function calculateDP(T,H)
	Dim Tc
	Tc = (T - 32) * 5 / 9
	calculateDP = RoundToOneDecimal(((Tc  -  ((100 - H) / 5)) * 1.8) + 32)
End Function

Function RoundUp(number)
	roundedValue = Int(number + 0.999999999)
	RoundUp = CStr(roundedValue) & ".0"
End Function

Function numPad(number)
	numPad = FormatNumber(number, 1)
End Function
' Create a sleep function
Sub sleep(x)
	'WScript.Sleep(x)
End Sub

' Function to change control settings\
' status is the override checkbox
' control is the name of the control
' sp is the setting to change to
' mode is what mode we want, auto or manual
Sub changeControl(status, control, sp, mode)
	Dim Setpoint
	Setpoint = numPad(sp)
	If mode = "manual" Then
		If status.value <> True Then
			status.value = True
			If control.value <> Setpoint Then
				control.value = Setpoint 
				control.linecolor = RGB(0,255,0)
				sleep(sleepX)
			End If
		Else
			If control.value <> Setpoint Then
				control.value = Setpoint 
				control.linecolor = RGB(0,255,0)
				'sleep(sleepX)
			End If
		End If
	' If mode is "Auto", check to see if we are in "manual"
	' If we are in manual ,change to auto.
	Else 
		If status.value = True Then
			status.value = False
			sleep(sleepX)
		End If
	End If
End Sub
' Cooling scenario A
Sub ScenarioA(raDamper1,raDamper2,oaDamper1,oaDamper2,blowerControl1,blowerControl2,hv4Damper,rhSp1,rhSp2,clgSp1,clgSp2,htgSp1,htgSp2,nosp)
	changeControl ahu1RAdamperMode, ahu1RAdamper, raDamper1, "manual"
	changeControl ahu2RAdamperMode, ahu2RAdamper, raDamper2, "manual"
	changeControl ahu1OAdamperMode, ahu1OAdamper, oaDamper1, "manual"
	changeControl ahu2OAdamperMode, ahu2OAdamper, oaDamper2, "manual"
	If nosp = True Then
		changeControl ahu1BlowerMode, ahu1BlowerControl, blowerControl1, "manual"
		changeControl ahu2BlowerMode, ahu2BlowerControl, blowerControl2, "manual"
	End If
	changeControl hv4DamperMode, hv4MAdamper, hv4Damper, "manual"
	changeControl ahu1RhSpMode, ahu1RhSp, rhSp1, "manual"
	changeControl ahu2RhSpMode, ahu2RhSp, rhSp2, "manual"
	changeControl ahu1ClgSpMode, ahu1ClgSp, clgSp1, "manual"
	changeControl ahu2ClgSpMode, ahu2ClgSp, clgSp2, "manual"
	changeControl ahu1HtgSpMode, ahu1HtgSp, htgSp1, "manual"
	changeControl ahu2HtgSpMode, ahu2HtgSp, htgSp2, "manual"

End Sub
Function condA( )
			dataWindow.value = "Cooling - Extreme Load"
			raDamper1 = 40
			raDamper2 = 40
			oaDamper1 = 45
			oaDamper2 = 45
			blowerControl1 = 95
			blowerControl2 = 95
			If outsideTemp.value > alphanum3.value Then 
				hv4Damper = 50 
			}
			Else
				hv4Damper = 50
			End If
			rhSp1 = 55
			rhSp2 = 55
			clgSp1 = 55
			clgSp2 = 55
			htgSp1 = 55
			htgSp2 = 55
			nosp = True
			ScenarioA raDamper1,raDamper2,oaDamper1,oaDamper2,blowerControl1,blowerControl2,hv4Damper,rhSp1,rhSp2,clgSp1,clgSp2,htgSp1,htgSp2,nosp
End Function
Function condA_1( )
			dataWindow.value = "Cooling - Extreme Load"
			raDamper1 = 41
			raDamper2 = 41
			oaDamper1 = 46
			oaDamper2 = 46
			blowerControl1 = 98
			blowerControl2 = 98
			If outsideTemp.value > alphanum3.value Then 
				hv4Damper = 50 
			}
			Else
				hv4Damper = 50
			End If
			rhSp1 = 55
			rhSp2 = 55
			clgSp1 = 55
			clgSp2 = 55
			htgSp1 = 55
			htgSp2 = 55
			nosp = True
			ScenarioA raDamper1,raDamper2,oaDamper1,oaDamper2,blowerControl1,blowerControl2,hv4Damper,rhSp1,rhSp2,clgSp1,clgSp2,htgSp1,htgSp2,nosp
End Function
Function condA_2( )
			dataWindow.value = "Cooling - Extreme Load"
			raDamper1 = 42
			raDamper2 = 42
			oaDamper1 = 47
			oaDamper2 = 47
			blowerControl1 = 98
			blowerControl2 = 98
			If outsideTemp.value > alphanum3.value Then 
				hv4Damper = 50 
			}
			Else
				hv4Damper = 50
			End If
			rhSp1 = 55
			rhSp2 = 55
			clgSp1 = 55
			clgSp2 = 55
			htgSp1 = 55
			htgSp2 = 55
			nosp = True
			ScenarioA raDamper1,raDamper2,oaDamper1,oaDamper2,blowerControl1,blowerControl2,hv4Damper,rhSp1,rhSp2,clgSp1,clgSp2,htgSp1,htgSp2,nosp
End Function
Function condA_3( )
			dataWindow.value = "Cooling - Extreme Load"
			raDamper1 = 43
			raDamper2 = 43
			oaDamper1 = 48
			oaDamper2 = 48
			blowerControl1 = 98
			blowerControl2 = 98
			If outsideTemp.value > alphanum3.value Then 
				hv4Damper = 50 
			}
			Else
				hv4Damper = 50
			End If
			rhSp1 = 55
			rhSp2 = 55
			clgSp1 = 55
			clgSp2 = 55
			htgSp1 = 55
			htgSp2 = 55
			nosp = True
			ScenarioA raDamper1,raDamper2,oaDamper1,oaDamper2,blowerControl1,blowerControl2,hv4Damper,rhSp1,rhSp2,clgSp1,clgSp2,htgSp1,htgSp2,nosp
End Function
Function condA_4( )
			dataWindow.value = "Cooling - Extreme Load"
			raDamper1 = 44
			raDamper2 = 44
			oaDamper1 = 49
			oaDamper2 = 49
			blowerControl1 = 98
			blowerControl2 = 98
			If outsideTemp.value > alphanum3.value Then 
				hv4Damper = 50 
			}
			Else
				hv4Damper = 50
			End If
			rhSp1 = 55
			rhSp2 = 55
			clgSp1 = 55
			clgSp2 = 55
			htgSp1 = 55
			htgSp2 = 55
			nosp = True
			ScenarioA raDamper1,raDamper2,oaDamper1,oaDamper2,blowerControl1,blowerControl2,hv4Damper,rhSp1,rhSp2,clgSp1,clgSp2,htgSp1,htgSp2,nosp
End Function
Function condB ( )
			dataWindow.value = "Cooling - Heavy Load"
			raDamper1 = 45
			raDamper2 = 45
			oaDamper1 = 50
			oaDamper2 = 50
			blowerControl1 = 98
			blowerControl2 = 98
			If outsideTemp.value > alphanum3.value Then 
				hv4Damper = 50 
			}
			Else
				hv4Damper = 50
			End If
			rhSp1 = 55
			rhSp2 = 55
			clgSp1 = 55
			clgSp2 = 55
			htgSp1 = 55
			htgSp2 = 55
			nosp = True
			ScenarioA raDamper1,raDamper2,oaDamper1,oaDamper2,blowerControl1,blowerControl2,hv4Damper,rhSp1,rhSp2,clgSp1,clgSp2,htgSp1,htgSp2,nosp
End Function
Function condB_1 ( )
			dataWindow.value = "Cooling - Heavy Load"
			raDamper1 = 44
			raDamper2 = 44
			oaDamper1 = 51
			oaDamper2 = 51
			blowerControl1 = 98
			blowerControl2 = 98
			If outsideTemp.value > alphanum3.value Then 
				hv4Damper = 50 
			}
			Else
				hv4Damper = 50
			End If
			rhSp1 = 55
			rhSp2 = 55
			clgSp1 = 55
			clgSp2 = 55
			htgSp1 = 55
			htgSp2 = 55
			nosp = True
			ScenarioA raDamper1,raDamper2,oaDamper1,oaDamper2,blowerControl1,blowerControl2,hv4Damper,rhSp1,rhSp2,clgSp1,clgSp2,htgSp1,htgSp2,nosp
End Function
Function condB_2 ( )
			dataWindow.value = "Cooling - Heavy Load"
			raDamper1 = 43
			raDamper2 = 43
			oaDamper1 = 52
			oaDamper2 = 52
			blowerControl1 = 98
			blowerControl2 = 98
			If outsideTemp.value > alphanum3.value Then 
				hv4Damper = 50 
			}
			Else
				hv4Damper = 50
			End If
			rhSp1 = 55
			rhSp2 = 55
			clgSp1 = 55
			clgSp2 = 55
			htgSp1 = 55
			htgSp2 = 55
			nosp = True
			ScenarioA raDamper1,raDamper2,oaDamper1,oaDamper2,blowerControl1,blowerControl2,hv4Damper,rhSp1,rhSp2,clgSp1,clgSp2,htgSp1,htgSp2,nosp
End Function
Function condB_3 ( )
			dataWindow.value = "Cooling - Heavy Load"
			raDamper1 = 42
			raDamper2 = 42
			oaDamper1 = 53
			oaDamper2 = 53
			blowerControl1 = 98
			blowerControl2 = 98
			If outsideTemp.value > alphanum3.value Then 
				hv4Damper = 50 
			}
			Else
				hv4Damper = 50
			End If
			rhSp1 = 55
			rhSp2 = 55
			clgSp1 = 55
			clgSp2 = 55
			htgSp1 = 55
			htgSp2 = 55
			nosp = True
			ScenarioA raDamper1,raDamper2,oaDamper1,oaDamper2,blowerControl1,blowerControl2,hv4Damper,rhSp1,rhSp2,clgSp1,clgSp2,htgSp1,htgSp2,nosp
End Function
Function condB_4 ( )
			dataWindow.value = "Cooling - Heavy Load"
			raDamper1 = 41
			raDamper2 = 41
			oaDamper1 = 54
			oaDamper2 = 54
			blowerControl1 = 98
			blowerControl2 = 98
			If outsideTemp.value > alphanum3.value Then 
				hv4Damper = 50 
			}
			Else
				hv4Damper = 50
			End If
			rhSp1 = 55
			rhSp2 = 55
			clgSp1 = 55
			clgSp2 = 55
			htgSp1 = 55
			htgSp2 = 55
			nosp = True
			ScenarioA raDamper1,raDamper2,oaDamper1,oaDamper2,blowerControl1,blowerControl2,hv4Damper,rhSp1,rhSp2,clgSp1,clgSp2,htgSp1,htgSp2,nosp
End Function
Function condC ( )
			dataWindow.value = "Cooling Mode"
			raDamper1 = 40
			raDamper2 = 40
			oaDamper1 = 55
			oaDamper2 = 55
			blowerControl1 = 98
			blowerControl2 = 98
			If outsideTemp.value > alphanum3.value Then 
				hv4Damper = 50 
			}
			Else
				hv4Damper = 50
			End If
			rhSp1 = 55
			rhSp2 = 55
			clgSp1 = 55
			clgSp2 = 55
			htgSp1 = 55
			htgSp2 = 55
			nosp = True
			ScenarioA raDamper1,raDamper2,oaDamper1,oaDamper2,blowerControl1,blowerControl2,hv4Damper,rhSp1,rhSp2,clgSp1,clgSp2,htgSp1,htgSp2,nosp
End Function
Function condD ( )
			dataWindow.value = "Cooling Mode"
			raDamper1 = 45
			raDamper2 = 45
			oaDamper1 = 55
			oaDamper2 = 55
			blowerControl1 = 98
			blowerControl2 = 98
			If outsideTemp.value > alphanum3.value Then 
				hv4Damper = 50 
			}
			Else
				hv4Damper = 100
			End If
			rhSp1 = 55
			rhSp2 = 55
			clgSp1 = 55
			clgSp2 = 55
			htgSp1 = 55
			htgSp2 = 55
			nosp = True
			ScenarioA raDamper1,raDamper2,oaDamper1,oaDamper2,blowerControl1,blowerControl2,hv4Damper,rhSp1,rhSp2,clgSp1,clgSp2,htgSp1,htgSp2,nosp
End Function
Function condE ( )
				dataWindow.value = "Free Cooling Mode"
				raDamper1 = 95
				raDamper2 = 95
				oaDamper1 = 100
				oaDamper2 = 100
				blowerControl1 = 98
				blowerControl2 = 98
				hv4Damper = 100
				rhSp1 = 35
				rhSp2 = 35
				clgSp1 = 55
				clgSp2 = 55
				htgSp1 = 55
				htgSp2 = 55
				nosp = True
				ScenarioA raDamper1,raDamper2,oaDamper1,oaDamper2,blowerControl1,blowerControl2,hv4Damper,rhSp1,rhSp2,clgSp1,clgSp2,htgSp1,htgSp2,nosp
End Function
Function condF ( )
				dataWindow.value = "Cooling - Low Load"
				raDamper1 = 50
				raDamper2 = 50
				oaDamper1 = 55
				oaDamper2 = 55
				blowerControl1 = 98
				blowerControl2 = 98
				hv4Damper = 100
				rhSp1 = 55
				rhSp2 = 55
				clgSp1 = 55
				clgSp2 = 55
				htgSp1 = 55
				htgSp2 = 55
				nosp = True
				ScenarioA raDamper1,raDamper2,oaDamper1,oaDamper2,blowerControl1,blowerControl2,hv4Damper,rhSp1,rhSp2,clgSp1,clgSp2,htgSp1,htgSp2,nosp
End Function
Function condG ( )
				dataWindow.value = "Free Cooling Mode"
				raDamper1 = 85
				raDamper2 = 85
				oaDamper1 = 100
				oaDamper2 = 100
				blowerControl1 = 88
				blowerControl2 = 88
				hv4Damper = 50
				rhSp1 = 35
				rhSp2 = 35
				clgSp1 = 53
				clgSp2 = 53
				htgSp1 = 53
				htgSp2 = 53
				nosp = True
				ScenarioA raDamper1,raDamper2,oaDamper1,oaDamper2,blowerControl1,blowerControl2,hv4Damper,rhSp1,rhSp2,clgSp1,clgSp2,htgSp1,htgSp2,nosp
End Function
Function condH ( )
				dataWindow.value = "Cooling -  Extreme Low load"
				raDamper1 = 40
				raDamper2 = 40
				oaDamper1 = 40
				oaDamper2 = 40
				blowerControl1 = 98
				blowerControl2 = 98
				hv4Damper = 50
				rhSp1 = 50
				rhSp2 = 50
				clgSp1 = 55
				clgSp2 = 55
				htgSp1 = 53
				htgSp2 = 53
				nosp = True
				ScenarioA raDamper1,raDamper2,oaDamper1,oaDamper2,blowerControl1,blowerControl2,hv4Damper,rhSp1,rhSp2,clgSp1,clgSp2,htgSp1,htgSp2,nosp
End Function
Function condI ( )
				dataWindow.value = "Free Cooling Mode"
				raDamper1 = 65
				raDamper2 = 65
				oaDamper1 = 100
				oaDamper2 = 100
				nosp = True
				blowerControl1 = 82
				blowerControl2 = 82
				hv4Damper = 50
				rhSp1 =  35
				rhSp2 = 35
				clgSp1 = 55
				clgSp2 = 55
				htgSp1 = 53
				htgSp2 = 53
				ScenarioA raDamper1,raDamper2,oaDamper1,oaDamper2,blowerControl1,blowerControl2,hv4Damper,rhSp1,rhSp2,clgSp1,clgSp2,htgSp1,htgSp2,nosp
End Function
Function condJ ( )
				dataWindow.value = "Cooling -  Extreme Low load"
				raDamper1 = 40
				raDamper2 = 40
				oaDamper1 = 40
				oaDamper2 = 40
				blowerControl1 = 95
				blowerControl2 = 95
				hv4Damper = 50
				rhSp1 = 50
				rhSp2 = 50
				clgSp1 = 55
				clgSp2 = 55
				htgSp1 = 53
				htgSp2 = 53
				nosp = True
				ScenarioA raDamper1,raDamper2,oaDamper1,oaDamper2,blowerControl1,blowerControl2,hv4Damper,rhSp1,rhSp2,clgSp1,clgSp2,htgSp1,htgSp2,nosp
End Function
Function condK ( )
				raDamper1 = 50
				raDamper2 = 50
				oaDamper1 = 100
				oaDamper2 = 100
				nosp = True
				blowerControl1 = 82
				blowerControl2 = 82
				hv4Damper = 50
				rhSp1 =  HumidityAdjust( outsideTemp.value )
				rhSp2 = HumidityAdjust( outsideTemp.value )
				clgSp1 = 55
				clgSp2 = 55
				htgSp1 = 55
				htgSp2 = 55
				ScenarioA raDamper1,raDamper2,oaDamper1,oaDamper2,blowerControl1,blowerControl2,hv4Damper,rhSp1,rhSp2,clgSp1,clgSp2,htgSp1,htgSp2,nosp
End Function
Function condL ( )
				raDamper1 = 50
				raDamper2 = 50
				oaDamper1 = 70
				oaDamper2 = 70
				nosp = True
				blowerControl1 = 82
				blowerControl2 = 82
				hv4Damper = 60
				rhSp1 =  HumidityAdjust( outsideTemp.value )
				rhSp2 = HumidityAdjust( outsideTemp.value )
				clgSp1 = 55
				clgSp2 = 55
				htgSp1 = 55
				htgSp2 = 55
				ScenarioA raDamper1,raDamper2,oaDamper1,oaDamper2,blowerControl1,blowerControl2,hv4Damper,rhSp1,rhSp2,clgSp1,clgSp2,htgSp1,htgSp2,nosp
End Function
Function condM ( )
				raDamper1 = 50
				raDamper2 = 50
				oaDamper1 = 65
				oaDamper2 = 65
				nosp = True
				blowerControl1 = 82
				blowerControl2 = 82
				hv4Damper = 60
				rhSp1 =  HumidityAdjust( outsideTemp.value )
				rhSp2 = HumidityAdjust( outsideTemp.value )
				clgSp1 = 55
				clgSp2 = 55
				htgSp1 = 58
				htgSp2 = 58
				ScenarioA raDamper1,raDamper2,oaDamper1,oaDamper2,blowerControl1,blowerControl2,hv4Damper,rhSp1,rhSp2,clgSp1,clgSp2,htgSp1,htgSp2,nosp
End Function
Function condN ( )
				raDamper1 = 50
				raDamper2 = 50
				oaDamper1 = 60
				oaDamper2 = 60
				nosp = True
				blowerControl1 = 88
				blowerControl2 = 88
				hv4Damper = 60
				rhSp1 =  HumidityAdjust( outsideTemp.value )
				rhSp2 = HumidityAdjust( outsideTemp.value )
				clgSp1 = 55
				clgSp2 = 55
				htgSp1 = 60
				htgSp2 = 60
				ScenarioA raDamper1,raDamper2,oaDamper1,oaDamper2,blowerControl1,blowerControl2,hv4Damper,rhSp1,rhSp2,clgSp1,clgSp2,htgSp1,htgSp2,nosp
End Function
Function condO ( )
				raDamper1 = 50
				raDamper2 = 50
				oaDamper1 = 55
				oaDamper2 = 55
				nosp = True
				blowerControl1 = 88
				blowerControl2 = 88
				hv4Damper = 60
				rhSp1 =  HumidityAdjust( outsideTemp.value )
				rhSp2 = HumidityAdjust( outsideTemp.value )
				clgSp1 = 55
				clgSp2 = 55
				htgSp1 = 60
				htgSp2 = 60
				ScenarioA raDamper1,raDamper2,oaDamper1,oaDamper2,blowerControl1,blowerControl2,hv4Damper,rhSp1,rhSp2,clgSp1,clgSp2,htgSp1,htgSp2,nosp
End Function
Function condP ( )
				raDamper1 = 50
				raDamper2 = 50
				oaDamper1 = 50
				oaDamper2 = 50
				nosp = True
				blowerControl1 = 86
				blowerControl2 = 86
				hv4Damper = 60
				rhSp1 =  HumidityAdjust( outsideTemp.value )
				rhSp2 = HumidityAdjust( outsideTemp.value )
				clgSp1 = 55
				clgSp2 = 55
				htgSp1 = 60
				htgSp2 = 60
				ScenarioA raDamper1,raDamper2,oaDamper1,oaDamper2,blowerControl1,blowerControl2,hv4Damper,rhSp1,rhSp2,clgSp1,clgSp2,htgSp1,htgSp2,nosp
End Function
Function condQ ( )
				raDamper1 = 45
				raDamper2 =45
				oaDamper1 = 45
				oaDamper2 = 45
				nosp = True
				blowerControl1 = 84
				blowerControl2 = 84
				hv4Damper = 60
				rhSp1 =  HumidityAdjust( outsideTemp.value )
				rhSp2 = HumidityAdjust( outsideTemp.value )
				clgSp1 = 55
				clgSp2 = 55
				htgSp1 = 60
				htgSp2 = 60
				ScenarioA raDamper1,raDamper2,oaDamper1,oaDamper2,blowerControl1,blowerControl2,hv4Damper,rhSp1,rhSp2,clgSp1,clgSp2,htgSp1,htgSp2,nosp
End Function
Function condR ( )
				raDamper1 = 45
				raDamper2 = 45
				oaDamper1 = 45
				oaDamper2 = 45
				nosp = True
				blowerControl1 = 84
				blowerControl2 = 84
				hv4Damper = 60
				rhSp1 =  HumidityAdjust( outsideTemp.value )
				rhSp2 = HumidityAdjust( outsideTemp.value )
				clgSp1 = 55
				clgSp2 = 55
				htgSp1 = 60
				htgSp2 = 60
				ScenarioA raDamper1,raDamper2,oaDamper1,oaDamper2,blowerControl1,blowerControl2,hv4Damper,rhSp1,rhSp2,clgSp1,clgSp2,htgSp1,htgSp2,nosp
End Function
Function condT ( )
			raDamper1 = 45
			raDamper2 = 45
			oaDamper1 = 45
			oaDamper2 = 45
			blowerControl1 = 84
			blowerControl2 = 84
			hv4Damper = 60
			rhSp1 = HumidityAdjust( outsideTemp.value )
			rhSp2 = HumidityAdjust( outsideTemp.value )
			clgSp1 = 55
			clgSp2 = 55
			htgSp1 = 60
			htgSp2 = 60
			nosp = True
			ScenarioA raDamper1,raDamper2,oaDamper1,oaDamper2,blowerControl1,blowerControl2,hv4Damper,rhSp1,rhSp2,clgSp1,clgSp2,htgSp1,htgSp2,nosp
End Function
