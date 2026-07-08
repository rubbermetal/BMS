Class ControlField
	Public value
End Class

Dim control, status, steamPressure
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
Function StaticPressureEast(timeCheck )
	Dim spSp, minSpeed, maxSpeed
	minSpeed = "83.0"
	maxSpeed = "99.0"
	If timeCheck = 30 Or timeCheck = 0 Then
			If alphanum10.value > eastBuldingStatic.value And ahu1BlowerControl.value < maxSpeed Then
				speed1 = ahu1BlowerControl.value + 1
			ElseIf alphanum10.value < eastBuldingStatic.value And ahu1BlowerControl.value > minSpeed Then
				speed1 = ahu1BlowerControl.value  - 1
			Else
				speed1 = ahu1BlowerControl.value
			End If
	Else
		speed1 = ahu1BlowerControl.value
	End If
	changeControl ahu1BlowerMode, ahu1BlowerControl, speed1, "manual"
End Function

Function StaticPressureWest(timeCheck )
	Dim spSp, minSpeed, maxSpeed
	minSpeed = "83.0"
	maxSpeed = "99.0"
	If timeCheck = 30 Or timeCheck = 0 Then
			If alphanum100.value > alphanum21.value And ahu2BlowerControl.value < maxSpeed Then
				speed2 = ahu2BlowerControl.value + 1
			ElseIf alphanum100.value < alphanum21.value And ahu2BlowerControl.value > minSpeed Then
				speed2 = ahu2BlowerControl.value  - 1
			Else
				speed2 = ahu2BlowerControl.value
			End If
	Else
		speed2 = ahu2BlowerControl.value
	End If
	changeControl ahu2BlowerMode, ahu2BlowerControl, speed2, "manual"
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
		'MsgBox "HV4 is in alarm, blower is offline."
		Set player = Nothing
		alarmSilence.value = True
		Exit Function
	ElseIf ahu2Alarm.value <> 3 And alarmSilence.value <> True Then
		Set player = CreateObject("WMPlayer.OCX")
		player.URL = "C:\Program Files\Honeywell\client\abstract\ahu2-blower-alarm.mp3"
		player.Controls.play
		'MsgBox "AHU-2 is in alarm, blower is offline."
		Set player = Nothing
		alarmSilence.value = True
		Exit Function
	ElseIf ahu1Alarm.value <> 3 And alarmSilence.value <> True Then
		Set player = CreateObject("WMPlayer.OCX")
		player.URL = "C:\Program Files\Honeywell\client\abstract\ahu1-blower-alarm.mp3"
		player.Controls.play
		'MsgBox "AHU-1 is in alarm, blower is offline."
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

' Dew point (deg F) from dry-bulb (deg F) and RH (percent) via Magnus-Tetens.
Function calculateDP(T,H)
	Dim Tc, a, b, gamma, TdC
	a = 17.27
	b = 237.7
	If H < 1 Then H = 1
	Tc = (T - 32) * 5 / 9
	gamma = Log(H / 100) + (a * Tc) / (b + Tc)
	TdC = (b * gamma) / (a - gamma)
	calculateDP = RoundToOneDecimal((TdC * 1.8) + 32)
End Function

Function RoundUp(number)
	roundedValue = Int(number + 0.999999999)
	RoundUp = CStr(roundedValue) & ".0"
End Function

Function numPad(number)
	numPad = FormatNumber(number, 1)
End Function
' Discharge setpoint with return-RH reset.
' Returns lowSp when the average return RH is below 55%, otherwise highSp.
Function dischargeSp(highSp, lowSp)
	Dim avgRH
	avgRH = CDbl(ahu2RH.value)
	If avgRH < 55 Then
		dischargeSp = lowSp
	Else
		dischargeSp = highSp
	End If
End Function
' Moist-air enthalpy in Btu/lb dry air, from dry-bulb temp (deg F) and
' relative humidity (percent). ASHRAE IP correlations, sea-level pressure.
Function enthalpyIP(tempF, rhPercent)
	Dim Patm, tRankine, lnPws, Pws, Pw, W
	Patm = 14.696
	tRankine = tempF + 459.67
	lnPws = -1.0440397E4 / tRankine
	lnPws = lnPws - 1.1294650E1
	lnPws = lnPws - 2.7022355E-2 * tRankine
	lnPws = lnPws + 1.2890360E-5 * tRankine ^ 2
	lnPws = lnPws - 2.4780681E-9 * tRankine ^ 3
	lnPws = lnPws + 6.5459673E0 * Log(tRankine)
	Pws = Exp(lnPws)
	Pw = (rhPercent / 100) * Pws
	W = 0.621945 * Pw / (Patm - Pw)
	enthalpyIP = 0.240 * tempF + W * (1061 + 0.444 * tempF)
End Function
' Saturation humidity ratio (lb water / lb dry air) at a dry-bulb temp (deg F).
Function satW(tempF)
	Dim Patm, tRankine, lnPws, Pws
	Patm = 14.696
	tRankine = tempF + 459.67
	lnPws = -1.0440397E4 / tRankine
	lnPws = lnPws - 1.1294650E1
	lnPws = lnPws - 2.7022355E-2 * tRankine
	lnPws = lnPws + 1.2890360E-5 * tRankine ^ 2
	lnPws = lnPws - 2.4780681E-9 * tRankine ^ 3
	lnPws = lnPws + 6.5459673E0 * Log(tRankine)
	Pws = Exp(lnPws)
	satW = 0.621945 * Pws / (Patm - Pws)
End Function
' Leaving-air enthalpy (Btu/lb) across a coil. Wet (leaving ~91% RH) when the
' discharge temp is below the entering dew point (satW at discharge < Win);
' otherwise dry, so leaving humidity stays at the entering Win (sensible only).
Function leavingEnthalpy(daTemp, Win)
	If satW(daTemp) < Win Then
		leavingEnthalpy = enthalpyIP(daTemp, 91)
	Else
		leavingEnthalpy = 0.270 * daTemp + Win * (1061 + 0.444 * daTemp)
	End If
End Function
' Total airside heat load (Btu/hr), computed per air handler then summed.
' Each unit: entering enthalpy from its own HtgDsch temp + ahu2RH, leaving by
' its own discharge temp (wet/dry judged independently). Each AHU = 55,000 CFM.
Function calculateLoad()
	Dim s1, s2, da1, da2, t1, t2, rh
	Dim cfm1, cfm2, hIn1, hIn2, Win1, Win2, hOut1, hOut2, load1, load2

	' --- Read raw point values (text) ---
	s1  = ahu1BlowerSpeed.value
	s2  = ahu2BlowerSpeed.value
	da1 = ahu1DA.value
	da2 = ahu2DA.value
	t1  = HtgDsch1.value
	t2  = HtgDsch2.value
	rh  = ahu2RH.value

	' --- Bail out safely if any point is not a valid number ---
	If Not (IsNumeric(s1) And IsNumeric(s2) And IsNumeric(da1) And IsNumeric(da2) And IsNumeric(t1) And IsNumeric(t2) And IsNumeric(rh)) Then
		calculateLoad = 0
		Exit Function
	End If

	' --- Airflow per unit. 51,000 = 55k design x 0.926, calibrated to waterside
	'     chiller tons: zeros the ~7% full-load overread. Light-load still reads
	'     high (enthalpy small-diff + CFM falls faster than VFD speed at part load). ---
	cfm1 = 51000 * (CDbl(s1) / 100)
	cfm2 = 51000 * (CDbl(s2) / 100)

	' --- AHU1: own entering enthalpy (HtgDsch1 + ahu2RH), own wet/dry ---
	hIn1 = enthalpyIP(CDbl(t1), CDbl(rh))
	Win1 = (hIn1 - 0.240 * CDbl(t1)) / (1061 + 0.444 * CDbl(t1))
	hOut1 = leavingEnthalpy(CDbl(da1), Win1)
	load1 = 4.5 * cfm1 * (hIn1 - hOut1)

	' --- AHU2: own entering enthalpy (HtgDsch2 + ahu2RH), own wet/dry ---
	hIn2 = enthalpyIP(CDbl(t2), CDbl(rh))
	Win2 = (hIn2 - 0.240 * CDbl(t2)) / (1061 + 0.444 * CDbl(t2))
	hOut2 = leavingEnthalpy(CDbl(da2), Win2)
	load2 = 4.5 * cfm2 * (hIn2 - hOut2)

	calculateLoad = Round(load1 + load2, 0)
End Function
' Cooling load as % of total (two-chiller, 500-ton) plant capacity.
' Derived from calculateLoad (actual leaving-air enthalpy and actual airflow).
' 500 tons = 6,000,000 Btu/hr, so % = load / 60000.
Function pctCapacity()
	pctCapacity = Round(calculateLoad() / 60000, 1)
End Function
' Mode + load descriptor for dataWindow, tiered by % of plant capacity.
Function coolingLabel()
	Dim pct
	pct = pctCapacity()
	If pct >= 65 Then
		coolingLabel = "Cooling - Extreme Load"
	ElseIf pct >= 55 Then
		coolingLabel = "Cooling - Heavy Load"
	ElseIf pct >= 40 Then
		coolingLabel = "Cooling - Normal Load"
	ElseIf pct >= 25 Then
		coolingLabel = "Cooling - Moderate Load"
	ElseIf pct >= 10 Then
		coolingLabel = "Cooling - Low Load"
	Else
		coolingLabel = "Cooling - Light Load"
	End If
End Function

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
			End If
		Else
			If control.value <> Setpoint Then
				control.value = Setpoint 
				control.linecolor = RGB(0,255,0)
			End If
		End If
	' If mode is "Auto", check to see if we are in "manual"
	' If we are in manual ,change to auto.
	Else 
		If status.value = True Then
			status.value = False
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
' Apply one cooling schedule row (replaces condA-condD and condF/H/J).
Sub applyCoolingRow(r)
	Dim hv4, sp
	If r(4) = -1 Then
		If outsideTemp.value > alphanum3.value Then
			hv4 = 50
		Else
			hv4 = 100
		End If
	Else
		hv4 = r(4)
	End If
	sp = dischargeSp(r(6), r(7))
	dataWindow.value = coolingLabel()
	ScenarioA r(1), r(1), r(2), r(2), r(3), r(3), hv4, r(5), r(5), sp, sp, sp, sp, True
End Sub
' Cooling schedule dispatch. Returns True if a row was applied; returns
' False ONLY for a chiller-OFF sub-72 band, so the caller runs free cooling.
' Row cols: lowBound, RAclosed, OAopen, fan, hv4 (-1 = condD rule:
' 50 if OAT>alphanum3 else 100), rhSp, dischHi, dischLo, chillerOnly,
' inclLowEdge (True = temp >= lowBound, False = temp > lowBound).
Function dispatchCooling(oat)
	Dim clgSched(15), i, r, reached
	dispatchCooling = False
	If Not IsNumeric(oat) Then Exit Function
	oat = CDbl(oat)
	clgSched(0)  = Array(92, 40, 45, 98,  50, 55, 55, 55, False, True)
	clgSched(1)  = Array(91, 41, 46, 98,  50, 55, 55, 55, False, True)
	clgSched(2)  = Array(90, 42, 47, 98,  50, 55, 55, 55, False, True)
	clgSched(3)  = Array(89, 43, 48, 98,  50, 55, 55, 55, False, True)
	clgSched(4)  = Array(88, 44, 49, 98,  50, 55, 55, 55, False, True)
	clgSched(5)  = Array(87, 45, 50, 98,  50, 55, 55, 55, False, True)
	clgSched(6)  = Array(86, 44, 51, 98,  50, 55, 55, 55, False, True)
	clgSched(7)  = Array(85, 43, 52, 98,  50, 55, 55, 55, False, True)
	clgSched(8)  = Array(84, 42, 53, 98,  50, 55, 55, 55, False, True)
	clgSched(9)  = Array(83, 41, 54, 98,  50, 55, 55, 55, False, True)
	clgSched(10) = Array(80, 40, 55, 98,  50, 55, 55, 55, False, True)
	clgSched(11) = Array(72, 45, 55, 98,  -1, 55, 55, 55, False, False)
	clgSched(12) = Array(62, 50, 50, 98, 100, 55, 55, 53, True,  False)
	clgSched(13) = Array(55, 45, 45, 98, 100, 55, 53, 50, True,  False)
	clgSched(14) = Array(50, 40, 40, 98,  50, 50, 55, 50, True,  False)
	clgSched(15) = Array(45, 40, 40, 95,  50, 50, 55, 50, True,  True)
	For i = 0 To UBound(clgSched)
		r = clgSched(i)
		If r(9) Then
			reached = (oat >= r(0))
		Else
			reached = (oat > r(0))
		End If
		If reached Then
			If r(8) And chillerStatus.value <> "ON" Then
				Exit Function
			End If
			applyCoolingRow r
			dispatchCooling = True
			Exit Function
		End If
	Next
End Function
Function condE ( )
				dataWindow.value = "Free Cooling Mode"
				raDamper1 = 95
				raDamper2 = 95
				oaDamper1 = 100
				oaDamper2 = 100
				blowerControl1 = 82
				blowerControl2 = 82
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
Function condG ( )
				dataWindow.value = "Free Cooling Mode"
				raDamper1 = 85
				raDamper2 = 85
				oaDamper1 = 100
				oaDamper2 = 100
				blowerControl1 = 85
				blowerControl2 = 85
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
Function condI ( )
				dataWindow.value = "Free Cooling Mode"
				raDamper1 = 65
				raDamper2 = 65
				oaDamper1 = 100
				oaDamper2 = 100
				If spMaintain.value = True Then
					Call StaticPressureEast(currentMinutes)
					Call StaticPressureWest(currentMinutes)
					nosp = False
				Else
					nosp = True
				End If
				blowerControl1 = 88
				blowerControl2 = 88
				hv4Damper = 50
				rhSp1 =  35
				rhSp2 = 35
				clgSp1 = 55
				clgSp2 = 55
				htgSp1 = 53
				htgSp2 = 53
				ScenarioA raDamper1,raDamper2,oaDamper1,oaDamper2,blowerControl1,blowerControl2,hv4Damper,rhSp1,rhSp2,clgSp1,clgSp2,htgSp1,htgSp2,nosp
End Function
Function condK ( )
				raDamper1 = 50
				raDamper2 = 50
				oaDamper1 = 100
				oaDamper2 = 100
				If spMaintain.value = True Then
					Call StaticPressureEast(currentMinutes)
					Call StaticPressureWest(currentMinutes)
					nosp = False
				Else
					nosp = True
				End If
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
				If spMaintain.value = True Then
					Call StaticPressureEast(currentMinutes)
					Call StaticPressureWest(currentMinutes)
					nosp = False
				Else
					nosp = True
				End If
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
				If spMaintain.value = True Then
					Call StaticPressureEast(currentMinutes)
					Call StaticPressureWest(currentMinutes)
					nosp = False
				Else
					nosp = True
				End If
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
				If spMaintain.value = True Then
					Call StaticPressureEast(currentMinutes)
					Call StaticPressureWest(currentMinutes)
					nosp = False
				Else
					nosp = True
				End If
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
				If spMaintain.value = True Then
					Call StaticPressureEast(currentMinutes)
					Call StaticPressureWest(currentMinutes)
					nosp = False
				Else
					nosp = True
				End If
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
				If spMaintain.value = True Then
					Call StaticPressureEast(currentMinutes)
					Call StaticPressureWest(currentMinutes)
					nosp = False
				Else
					nosp = True
				End If
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
				If spMaintain.value = True Then
					Call StaticPressureEast(currentMinutes)
					Call StaticPressureWest(currentMinutes)
					nosp = False
				Else
					nosp = True
				End If
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
				If spMaintain.value = True Then
					Call StaticPressureEast(currentMinutes)
					Call StaticPressureWest(currentMinutes)
					nosp = False
				Else
					nosp = True
				End If
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
Function condS ( )
				raDamper1 = 45
				raDamper2 = 45
				oaDamper1 = 45
				oaDamper2 = 45
				If spMaintain.value = True Then
					Call StaticPressureEast(currentMinutes)
					Call StaticPressureWest(currentMinutes)
					nosp = False
				Else
					nosp = True
				End If
				blowerControl1 = 84
				blowerControl2 = 84
				hv4Damper = 60
				rhSp1 = HumidityAdjust( outsideTemp.value )
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
				If spMaintain.value = True Then
					Call StaticPressureEast(currentMinutes)
					Call StaticPressureWest(currentMinutes)
					nosp = False
				Else
					nosp = True
				End If
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

'==============================================================================
' Chiller staging (lift-aware) & free-cooling helpers   (added 2026-07-08)
' Reuses the existing enthalpyIP(). Calibrated to actual chiller runtime: the
' 2nd-chiller flag matched 92% vs 85% for MA enthalpy alone. Adds the term the
' load-only pctCapacity() can't see: OA wet-bulb proxies condenser lift (cond
' water tracks wet-bulb + tower approach), and a chiller makes fewer tons as lift
' rises - so 2 chillers stage EARLIER on hot/humid days at the same load.
'   staging index = MA enthalpy + STAGE_K*(OA wet-bulb - 65)
' NEEDS A RELIABLE OA RH (or dewpoint). AHU1OaRh reads ~99% (bad) - feed a good
' source or the wet-bulb / OA-enthalpy / free-cooling values will be wrong.
'
' HOW TO CALL (wire the point reads / display yourself):
'   maE = enthalpyIP(HtgDsch2.value, ahu2RH.value)      ' MA enthalpy, BTU/lb
'   oaE = enthalpyIP(outsideTemp.value, oaRh.value)     ' OA enthalpy, BTU/lb
'   wb  = WetBulbF(outsideTemp.value, oaRh.value)       ' OA wet-bulb, deg F
'   n   = ChillerStage( StageIndex(maE, wb) )           ' 0=off, 1=one, 2=two
'   fc  = FreeCoolPct(maE, oaE)                          ' 0-100 % (-1 = n/a)
' (Load-based view stays: pctCapacity() > 50 ~ 2 chillers. This is the lift-aware
'  cross-check; fuse once re-fit against calculateLoad if wanted.)
'==============================================================================

Const STAGE_K   = 0.35
Const STAGE_ON  = 20
Const STAGE_2ND = 27.2

' Outdoor wet-bulb (deg F), Stull approximation = condenser-lift proxy.
'   Call: WetBulbF(outsideTemp.value, oaRh.value)
Function WetBulbF(tempF, rhPercent)
	Dim r, c, tw
	r = rhPercent
	If r < 1 Then r = 1
	If r > 100 Then r = 100
	c = (tempF - 32) * 5 / 9
	tw = c * Atn(0.151977 * Sqr(r + 8.313659)) + Atn(c + r) - Atn(r - 1.676331) _
		+ 0.00391838 * (r ^ 1.5) * Atn(0.023101 * r) - 4.686035
	WetBulbF = tw * 9 / 5 + 32
End Function

' Lift-aware staging index = MA enthalpy + STAGE_K*(wet-bulb - 65).
'   Call: StageIndex(maEnthalpy, oaWetBulb)
Function StageIndex(maE, wb)
	StageIndex = maE + STAGE_K * (wb - 65)
End Function

' Chiller stage from the index: 0 = off, 1 = one chiller, 2 = two chillers.
'   Call: ChillerStage( StageIndex(maE, wb) )
Function ChillerStage(idx)
	If idx < STAGE_ON Then
		ChillerStage = 0
	ElseIf idx < STAGE_2ND Then
		ChillerStage = 1
	Else
		ChillerStage = 2
	End If
End Function

' Free-cooling availability, 0-100 %  (returns -1 = n/a when chiller already off).
'   100 = OA cold/dry enough to hold without the chiller; 1-99 = partial trim only.
'   Call: FreeCoolPct(maEnthalpy, oaEnthalpy)
Function FreeCoolPct(maE, oaE)
	Dim denom, p
	If maE <= STAGE_ON Then
		FreeCoolPct = -1
		Exit Function
	End If
	denom = maE - STAGE_ON
	If denom <= 0 Then
		FreeCoolPct = 0
		Exit Function
	End If
	p = (maE - oaE) / denom * 100
	If p < 0 Then p = 0
	If p > 100 Then p = 100
	FreeCoolPct = p
End Function
