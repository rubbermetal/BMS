Class ControlField
	Public value
End Class

Dim control, status, steamPressure, skipDampers
steamPressure = alphanum11.value - 1

Function timeFetch ( )
	
	currentTime = Time
	timeNow.value = currentTime

End Function

' Return-air dew point (deg F) for the DewPoint text box. Thin wrapper over
' calculateDP (Magnus-Tetens) with the same inputs the handlers use:
' ahu2RH only - ahu1RH reads high (boiler-room humidity).
Function dp ( )
	Dim avgRaTemp
	avgRaTemp = (CDbl(Abs(ahu2RAtemp.value)) + CDbl(Abs(ahu1RAtemp.value))) / 2
	dp = calculateDP(avgRaTemp, CDbl(ahu2RH.value))
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
' its own discharge temp (wet/dry judged independently). Effective 51,000 CFM/AHU.
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
	' skipDampers is set True by applyCoolingRow when the enthalpy trim
	' loops own the RA/OA dampers; every other caller gets scheduled positions.
	If skipDampers <> True Then
		changeControl ahu1RAdamperMode, ahu1RAdamper, raDamper1, "manual"
		changeControl ahu2RAdamperMode, ahu2RAdamper, raDamper2, "manual"
		changeControl ahu1OAdamperMode, ahu1OAdamper, oaDamper1, "manual"
		changeControl ahu2OAdamperMode, ahu2OAdamper, oaDamper2, "manual"
	End If
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
	skipDampers = (enthMaintain.value = True)
	ScenarioA r(1), r(1), r(2), r(2), r(3), r(3), hv4, r(5), r(5), sp, sp, sp, sp, True
	If skipDampers = True Then
		Call EnthalpyTrim1()
		Call EnthalpyTrim2()
	End If
	skipDampers = False
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
' Apply one free-cooling schedule row (replaces condE/condG/condI).
Sub applyFreeCoolRow(r)
	Dim nosp
	If r(8) Then
		nosp = spMaintainNosp()
	Else
		nosp = True
	End If
	dataWindow.value = "Free Cooling Mode"
	skipDampers = False
	ScenarioA r(1), r(1), r(2), r(2), r(3), r(3), r(4), r(5), r(5), r(6), r(6), r(7), r(7), nosp
End Sub
' Free-cooling schedule dispatch (replaces condE/condG/condI). Only reached
' when dispatchCooling returned False (chiller OFF, sub-72 band) and OAT is
' >= 45; the last row is the unconditional catch-all, same band edges as the
' old condE (>55) / condG (>50) / else condI ladder.
' Row cols: lowBound, RAdamper, OAdamper, fan, hv4, rhSp, clgSp, htgSp,
' spAllowed (True = spMaintain checkbox may hand the blowers to the
' static-pressure loops).
Function dispatchFreeCooling(oat)
	Dim fcSched(2), i, r
	dispatchFreeCooling = False
	If Not IsNumeric(oat) Then Exit Function
	oat = CDbl(oat)
	fcSched(0) = Array(55, 95, 100, 82, 100, 35, 55, 55, False)
	fcSched(1) = Array(50, 85, 100, 85,  50, 35, 53, 53, False)
	fcSched(2) = Array(50, 65, 100, 88,  50, 35, 55, 53, True)
	For i = 0 To UBound(fcSched)
		r = fcSched(i)
		If oat > r(0) Or i = UBound(fcSched) Then
			applyFreeCoolRow r
			dispatchFreeCooling = True
			Exit Function
		End If
	Next
End Function
' spMaintain checkbox -> hand the blowers to the static-pressure loops for
' this pass. Returns the nosp flag for ScenarioA (False = static-pressure
' loops own the fans, so skip the scheduled blower write).
Function spMaintainNosp()
	If spMaintain.value = True Then
		Call StaticPressureEast(currentMinutes)
		Call StaticPressureWest(currentMinutes)
		spMaintainNosp = False
	Else
		spMaintainNosp = True
	End If
End Function
' Apply one heating schedule row (replaces condK-condT). rhSp is always
' HumidityAdjust(oat); heating never runs the enthalpy trim loops.
Sub applyHeatingRow(r, oat)
	Dim nosp, rh
	rh = HumidityAdjust(oat)
	If r(7) Then
		nosp = spMaintainNosp()
	Else
		nosp = True
	End If
	dataWindow.value = "Normal Heating Mode"
	skipDampers = False
	ScenarioA r(1), r(1), r(2), r(2), r(3), r(3), r(4), rh, rh, r(5), r(5), r(6), r(6), nosp
End Sub
' Heating schedule dispatch (mirrors dispatchCooling). Rows descend; the
' first row with oat > lowBound wins, so bands are (lowBound, prevBound] -
' identical to the old condK-condT ladder. The last row is an unconditional
' deep-cold catch-all (old condT: any oat =< 0; spAllowed=False preserves the
' net effect of condT's nosp=True override, which clobbered its
' static-pressure writes anyway).
' Row cols: lowBound, RAdamper, OAdamper, fan, hv4, clgSp, htgSp, spAllowed.
Function dispatchHeating(oat)
	Dim htgSched(8), i, r
	dispatchHeating = False
	If Not IsNumeric(oat) Then Exit Function
	oat = CDbl(oat)
	htgSched(0) = Array(40, 50, 100, 82, 50, 55, 55, True)
	htgSched(1) = Array(35, 50,  70, 82, 60, 55, 55, True)
	htgSched(2) = Array(30, 50,  65, 82, 60, 55, 58, True)
	htgSched(3) = Array(25, 50,  60, 88, 60, 55, 60, True)
	htgSched(4) = Array(20, 50,  55, 88, 60, 55, 60, True)
	htgSched(5) = Array(15, 50,  50, 86, 60, 55, 60, True)
	htgSched(6) = Array(10, 45,  45, 84, 60, 55, 60, True)
	htgSched(7) = Array(0,  45,  45, 84, 60, 55, 60, True)
	htgSched(8) = Array(0,  45,  45, 84, 60, 55, 60, False)
	For i = 0 To UBound(htgSched)
		r = htgSched(i)
		If oat > r(0) Or i = UBound(htgSched) Then
			applyHeatingRow r, oat
			dispatchHeating = True
			Exit Function
		End If
	Next
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
	tw = c * Atn(0.151977 * Sqr(r + 8.313659)) + Atn(c + r) - Atn(r - 1.676331)
	tw = tw + 0.00391838 * (r ^ 1.5) * Atn(0.023101 * r) - 4.686035
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
	p = (maE - oaE) / denom * 100
	If p < 0 Then p = 0
	If p > 100 Then p = 100
	FreeCoolPct = p
End Function

' --- Enthalpy-driven damper trim (mechanical cooling only) ------------------
' Feedback loop on each AHU's own mixed-air enthalpy vs the maEnthSp target.
' Enabled by the enthMaintain checkbox; applyCoolingRow then skips the
' scheduled RA/OA damper writes and these loops nudge the dampers instead.
' Steps +/-1 per pass (the script runs every minute, so max 1%/min slew),
' toward the cooler airstream when above target, warmer when below, held
' inside the pressurization clamps. Free cooling and heating are untouched.
Const ENTH_DB     = 0.5    ' deadband, Btu/lb - no move inside target +/- this
Const ENTH_RA_MIN = 40     ' RA % closed floor  (40/40 building-pressure floor)
Const ENTH_RA_MAX = 50     ' RA % closed ceiling
Const ENTH_OA_MIN = 40     ' OA % open floor
Const ENTH_OA_MAX = 55     ' OA % open ceiling (east-static envelope)

' Shared step logic: returns -1, 0, or +1 (direction to shift OA fraction).
Function enthTrimStep(hMA, tgt, oaT, raT)
	enthTrimStep = 0
	If hMA > tgt + ENTH_DB Then
		' MA too warm: shift toward the cooler airstream
		If oaT < raT Then enthTrimStep = 1 Else enthTrimStep = -1
	ElseIf hMA < tgt - ENTH_DB Then
		' MA too cold: shift toward the warmer airstream
		If oaT < raT Then enthTrimStep = -1 Else enthTrimStep = 1
	End If
End Function

Function EnthalpyTrim1( )
	Dim hMA, stp, raPos, oaPos
	If Not (IsNumeric(HtgDsch1.value) And IsNumeric(ahu2RH.value) And IsNumeric(maEnthSp.value) And IsNumeric(outsideTemp.value) And IsNumeric(ahu1RAtemp.value)) Then Exit Function
	hMA = enthalpyIP(CDbl(HtgDsch1.value), CDbl(ahu2RH.value))
	stp = enthTrimStep(hMA, CDbl(maEnthSp.value), CDbl(outsideTemp.value), CDbl(Abs(ahu1RAtemp.value)))
	If stp = 0 Then Exit Function
	raPos = CDbl(ahu1RAdamper.value) + stp
	oaPos = CDbl(ahu1OAdamper.value) + stp
	If raPos < ENTH_RA_MIN Then raPos = ENTH_RA_MIN
	If raPos > ENTH_RA_MAX Then raPos = ENTH_RA_MAX
	If oaPos < ENTH_OA_MIN Then oaPos = ENTH_OA_MIN
	If oaPos > ENTH_OA_MAX Then oaPos = ENTH_OA_MAX
	changeControl ahu1RAdamperMode, ahu1RAdamper, raPos, "manual"
	changeControl ahu1OAdamperMode, ahu1OAdamper, oaPos, "manual"
End Function

Function EnthalpyTrim2( )
	Dim hMA, stp, raPos, oaPos
	If Not (IsNumeric(HtgDsch2.value) And IsNumeric(ahu2RH.value) And IsNumeric(maEnthSp.value) And IsNumeric(outsideTemp.value) And IsNumeric(ahu2RAtemp.value)) Then Exit Function
	hMA = enthalpyIP(CDbl(HtgDsch2.value), CDbl(ahu2RH.value))
	stp = enthTrimStep(hMA, CDbl(maEnthSp.value), CDbl(outsideTemp.value), CDbl(Abs(ahu2RAtemp.value)))
	If stp = 0 Then Exit Function
	raPos = CDbl(ahu2RAdamper.value) + stp
	oaPos = CDbl(ahu2OAdamper.value) + stp
	If raPos < ENTH_RA_MIN Then raPos = ENTH_RA_MIN
	If raPos > ENTH_RA_MAX Then raPos = ENTH_RA_MAX
	If oaPos < ENTH_OA_MIN Then oaPos = ENTH_OA_MIN
	If oaPos > ENTH_OA_MAX Then oaPos = ENTH_OA_MAX
	changeControl ahu2RAdamperMode, ahu2RAdamper, raPos, "manual"
	changeControl ahu2OAdamperMode, ahu2OAdamper, oaPos, "manual"
End Function
