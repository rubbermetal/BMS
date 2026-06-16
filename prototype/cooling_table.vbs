' =====================================================================
'  PROTOTYPE - cooling schedule as data, not 15 near-identical functions
'  NOT loaded by the BMS. For review only. If approved, this folds into
'  general.vbs (table + the two subs) and the outsideTemp control field
'  (the dispatch snippet near the bottom), replacing condA..condD and
'  condF/condH/condJ plus the cooling half of the ElseIf ladder.
'
'  NOTE ON BOUNDARIES / NO FALL-THROUGH:
'  Unlike the If/ElseIf ladder, this loop has NO trailing Else, so an
'  exact boundary value can never "escape" to a default - the first row
'  the temperature reaches always wins. The fall-through bug the old
'  =< / >= operators were guarding against cannot occur here. The "incl"
'  column below reproduces the EXACT band each integer boundary lands in
'  under the original ladder (e.g. exactly 72 -> the 55-72 band, not condD).
' =====================================================================

' --- The schedule (mechanical cooling). Rows ordered HIGH band -> LOW. ---
' Columns:
'   0 lowBound   - see "incl"
'   1 RAclosed   - return-air damper % closed   (applied to AHU1 and AHU2)
'   2 OAopen     - outside-air damper % open     (applied to AHU1 and AHU2)
'   3 fan        - blower %                       (applied to AHU1 and AHU2)
'   4 hv4        - HV4 % (-1 = old condD rule: 50 if OAT>alphanum3 else 100)
'   5 rhSp       - RH setpoint
'   6 dischHi    - discharge setpoint when return RH >= 55%
'   7 dischLo    - discharge setpoint when return RH < 55%  (dischargeSp reset)
'   8 chillerOnly- True = only runs when chiller ON; if OFF, caller free-cools
'   9 incl       - True: row claims temp >= lowBound (inclusive lower edge)
'                  False: row claims temp >  lowBound (exclusive; matches the
'                  old > 72 / > 55 / > 50 edges so 72,55,50 land as they do today)
Dim clgSched
clgSched = Array( _
    Array(92, 40, 45, 98,  50, 55, 55, 55, False, True), _
    Array(91, 41, 46, 98,  50, 55, 55, 55, False, True), _
    Array(90, 42, 47, 98,  50, 55, 55, 55, False, True), _
    Array(89, 43, 48, 98,  50, 55, 55, 55, False, True), _
    Array(88, 44, 49, 98,  50, 55, 55, 55, False, True), _
    Array(87, 45, 50, 98,  50, 55, 55, 55, False, True), _
    Array(86, 44, 51, 98,  50, 55, 55, 55, False, True), _
    Array(85, 43, 52, 98,  50, 55, 55, 55, False, True), _
    Array(84, 42, 53, 98,  50, 55, 55, 55, False, True), _
    Array(83, 41, 54, 98,  50, 55, 55, 55, False, True), _
    Array(80, 40, 55, 98,  50, 55, 55, 55, False, True), _
    Array(72, 45, 55, 98,  -1, 55, 55, 55, False, False), _
    Array(62, 50, 50, 98, 100, 55, 55, 53, True,  False), _
    Array(55, 45, 45, 98, 100, 55, 53, 50, True,  False), _
    Array(50, 40, 40, 98,  50, 50, 55, 50, True,  False), _
    Array(45, 40, 40, 95,  50, 50, 55, 50, True,  True) )

' --- Apply one schedule row (replaces the body of each condX function) ---
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
	sp = dischargeSp(r(6), r(7))   ' equals dischHi when hi=lo (no reset)
	ScenarioA r(1), r(1), r(2), r(2), r(3), r(3), hv4, _
	          r(5), r(5), sp, sp, sp, sp, True
End Sub

' --- Dispatch: find the band and apply it. ---
' Returns True if handled. Returns False ONLY for a chiller-OFF sub-72 band,
' which means the caller should run the matching free-cooling function.
Function dispatchCooling(oat)
	Dim i, r, reached
	dispatchCooling = False
	For i = 0 To UBound(clgSched)
		r = clgSched(i)
		If r(9) Then
			reached = (oat >= r(0))     ' inclusive lower edge
		Else
			reached = (oat > r(0))      ' exclusive (old > 72 / > 55 / > 50)
		End If
		If reached Then
			If r(8) And chillerStatus.value <> "ON" Then
				Exit Function           ' chiller off -> free cooling (caller)
			End If
			applyCoolingRow r
			dispatchCooling = True
			Exit Function
		End If
	Next
End Function

' =====================================================================
'  How the outsideTemp control field's cooling section would look.
'  Replaces every condA..condD call AND the condE/F/G/H/I/J chiller
'  branches. Heating (< 45) and everything else stays exactly as-is.
'  No trailing Else can be reached for any temp >= 45.
' =====================================================================
'
'	If outsideTemp.value >= 45 Then
'		If Not dispatchCooling(outsideTemp.value) Then
'			' chiller OFF in a 45-72 band -> free cooling (unchanged funcs)
'			If outsideTemp.value > 55 Then
'				Call condE()
'			ElseIf outsideTemp.value > 50 Then
'				Call condG()
'			Else
'				Call condI()
'			End If
'		End If
'	ElseIf outsideTemp.value > 0 Then
'		' ... existing heating bands condK..condS, unchanged ...
'	Else
'		Call condT()
'	End If
