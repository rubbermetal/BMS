Class ControlField
	Public value
End Class

Dim sleepX, control, status
 ' Time to sleep between setting changes
sleepX = 2000


Function CustomLog(x)
	Dim result, y, power, term, n, maxIterations

	maxIterations = 100
	y = (x - 1) / (x + 1)
	result = 0
	power = y
	term = 2 * power / 1
	For n = 1 To maxIterations
		result = result + term
		power = power * y * y
		term = 2 * power / (2 * n + 1)
		Next
	CustomLog = result
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
				sleep(sleepX)
			End If
		Else
			If control.value <> Setpoint Then
				control.value = Setpoint 
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
Sub ScenarioA(raDamper1,raDamper2,oaDamper1,oaDamper2,blowerControl1,blowerControl2,hv4Damper,rhSp1,rhSp2,clgSp1,clgSp2,htgSp1,htgSp2)
	changeControl ahu1RAdamperMode, ahu1RAdamper, raDamper1, "manual"
	changeControl ahu2RAdamperMode, ahu2RAdamper, raDamper2, "manual"
	changeControl ahu1OAdamperMode, ahu1OAdamper, oaDamper1, "manual"
	changeControl ahu2OAdamperMode, ahu2OAdamper, oaDamper2, "manual"
	changeControl ahu1BlowerMode, ahu1BlowerControl, blowerControl1, "manual"
	changeControl ahu2BlowerMode, ahu2BlowerControl, blowerControl2, "manual"
	changeControl hv4DamperMode, hv4MAdamper, hv4Damper, "manual"
	changeControl ahu1RhSpMode, ahu1RhSp, rhSp1, "manual"
	changeControl ahu2RhSpMode, ahu2RhSp, rhSp2, "manual"
	changeControl ahu1ClgSpMode, ahu1ClgSp, clgSp1, "manual"
	changeControl ahu2ClgSpMode, ahu2ClgSp, clgSp2, "manual"
	changeControl ahu1HtgSpMode, ahu1HtgSp, htgSp1, "manual"
	changeControl ahu2HtgSpMode, ahu2HtgSp, htgSp2, "manual"

End Sub

