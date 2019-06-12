'=== Calc Group Types ===
CAW PT
CAW 4 & 4
CAW 5 & 2
CAW CASUAL
CAW FT OVER 80
CAW FT OVER 40
CAW 6 WK ROTATION
CAW FT TEMP


'----------------------------------------------- OverTime -----------------------------------
'Biweekly type
If strCalcGrp = "CAW FT OVER 80" Then
	If (there is overtime for an employee) then
		'Check Bi-Weekly work hours first
		If (Bi-Weekly reg hours < 80) ' Do we need to check it has to be 80, no more and no less?
			Error this employee (fill color)
		Else
			Nothing To do. No Per Day Check for 'OVER 80'
		End If
	End If
End If


'weekly type
If strCalcGrp = "CAW 5 & 2" Or strCalcGrp = "PT" Or strCalcGrp = "FT Temp" Or strCalcGrp = "Casual" Then
    If (there is overtime for an employee)  then
		If (daily reg >= 8) or (this pay week reg >= 40) or (this whole pay period reg >= 80) or (that work date timeCode is "STA/STA_BANK/SDW")
			ok
		Else
			Error the overtime row (Fill color)
			Error this employee (Font color)
		End If
	End If
End If

If strCalcGrp = "FT Over 40" Then
    'equal to "CAW 5 & 2", but different criteria
	daily reg criteria is 10, not 8
	week criteria is 40
	this whole pay is 80
End If

If strCalcGrp = "4&4" Then
	If (there is overtime for an employee) then
		If (daily reg does not exsit) Or (daily reg >= 11.5) Then
			ok
		Else
		    Error the overtime row (Fill Color)
			Error this employee (Font color)
		End If
	End If
End If

If strCalcGrp = "6WK" Then
	Do Nothing
End If

'Per Day Type



'----------------------------------------------- PaidLeave -----------------------------------

If PaidLeave(TBD/FAM/SCK/VAC/CXL) in someday
	Total hours(all hour type except PREM and STA/STA - BANK) for that day cannot exceed the criteria in the column on the paper
End If

'----------------------------------------------- TotalOT -----------------------------------
Follow Total OT column on the paper
   errMsg = "TotalOT = x, Total REG=y, Threshold=20/40/10"

'----------------------------------------------- Rounding -----------------------------------
The last 2 digits should be 0.25/0.5/0.75/0.0
        If (calcGrp = 4&4) or (calcGrp = 6WK) Then
			If (hourType is OT)
				ignore rounding check for this row
			End If
			
		End If

 
If (calcGrp = PT Or calcGrp = CASUAL) And (timeCode = "STA" or "STA_BANK") Then    'SDW needs check
	Do not check rounding (because it's inherited from previous pay period)
	staDayIncludedForPtCasual = True
End If

If (calcGrp = PT Or calcGrp = CASUAL)
	If (checkNeeded)
	{
	    If ((weekly reg >= 40) Or (Biweekly >= 80))
			errMsg = "STA And weekly >=40 or BiWeekly >= 80"
	}
	If (rounding error found)
		color this line with errMsg
End If

@todo: 
'----------------------------------------------- DailyHours -----------------------------------
check daily total hours (except PREM and STA) < column value in the table



???'@todo: How to handle EXCESS_HRS_OT?
