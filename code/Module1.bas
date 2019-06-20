Attribute VB_Name = "Module1"
' this is the VB6/VBA equivalent of a struct data, no methods.
' column No. of each field in over time table
Private Type OverTimeTableCfgType
    titleRowNo As Integer
    firstDataRowNo As Integer
    
    calcGrpColNo As Integer
    payGrpDesColNo As Integer
    employeeIdColNo As Integer
    fullNameColNo  As Integer
    hourTypeColNo As Integer
    workDateColNo As Integer
    timeCodeColNo As Integer
    hoursColNo As Integer
    teamNameColNo As Integer
    depNameColNo As Integer
    roundingColNo As Integer
    errMsgColNo As Integer
End Type

Public Enum OtByDayChkType
    OT_BY_DAY_CHK_PAID_LEAVE = 1
    OT_BY_DAY_CHK_REG_HOURS_OT = 2
End Enum

'We assume there are no more rows than MAX_ROWS_FOR_ONE_EMPLOYEE for each employee to avoid Transpose and ReDim Preserve the array Size
'@TODO: If some table exceeds this limit, we have to take a look and probably change this value accordingly.
Private Const DEFAULT_ERR_MSG_COL_NAME = "ERROR MESSAGE"
Public Const MAX_ROWS_FOR_ONE_EMPLOYEE = 50
Public Const REG_WORK_HOURS_LOW_LIMIT_PER_DAY = 8
Public Const REG_WORK_HOURS_LOW_LIMIT_WHOLE_PAY = 80
Public Const REG_HOURS_CRITERIA_WEEKLY = 40
Public Const REG_HOURS_CRITERIA_BIWEEKLY = 80
Public Const ABM_OT_HOURS_FOR_6WK = 40
Public Const ABM_OT_HOURS_FOR_CASUAL = 10
Public Const ABM_OT_HOURS_FOR_OTHERS = 20
Public Const DEFAULT_FLOAT_PRECISION = 0.001
Public Const ROUNDING_PRECISION = 0.00001
Public Const DEFAULT_FONT_COLOR_FOR_ABNORMAL_ROW = vbRed
Public Const DEFAULT_BG_COLOR_FOR_ABNORMAL_ROW = vbYellow

'Helper function: clear Immediate window
Function cw()
    Application.VBE.Windows("Immediate").SetFocus
    If Application.VBE.ActiveWindow.Caption = "Immediate" And Application.VBE.ActiveWindow.Visible Then
        Application.SendKeys "^g ^a {DEL}"
    End If
End Function

'Find the coloumn No. of each field in over time table.
'Only search the first maxRowNum of Rows because the table title
'should be in the first few rows.
Private Function getTableCfg(lastRowNo As Long, lastColNo As Long, ByRef tblCfg As OverTimeTableCfgType)
    Dim found As Boolean: found = False
    Dim i As Long, j As Long
    'Find the correct column we care
    For i = 1 To lastRowNo
        For j = 1 To lastColNo
            With tblCfg
                If ActiveSheet.Cells(i, j).Value = "Calc Group Name" Then
                    .firstDataRowNo = i + 1
                    .titleRowNo = i
                    .calcGrpColNo = j
                    found = True
                ElseIf ActiveSheet.Cells(i, j).Value = "Pay Group Description" Then
                    .payGrpDesColNo = j
                ElseIf ActiveSheet.Cells(i, j).Value = "Employee ID" Then
                    .employeeIdColNo = j
                ElseIf ActiveSheet.Cells(i, j).Value = "Full Name" Then
                    .fullNameColNo = j
                ElseIf ActiveSheet.Cells(i, j).Value = "Hour Type" Then
                    .hourTypeColNo = j
                ElseIf ActiveSheet.Cells(i, j).Value = "Work Date" Then
                    .workDateColNo = j
                ElseIf ActiveSheet.Cells(i, j).Value = "Time Code" Then
                    .timeCodeColNo = j
                ElseIf ActiveSheet.Cells(i, j).Value = "Hours" Then
                    .hoursColNo = j
                ElseIf ActiveSheet.Cells(i, j).Value = "Team Name" Then
                    .teamNameColNo = j
                ElseIf ActiveSheet.Cells(i, j).Value = "Department Name" Then
                    .depNameColNo = j
                ElseIf ActiveSheet.Cells(i, j).Value = DEFAULT_ERR_MSG_COL_NAME Then
                    .errMsgColNo = j
                End If
            End With
        Next j
        If found = True Then Exit For
    Next i
End Function

'Helper function: check table Colums has the fields we need
Private Function chkTbleCfg(ByRef tblCfg As OverTimeTableCfgType) As Boolean
    Dim result As Boolean: result = True
    
    With tblCfg
        If ((.firstDataRowNo <= 0) Or (.titleRowNo <= 0) Or _
            (.calcGrpColNo <= 0) Or _
            (.employeeIdColNo <= 0) Or _
            (.hourTypeColNo <= 0) Or (.workDateColNo <= 0) Or _
            (.timeCodeColNo <= 0) Or _
            (.hoursColNo <= 0)) Then
            result = False
        End If
    End With
    
    chkTbleCfg = result
End Function

'Helper function: Debug Print one row
Function dbgPrintRow(rowNo As Integer, ByRef rngToPrint As Range, ByRef prefix As String)
    Dim r As Range
    Dim c As Range
    Dim outStr As String: outStr = vbNullString
    Debug.Print "dbgPrintRow: rowNo=" & rowNo & ". prefix= " & prefix

    Set r = rngToPrint.Rows(rowNo)
    For Each c In r.Cells
        If (IsEmpty(c)) Then Exit For
        outStr = outStr & c.Value & " | "
    Next c
    
    Debug.Print outStr
End Function

'Helper function: Print an String Array to a File
Function printArray2File(ByRef arrIn() As String, ByRef arrName As String, arrLen As Integer)
    Dim elem As String
    Dim i As Integer
    Dim saveFile As String
    saveFile = Application.ActiveWorkbook.path & "\" & arrName & ".txt"
    
    If Dir(saveFile) = "" Then
        Open saveFile For Output As #1
    Else
        Open saveFile For Append As #1
    End If
    
    Print #1, "==Start printArray: arrName= " & arrName & ", arrLen= " & arrLen

    For i = 1 To arrLen
        Print #1, arrIn(i)
    Next i
    
    Print #1, "==End printArray: arrName= " & arrName & " arrLen=" & arrLen
    Close #1
End Function

'Get all unique Strings in an array by target column
'arrDataOut is a 1-D array
Function getAllUniqueStrsByColInArray(titleRowNo As Integer, targetCol As Integer, _
                                      ByRef arrDataIn As Variant, nrRowsArrIn As Integer, _
                                      ByRef arrDataOut() As String)
    Dim elem As String
    Dim foundMatch As Boolean
    Dim nrUnique As Integer: nrUnique = 0
    
    Dim i As Integer
    For i = 1 To nrRowsArrIn
        foundMatch = False
        elem = arrDataIn(i, targetCol)
        For j = 1 To nrUnique
            If elem = arrDataOut(j) Then
                foundMatch = True
            End If
        Next j

        'new element And Not Null String
        If Not foundMatch Then
            If Not Trim(elem & vbNullString) = vbNullString Then
                nrUnique = nrUnique + 1
                ReDim Preserve arrDataOut(nrUnique)
                arrDataOut(nrUnique) = elem
            End If
        End If
    Next i

'    Dim titleStr As String
'    titleStr = ActiveSheet.Cells(titleRowNo, targetCol).Value
'    Call printArray2File(unique, titleStr, nrUnique)
End Function

'Get all rows that have same value in a certain column of a Array
'e.g., Get all rows in an employee range with same Work Date Field
'arrDataOut is a 2-D array
Function getAllRowsHasSameOneColValInArray(titleRowNo As Integer, targetCol As Integer, ByRef targetStr As String, _
                                        ByRef arrDataIn As Variant, nrRowsArrIn As Integer, _
                                        ByRef arrDataOut As Variant, ByRef nrOfRealRowsOut As Integer)
    Dim i As Integer
    Dim nrColsArrIn As Integer: nrColsArrIn = UBound(arrDataIn, 2) 'ActiveSheet.UsedRange.Columns.Count
    
    nrOfRealRowsOut = 0
    For i = 1 To nrRowsArrIn
        If arrDataIn(i, targetCol) = targetStr Then
            nrOfRealRowsOut = nrOfRealRowsOut + 1
            If nrOfRealRowsOut > MAX_ROWS_FOR_ONE_EMPLOYEE Then
                MsgBox "Cannot Handle more than " & MAX_ROWS_FOR_ONE_EMPLOYEE & "records per employee Now"
                Exit Function
            End If
            'Copy 1 row to arrDataOut as output
            For j = 1 To nrColsArrIn
                arrDataOut(nrOfRealRowsOut, j) = arrDataIn(i, j)
            Next j
        End If
    Next i

    'Get the title for targetCol and print debug info
'    Dim titleStr As String
'    titleStr = ActiveSheet.Cells(titleRowNo, targetCol).Value
'    Debug.Print "titleStr: " & titleStr & " targetStr: " & targetStr & ". " & cnt & " Records that has same value in Column " & targetCol
End Function

'Check if hour type is over time: (SDW) Stat Holody Worked and any OT in STA day are not considered as normal overtime
Private Function isNormalOverTimeHour(ByRef tblCfg As OverTimeTableCfgType, ByRef oneEmpRecords As Variant, _
                                        nrOfRealRowsForEmp As Integer, ByRef tgtWkDate As Date, _
                                        ByRef hourType As String, ByRef timeCode As String) As Boolean
    isNormalOverTimeHour = False
    
    If isStaWorkDate(tblCfg, oneEmpRecords, nrOfRealRowsForEmp, tgtWkDate) Then
        'Stat holiday is normal overtime hour, return False in this case,
        'so do nothing and use default return value
    ElseIf hourType = "OT1.5" Or hourType = "BANK_OT" Or hourType = "OT1.5_AVG" Then
        If timeCode <> "SDW" Then
            isNormalOverTimeHour = True
        End If
    End If
End Function

'Check if current work date is a STA (stat) date
Private Function isStaWorkDate(ByRef tblCfg As OverTimeTableCfgType, ByRef oneEmpRecords As Variant, _
                                nrOfRealRowsForEmp As Integer, ByRef tgtWkDate As Date) As Boolean
    isStaWorkDate = False
    
    Dim j As Integer
    Dim strTimeCode As String
    Dim currDate As Date
    
    For j = 1 To nrOfRealRowsForEmp
        strTimeCode = oneEmpRecords(j, tblCfg.timeCodeColNo)
        currDate = oneEmpRecords(j, tblCfg.workDateColNo)
        
        If (currDate = tgtWkDate) Then
            If (strTimeCode = "STA") Or (strTimeCode = "STA_BANK") Then
                isStaWorkDate = True
                Exit Function
            End If
        End If
    Next j
End Function

'Check if hour type is regular work hour (not over time)
Private Function isRegularHour(ByRef hourType As String, ByRef timeCode As String) As Boolean
    isRegularHour = False
    
    If hourType = "REG" Or hourType = "BANK_REG" Then
        If timeCode <> "TRV" Then
            isRegularHour = True
        End If
    End If
End Function

Private Function isWorkDateFirstWeek(ByRef currWorkDate As Date, ByRef payPeriodStartDate As Date) As Boolean
    Dim payPeriodFirstWeekEndDate As Date
    payPeriodFirstWeekEndDate = DateAdd("d", 7 - 1, payPeriodStartDate)
    
    If currWorkDate <= payPeriodFirstWeekEndDate Then
        isWorkDateFirstWeek = True
    Else
        isWorkDateFirstWeek = False
    End If
End Function

Private Function isTimeCodePaidLeave(ByRef timeCode As String) As Boolean
    'For now, we only think these 5 'Time Code' are Paid Leave
    If (timeCode = "TBD") Or (timeCode = "FAM") Or (timeCode = "SCK") Or _
       (timeCode = "VAC") Or (timeCode = "VAC - PAY PERIOD") Or _
       (timeCode = "CXL") Then
        isTimeCodePaidLeave = True
    Else
        isTimeCodePaidLeave = False
    End If
End Function

Private Function colorEmpOTOneDay(ByRef tblCfg As OverTimeTableCfgType, ByRef arrAllValidData As Variant, _
                                ByRef oneEmpRecords As Variant, nrOfRealRowsForEmp As Integer, _
                                ByRef tgtStrEmpId As String, _
                                ByRef tgtWorkDate As Date, ByRef errMsg As String)
    Dim i As Integer
    Dim strHourType As String
    Dim strTimeCode As String
    Dim currWkDate As Date
    
    For i = 1 To UBound(arrAllValidData)
        strHourType = arrAllValidData(i, tblCfg.hourTypeColNo)
        strTimeCode = arrAllValidData(i, tblCfg.timeCodeColNo)
        currWkDate = arrAllValidData(i, tblCfg.workDateColNo)
        
        If arrAllValidData(i, tblCfg.employeeIdColNo) = tgtStrEmpId And strHourType <> "REG_PREM" Then
            'Change Font Color for this employee
            ActiveSheet.UsedRange.Rows(i + tblCfg.titleRowNo).Font.Color = DEFAULT_FONT_COLOR_FOR_ABNORMAL_ROW
            If tgtWorkDate = currWkDate Then
                'Change Background color for overtime hours in this day
                If isNormalOverTimeHour(tblCfg, oneEmpRecords, nrOfRealRowsForEmp, currWkDate, strHourType, strTimeCode) Then
                    ActiveSheet.UsedRange.Rows(i + tblCfg.titleRowNo).Interior.Color = DEFAULT_BG_COLOR_FOR_ABNORMAL_ROW
                    ActiveSheet.Cells(i + tblCfg.titleRowNo, tblCfg.errMsgColNo) = errMsg
                End If
            End If
        End If
    Next i
End Function

Private Function colorEmpOTWholePay(ByRef tblCfg As OverTimeTableCfgType, ByRef arrAllValidData As Variant, _
                                    ByRef oneEmpRecords As Variant, nrOfRealRowsForEmp As Integer, _
                                    ByRef tgtStrEmpId As String, ByRef errMsg As String)
    Dim i As Integer
    Dim strHourType As String
    Dim strTimeCode As String
    Dim currWkDate As Date
    
    For i = 1 To UBound(arrAllValidData)
        strHourType = arrAllValidData(i, tblCfg.hourTypeColNo)
        strTimeCode = arrAllValidData(i, tblCfg.timeCodeColNo)
        currWkDate = arrAllValidData(i, tblCfg.workDateColNo)
        
        If arrAllValidData(i, tblCfg.employeeIdColNo) = tgtStrEmpId And strHourType <> "REG_PREM" Then
            'Change Font Color for this day
            ActiveSheet.UsedRange.Rows(i + tblCfg.titleRowNo).Font.Color = DEFAULT_FONT_COLOR_FOR_ABNORMAL_ROW
            'Change Background color for overtime hours in this day
            If isNormalOverTimeHour(tblCfg, oneEmpRecords, nrOfRealRowsForEmp, currWkDate, strHourType, strTimeCode) Then
                ActiveSheet.UsedRange.Rows(i + tblCfg.titleRowNo).Interior.Color = DEFAULT_BG_COLOR_FOR_ABNORMAL_ROW
                ActiveSheet.Cells(i + tblCfg.titleRowNo, tblCfg.errMsgColNo) = errMsg
            End If
        End If
    Next i
End Function

'Color employee paid leave for one day
Function colorEmpPLOneDay(ByRef tblCfg As OverTimeTableCfgType, ByRef arrAllValidData As Variant, _
                            ByRef oneEmpRecords As Variant, _
                            ByRef tgtStrEmpId As String, _
                            ByRef tgtWorkDate As Date, ByRef errMsg As String)
    Dim i As Integer
    Dim strHourType As String
    Dim strTimeCode As String
    Dim currWkDate As Date
    
    For i = 1 To UBound(arrAllValidData)
        strHourType = arrAllValidData(i, tblCfg.hourTypeColNo)
        strTimeCode = arrAllValidData(i, tblCfg.timeCodeColNo)
        currWkDate = arrAllValidData(i, tblCfg.workDateColNo)
        
        If arrAllValidData(i, tblCfg.employeeIdColNo) = tgtStrEmpId And strHourType <> "REG_PREM" Then
            'Change Font Color for this employee
            ActiveSheet.UsedRange.Rows(i + tblCfg.titleRowNo).Font.Color = DEFAULT_FONT_COLOR_FOR_ABNORMAL_ROW
            If (tgtWorkDate = currWkDate) Then
                'Change Background color for Paid Leave Hours hours in this day
                If isTimeCodePaidLeave(strTimeCode) Then
                    ActiveSheet.UsedRange.Rows(i + tblCfg.titleRowNo).Interior.Color = DEFAULT_BG_COLOR_FOR_ABNORMAL_ROW
                    ActiveSheet.Cells(i + tblCfg.titleRowNo, tblCfg.errMsgColNo) = errMsg
                End If
            End If
        End If
    Next i

End Function

Private Function colorEmpOneRow(ByRef tblCfg As OverTimeTableCfgType, ByRef strEmpId As String, _
                              ByRef oneEmpRecords As Variant, nrOfRealRowsForEmp As Integer, _
                              rowNoInOneEmpRec As Integer, ByRef arrAllValidData As Variant, _
                              ByRef errMsg As String)
    Dim strHourType As String
    Dim strTimeCode As String
    Dim currWkDate As Date
    Dim hours As Single
    
    Dim tmpEmpId As String
    Dim tmpHourType As String
    Dim tmpTimeCode As String
    Dim tmpWkDate As Date
    Dim tmpHours As Single
    
    strHourType = oneEmpRecords(rowNoInOneEmpRec, tblCfg.hourTypeColNo)
    strTimeCode = oneEmpRecords(rowNoInOneEmpRec, tblCfg.timeCodeColNo)
    currWkDate = oneEmpRecords(rowNoInOneEmpRec, tblCfg.workDateColNo)
    hours = oneEmpRecords(rowNoInOneEmpRec, tblCfg.hoursColNo)
    
    'loop all the data to color all records of the same employee
    For j = 1 To UBound(arrAllValidData)
        tmpEmpId = arrAllValidData(j, tblCfg.employeeIdColNo)
        tmpHourType = arrAllValidData(j, tblCfg.hourTypeColNo)
        tmpTimeCode = arrAllValidData(j, tblCfg.timeCodeColNo)
        tmpWkDate = arrAllValidData(j, tblCfg.workDateColNo)
        tmpHours = arrAllValidData(j, tblCfg.hoursColNo)
        
        If (tmpEmpId = strEmpId) And (tmpHourType <> "REG_PREM") Then
            'color all font for this employee
            ActiveSheet.UsedRange.Rows(j + tblCfg.titleRowNo).Font.Color = DEFAULT_FONT_COLOR_FOR_ABNORMAL_ROW
            If (tmpHourType = strHourType) And (tmpTimeCode = strTimeCode) And _
               (tmpWkDate = currWkDate) And (tmpHours = hours) Then
               'color error row's background
                ActiveSheet.UsedRange.Rows(j + tblCfg.titleRowNo).Interior.Color = DEFAULT_BG_COLOR_FOR_ABNORMAL_ROW
                If Not Trim(errMsg & vbNullString) = vbNullString Then
                    ActiveSheet.Cells(j + tblCfg.titleRowNo, tblCfg.errMsgColNo) = errMsg
                End If
            End If
        End If
    Next j
End Function

Private Function calcEmpHoursInOneDay(ByRef tblCfg As OverTimeTableCfgType, _
                                        ByRef oneEmpRecords As Variant, nrOfRealRowsForEmp As Integer, _
                                        ByRef tgtDate As Date, _
                                        ByRef regHourInOneDay As Single, ByRef otHourInOneDay As Single, _
                                        ByRef wkHourInOneDay As Single)
    regHourInOneDay = 0
    otHourInOneDay = 0
    wkHourInOneDay = 0

    Dim strHourType As String
    Dim strTimeCode As String
    Dim currDate As Date
    Dim j As Integer
    
    For j = 1 To nrOfRealRowsForEmp
        strHourType = oneEmpRecords(j, tblCfg.hourTypeColNo)
        strTimeCode = oneEmpRecords(j, tblCfg.timeCodeColNo)
        currDate = oneEmpRecords(j, tblCfg.workDateColNo)
        
        If currDate = tgtDate Then
            If isNormalOverTimeHour(tblCfg, oneEmpRecords, nrOfRealRowsForEmp, currDate, strHourType, strTimeCode) Then
                otHourInOneDay = otHourInOneDay + oneEmpRecords(j, tblCfg.hoursColNo)
            ElseIf isRegularHour(strHourType, strTimeCode) Then
                regHourInOneDay = regHourInOneDay + oneEmpRecords(j, tblCfg.hoursColNo)
                'Define Work hours: Subset of Regular hours but it's real work hours, Not STA, Not Paid Leave.
                If Not isTimeCodePaidLeave(strTimeCode) And strTimeCode <> "STA" And (strTimeCode <> "STA_BANK") Then
                    wkHourInOneDay = wkHourInOneDay + oneEmpRecords(j, tblCfg.hoursColNo)
                End If
            End If
        End If
    Next j
End Function

Private Function calcEmpHoursInWholePay(ByRef tblCfg As OverTimeTableCfgType, _
                                        ByRef oneEmpRecords As Variant, nrOfRealRowsForEmp As Integer, _
                                        ByRef regHourInWholePay As Single, ByRef otHourInWholePay As Single)
    otHourInWholePay = 0
    regHourInWholePay = 0
    Dim strHourType As String
    Dim strTimeCode As String
    Dim currWkDate As Date
    Dim j As Integer
    
    For j = 1 To nrOfRealRowsForEmp
        strHourType = oneEmpRecords(j, tblCfg.hourTypeColNo)
        strTimeCode = oneEmpRecords(j, tblCfg.timeCodeColNo)
        currWkDate = oneEmpRecords(j, tblCfg.workDateColNo)
        
        If isNormalOverTimeHour(tblCfg, oneEmpRecords, nrOfRealRowsForEmp, currWkDate, strHourType, strTimeCode) Then
            otHourInWholePay = otHourInWholePay + oneEmpRecords(j, tblCfg.hoursColNo)
        ElseIf isRegularHour(strHourType, strTimeCode) Then
            regHourInWholePay = regHourInWholePay + oneEmpRecords(j, tblCfg.hoursColNo)
        End If
    Next j
End Function

Private Function calcEmpHoursIn2Weeks(ByRef tblCfg As OverTimeTableCfgType, _
                                        ByRef oneEmpRecords As Variant, nrOfRealRowsForEmp As Integer, _
                                        ByRef payPeriodStartDate As Date, _
                                        ByRef regHourInWk1 As Single, ByRef otHourInWk1 As Single, _
                                        ByRef regHourInWk2 As Single, ByRef otHourInWk2 As Single)
    regHourInWk1 = 0
    otHourInWk1 = 0
    regHourInWk2 = 0
    otHourInWk2 = 0
    
    Dim strHourType As String
    Dim strTimeCode As String
    Dim currDate As Date
    Dim j As Integer
    Dim payPeriodFirstWeekEndDate As Date
    
    payPeriodFirstWeekEndDate = DateAdd("d", 7 - 1, payPeriodStartDate)

    For j = 1 To nrOfRealRowsForEmp
        strHourType = oneEmpRecords(j, tblCfg.hourTypeColNo)
        strTimeCode = oneEmpRecords(j, tblCfg.timeCodeColNo)
        currDate = oneEmpRecords(j, tblCfg.workDateColNo)
        
        If isWorkDateFirstWeek(currDate, payPeriodStartDate) Then
            If isNormalOverTimeHour(tblCfg, oneEmpRecords, nrOfRealRowsForEmp, currDate, strHourType, strTimeCode) Then
                otHourInWk1 = otHourInWk1 + oneEmpRecords(j, tblCfg.hoursColNo)
            ElseIf isRegularHour(strHourType, strTimeCode) Then
                regHourInWk1 = regHourInWk1 + oneEmpRecords(j, tblCfg.hoursColNo)
            End If
        Else
            If isNormalOverTimeHour(tblCfg, oneEmpRecords, nrOfRealRowsForEmp, currDate, strHourType, strTimeCode) Then
                otHourInWk2 = otHourInWk2 + oneEmpRecords(j, tblCfg.hoursColNo)
            ElseIf isRegularHour(strHourType, strTimeCode) Then
                regHourInWk2 = regHourInWk2 + oneEmpRecords(j, tblCfg.hoursColNo)
            End If
        End If
    Next j
End Function

Private Function handleEmpHoursRounding(ByRef tblCfg As OverTimeTableCfgType, ByRef strEmpId As String, _
                              ByRef oneEmpRecords As Variant, nrOfRealRowsForEmp As Integer, _
                              ByRef payPeriodStartDate As Date, _
                              ByRef arrAllValidData As Variant)
    Dim hours As Single
    Dim errHour As Single
    Dim strCalcGrp As String
    Dim strHourType As String
    Dim strTimeCode As String
    Dim currWkDate As Date
    Dim i As Integer
    Dim chkNeeded As Boolean
    Dim staDayIncludedForPtCasual As Boolean: staDayIncludedForPtCasual = False
    Dim errMsg As String: errMsg = vbNullString

    Dim regHourInWk1 As Single
    Dim otHourInWk1 As Single
    Dim regHourInWk2 As Single
    Dim otHourInWk2 As Single

    strCalcGrp = oneEmpRecords(1, tblCfg.calcGrpColNo)
    
    Call calcEmpHoursIn2Weeks(tblCfg, oneEmpRecords, nrOfRealRowsForEmp, payPeriodStartDate, _
                            regHourInWk1, otHourInWk1, regHourInWk2, otHourInWk2)
    
    For i = 1 To nrOfRealRowsForEmp
        strHourType = oneEmpRecords(i, tblCfg.hourTypeColNo)
        strTimeCode = oneEmpRecords(i, tblCfg.timeCodeColNo)
        currWkDate = oneEmpRecords(i, tblCfg.workDateColNo)
        hours = oneEmpRecords(i, tblCfg.hoursColNo)
        
        '@todo: How to handle EXCESS_HRS_OT?
        'Since the decimal part has to be 0.25/0.5/0.75/0.0, we can multiply 4 to make it always 0.0
        chkNeeded = True
        
        If (strCalcGrp = "CAW 4 & 4") Or (strCalcGrp = "CAW 6 WK ROTATION") Then
            'Overtime is carried from last pay period or reconciled automatically, no rounding check
            If isNormalOverTimeHour(tblCfg, oneEmpRecords, nrOfRealRowsForEmp, currWkDate, strHourType, strTimeCode) Then
                chkNeeded = False
            End If
        ElseIf (strCalcGrp = "CAW PT") Or (strCalcGrp = "CAW CASUAL") Then
            'Overtime is carried from last 4 weeks automatically, no rounding check
            If (strTimeCode = "STA") Or (strTimeCode = "STA_BANK") Then
                chkNeeded = False
                staDayIncludedForPtCasual = True
            End If
        End If

        If (strHourType = "REG_PREM") Then  'No check for premium(duplicate) rows
            chkNeeded = False
        End If

        If chkNeeded Then
            errHour = hours * 4 - Int(hours * 4)
            If (errHour > ROUNDING_PRECISION) Then
                If staDayIncludedForPtCasual Then
                    If isWorkDateFirstWeek(currWkDate, payPeriodStartDate) Then 'this day is in first week
                        If (regHourInWk1 >= REG_HOURS_CRITERIA_WEEKLY) Then
                            errMsg = "STA included, RegHourWeekly >= " & REG_HOURS_CRITERIA_WEEKLY & ";"
                        End If
                    Else 'this day is in second week
                        If (regHourInWk2 >= REG_HOURS_CRITERIA_WEEKLY) Then
                            errMsg = "STA included, RegHourWeekly >= " & REG_HOURS_CRITERIA_WEEKLY & ";"
                        End If
                    End If
                    
                    If (regHourInWk1 + regHourInWk2 >= REG_HOURS_CRITERIA_BIWEEKLY) Then
                        errMsg = errMsg & " RegHourBiWeekly >= " & REG_HOURS_CRITERIA_BIWEEKLY
                    End If
                End If

                Call colorEmpOneRow(tblCfg, strEmpId, oneEmpRecords, nrOfRealRowsForEmp, i, arrAllValidData, errMsg)
            End If
        End If
    Next i
End Function

Function hanldeEmpHoursByDay(ByRef tblCfg As OverTimeTableCfgType, ByRef strEmpId As String, _
                             ByRef oneEmpRecords As Variant, nrOfRealRowsForEmp As Integer, _
                             ByRef arrAllValidData As Variant, ByRef chkType As OtByDayChkType, _
                             ByRef dictDailyHoursForCalcGrp As Object)
    Dim errMsg As String
    
    'Get Work Date For this Employee
    Dim strUniqueWorkDateForOneEmp() As String

    ' - Get all unique work date strings for this employee
    Call getAllUniqueStrsByColInArray(tblCfg.titleRowNo, tblCfg.workDateColNo, oneEmpRecords, _
                                      nrOfRealRowsForEmp, strUniqueWorkDateForOneEmp)
    If IsEmpty(strUniqueWorkDateForOneEmp) Then
        MsgBox "Work Date for Employee ID: " & strEmpId & "Is Empty"
        Exit Function
    End If
    
    Dim arrSameWorkDate() As Variant
    Dim nrOfRealRowsSameWorkDate As Integer
    ReDim arrSameWorkDate(MAX_ROWS_FOR_ONE_EMPLOYEE, ActiveSheet.UsedRange.Columns.Count)
    
    ' - Get range for each work date of this employee
    For j = 1 To UBound(strUniqueWorkDateForOneEmp)
        nrOfRealRowsSameWorkDate = 0
        Call getAllRowsHasSameOneColValInArray(tblCfg.titleRowNo, tblCfg.workDateColNo, strUniqueWorkDateForOneEmp(j), _
                                            oneEmpRecords, nrOfRealRowsForEmp, _
                                            arrSameWorkDate, nrOfRealRowsSameWorkDate)
        'Calculate total hours in one day.
        Dim otHourInOneDay As Single: otHourInOneDay = 0
        Dim regHourInOneDay As Single: regHourInOneDay = 0
        Dim paidLeaveHourInOneDay As Single: paidLeaveHourInOneDay = 0
        Dim normalHourInOneDay As Single: normalHourInOneDay = 0
        Dim currWkDate As Date: currWkDate = CDate(strUniqueWorkDateForOneEmp(j))
        Dim strHourType As String
        Dim strTimeCode As String
        '@TODO: use the first row as calc group because we asume one employee belongs to one calc group
        Dim strCalcGrp As String: strCalcGrp = arrSameWorkDate(1, tblCfg.calcGrpColNo)

        For k = 1 To nrOfRealRowsSameWorkDate
            strHourType = arrSameWorkDate(k, tblCfg.hourTypeColNo)
            strTimeCode = arrSameWorkDate(k, tblCfg.timeCodeColNo)
            
            'Non duplicate (PREM), Non STA will be taken as normal hour for paid leave
            If strHourType <> "REG_PREM" Then
                If Not isStaWorkDate(tblCfg, oneEmpRecords, nrOfRealRowsForEmp, currWkDate) Then
                    normalHourInOneDay = normalHourInOneDay + arrSameWorkDate(k, tblCfg.hoursColNo)
                End If
            End If
            
            If isNormalOverTimeHour(tblCfg, oneEmpRecords, nrOfRealRowsForEmp, currWkDate, strHourType, strTimeCode) Then
                otHourInOneDay = otHourInOneDay + arrSameWorkDate(k, tblCfg.hoursColNo)
            ElseIf isRegularHour(strHourType, strTimeCode) Then
                regHourInOneDay = regHourInOneDay + arrSameWorkDate(k, tblCfg.hoursColNo)
            End If
            
            If isTimeCodePaidLeave(strTimeCode) Then
                paidLeaveHourInOneDay = paidLeaveHourInOneDay + arrSameWorkDate(k, tblCfg.hoursColNo)
            End If
        Next k

        If (chkType = OT_BY_DAY_CHK_PAID_LEAVE) Then
            Dim dailyHourCriteria As Single: dailyHourCriteria = dictDailyHoursForCalcGrp(strCalcGrp)
            If strCalcGrp = "CAW 4 & 4" Then
                dailyHourCriteria = dailyHourCriteria + 0.29 'Add extra 0.29 for reconcile
            End If
            'If there is paid leave hour for a day, all normal hours should not exceed a certain criteria
            If (paidLeaveHourInOneDay > 0) And (normalHourInOneDay > dailyHourCriteria) Then
                errMsg = "PaidLeaveHourInOneDay > 0, then TotalHours should be <= " & dailyHourCriteria
                Call colorEmpPLOneDay(tblCfg, arrAllValidData, oneEmpRecords, strEmpId, currWkDate, errMsg)
            End If
        ElseIf (chkType = OT_BY_DAY_CHK_REG_HOURS_OT) Then
            'This branch not used anymore
        End If
    Next j

End Function

Function handleBiweeklyOnlyCalcGrpEmp(ByRef tblCfg As OverTimeTableCfgType, ByRef strEmpId As String, _
                                        ByRef oneEmpRecords As Variant, nrOfRealRowsForEmp As Integer, _
                                        ByRef arrAllValidData As Variant)
    Dim errMsg As String
    Dim regHourInWholePay As Single
    Dim otHourInWholePay As Single
    
    Call calcEmpHoursInWholePay(tblCfg, oneEmpRecords, nrOfRealRowsForEmp, regHourInWholePay, otHourInWholePay)

    'If there is overtime in this pay, the regular work hour has to be more than a criteria (e.g. 80 hours)
    If (otHourInWholePay > 0) And (regHourInWholePay < REG_WORK_HOURS_LOW_LIMIT_WHOLE_PAY) Then
        errMsg = "OvertimeInWholePay > 0, then RegHoursInWholePay should be >= " & REG_WORK_HOURS_LOW_LIMIT_WHOLE_PAY
        Call colorEmpOTWholePay(tblCfg, arrAllValidData, oneEmpRecords, nrOfRealRowsForEmp, strEmpId, errMsg)
    End If
End Function

Function handleDailyWeeklyBiweeklyCalcGrpEmp(ByRef tblCfg As OverTimeTableCfgType, ByRef strEmpId As String, _
                                            ByRef oneEmpRecords As Variant, nrOfRealRowsForEmp As Integer, _
                                            ByRef payPeriodStartDate As Date, _
                                            ByRef arrAllValidData As Variant, ByRef dictDailyHoursForCalcGrp As Object)
    Dim regHourInWholePay As Single
    Dim otHourInWholePay As Single
    
    Dim regHourInWk1 As Single
    Dim otHourInWk1 As Single
    Dim regHourInWk2 As Single
    Dim otHourInWk2 As Single

    Call calcEmpHoursIn2Weeks(tblCfg, oneEmpRecords, nrOfRealRowsForEmp, payPeriodStartDate, _
                                regHourInWk1, otHourInWk1, regHourInWk2, otHourInWk2)
    
    regHourInWholePay = regHourInWk1 + regHourInWk2
    otHourInWholePay = otHourInWk1 + otHourInWk2
    
    Dim j As Integer
    'Get Work Date For this Employee
    Dim strUniqueWorkDateForOneEmp() As String

    ' - Get all unique work date strings for this employee
    Call getAllUniqueStrsByColInArray(tblCfg.titleRowNo, tblCfg.workDateColNo, _
                                    oneEmpRecords, nrOfRealRowsForEmp, _
                                    strUniqueWorkDateForOneEmp)
    Dim regHourInWeek As Single
    Dim currWorkDate As Date
    Dim regHourInOneDay As Single
    Dim otHourInOneDay As Single
    Dim wkHourInOneDay As Single
    Dim strCalcGrp As String: strCalcGrp = oneEmpRecords(1, tblCfg.calcGrpColNo)
    Dim dailHourCriteria As Single: dailHourCriteria = dictDailyHoursForCalcGrp(strCalcGrp)
    Dim errMsg As String

    For j = 1 To UBound(strUniqueWorkDateForOneEmp)
        currWorkDate = CDate(strUniqueWorkDateForOneEmp(j))
        Call calcEmpHoursInOneDay(tblCfg, oneEmpRecords, nrOfRealRowsForEmp, currWorkDate, regHourInOneDay, otHourInOneDay, wkHourInOneDay)

        'when there is overtime, daily work hour should NOT be less than the criteria
        If (otHourInOneDay > 0) Then
            If isWorkDateFirstWeek(currWorkDate, payPeriodStartDate) Then
                regHourInWeek = regHourInWk1
            Else
                regHourInWeek = regHourInWk2
            End If
            
            If (regHourInOneDay >= dailHourCriteria) Or (regHourInWeek >= REG_HOURS_CRITERIA_WEEKLY) Or _
                (regHourInWholePay >= REG_HOURS_CRITERIA_BIWEEKLY) Or _
                (isStaWorkDate(tblCfg, oneEmpRecords, nrOfRealRowsForEmp, currWorkDate)) Then
                'This is the right branch, do nothing
            Else
                'Create error message as hint for users.
                errMsg = "OvertimeInOneDay > 0, RegHour Threshold is (" & dailHourCriteria & "/" _
                         & REG_HOURS_CRITERIA_WEEKLY & "/" & REG_HOURS_CRITERIA_BIWEEKLY & ")"
                'Error Branch, color this day
                Call colorEmpOTOneDay(tblCfg, arrAllValidData, oneEmpRecords, nrOfRealRowsForEmp, strEmpId, currWorkDate, errMsg)
            End If
        End If
    Next j
End Function

Function handleDailyOnlyCalcGrpEmp(ByRef tblCfg As OverTimeTableCfgType, ByRef strEmpId As String, _
                                    ByRef oneEmpRecords As Variant, nrOfRealRowsForEmp As Integer, _
                                    ByRef payPeriodStartDate As Date, _
                                    ByRef arrAllValidData As Variant, ByRef dictDailyHoursForCalcGrp As Object)
    Dim j As Integer
    'Get Work Date For this Employee
    Dim strUniqueWorkDateForOneEmp() As String

    ' - Get all unique work date strings for this employee
    Call getAllUniqueStrsByColInArray(tblCfg.titleRowNo, tblCfg.workDateColNo, oneEmpRecords, _
                                      nrOfRealRowsForEmp, strUniqueWorkDateForOneEmp)
    Dim errMsg As String
    Dim currWorkDate As Date
    Dim regHourInOneDay As Single
    Dim otHourInOneDay As Single
    Dim wkHourInOneDay As Single
    Dim strCalcGrp As String: strCalcGrp = oneEmpRecords(1, tblCfg.calcGrpColNo)
    Dim dailHourCriteria As Single: dailHourCriteria = dictDailyHoursForCalcGrp(strCalcGrp)
    
    For j = 1 To UBound(strUniqueWorkDateForOneEmp)
        currWorkDate = CDate(strUniqueWorkDateForOneEmp(j))
        Call calcEmpHoursInOneDay(tblCfg, oneEmpRecords, nrOfRealRowsForEmp, currWorkDate, regHourInOneDay, otHourInOneDay, wkHourInOneDay)

        'when there is overtime, daily work hour should NOT be less than the criteria
        If (otHourInOneDay > 0) Then
            '(regHourInOneDay <= DEFAULT_FLOAT_PRECISION) is equal to (Abs(regHourInOneDay - 0) <= DEFAULT_FLOAT_PRECISION)
            If (regHourInOneDay >= dailHourCriteria) Or (regHourInOneDay <= DEFAULT_FLOAT_PRECISION) Then
                'This is the right branch, do nothing
            Else
                errMsg = "OvertimeInOneDay > 0, then Regular Hour should be 0 or >= " & dailHourCriteria
                Call colorEmpOTOneDay(tblCfg, arrAllValidData, oneEmpRecords, nrOfRealRowsForEmp, strEmpId, currWorkDate, errMsg)
            End If
        End If
    Next j
End Function

Function chkDailyHoursUpperLimitForOneEmp(ByRef tblCfg As OverTimeTableCfgType, ByRef strEmpId As String, _
                                        ByRef oneEmpRecords As Variant, nrOfRealRowsForEmp As Integer, _
                                        ByRef arrAllValidData As Variant, ByRef dictDailyHoursUpperLimitForCalcGrp As Object)
    'Get Work Date For this Employee
    Dim strUniqueWorkDateForOneEmp() As String

    ' - Get all unique work date strings for this employee
    Call getAllUniqueStrsByColInArray(tblCfg.titleRowNo, tblCfg.workDateColNo, oneEmpRecords, _
                                      nrOfRealRowsForEmp, strUniqueWorkDateForOneEmp)
    Dim errMsg As String
    Dim currWorkDate As Date
    Dim tmpWorkDate As Date
    Dim tmpHourType As String
    Dim regHourInOneDay As Single
    Dim otHourInOneDay As Single
    Dim wkHourInOneDay As Single
    Dim strCalcGrp As String: strCalcGrp = oneEmpRecords(1, tblCfg.calcGrpColNo)
    Dim dailHourUpperLimit As Single: dailHourUpperLimit = dictDailyHoursUpperLimitForCalcGrp(strCalcGrp)
    Dim j As Integer
    Dim k As Integer
    
    For j = 1 To UBound(strUniqueWorkDateForOneEmp)
        currWorkDate = CDate(strUniqueWorkDateForOneEmp(j))
        Call calcEmpHoursInOneDay(tblCfg, oneEmpRecords, nrOfRealRowsForEmp, currWorkDate, regHourInOneDay, otHourInOneDay, wkHourInOneDay)

        'When daily total hours is higher than upper threshold, color it
        dailyTotalHours = otHourInOneDay + wkHourInOneDay
        If (dailyTotalHours > dailHourUpperLimit) Then
            errMsg = "DailyTotalHour = " & dailyTotalHours & ", UpperLimit = " & dailHourUpperLimit
            'Find the first row in that day and color it
            For k = 1 To nrOfRealRowsForEmp
                tmpWorkDate = oneEmpRecords(k, tblCfg.workDateColNo)
                tmpHourType = oneEmpRecords(k, tblCfg.hourTypeColNo)
                If (tmpWorkDate = currWorkDate) And (tmpHourType <> "REG_PREM") Then
                    Exit For
                End If
            Next k
            Call colorEmpOneRow(tblCfg, strEmpId, oneEmpRecords, nrOfRealRowsForEmp, k, arrAllValidData, errMsg)
        End If
    Next j
End Function

Function handleEmpTimeRecords(ByRef tblCfg As OverTimeTableCfgType, ByRef strEmpId As String, _
                              ByRef oneEmpRecords As Variant, nrOfRealRowsForEmp As Integer, _
                              ByRef payPeriodStartDate As Date, _
                              ByRef arrAllValidData As Variant, ByRef dictDailyHoursForCalcGrp As Object)
    '@TODO: we assume one employee only belongs to one calc Group
    Dim strCalcGrp As String: strCalcGrp = oneEmpRecords(1, tblCfg.calcGrpColNo)

    If strCalcGrp = "CAW FT OVER 80" Then
        Call handleBiweeklyOnlyCalcGrpEmp(tblCfg, strEmpId, oneEmpRecords, nrOfRealRowsForEmp, arrAllValidData)
    ElseIf strCalcGrp = "CAW 4 & 4" Then
        Call handleDailyOnlyCalcGrpEmp(tblCfg, strEmpId, oneEmpRecords, nrOfRealRowsForEmp, payPeriodStartDate, _
                                        arrAllValidData, dictDailyHoursForCalcGrp)
    ElseIf strCalcGrp = "CAW 6 WK ROTATION" Then
        'Do Nothing because this type has no enough info in the table and it has to be done manually
    Else    'Other five calc groups: CAW PT, CAW 5 & 2, CAW CASUAL, CAW FT OVER 40, CAW FT TEMP
        Call handleDailyWeeklyBiweeklyCalcGrpEmp(tblCfg, strEmpId, oneEmpRecords, nrOfRealRowsForEmp, payPeriodStartDate, _
                                                arrAllValidData, dictDailyHoursForCalcGrp)
    End If

End Function

Function getPayPeriod(ByRef payPeriodStartDate As Date, nrPayPeriodOffset As Integer)
    Dim firstPayDay As Date: firstPayDay = DateValue(DEFAULT_FIRST_PAY_DAY_IN_2019)
    Dim offset As Integer

    'Guess the most possible pay day the user may input as default.
    'User will run this command 2 Pay Period later from now
    offset = DateDiff("d", firstPayDay, Now())
    offset = offset - (offset Mod DEFAULT_PAY_PERIOD_IN_DAYS) + nrPayPeriodOffset * DEFAULT_PAY_PERIOD_IN_DAYS
    firstPayDay = DateAdd("d", offset, firstPayDay)

    'Pay Day is the input from User
    Dim strInputDate As String
    strInputDate = InputBox("Input Pay Day in format YYYY/MM/DD" & vbCrLf & vbCrLf & "Note: Pay period starts from the Pay Day you input subtracts 27", _
                            "User Input", Format(firstPayDay, "yyyy-mm-dd"))

    If strInputDate = vbNullString Then
        MsgBox "Cancel button pressed and Will Do Nothing!"
        Exit Function
    End If
    
    'Update pay start date according to user's input
    payPeriodStartDate = DateAdd("d", DEFAULT_PAY_DAY_START_OFFSET, CDate(strInputDate))
End Function

Sub OverTime()
    'Turn off some global switches in order to run faster
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    'Clear background/Font color of the table
    ActiveSheet.Cells.Interior.ColorIndex = 0
    ActiveSheet.Cells.Font.ColorIndex = 0
    
    Dim payPeriodStartDate As Date: payPeriodStartDate = 0
    
    'Get Pay period date
    Call getPayPeriod(payPeriodStartDate, 1)
    If payPeriodStartDate = 0 Then
        Exit Sub
    End If
    
    Dim StartTime As Double
    StartTime = Timer
    
    Dim lastRowNo As Long: lastRowNo = ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row
    Dim lastColNo As Long: lastColNo = ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Column
    
    Dim tblCfg As OverTimeTableCfgType
    
    'Find the correct column we care
    tblCfg.firstDataRowNo = -1      'Initilize it to an invalid value first
    tblCfg.errMsgColNo = -1
    Call getTableCfg(lastRowNo, lastColNo, tblCfg)

    'Error handling to prevent operating on a wrong Table
    If Not chkTbleCfg(tblCfg) Then
        MsgBox "Unable to process this table by this Command because " & vbCrLf & vbCrLf & _
               "Column  'Calc Group Name', 'Employee ID', 'Hour Type', 'Work Date', 'Time Code' and 'Hours' are mandatory!" & vbCrLf & vbCrLf & _
               "Please Check if you Opened a correct Table WorkSeet!"
        Exit Sub
    End If
    
    'Clear or Create a column to save result
    If (tblCfg.errMsgColNo <= 0) Then
        tblCfg.errMsgColNo = lastColNo + 1
        lastColNo = lastColNo + 1
    End If
    ActiveSheet.Columns(tblCfg.errMsgColNo).ClearContents
    ActiveSheet.Cells(tblCfg.titleRowNo, tblCfg.errMsgColNo) = DEFAULT_ERR_MSG_COL_NAME
    
    'Create daily hours criteria for each calc group
    Dim dictDailyHoursForCalcGrp As Object
    Set dictDailyHoursForCalcGrp = CreateObject("Scripting.Dictionary")
    Call initCalcGrpDailyHoursLowerLimit(dictDailyHoursForCalcGrp)
    
    Dim nrRowsAllValidData As Integer: nrRowsAllValidData = lastRowNo - tblCfg.firstDataRowNo + 1
    
    'Delete Existed debug output text files
    Dim dbgFile As String
    dbgFile = Application.ActiveWorkbook.path & "\*.txt"
    If Len(Dir(dbgFile)) > 0 Then
        SetAttr dbgFile, vbNormal
        Kill dbgFile
    End If

    Debug.Print "OverTime Start, Rows/Colums= " & ActiveSheet.UsedRange.Rows.Count & "/" & ActiveSheet.UsedRange.Columns.Count

    'Read all valid data to memory. @TODO: we may combine the following two steps into one to save some memory in Future?
    ' - First, Read data to a range by removing unneeded headers
    '     Using Resize to remove headers in the first few rows that we do not need
    Dim rngAllValidData As Range
    Set rngAllValidData = ActiveSheet.UsedRange.Rows(tblCfg.firstDataRowNo).Resize(nrRowsAllValidData)
    Set rngAllValidData = rngAllValidData.Resize(nrRowsAllValidData, lastColNo)
    
    ' - Second, Read data to an array to make later process faster
    Dim arrAllValidData As Variant: arrAllValidData = rngAllValidData.Value
    
    'Find all unique employee IDs. allEmpIDs is a 1-D array
    Dim allEmpIDs() As String
    Call getAllUniqueStrsByColInArray(tblCfg.titleRowNo, tblCfg.employeeIdColNo, arrAllValidData, nrRowsAllValidData, allEmpIDs)

    Dim i As Integer
    Dim nrOfRealRowsForEmp As Integer
    Dim oneEmpRecords As Variant
    ReDim oneEmpRecords(MAX_ROWS_FOR_ONE_EMPLOYEE, lastColNo)
    
    'handle each employee
    For i = 1 To UBound(allEmpIDs)
        'Find all rows of this employee
        nrOfRealRowsForEmp = 0
        Call getAllRowsHasSameOneColValInArray(tblCfg.titleRowNo, tblCfg.employeeIdColNo, allEmpIDs(i), _
                                            arrAllValidData, nrRowsAllValidData, _
                                            oneEmpRecords, nrOfRealRowsForEmp)
        'Error Handling
        If nrOfRealRowsForEmp = 0 Then
            MsgBox "Employee Records for Employee ID: " & allEmpIDs(i) & "Is Empty"
            Exit For
        End If
        
        'Process this employee's time records
        Call handleEmpTimeRecords(tblCfg, allEmpIDs(i), oneEmpRecords, nrOfRealRowsForEmp, _
                                payPeriodStartDate, arrAllValidData, dictDailyHoursForCalcGrp)
    Next i

    Debug.Print "OverTime End, Rows(Sheet)/Rows(Real)/Colums= " & ActiveSheet.UsedRange.Rows.Count & _
                "/" & rngAllValidData.Rows.Count & "/" & ActiveSheet.UsedRange.Columns.Count
    
    MsgBox ">>>>     OverTime     <<<<" & vbCrLf & vbCrLf & vbCrLf & _
            "Done Normally in  " & Format(Timer - StartTime, "0.00") & "  seconds!"
    
    'Turn On screen updates after running to the end
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub

Private Function handleEmpPaidLeaveOT(ByRef tblCfg As OverTimeTableCfgType, ByRef strEmpId As String, _
                              ByRef oneEmpRecords As Variant, nrOfRealRowsForEmp As Integer, _
                              ByRef arrAllValidData As Variant, ByRef dictDailyHoursForCalcGrp As Object)
    'Handle employee hours by day and check paid leave
    Call hanldeEmpHoursByDay(tblCfg, strEmpId, oneEmpRecords, nrOfRealRowsForEmp, arrAllValidData, _
                             OT_BY_DAY_CHK_PAID_LEAVE, dictDailyHoursForCalcGrp)
End Function

'dict object should be created outside
Private Function initCalcGrpDailyHoursUpperLimit(ByRef dict As Object)
    dict.Add "CAW 6 WK ROTATION", 15
    dict.Add "CAW 4 & 4", 15
    dict.Add "CAW 5 & 2", 15
    dict.Add "CAW FT OVER 40", 15
    dict.Add "CAW FT OVER 80", 15
    dict.Add "CAW PT", 12
    dict.Add "CAW FT TEMP", 12
    dict.Add "CAW CASUAL", 10
End Function

'dict object should be created outside
Private Function initCalcGrpDailyHoursLowerLimit(ByRef dict As Object)
    dict.Add "CAW 6 WK ROTATION", 12
    dict.Add "CAW 4 & 4", (11.5) 'there may be reconcile for 4&4, so an extra 0.29 is added when needed
    dict.Add "CAW 5 & 2", 8
    dict.Add "CAW FT OVER 40", 10
    dict.Add "CAW FT OVER 80", 12
    dict.Add "CAW PT", 8
    dict.Add "CAW FT TEMP", 8
    dict.Add "CAW CASUAL", 8
End Function

'dict object should be created outside
Private Function initCalcGrpTotalRegHoursLowerLimit(ByRef dict As Object)
    dict.Add "CAW 6 WK ROTATION", 0 'we don't care lower limit for this type
    dict.Add "CAW 4 & 4", 0
    dict.Add "CAW 5 & 2", 76
    dict.Add "CAW FT OVER 40", 76
    dict.Add "CAW FT OVER 80", 76
    dict.Add "CAW PT", 30
    dict.Add "CAW FT TEMP", 30
    dict.Add "CAW CASUAL", 0
End Function

Sub PaidLeave()
    'Turn off some global switches in order to run faster
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    'Clear background color of the table to 'No Fill' first
    ActiveSheet.Cells.Interior.ColorIndex = 0
    ActiveSheet.Cells.Font.ColorIndex = 0

    'Measure time performance
    Dim StartTime As Double
    StartTime = Timer
    
    Dim tblCfg As OverTimeTableCfgType
    Dim lastRowNo As Long: lastRowNo = ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row
    Dim lastColNo As Long: lastColNo = ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Column

    'Find the correct column we care
    tblCfg.firstDataRowNo = -1      'Initilize it to an invalid value first
    tblCfg.errMsgColNo = -1
    Call getTableCfg(lastRowNo, lastColNo, tblCfg)
    
    'Error handling to prevent operating on a wrong Table
    If Not chkTbleCfg(tblCfg) Then
        MsgBox "Unable to process this table by this Command Because Critical Columns is missing" & vbCrLf & vbCrLf & _
                "Column  'Calc Group Name', 'Employee ID', 'Hour Type', 'Work Date', 'Time Code' and 'Hours' are mandatory" & vbCrLf & vbCrLf & _
                "Please Check if you Opened a correct Table WorkSeet!"
        Exit Sub
    End If

    'Clear or Create a column to save result
    If (tblCfg.errMsgColNo <= 0) Then
        tblCfg.errMsgColNo = lastColNo + 1
        lastColNo = lastColNo + 1
    End If
    ActiveSheet.Columns(tblCfg.errMsgColNo).ClearContents
    ActiveSheet.Cells(tblCfg.titleRowNo, tblCfg.errMsgColNo) = DEFAULT_ERR_MSG_COL_NAME

    'Create daily hours criteria for each calc group
    Dim dictDailyHoursForCalcGrp As Object
    Set dictDailyHoursForCalcGrp = CreateObject("Scripting.Dictionary")
    Call initCalcGrpDailyHoursLowerLimit(dictDailyHoursForCalcGrp)
        
    Dim nrRowsAllValidData As Integer: nrRowsAllValidData = lastRowNo - tblCfg.firstDataRowNo + 1
    
    'Read all valid data to memory. @TODO: we may combine the following two steps into one to save some memory in Future?
    ' - First, Read data to a range by removing unneeded headers
    '     Using Resize to remove headers in the first few rows that we do not need
    Dim rngAllValidData As Range
    Set rngAllValidData = ActiveSheet.Rows(tblCfg.firstDataRowNo).Resize(nrRowsAllValidData)
    Set rngAllValidData = rngAllValidData.Resize(nrRowsAllValidData, lastColNo)
    
    ' - Second, Read data to an array to make later process faster
    Dim arrAllValidData As Variant: arrAllValidData = rngAllValidData.Value
    
    'Find all unique employee IDs. allEmpIDs is a 1-D array
    Dim allEmpIDs() As String
    Call getAllUniqueStrsByColInArray(tblCfg.titleRowNo, tblCfg.employeeIdColNo, arrAllValidData, nrRowsAllValidData, allEmpIDs)
    
    Dim i As Integer
    Dim nrOfRealRowsForEmp As Integer
    Dim oneEmpRecords As Variant
    ReDim oneEmpRecords(MAX_ROWS_FOR_ONE_EMPLOYEE, ActiveSheet.UsedRange.Columns.Count)
    
    'handle each employee
    For i = 1 To UBound(allEmpIDs)
        'Find all rows of this employee
        nrOfRealRowsForEmp = 0
        Call getAllRowsHasSameOneColValInArray(tblCfg.titleRowNo, tblCfg.employeeIdColNo, allEmpIDs(i), _
                                            arrAllValidData, nrRowsAllValidData, _
                                            oneEmpRecords, nrOfRealRowsForEmp)
        'Error Handling
        If nrOfRealRowsForEmp = 0 Then
            MsgBox "Employee Records for Employee ID: " & allEmpIDs(i) & "Is Empty"
            Exit For
        End If

        'Process this employee's time records
        Call handleEmpPaidLeaveOT(tblCfg, allEmpIDs(i), oneEmpRecords, nrOfRealRowsForEmp, arrAllValidData, dictDailyHoursForCalcGrp)
    Next i

    
    MsgBox ">>>>     PaidLeave     <<<<" & vbCrLf & vbCrLf & vbCrLf & _
            "Done Normally in  " & Format(Timer - StartTime, "0.00") & "  seconds!"
    
    'Turn On screen updates after running to the end
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub

'Total OT calculation
Sub TotalOT()
    'Turn off some global switches in order to run faster
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    'Clear background color of the table to 'No Fill' first
    ActiveSheet.Cells.Interior.ColorIndex = 0
    ActiveSheet.Cells.Font.ColorIndex = 0
    
    'Measure time performance
    Dim StartTime As Double
    StartTime = Timer
    
    Dim tblCfg As OverTimeTableCfgType
    Dim lastRowNo As Long: lastRowNo = ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row
    Dim lastColNo As Long: lastColNo = ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Column

    'Find the correct column we care
    tblCfg.firstDataRowNo = -1      'Initilize it to an invalid value first
    tblCfg.errMsgColNo = -1
    Call getTableCfg(lastRowNo, lastColNo, tblCfg)
    
    'Error handling to prevent operating on a wrong Table
    If Not chkTbleCfg(tblCfg) Then
        MsgBox "Unable to process this table by this Command Because Critical Columns is missing" & vbCrLf & vbCrLf & _
                "Column  'Calc Group Name', 'Employee ID', 'Hour Type', 'Work Date', 'Time Code' and 'Hours' are mandatory" & vbCrLf & vbCrLf & _
                "Please Check if you Opened a correct Table WorkSeet!"
        Exit Sub
    End If

    'Clear or Create a column to save result
    If (tblCfg.errMsgColNo < 0) Then
        tblCfg.errMsgColNo = lastColNo + 1
        lastColNo = lastColNo + 1
    End If
    ActiveSheet.Columns(tblCfg.errMsgColNo).ClearContents
    ActiveSheet.Cells(tblCfg.titleRowNo, tblCfg.errMsgColNo) = DEFAULT_ERR_MSG_COL_NAME

    Dim nrRowsAllValidData As Integer: nrRowsAllValidData = lastRowNo - tblCfg.firstDataRowNo + 1
    
    'Read all valid data to memory. @TODO: we may combine the following two steps into one to save some memory in Future?
    ' - First, Read data to a range by removing unneeded headers
    '     Using Resize to remove headers in the first few rows that we do not need
    Dim rngAllValidData As Range
    Set rngAllValidData = ActiveSheet.Rows(tblCfg.firstDataRowNo).Resize(nrRowsAllValidData)
    Set rngAllValidData = rngAllValidData.Resize(nrRowsAllValidData, lastColNo)
    
    ' - Second, Read data to an array to make later process faster
    Dim arrAllValidData As Variant: arrAllValidData = rngAllValidData.Value
    
    'Find all unique employee IDs. allEmpIDs is a 1-D array
    Dim allEmpIDs() As String
    Call getAllUniqueStrsByColInArray(tblCfg.titleRowNo, tblCfg.employeeIdColNo, arrAllValidData, nrRowsAllValidData, allEmpIDs)
    
    Dim i As Integer
    Dim nrOfRealRowsForEmp As Integer
    Dim oneEmpRecords As Variant
    ReDim oneEmpRecords(MAX_ROWS_FOR_ONE_EMPLOYEE, ActiveSheet.UsedRange.Columns.Count)
    
    Dim regHourInWholePay As Single
    Dim otHourInWholePay As Single
    Dim strCalcGrp As String
    Dim otAbnormal As Boolean
    Dim errMsg As String
    
    'handle each employee
    For i = 1 To UBound(allEmpIDs)
        'Find all rows of this employee
        nrOfRealRowsForEmp = 0
        Call getAllRowsHasSameOneColValInArray(tblCfg.titleRowNo, tblCfg.employeeIdColNo, allEmpIDs(i), _
                                            arrAllValidData, nrRowsAllValidData, _
                                            oneEmpRecords, nrOfRealRowsForEmp)
        'Error Handling
        If nrOfRealRowsForEmp = 0 Then
            MsgBox "Employee Records for Employee ID: " & allEmpIDs(i) & "Is Empty"
            Exit For
        End If

        'Process this employee's time records
        Call calcEmpHoursInWholePay(tblCfg, oneEmpRecords, nrOfRealRowsForEmp, _
                                    regHourInWholePay, otHourInWholePay)
        
        strCalcGrp = oneEmpRecords(1, tblCfg.calcGrpColNo)
        
        otAbnormal = False
        errMsg = "Total OT = " & otHourInWholePay & ", Total Reg = " & regHourInWholePay & ", Threshold = "

        If (strCalcGrp = "CAW 6 WK ROTATION") Or (strCalcGrp = "CAW 4 & 4") Then
            'Do nothing for these two
        ElseIf (strCalcGrp = "CAW 5 & 2") Or (strCalcGrp = "CAW FT OVER 40") Or _
                (strCalcGrp = "CAW FT OVER 80") Or (strCalcGrp = "CAW PT") Or _
                (strCalcGrp = "CAW FT TEMP") Then
            If (otHourInWholePay >= ABM_OT_HOURS_FOR_OTHERS) Then
                otAbnormal = True
                errMsg = errMsg & ABM_OT_HOURS_FOR_OTHERS
            End If
        ElseIf (strCalcGrp = "CAW CASUAL") Then
            If (otHourInWholePay >= ABM_OT_HOURS_FOR_CASUAL) Then
                otAbnormal = True
                errMsg = errMsg & ABM_OT_HOURS_FOR_CASUAL
            End If
        End If
        
        'when total overtime is abnormal, color whole pay for this employee
        If otAbnormal = True Then
            Call colorEmpOTWholePay(tblCfg, arrAllValidData, oneEmpRecords, nrOfRealRowsForEmp, allEmpIDs(i), errMsg)
        End If
    Next i
    
    MsgBox ">>>>     TotalOT     <<<<" & vbCrLf & vbCrLf & vbCrLf & _
           "Done Normally in  " & Format(Timer - StartTime, "0.00") & "  seconds!"
    
    'Turn On screen updates after running to the end
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
End Sub


'Check rounding: if the last 2 digits are equal to 0.25/0.50/0.75/0
Sub Rounding()
    'Turn off some global switches in order to run faster
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    'Clear background color of the table to 'No Fill' first
    ActiveSheet.Cells.Interior.ColorIndex = 0
    ActiveSheet.Cells.Font.ColorIndex = 0
    
    Dim payPeriodStartDate As Date: payPeriodStartDate = 0
    
    'Get Pay period date
    Call getPayPeriod(payPeriodStartDate, 1)
    If payPeriodStartDate = 0 Then
        Exit Sub
    End If
    
    'Measure time performance
    Dim StartTime As Double
    StartTime = Timer
    
    Dim tblCfg As OverTimeTableCfgType
    Dim lastRowNo As Long: lastRowNo = ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row
    Dim lastColNo As Long: lastColNo = ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Column

    'Find the correct column we care
    tblCfg.firstDataRowNo = -1      'Initilize it to an invalid value first
    tblCfg.errMsgColNo = -1
    Call getTableCfg(lastRowNo, lastColNo, tblCfg)
    
    'Error handling to prevent operating on a wrong Table
    
    If Not chkTbleCfg(tblCfg) Then
        MsgBox "Unable to process this table by this Command Because Critical Columns is missing" & vbCrLf & vbCrLf & _
                "Column  'Calc Group Name', 'Employee ID', 'Hour Type', 'Work Date', 'Time Code' and 'Hours' are mandatory" & vbCrLf & vbCrLf & _
                "Please Check if you Opened a correct WorkSeet!"
        Exit Sub
    End If
    
    'Clear or Create a column to save result
    If (tblCfg.errMsgColNo < 0) Then
        tblCfg.errMsgColNo = lastColNo + 1
        lastColNo = lastColNo + 1
    End If
    ActiveSheet.Columns(tblCfg.errMsgColNo).ClearContents
    ActiveSheet.Cells(tblCfg.titleRowNo, tblCfg.errMsgColNo) = DEFAULT_ERR_MSG_COL_NAME

    Dim nrRowsAllValidData As Integer: nrRowsAllValidData = lastRowNo - tblCfg.firstDataRowNo + 1
    
    'Read all valid data to memory. @TODO: we may combine the following two steps into one to save some memory in Future?
    ' - First, Read data to a range by removing unneeded headers
    '     Using Resize to remove headers in the first few rows that we do not need
    Dim rngAllValidData As Range
    Set rngAllValidData = ActiveSheet.Rows(tblCfg.firstDataRowNo).Resize(nrRowsAllValidData)
    Set rngAllValidData = rngAllValidData.Resize(nrRowsAllValidData, lastColNo)
    
    ' - Second, Read data to an array to make later process faster
    Dim arrAllValidData As Variant: arrAllValidData = rngAllValidData.Value
    
    'Find all unique employee IDs. allEmpIDs is a 1-D array
    Dim allEmpIDs() As String
    Call getAllUniqueStrsByColInArray(tblCfg.titleRowNo, tblCfg.employeeIdColNo, arrAllValidData, nrRowsAllValidData, allEmpIDs)
    
    Dim nrOfRealRowsForEmp As Integer
    Dim oneEmpRecords As Variant
    ReDim oneEmpRecords(MAX_ROWS_FOR_ONE_EMPLOYEE, ActiveSheet.UsedRange.Columns.Count)
    Dim i As Integer
    
    'handle each employee
    For i = 1 To UBound(allEmpIDs)
        'Find all rows of this employee
        nrOfRealRowsForEmp = 0
        Call getAllRowsHasSameOneColValInArray(tblCfg.titleRowNo, tblCfg.employeeIdColNo, allEmpIDs(i), _
                                            arrAllValidData, nrRowsAllValidData, _
                                            oneEmpRecords, nrOfRealRowsForEmp)
        'Error Handling
        If nrOfRealRowsForEmp = 0 Then
            MsgBox "Employee Records for Employee ID: " & allEmpIDs(i) & "Is Empty"
            Exit For
        End If
        
        'Process this employee's time records
        Call handleEmpHoursRounding(tblCfg, allEmpIDs(i), oneEmpRecords, nrOfRealRowsForEmp, payPeriodStartDate, arrAllValidData)
    Next i
    
    MsgBox ">>>>     Rounding     <<<<" & vbCrLf & vbCrLf & vbCrLf & _
            "Done Normally in  " & Format(Timer - StartTime, "0.00") & "  seconds!"
    
    'Turn On screen updates after running to the end
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub

'Check daily total hours (except PREM and STA) < column value in the table
Sub DailyHours()
    'Turn off some global switches in order to run faster
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    'Clear background color of the table to 'No Fill' first
    ActiveSheet.Cells.Interior.ColorIndex = 0
    ActiveSheet.Cells.Font.ColorIndex = 0
    
    'Measure time performance
    Dim StartTime As Double
    StartTime = Timer
    
    Dim tblCfg As OverTimeTableCfgType
    Dim lastRowNo As Long: lastRowNo = ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row
    Dim lastColNo As Long: lastColNo = ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Column

    'Find the correct column we care
    tblCfg.firstDataRowNo = -1      'Initilize it to an invalid value first
    tblCfg.errMsgColNo = -1
    Call getTableCfg(lastRowNo, lastColNo, tblCfg)
    
    'Error handling to prevent operating on a wrong Table
    If Not chkTbleCfg(tblCfg) Then
        MsgBox "Unable to process this table by this Command Because Critical Columns is missing" & vbCrLf & vbCrLf & _
                "Column  'Calc Group Name', 'Employee ID', 'Hour Type', 'Work Date', 'Time Code' and 'Hours' are mandatory" & vbCrLf & vbCrLf & _
                "Please Check if you Opened a correct WorkSeet!"
        Exit Sub
    End If
    
    'Clear or Create a column to save result
    If (tblCfg.errMsgColNo < 0) Then
        tblCfg.errMsgColNo = lastColNo + 1
        lastColNo = lastColNo + 1
    End If
    ActiveSheet.Columns(tblCfg.errMsgColNo).ClearContents
    ActiveSheet.Cells(tblCfg.titleRowNo, tblCfg.errMsgColNo) = DEFAULT_ERR_MSG_COL_NAME

    Dim nrRowsAllValidData As Integer: nrRowsAllValidData = lastRowNo - tblCfg.firstDataRowNo + 1
    
    'Read all valid data to memory. @TODO: we may combine the following two steps into one to save some memory in Future?
    ' - First, Read data to a range by removing unneeded headers
    '     Using Resize to remove headers in the first few rows that we do not need
    Dim rngAllValidData As Range
    Set rngAllValidData = ActiveSheet.Rows(tblCfg.firstDataRowNo).Resize(nrRowsAllValidData)
    Set rngAllValidData = rngAllValidData.Resize(nrRowsAllValidData, lastColNo)
    
    ' - Second, Read data to an array to make later process faster
    Dim arrAllValidData As Variant: arrAllValidData = rngAllValidData.Value
    
    'Find all unique employee IDs. allEmpIDs is a 1-D array
    Dim allEmpIDs() As String
    Call getAllUniqueStrsByColInArray(tblCfg.titleRowNo, tblCfg.employeeIdColNo, arrAllValidData, nrRowsAllValidData, allEmpIDs)
    
    Dim nrOfRealRowsForEmp As Integer
    Dim oneEmpRecords As Variant
    ReDim oneEmpRecords(MAX_ROWS_FOR_ONE_EMPLOYEE, ActiveSheet.UsedRange.Columns.Count)
    
    Dim dictDailyHoursUpperLimitForCalcGrp As Object
    Set dictDailyHoursUpperLimitForCalcGrp = CreateObject("Scripting.Dictionary")
    Call initCalcGrpDailyHoursUpperLimit(dictDailyHoursUpperLimitForCalcGrp)
    Dim i As Integer

    'handle each employee
    For i = 1 To UBound(allEmpIDs)
        'Find all rows of this employee
        nrOfRealRowsForEmp = 0
        Call getAllRowsHasSameOneColValInArray(tblCfg.titleRowNo, tblCfg.employeeIdColNo, allEmpIDs(i), _
                                            arrAllValidData, nrRowsAllValidData, _
                                            oneEmpRecords, nrOfRealRowsForEmp)
        'Error Handling
        If nrOfRealRowsForEmp = 0 Then
            MsgBox "Employee Records for Employee ID: " & allEmpIDs(i) & "Is Empty"
            Exit For
        End If
        
        'Check if this employee's daily hours exceed upper limit
        Call chkDailyHoursUpperLimitForOneEmp(tblCfg, allEmpIDs(i), oneEmpRecords, nrOfRealRowsForEmp, _
                                                arrAllValidData, dictDailyHoursUpperLimitForCalcGrp)
    Next i
    
    MsgBox ">>>>     DailyHours     <<<<" & vbCrLf & vbCrLf & vbCrLf & _
            "Done Normally in  " & Format(Timer - StartTime, "0.00") & "  seconds!"
    
    'Turn On screen updates after running to the end
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub

'Check total regular hours in the whole pay
Sub TotalREG()
    'Turn off some global switches in order to run faster
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    'Clear background color of the table to 'No Fill' first
    ActiveSheet.Cells.Interior.ColorIndex = 0
    ActiveSheet.Cells.Font.ColorIndex = 0
    
    'Measure time performance
    Dim StartTime As Double
    StartTime = Timer
    
    Dim tblCfg As OverTimeTableCfgType
    Dim lastRowNo As Long: lastRowNo = ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row
    Dim lastColNo As Long: lastColNo = ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Column

    'Find the correct column we care
    tblCfg.firstDataRowNo = -1      'Initilize it to an invalid value first
    tblCfg.errMsgColNo = -1
    Call getTableCfg(lastRowNo, lastColNo, tblCfg)
    
    'Error handling to prevent operating on a wrong Table
    If Not chkTbleCfg(tblCfg) Then
        MsgBox "Unable to process this table by this Command Because Critical Columns is missing" & vbCrLf & vbCrLf & _
                "Column  'Calc Group Name', 'Employee ID', 'Hour Type', 'Work Date', 'Time Code' and 'Hours' are mandatory" & vbCrLf & vbCrLf & _
                "Please Check if you Opened a correct WorkSeet!"
        Exit Sub
    End If
    
    'Clear or Create a column to save result
    If (tblCfg.errMsgColNo < 0) Then
        tblCfg.errMsgColNo = lastColNo + 1
        lastColNo = lastColNo + 1
    End If
    ActiveSheet.Columns(tblCfg.errMsgColNo).ClearContents
    ActiveSheet.Cells(tblCfg.titleRowNo, tblCfg.errMsgColNo) = DEFAULT_ERR_MSG_COL_NAME

    Dim nrRowsAllValidData As Integer: nrRowsAllValidData = lastRowNo - tblCfg.firstDataRowNo + 1
    
    'Read all valid data to memory. @TODO: we may combine the following two steps into one to save some memory in Future?
    ' - First, Read data to a range by removing unneeded headers
    '     Using Resize to remove headers in the first few rows that we do not need
    Dim rngAllValidData As Range
    Set rngAllValidData = ActiveSheet.Rows(tblCfg.firstDataRowNo).Resize(nrRowsAllValidData)
    Set rngAllValidData = rngAllValidData.Resize(nrRowsAllValidData, lastColNo)
    
    ' - Second, Read data to an array to make later process faster
    Dim arrAllValidData As Variant: arrAllValidData = rngAllValidData.Value
    
    'Find all unique employee IDs. allEmpIDs is a 1-D array
    Dim allEmpIDs() As String
    Call getAllUniqueStrsByColInArray(tblCfg.titleRowNo, tblCfg.employeeIdColNo, arrAllValidData, nrRowsAllValidData, allEmpIDs)
    
    Dim nrOfRealRowsForEmp As Integer
    Dim oneEmpRecords As Variant
    ReDim oneEmpRecords(MAX_ROWS_FOR_ONE_EMPLOYEE, ActiveSheet.UsedRange.Columns.Count)
    
    Dim dictTotalRegHoursUpperLimitForCalcGrp As Object
    Set dictTotalRegHoursUpperLimitForCalcGrp = CreateObject("Scripting.Dictionary")
    Call initCalcGrpTotalRegHoursLowerLimit(dictTotalRegHoursUpperLimitForCalcGrp)

    Dim errMsg As String
    Dim strCalcGrp As String
    Dim strHourType As String
    Dim regHourInWholePay As Single
    Dim otHourInWholePay As Single
    Dim totalHourInWholePay As Single
    Dim totalRegHourLowLimit As String
    Dim i As Integer
    Dim j As Integer
    
    'handle each employee
    For i = 1 To UBound(allEmpIDs)
        'Find all rows of this employee
        nrOfRealRowsForEmp = 0
        Call getAllRowsHasSameOneColValInArray(tblCfg.titleRowNo, tblCfg.employeeIdColNo, allEmpIDs(i), _
                                            arrAllValidData, nrRowsAllValidData, _
                                            oneEmpRecords, nrOfRealRowsForEmp)
        'Error Handling
        If nrOfRealRowsForEmp = 0 Then
            MsgBox "Employee Records for Employee ID: " & allEmpIDs(i) & "Is Empty"
            Exit For
        End If
        
        strCalcGrp = oneEmpRecords(1, tblCfg.calcGrpColNo)
        totalRegHourLowLimit = dictTotalRegHoursUpperLimitForCalcGrp(strCalcGrp)
                
        'Do not check these types
        If strCalcGrp <> "CAW 6 WK ROTATION" And strCalcGrp <> "CAW 4 & 4" And strCalcGrp <> "CAW CASUAL" Then
            Call calcEmpHoursInWholePay(tblCfg, oneEmpRecords, nrOfRealRowsForEmp, regHourInWholePay, otHourInWholePay)
            'For these two types, we consider overtime hours to filter more results
            If strCalcGrp = "CAW PT" Or strCalcGrp = "CAW FT TEMP" Then
                totalHourInWholePay = regHourInWholePay + otHourInWholePay
                errMsg = "Total (Reg+OT) = "
            Else
                totalHourInWholePay = regHourInWholePay
                errMsg = "Total Reg = "
            End If
            If (totalHourInWholePay < totalRegHourLowLimit) Then
                errMsg = errMsg & totalHourInWholePay & ", Threshold = " & totalRegHourLowLimit
                'Color the first Non-Prem row of this employee because this is a whole pay check
                For j = 1 To nrOfRealRowsForEmp
                    strHourType = oneEmpRecords(j, tblCfg.hourTypeColNo)
                    If strHourType <> "REG_PREM" Then
                        Exit For
                    End If
                Next j
                Call colorEmpOneRow(tblCfg, allEmpIDs(i), oneEmpRecords, nrOfRealRowsForEmp, 1, arrAllValidData, errMsg)
            End If
        End If
    Next i
    
    MsgBox ">>>>     TotalREG     <<<<" & vbCrLf & vbCrLf & vbCrLf & _
            "Done Normally in  " & Format(Timer - StartTime, "0.00") & "  seconds!"
    
    'Turn On screen updates after running to the end
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub
