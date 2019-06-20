Attribute VB_Name = "Module4"
'PTO Check

'Mandatory columns in PTO Balance table. e.g., ORG
Private Const PTO_BALANCE_TBL_MAN_COL_START = 1
Private Const PTO_BALANCE_TBL_MAN_COL_END = 4

'Mandatory columns in employee current hours table. e.g., BWFT_Current_Hours_xxxx
Private Const EMP_CURR_HOURS_TBL_MAN_COL_START = 1
Private Const EMP_CURR_HOURS_TBL_MAN_COL_END = 4

'Max number of current hour table supported once
Private Const MAX_NOF_CURR_HOURS_TABLE = 2

'Max number of code strings for each PTO BALANCE TYPE (BANKH/FAMILY/SICK/VACAT)
'in Current Hours Table
Private Const MAX_NOF_CODES_IN_CURR_HOURS = 2

'Only look for in the first few rows for each table
Private Const MAX_NO_ROWS_FOR_TABLE_TITLE = 10

Private Const DEFAULT_BG_COLOR_FOR_MATCHED_CELL = vbGreen
Private Const DEFAULT_FONT_COLOR_FOR_MATCHED_CELL = vbRed

Private Const DEFAULT_ERR_MSG_COL_NAME = "ERROR MESSAGE"

Private Type PtoBalanceTbl  'e.g., ORG table
    titleRowNo As Integer
    firstDataRowNo As Integer
    lastRowNo As Long
    lastColNo As Long
    
    unionCodeColNo As Integer
    empNumColNo As Integer
    ptoPlanCodeColNo As Integer
    availBalColNo As Integer
End Type

Private Type EmpCurrHoursTbl    'e.g., BWFT_Current_Hours_xxxx
    titleRowNo As Integer
    firstDataRowNo As Integer
    lastRowNo As Long
    lastColNo As Long
    
    empNumColNo As Integer
    codeColNo As Integer
    hoursColNo As Integer
    calcGrpColNo As Integer
    errMsgColNo As Integer
End Type

Private Type PtoAvailBalanceType
    familyH As Single
    bankH As Single
    sickH As Single
    vacatH As Single
End Type

Private Type CurrHoursUsedType
    famH As Single
    otuH As Single
    sckH As Single
    vacH As Single
End Type

'Find the coloumn No. of each field in the PTO Balance table
Function getPtoBalanceTableCfg(ByRef ws As Worksheet, _
                               lastRowNo As Long, lastColNo As Long, _
                               ByRef ptoBalanceTblCfg As PtoBalanceTbl) As Boolean
    Dim k As Integer
    Dim found(PTO_BALANCE_TBL_MAN_COL_START To PTO_BALANCE_TBL_MAN_COL_END) As Boolean
    For k = PTO_BALANCE_TBL_MAN_COL_START To PTO_BALANCE_TBL_MAN_COL_END
        found(k) = False
    Next k
    
    With ptoBalanceTblCfg
        .titleRowNo = 0
        .firstDataRowNo = 0
        .lastRowNo = 0
        .lastColNo = 0
        .unionCodeColNo = 0
        .empNumColNo = 0
        .ptoPlanCodeColNo = 0
        .availBalColNo = 0
    End With
    
    Dim result As Boolean
    Dim i As Long, j As Long
    'Find the correct column we care
    For i = 1 To MAX_NO_ROWS_FOR_TABLE_TITLE
        result = True
        For j = 1 To lastColNo
            With ptoBalanceTblCfg
                If ws.Cells(i, j).Value = "Local Union Code" Then
                    .firstDataRowNo = i + 1
                    .titleRowNo = i
                    .unionCodeColNo = j
                    found(1) = True
                ElseIf ws.Cells(i, j).Value = "Employee Number" Then
                    .empNumColNo = j
                    found(2) = True
                ElseIf ws.Cells(i, j).Value = "PTO Plan Code" Then
                    .ptoPlanCodeColNo = j
                    found(3) = True
                ElseIf ws.Cells(i, j).Value = "Available Balance" Then
                    .availBalColNo = j
                    found(4) = True
                End If
            End With
        Next j
        
        For k = PTO_BALANCE_TBL_MAN_COL_START To PTO_BALANCE_TBL_MAN_COL_END
            If found(k) = False Then
                result = False
                Exit For
            End If
        Next k
        
        If result = True Then
            ptoBalanceTblCfg.lastRowNo = lastRowNo
            ptoBalanceTblCfg.lastColNo = lastColNo
            Exit For
        End If
    Next i
    
    getPtoBalanceTableCfg = result
End Function

'Find the coloumn No. of each field in the Employee Current Hours table
Function getEmpCurrHoursTableCfg(ByRef ws As Worksheet, _
                                lastRowNo As Long, lastColNo As Long, _
                                ByRef empCurrHoursTblCfg As EmpCurrHoursTbl) As Boolean
    Dim k As Integer
    Dim found(EMP_CURR_HOURS_TBL_MAN_COL_START To EMP_CURR_HOURS_TBL_MAN_COL_END) As Boolean
    For k = EMP_CURR_HOURS_TBL_MAN_COL_START To EMP_CURR_HOURS_TBL_MAN_COL_END
        found(k) = False
    Next k
    
    With empCurrHoursTblCfg
        .titleRowNo = 0
        .firstDataRowNo = 0
        .lastRowNo = 0
        .lastColNo = 0
        
        .empNumColNo = 0
        .codeColNo = 0
        .hoursColNo = 0
        .calcGrpColNo = 0
        .errMsgColNo = 0
    End With
    
    Dim result As Boolean
    Dim i As Long, j As Long
    'Find the correct column we care
    For i = 1 To MAX_NO_ROWS_FOR_TABLE_TITLE
        result = True
        For j = 1 To lastColNo
            With empCurrHoursTblCfg
                If ws.Cells(i, j).Value = "Emp Num" Then
                    .firstDataRowNo = i + 1
                    .titleRowNo = i
                    .empNumColNo = j
                    found(1) = True
                ElseIf ws.Cells(i, j).Value = "Code" Then
                    .codeColNo = j
                    found(2) = True
                ElseIf ws.Cells(i, j).Value = "Hours" Then
                    .hoursColNo = j
                    found(3) = True
                ElseIf ws.Cells(i, j).Value = "Calc Group" Then
                    .calcGrpColNo = j
                    found(4) = True
                ElseIf ws.Cells(i, j).Value = DEFAULT_ERR_MSG_COL_NAME Then
                    .errMsgColNo = j
                End If
            End With
        Next j

        For k = EMP_CURR_HOURS_TBL_MAN_COL_START To EMP_CURR_HOURS_TBL_MAN_COL_END
            If found(k) = False Then
                result = False
            End If
        Next k
        
        If result = True Then
            empCurrHoursTblCfg.lastRowNo = lastRowNo
            empCurrHoursTblCfg.lastColNo = lastColNo

            If empCurrHoursTblCfg.errMsgColNo <= 0 Then
                empCurrHoursTblCfg.errMsgColNo = lastColNo + 1
                empCurrHoursTblCfg.lastColNo = lastColNo + 1
            End If
            
            Exit For
        End If
    Next i
    
    getEmpCurrHoursTableCfg = result
End Function

'Find all the talbes we need to handle:
'  PTOBalance is mandatory, any one of BWFT/BWPT or both
Function findTargetTables(ByRef ptoBalanceWorkSheet As Worksheet, _
                        ByRef ptoBalanceTblCfg As PtoBalanceTbl, _
                        ByRef currHoursWorkSheet() As Worksheet, _
                        ByRef currHoursTblCfg() As EmpCurrHoursTbl, _
                        ByRef noOfCurrHoursWorkSheets As Integer) As Boolean
    Dim book As Workbook
    Dim sheet As Worksheet
    Dim lastRowNo As Long
    Dim lastColNo As Long
    
    Dim k As Integer
    Dim result As Boolean
    Dim found(0 To MAX_NOF_CURR_HOURS_TABLE) As Boolean
    'found(0) is for PtoBalance Table;
    'found(1 To MAX_NOF_CURR_HOURS_TABLE) is current balane Table
    For k = 0 To MAX_NOF_CURR_HOURS_TABLE
        found(k) = False
    Next k

    noOfCurrHoursWorkSheets = 0

    For Each book In Workbooks
        For Each sheet In book.Worksheets
            lastRowNo = sheet.Cells.SpecialCells(xlCellTypeLastCell).Row
            lastColNo = sheet.Cells.SpecialCells(xlCellTypeLastCell).Column
            If Not found(0) Then
                result = getPtoBalanceTableCfg(sheet, lastRowNo, lastColNo, ptoBalanceTblCfg)
                If result Then
                    found(0) = True
                    Set ptoBalanceWorkSheet = sheet
                End If
            End If
            For k = 1 To MAX_NOF_CURR_HOURS_TABLE
                If found(k) = False Then
                    result = getEmpCurrHoursTableCfg(sheet, lastRowNo, lastColNo, currHoursTblCfg(k))
                    If result Then
                        found(k) = True
                        Set currHoursWorkSheet(k) = sheet
                        noOfCurrHoursWorkSheets = noOfCurrHoursWorkSheets + 1
                        'same sheet can only be pointed by one element in currHoursTblCfg
                        Exit For
                    End If
                End If
            Next k
        Next sheet
    Next book
    
    'If PtoBalance Table is found and at least 1 current hours table is found
    If found(0) = True And noOfCurrHoursWorkSheets >= 1 Then
        findTargetTables = True
    Else
        findTargetTables = False
    End If
End Function

Private Function getPtoTblAvailBalance(ByRef tgtStrEmpNum As String, _
                                        ByRef arrCurrHoursData As Variant, _
                                        ByRef hoursWS As Worksheet, _
                                        ByRef currHoursTblCfg As EmpCurrHoursTbl, _
                                        ByRef ptoBalanceTblCfg As PtoBalanceTbl, _
                                        ByRef oneEmpPtoBalRecords As Variant, _
                                        nrOfRealPtoBalRowsForEmp As Integer, _
                                        ByRef ptoAvailBalance As PtoAvailBalanceType) As Boolean
    Dim i As Integer
    Dim strPlanCode As String
    Dim strUnionCode As String
    Dim errMsg As String
    Dim balance As Single

    With ptoAvailBalance
        .familyH = 0
        .bankH = 0
        .sickH = 0
        .vacatH = 0
    End With
    
    strUnionCode = oneEmpPtoBalRecords(1, ptoBalanceTblCfg.unionCodeColNo)
    If strUnionCode <> "CAW" Then
        errMsg = "Unsupported Union"
        Call colorOneEmpInCurrHoursTable(tgtStrEmpNum, arrCurrHoursData, _
                                        hoursWS, currHoursTblCfg, errMsg)
        getPtoTblAvailBalance = False
        Exit Function
    End If

    For i = 1 To nrOfRealPtoBalRowsForEmp
        strPlanCode = oneEmpPtoBalRecords(i, ptoBalanceTblCfg.ptoPlanCodeColNo)
        balance = oneEmpPtoBalRecords(i, ptoBalanceTblCfg.availBalColNo)
        If strPlanCode = "BANKH" Then
            ptoAvailBalance.bankH = ptoAvailBalance.bankH + balance
        ElseIf strPlanCode = "FAMILY" Then
            ptoAvailBalance.familyH = ptoAvailBalance.familyH + balance
        ElseIf strPlanCode = "SICK" Then
            ptoAvailBalance.sickH = ptoAvailBalance.sickH + balance
        ElseIf strPlanCode = "VACAT" Then
            ptoAvailBalance.vacatH = ptoAvailBalance.vacatH + balance
        End If
    Next i
    
    getPtoTblAvailBalance = True
End Function

Private Function getCurrHoursTableUsed(ByRef currHoursTblCfg As EmpCurrHoursTbl, _
                                        ByRef oneEmpHoursRecords As Variant, _
                                        nrOfRealCurrHoursRowsForEmp As Integer, _
                                        ByRef currHoursUsed As CurrHoursUsedType)
    Dim i As Integer
    Dim strCode As String
    Dim usedHours As Single

    With currHoursUsed
        .famH = 0
        .otuH = 0
        .sckH = 0
        .vacH = 0
    End With

    For i = 1 To nrOfRealCurrHoursRowsForEmp
        strCode = oneEmpHoursRecords(i, currHoursTblCfg.codeColNo)
        usedHours = oneEmpHoursRecords(i, currHoursTblCfg.hoursColNo)
        If strCode = "FAMAV" Or strCode = "FAMLY" Then
            currHoursUsed.famH = currHoursUsed.famH + usedHours
        ElseIf strCode = "OTU" Or strCode = "OTUAV" Then
            currHoursUsed.otuH = currHoursUsed.otuH + usedHours
        ElseIf strCode = "SCKAV" Or strCode = "SICK" Then
            currHoursUsed.sckH = currHoursUsed.sckH + usedHours
        ElseIf strCode = "VACAV" Or strCode = "VACH" Then
            currHoursUsed.vacH = currHoursUsed.vacH + usedHours
        End If
    Next i
End Function

Private Function colorOneRowInCurrHoursTable(ByRef tgtStrEmpNum As String, _
                                            ByRef arrCurrHoursData As Variant, _
                                            ByRef hoursWS As Worksheet, _
                                            ByRef currHoursTblCfg As EmpCurrHoursTbl, _
                                            ByRef tgtStrCode() As String, _
                                            ByRef usedH As Single, _
                                            ByRef balanceH As Single)
    Dim i As Integer
    Dim strEmpNum As String
    Dim strCode As String
    Dim hours As Single
    Dim errMsg As String
    Dim found As Boolean: found = False

    For i = 1 To UBound(arrCurrHoursData)
        strEmpNum = arrCurrHoursData(i, currHoursTblCfg.empNumColNo)
        strCode = arrCurrHoursData(i, currHoursTblCfg.codeColNo)
        hours = arrCurrHoursData(i, currHoursTblCfg.hoursColNo)
        If strEmpNum = tgtStrEmpNum And hours > 0 Then
            If strCode = tgtStrCode(1) Or strCode = tgtStrCode(2) Then
                found = True
                errMsg = strCode & " Overused = " & (usedH - balanceH) & ", Used/Balance = (" & usedH & "/" & balanceH & ")"
                hoursWS.UsedRange.Rows(i + currHoursTblCfg.titleRowNo).Font.Color = DEFAULT_FONT_COLOR_FOR_ABNORMAL_ROW
                hoursWS.UsedRange.Rows(i + currHoursTblCfg.titleRowNo).Interior.Color = DEFAULT_BG_COLOR_FOR_ABNORMAL_ROW
                hoursWS.Cells(i + currHoursTblCfg.titleRowNo, currHoursTblCfg.errMsgColNo) = errMsg
            End If
        End If
    Next i
    
    'when we have not find the record, maybe because there is no record in current hour table. We color the first row
    If found = False Then
        For i = 1 To UBound(arrCurrHoursData)
            strEmpNum = arrCurrHoursData(i, currHoursTblCfg.empNumColNo)
            If strEmpNum = tgtStrEmpNum Then
                errMsg = tgtStrCode(1) & " Overused Before = " & (usedH - balanceH) & ", Used/Balance = (" & usedH & "/" & balanceH & ")"
                hoursWS.UsedRange.Rows(i + currHoursTblCfg.titleRowNo).Font.Color = DEFAULT_FONT_COLOR_FOR_ABNORMAL_ROW
                hoursWS.UsedRange.Rows(i + currHoursTblCfg.titleRowNo).Interior.Color = DEFAULT_BG_COLOR_FOR_ABNORMAL_ROW
                hoursWS.Cells(i + currHoursTblCfg.titleRowNo, currHoursTblCfg.errMsgColNo) = errMsg
                Exit For
            End If
        Next i
    End If
End Function

Private Function colorOneEmpInCurrHoursTable(ByRef tgtStrEmpNum As String, _
                                            ByRef arrCurrHoursData As Variant, _
                                            ByRef hoursWS As Worksheet, _
                                            ByRef currHoursTblCfg As EmpCurrHoursTbl, _
                                            ByRef errMsg As String)
    Dim i As Integer
    Dim strEmpNum As String

    For i = 1 To UBound(arrCurrHoursData)
        strEmpNum = arrCurrHoursData(i, currHoursTblCfg.empNumColNo)
        If strEmpNum = tgtStrEmpNum Then
            hoursWS.UsedRange.Rows(i + currHoursTblCfg.titleRowNo).Font.Color = DEFAULT_FONT_COLOR_FOR_ABNORMAL_ROW
            hoursWS.UsedRange.Rows(i + currHoursTblCfg.titleRowNo).Interior.Color = DEFAULT_BG_COLOR_FOR_ABNORMAL_ROW
            hoursWS.Cells(i + currHoursTblCfg.titleRowNo, currHoursTblCfg.errMsgColNo) = errMsg
        End If
    Next i
End Function

Private Function handleHoursBalanceDiff(ByRef tgtStrEmpNum As String, _
                                        ByRef arrCurrHoursData As Variant, _
                                        ByRef hoursWS As Worksheet, _
                                        ByRef currHoursTblCfg As EmpCurrHoursTbl, _
                                        ByRef ptoAvailBalance As PtoAvailBalanceType, _
                                        ByRef currHoursUsed As CurrHoursUsedType)
    Dim strCode(MAX_NOF_CODES_IN_CURR_HOURS) As String
    Dim ret As Single
    
    If (ptoAvailBalance.familyH - currHoursUsed.famH) < 0 Then
        strCode(1) = "FAMAV"
        strCode(2) = "FAMLY"
        Call colorOneRowInCurrHoursTable(tgtStrEmpNum, arrCurrHoursData, _
                                        hoursWS, currHoursTblCfg, _
                                        strCode, currHoursUsed.famH, ptoAvailBalance.familyH)
    End If
    
    If (ptoAvailBalance.bankH - currHoursUsed.otuH) < 0 Then
        strCode(1) = "OTU"
        strCode(2) = "OTUAV"
        Call colorOneRowInCurrHoursTable(tgtStrEmpNum, arrCurrHoursData, _
                                        hoursWS, currHoursTblCfg, _
                                        strCode, currHoursUsed.otuH, ptoAvailBalance.bankH)
    End If
    
    If (ptoAvailBalance.sickH - currHoursUsed.sckH) < 0 Then
        strCode(1) = "SICK"
        strCode(2) = "SCKAV"
        Call colorOneRowInCurrHoursTable(tgtStrEmpNum, arrCurrHoursData, _
                                        hoursWS, currHoursTblCfg, _
                                        strCode, currHoursUsed.sckH, ptoAvailBalance.sickH)
    End If
    
    If (ptoAvailBalance.vacatH - currHoursUsed.vacH) < 0 Then
        strCode(1) = "VACH"
        strCode(2) = "VACAV"
        Call colorOneRowInCurrHoursTable(tgtStrEmpNum, arrCurrHoursData, _
                                        hoursWS, currHoursTblCfg, _
                                        strCode, currHoursUsed.vacH, ptoAvailBalance.vacatH)
    End If
End Function

Function handleEmpCurrHoursRecordsPtoChk(ByRef tgtStrEmpNum As String, _
                                        ByRef arrCurrHoursData As Variant, _
                                        ByRef hoursWS As Worksheet, _
                                        ByRef currHoursTblCfg As EmpCurrHoursTbl, _
                                        ByRef oneEmpHoursRecords As Variant, _
                                        nrOfRealCurrHoursRowsForEmp As Integer, _
                                        ByRef ptoBalanceTblCfg As PtoBalanceTbl, _
                                        ByRef oneEmpPtoBalRecords As Variant, _
                                        nrOfRealPtoBalRowsForEmp As Integer)

    Dim ptoAvailBalance As PtoAvailBalanceType
    Dim currHoursUsed As CurrHoursUsedType
    Dim result As Boolean
    
    result = getPtoTblAvailBalance(tgtStrEmpNum, arrCurrHoursData, hoursWS, currHoursTblCfg, _
                                    ptoBalanceTblCfg, oneEmpPtoBalRecords, _
                                    nrOfRealPtoBalRowsForEmp, ptoAvailBalance)
    If result Then
        Call getCurrHoursTableUsed(currHoursTblCfg, oneEmpHoursRecords, _
                                    nrOfRealCurrHoursRowsForEmp, currHoursUsed)
    
        Call handleHoursBalanceDiff(tgtStrEmpNum, arrCurrHoursData, hoursWS, _
                                    currHoursTblCfg, ptoAvailBalance, currHoursUsed)
    End If
End Function

'Check total regular hours in the whole pay
Sub PtoCheck()
    'Turn off some global switches in order to run faster
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    'Measure time performance
    Dim StartTime As Double
    StartTime = Timer
    
    Dim ptoBalanceWorkSheet As Worksheet
    Dim ptoBalanceTblCfg As PtoBalanceTbl

    Dim currHoursWorkSheet(1 To MAX_NOF_CURR_HOURS_TABLE) As Worksheet
    Dim currHoursTblCfg(1 To MAX_NOF_CURR_HOURS_TABLE) As EmpCurrHoursTbl
    Dim noOfCurrHoursWorkSheets As Integer
    Dim result As Boolean: result = False
    
    result = findTargetTables(ptoBalanceWorkSheet, ptoBalanceTblCfg, _
                            currHoursWorkSheet, currHoursTblCfg, _
                            noOfCurrHoursWorkSheets)
    If result = False Then
        Dim errOut As String
        errOut = "Unable to handle PtoCheck Request" & vbCrLf & vbCrLf
        If IsEmpty(ptoBalanceWorkSheet) Then
            errOut = errOut & "PTO Balances Table is NOT Open" & vbCrLf
            errOut = errOut & "Column  'Local Union Code', 'Employee Number'," & _
                              "'PTO Plan Code' and 'Available Balance' are mandatory" & vbCrLf
        End If
        
        If IsEmpty(currHoursWorkSheet(1)) Then
            errOut = errOut & "At least one of BWFT/BWPT Current Hours Table has to be Open" & vbCrLf
            errOut = errOut & "Column  'Emp Num', 'Code', 'Hours' and 'Calc Group' are mandatory" & vbCrLf
        End If

        MsgBox errOut
        Exit Sub
    End If

    Dim i As Integer
    Dim k As Integer

    'Read all data in from PTO Check Table
    Dim nrValidRowsInPtoBalanceTbl As Integer
    nrValidRowsInPtoBalanceTbl = ptoBalanceTblCfg.lastRowNo - ptoBalanceTblCfg.firstDataRowNo + 1
    
    Dim rngPtoBalanceTbl As Range
    Set rngPtoBalanceTbl = ptoBalanceWorkSheet.Rows(ptoBalanceTblCfg.firstDataRowNo).Resize(nrValidRowsInPtoBalanceTbl)
    Set rngPtoBalanceTbl = rngPtoBalanceTbl.Resize(nrValidRowsInPtoBalanceTbl, ptoBalanceTblCfg.lastColNo)
    Dim arrPtoBalanceData As Variant: arrPtoBalanceData = rngPtoBalanceTbl.Value

    'Read all data in from Current Hours Table
    Dim colorOfFirstRow() As Long
    Dim nrValidRowsInCurrHoursTbl() As Integer
    Dim rngCurrHoursTbl() As Range
    Dim arrCurrHoursData() As Variant

    ReDim colorOfFirstRow(1 To noOfCurrHoursWorkSheets)
    ReDim nrValidRowsInCurrHoursTbl(1 To noOfCurrHoursWorkSheets)
    ReDim rngCurrHoursTbl(1 To noOfCurrHoursWorkSheets)
    ReDim arrCurrHoursData(1 To noOfCurrHoursWorkSheets)

    Dim allEmpIDs() As String
    Dim nrOfRealCurrHoursRowsForEmp As Integer
    Dim oneEmpHoursRecords As Variant
    
    Dim nrOfRealPtoBalRowsForEmp As Integer
    Dim oneEmpPtoBalRecords As Variant
    ReDim oneEmpPtoBalRecords(MAX_ROWS_FOR_ONE_EMPLOYEE, ptoBalanceTblCfg.lastColNo)

    For k = 1 To noOfCurrHoursWorkSheets
        nrValidRowsInCurrHoursTbl(k) = currHoursTblCfg(k).lastRowNo - currHoursTblCfg(k).firstDataRowNo + 1
        Set rngCurrHoursTbl(k) = currHoursWorkSheet(k).Rows(currHoursTblCfg(k).firstDataRowNo).Resize(nrValidRowsInCurrHoursTbl(k))
        Set rngCurrHoursTbl(k) = rngCurrHoursTbl(k).Resize(nrValidRowsInCurrHoursTbl(k), currHoursTblCfg(k).lastColNo)
        arrCurrHoursData(k) = rngCurrHoursTbl(k).Value
        
        'Recovering colors
        colorOfFirstRow(k) = currHoursWorkSheet(k).Rows(currHoursTblCfg(k).titleRowNo).Interior.Color
        currHoursWorkSheet(k).Cells.Font.ColorIndex = 0
        currHoursWorkSheet(k).Cells.Interior.ColorIndex = 0
        currHoursWorkSheet(k).Rows(currHoursTblCfg(k).titleRowNo).Interior.Color = colorOfFirstRow(k)
        
        'Recovering error message
        currHoursWorkSheet(k).Columns(currHoursTblCfg(k).errMsgColNo).ClearContents
        currHoursWorkSheet(k).Cells(currHoursTblCfg(k).titleRowNo, currHoursTblCfg(k).errMsgColNo) = DEFAULT_ERR_MSG_COL_NAME

        
        Call getAllUniqueStrsByColInArray(currHoursTblCfg(k).titleRowNo, currHoursTblCfg(k).empNumColNo, _
                                        arrCurrHoursData(k), nrValidRowsInCurrHoursTbl(k), allEmpIDs)
    
        ReDim oneEmpHoursRecords(MAX_ROWS_FOR_ONE_EMPLOYEE, currHoursTblCfg(k).lastColNo)
        'Find one employee's records in current hour table
        For i = 1 To UBound(allEmpIDs)
            'Get one employee record in Current Hours Table
            Call getAllRowsHasSameOneColValInArray(currHoursTblCfg(k).titleRowNo, currHoursTblCfg(k).empNumColNo, _
                                                allEmpIDs(i), arrCurrHoursData(k), nrValidRowsInCurrHoursTbl(k), _
                                                oneEmpHoursRecords, nrOfRealCurrHoursRowsForEmp)
            'Get one employee record in PTO Balance Table
            Call getAllRowsHasSameOneColValInArray(ptoBalanceTblCfg.titleRowNo, ptoBalanceTblCfg.empNumColNo, _
                                                allEmpIDs(i), arrPtoBalanceData, nrValidRowsInPtoBalanceTbl, _
                                                oneEmpPtoBalRecords, nrOfRealPtoBalRowsForEmp)
            
            Call handleEmpCurrHoursRecordsPtoChk(allEmpIDs(i), arrCurrHoursData(k), _
                                                currHoursWorkSheet(k), currHoursTblCfg(k), _
                                                oneEmpHoursRecords, nrOfRealCurrHoursRowsForEmp, _
                                                ptoBalanceTblCfg, oneEmpPtoBalRecords, nrOfRealPtoBalRowsForEmp)
        Next i
    Next k

    MsgBox ">>>>     PtoCheck     <<<<" & vbCrLf & vbCrLf & vbCrLf & _
            "Done Normally in  " & Format(Timer - StartTime, "0.00") & "  seconds!"
    
    'Turn On screen updates after running to the end
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub
