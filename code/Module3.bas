Attribute VB_Name = "Module3"
' this is the VB6/VBA equivalent of a struct data, no methods.
' column No. of each field in the manually maintained table called Increase Dates Table

Private Const EMP_PAY_INC_LEVEL_START = 2
Private Const EMP_PAY_INC_LEVEL_END = 11
Private Const DEFAULT_BG_COLOR_FOR_MATCHED_CELL = vbGreen
Private Const DEFAULT_FONT_COLOR_FOR_MATCHED_CELL = vbRed

Private Type IncreaseDateTbl
    titleRowNo As Integer
    firstDataRowNo As Integer
    
    empNameColNo As Integer
    'define so many level columns so that we can support column sequence change in the table
    levelColNo(EMP_PAY_INC_LEVEL_START To EMP_PAY_INC_LEVEL_END) As Integer
End Type

Private Const DEFAULT_BG_COLOR_FOR_MATCHED_CELL = vbGreen
Private Const DEFAULT_FONT_COLOR_FOR_MATCHED_CELL = vbRed

'Find the coloumn No. of each field in the table
Private Function getTableCfg(lastRowNo As Long, lastColNo As Long, ByRef tblCfg As IncreaseDateTbl)
    Dim found As Boolean: found = False
    Dim i As Long, j As Long
    'Find the correct column we care
    For i = 1 To lastRowNo
        For j = 1 To lastColNo
            With tblCfg
                If ActiveSheet.Cells(i, j).Value = "Employee Name" Then
                    .firstDataRowNo = i + 1
                    .titleRowNo = i
                    .empNameColNo = j
                    found = True
                ElseIf ActiveSheet.Cells(i, j).Value = "Level2" Then
                    .levelColNo(2) = j
                ElseIf ActiveSheet.Cells(i, j).Value = "Level3" Then
                    .levelColNo(3) = j
                ElseIf ActiveSheet.Cells(i, j).Value = "Level4" Then
                    .levelColNo(4) = j
                ElseIf ActiveSheet.Cells(i, j).Value = "Level5" Then
                    .levelColNo(5) = j
                ElseIf ActiveSheet.Cells(i, j).Value = "Level6" Then
                    .levelColNo(6) = j
                ElseIf ActiveSheet.Cells(i, j).Value = "Level7" Then
                    .levelColNo(7) = j
                ElseIf ActiveSheet.Cells(i, j).Value = "Level8" Then
                    .levelColNo(8) = j
                ElseIf ActiveSheet.Cells(i, j).Value = "Level9" Then
                    .levelColNo(9) = j
                ElseIf ActiveSheet.Cells(i, j).Value = "Level10" Then
                    .levelColNo(10) = j
                ElseIf ActiveSheet.Cells(i, j).Value = "Level11" Then
                    .levelColNo(11) = j
                End If
            End With
        Next j
        If found = True Then Exit For
    Next i
End Function

'Helper function: check table Colums has the fields we need
Private Function chkTbleCfg(ByRef tblCfg As IncreaseDateTbl) As Boolean
    Dim result As Boolean: result = True

    With tblCfg
        If  (.firstDataRowNo <= 0) Or (.titleRowNo <= 0) Or _
            (.empNameColNo <= 0) Then
            result = False
        Else
            For j = EMP_PAY_INC_LEVEL_START To EMP_PAY_INC_LEVEL_END
                If .levelColNo(j) <= 0 Then 
                    result = False
                    Exit For
                End If
            Next j
    End With

    chkTbleCfg = result
End Function

Private Function isCurrDateInPayPeriod(ByRef currDate As Date, _
                                        ByRef payPeriodStartDate As Date) As Boolean
    Dim payPeriodEndDate As Date
    payPeriodEndDate = DateAdd("d", 14 - 1, payPeriodStartDate)
    
    If currDate >= payPeriodStartDate And currDate <= payPeriodEndDate Then
        isCurrDateInPayPeriod = True
    Else
        isCurrDateInPayPeriod = False
    End If
End Function

'Check total regular hours in the whole pay
Sub PayIncForPosTransfer()
    'Turn off some global switches in order to run faster
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    'Clear background color of the table to 'No Fill' first
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
    
    Dim tblCfg As IncreaseDateTbl
    Dim lastRowNo As Long: lastRowNo = ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row
    Dim lastColNo As Long: lastColNo = ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Column

    'Find the correct column we care
    tblCfg.firstDataRowNo = -1      'Initilize it to an invalid value first
    tblCfg.empNameColNo = -1
    Call getTableCfg(lastRowNo, lastColNo, tblCfg)
    
    'Error handling to prevent operating on a wrong Table
    If Not chkTbleCfg(tblCfg) Then
        MsgBox "Unable to process this table by this Command Because Critical Columns is missing" & vbCrLf & vbCrLf & _
                "Column  'Employee Name' and 'Level2' ~ 'Level11' are mandatory" & vbCrLf & vbCrLf & _
                "Please Check if you Opened a correct WorkSeet!"
        Exit Sub
    End If

    Dim nrRowsAllValidData As Integer: nrRowsAllValidData = lastRowNo - tblCfg.firstDataRowNo + 1

    'Read all valid data to memory.
    ' - First, Read data to a range by removing unneeded headers
    '     Using Resize to remove headers in the first few rows that we do not need
    Dim rngAllValidData As Range
    Set rngAllValidData = ActiveSheet.Rows(tblCfg.firstDataRowNo).Resize(nrRowsAllValidData)
    Set rngAllValidData = rngAllValidData.Resize(nrRowsAllValidData, lastColNo)
    
    ' - Second, Read data to an array to make later process faster
    Dim arrAllValidData As Variant: arrAllValidData = rngAllValidData.Value
    
    Dim i As Integer
    Dim j As Integer
    Dim rowMatched As Boolean
    Dim empName As String
    Dim levelDate(EMP_PAY_INC_LEVEL_START To EMP_PAY_INC_LEVEL_END) As Date

    'Read each data row
    For i = 1 To UBound(arrAllValidData)
        rowMatched = False

        empName = arrAllValidData(i, empNameColNo)
        'See if we have any dates can match this pay period
        For j = EMP_PAY_INC_LEVEL_START To EMP_PAY_INC_LEVEL_END
            levelDate(j) = arrAllValidData(i, levelColNo(j))
            If isCurrDateInPayPeriod(levelDate(j), payPeriodStartDate) Then
                rowMatched = True
                ActiveSheet.Cells(i + tblCfg.titleRowNo, tblCfg.levelColNo(j)).Interior.Color = _
                                DEFAULT_BG_COLOR_FOR_MATCHED_CELL
            End If
        Next j
        'When there is any match in this row, we color the employee Name column
        If rowMatched Then
            ActiveSheet.Cells(i + tblCfg.titleRowNo, tblCfg.empNameColNo).Interior.Color = _
                                DEFAULT_BG_COLOR_FOR_MATCHED_CELL
            ActiveSheet.Cells(i + tblCfg.titleRowNo, tblCfg.empNameColNo).Font.Color = _
                                DEFAULT_FONT_COLOR_FOR_MATCHED_CELL
        End If
    Next i
    
    MsgBox ">>>>     PayIncForPosTransfer     <<<<" & vbCrLf & vbCrLf & vbCrLf & _
            "Done Normally in  " & Format(Timer - StartTime, "0.00") & "  seconds!"
    
    'Turn On screen updates after running to the end
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub
