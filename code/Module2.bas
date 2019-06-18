Attribute VB_Name = "Module2"
' this is the VB6/VBA equivalent of a struct data, no methods.
' column No. of each field in pay increase table
Private Type SalaryIncreaseTable
    titleRowNo As Integer
    firstDataRowNo As Integer
    calcPayIncDateColNo As Integer
    
    lastHireDateColNo As Integer
    levelColNo As Integer
    employeeTypeColNo As Integer
End Type

'Define the calculation start point for pay date in yyyy-mm-dd format
Private Const ERR_HIRE_DATE_LATER_THAN_PAY_DAY = "Last Hire Date is Later Than Pay Day you specified "
Private Const DEFAULT_PAY_INCREASE_COL_NAME = "RATE INCREASE DATE"
Public Const DEFAULT_FIRST_PAY_DAY_IN_2019 = "2019-01-04"
Public Const DEFAULT_PAY_DAY_START_OFFSET = -27
Public Const EMP_LEVEL_TP_1 = 4     'Employee level turning point: salary increase has different rules beyond this level
Public Const EMP_LEVEL_TP_TOP = 10   'Employee level pay increase top: salary won't increase over this level
Public Const EMP_PAY_INC_PERIOD_1 = 6  'Employee pay increase period 1: every 6 months
Public Const EMP_PAY_INC_PERIOD_2 = 12 'Employee pay increase period 2: every 12 months
Public Const DEFAULT_PAY_PERIOD_IN_DAYS = 14    'Employee pay period: every 14 days (biweekly)
Public Const DEFAULT_COLOR_FOR_SALARY_INC = vbYellow
Public Const EMP_TYPE_CASUAL As String = "Casual"

'Read table header
Private Function getTableCfg(lastRowNo As Long, lastColNo As Long, ByRef tblCfg As SalaryIncreaseTable)
    Dim i, j As Long
    Dim found As Boolean: found = False
    
    'Find the correct column we care
    For i = 1 To lastRowNo
        For j = 1 To lastColNo
            With tblCfg
                If ActiveSheet.Cells(i, j).Value = "Last Hire Date" Then
                    .firstDataRowNo = i + 1
                    .titleRowNo = i
                    .lastHireDateColNo = j
                    found = True
                ElseIf ActiveSheet.Cells(i, j).Value = "Step Number" Then
                    .levelColNo = j
                ElseIf ActiveSheet.Cells(i, j).Value = "Employment Type" Then
                    .employeeTypeColNo = j
                ElseIf ActiveSheet.Cells(i, j).Value = DEFAULT_PAY_INCREASE_COL_NAME Then
                    .calcPayIncDateColNo = j
                End If
            End With
        Next j
        If found = True Then Exit For
    Next i
End Function

'Helper function: check table Colums has the fields we need
Private Function chkTbleCfg(ByRef tblCfg As SalaryIncreaseTable) As Boolean
    Dim result As Boolean: result = True
    
    With tblCfg
        If ((.firstDataRowNo <= 0) Or (.titleRowNo <= 0) Or _
            (.lastHireDateColNo <= 0) Or (.levelColNo <= 0) Or _
            (.employeeTypeColNo <= 0)) Then
            result = False
        End If
    End With
    
    chkTbleCfg = result
End Function

Private Function isSpecialDateToIncreaseSalary(div12Zero As Boolean, ByRef inDate As Date, tgtYear As Integer) As Boolean
    Dim result As Boolean: result = False
    Dim tempDay, tempMonth As Integer
    'Calculate if a year is a leap year
    Dim isLeapYear As Boolean: isLeapYear = (Month(DateSerial(tgtYear, 2, 29)) = 2)
    
    tempMonth = Month(inDate)
    tempDay = Day(inDate)
    
    'If Increase by 12 Month, there is only one special day which is Feb. 29. It depends on if current year is Leap Year as well
    If div12Zero = True And isLeapYear = False Then
        If (tempMonth = 2 And tempDay = 29) Then
            result = True
        End If
    'If Increase by 6 Month, there are several special days
    ElseIf ((tempMonth = 3 And tempDay = 31) Or _
        (tempMonth = 5 And tempDay = 31) Or _
        (tempMonth = 8 And tempDay = 29 And isLeapYear = False) Or _
        (tempMonth = 8 And tempDay >= 30) Or _
        (tempMonth = 10 And tempDay = 31) Or _
        (tempMonth = 12 And tempDay = 31)) Then
        result = True
    End If
    
    isSpecialDateToIncreaseSalary = result
End Function

'Calculate pay increase date: If it's not in this pay period, return Empty
Private Function calcPayIncreaseDate(ByRef lastHireDate As Date, ByRef payPeriodStartDate As Date, empLevel As Integer, ByRef payIncDate As Date)
    payIncDate = 0
    Dim payPeriodEndDate As Date: payPeriodEndDate = 0
    
    Dim monthOffSetStart As Long
    Dim monthOffSetEnd As Long
    Dim yearOffSet As Long

    payPeriodEndDate = DateAdd("d", DEFAULT_PAY_PERIOD_IN_DAYS - 1, payPeriodStartDate)

    monthOffSetStart = DateDiff("m", lastHireDate, payPeriodStartDate)
    monthOffSetEnd = DateDiff("m", lastHireDate, payPeriodEndDate)
    yearOffSet = DateDiff("yyyy", lastHireDate, payPeriodEndDate)

    Dim div12Zero As Boolean: div12Zero = False
    Dim tempDate As Date: tempDate = 0
    
    'Increase salary every 6 months for level 4 and below.
    'Month offset must be greater than 0 because we don't increase salary for new hire in this period.
    If (empLevel <= EMP_LEVEL_TP_1) And (monthOffSetStart > 0) Then
        If (monthOffSetStart Mod EMP_PAY_INC_PERIOD_1) = 0 Then
            tempDate = WorksheetFunction.EDate(lastHireDate, monthOffSetStart)
            If monthOffSetStart Mod EMP_PAY_INC_PERIOD_2 = 0 Then div12Zero = True
        ElseIf (monthOffSetEnd Mod EMP_PAY_INC_PERIOD_1) = 0 Then
            tempDate = WorksheetFunction.EDate(lastHireDate, monthOffSetEnd)
            If monthOffSetEnd Mod EMP_PAY_INC_PERIOD_2 = 0 Then div12Zero = True
        End If
        
        If Not tempDate = 0 Then
            'For these special days, extra 1 is needed
            If isSpecialDateToIncreaseSalary(div12Zero, lastHireDate, Year(tempDate)) Then
                tempDate = DateAdd("d", 1, tempDate)
            End If
            If tempDate >= payPeriodStartDate And tempDate <= payPeriodEndDate Then
                payIncDate = tempDate
            End If
        End If
    'Increase salary every 12 months for level 5~10.
    'Year offset must be greater than 0 for the same reason as above
    ElseIf (empLevel <= EMP_LEVEL_TP_TOP) Then
        If (yearOffSet > 0) Then
            tempDate = WorksheetFunction.EDate(lastHireDate, yearOffSet * EMP_PAY_INC_PERIOD_2)
            'For these special days, extra 1 is needed
            If isSpecialDateToIncreaseSalary(True, lastHireDate, Year(tempDate)) Then
                tempDate = DateAdd("d", 1, tempDate)
            End If
            If tempDate >= payPeriodStartDate And tempDate <= payPeriodEndDate Then
                payIncDate = tempDate
            End If
        End If
    End If
End Function

'Color a target row and insert 1 column in the end
Function handlePayIncDateResult(titleRowNo As Integer, payRowNo As Integer, _
                                ByRef payIncreaseDate As Date, payIncResultCol As Integer, _
                                ByRef thisPayStartDate As Date)
    If Not payIncreaseDate = 0 Then
        ActiveSheet.UsedRange.Rows(titleRowNo + payRowNo).Interior.Color = DEFAULT_COLOR_FOR_SALARY_INC
        ActiveSheet.Cells(payRowNo + titleRowNo, payIncResultCol) = payIncreaseDate
    End If
End Function


'Input: Pay Day, default is May 24 for today.
Sub IncreaseSalary()
    'Turn off some global switches in order to run faster
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    'Clear background color of the table to 'No Fill' first
    ActiveSheet.Cells.Interior.ColorIndex = 0
    'Get last Row/Column No. First
    Dim lastColNo As Long: lastColNo = ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Column
    Dim lastRowNo As Long: lastRowNo = ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row
    
    Dim firstPayDay As Date: firstPayDay = DateValue(DEFAULT_FIRST_PAY_DAY_IN_2019)
    Dim thisPayStartDate As Date
    Dim offset As Integer

    'Guess the most possible pay day the user may input as default.
    'User will run this command 2 Pay Period later from now
    offset = DateDiff("d", firstPayDay, Now())
    offset = offset - (offset Mod DEFAULT_PAY_PERIOD_IN_DAYS) + 2 * DEFAULT_PAY_PERIOD_IN_DAYS
    firstPayDay = DateAdd("d", offset, firstPayDay)

    'Pay Day is the input from User
    Dim strInputDate As String
    strInputDate = InputBox("Input Pay Day in format YYYY/MM/DD" & vbCrLf & vbCrLf & "Note: Pay period starts from the Pay Day you input subtracts 27", _
                            "User Input", Format(firstPayDay, "yyyy-mm-dd"))

    If strInputDate = vbNullString Then
        MsgBox "Cancel button pressed and Will Do Nothing!"
        Exit Sub
    End If
    
    'Update pay start date according to user's input
    thisPayStartDate = DateAdd("d", DEFAULT_PAY_DAY_START_OFFSET, CDate(strInputDate))

    Dim StartTime As Double
    StartTime = Timer
    
    Dim tblCfg As SalaryIncreaseTable
    'Find the correct column we care @todo: assume tile must exist in first 10 lines
    tblCfg.firstDataRowNo = -1      'Initilize it to an invalid value first
    tblCfg.calcPayIncDateColNo = -1      'Initilize it to an invalid value first
    Call getTableCfg(lastRowNo, lastColNo, tblCfg)

    'Error handling to prevent operating on a wrong Table
    If Not chkTbleCfg(tblCfg) Then
        MsgBox "Unable to process this table by this Command" & vbCrLf & vbCrLf & _
                "Column  'Last Hire Date', 'Step Number' and 'Employment Type' are mandatory" & vbCrLf & vbCrLf & _
                "Please Check if you Opened a correct Table WorkSeet!"
        Exit Sub
    End If

    'Clear or Create a column to save result
    If (tblCfg.calcPayIncDateColNo < 0) Then
        tblCfg.calcPayIncDateColNo = ActiveSheet.UsedRange.Columns.Count + 1
    End If
    ActiveSheet.Columns(tblCfg.calcPayIncDateColNo).ClearContents
    ActiveSheet.Cells(tblCfg.titleRowNo, tblCfg.calcPayIncDateColNo) = DEFAULT_PAY_INCREASE_COL_NAME

    Dim nrRowsAllValidData As Integer: nrRowsAllValidData = lastRowNo - tblCfg.firstDataRowNo + 1
    
    'Read all valid data to memory. @TODO: we may combine the following two steps into one to save some memory in Future?
    ' - First, Read data to a range by removing unneeded headers
    '     Using Resize to remove headers in the first few rows that we do not need
    Dim rngAllValidData As Range
    Set rngAllValidData = ActiveSheet.Rows(tblCfg.firstDataRowNo).Resize(nrRowsAllValidData)
    Set rngAllValidData = rngAllValidData.Resize(nrRowsAllValidData, lastColNo)
    
    ' - Second, Read data to an array to make later process faster
    Dim arrAllValidData As Variant: arrAllValidData = rngAllValidData.Value

    Dim i As Integer
    Dim payIncreaseDate As Date: payIncreaseDate = 0
    Dim lastHireDate As Date: lastHireDate = 0
    Dim empLevel As Integer
    Dim empType As String
    
    For i = 1 To UBound(arrAllValidData)
        lastHireDate = arrAllValidData(i, tblCfg.lastHireDateColNo)
        empLevel = arrAllValidData(i, tblCfg.levelColNo)
        empType = arrAllValidData(i, tblCfg.employeeTypeColNo)
        'Only calculate Non-Casual Employee's Pay Increasing Date. Casual will be calculated by User Manually
        If Not empType = EMP_TYPE_CASUAL Then
            Call calcPayIncreaseDate(lastHireDate, thisPayStartDate, empLevel, payIncreaseDate)
            
            Call handlePayIncDateResult(tblCfg.titleRowNo, i, payIncreaseDate, tblCfg.calcPayIncDateColNo, thisPayStartDate)
        End If
    Next i
    
    MsgBox ">>>>     IncreaseSalary     <<<<" & vbCrLf & vbCrLf & vbCrLf & _
            "Done Normally in " & Format(Timer - StartTime, "0.00") & " seconds!"
    
    'Turn On screen updates after running to the end
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub
