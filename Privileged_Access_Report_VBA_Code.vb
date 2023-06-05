Privileged_Access_Report_VBA_Code.vb

------------------
Edit_CA.cls

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Edit_CA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@IgnoreModule UndeclaredVariable, UnassignedVariableUsage
Option Explicit

' Procedure purpose:  To reconnect/refresh data sources
Private Sub Connect_CA_Click()

    Dim lResult As Long
    Dim lRet As Boolean
    Dim i As Integer
    Dim ds As String
    Dim ds_concat As String
    Dim InfoBox As VbMsgBoxResult
    
    On Error GoTo ErrorHandler
    
    Const AckTime As Integer = 3
    
    Call OnStart
    
    lRet = True
    For i = 1 To 7
        ds = "DS_" & i
        lResult = Application.Run("SAPGetProperty", "IsDataSourceActive", ds)
        lRet = lRet And lResult
        If lRet = False Then
            If ds_concat <> vbNullString Then
                ds_concat = ds_concat & ", " & vbCrLf & "'" & ds & "'"
            Else
                ds_concat = "'" & ds & "'"
            End If
        End If
    Next i
    
    If ds_concat <> vbNullString Then
        MsgBox "Data Sources: " & vbCrLf & ds_concat & vbCrLf & " are inactive"
    End If
    
    If lResult = False Then
        
        ThisWorkbook.Sheets("Edit_CA").Activate
        ActiveSheet.Range("A1").Activate
        
        lResult = Application.Run("SAPLogOff", "True")
        lResult = Application.Run("SAPSetRefreshBehaviour", "Off")
        lResult = Application.Run("SAPExecuteCommand", "Refresh", "ALL")
        
        Call DataValidationList
        
        lResult = Application.Run("SAPSetRefreshBehaviour", "On")
        
    Else
        InfoBox = TimedMsgBox("You are already connected to the system" _
                            & vbCrLf & vbCrLf & "(This window will close in " & AckTime & " seconds)", _
                              "Connection Status", , AckTime)
    End If
    
    Call Alignment
    Call OnEnd
ErrorHandler:
    ' Call the global error handling procedure
    Call HandleError
End Sub

' Procedure purpose:  To save data in Planning Query
Private Sub Save_CA_Click()

    Dim EndTime As Double
    Dim StartTime As Double
    Dim ds_alias As String: ds_alias = "DS_3"
    Dim lResult As Long
    Dim wb As Workbook
    Dim InfoBox As VbMsgBoxResult
    
    On Error GoTo ErrorHandler
    
    Const AckTime As Integer = 3
    
    Set wb = ThisWorkbook
    
    Call OnStart
    
    lResult = Application.Run("SAPGetProperty", "IsDataSourceActive", ds_alias)
    If lResult = True Then
        
        StartTime = Timer
        lResult = Application.Run("SAPSetRefreshBehaviour", "Off")
        lResult = Application.Run("SAPDeleteDesignRule", ds_alias)
        lResult = Application.Run("SAPGetProperty", "IsDataSourceEditable", ds_alias)
        If lResult = True Then
            lResult = Application.Run("SAPGetProperty", "HasChangedPlanData", ds_alias)
            If lResult = True Then
                lResult = Application.Run("SAPExecuteCommand", "PlanDataSave")
                lResult = Application.Run("SAPExecuteCommand", "Restart", "ALL")
            
                Call DataValidationList
                
                EndTime = Timer
                lResult = Application.Run("SAPSetRefreshBehaviour", "On")
                wb.Sheets("Edit_CA").Range("E1").Value = Format((EndTime - StartTime) / 86400, "hh:mm:ss") & " [hh:mm:ss]"
                
                InfoBox = TimedMsgBox("Data saved in : " & Format((EndTime - StartTime) / 86400, "hh:mm:ss") & " [hh:mm:ss]" _
                                    & vbCrLf & vbCrLf & "(This window will close in " & AckTime & " seconds)", _
                                      "Save Status", , AckTime)
            Else
                InfoBox = TimedMsgBox("No data has been changed" _
                                    & vbCrLf & vbCrLf & "(This window will close in " & AckTime & " seconds)", _
                                      "Save Status", , AckTime)
            End If
            
            lResult = Application.Run("SAPSetRefreshBehaviour", "On")
        
        Else
            InfoBox = TimedMsgBox("Cannot save the data, please check if the query is in 'change mode' (Analysis ribbon)" _
                                & vbCrLf & vbCrLf & "(This window will close in " & AckTime & " seconds)", _
                                  "Connection Status", , AckTime)
            lResult = Application.Run("SAPSetRefreshBehaviour", "On")
        End If
        
    Else
        InfoBox = TimedMsgBox("You are not connected to the system" _
                            & vbCrLf & vbCrLf & "(This window will close in " & AckTime & " seconds)", _
                              "Connection Status", , AckTime)
    End If
    
    Call Alignment
    Call OnEnd
    
ErrorHandler:
    ' Call the global error handling procedure
    Call HandleError
End Sub

' Procedure purpose:  To enable floating buttons
Private Sub Worksheet_SelectionChange(ByVal target As Excel.Range)

    On Error GoTo 0
    With Cells(Windows(1).ScrollRow, Windows(1).ScrollColumn)
        Connect_CA.Top = .Top + 0
        Connect_CA.Left = .Left + 0
        Save_CA.Top = .Top + 0
        Save_CA.Left = .Left + 80
    End With
End Sub




------------------
Edit_SP.cls

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Edit_SP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@IgnoreModule UndeclaredVariable
' Procedure purpose:  To reconnect/refresh data sources
Private Sub Connect_SP_Click()

    Dim InfoBox As VbMsgBoxResult
    Dim ds As String
    Dim ds_concat As String
    Dim i As Integer
    Dim lResult As Long
    Dim lRet As Boolean
    
    On Error GoTo ErrorHandler
    
    Const AckTime As Integer = 3
    
    Call OnStart
    
    lRet = True
    For i = 1 To 7
        ds = "DS_" & i
        lResult = Application.Run("SAPGetProperty", "IsDataSourceActive", ds)
        lRet = lRet And lResult
        If lRet = False Then
            If ds_concat <> vbNullString Then
                ds_concat = ds_concat & ", " & vbCrLf & "'" & ds & "'"
            Else
                ds_concat = "'" & ds & "'"
            End If
        End If
    Next i
    
    If ds_concat <> vbNullString Then
        MsgBox "Data Sources: " & vbCrLf & ds_concat & vbCrLf & " are inactive"
    End If
    
    If lResult = False Then
        
        ThisWorkbook.Sheets("Edit_SP").Activate
        ActiveSheet.Range("A1").Activate
        
        lResult = Application.Run("SAPLogOff", "True")
        lResult = Application.Run("SAPSetRefreshBehaviour", "Off")
        lResult = Application.Run("SAPExecuteCommand", "Refresh", "ALL")
        
        Call DataValidationList
        
        lResult = Application.Run("SAPSetRefreshBehaviour", "On")
        
    Else
        InfoBox = TimedMsgBox("You are already connected to the system" _
                            & vbCrLf & vbCrLf & "(This window will close in " & AckTime & " seconds)", _
                              "Connection Status", , AckTime)
    End If
    
    Call Alignment
    Call OnEnd
    
ErrorHandler:
    ' Call the global error handling procedure
    Call HandleError
End Sub

' Procedure purpose:  To save data in Planning Query
Private Sub Save_SP_Click()

    Dim EndTime As Double
    Dim InfoBox As VbMsgBoxResult
    Dim StartTime As Double
    Dim ds_alias As String: ds_alias = "DS_5"
    Dim lResult As Long
    Dim wb As Workbook
    
    On Error GoTo ErrorHandler
    
    Const AckTime As Integer = 3
    
    Set wb = ThisWorkbook
    
    Call OnStart
    
    lResult = Application.Run("SAPGetProperty", "IsDataSourceActive", ds_alias)
    If lResult = True Then
        
        StartTime = Timer
        lResult = Application.Run("SAPSetRefreshBehaviour", "Off")
        lResult = Application.Run("SAPDeleteDesignRule", ds_alias)
        lResult = Application.Run("SAPGetProperty", "IsDataSourceEditable", ds_alias)
        If lResult = True Then
            lResult = Application.Run("SAPGetProperty", "HasChangedPlanData", ds_alias)
            If lResult = True Then
                lResult = Application.Run("SAPExecuteCommand", "PlanDataSave")
                lResult = Application.Run("SAPExecuteCommand", "Restart", "ALL")
            
                Call DataValidationList
                
                EndTime = Timer
                lResult = Application.Run("SAPSetRefreshBehaviour", "On")
                wb.Sheets("Edit_SP").Range("E1").Value = Format((EndTime - StartTime) / 86400, "hh:mm:ss") & " [hh:mm:ss]"
                
                InfoBox = TimedMsgBox("Data saved in : " & Format((EndTime - StartTime) / 86400, "hh:mm:ss") & " [hh:mm:ss]" _
                                    & vbCrLf & vbCrLf & "(This window will close in " & AckTime & " seconds)", _
                                      "Save Status", , AckTime)
            Else
                InfoBox = TimedMsgBox("No data has been changed" _
                                    & vbCrLf & vbCrLf & "(This window will close in " & AckTime & " seconds)", _
                                      "Save Status", , AckTime)
            End If
            
            lResult = Application.Run("SAPSetRefreshBehaviour", "On")
        
        Else
            InfoBox = TimedMsgBox("Cannot save the data, please check if the query is in 'change mode' (Analysis ribbon)" _
                                & vbCrLf & vbCrLf & "(This window will close in " & AckTime & " seconds)", _
                                  "Connection Status", , AckTime)
            lResult = Application.Run("SAPSetRefreshBehaviour", "On")
        End If
        
    Else
        InfoBox = TimedMsgBox("You are not connected to the system" _
                            & vbCrLf & vbCrLf & "(This window will close in " & AckTime & " seconds)", _
                              "Connection Status", , AckTime)
    End If
    
    Call Alignment
    Call OnEnd
    
ErrorHandler:
    ' Call the global error handling procedure
    Call HandleError
End Sub

' Procedure purpose:  To enable floating buttons
Private Sub Worksheet_SelectionChange(ByVal target As Excel.Range)

    On Error GoTo 0
    With Cells(Windows(1).ScrollRow, Windows(1).ScrollColumn)
        Connect_SP.Top = .Top + 0
        Connect_SP.Left = .Left + 0
        Save_SP.Top = .Top + 0
        Save_SP.Left = .Left + 80
    End With
End Sub




------------------
My_Module.bas

Attribute VB_Name = "My_Module"
Option Explicit
Global vFlag As Integer
Public gErrorNumber As Long
Public gErrorDescription As String
Public gErrorSource As String

Public Declare PtrSafe Function CustomTimeOffMsgBox Lib "user32" Alias "MessageBoxTimeoutA" ( _
ByVal xHwnd As LongPtr, _
ByVal xText As String, _
ByVal xCaption As String, _
ByVal xMsgBoxStyle As VbMsgBoxStyle, _
ByVal xwlange As Long, _
ByVal xTimeOut As Long) _
As Long

' Store the error details in global variables
Public Sub HandleError()
    gErrorNumber = Err.Number
    gErrorDescription = Err.Description
    gErrorSource = Erl & ": " & Err.Source
    
    ' Display or handle the error as per your requirements
    Debug.Print "Error Number: " & gErrorNumber & vbNewLine & _
           "Description: " & gErrorDescription & vbNewLine & _
           "Source: " & gErrorSource, vbCritical, "Error"
    
    ' Reset the error object
    Err.Clear
    Exit Sub
End Sub

' Function purpose:  To determine first filled row and last filled column
Public Function GetLastFilledColumnAndFirstFilledRow() As Variant
    
    Dim firstFilledRow As Long
    Dim lastFilledColumn As Long
    Dim lastNonEmptyRow As Long
    Dim ws As Worksheet
    
    On Error GoTo ErrorHandler
    Set ws = ActiveSheet
    
    ' Find the last filled row in the worksheet
    lastNonEmptyRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    ' Find the last filled column in the worksheet
    lastFilledColumn = ws.Cells(lastNonEmptyRow, ws.Columns.count).End(xlToLeft).Column
    
    ' Find the first filled row in the last filled column
    firstFilledRow = Cells(lastNonEmptyRow, lastFilledColumn).End(xlUp).Row
    
    ' Return the last filled column and first filled row as an array
    GetLastFilledColumnAndFirstFilledRow = Array(lastFilledColumn, firstFilledRow)
    
    Exit Function
    
ErrorHandler:
    ' Call the global error handling procedure
    Call HandleError
End Function

Public Function TimedMsgBox( _
       Prompt As String, _
       Optional Title As String = "Pop-up message", _
       Optional Buttons As VbMsgBoxStyle = vbOKOnly, _
       Optional Timeout As Long = 5000) _
        As VbMsgBoxResult
    ' Function purpose:  To create custom MsgBox with autoclose option
    
    TimedMsgBox = CustomTimeOffMsgBox(0&, Prompt, Title, Buttons, 0, Timeout * 1000)
    
End Function

' Function purpose:  To evaluate if a worksheet is protected
Public Function SheetProtected(TargetSheet As Worksheet) As Boolean
    
    On Error GoTo ErrorHandler
    
    SheetProtected = TargetSheet.ProtectContents
    
    Exit Function

ErrorHandler:
    ' Call the global error handling procedure
    Call HandleError
End Function

' Procedure purpose:  To unlock all worksheets in this workbook
Public Sub UnlockSheets()
    
    Dim wb          As Workbook
    Dim ws          As Worksheet
    
    On Error GoTo ErrorHandler
    
    Set wb = ThisWorkbook
    Set ws = ActiveSheet
    
    wb.Activate
    
    For Each ws In wb.Worksheets
        If ActiveSheet.ProtectContents = True Then
            ActiveSheet.Unprotect
        End If
        If ws.Cells.Locked = True Then
            ws.Cells.Locked = False
        End If
        If ws.Cells.FormulaHidden = True Then
            ws.Cells.FormulaHidden = False
        End If
    Next ws
    
    Exit Sub
    
ErrorHandler:
    ' Call the global error handling procedure
    Call HandleError
End Sub


' Procedure purpose:  To reconnect with the SAP data source
Public Sub Reconnect()
    Dim lResult     As Long
    
    On Error GoTo ErrorHandler
    
    lResult = Application.Run("SAPSetRefreshBehaviour", "Off")
    lResult = Application.Run("SAPLogOff", "True")
    lResult = Application.Run("SAPExecuteCommand", "Refresh", "All")
    If ActiveSheet.Name = "Edit_SP" Then
        Call DataValidationList
    End If
    lResult = Application.Run("SAPSetRefreshBehaviour", "On")
    
    Exit Sub
    
ErrorHandler:
    ' Call the global error handling procedure
    Call HandleError
End Sub

' Procedure purpose:  To disable immediate calculations, screen updates, events, messages
Public Sub OnStart()

    On Error GoTo ErrorHandler
    
    ThisWorkbook.Activate
    
    ActiveSheet.EnableCalculation = False
    Application.AskToUpdateLinks = False
    Application.Calculation = xlCalculationManual ' xlAutomatic
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    Call UnlockSheets
    
    Exit Sub
    
ErrorHandler:
    ' Call the global error handling procedure
    Call HandleError
    
End Sub

' Procedure purpose:  To enable immediate calculations, screen updates, events, messages
Public Sub OnEnd()
    
    On Error GoTo ErrorHandler
    
    ThisWorkbook.Activate
    
    ActiveSheet.EnableCalculation = True
    Application.AskToUpdateLinks = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    Exit Sub
    
ErrorHandler:
    ' Call the global error handling procedure
    Call HandleError
    
End Sub

' Procedure purpose:  To add data validation list on certain range (Approval flag)
Public Sub DataValidationList()
    
    Dim cell As Range
    Dim firstRow As Long
    Dim headerRow As Long
    Dim lastColumn As Long
    Dim llastrow As Long
    Dim searchRange As Range
    Dim targetColumn As Range
    Dim targetHeader As String
    Dim ws As Worksheet
    
    On Error GoTo ErrorHandler
    
    targetHeader = "Approval Flag"
    
    ' Set the search range as the entire worksheet
    Set searchRange = ActiveSheet.UsedRange
    
    ' Find the last used column in the search range
    lastColumn = searchRange.Columns(searchRange.Columns.count).Column
    
    ' Loop through each filled column to find the target header
    For headerRow = 1 To searchRange.Rows.count
        For Each cell In searchRange.Range(searchRange.Cells(headerRow, 1), searchRange.Cells(headerRow, lastColumn))
            If StrComp(cell.Value, targetHeader, vbTextCompare) = 0 Then
                ' Set the target column and row and exit the loop
                Set targetColumn = searchRange.Columns(cell.Column)
                firstRow = headerRow
                Exit For
            End If
        Next cell
        
        If Not targetColumn Is Nothing Then
            Exit For
        End If
    Next headerRow
    
    Set ws = ThisWorkbook.ActiveSheet
    llastrow = ActiveSheet.Range(ActiveSheet.Range("A65536").End(XlDirection.xlUp).Address).Row

    targetColumn.WrapText = True
    
    targetColumn.Select
    
    If SheetProtected(ws) Then
        Call UnlockSheets
    End If
    
    Set targetColumn = targetColumn.Resize(llastrow - firstRow).Offset(firstRow, 0)
    
    With targetColumn.Validation
        .Delete
        .Add _
        Type:=xlValidateList, _
        AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, _
        Formula1:="0,1,2"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With
    
    Call Comments
    
    Exit Sub
    
ErrorHandler:
    ' Call the global error handling procedure
    Call HandleError
    
End Sub

' Procedure purpose:  To add comment on "Approval Flag" header
Public Sub Comments()

    Dim cell As Range
    Dim firstRow As Long
    Dim headerRow As Long
    Dim lastColumn As Long
    Dim searchRange As Range
    Dim targetColumn As Range
    Dim targetHeader As String
    Dim targetHeaderCell As Range
    
    On Error GoTo ErrorHandler
    
    targetHeader = "Approval Flag"
    
    ' Set the search range as the entire worksheet
    Set searchRange = ActiveSheet.UsedRange
    
    ' Find the last used column in the search range
    lastColumn = searchRange.Columns(searchRange.Columns.count).Column
    
    ' Loop through each filled column to find the target header
    For headerRow = 1 To searchRange.Rows.count
        For Each cell In searchRange.Range(searchRange.Cells(headerRow, 1), searchRange.Cells(headerRow, lastColumn))
            If StrComp(cell.Value, targetHeader, vbTextCompare) = 0 Then
                ' Set the target column and row and exit the loop
                Set targetColumn = searchRange.Columns(cell.Column)
                firstRow = headerRow
                Exit For
            End If
        Next cell
        
        If Not targetColumn Is Nothing Then
            Exit For
        End If
    Next headerRow
    
    Set targetHeaderCell = targetColumn.Cells(firstRow)
    
    With targetHeaderCell
        .ClearComments
        If .Comment Is Nothing Then
            .AddCommentThreaded ( _
                                "0 - Not approved;" & vbLf & "1 - Pending approval;" & vbLf & "2 - Approved;" _
                                )
        End If
    End With

    Exit Sub
    
ErrorHandler:
    ' Call the global error handling procedure
    Call HandleError
End Sub

' Procedure purpose:  To align columns
Public Sub Alignment()

    Dim StartRow As String
    Dim cell As Range
    Dim col As Range
    Dim firstRow As Long
    Dim headerRow As Long
    Dim lastColumn As Long
    Dim llastcol As String
    Dim llastrow As Long
    Dim rng As Range
    Dim searchRange As Range
    Dim targetColumn As Range
    Dim targetHeader As String
    Dim targetHeaderCell As Range
    
    On Error GoTo ErrorHandler
    
    targetHeader = "Data Snapshot Time (UTC)"

    ' Set the search range as the entire worksheet
    Set searchRange = ActiveSheet.UsedRange
    
    ' Find the last used column in the search range
    lastColumn = searchRange.Columns(searchRange.Columns.count).Column
    
    ' Loop through each filled column to find the target header
    For headerRow = 1 To searchRange.Rows.count
        For Each cell In searchRange.Range(searchRange.Cells(headerRow, 1), searchRange.Cells(headerRow, lastColumn))
            If StrComp(cell.Value, targetHeader, vbTextCompare) = 0 Then
                ' Set the target column and row and exit the loop
                Set targetColumn = searchRange.Columns(cell.Column)
                firstRow = headerRow
                Exit For
            End If
        Next cell
        
        If Not targetColumn Is Nothing Then
            Exit For
        End If
    Next headerRow
    
    Set targetHeaderCell = targetColumn.Cells(firstRow)
    
    StartRow = "A" & firstRow
    
    llastrow = ActiveSheet.Range(ActiveSheet.Range("A65536").End(XlDirection.xlUp).Address).Row
    llastcol = Split(ActiveSheet.Cells(1, ActiveSheet.Range(StartRow).End(xlToRight).Column).Address, "$")(1)
    Set rng = ActiveSheet.Range(targetHeaderCell.Offset(, 1), ActiveSheet.Range(llastcol & llastrow))

    With rng
        .WrapText = True
        .HorizontalAlignment = xlLeft
        .ReadingOrder = xlContext
        With .Columns
            For Each col In .Columns
                col.EntireColumn.ColumnWidth = 30
            Next col
        End With
        .EntireRow.AutoFit
    End With

    Exit Sub
    
ErrorHandler:
    ' Call the global error handling procedure
    Call HandleError
End Sub

' Procedure purpose:  To change columns visiblity
Public Sub Columns_Visibility()
    
    On Error GoTo ErrorHandler
    
    Dim StartRow As Integer
    Dim lastColumn As Variant
    Dim lastcell As Range
    Dim result As Variant
    Dim searchRange As Range
    Dim targetColumn As Range
    Dim targetColumnLetter As String
    Dim targetColumnSet As Range       ' Variable to hold each set of target column
    Dim targetColumns As Collection    ' Store multiple sets of target columns
    Dim targetHeader As String
    Dim targetHeaderCell As Range
    
    result = GetLastFilledColumnAndFirstFilledRow()
    StartRow = result(1)
    targetHeader = "(extended)"

    ' Set the search range as the entire worksheet
    Set searchRange = ActiveSheet.UsedRange
    
    ' Find the last used column in the search range
    Set lastcell = ActiveSheet.Cells(StartRow, ActiveSheet.Range("A" & StartRow).End(xlToRight).Column)
    lastColumn = Split(lastcell.Address, "$")(1)
    If lastcell.Value <> vbNullString Then
        targetColumnLetter = Split(ActiveSheet.Cells(1, lastColumn).Offset(, 1).Address, "$")(1)
    Else
        targetColumnLetter = Split(ActiveSheet.Cells(1, lastColumn).Address, "$")(1)
    End If
    
    If (Right(ActiveSheet.Name, 2) = "CA") Or (Right(ActiveSheet.Name, 2) = "SP") Then
        Call Alignment
    End If
    
    If ActiveSheet.Name = "Export_CA" Then
        ActiveSheet.Range(targetColumnLetter & ":XFD").EntireColumn.Hidden = True
        ' Initialize the collection to store target columns
        Set targetColumns = New Collection

        ' Loop through each filled column to find the target header
        For Each targetColumn In searchRange.Columns
            ' Find the target header in the current column
            Set targetHeaderCell = targetColumn.Find(targetHeader, LookIn:=xlValues, LookAt:=xlPart)
            
            If Not targetHeaderCell Is Nothing Then
                ' Add the target column to the collection
                targetColumns.Add targetHeaderCell.EntireColumn
            End If
        Next targetColumn
        
        ' Loop through each set of target columns
        For Each targetColumnSet In targetColumns
            ' Loop through each target column in the set
            For Each targetColumn In targetColumnSet
                targetColumnLetter = targetColumn.Address
                ActiveSheet.Range(targetColumnLetter).EntireColumn.Hidden = True
            Next targetColumn
        Next targetColumnSet
        
    ElseIf (ActiveSheet.Name = "Export_SP") Or (ActiveSheet.Name Like "Changelog*") Then
        ActiveSheet.Range(targetColumnLetter & ":XFD").EntireColumn.Hidden = True
    End If

    Exit Sub
    
ErrorHandler:
    ' Call the global error handling procedure
    Call HandleError
End Sub

' Procedure purpose:  To replace "#" characters with blank value in particular range on "Export" tab
Public Sub RemoveHashCharacters(rng As Range)

    Dim Cell_1 As String
    Dim Cell_2 As String
    Dim colcnt As Integer
    Dim DataRange As Variant
    Dim Icol As Integer
    Dim Irow As Long
    Dim rowcnt As Integer
    
    On Error GoTo ErrorHandler
    
    Call ReplaceHeaderValue
    
    DataRange = rng.Value
    rowcnt = rng.Rows.count
    colcnt = rng.Columns.count
    
    For Irow = 1 To rowcnt
        For Icol = 1 To colcnt
            Cell_1 = DataRange(Irow, Icol)
            If Cell_1 <> vbNullString Then
                If Cell_1 = "#" Then
                    Cell_1 = vbNullString
                End If
                DataRange(Irow, Icol) = Cell_1
            End If
        Next Icol
    Next Irow
    rng.Value = DataRange
    
    rng.WrapText = True
    rng.Columns.EntireColumn.AutoFit
    rng.Rows.EntireRow.AutoFit

    Exit Sub
    
ErrorHandler:
    ' Call the global error handling procedure
    Call HandleError
End Sub

' Procedure purpose:  To replace "&&" with line breaks (VbCrLf) in certain range on the "Export" tab _
Public Sub RestoreLineBreaks(rng As Range)
    and to merge cells that were splitted due to max. field length (250)
    
    On Error GoTo ErrorHandler
    
    Dim Cell_1 As String
    Dim Cell_2 As String
    Dim DataRange As Variant
    Dim Icol As Integer
    Dim Irow As Long
    Dim colcnt As Integer
    Dim rowcnt As Integer
    
    Const LineSeparator As String = "&&"

    Set rng = rng
    rowcnt = rng.Rows.count
    colcnt = rng.Columns.count
    
    DataRange = rng.Value
    
    For Irow = 1 To rowcnt
        For Icol = 1 To colcnt
            Cell_1 = DataRange(Irow, Icol)
            If Cell_1 <> vbNullString Then
                If (Icol = 1 Or Icol = 3) Then
                    Cell_2 = DataRange(Irow, Icol + 1)
                    Cell_1 = Cell_1 & Cell_2
                End If
                
                If (Icol = 2 Or Icol = 4) Then
                    Cell_1 = Empty
                End If
                If Len(Replace(Cell_1, LineSeparator, vbNullString)) <> Len(Cell_1) Then
                    Cell_1 = Replace(Cell_1, LineSeparator, vbCrLf)
                End If
                DataRange(Irow, Icol) = Cell_1
            End If
        Next Icol
    Next Irow
    rng.Value = DataRange

    Exit Sub
    
ErrorHandler:
    ' Call the global error handling procedure
    Call HandleError
End Sub

' Procedure purpose:  To set print area, header information (Query last refresh date & time) on the "Export" tab
Public Sub SetPrintLayout(lRefreshDate As Double)
    
    Dim llastrow    As Long
    Dim rng         As Range
    Dim ws          As Worksheet
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.ActiveSheet
    
    ws.Activate
    
    llastrow = ActiveSheet.Range(ActiveSheet.Range("A65536").End(XlDirection.xlUp).Address).Row
    Set rng = ActiveSheet.Range("A1", ActiveSheet.Range("A" & llastrow).End(xlToRight))
    ws.PageSetup.PrintArea = rng.Address
    ws.PageSetup.CenterHeader = "QueryLastRefreshedAt: " & Format(lRefreshDate, "dddd, mmmm d, yyyy h:mm:ss")

    Exit Sub
    
ErrorHandler:
    ' Call the global error handling procedure
    Call HandleError
End Sub

Public Sub ReplaceHeaderValue()

    Dim cell As Range
    Dim dataArr As Variant
    Dim firstRow As Long
    Dim headerRow As Long
    Dim i As Long, j As Long
    Dim lastColumn As Long
    Dim searchRange As Range
    Dim targetColumn As Range
    Dim targetHeader As String
    
    On Error GoTo ErrorHandler
    
    targetHeader = "Data Snapshot Time (UTC)"
    
    ' Set the search range as the entire worksheet
    Set searchRange = ActiveSheet.UsedRange
    
    ' Find the last used column in the search range
    lastColumn = searchRange.Columns(searchRange.Columns.count).Column
    
    ' Loop through each filled column to find the target header
    For headerRow = 1 To searchRange.Rows.count
        For Each cell In searchRange.Range(searchRange.Cells(headerRow, 1), searchRange.Cells(headerRow, lastColumn))
            If StrComp(cell.Value, targetHeader, vbTextCompare) = 0 Then
                ' Set the target column and row and exit the loop
                Set targetColumn = searchRange.Columns(cell.Column)
                firstRow = headerRow
                Exit For
            End If
        Next cell
        
        If Not targetColumn Is Nothing Then
            Exit For
        End If
    Next headerRow
    
    ' Check if the target header is found
    If Not targetColumn Is Nothing Then
        ' Read the data from the target column into an array
        dataArr = targetColumn.Value
        
        ' Loop through the array to replace "#" with "00:00:00"
        For i = 2 To UBound(dataArr, 1)
            If dataArr(i, 1) = "#" Then
                dataArr(i, 1) = "00:00:00"
            End If
        Next i
        
        ' Update the values in the target column with the modified array
        targetColumn.Value = dataArr
        
        ' Select the entire column
        targetColumn.Select
    End If

    Exit Sub
    
ErrorHandler:
    ' Call the global error handling procedure
    Call HandleError
End Sub




------------------
Privileged_Access_Report_VBA_Code.vb

------------------
Edit_CA.cls

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Edit_CA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@IgnoreModule UndeclaredVariable, UnassignedVariableUsage
Option Explicit

' Procedure purpose:  To reconnect/refresh data sources
Private Sub Connect_CA_Click()

    Dim lResult As Long
    Dim lRet As Boolean
    Dim i As Integer
    Dim ds As String
    Dim ds_concat As String
    Dim InfoBox As VbMsgBoxResult
    
    On Error GoTo ErrorHandler
    
    Const AckTime As Integer = 3
    
    Call OnStart
    
    lRet = True
    For i = 1 To 7
        ds = "DS_" & i
        lResult = Application.Run("SAPGetProperty", "IsDataSourceActive", ds)
        lRet = lRet And lResult
        If lRet = False Then
            If ds_concat <> vbNullString Then
                ds_concat = ds_concat & ", " & vbCrLf & "'" & ds & "'"
            Else
                ds_concat = "'" & ds & "'"
            End If
        End If
    Next i
    
    If ds_concat <> vbNullString Then
        MsgBox "Data Sources: " & vbCrLf & ds_concat & vbCrLf & " are inactive"
    End If
    
    If lResult = False Then
        
        ThisWorkbook.Sheets("Edit_CA").Activate
        ActiveSheet.Range("A1").Activate
        
        lResult = Application.Run("SAPLogOff", "True")
        lResult = Application.Run("SAPSetRefreshBehaviour", "Off")
        lResult = Application.Run("SAPExecuteCommand", "Refresh", "ALL")
        
        Call DataValidationList
        
        lResult = Application.Run("SAPSetRefreshBehaviour", "On")
        
    Else
        InfoBox = TimedMsgBox("You are already connected to the system" _
                            & vbCrLf & vbCrLf & "(This window will close in " & AckTime & " seconds)", _
                              "Connection Status", , AckTime)
    End If
    
    Call Alignment
    Call OnEnd
ErrorHandler:
    ' Call the global error handling procedure
    Call HandleError
End Sub

' Procedure purpose:  To save data in Planning Query
Private Sub Save_CA_Click()

    Dim EndTime As Double
    Dim StartTime As Double
    Dim ds_alias As String: ds_alias = "DS_3"
    Dim lResult As Long
    Dim wb As Workbook
    Dim InfoBox As VbMsgBoxResult
    
    On Error GoTo ErrorHandler
    
    Const AckTime As Integer = 3
    
    Set wb = ThisWorkbook
    
    Call OnStart
    
    lResult = Application.Run("SAPGetProperty", "IsDataSourceActive", ds_alias)
    If lResult = True Then
        
        StartTime = Timer
        lResult = Application.Run("SAPSetRefreshBehaviour", "Off")
        lResult = Application.Run("SAPDeleteDesignRule", ds_alias)
        lResult = Application.Run("SAPGetProperty", "IsDataSourceEditable", ds_alias)
        If lResult = True Then
            lResult = Application.Run("SAPGetProperty", "HasChangedPlanData", ds_alias)
            If lResult = True Then
                lResult = Application.Run("SAPExecuteCommand", "PlanDataSave")
                lResult = Application.Run("SAPExecuteCommand", "Restart", "ALL")
            
                Call DataValidationList
                
                EndTime = Timer
                lResult = Application.Run("SAPSetRefreshBehaviour", "On")
                wb.Sheets("Edit_CA").Range("E1").Value = Format((EndTime - StartTime) / 86400, "hh:mm:ss") & " [hh:mm:ss]"
                
                InfoBox = TimedMsgBox("Data saved in : " & Format((EndTime - StartTime) / 86400, "hh:mm:ss") & " [hh:mm:ss]" _
                                    & vbCrLf & vbCrLf & "(This window will close in " & AckTime & " seconds)", _
                                      "Save Status", , AckTime)
            Else
                InfoBox = TimedMsgBox("No data has been changed" _
                                    & vbCrLf & vbCrLf & "(This window will close in " & AckTime & " seconds)", _
                                      "Save Status", , AckTime)
            End If
            
            lResult = Application.Run("SAPSetRefreshBehaviour", "On")
        
        Else
            InfoBox = TimedMsgBox("Cannot save the data, please check if the query is in 'change mode' (Analysis ribbon)" _
                                & vbCrLf & vbCrLf & "(This window will close in " & AckTime & " seconds)", _
                                  "Connection Status", , AckTime)
            lResult = Application.Run("SAPSetRefreshBehaviour", "On")
        End If
        
    Else
        InfoBox = TimedMsgBox("You are not connected to the system" _
                            & vbCrLf & vbCrLf & "(This window will close in " & AckTime & " seconds)", _
                              "Connection Status", , AckTime)
    End If
    
    Call Alignment
    Call OnEnd
    
ErrorHandler:
    ' Call the global error handling procedure
    Call HandleError
End Sub

' Procedure purpose:  To enable floating buttons
Private Sub Worksheet_SelectionChange(ByVal target As Excel.Range)

    On Error GoTo 0
    With Cells(Windows(1).ScrollRow, Windows(1).ScrollColumn)
        Connect_CA.Top = .Top + 0
        Connect_CA.Left = .Left + 0
        Save_CA.Top = .Top + 0
        Save_CA.Left = .Left + 80
    End With
End Sub




------------------
Edit_SP.cls

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Edit_SP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@IgnoreModule UndeclaredVariable
' Procedure purpose:  To reconnect/refresh data sources
Private Sub Connect_SP_Click()

    Dim InfoBox As VbMsgBoxResult
    Dim ds As String
    Dim ds_concat As String
    Dim i As Integer
    Dim lResult As Long
    Dim lRet As Boolean
    
    On Error GoTo ErrorHandler
    
    Const AckTime As Integer = 3
    
    Call OnStart
    
    lRet = True
    For i = 1 To 7
        ds = "DS_" & i
        lResult = Application.Run("SAPGetProperty", "IsDataSourceActive", ds)
        lRet = lRet And lResult
        If lRet = False Then
            If ds_concat <> vbNullString Then
                ds_concat = ds_concat & ", " & vbCrLf & "'" & ds & "'"
            Else
                ds_concat = "'" & ds & "'"
            End If
        End If
    Next i
    
    If ds_concat <> vbNullString Then
        MsgBox "Data Sources: " & vbCrLf & ds_concat & vbCrLf & " are inactive"
    End If
    
    If lResult = False Then
        
        ThisWorkbook.Sheets("Edit_SP").Activate
        ActiveSheet.Range("A1").Activate
        
        lResult = Application.Run("SAPLogOff", "True")
        lResult = Application.Run("SAPSetRefreshBehaviour", "Off")
        lResult = Application.Run("SAPExecuteCommand", "Refresh", "ALL")
        
        Call DataValidationList
        
        lResult = Application.Run("SAPSetRefreshBehaviour", "On")
        
    Else
        InfoBox = TimedMsgBox("You are already connected to the system" _
                            & vbCrLf & vbCrLf & "(This window will close in " & AckTime & " seconds)", _
                              "Connection Status", , AckTime)
    End If
    
    Call Alignment
    Call OnEnd
    
ErrorHandler:
    ' Call the global error handling procedure
    Call HandleError
End Sub

' Procedure purpose:  To save data in Planning Query
Private Sub Save_SP_Click()

    Dim EndTime As Double
    Dim InfoBox As VbMsgBoxResult
    Dim StartTime As Double
    Dim ds_alias As String: ds_alias = "DS_5"
    Dim lResult As Long
    Dim wb As Workbook
    
    On Error GoTo ErrorHandler
    
    Const AckTime As Integer = 3
    
    Set wb = ThisWorkbook
    
    Call OnStart
    
    lResult = Application.Run("SAPGetProperty", "IsDataSourceActive", ds_alias)
    If lResult = True Then
        
        StartTime = Timer
        lResult = Application.Run("SAPSetRefreshBehaviour", "Off")
        lResult = Application.Run("SAPDeleteDesignRule", ds_alias)
        lResult = Application.Run("SAPGetProperty", "IsDataSourceEditable", ds_alias)
        If lResult = True Then
            lResult = Application.Run("SAPGetProperty", "HasChangedPlanData", ds_alias)
            If lResult = True Then
                lResult = Application.Run("SAPExecuteCommand", "PlanDataSave")
                lResult = Application.Run("SAPExecuteCommand", "Restart", "ALL")
            
                Call DataValidationList
                
                EndTime = Timer
                lResult = Application.Run("SAPSetRefreshBehaviour", "On")
                wb.Sheets("Edit_SP").Range("E1").Value = Format((EndTime - StartTime) / 86400, "hh:mm:ss") & " [hh:mm:ss]"
                
                InfoBox = TimedMsgBox("Data saved in : " & Format((EndTime - StartTime) / 86400, "hh:mm:ss") & " [hh:mm:ss]" _
                                    & vbCrLf & vbCrLf & "(This window will close in " & AckTime & " seconds)", _
                                      "Save Status", , AckTime)
            Else
                InfoBox = TimedMsgBox("No data has been changed" _
                                    & vbCrLf & vbCrLf & "(This window will close in " & AckTime & " seconds)", _
                                      "Save Status", , AckTime)
            End If
            
            lResult = Application.Run("SAPSetRefreshBehaviour", "On")
        
        Else
            InfoBox = TimedMsgBox("Cannot save the data, please check if the query is in 'change mode' (Analysis ribbon)" _
                                & vbCrLf & vbCrLf & "(This window will close in " & AckTime & " seconds)", _
                                  "Connection Status", , AckTime)
            lResult = Application.Run("SAPSetRefreshBehaviour", "On")
        End If
        
    Else
        InfoBox = TimedMsgBox("You are not connected to the system" _
                            & vbCrLf & vbCrLf & "(This window will close in " & AckTime & " seconds)", _
                              "Connection Status", , AckTime)
    End If
    
    Call Alignment
    Call OnEnd
    
ErrorHandler:
    ' Call the global error handling procedure
    Call HandleError
End Sub

' Procedure purpose:  To enable floating buttons
Private Sub Worksheet_SelectionChange(ByVal target As Excel.Range)

    On Error GoTo 0
    With Cells(Windows(1).ScrollRow, Windows(1).ScrollColumn)
        Connect_SP.Top = .Top + 0
        Connect_SP.Left = .Left + 0
        Save_SP.Top = .Top + 0
        Save_SP.Left = .Left + 80
    End With
End Sub




------------------
My_Module.bas

Attribute VB_Name = "My_Module"
Option Explicit
Global vFlag As Integer
Public gErrorNumber As Long
Public gErrorDescription As String
Public gErrorSource As String

Public Declare PtrSafe Function CustomTimeOffMsgBox Lib "user32" Alias "MessageBoxTimeoutA" ( _
ByVal xHwnd As LongPtr, _
ByVal xText As String, _
ByVal xCaption As String, _
ByVal xMsgBoxStyle As VbMsgBoxStyle, _
ByVal xwlange As Long, _
ByVal xTimeOut As Long) _
As Long

' Store the error details in global variables
Public Sub HandleError()
    gErrorNumber = Err.Number
    gErrorDescription = Err.Description
    gErrorSource = Erl & ": " & Err.Source
    
    ' Display or handle the error as per your requirements
    Debug.Print "Error Number: " & gErrorNumber & vbNewLine & _
           "Description: " & gErrorDescription & vbNewLine & _
           "Source: " & gErrorSource, vbCritical, "Error"
    
    ' Reset the error object
    Err.Clear
    Exit Sub
End Sub

' Function purpose:  To determine first filled row and last filled column
Public Function GetLastFilledColumnAndFirstFilledRow() As Variant
    
    Dim firstFilledRow As Long
    Dim lastFilledColumn As Long
    Dim lastNonEmptyRow As Long
    Dim ws As Worksheet
    
    On Error GoTo ErrorHandler
    Set ws = ActiveSheet
    
    ' Find the last filled row in the worksheet
    lastNonEmptyRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    ' Find the last filled column in the worksheet
    lastFilledColumn = ws.Cells(lastNonEmptyRow, ws.Columns.count).End(xlToLeft).Column
    
    ' Find the first filled row in the last filled column
    firstFilledRow = Cells(lastNonEmptyRow, lastFilledColumn).End(xlUp).Row
    
    ' Return the last filled column and first filled row as an array
    GetLastFilledColumnAndFirstFilledRow = Array(lastFilledColumn, firstFilledRow)
    
    Exit Function
    
ErrorHandler:
    ' Call the global error handling procedure
    Call HandleError
End Function

Public Function TimedMsgBox( _
       Prompt As String, _
       Optional Title As String = "Pop-up message", _
       Optional Buttons As VbMsgBoxStyle = vbOKOnly, _
       Optional Timeout As Long = 5000) _
        As VbMsgBoxResult
    ' Function purpose:  To create custom MsgBox with autoclose option
    
    TimedMsgBox = CustomTimeOffMsgBox(0&, Prompt, Title, Buttons, 0, Timeout * 1000)
    
End Function

' Function purpose:  To evaluate if a worksheet is protected
Public Function SheetProtected(TargetSheet As Worksheet) As Boolean
    
    On Error GoTo ErrorHandler
    
    SheetProtected = TargetSheet.ProtectContents
    
    Exit Function

ErrorHandler:
    ' Call the global error handling procedure
    Call HandleError
End Function

' Procedure purpose:  To unlock all worksheets in this workbook
Public Sub UnlockSheets()
    
    Dim wb          As Workbook
    Dim ws          As Worksheet
    
    On Error GoTo ErrorHandler
    
    Set wb = ThisWorkbook
    Set ws = ActiveSheet
    
    wb.Activate
    
    For Each ws In wb.Worksheets
        If ActiveSheet.ProtectContents = True Then
            ActiveSheet.Unprotect
        End If
        If ws.Cells.Locked = True Then
            ws.Cells.Locked = False
        End If
        If ws.Cells.FormulaHidden = True Then
            ws.Cells.FormulaHidden = False
        End If
    Next ws
    
    Exit Sub
    
ErrorHandler:
    ' Call the global error handling procedure
    Call HandleError
End Sub


' Procedure purpose:  To reconnect with the SAP data source
Public Sub Reconnect()
    Dim lResult     As Long
    
    On Error GoTo ErrorHandler
    
    lResult = Application.Run("SAPSetRefreshBehaviour", "Off")
    lResult = Application.Run("SAPLogOff", "True")
    lResult = Application.Run("SAPExecuteCommand", "Refresh", "All")
    If ActiveSheet.Name = "Edit_SP" Then
        Call DataValidationList
    End If
    lResult = Application.Run("SAPSetRefreshBehaviour", "On")
    
    Exit Sub
    
ErrorHandler:
    ' Call the global error handling procedure
    Call HandleError
End Sub

' Procedure purpose:  To disable immediate calculations, screen updates, events, messages
Public Sub OnStart()

    On Error GoTo ErrorHandler
    
    ThisWorkbook.Activate
    
    ActiveSheet.EnableCalculation = False
    Application.AskToUpdateLinks = False
    Application.Calculation = xlCalculationManual ' xlAutomatic
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    Call UnlockSheets
    
    Exit Sub
    
ErrorHandler:
    ' Call the global error handling procedure
    Call HandleError
    
End Sub

' Procedure purpose:  To enable immediate calculations, screen updates, events, messages
Public Sub OnEnd()
    
    On Error GoTo ErrorHandler
    
    ThisWorkbook.Activate
    
    ActiveSheet.EnableCalculation = True
    Application.AskToUpdateLinks = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    Exit Sub
    
ErrorHandler:
    ' Call the global error handling procedure
    Call HandleError
    
End Sub

' Procedure purpose:  To add data validation list on certain range (Approval flag)
Public Sub DataValidationList()
    
    Dim cell As Range
    Dim firstRow As Long
    Dim headerRow As Long
    Dim lastColumn As Long
    Dim llastrow As Long
    Dim searchRange As Range
    Dim targetColumn As Range
    Dim targetHeader As String
    Dim ws As Worksheet
    
    On Error GoTo ErrorHandler
    
    targetHeader = "Approval Flag"
    
    ' Set the search range as the entire worksheet
    Set searchRange = ActiveSheet.UsedRange
    
    ' Find the last used column in the search range
    lastColumn = searchRange.Columns(searchRange.Columns.count).Column
    
    ' Loop through each filled column to find the target header
    For headerRow = 1 To searchRange.Rows.count
        For Each cell In searchRange.Range(searchRange.Cells(headerRow, 1), searchRange.Cells(headerRow, lastColumn))
            If StrComp(cell.Value, targetHeader, vbTextCompare) = 0 Then
                ' Set the target column and row and exit the loop
                Set targetColumn = searchRange.Columns(cell.Column)
                firstRow = headerRow
                Exit For
            End If
        Next cell
        
        If Not targetColumn Is Nothing Then
            Exit For
        End If
    Next headerRow
    
    Set ws = ThisWorkbook.ActiveSheet
    llastrow = ActiveSheet.Range(ActiveSheet.Range("A65536").End(XlDirection.xlUp).Address).Row

    targetColumn.WrapText = True
    
    targetColumn.Select
    
    If SheetProtected(ws) Then
        Call UnlockSheets
    End If
    
    Set targetColumn = targetColumn.Resize(llastrow - firstRow).Offset(firstRow, 0)
    
    With targetColumn.Validation
        .Delete
        .Add _
        Type:=xlValidateList, _
        AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, _
        Formula1:="0,1,2"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With
    
    Call Comments
    
    Exit Sub
    
ErrorHandler:
    ' Call the global error handling procedure
    Call HandleError
    
End Sub

' Procedure purpose:  To add comment on "Approval Flag" header
Public Sub Comments()

    Dim cell As Range
    Dim firstRow As Long
    Dim headerRow As Long
    Dim lastColumn As Long
    Dim searchRange As Range
    Dim targetColumn As Range
    Dim targetHeader As String
    Dim targetHeaderCell As Range
    
    On Error GoTo ErrorHandler
    
    targetHeader = "Approval Flag"
    
    ' Set the search range as the entire worksheet
    Set searchRange = ActiveSheet.UsedRange
    
    ' Find the last used column in the search range
    lastColumn = searchRange.Columns(searchRange.Columns.count).Column
    
    ' Loop through each filled column to find the target header
    For headerRow = 1 To searchRange.Rows.count
        For Each cell In searchRange.Range(searchRange.Cells(headerRow, 1), searchRange.Cells(headerRow, lastColumn))
            If StrComp(cell.Value, targetHeader, vbTextCompare) = 0 Then
                ' Set the target column and row and exit the loop
                Set targetColumn = searchRange.Columns(cell.Column)
                firstRow = headerRow
                Exit For
            End If
        Next cell
        
        If Not targetColumn Is Nothing Then
            Exit For
        End If
    Next headerRow
    
    Set targetHeaderCell = targetColumn.Cells(firstRow)
    
    With targetHeaderCell
        .ClearComments
        If .Comment Is Nothing Then
            .AddCommentThreaded ( _
                                "0 - Not approved;" & vbLf & "1 - Pending approval;" & vbLf & "2 - Approved;" _
                                )
        End If
    End With

    Exit Sub
    
ErrorHandler:
    ' Call the global error handling procedure
    Call HandleError
End Sub

' Procedure purpose:  To align columns
Public Sub Alignment()

    Dim StartRow As String
    Dim cell As Range
    Dim col As Range
    Dim firstRow As Long
    Dim headerRow As Long
    Dim lastColumn As Long
    Dim llastcol As String
    Dim llastrow As Long
    Dim rng As Range
    Dim searchRange As Range
    Dim targetColumn As Range
    Dim targetHeader As String
    Dim targetHeaderCell As Range
    
    On Error GoTo ErrorHandler
    
    targetHeader = "Data Snapshot Time (UTC)"

    ' Set the search range as the entire worksheet
    Set searchRange = ActiveSheet.UsedRange
    
    ' Find the last used column in the search range
    lastColumn = searchRange.Columns(searchRange.Columns.count).Column
    
    ' Loop through each filled column to find the target header
    For headerRow = 1 To searchRange.Rows.count
        For Each cell In searchRange.Range(searchRange.Cells(headerRow, 1), searchRange.Cells(headerRow, lastColumn))
            If StrComp(cell.Value, targetHeader, vbTextCompare) = 0 Then
                ' Set the target column and row and exit the loop
                Set targetColumn = searchRange.Columns(cell.Column)
                firstRow = headerRow
                Exit For
            End If
        Next cell
        
        If Not targetColumn Is Nothing Then
            Exit For
        End If
    Next headerRow
    
    Set targetHeaderCell = targetColumn.Cells(firstRow)
    
    StartRow = "A" & firstRow
    
    llastrow = ActiveSheet.Range(ActiveSheet.Range("A65536").End(XlDirection.xlUp).Address).Row
    llastcol = Split(ActiveSheet.Cells(1, ActiveSheet.Range(StartRow).End(xlToRight).Column).Address, "$")(1)
    Set rng = ActiveSheet.Range(targetHeaderCell.Offset(, 1), ActiveSheet.Range(llastcol & llastrow))

    With rng
        .WrapText = True
        .HorizontalAlignment = xlLeft
        .ReadingOrder = xlContext
        With .Columns
            For Each col In .Columns
                col.EntireColumn.ColumnWidth = 30
            Next col
        End With
        .EntireRow.AutoFit
    End With

    Exit Sub
    
ErrorHandler:
    ' Call the global error handling procedure
    Call HandleError
End Sub

' Procedure purpose:  To change columns visiblity
Public Sub Columns_Visibility()
    
    On Error GoTo ErrorHandler
    
    Dim StartRow As Integer
    Dim lastColumn As Variant
    Dim lastcell As Range
    Dim result As Variant
    Dim searchRange As Range
    Dim targetColumn As Range
    Dim targetColumnLetter As String
    Dim targetColumnSet As Range       ' Variable to hold each set of target column
    Dim targetColumns As Collection    ' Store multiple sets of target columns
    Dim targetHeader As String
    Dim targetHeaderCell As Range
    
    result = GetLastFilledColumnAndFirstFilledRow()
    StartRow = result(1)
    targetHeader = "(extended)"

    ' Set the search range as the entire worksheet
    Set searchRange = ActiveSheet.UsedRange
    
    ' Find the last used column in the search range
    Set lastcell = ActiveSheet.Cells(StartRow, ActiveSheet.Range("A" & StartRow).End(xlToRight).Column)
    lastColumn = Split(lastcell.Address, "$")(1)
    If lastcell.Value <> vbNullString Then
        targetColumnLetter = Split(ActiveSheet.Cells(1, lastColumn).Offset(, 1).Address, "$")(1)
    Else
        targetColumnLetter = Split(ActiveSheet.Cells(1, lastColumn).Address, "$")(1)
    End If
    
    If (Right(ActiveSheet.Name, 2) = "CA") Or (Right(ActiveSheet.Name, 2) = "SP") Then
        Call Alignment
    End If
    
    If ActiveSheet.Name = "Export_CA" Then
        ActiveSheet.Range(targetColumnLetter & ":XFD").EntireColumn.Hidden = True
        ' Initialize the collection to store target columns
        Set targetColumns = New Collection

        ' Loop through each filled column to find the target header
        For Each targetColumn In searchRange.Columns
            ' Find the target header in the current column
            Set targetHeaderCell = targetColumn.Find(targetHeader, LookIn:=xlValues, LookAt:=xlPart)
            
            If Not targetHeaderCell Is Nothing Then
                ' Add the target column to the collection
                targetColumns.Add targetHeaderCell.EntireColumn
            End If
        Next targetColumn
        
        ' Loop through each set of target columns
        For Each targetColumnSet In targetColumns
            ' Loop through each target column in the set
            For Each targetColumn In targetColumnSet
                targetColumnLetter = targetColumn.Address
                ActiveSheet.Range(targetColumnLetter).EntireColumn.Hidden = True
            Next targetColumn
        Next targetColumnSet
        
    ElseIf (ActiveSheet.Name = "Export_SP") Or (ActiveSheet.Name Like "Changelog*") Then
        ActiveSheet.Range(targetColumnLetter & ":XFD").EntireColumn.Hidden = True
    End If

    Exit Sub
    
ErrorHandler:
    ' Call the global error handling procedure
    Call HandleError
End Sub

' Procedure purpose:  To replace "#" characters with blank value in particular range on "Export" tab
Public Sub RemoveHashCharacters(rng As Range)

    Dim Cell_1 As String
    Dim Cell_2 As String
    Dim colcnt As Integer
    Dim DataRange As Variant
    Dim Icol As Integer
    Dim Irow As Long
    Dim rowcnt As Integer
    
    On Error GoTo ErrorHandler
    
    Call ReplaceHeaderValue
    
    DataRange = rng.Value
    rowcnt = rng.Rows.count
    colcnt = rng.Columns.count
    
    For Irow = 1 To rowcnt
        For Icol = 1 To colcnt
            Cell_1 = DataRange(Irow, Icol)
            If Cell_1 <> vbNullString Then
                If Cell_1 = "#" Then
                    Cell_1 = vbNullString
                End If
                DataRange(Irow, Icol) = Cell_1
            End If
        Next Icol
    Next Irow
    rng.Value = DataRange
    
    rng.WrapText = True
    rng.Columns.EntireColumn.AutoFit
    rng.Rows.EntireRow.AutoFit

    Exit Sub
    
ErrorHandler:
    ' Call the global error handling procedure
    Call HandleError
End Sub

' Procedure purpose:  To replace "&&" with line breaks (VbCrLf) in certain range on the "Export" tab _
Public Sub RestoreLineBreaks(rng As Range)
    and to merge cells that were splitted due to max. field length (250)
    
    On Error GoTo ErrorHandler
    
    Dim Cell_1 As String
    Dim Cell_2 As String
    Dim DataRange As Variant
    Dim Icol As Integer
    Dim Irow As Long
    Dim colcnt As Integer
    Dim rowcnt As Integer
    
    Const LineSeparator As String = "&&"

    Set rng = rng
    rowcnt = rng.Rows.count
    colcnt = rng.Columns.count
    
    DataRange = rng.Value
    
    For Irow = 1 To rowcnt
        For Icol = 1 To colcnt
            Cell_1 = DataRange(Irow, Icol)
            If Cell_1 <> vbNullString Then
                If (Icol = 1 Or Icol = 3) Then
                    Cell_2 = DataRange(Irow, Icol + 1)
                    Cell_1 = Cell_1 & Cell_2
                End If
                
                If (Icol = 2 Or Icol = 4) Then
                    Cell_1 = Empty
                End If
                If Len(Replace(Cell_1, LineSeparator, vbNullString)) <> Len(Cell_1) Then
                    Cell_1 = Replace(Cell_1, LineSeparator, vbCrLf)
                End If
                DataRange(Irow, Icol) = Cell_1
            End If
        Next Icol
    Next Irow
    rng.Value = DataRange

    Exit Sub
    
ErrorHandler:
    ' Call the global error handling procedure
    Call HandleError
End Sub

' Procedure purpose:  To set print area, header information (Query last refresh date & time) on the "Export" tab
Public Sub SetPrintLayout(lRefreshDate As Double)
    
    Dim llastrow    As Long
    Dim rng         As Range
    Dim ws          As Worksheet
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.ActiveSheet
    
    ws.Activate
    
    llastrow = ActiveSheet.Range(ActiveSheet.Range("A65536").End(XlDirection.xlUp).Address).Row
    Set rng = ActiveSheet.Range("A1", ActiveSheet.Range("A" & llastrow).End(xlToRight))
    ws.PageSetup.PrintArea = rng.Address
    ws.PageSetup.CenterHeader = "QueryLastRefreshedAt: " & Format(lRefreshDate, "dddd, mmmm d, yyyy h:mm:ss")

    Exit Sub
    
ErrorHandler:
    ' Call the global error handling procedure
    Call HandleError
End Sub

Public Sub ReplaceHeaderValue()

    Dim cell As Range
    Dim dataArr As Variant
    Dim firstRow As Long
    Dim headerRow As Long
    Dim i As Long, j As Long
    Dim lastColumn As Long
    Dim searchRange As Range
    Dim targetColumn As Range
    Dim targetHeader As String
    
    On Error GoTo ErrorHandler
    
    targetHeader = "Data Snapshot Time (UTC)"
    
    ' Set the search range as the entire worksheet
    Set searchRange = ActiveSheet.UsedRange
    
    ' Find the last used column in the search range
    lastColumn = searchRange.Columns(searchRange.Columns.count).Column
    
    ' Loop through each filled column to find the target header
    For headerRow = 1 To searchRange.Rows.count
        For Each cell In searchRange.Range(searchRange.Cells(headerRow, 1), searchRange.Cells(headerRow, lastColumn))
            If StrComp(cell.Value, targetHeader, vbTextCompare) = 0 Then
                ' Set the target column and row and exit the loop
                Set targetColumn = searchRange.Columns(cell.Column)
                firstRow = headerRow
                Exit For
            End If
        Next cell
        
        If Not targetColumn Is Nothing Then
            Exit For
        End If
    Next headerRow
    
    ' Check if the target header is found
    If Not targetColumn Is Nothing Then
        ' Read the data from the target column into an array
        dataArr = targetColumn.Value
        
        ' Loop through the array to replace "#" with "00:00:00"
        For i = 2 To UBound(dataArr, 1)
            If dataArr(i, 1) = "#" Then
                dataArr(i, 1) = "00:00:00"
            End If
        Next i
        
        ' Update the values in the target column with the modified array
        targetColumn.Value = dataArr
        
        ' Select the entire column
        targetColumn.Select
    End If

    Exit Sub
    
ErrorHandler:
    ' Call the global error handling procedure
    Call HandleError
End Sub






------------------
ThisWorkbook.cls

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' Procedure purpose:  To enable SAP Analysis plug-in, reconnect/refresh all data sources to BW system
Public Sub Workbook_open()
    vFlag = 1
    
    Dim StartRow As Integer
    Dim addin As COMAddIn
    Dim cell As Range
    Dim cofCom As Object
    Dim firstRow As Long
    Dim headerRow As Long
    Dim lastColumn As Long
    Dim result As Variant
    Dim rng As Range
    Dim searchRange As Range
    Dim targetColumn As Range
    Dim targetHeader As String
    Dim targetHeaderCell As Range
    Dim wb As Workbook

    On Error GoTo ErrorHandler
    
    Set cofCom = Application.COMAddIns("SapExcelAddIn").Object
    Set wb = ThisWorkbook
    
    If wb.Windows(1).Visible = False Then
        wb.Windows(1).Visible = True
    End If
    
    For Each addin In Application.COMAddIns
        If addin.progID = "SapExcelAddIn" Then
            If addin.Connect = False Then
                addin.Connect = True
            ElseIf addin.Connect = True Then
                addin.Connect = False
                addin.Connect = True
            End If
        End If
    Next

    Call OnStart
    
    Call Reconnect
    
    result = GetLastFilledColumnAndFirstFilledRow()
    StartRow = result(1)
    
    With ThisWorkbook.Sheets("Edit_CA")
        Set rng = ActiveSheet.Range("A1")
        rng.Select
    End With
    
    If (ActiveSheet.Name = "Export_CA") Or (ActiveSheet.Name = "Export_SP") Then
        Dim llastrow As Variant
        llastrow = ActiveSheet.Range(ActiveSheet.Range("A65536").End(XlDirection.xlUp).Address).Row
        Dim llastcol As Variant
        llastcol = Split(ActiveSheet.Cells(1, ActiveSheet.Range("A" & StartRow).End(xlToRight).Column).Address, "$")(1)
        Set rng = ActiveSheet.Range("A" & StartRow + 1, ActiveSheet.Range(llastcol & llastrow))

        Call RemoveHashCharacters(rng)
        
        targetHeader = "Data Snapshot Time (UTC)"

        ' Set the search range as the entire worksheet
        Set searchRange = ActiveSheet.UsedRange
        
        ' Find the last used column in the search range
        lastColumn = searchRange.Columns(searchRange.Columns.count).Column
        
        ' Loop through each filled column to find the target header
        For headerRow = 1 To searchRange.Rows.count
            For Each cell In searchRange.Range(searchRange.Cells(headerRow, 1), searchRange.Cells(headerRow, lastColumn))
                If StrComp(cell.Value, targetHeader, vbTextCompare) = 0 Then
                    ' Set the target column and row and exit the loop
                    Set targetColumn = searchRange.Columns(cell.Column)
                    firstRow = headerRow
                    Exit For
                End If
            Next cell
            
            If Not targetColumn Is Nothing Then
                Exit For
            End If
        Next headerRow
        
        Set targetHeaderCell = targetColumn.Cells(firstRow)
        
        StartRow = "A" & firstRow
        
        llastrow = ActiveSheet.Range(ActiveSheet.Range("A65536").End(XlDirection.xlUp).Address).Row
        llastcol = Split(ActiveSheet.Cells(1, ActiveSheet.Range(StartRow).End(xlToRight).Column).Address, "$")(1)
        Set rng = ActiveSheet.Range(targetHeaderCell.Offset(1, 1), ActiveSheet.Range(llastcol & llastrow))
        Call RestoreLineBreaks(rng)
    End If
    Call Columns_Visibility
    Call ReplaceHeaderValue
    Call OnEnd
    
    vFlag = 0
    
    Exit Sub
    
ErrorHandler:
    ' Call the global error handling procedure
    Call HandleError
    
End Sub

' Procedure purpose:  To refresh data during switch between worksheets
Public Sub Workbook_SheetActivate(ByVal Sh As Object)

    vFlag = 1
    
    Dim InfoBox As VbMsgBoxResult
    Dim StartRow As Integer
    Dim item As Variant, arr
    Dim lDS_Alias As String
    Dim lRefreshDate As Double
    Dim lResult As Long
    Dim lRet As Boolean
    Dim lSeparator As String * 1
    Dim lSeparator_Count As Integer
    Dim llastcol As String
    Dim llastrow As Long
    Dim result As Variant
    Dim rng As Range
    
    On Error GoTo ErrorHandler
    
    result = GetLastFilledColumnAndFirstFilledRow()
    StartRow = result(1)
    
    Const AckTime As Integer = 3
    
    Call OnStart
    
    Dim StartTime As Variant
    StartTime = Timer
    
    AppActivate Application.Caption
    DoEvents
    InfoBox = TimedMsgBox("Refresh on tab """ & Sh.Name & """ in progress..." _
                        & vbCrLf & vbCrLf & "(This window will close in " & AckTime & " seconds)", _
                          "Data refresh on tab """ & Sh.Name & """", , AckTime)
     
    Select Case Sh.Name
    Case "Edit_CA"
        lDS_Alias = "DS_3"
    Case "Export_CA"
        lDS_Alias = "DS_1;DS_2"
    Case "Changelog_CA"
        lDS_Alias = "DS_4"
    Case "Edit_SP"
        lDS_Alias = "DS_5"
    Case "Export_SP"
        lDS_Alias = "DS_6"
    Case "Changelog_SP"
        lDS_Alias = "DS_7"
    Case "DevAccess"
        lDS_Alias = "DS_8"
    End Select
    
    lSeparator = ";"
    lSeparator_Count = Len(lDS_Alias) - Len(Replace(lDS_Alias, lSeparator, vbNullString))
    lRet = True
    
    If lDS_Alias <> vbNullString Then
        If lSeparator_Count > 0 Then
            arr = Split(lDS_Alias, lSeparator)
            For Each item In arr
                lRet = Application.Run("SAPGetProperty", "IsConnected", item) And lRet
            Next item
        Else
            lRet = Application.Run("SAPGetProperty", "IsConnected", lDS_Alias)
        End If
    End If

    If lRet = True Then
        
        lResult = Empty
        If lDS_Alias <> vbNullString Then
            lResult = Application.Run("SAPExecuteCommand", "Restart", lDS_Alias)
        
            Select Case Sh.Name
            Case "Export_CA"
                lDS_Alias = "DS_2"
                lRefreshDate = Application.Run("SAPGetSourceInfo", lDS_Alias, "QueryLastRefreshedAt")
                Call SetPrintLayout(lRefreshDate)
            Case "Export_SP"
                lRefreshDate = Application.Run("SAPGetSourceInfo", lDS_Alias, "QueryLastRefreshedAt")
                Call SetPrintLayout(lRefreshDate)
            Case Else
                lRefreshDate = Application.Run("SAPGetSourceInfo", lDS_Alias, "QueryLastRefreshedAt")
            End Select
        End If
        
        If (Sh.Name = "Edit_CA") Or (Sh.Name = "Edit_SP") Then
            Call DataValidationList
        ElseIf Sh.Name = "Export_CA" Then
        
            llastrow = ActiveSheet.Range(ActiveSheet.Range("A65536").End(XlDirection.xlUp).Address).Row
            llastcol = Split(ActiveSheet.Cells(1, ActiveSheet.Range("A" & StartRow).End(xlToRight).Column).Address, "$")(1)
            Set rng = ActiveSheet.Range("A" & StartRow + 1, ActiveSheet.Range(llastcol & llastrow))
            
            Call RemoveHashCharacters(rng)

            Set rng = ActiveSheet.Range("S" & StartRow + 1, ActiveSheet.Range(llastcol & llastrow))
            
            Call RestoreLineBreaks(rng)
            
        ElseIf Sh.Name = "Export_SP" Then
        
            llastrow = ActiveSheet.Range(ActiveSheet.Range("A65536").End(XlDirection.xlUp).Address).Row
            llastcol = Split(ActiveSheet.Cells(1, ActiveSheet.Range("A" & StartRow).End(xlToRight).Column).Address, "$")(1)
            Set rng = ActiveSheet.Range("A" & StartRow + 1, ActiveSheet.Range(llastcol & llastrow))
            Call RemoveHashCharacters(rng)
            
        End If
        
        lResult = Application.Run("SAPSetRefreshBehaviour", "On")
        
        Call Columns_Visibility
        
        AppActivate Application.Caption
        DoEvents
        Dim EndTime As Variant
        EndTime = Timer
        InfoBox = TimedMsgBox("Data refreshed in : " & Format((EndTime - StartTime) / 86400, "hh:mm:ss") & " [hh:mm:ss]" _
                            & vbCrLf & vbCrLf & "(This window will close in " & AckTime & " seconds)", _
                              "Data refresh on tab """ & Sh.Name & """", , AckTime)
    Else
        AppActivate Application.Caption
        DoEvents
        lResult = Application.Run("SAPLogOff", "True")
        lResult = Application.Run("SAPExecuteCommand", "Refresh", "All")
    End If
    
    Call ReplaceHeaderValue
    Call OnEnd
    
    vFlag = 0
    
    Exit Sub
    
ErrorHandler:
    ' Call the global error handling procedure
    Call HandleError
End Sub

' Procedure purpose:  To lock edition on export tabs and to validate data in input-ready fields _
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal target As Range)
    (check max. field length, split values in cells per columns, replace line breaks with "&&")

    Dim CondOpt As String
    Dim CurLenLim As Integer
    Dim Mssg As String
    Dim NewCellValue As String
    Dim NewCellValue2 As String
    Dim StartRow As Integer
    Dim ValidatedCells As Range
    Dim ValidatedCells_2 As Range
    Dim cell As Range
    Dim llastrow As Long
    Dim LastCol As String
    Dim result As Variant
    Dim x As Range
    
    On Error GoTo ErrorHandler
    
    Const StringLenLim As Integer = 250
    Const LineSeparator As String = "&&"
    Const Col_1 As Integer = 19
    Const Col_2 As Integer = 22
        
    If Sh.Name Like "Export*" Or Sh.Name Like "Changelog*" Or Sh.Name = "DevAccess" Then
        If vFlag = 0 Then
            Call OnStart
            Set x = ActiveSheet.UsedRange
            Set ValidatedCells = Intersect(target, Range(x.Address))
            If Not ValidatedCells Is Nothing Then
                For Each cell In ValidatedCells
                    result = MsgBox("Attempted to change the value of the cell: " & target.Address, vbExclamation, "Warning")
                    If result = vbOK Then
                        Application.Undo
                        Call OnEnd
                        Exit Sub
                    End If
                Next cell
            End If
            Call OnEnd
        End If
    End If
    
    If Sh.Name Like "Edit*" Then
        If vFlag = 0 Then
            Call OnStart
                
            result = GetLastFilledColumnAndFirstFilledRow()
            LastCol = Split(Cells(1, result(0)).Address, "$")(1)
            StartRow = result(1)
                
            llastrow = ActiveSheet.Range(ActiveSheet.Range("A65536").End(XlDirection.xlUp).Address).Row
            If Right(Sh.Name, 2) = "SP" Then
                Set ValidatedCells = Intersect(target, target.Parent.Range("P" & StartRow & ":Q" & llastrow))
                Set ValidatedCells_2 = Intersect(target, Union(target.Parent.Range("A2" & ":O" & llastrow), target.Parent.Range("A" & StartRow & ":" & LastCol & StartRow)))
            ElseIf Right(Sh.Name, 2) = "CA" Then
                Set ValidatedCells = Intersect(target, target.Parent.Range("S" & StartRow & ":T" & llastrow, "V" & StartRow & ":X" & llastrow))
                Set ValidatedCells_2 = Intersect(target, Union(target.Parent.Range("A2" & ":R" & llastrow), target.Parent.Range("A" & StartRow & ":" & LastCol & StartRow)))
            End If

            If Not ValidatedCells_2 Is Nothing Then
                For Each cell In ValidatedCells_2
                    result = MsgBox("Attempted to change the value of the cell: " & target.Address, vbExclamation, "Warning")
                    If result = vbOK Then
                        Application.Undo
                        Call OnEnd
                        Exit Sub
                    End If
                Next cell
            End If
            If Not ValidatedCells Is Nothing Then
                For Each cell In ValidatedCells
                    If cell.Value <> vbNullString Then
                        NewCellValue = Replace(Replace(cell.Value, vbCr, LineSeparator), vbLf, LineSeparator)
                    End If
                    If Len(NewCellValue) > 250 Then
                        NewCellValue2 = Right(NewCellValue, Len(NewCellValue) - StringLenLim)
                    End If
                    If (cell.Column = Col_1 - 1 Or cell.Column = Col_2 - 1) Then
                        CurLenLim = StringLenLim * 2
                        CondOpt = " and split it into 2 columns "
                    Else
                        CurLenLim = StringLenLim
                        CondOpt = Empty
                    End If
                            
                    Mssg = "The information" & _
                           " inserted in cell " & cell.Address & _
                           " exceeds accepted field length (250 characters) by " & _
                           Len(NewCellValue) - StringLenLim & " characters."
                            
                    If (Len(NewCellValue) > StringLenLim And Len(NewCellValue2) <= StringLenLim And (cell.Column = Col_1 Or cell.Column = Col_2)) Then
                        result = MsgBox(Mssg & _
                                        vbCrLf & vbCrLf & _
                                        "Split it into the next cell (Ok) or undo (Cancel)?", _
                                        vbQuestion + vbOKCancel)
                        If result = vbOK Then
                            cell.Offset(, 1).Value = NewCellValue2
                            NewCellValue = Left(NewCellValue, StringLenLim)
                            cell.Value = NewCellValue
                        Else
                            Application.Undo
                            Call OnEnd
                            Exit Sub
                        End If
                        Exit Sub
                    ElseIf (Len(NewCellValue) > StringLenLim And Len(NewCellValue2) <= StringLenLim And Not (cell.Column = Col_1 And cell.Column = Col_2)) Then
                        result = MsgBox(Mssg & _
                                        vbCrLf & vbCrLf & _
                                        "Trim to " & CurLenLim & CondOpt & " (Ok) or undo (Cancel)?", _
                                        vbQuestion + vbOKCancel)
                        If result = vbOK Then
                            NewCellValue = Left(NewCellValue, StringLenLim)
                            cell.Value = NewCellValue
                        Else
                            Application.Undo
                            Call OnEnd
                            Exit Sub
                        End If
                        Exit Sub
                    ElseIf (Len(NewCellValue) > StringLenLim And Len(NewCellValue2) > StringLenLim And (cell.Column = Col_1 Or cell.Column = Col_2)) Then
                        result = MsgBox(Mssg & _
                                        vbCrLf & vbCrLf & _
                                        "Trim to " & CurLenLim & CondOpt & " (Ok) or undo (Cancel)?", _
                                        vbQuestion + vbOKCancel)
                        If result = vbOK Then
                            cell.Offset(, 1).Value = Left(NewCellValue2, StringLenLim)
                            NewCellValue = Left(NewCellValue, StringLenLim)
                            cell.Value = NewCellValue
                        Else
                            Application.Undo
                            Call OnEnd
                            Exit Sub
                        End If
                        Exit Sub
                    ElseIf (Len(NewCellValue) > StringLenLim And Len(NewCellValue2) > StringLenLim And Not (cell.Column = Col_1 And cell.Column = Col_2)) Then
                        result = MsgBox(Mssg & _
                                        vbCrLf & vbCrLf & _
                                        "Trim to " & CurLenLim & CondOpt & " (Ok) or undo (Cancel)?", _
                                        vbQuestion + vbOKCancel)
                        If result = vbOK Then
                            NewCellValue = Left(NewCellValue, StringLenLim)
                            cell.Value = NewCellValue
                        Else
                            Application.Undo
                            Call OnEnd
                            Exit Sub
                        End If
                        Exit Sub
                    End If
                            
                Next cell
            End If
            Call OnEnd
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    ' Call the global error handling procedure
    Call HandleError
End Sub




