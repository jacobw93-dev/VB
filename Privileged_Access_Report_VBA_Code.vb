ThisWorkBook

Public Sub Workbook_open()

    Dim wb          As Workbook
    Dim ws          As Worksheet
    
    Dim lResult     As Long
    Dim addin       As COMAddIn
    Dim rng As Range
    
    Call OnStart
    
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
    
    lResult = Application.Run("SAPSetRefreshBehaviour", "Off")
    lResult = Application.Run("SAPLogOff", "True")
    lResult = Application.Run("SAPExecuteCommand", "Refresh", "All")
    lResult = Application.Run("SAPSetRefreshBehaviour", "On")
    
    With ThisWorkbook.Sheets("Edit")
        Set rng = Range("A1")
        rng.Select
    End With
    
    Call DataValidationList
    
    Call OnEnd
    
End Sub

Public Sub Workbook_SheetActivate(ByVal Sh As Object)
    
    Dim lResult     As Long, lRet As Boolean
    Dim lRefreshDate As Double
    Dim lDS_Alias   As String
    Dim lSeparator  As String * 1
    Dim lSeparator_Count As Integer
    Dim item        As Variant, arr, c As Collection
    Dim rng         As Range
    
    Dim AckTime     As Integer, InfoBox As Object
    
    Call OnStart
    
    Set InfoBox = CreateObject("WScript.Shell")
    AckTime = 3
    AppActivate Application.Caption
    DoEvents
    Select Case InfoBox.Popup("Refresh On tab """ & Sh.Name & """ in progress..." _
         & vbCrLf & vbCrLf & "(This window will close in " & AckTime & " seconds)", _
           AckTime, "Data refresh", 0)
    End Select
     
    Select Case Sh.Name
        Case "Edit"
            lDS_Alias = "DS_3"
        Case "Export"
            lDS_Alias = "DS_1;DS_2"
        Case "Changelog"
            lDS_Alias = "DS_4"
        Case "Sensitive Profiles"
            lDS_Alias = "DS_5"
    End Select
    
    lSeparator = ";"
    lSeparator_Count = Len(lDS_Alias) - Len(Replace(lDS_Alias, lSeparator, ""))
    lRet = True
    
    If lSeparator_Count > 0 Then
        arr = Split(lDS_Alias, lSeparator)
        For Each item In arr
            lRet = Application.Run("SAPGetProperty", "IsConnected", item) And lRet
        Next item
    Else
        lRet = Application.Run("SAPGetProperty", "IsConnected", lDS_Alias)
    End If
    

    If lRet = True Then
        
        StartTime = Timer
        
        lResult = Application.Run("SAPSetRefreshBehaviour", "Off")
        lResult = Application.Run("SAPExecuteCommand", "Refresh", lDS_Alias)
        lResult = Application.Run("SAPExecuteCommand", "RefreshData", lDS_Alias)
        lResult = Application.Run("SAPExecuteCommand", "Restart", lDS_Alias)
        
        Select Case Sh.Name
            Case "Export"
                lDS_Alias = "DS_2"
                lRefreshDate = Application.Run("SAPGetSourceInfo", lDS_Alias, "QueryLastRefreshedAt")
                Set rng = ThisWorkbook.Sheets("Export").Range("A1", Range("A1").End(xlDown).End(xlToRight).End(xlDown).End(xlToRight).End(xlDown))
                '                Debug.Print (rng.Address)
                ThisWorkbook.Sheets("Export").PageSetup.PrintArea = rng.Address
            Case Else
                lRefreshDate = Application.Run("SAPGetSourceInfo", lDS_Alias, "QueryLastRefreshedAt")
        End Select
        
        For Each ws In Application.ActiveWorkbook.Worksheets
            ws.PageSetup.CenterHeader = "QueryLastRefreshedAt: " & Format(lRefreshDate, "dddd, mmmm d, yyyy h:mm:ss")
        Next
        
        If Sh.Name = "Edit" Then
            Call DataValidationList
            Call LockSheets
        ElseIf Sh.Name = "Export" Then
            Call Remove_Hash_Characters
            Call RestoreLineBreaks
        End If
        
        EndTime = Timer
        lResult = Application.Run("SAPSetRefreshBehaviour", "On")
        AppActivate Application.Caption
        DoEvents
        Select Case InfoBox.Popup("Data refreshed in : " & Format((EndTime - StartTime) / 86400, "hh:mm:ss") & " [hh:mm:ss]" _
             & vbCrLf & vbCrLf & "(This window will close in " & AckTime & " seconds)", _
               AckTime, "Data refresh", 0)
        End Select
    
    Else
        AppActivate Application.Caption
        DoEvents
        Select Case InfoBox.Popup("You are not connected to the system" _
             & vbCrLf & vbCrLf & "(This window will close in " & AckTime & " seconds)", _
               AckTime, "Connection status", 0)
        End Select
        lResult = Application.Run("SAPLogOff", "True")
        lResult = Application.Run("SAPExecuteCommand", "Refresh", "All")
    End If
    
    Call OnEnd
    ActiveSheet.Calculate

End Sub

-------------
Sheet3(Edit)

Private Sub CommandButton1_Click()
    
    Dim lResult     As Long, lRet As Boolean
    Dim range_1     As Range
    Dim i           As Integer, ds As String, ds_concat As String
    
    Call OnStart
    
    lRet = True
    
    For i = 1 To 5
        ds = "DS_" & i
        lResult = Application.Run("SAPGetProperty", "IsConnected", ds)
        lRet = lRet And lResult
        If lResult = False Then
            If ds_concat <> "" Then
                ds_concat = ds_concat & ", " & ds
            Else
                ds_concat = ds
            End If
        End If
    Next i
    
    If ds_concat <> "" Then
        MsgBox "Data Sources: """ & ds_concat & """ are inactive"
    End If
    
    If lResult = False Then
        
        ThisWorkbook.Sheets("Edit").Activate
        ActiveSheet.Range("A1").Activate
        
        lResult = Application.Run("SAPLogOff", "True")
        lResult = Application.Run("SAPSetRefreshBehaviour", "Off")
        lResult = Application.Run("SAPExecuteCommand", "Refresh", "ALL")
        
        Call DataValidationList
        
        lResult = Application.Run("SAPSetRefreshBehaviour", "On")
        
    Else
        MsgBox "You are already connected to the system"
    End If
    
    Call OnEnd
    
End Sub

Private Sub CommandButton2_Click()
    
    Dim lResult     As Long
    Dim lRefreshDate As Double
    Dim StartTime   As Double
    Dim EndTime     As Double
    Dim wb          As Workbook
    Dim AckTime     As Integer, InfoBox As Object
    
    Set wb = ThisWorkbook
    
    Call OnStart
    
    lResult = Application.Run("SAPGetProperty", "IsConnected", "DS_2")
    If lResult = True Then
        
        StartTime = Timer
        lResult = Application.Run("SAPSetRefreshBehaviour", "Off")
        lResult = Application.Run("SAPExecuteCommand", "PlanDataSave")
        lResult = Application.Run("SAPExecuteCommand", "Restart", "ALL")
        
        Call DataValidationList
        Call Remove_Hash_Characters
        
        EndTime = Timer
        lResult = Application.Run("SAPSetRefreshBehaviour", "On")
        wb.Sheets("Edit").Range("E1").Value = Format((EndTime - StartTime) / 86400, "hh:mm:ss") & " [hh:mm:ss]"
        
        Set InfoBox = CreateObject("WScript.Shell")
        AckTime = 3
        Select Case InfoBox.Popup("Data saved in : " & Format((EndTime - StartTime) / 86400, "hh:mm:ss") & " [hh:mm:ss]" _
             & vbCrLf & vbCrLf & "(This window will close in " & AckTime & " seconds)", _
               AckTime, "Data saved", 0)
        End Select
        
    Else
        MsgBox "You are not connected to the system"
    End If
    
    Call OnEnd

End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    
    Dim ValidatedCells As Range
    Dim Cell        As Range
    
    Set ValidatedCells = Intersect(Target, Target.Parent.Range("Q:R,T:V"))
    If Not ValidatedCells Is Nothing Then
        
        For Each Cell In ValidatedCells
            If Not Len(Cell.Value) <= 250 Then
                MsgBox "The value" & _
                       " inserted in cell " & Cell.Address & _
                       " exceeds accepted field length by " & _
                       Len(Cell.Value) - 250 & " characters." & _
                       vbCrLf & vbCrLf & "Undo!", vb
                Application.Undo
                Exit Sub
            End If
        Next Cell
        
        For Each Cell In ValidatedCells.Cells
            If Len(Replace(Replace(Cell.Value, vbCr, ""), vbLf, "")) <> Len(Cell.Value) Then
                Cell.Value = Replace(Replace(Cell.Value, vbCr, "&&"), vbLf, "&&")
            End If
        Next Cell
    End If
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Excel.Range)
    On Error GoTo 0
    With Cells(Windows(1).ScrollRow, Windows(1).ScrollColumn)
        CommandButton1.Top = .Top + 0
        CommandButton1.Left = .Left + 0
        CommandButton2.Top = .Top + 0
        CommandButton2.Left = .Left + 80
    End With
End Sub

-------------
Module1

Private Function SheetProtected(TargetSheet As Worksheet) As Boolean
     'Function purpose:  To evaluate if a worksheet is protected
     
    If TargetSheet.ProtectContents = True Then
        SheetProtected = True
    Else
        SheetProtected = False
    End If
     
End Function

Public Sub UnlockSheets()
    Dim wb          As Workbook
    Dim ws          As Worksheet
    
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

    
End Sub

Public Sub LockSheets()
    Dim rng As Range
    Dim lResult     As Long
    
    ThisWorkbook.Activate
    
    With ThisWorkbook.Sheets("Edit")
        Set rng = Range("A1")
        rng.Select
    End With
    
    With ThisWorkbook.Sheets("Edit").Rows(1)
        If Selection.Locked = False Then
            Selection.Locked = True
        End If
        If Selection.FormulaHidden = False Then
            Selection.FormulaHidden = True
        End If
    End With
    If ActiveSheet.ProtectContents = False Then
        ThisWorkbook.Sheets("Edit").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    End If
    
End Sub
Public Sub OnStart()
    
    Dim wb          As Workbook
    Dim ws          As Worksheet
    
    Set wb = ThisWorkbook
    
    ThisWorkbook.Activate
    
    ActiveSheet.EnableCalculation = False
    Application.AskToUpdateLinks = False
    Application.Calculation = xlCalculationManual ' xlAutomatic
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.StatusBar = False
    
    Call UnlockSheets
    
End Sub

Public Sub OnEnd()
    
    Dim wb          As Workbook
    Dim ws          As Worksheet
    
    Set wb = ThisWorkbook
    
    ThisWorkbook.Activate
       
    Call LockSheets
    
    ActiveSheet.EnableCalculation = True
    Application.AskToUpdateLinks = True
    Application.Calculation = xlAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.StatusBar = True
    
End Sub

Public Sub DataValidationList()
    Dim rng As Range
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets("Edit")
    Set rng = ws.Range("S3", ws.Range("S3").End(xlDown))
    
    ws.Activate
    rng.Select
    
    If SheetProtected(ws) Then
       Call UnlockSheets
    End If
    
    With rng.Validation
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
    
    Set rng = Sheets("Edit").Range("S2")
    
    With rng
        .ClearComments
        If .Comment Is Nothing Then
            .AddCommentThreaded ( _
                "0 - Not approved;" & vbLf & "1 - Pending approval;" & vbLf & "2 - Approved;" _
                )
        End If
    End With
    
End Sub

Public Sub Remove_Hash_Characters()
    
    Dim rng         As Range
    Dim Cell           As Range
    
    Set rng = ActiveSheet.Range("A6", Range("A6").End(xlDown).End(xlToRight))
    
    For Each Cell In rng.Cells
        If Cell.Value = "#" Then
            Cell.Value = ""
        End If
    Next Cell
    
End Sub

Public Sub RestoreLineBreaks()
    
    Dim rng         As Range
    Dim Cell           As Range
    
    Set rng = ActiveSheet.Range("Q6", Range("Q6").SpecialCells(xlLastCell))
    
    For Each Cell In rng.Cells
        If Len(Replace(Cell.Value, "&&", "")) <> Len(Cell.Value) Then
            Cell.Value = Replace(Cell.Value, "&&", vbCrLf)
        End If
    Next Cell
    
End Sub