Privileged_Access_Report_VBA_Code.vb

------------------
My_Module.bas

Attribute VB_Name = "My_Module"
Option Explicit
    Global vFlag As Integer

Public Declare PtrSafe Function CustomTimeOffMsgBox Lib "user32" Alias "MessageBoxTimeoutA" ( _
            ByVal xHwnd As LongPtr, _
            ByVal xText As String, _
            ByVal xCaption As String, _
            ByVal xMsgBoxStyle As VbMsgBoxStyle, _
            ByVal xwlange As Long, _
            ByVal xTimeOut As Long) _
    As Long
    


Public Function TimedMsgBox( _
        Prompt As String, _
        Optional Title As String = "Pop-up message", _
        Optional Buttons As VbMsgBoxStyle = vbOKOnly, _
        Optional Timeout As Long = 5000) _
    As VbMsgBoxResult
    ' Function purpose:  To create custom MsgBox with autoclose option
    
    TimedMsgBox = CustomTimeOffMsgBox(0&, Prompt, Title, Buttons, 0, Timeout * 1000)
    
End Function

Public Function SheetProtected(TargetSheet As Worksheet) As Boolean
    ' Function purpose:  To evaluate if a worksheet is protected
    
    If TargetSheet.ProtectContents = True Then
        SheetProtected = True
    Else
        SheetProtected = False
    End If
    
End Function

Public Sub UnlockSheets()
' Procedure purpose:  To unlock all worksheets in this workbook
    
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

Public Sub Reconnect()
' Procedure purpose:  To reconnect with the SAP data source
    Dim lResult     As Long
    
    lResult = Application.Run("SAPSetRefreshBehaviour", "Off")
    lResult = Application.Run("SAPLogOff", "True")
    lResult = Application.Run("SAPExecuteCommand", "Refresh", "All")
    If ActiveSheet.Name = "Edit_SP" Then
        Call DataValidationList
    End If
    lResult = Application.Run("SAPSetRefreshBehaviour", "On")
    
    End Sub
Public Sub OnStart()
' Procedure purpose:  To disable immediate calculations, screen updates, events, messages
    
    Dim wb          As Workbook
    Dim ws          As Worksheet
    
    Set wb = ThisWorkbook
    
    ThisWorkbook.Activate
    
    ActiveSheet.EnableCalculation = False
    Application.AskToUpdateLinks = False
    Application.Calculation = xlCalculationManual        ' xlAutomatic
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    Call UnlockSheets
    
End Sub

Public Sub OnEnd()
' Procedure purpose:  To enable immediate calculations, screen updates, events, messages
    
    Dim wb          As Workbook
    Dim ws          As Worksheet
    
    Set wb = ThisWorkbook
    
    ThisWorkbook.Activate
    
    ' If Left(ActiveSheet.Name, 6) <> "Export" Or ActiveSheet.Name <> "DevAccess" Then
        ActiveSheet.EnableCalculation = True
        Application.AskToUpdateLinks = True
        Application.Calculation = xlAutomatic
        Application.DisplayAlerts = True
        Application.EnableEvents = True
        Application.ScreenUpdating = True
    ' End If
    
End Sub

Public Sub DataValidationList()
' Procedure purpose:  To add data validation list on certain range (Approval flag)
    
    Dim rng         As Range
    Dim ws          As Worksheet
    Dim llastcol    As String
    Dim llastrow    As Long
    
    Const Col_Approval As String = "U"
    
    Set ws = ThisWorkbook.ActiveSheet
    llastrow = Range(Range("A65536").End(XlDirection.xlUp).Address).Row
    llastcol = Split(Cells(1, Range("A5").End(xlToRight).Column).Address, "$")(1)
    Set rng = ActiveSheet.Range("A3", Range(llastcol & llastrow))
    
    rng.WrapText = True
    
    Set rng = ActiveSheet.Range(Col_Approval & 3, Range(Col_Approval & llastrow))
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
    
    Call Comments
    
End Sub

Public Sub Comments()
' Procedure purpose:  To add comment on "Approval Flag" header

    Dim rng         As Range
    Const Col_Approval As String = "U"
    
    Set rng = Sheets("Edit_CA").Range(Col_Approval & 2)
    
    With rng
        .ClearComments
        If .Comment Is Nothing Then
            .AddCommentThreaded ( _
                                "0 - Not approved;" & vbLf & "1 - Pending approval;" & vbLf & "2 - Approved;" _
                                )
        End If
    End With
End Sub

Public Sub Alignment()
' Procedure purpose:  To align columns

    Dim llastcol    As String
    Dim llastrow    As Long
    Dim StartRow    As String
    Dim rng         As Range
    
    Const Col_1 As String = "S"
    
    If ActiveSheet.Name = "Export_CA" Then
        StartRow = "A7"
    ElseIf ActiveSheet.Name = "Export_SP" Then
        StartRow = "A1"
    Else
        StartRow = "A1"
    End If
    
    
    llastrow = Range(Range("A65536").End(XlDirection.xlUp).Address).Row
    llastcol = Split(Cells(1, Range(StartRow).End(xlToRight).Column).Address, "$")(1)
    Set rng = ActiveSheet.Range(Col_1 & 1, Range(llastcol & llastrow))
    
    With rng
        .WrapText = True
        .HorizontalAlignment = xlLeft
        .ReadingOrder = xlContext
        .EntireColumn.ColumnWidth = 30
        .EntireRow.AutoFit
    End With
    
End Sub

Public Sub Columns_Visibility()
' Procedure purpose:  To change columns visiblity

    If Right(ActiveSheet.Name, 2) = "CA" Then
            Call Alignment
    End If
    
    If ActiveSheet.Name = "Export_CA" Then
        ActiveSheet.Range("T:T,V:V").EntireColumn.Hidden = True
    ElseIf ActiveSheet.Name = "Export_SP" Then
        ActiveSheet.Range("Q:Z").EntireColumn.Hidden = True
    ElseIf ActiveSheet.Name = "Changelog_SP" Then
        ActiveSheet.Range("U:Z").EntireColumn.Hidden = True
    End If
    
End Sub

Public Sub RemoveHashCharacters(rng As Range)
' Procedure purpose:  To replace "#" characters with blank value in particular range on "Export" tab

    Dim Cell           As Range
    Dim llastrow    As Long
    Dim DataRange   As Variant
    Dim Irow        As Long, rowcnt As Integer
    Dim Icol        As Integer, colcnt As Integer
    Dim Cell_1      As String
    
'    llastrow = Range(Range("A65536").End(XlDirection.xlUp).Address).Row
'    Set rng = ActiveSheet.Range("A6", Range("A" & llastrow).End(xlToRight))
    
    DataRange = rng.Value
    rowcnt = rng.Rows.count
    colcnt = rng.Columns.count
    
    For Irow = 1 To rowcnt
        For Icol = 1 To colcnt
            Cell_1 = DataRange(Irow, Icol)
            If Cell_1 <> "" Then
                If Cell_1 = "#" Then
                    Cell_1 = ""
                End If
                DataRange(Irow, Icol) = Cell_1
            End If
        Next Icol
    Next Irow
    rng.Value = DataRange
    
    rng.WrapText = True
    rng.Columns.EntireColumn.AutoFit
    rng.Rows.EntireRow.AutoFit
    
End Sub

Public Sub RestoreLineBreaks(rng As Range)
' Procedure purpose:  To replace "&&" with line breaks (VbCrLf) in certain range on the "Export" tab _
and to merge cells that were splitted due to max. field length (250)
    
    Dim Cell           As Range
    Dim llastcol    As String
    Dim llastrow    As Long
    Dim DataRange   As Variant
    Dim Irow        As Long, rowcnt As Integer
    Dim Icol        As Integer, colcnt As Integer
    Dim Cell_1      As String, Cell_2 As String
    
    Const LineSeparator As String = "&&"

    Set rng = rng
    rowcnt = rng.Rows.count
    colcnt = rng.Columns.count
    
    DataRange = rng.Value
    
    For Irow = 1 To rowcnt
        For Icol = 1 To colcnt
            Cell_1 = DataRange(Irow, Icol)
            If Cell_1 <> "" Then
                If (Icol = 1 Or Icol = 3) Then
                    Debug.Print (DataRange(Irow, Icol + 1))
                    Cell_2 = DataRange(Irow, Icol + 1)
                    Cell_1 = Cell_1 & Cell_2
                End If
                
                If (Icol = 2 Or Icol = 4) Then
                    Cell_1 = Empty
                End If
                If Len(Replace(Cell_1, LineSeparator, "")) <> Len(Cell_1) Then
                    Cell_1 = Replace(Cell_1, LineSeparator, vbCrLf)
                End If
                DataRange(Irow, Icol) = Cell_1
            End If
        Next Icol
    Next Irow
    rng.Value = DataRange
    
End Sub

Public Sub SetPrintLayout(lRefreshDate As Double)
' Procedure purpose:  To set print area, header information (Query last refresh date & time) on the "Export" tab
    
    Dim llastrow    As Long
    Dim rng         As Range
    Dim ws          As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ws.Activate
    
    llastrow = Range(ws.Range("A65536").End(XlDirection.xlUp).Address).Row
    Set rng = ws.Range("A1", Range("A" & llastrow).End(xlToRight))
    ws.PageSetup.PrintArea = rng.Address
    ws.PageSetup.CenterHeader = "QueryLastRefreshedAt: " & Format(lRefreshDate, "dddd, mmmm d, yyyy h:mm:ss")
    
End Sub



------------------
Privileged_Access_Report_VBA_Code.vb

------------------
ExportVBA.bas

Attribute VB_Name = "ExportVBA"
' Excel macro to export all VBA source code in this project to text files for proper source control versioning
' Requires enabling the Excel setting in Options/Trust Center/Trust Center Settings/Macro Settings/Trust access to the VBA project object model
Public Sub ExportVisualBasicCode()
    Const Module = 1
    Const ClassModule = 2
    Const Form = 3
    Const Document = 100
    Const Padding = 24
    
    Dim VBComponent As Object
    Dim count As Integer
    Dim path As String
    Dim directory As String
    Dim extension As String
    Dim fso As New FileSystemObject
    
    directory = ActiveWorkbook.path & "\VisualBasic"
    count = 0
    
    If Not fso.FolderExists(directory) Then
        Call fso.CreateFolder(directory)
    End If
    Set fso = Nothing
    
    For Each VBComponent In ActiveWorkbook.VBProject.VBComponents
        Select Case VBComponent.Type
            Case ClassModule, Document
                extension = ".cls"
            Case Form
                extension = ".frm"
            Case Module
                extension = ".bas"
            Case Else
                extension = ".txt"
        End Select
            
                
        On Error Resume Next
        Err.Clear
        
        path = directory & "\" & VBComponent.Name & extension
        Call VBComponent.Export(path)
        
        If Err.Number <> 0 Then
            Call MsgBox("Failed to export " & VBComponent.Name & " to " & path, vbCritical)
        Else
            count = count + 1
            Debug.Print "Exported " & Left$(VBComponent.Name & ":" & Space(Padding), Padding) & path
        End If

        On Error GoTo 0
    Next
    
    Application.StatusBar = "Successfully exported " & CStr(count) & " VBA files to " & directory
    Application.OnTime Now + TimeSerial(0, 0, 10), "ClearStatusBar"
End Sub


------------------
My_Module.bas

Attribute VB_Name = "My_Module"
Option Explicit
    Global vFlag As Integer

Public Declare PtrSafe Function CustomTimeOffMsgBox Lib "user32" Alias "MessageBoxTimeoutA" ( _
            ByVal xHwnd As LongPtr, _
            ByVal xText As String, _
            ByVal xCaption As String, _
            ByVal xMsgBoxStyle As VbMsgBoxStyle, _
            ByVal xwlange As Long, _
            ByVal xTimeOut As Long) _
    As Long
    


Public Function TimedMsgBox( _
        Prompt As String, _
        Optional Title As String = "Pop-up message", _
        Optional Buttons As VbMsgBoxStyle = vbOKOnly, _
        Optional Timeout As Long = 5000) _
    As VbMsgBoxResult
    ' Function purpose:  To create custom MsgBox with autoclose option
    
    TimedMsgBox = CustomTimeOffMsgBox(0&, Prompt, Title, Buttons, 0, Timeout * 1000)
    
End Function

Public Function SheetProtected(TargetSheet As Worksheet) As Boolean
    ' Function purpose:  To evaluate if a worksheet is protected
    
    If TargetSheet.ProtectContents = True Then
        SheetProtected = True
    Else
        SheetProtected = False
    End If
    
End Function

Public Sub UnlockSheets()
' Procedure purpose:  To unlock all worksheets in this workbook
    
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

Public Sub Reconnect()
' Procedure purpose:  To reconnect with the SAP data source
    Dim lResult     As Long
    
    lResult = Application.Run("SAPSetRefreshBehaviour", "Off")
    lResult = Application.Run("SAPLogOff", "True")
    lResult = Application.Run("SAPExecuteCommand", "Refresh", "All")
    If ActiveSheet.Name = "Edit_SP" Then
        Call DataValidationList
    End If
    lResult = Application.Run("SAPSetRefreshBehaviour", "On")
    
    End Sub
Public Sub OnStart()
' Procedure purpose:  To disable immediate calculations, screen updates, events, messages
    
    Dim wb          As Workbook
    Dim ws          As Worksheet
    
    Set wb = ThisWorkbook
    
    ThisWorkbook.Activate
    
    ActiveSheet.EnableCalculation = False
    Application.AskToUpdateLinks = False
    Application.Calculation = xlCalculationManual        ' xlAutomatic
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    Call UnlockSheets
    
End Sub

Public Sub OnEnd()
' Procedure purpose:  To enable immediate calculations, screen updates, events, messages
    
    Dim wb          As Workbook
    Dim ws          As Worksheet
    
    Set wb = ThisWorkbook
    
    ThisWorkbook.Activate
    
    ' If Left(ActiveSheet.Name, 6) <> "Export" Or ActiveSheet.Name <> "DevAccess" Then
        ActiveSheet.EnableCalculation = True
        Application.AskToUpdateLinks = True
        Application.Calculation = xlAutomatic
        Application.DisplayAlerts = True
        Application.EnableEvents = True
        Application.ScreenUpdating = True
    ' End If
    
End Sub

Public Sub DataValidationList()
' Procedure purpose:  To add data validation list on certain range (Approval flag)
    
    Dim rng         As Range
    Dim ws          As Worksheet
    Dim llastcol    As String
    Dim llastrow    As Long
    
    Const Col_Approval As String = "U"
    
    Set ws = ThisWorkbook.ActiveSheet
    llastrow = Range(Range("A65536").End(XlDirection.xlUp).Address).Row
    llastcol = Split(Cells(1, Range("A5").End(xlToRight).Column).Address, "$")(1)
    Set rng = ActiveSheet.Range("A3", Range(llastcol & llastrow))
    
    rng.WrapText = True
    
    Set rng = ActiveSheet.Range(Col_Approval & 3, Range(Col_Approval & llastrow))
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
    
    Call Comments
    
End Sub

Public Sub Comments()
' Procedure purpose:  To add comment on "Approval Flag" header

    Dim rng         As Range
    Const Col_Approval As String = "U"
    
    Set rng = Sheets("Edit_CA").Range(Col_Approval & 2)
    
    With rng
        .ClearComments
        If .Comment Is Nothing Then
            .AddCommentThreaded ( _
                                "0 - Not approved;" & vbLf & "1 - Pending approval;" & vbLf & "2 - Approved;" _
                                )
        End If
    End With
End Sub

Public Sub Alignment()
' Procedure purpose:  To align columns

    Dim llastcol    As String
    Dim llastrow    As Long
    Dim StartRow    As String
    Dim rng         As Range
    
    Const Col_1 As String = "S"
    
    If ActiveSheet.Name = "Export_CA" Then
        StartRow = "A7"
    ElseIf ActiveSheet.Name = "Export_SP" Then
        StartRow = "A1"
    Else
        StartRow = "A1"
    End If
    
    
    llastrow = Range(Range("A65536").End(XlDirection.xlUp).Address).Row
    llastcol = Split(Cells(1, Range(StartRow).End(xlToRight).Column).Address, "$")(1)
    Set rng = ActiveSheet.Range(Col_1 & 1, Range(llastcol & llastrow))
    
    With rng
        .WrapText = True
        .HorizontalAlignment = xlLeft
        .ReadingOrder = xlContext
        .EntireColumn.ColumnWidth = 30
        .EntireRow.AutoFit
    End With
    
End Sub

Public Sub Columns_Visibility()
' Procedure purpose:  To change columns visiblity

    If Right(ActiveSheet.Name, 2) = "CA" Then
            Call Alignment
    End If
    
    If ActiveSheet.Name = "Export_CA" Then
        ActiveSheet.Range("T:T,V:V").EntireColumn.Hidden = True
    ElseIf ActiveSheet.Name = "Export_SP" Then
        ActiveSheet.Range("Q:Z").EntireColumn.Hidden = True
    ElseIf ActiveSheet.Name = "Changelog_SP" Then
        ActiveSheet.Range("U:Z").EntireColumn.Hidden = True
    End If
    
End Sub

Public Sub RemoveHashCharacters(rng As Range)
' Procedure purpose:  To replace "#" characters with blank value in particular range on "Export" tab

    Dim Cell           As Range
    Dim llastrow    As Long
    Dim DataRange   As Variant
    Dim Irow        As Long, rowcnt As Integer
    Dim Icol        As Integer, colcnt As Integer
    Dim Cell_1      As String
    
'    llastrow = Range(Range("A65536").End(XlDirection.xlUp).Address).Row
'    Set rng = ActiveSheet.Range("A6", Range("A" & llastrow).End(xlToRight))
    
    DataRange = rng.Value
    rowcnt = rng.Rows.count
    colcnt = rng.Columns.count
    
    For Irow = 1 To rowcnt
        For Icol = 1 To colcnt
            Cell_1 = DataRange(Irow, Icol)
            If Cell_1 <> "" Then
                If Cell_1 = "#" Then
                    Cell_1 = ""
                End If
                DataRange(Irow, Icol) = Cell_1
            End If
        Next Icol
    Next Irow
    rng.Value = DataRange
    
    rng.WrapText = True
    rng.Columns.EntireColumn.AutoFit
    rng.Rows.EntireRow.AutoFit
    
End Sub

Public Sub RestoreLineBreaks(rng As Range)
' Procedure purpose:  To replace "&&" with line breaks (VbCrLf) in certain range on the "Export" tab _
and to merge cells that were splitted due to max. field length (250)
    
    Dim Cell           As Range
    Dim llastcol    As String
    Dim llastrow    As Long
    Dim DataRange   As Variant
    Dim Irow        As Long, rowcnt As Integer
    Dim Icol        As Integer, colcnt As Integer
    Dim Cell_1      As String, Cell_2 As String
    
    Const LineSeparator As String = "&&"

    Set rng = rng
    rowcnt = rng.Rows.count
    colcnt = rng.Columns.count
    
    DataRange = rng.Value
    
    For Irow = 1 To rowcnt
        For Icol = 1 To colcnt
            Cell_1 = DataRange(Irow, Icol)
            If Cell_1 <> "" Then
                If (Icol = 1 Or Icol = 3) Then
                    Debug.Print (DataRange(Irow, Icol + 1))
                    Cell_2 = DataRange(Irow, Icol + 1)
                    Cell_1 = Cell_1 & Cell_2
                End If
                
                If (Icol = 2 Or Icol = 4) Then
                    Cell_1 = Empty
                End If
                If Len(Replace(Cell_1, LineSeparator, "")) <> Len(Cell_1) Then
                    Cell_1 = Replace(Cell_1, LineSeparator, vbCrLf)
                End If
                DataRange(Irow, Icol) = Cell_1
            End If
        Next Icol
    Next Irow
    rng.Value = DataRange
    
End Sub

Public Sub SetPrintLayout(lRefreshDate As Double)
' Procedure purpose:  To set print area, header information (Query last refresh date & time) on the "Export" tab
    
    Dim llastrow    As Long
    Dim rng         As Range
    Dim ws          As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    ws.Activate
    
    llastrow = Range(ws.Range("A65536").End(XlDirection.xlUp).Address).Row
    Set rng = ws.Range("A1", Range("A" & llastrow).End(xlToRight))
    ws.PageSetup.PrintArea = rng.Address
    ws.PageSetup.CenterHeader = "QueryLastRefreshedAt: " & Format(lRefreshDate, "dddd, mmmm d, yyyy h:mm:ss")
    
End Sub





------------------
Sheet1.cls

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


------------------
Sheet2.cls

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


------------------
Sheet3.cls

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    ' Procedure purpose:  To validate data in input-ready fields _
    (check max. field length, split values in cells per columns, replace line breaks with "&&")
    
    Dim ValidatedCells As Range
    Dim Cell        As Range
    Dim Result      As Integer
    Dim NewCellValue As String, NewCellValue2 As String
    Dim rng         As Range
    Dim llastrow    As Long
    
    Const StringLenLim As Integer = 250
    Const LineSeparator As String = "&&"
    Const Col_1 As Integer = 19
    Const Col_2 As Integer = 22
    
    Dim CurLenLim   As Integer
    Dim CondOpt     As String
    Dim Mssg        As String
    
    llastrow = Range(ActiveSheet.Range("A65536").End(XlDirection.xlUp).Address).Row
    
    Set ValidatedCells = Intersect(Target, Target.Parent.Range("S3:T" & llastrow, "V3:X" & llastrow))
    If Not ValidatedCells Is Nothing Then
        For Each Cell In ValidatedCells
            If Cell.Value <> "" Then
                NewCellValue = Replace(Replace(Cell.Value, vbCr, LineSeparator), vbLf, LineSeparator)
            End If
            If Len(NewCellValue) > 250 Then
                NewCellValue2 = Right(NewCellValue, Len(NewCellValue) - StringLenLim)
            End If
            If (Cell.Column = 18 Or Cell.Column = 21) Then
                CurLenLim = StringLenLim * 2
                CondOpt = " and split it into 2 columns "
            Else
                CurLenLim = StringLenLim
                CondOpt = Empty
            End If
            
            Mssg = "The information" & _
                    " inserted in cell " & Cell.Address & _
                    " exceeds accepted field length (250 characters) by " & _
                    Len(NewCellValue) - StringLenLim & " characters."
            
            If (Len(NewCellValue) > StringLenLim And Len(NewCellValue2) <= StringLenLim And (Cell.Column = Col_1 Or Cell.Column = Col_2)) Then
                Result = MsgBox(Mssg & _
                         vbCrLf & vbCrLf & _
                         "Split it into the next cell (Ok) or undo (Cancel)?", _
                         vbQuestion + vbOKCancel)
                If Result = vbOK Then
                    Cell.Offset(, 1).Value = NewCellValue2
                    NewCellValue = Left(NewCellValue, StringLenLim)
                    Cell.Value = NewCellValue
                Else
                    Application.Undo
                    Exit Sub
                End If
                Exit Sub
            ElseIf (Len(NewCellValue) > StringLenLim And Len(NewCellValue2) <= StringLenLim And Not (Cell.Column = Col_1 And Cell.Column = Col_2)) Then
                Result = MsgBox(Mssg & _
                         vbCrLf & vbCrLf & _
                         "Trim to " & CurLenLim & CondOpt & " (Ok) or undo (Cancel)?", _
                         vbQuestion + vbOKCancel)
                If Result = vbOK Then
                    NewCellValue = Left(NewCellValue, StringLenLim)
                    Cell.Value = NewCellValue
                Else
                    Application.Undo
                    Exit Sub
                End If
                Exit Sub
            ElseIf (Len(NewCellValue) > StringLenLim And Len(NewCellValue2) > StringLenLim And (Cell.Column = Col_1 Or Cell.Column = Col_2)) Then
                Result = MsgBox(Mssg & _
                         vbCrLf & vbCrLf & _
                         "Trim to " & CurLenLim & CondOpt & " (Ok) or undo (Cancel)?", _
                         vbQuestion + vbOKCancel)
                If Result = vbOK Then
                    Cell.Offset(, 1).Value = Left(NewCellValue2, StringLenLim)
                    NewCellValue = Left(NewCellValue, StringLenLim)
                    Cell.Value = NewCellValue
                Else
                    Application.Undo
                    Exit Sub
                End If
                Exit Sub
            ElseIf (Len(NewCellValue) > StringLenLim And Len(NewCellValue2) > StringLenLim And Not (Cell.Column = Col_1 And Cell.Column = Col_2)) Then
                Result = MsgBox(Mssg & _
                         vbCrLf & vbCrLf & _
                         "Trim to " & CurLenLim & CondOpt & " (Ok) or undo (Cancel)?", _
                         vbQuestion + vbOKCancel)
                If Result = vbOK Then
                    NewCellValue = Left(NewCellValue, StringLenLim)
                    Cell.Value = NewCellValue
                Else
                    Application.Undo
                    Exit Sub
                End If
                Exit Sub
            End If
            
        Next Cell
    End If
End Sub

Private Sub Connect_CA_Click()
' Procedure purpose:  To reconnect/refresh data sources

    Dim lResult     As Long, lRet As Boolean
    Dim range_1     As Range
    Dim i As Integer, ds As String, ds_name As String, ds_concat As String
    Dim InfoBox As VbMsgBoxResult
    Const AckTime As Integer = 3
    
    Call OnStart
    
    lRet = True
    For i = 1 To 7
        ds = "DS_" & i
        lResult = Application.Run("SAPGetProperty", "IsDataSourceActive", ds)
        lRet = lRet And lResult
        If lRet = False Then
            If ds_concat <> "" Then
                ds_concat = ds_concat & ", " & vbCrLf & "'" & ds & "'"
            Else
                ds_concat = "'" & ds & "'"
            End If
        End If
    Next i
    
    If ds_concat <> "" Then
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
End Sub

Private Sub Save_CA_Click()
' Procedure purpose:  To save data in Planning Query

    Dim lResult     As Long
    Dim lRefreshDate As Double
    Dim StartTime   As Double
    Dim EndTime     As Double
    Dim wb          As Workbook
    Dim ds_alias As String: ds_alias = "DS_3"
    Dim InfoBox As VbMsgBoxResult
    
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
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Excel.Range)
' Procedure purpose:  To enable floating buttons

    On Error GoTo 0
    With Cells(Windows(1).ScrollRow, Windows(1).ScrollColumn)
        Connect_CA.Top = .Top + 0
        Connect_CA.Left = .Left + 0
        Save_CA.Top = .Top + 0
        Save_CA.Left = .Left + 80
    End With
End Sub


------------------
Sheet4.cls

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


------------------
Sheet5.cls

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Connect_SP_Click()
' Procedure purpose:  To reconnect/refresh data sources

    Dim lResult     As Long, lRet As Boolean
    Dim range_1     As Range
    Dim i As Integer, ds As String, ds_name As String, ds_concat As String
    Dim InfoBox As VbMsgBoxResult
    Const AckTime As Integer = 3
    
    Call OnStart
    
    lRet = True
    For i = 1 To 7
        ds = "DS_" & i
        lResult = Application.Run("SAPGetProperty", "IsDataSourceActive", ds)
        lRet = lRet And lResult
        If lRet = False Then
            If ds_concat <> "" Then
                ds_concat = ds_concat & ", " & vbCrLf & "'" & ds & "'"
            Else
                ds_concat = "'" & ds & "'"
            End If
        End If
    Next i
    
    If ds_concat <> "" Then
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
End Sub

Private Sub Save_SP_Click()
' Procedure purpose:  To save data in Planning Query

    Dim lResult     As Long
    Dim lRefreshDate As Double
    Dim StartTime   As Double
    Dim EndTime     As Double
    Dim wb          As Workbook
    Dim ds_alias As String: ds_alias = "DS_5"
    Dim InfoBox As VbMsgBoxResult
    
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
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Excel.Range)
' Procedure purpose:  To enable floating buttons

    On Error GoTo 0
    With Cells(Windows(1).ScrollRow, Windows(1).ScrollColumn)
        Connect_SP.Top = .Top + 0
        Connect_SP.Left = .Left + 0
        Save_SP.Top = .Top + 0
        Save_SP.Left = .Left + 80
    End With
End Sub



------------------
Sheet6.cls

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


------------------
Sheet7.cls

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


------------------
Sheet8.cls

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


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
Public Sub Workbook_open()
' Procedure purpose:  To enable SAP Analysis plug-in, reconnect/refresh all data sources to BW system
    vFlag = 1
    
    Dim wb          As Workbook
    Dim ws          As Worksheet
    Dim lResult     As Long
    Dim addin       As COMAddIn
    Dim rng         As Range
    Dim cofCom      As Object
    Dim StartRow    As Integer
    
    On Error Resume Next
    Set cofCom = Application.COMAddIns("SapExcelAddIn").Object
    Set wb = ThisWorkbook
    
    If ActiveSheet.Name = "Export_CA" Then
        StartRow = 7
    ElseIf ActiveSheet.Name = "Export_SP" Then
        StartRow = 1
    Else
        StartRow = 1
    End If
    
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
    
    With ThisWorkbook.Sheets("Edit_CA")
        Set rng = Range("A1")
        rng.Select
    End With
    
    If (ActiveSheet.Name = "Export_CA") Or (ActiveSheet.Name = "Export_SP") Then
        llastrow = Range(Range("A65536").End(XlDirection.xlUp).Address).Row
        llastcol = Split(Cells(1, Range("A" & StartRow).End(xlToRight).Column).Address, "$")(1)
        Set rng = ActiveSheet.Range("A" & StartRow + 1, Range(llastcol & llastrow))
        Call RemoveHashCharacters(rng)
        Set rng = ActiveSheet.Range("S" & StartRow + 1, Range(llastcol & llastrow))
        Call RestoreLineBreaks(rng)
    End If
    
    Call Columns_Visibility
    
    Call OnEnd
    
    vFlag = 0
    
End Sub

Public Sub Workbook_SheetActivate(ByVal Sh As Object)
' Procedure purpose:  To refresh data during switch between worksheets

    vFlag = 1
    
    Dim lResult As Long, lRet As Boolean
    Dim lRefreshDate As Double
    Dim lDS_Alias   As String
    Dim lSeparator  As String * 1
    Dim lSeparator_Count As Integer
    Dim item As Variant, arr, c As Collection
    Dim rng As Range
    Dim Cell           As Range
    Dim llastcol    As String
    Dim llastrow    As Long
    Dim StartRow    As Integer
    Dim ws As Worksheet
    Dim InfoBox As VbMsgBoxResult
    
    Const AckTime As Integer = 3
    
    If ActiveSheet.Name = "Export_CA" Then
        StartRow = 7
    ElseIf ActiveSheet.Name = "Export_SP" Then
        StartRow = 1
    Else
        StartRow = 1
    End If
    
    Call OnStart
    
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
    lSeparator_Count = Len(lDS_Alias) - Len(Replace(lDS_Alias, lSeparator, ""))
    lRet = True
    
    If lDS_Alias <> "" Then
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
        If lDS_Alias <> "" Then
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
        
        If Sh.Name = "Edit_CA" Then
            Call DataValidationList
        ElseIf Sh.Name = "Export_CA" Then
        
            llastrow = Range(Range("A65536").End(XlDirection.xlUp).Address).Row
            llastcol = Split(Cells(1, Range("A" & StartRow).End(xlToRight).Column).Address, "$")(1)
            Set rng = ActiveSheet.Range("A" & StartRow + 1, Range(llastcol & llastrow))
            
            Call RemoveHashCharacters(rng)

            Set rng = ActiveSheet.Range("S" & StartRow + 1, Range(llastcol & llastrow))
            
            Call RestoreLineBreaks(rng)
            
        ElseIf Sh.Name = "Export_SP" Then
        
            llastrow = Range(Range("A65536").End(XlDirection.xlUp).Address).Row
            llastcol = Split(Cells(1, Range("A" & StartRow).End(xlToRight).Column).Address, "$")(1)
            Set rng = ActiveSheet.Range("A" & StartRow + 1, Range(llastcol & llastrow))
            
            Call RemoveHashCharacters(rng)
            
        End If
        
        
        lResult = Application.Run("SAPSetRefreshBehaviour", "On")
        
        Call Columns_Visibility
        
        AppActivate Application.Caption
        DoEvents
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
    
    Call OnEnd
    
    vFlag = 0

End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)

    Dim x As Range
    Dim Mssg        As String
    Dim ValidatedCells As Range
    Dim Cell        As Range
    
    Set x = ActiveSheet.UsedRange
    Set ValidatedCells = Intersect(Target, Range(x.Address))
    
     If Sh.Name Like "Export*" Or Sh.Name = "DevAccess" Then
        If vFlag = 0 Then
        Call OnStart
            If Not ValidatedCells Is Nothing Then
                For Each Cell In ValidatedCells
                    Result = MsgBox("Attempted to change the value of the cell: " & Target.Address, vbExclamation, "Warning")
                    If Result = vbOK Then
                        Application.Undo
                        Call OnEnd
                        Exit Sub
                    End If
                Next Cell
            End If
        Call OnEnd
        End If
    End If
    
End Sub