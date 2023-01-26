ThisWorkBook

	Public Sub Workbook_open()
	' Procedure purpose:  To enable SAP Analysis plug-in, reconnect/refresh all data sources to BW system

		Dim wb          As Workbook
		Dim ws          As Worksheet
		Dim lResult     As Long
		Dim addin       As COMAddIn
		Dim rng As Range
		Dim cofCom As Object
		
		On Error Resume Next
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
	' Procedure purpose:  To refresh data during switch between worksheets
		
		Dim lResult As Long, lRet As Boolean
		Dim lRefreshDate As Double
		Dim lDS_Alias   As String
		Dim lSeparator  As String * 1
		Dim lSeparator_Count As Integer
		Dim item As Variant, arr, c As Collection
		Dim rng As Range
		Dim llastrow As Long
		Dim ws As Worksheet
		Dim InfoBox As VbMsgBoxResult
		
		Const AckTime As Integer = 3
		
		Call OnStart
		
		StartTime = Timer
		
		AppActivate Application.Caption
		DoEvents
		InfoBox = TimedMsgBox("Refresh on tab """ & Sh.Name & """ in progress..." _
			 & vbCrLf & vbCrLf & "(This window will close in " & AckTime & " seconds)", _
			 "Data refresh on tab """ & Sh.Name & """", , AckTime)
		 
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
			lResult = Application.Run("SAPExecuteCommand", "Restart", lDS_Alias)
			
			Select Case Sh.Name
				Case "Export"
					lDS_Alias = "DS_2"
					lRefreshDate = Application.Run("SAPGetSourceInfo", lDS_Alias, "QueryLastRefreshedAt")
					Call SetPrintLayout(lRefreshDate)
				Case Else
					lRefreshDate = Application.Run("SAPGetSourceInfo", lDS_Alias, "QueryLastRefreshedAt")
			End Select
			
			
			If Sh.Name = "Edit" Then
				Call DataValidationList
			ElseIf Sh.Name = "Export" Then
				Call RemoveHashCharacters
				Call RestoreLineBreaks
				Set ws = ThisWorkbook.Sheets("Export")
				ws.Range("R:R,T:T").Select
				Selection.EntireColumn.Hidden = True
			End If
			
			lResult = Application.Run("SAPSetRefreshBehaviour", "On")
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

	End Sub

-------------
Sheet3(Edit)

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
		
		Dim CurLenLim   As Integer
		Dim CondOpt     As String
		Dim Mssg        As String
		
		llastrow = Range(ActiveSheet.Range("A65536").End(XlDirection.xlUp).Address).Row
		
		Set ValidatedCells = Intersect(Target, Target.Parent.Range("Q3:R" & llastrow, "T3:V" & llastrow))
		If Not ValidatedCells Is Nothing Then
			For Each Cell In ValidatedCells
				If Cell.Value <> "" Then
					NewCellValue = Replace(Replace(Cell.Value, vbCr, LineSeparator), vbLf, LineSeparator)
				End If
				If Len(NewCellValue) > 250 Then
					NewCellValue2 = Right(NewCellValue, Len(NewCellValue) - StringLenLim)
				End If
				If (Cell.Column = 17 Or Cell.Column = 20) Then
					CurLenLim = StringLenLim * 2
					CondOpt = " and split it into 2 columns "
				Else
					CurLenLim = StringLenLim
					CondOpt = Empty
				End If
				
				Mssg = "The value" & _
						" inserted in cell " & Cell.Address & _
						" exceeds accepted field length by " & _
						Len(NewCellValue) - StringLenLim & " characters."
				
				If (Len(NewCellValue) > StringLenLim And Len(NewCellValue2) <= StringLenLim And (Cell.Column = 17 Or Cell.Column = 20)) Then
					Result = MsgBox(Mssg & _
							 vbCrLf & vbCrLf & _
							 "Split it into 2 columns (Ok) or undo (Cancel)?", _
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
				ElseIf (Len(NewCellValue) > StringLenLim And Len(NewCellValue2) <= StringLenLim And Not (Cell.Column = 17 And Cell.Column = 20)) Then
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
				ElseIf (Len(NewCellValue) > StringLenLim And Len(NewCellValue2) > StringLenLim And (Cell.Column = 17 Or Cell.Column = 20)) Then
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
				ElseIf (Len(NewCellValue) > StringLenLim And Len(NewCellValue2) > StringLenLim And Not (Cell.Column = 17 And Cell.Column = 20)) Then
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

	Private Sub Connect_Click()
	' Procedure purpose:  To reconnect/refresh data sources

		Dim lResult     As Long, lRet As Boolean
		Dim range_1     As Range
		Dim i As Integer, ds As String, ds_name As String, ds_concat As String
		Dim InfoBox As VbMsgBoxResult
		Const AckTime As Integer = 3
		
		Call OnStart
		
		lRet = True
		For i = 1 To 5
			ds = "DS_" & i
			lResult = Application.Run("SAPGetProperty", "IsConnected", ds)
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
			
			ThisWorkbook.Sheets("Edit").Activate
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
		
		Call OnEnd
	End Sub

	Private Sub Save_Click()
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
		
		lResult = Application.Run("SAPGetProperty", "IsConnected", ds_alias)
		If lResult = True Then
			
			StartTime = Timer
			lResult = Application.Run("SAPSetRefreshBehaviour", "Off")
			lResult = Application.Run("SAPDeleteDesignRule", ds_alias)
			lResult = Application.Run("SAPExecuteCommand", "PlanDataSave")
			lResult = Application.Run("SAPExecuteCommand", "Restart", "ALL")
			
			Call DataValidationList
			
			EndTime = Timer
			lResult = Application.Run("SAPSetRefreshBehaviour", "On")
			wb.Sheets("Edit").Range("E1").Value = Format((EndTime - StartTime) / 86400, "hh:mm:ss") & " [hh:mm:ss]"
			
			InfoBox = TimedMsgBox("Data saved in : " & Format((EndTime - StartTime) / 86400, "hh:mm:ss") & " [hh:mm:ss]" _
				& vbCrLf & vbCrLf & "(This window will close in " & AckTime & " seconds)", _
				"Data saved", , AckTime)
			
		Else
			InfoBox = TimedMsgBox("You are not connected to the system" _
				& vbCrLf & vbCrLf & "(This window will close in " & AckTime & " seconds)", _
				"Connection Status", , AckTime)
		End If
		
		Call OnEnd
	End Sub

	Private Sub Worksheet_SelectionChange(ByVal Target As Excel.Range)
	' Procedure purpose:  To enable floating buttons

		On Error GoTo 0
		With Cells(Windows(1).ScrollRow, Windows(1).ScrollColumn)
			Connect.Top = .Top + 0
			Connect.Left = .Left + 0
			Save.Top = .Top + 0
			Save.Left = .Left + 80
		End With
	End Sub

-------------
Module1

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
		
		'    Call LockSheets
		
		ActiveSheet.EnableCalculation = True
		Application.AskToUpdateLinks = True
		Application.Calculation = xlAutomatic
		Application.DisplayAlerts = True
		Application.EnableEvents = True
		Application.ScreenUpdating = True
		
	End Sub

	Public Sub DataValidationList()
	' Procedure purpose:  To add data validation list on certain range (Approval flag)
		
		Dim rng         As Range
		Dim ws          As Worksheet
		Dim llastrow    As Long
		
		Set ws = ThisWorkbook.Worksheets("Edit")
		ws.Activate
		llastrow = Range(ActiveSheet.Range("A65536").End(XlDirection.xlUp).Address).Row
		Set rng = ActiveSheet.Range("A3", Range("V" & llastrow))
		rng.WrapText = True
		Set rng = ActiveSheet.Range("S3", Range("S" & llastrow))
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

	Public Sub RemoveHashCharacters()
	' Procedure purpose:  To replace "#" characters with blank value in particular range on "Export" tab
		
		Dim rng         As Range
		Dim Cell           As Range
		Dim llastrow    As Long
		Dim DataRange   As Variant
		Dim Irow        As Long, rowcnt As Integer
		Dim Icol        As Integer, colcnt As Integer
		Dim Cell_1      As String
		
		llastrow = Range(Range("A65536").End(XlDirection.xlUp).Address).Row
		Set rng = ActiveSheet.Range("A6", Range("A" & llastrow).End(xlToRight))
		
		DataRange = rng.Value
		rowcnt = rng.Rows.Count
		colcnt = rng.Columns.Count
		
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

	Public Sub RestoreLineBreaks()
	' Procedure purpose:  To replace "&&" with line breaks (VbCrLf) in certain range on the "Export" tab _
	and to merge cells that were splitted due to max. field length (250)
		
		Dim rng         As Range
		Dim Cell           As Range
		Dim llastrow    As Long
		Dim DataRange   As Variant
		Dim Irow        As Long, rowcnt As Integer
		Dim Icol        As Integer, colcnt As Integer
		Dim Cell_1      As String, Cell_2 As String
		
		Const LineSeparator As String = "&&"
		
		llastrow = Range(Range("A65536").End(XlDirection.xlUp).Address).Row
		Set rng = ActiveSheet.Range("Q6", Range("Q" & llastrow).End(xlToRight))
		rowcnt = rng.Rows.Count
		colcnt = rng.Columns.Count
		
		DataRange = rng.Value
		
		For Irow = 1 To rowcnt
			For Icol = 1 To colcnt
				Cell_1 = DataRange(Irow, Icol)
				If Cell_1 <> "" Then
					If (Icol = 1 Or Icol = 3) Then
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
		
		rng.WrapText = True
		rng.Columns.EntireColumn.AutoFit
		rng.Rows.EntireRow.AutoFit
		
	End Sub

	Public Sub SetPrintLayout(lRefreshDate As Double)
	' Procedure purpose:  To set print area, header information (Query last refresh date & time) on the "Export" tab
		
		Dim ws          As Worksheet
		Set ws = ThisWorkbook.Sheets("Export")
		
		ws.Activate
		
		llastrow = Range(ws.Range("A65536").End(XlDirection.xlUp).Address).Row
		Set rng = ws.Range("A1", Range("A" & llastrow).End(xlToRight))
		ws.PageSetup.PrintArea = rng.Address
		ws.PageSetup.CenterHeader = "QueryLastRefreshedAt: " & Format(lRefreshDate, "dddd, mmmm d, yyyy h:mm:ss")
		
	End Sub