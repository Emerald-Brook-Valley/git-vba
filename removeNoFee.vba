'
' 移除自動扣繳或代扣清冊中，款項為0者
'
Sub removeNoFee(filePrefix As String)
	Dim feeColumn As String
	If filePrefix = "代扣通知" Then
		feeColumn = "D"
	ElseIf filePrefix = "自動扣繳" Then
		feeColumn = "E"
	Else
		Exit Sub
	End If
	
	Dim xlsxFile As String
    xlsxFile = Dir(ActiveWorkbook.Path & "\" & filePrefix & "*.xlsx")    
	Dim wb As Workbook
	Dim ws As Worksheet
	Dim lastRow As Long
    Dim i As Long
    Do While xlsxFile <> ""
		Set wb = Workbooks.Open(xlsxFile)
		Set ws = wb.Worksheets(1)
		lastRow = ws.Cells(ws.Rows.Count, feeColumn).End(xlUp).Row
		' Loop backwards from lastRow to the first row
		For i = lastRow To 1 Step -1
			If ws.Cells(i, feeColumn).Value = 0 Then
				Rows(i).Delete Shift:=xlUp
			End If
		Next i
		wb.Close SaveChanges:=True
		' Get next file
        xlsxFile = Dir()
    Loop
	
End Sub