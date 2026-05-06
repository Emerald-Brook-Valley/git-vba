'
' 移除自動扣繳或代扣清冊中，款項為0者
' 自動扣繳下方尚有資料，與代扣清冊刪除方式不同
' 同時產生sheet2, email銀行用。有帳號、身分證統一編號、金額、用戶號碼四欄
'
Sub removeNoFeeDeduct(filePrefix As String)
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
	Dim iFirst As Integer
    Dim iLast As Integer

    Do While xlsxFile <> ""
		Set wb = Workbooks.Open(xlsxFile)
		Set ws = wb.Worksheets(1)
		iFirst = FindFirstResid(ws.Range("D:D"))
		iLast = FindLastResid(ws.Range("D:D"), iFirst)
		'lastRow = ws.Cells(ws.Rows.Count, feeColumn).End(xlUp).Row
		' Loop backwards from lastRow to the first row
		For i = iLast To iFirst Step -1
			If ws.Cells(i, feeColumn).Value = 0 Then
				Rows(i).Delete Shift:=xlUp
			End If
		Next i
		' ---------------------------
		' 新增sheet2做email銀行用
		' ---------------------------
		Dim ws2 As Worksheet
		Set ws2 = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
		On Error Resume Next '忽略如果sheet名稱已經存在
		ws2.Name = "email2bank"
		On Error GoTo 0
		ws2.Range("A1") = "帳號"
		ws2.Range("B1") = "身分證統一編號"
		ws2.Range("C1") = "金額"
		ws2.Range("D1") = "用戶號碼"
		iFirst = FindFirstResid(ws.Range("D:D"))
		iLast = FindLastResid(ws.Range("D:D"), iFirst)
		ws2.Range("A2:" & "A" & iLast - iFirst + 2).NumberFormat = "@"
		ws2.Range("A2:" & "A" & iLast - iFirst + 2).Value = ws.Range("B" & iFirst & ":B" & iLast).Value
		ws2.Range("C2:" & "C" & iLast - iFirst + 2).Value = ws.Range("E" & iFirst & ":E" & iLast).Value
		ws2.Range("D2:" & "D" & iLast - iFirst + 2).Value = ws.Range("D" & iFirst & ":D" & iLast).Value
		ws2.Cells.EntireColumn.AutoFit
		With ws2.Range("A1:D" & iLast - iFirst + 2).Borders
			.LineStyle = xlContinuous
			.Weight = xlThin
			.ColorIndex = 0 ' Automatic (Black)
		End With
		
		ws.Activate
		wb.Close SaveChanges:=True
		' Get next file
        xlsxFile = Dir()
    Loop
	
End Sub