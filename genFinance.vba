'
' 產生繳費部分財務報表
'
Sub genFinance()
	Dim billSheet As Worksheet
	Set billSheet = ActiveWorkbook.Sheets("當月帳款")
	Dim finaSheet As Worksheet
	Set finaSheet = ActiveWorkbook.Sheets("繳費財報")
	Dim autoSheet As Worksheet
    Set autoSheet = ActiveWorkbook.Sheets("自動扣繳")
	' 自動扣繳戶別定義區間
    Dim auFirst As Integer
    Dim auLast As Integer
    auFirst = FindFirstResid(autoSheet.Range("D:D")) '用戶號碼所在欄位
    auLast = FindLastResid(autoSheet.Range("D:D"), auFirst)
    Dim auResid As Range
    Set auResid = autoSheet.Range("D" & auFirst & ":D" & auLast)
    Dim matchResult As Variant '儲存match結果，有找到會是行數
	' 未銷帳區間
	Dim tblDeficit As ListObject
    Set tblDeficit = ActiveWorkbook.Sheets("未銷帳").ListObjects("未銷帳戶")
	
	Dim rResid313 As Range
    Set rResid313 = billSheet.Range("A4:A316")
	Dim rngCell As Range
    'Dim sResid As String
	finaSheet.Range("A4:A316").Value = billSheet.Range("A4:A316").Value
	finaSheet.Range("B4:I316").ClearContents
	For Each rngCell In rResid313
		' 年季月繳打勾
		if billSheet.Range("I" & rngCell.Row).Value = 1 Then
			finaSheet.Range("B" & rngCell.Row).Value = ChrW(&H2713)
		ElseIf billSheet.Range("H" & rngCell.Row).Value = 1 Then
			finaSheet.Range("C" & rngCell.Row).Value = ChrW(&H2713)
		Else
			finaSheet.Range("D" & rngCell.Row).Value = ChrW(&H2713)
		End If
		' 代扣打勾
		matchResult = Application.Match(rngCell.Value, auResid, 0)
        If IsNumeric(matchResult) Then
			finaSheet.Range("E" & rngCell.Row).Value = ChrW(&H2713)
		End If
		' 應繳金額
		finaSheet.Range("F" & rngCell.Row).Value = billSheet.Range("D" & rngCell.Row).Value
		' 是否銷帳
		If isResid(ActiveWorkbook.Sheets("未銷帳").Range("A2").Value) Then
			matchResult = Application.Match(rngCell.Value, tblDeficit.ListColumns("繳款人識別碼").DataBodyRange, 0)
			If Not IsNumeric(matchResult) Then
				finaSheet.Range("G" & rngCell.Row).Value = ChrW(&H2713)
			Else
				finaSheet.Range("H" & rngCell.Row).Value = _
					tblDeficit.ListColumns("繳款總額").DataBodyRange(matchResult).Value
			End If
		Else
			finaSheet.Range("G" & rngCell.Row).Value = ChrW(&H2713)
		End If
	Next rngCell
End Sub