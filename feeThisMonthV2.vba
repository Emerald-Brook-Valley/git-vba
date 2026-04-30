'
' 當月帳款處理
' V2 將欠款，行政規費，其他費用改以313戶搜尋
'
Sub feeThisMonth()
    Dim iMonth As Integer
    iMonth = ActiveWorkbook.Sheets("基本資料").Range("B16").Value
    Dim iYear As Integer
    iYear = ActiveWorkbook.Sheets("基本資料").Range("B17").Value
    Dim seasonMonth As Variant
    seasonMonth = Array(1, 4, 7, 10)
    Dim rngCell As Range
    Dim sResid As String
    Dim billSheet As Worksheet
    Set billSheet = ActiveWorkbook.Sheets("當月帳款")
    Dim landSheet As Worksheet
    Set landSheet = ActiveWorkbook.Sheets("坪數")
    Dim yearDiscount As Double '年繳優惠
    yearDiscount = ActiveWorkbook.Worksheets("基本資料").Range("B10").Value
    Dim seasonDiscount As Double '季繳優惠
    seasonDiscount = ActiveWorkbook.Worksheets("基本資料").Range("B8").Value
    Dim rResid313 As Range
    Set rResid313 = billSheet.Range("A4:A316")
    Dim lValA As Long
    Dim lValB As Long
    Dim iPeriod As Integer
    Dim tblDeficit As ListObject
    Set tblDeficit = ActiveWorkbook.Sheets("未銷帳").ListObjects("未銷帳戶")
    Dim tblFine As ListObject
    Set tblFine = ActiveWorkbook.Sheets("行政規費").ListObjects("行政規費")
    Dim tblSpecial As ListObject
    Set tblSpecial = ActiveWorkbook.Sheets("其他費用").ListObjects("其他費用")
    Dim matchResult As Variant
	Dim iRow As Integer
    
	' 警告:檢查公告未繳
	Dim response As VbMsgBoxResult
	'MsgBox "請檢查sheet:公告未繳，是否資訊正確。請對照上期未繳公告完成修正後，回到起始表進行下一步驟。"
	response = MsgBox("請檢查sheet:公告未繳，是否資訊正確。請對照上期未繳公告完成修正後，回到起始表進行此步驟。" & vbCrLf & vbCrLf & _
		"是(Yes):前往公告未繳。    否(No):已完成修正，繼續執行。", vbYesNo + vbQuestion, "檢查公告未繳")
	If response = vbYes Then
		Worksheets("公告未繳").Activate
		Exit Sub
	End If
	
	' 複製上期累計欠款期數
    Call updateSumDebt

    ' ---------------------------
    ' 執行313戶之房屋、車位管理費計算。
    ' ---------------------------
    For Each rngCell In rResid313
		sResid = rngCell.Value
		iRow = rngCell.Row
        If billSheet.Range("I" & iRow) = 1 Then '年繳戶處理
            If iMonth = 10 Then
                billSheet.Range("B" & iRow).Value = _
                    CLng(CDbl(landSheet.Range("F" & iRow).Value * 12) * (1 - yearDiscount)) '房屋費
                billSheet.Range("C" & iRow).Value = _
					CLng(CDbl(landSheet.Range("G" & iRow).Value * 12))
                    'CLng(CDbl(landSheet.Range("G" & iRow).Value * 12) * (1 - yearDiscount)) '車位費
					' 目前車位無優惠
                ' 寫入通訊欄1
				billSheet.Range("L" & iRow).Value = _
                    comName & iYear & "/" & iMonth & "~" & iYear + 1 & "/" & iMonth - 1 & _
					"年繳管理費(" & Format(yearDiscount, "0%") & "優惠)。"
            Else
                billSheet.Range("B" & iRow).Value = 0
                billSheet.Range("C" & iRow).Value = 0
                billSheet.Range("L" & iRow).Value = comName & "管理費。"
            End If
        ElseIf billSheet.Range("I" & iRow) = -1 Then '年繳戶除名處理
            lValA = landSheet.Range("F" & iRow).Value
            lValB = landSheet.Range("G" & iRow).Value
            billSheet.Range("B" & iRow).Value = lValA
            billSheet.Range("C" & iRow).Value = lValB
			' ---------------------------
			' 除名年繳戶未銷帳特別處理
			' ---------------------------
            'billSheet.Range("E" & iRow).Value = lValA + lValB '前期欠款
			'確認未銷帳非空
			If isResid(ActiveWorkbook.Sheets("未銷帳").Range("A2").Value) Then
				matchResult = Application.Match(sResid, tblDeficit.ListColumns("繳款人識別碼").DataBodyRange, 0)
				If IsNumeric(matchResult) Then
					' 欠費金額為未銷帳扣除年繳金額，加上單月費
					billSheet.Range("E" & iRow).Value = _
						tblDeficit.ListColumns("繳款總額").DataBodyRange(matchResult).Value - _
						CLng(CDbl(landSheet.Range("F" & iRow).Value * 12) * (1 - yearDiscount)) - _
						CLng(CDbl(landSheet.Range("G" & iRow).Value * 12)) + _
						CLng(CDbl(landSheet.Range("F" & iRow).Value)) + CLng(CDbl(landSheet.Range("G" & iRow).Value))
					' 同步修正未銷帳
					tblDeficit.ListColumns("繳款總額").DataBodyRange(matchResult).Value = _
						billSheet.Range("E" & iRow).Value
					billSheet.Range("F" & iRow).Value = billSheet.Range("G" & iRow).Value + 1
					billSheet.Range("L" & iRow).Value = billSheet.Range("L" & iRow).Value & _
						"含前" & billSheet.Range("F" & iRow).Value & "期未繳金額。(失去年繳資格)" 'append to 通訊欄1
				End If
			Else
				billSheet.Range("E" & iRow).Value = 0
			End If
			' ---------------------------
            iPeriod = 1
            billSheet.Range("F" & iRow).Value = iPeriod '欠款期數
            billSheet.Range("I" & iRow) = 0 '轉為月繳
            billSheet.Range("L" & iRow).Value = _
                comName & iYear & "年" & iMonth & "月與前" & iPeriod & "期管理費。"
        ElseIf billSheet.Range("H" & iRow) = 1 Then '季繳戶處理
            If IsNumeric(Application.Match(iMonth, seasonMonth, 0)) Then '季繳月份
                billSheet.Range("B" & iRow).Value = _
                    CLng(CDbl(landSheet.Range("F" & iRow).Value * 3) * (1 - seasonDiscount))
                billSheet.Range("C" & iRow).Value = _
					CLng(CDbl(landSheet.Range("G" & iRow).Value * 3))
                    'CLng(CDbl(landSheet.Range("G" & iRow).Value * 3) * (1 - seasonDiscount))
					' 目前車位無優惠
                ' 寫入通訊欄1
				billSheet.Range("L" & iRow).Value = _
                    comName & iYear & "年" & iMonth & "~" & iMonth + 2 & _
					"月季繳管理費(" & Format(seasonDiscount, "0%") & "優惠)。"
            Else
                billSheet.Range("B" & iRow).Value = 0
                billSheet.Range("C" & iRow).Value = 0
                billSheet.Range("L" & iRow).Value = comName & "管理費。"
            End If
        ElseIf billSheet.Range("H" & iRow) = -1 Then '季繳戶除名處理
            lValA = landSheet.Range("F" & iRow).Value
            lValB = landSheet.Range("G" & iRow).Value
            billSheet.Range("B" & iRow).Value = lValA
            billSheet.Range("C" & iRow).Value = lValB
			' ---------------------------
			' 除名季繳戶未銷帳特別處理
			' ---------------------------
            'billSheet.Range("E" & iRow).Value = lValA + lValB '前期欠款
			'確認未銷帳非空
			If isResid(ActiveWorkbook.Sheets("未銷帳").Range("A2").Value) Then
				matchResult = Application.Match(sResid, tblDeficit.ListColumns("繳款人識別碼").DataBodyRange, 0)
				If IsNumeric(matchResult) Then
					' 欠費金額為未銷帳扣除季繳金額，加上單月費
					billSheet.Range("E" & iRow).Value = _
						tblDeficit.ListColumns("繳款總額").DataBodyRange(matchResult).Value - _
						CLng(CDbl(landSheet.Range("F" & iRow).Value * 3) * (1 - seasonDiscount)) - _
						CLng(CDbl(landSheet.Range("G" & iRow).Value * 3)) + _
						CLng(CDbl(landSheet.Range("F" & iRow).Value)) + CLng(CDbl(landSheet.Range("G" & iRow).Value))
					' 同步修正未銷帳
					tblDeficit.ListColumns("繳款總額").DataBodyRange(matchResult).Value = _
						billSheet.Range("E" & iRow).Value
					billSheet.Range("F" & iRow).Value = billSheet.Range("G" & iRow).Value + 1
					billSheet.Range("L" & iRow).Value = billSheet.Range("L" & iRow).Value & _
						"含前" & billSheet.Range("F" & iRow).Value & "期未繳金額。(失去季繳資格)" 'append to 通訊欄1
				End If
			Else
				billSheet.Range("E" & iRow).Value = 0
			End If
			' ---------------------------
            iPeriod = 1
            billSheet.Range("F" & iRow).Value = iPeriod '欠款期數
            billSheet.Range("L" & iRow).Value = _
                comName & iYear & "年" & iMonth & "月與前" & iPeriod & "期管理費。"
            billSheet.Range("H" & iRow) = 0 '轉為月繳
        Else '月繳處理
            lValA = landSheet.Range("F" & iRow).Value
            lValB = landSheet.Range("G" & iRow).Value
            billSheet.Range("B" & iRow).Value = lValA
            billSheet.Range("C" & iRow).Value = lValB
            ' 寫入通訊欄1
			billSheet.Range("L" & iRow).Value = _
                comName & iYear & "年" & iMonth & "月管理費。"
        End If
    Next rngCell
	
	' ---------------------------
    ' 填入未銷帳、行政規費、其他費用
    ' ---------------------------
	For Each rngCell In rResid313
		sResid = rngCell.Value
		iRow = rngCell.Row
		' 讀入未銷帳
		'確認未銷帳非空
		If isResid(ActiveWorkbook.Sheets("未銷帳").Range("A2").Value) Then
			matchResult = Application.Match(sResid, tblDeficit.ListColumns("繳款人識別碼").DataBodyRange, 0)
			If IsNumeric(matchResult) Then
				billSheet.Range("E" & iRow).Value = _
					tblDeficit.ListColumns("繳款總額").DataBodyRange(matchResult).Value
				billSheet.Range("F" & iRow).Value = billSheet.Range("G" & iRow).Value + 1
				billSheet.Range("L" & iRow).Value = billSheet.Range("L" & iRow).Value & _
					"含前" & billSheet.Range("F" & iRow).Value & "期未繳金額。" 'append to 通訊欄1
			End If
		Else
			billSheet.Range("E" & iRow).Value = 0
		End If
		' 讀入行政規費
		'確認行政規費非空
		If isResid(ActiveWorkbook.Sheets("行政規費").Range("A4").Value) Then
			matchResult = Application.Match(sResid, tblFine.ListColumns("戶別").DataBodyRange, 0)
			If IsNumeric(matchResult) Then
				billSheet.Range("J" & iRow).Value = _
					tblFine.ListColumns("行政規費").DataBodyRange(matchResult).Value
				billSheet.Range("M" & iRow).Value = "行政規費違規編號" & _
					tblFine.ListColumns("違規編號").DataBodyRange(matchResult).Value & "。" '寫入通訊欄2
			End If
		Else
			billSheet.Range("J" & iRow).Value = 0
		End If
		' 讀入其他費用
		'確認其他費用非空
		If isResid(ActiveWorkbook.Sheets("其他費用").Range("A4").Value) Then
			matchResult = Application.Match(sResid, tblSpecial.ListColumns("戶別").DataBodyRange, 0)
			If IsNumeric(matchResult) Then
				billSheet.Range("K" & iRow).Value = _
					tblSpecial.ListColumns("其他費用").DataBodyRange(matchResult).Value
				billSheet.Range("M" & iRow).Value = billSheet.Range("M" & iRow).Value & "其他費用:" & _
					tblSpecial.ListColumns("訊息").DataBodyRange(matchResult).Value 'append to 通訊欄2
			End If
		Else
			billSheet.Range("K" & iRow).Value = 0
		End If
	Next rngCell

    ' ---------------------------
    ' 管理費用加總
    ' ---------------------------
    For Each rngCell In rResid313
        billSheet.Range("D" & rngCell.Row).Value = _
            billSheet.Range("B" & rngCell.Row).Value + billSheet.Range("C" & rngCell.Row).Value + _
            billSheet.Range("E" & rngCell.Row).Value + billSheet.Range("J" & rngCell.Row).Value + _
            billSheet.Range("K" & rngCell.Row).Value
    Next rngCell
	
	' ---------------------------
	' 將住戶Email相關資訊寫入備註
	' ---------------------------
	Dim tblEmail As ListObject
    Set tblEmail = ActiveWorkbook.Sheets("Email").ListObjects("Email")
	For Each rngCell In rResid313
	'For Each rngCell In tblEmail.ListColumns("戶別").DataBodyRange
        sResid = rngCell.Value
        matchResult = Application.Match(sResid, tblEmail.ListColumns("戶別").DataBodyRange, 0)
        If IsNumeric(matchResult) Then
			' 寫入備註
            billSheet.Range("P" & rngCell.Row).Value = "貴戶電子郵件 " & _
                tblEmail.ListColumns("電子郵件").DataBodyRange(matchResult).Value & _
				"。如有錯誤請至櫃台更正。"
		Else
			billSheet.Range("P" & rngCell.Row).Value = "社區提供電子郵件收取繳費通知，請上 " & _
				"https://forms.gle/pDPZ5MfdHfVgCVJa8 登記，或至櫃台登記。"
        End If
    Next rngCell
    
    MsgBox "當月帳款計算完成"
    ' 插入完成check
    insertCheck (ActiveWorkbook.Sheets("起始表").Range("G6"))
	Dim logFile As String
	logFile = ActiveWorkbook.Path & "\logfile.txt"
	Call appendLog(logFile, "Finished feeThisMonth")
End Sub
