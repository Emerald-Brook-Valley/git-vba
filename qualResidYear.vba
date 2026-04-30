'
' 更新年繳名單(Month=10)
' 1. 檢查是否於未銷帳清單內
' 2. 應優先於季繳資料載入
' 3. 非資料更新月份，檢查名單與紀錄是否符合，不符合則提出警告，並以舊紀錄覆蓋
' 4. 處理11月資料時，年繳未繳標示為-1，並從年繳戶移除。計算管理費時須扣除未銷帳年費，並新增10月欠繳
'
Public Sub qualResidYear()
    Dim iMonth As Integer
    iMonth = ActiveWorkbook.Sheets("基本資料").Range("B16").Value
    Dim sFeePeriod As String
    sFeePeriod = ActiveWorkbook.Sheets("基本資料").Range("B27").Value
    Dim validMonth As Variant
    Dim checkMonth As Variant
    validMonth = Array(10)
    checkMonth = Array(11)
    Dim tblYear As ListObject
    Set tblYear = ActiveWorkbook.Sheets("年繳名單").ListObjects("年繳戶別")
    Dim tblDeficit As ListObject
    Set tblDeficit = ActiveWorkbook.Sheets("未銷帳").ListObjects("未銷帳戶")
    Dim rngCell As Range
    Dim sResid As String
    Dim rResid313 As Range
    Set rResid313 = ActiveWorkbook.Sheets("當月帳款").Range("A4:A316")
    Dim matchResult As Variant
	Dim logFile As String
	logFile = ActiveWorkbook.Path & "\logfile.txt"
    ' ---------------------------
    ' 更新年繳名單
    ' ---------------------------
    If IsNumeric(Application.Match(iMonth, validMonth, 0)) Then
        For Each rngCell In tblYear.ListColumns("年繳戶").DataBodyRange
            sResid = rngCell.Value
            ' 檢查未銷帳清單
			'確認未銷帳非空
			If isResid(ActiveWorkbook.Sheets("未銷帳").Range("A2").Value) Then
				If IsNumeric(Application.Match(sResid, tblDeficit.ListColumns("繳款人識別碼").DataBodyRange, 0)) Then
					Call appendLog(logFile, "年繳戶" & sResid & " Found in Deficit. address=" & rngCell.Address)
					ActiveWorkbook.Sheets("年繳名單").Range("D6").Insert Shift:=xlDown
					ActiveWorkbook.Sheets("年繳名單").Range("D6").Value = sResid & "因尚有欠款，自年繳名單移除" & sFeePeriod
					rngCell.ClearContents
				End If
			End If
        Next rngCell
        tblYear.Sort.Apply
        ' 不符資格年繳戶清除完畢，更新年繳欄位
        For Each rngCell In rResid313
            If IsNumeric(Application.Match(rngCell.Value, tblYear.ListColumns("年繳戶").DataBodyRange, 0)) Then
                ActiveWorkbook.Sheets("當月帳款").Range("I" & rngCell.Row).Value = 1
            Else
                ActiveWorkbook.Sheets("當月帳款").Range("I" & rngCell.Row).Value = 0
            End If
        Next rngCell
    ' ---------------------------
    ' 非年繳月份，名單須維持不變
    ' ---------------------------
    Else '非年繳處理月份，檢查清單是否與紀錄不符
        ' 是否有不當新增
        For Each rngCell In tblYear.ListColumns("年繳戶").DataBodyRange
            sResid = rngCell.Value
            'If ActiveWorkbook.Sheets("當月帳款").Range("I" & rResid313.Find(sResid).Row).Value = 1 Then
            matchResult = Application.Match(sResid, rResid313, 0)
            If IsNumeric(matchResult) Then
                If ActiveWorkbook.Sheets("當月帳款").Range("I" & matchResult + 3).Value = 1 Then
                    'MsgBox sResid & "正確"
                Else
					Call appendLog(logFile, "年繳戶" & sResid & " 不在上月名單中，將被移除")
                    ActiveWorkbook.Sheets("年繳名單").Range("D6").Insert Shift:=xlDown
                    ActiveWorkbook.Sheets("年繳名單").Range("D6").Value = sResid & "非年繳申請期間請勿更新，將自年繳名單移除" & sFeePeriod
                    rngCell.ClearContents
                End If
            Else
				Call appendLog(logFile, "年繳存在無法識別的戶名" & sResid & "或空白")
                'Exit Sub
            End If
        Next rngCell
        ' 是否有不當刪除
        For Each rngCell In ActiveWorkbook.Sheets("當月帳款").Range("I4:I316")
            If rngCell.Value = 1 Then
                sResid = ActiveWorkbook.Sheets("當月帳款").Range("A" & rngCell.Row).Value
                If IsNumeric(Application.Match(sResid, tblYear.ListColumns("年繳戶").DataBodyRange, 0)) Then
                    'MsgBox sResid & "正確"
                Else
					Call appendLog(logFile, "年繳戶" & sResid & " 不當移除，將被加回")
                    ActiveWorkbook.Sheets("年繳名單").Range("A4").Insert Shift:=xlDown
                    ActiveWorkbook.Sheets("年繳名單").Range("A4").Value = sResid
                    ActiveWorkbook.Sheets("年繳名單").Range("D6").Value = sResid & "非年繳申請期間請勿更新，將自年繳名單加回" & sFeePeriod
                End If
                
            End If
        Next rngCell
        tblYear.Sort.Apply
    End If
    ' ---------------------------
    ' 年繳未繳處理
    ' ---------------------------
    If IsNumeric(Application.Match(iMonth, checkMonth, 0)) Then
        For Each rngCell In tblYear.ListColumns("年繳戶").DataBodyRange
            sResid = rngCell.Value
            ' 檢查未銷帳清單
			'確認未銷帳非空
			If isResid(ActiveWorkbook.Sheets("未銷帳").Range("A2").Value) Then
				If IsNumeric(Application.Match(sResid, tblDeficit.ListColumns("繳款人識別碼").DataBodyRange, 0)) Then
					'MsgBox sResid & " Found in Deficit. address=" & rngCell.Address
					Call appendLog(logFile, "年繳戶" & sResid & " Found in Deficit. address=" & rngCell.Address)
					matchResult = Application.Match(sResid, rResid313, 0)
					ActiveWorkbook.Sheets("當月帳款").Range("I" & matchResult + 3).Value = -1
					ActiveWorkbook.Sheets("年繳名單").Range("D6").Insert Shift:=xlDown
					ActiveWorkbook.Sheets("年繳名單").Range("D6").Value = sResid & "因年繳未繳，改為月繳戶並加計欠款一期" & sFeePeriod
					rngCell.ClearContents
				End If
			End If
        Next rngCell
        tblYear.Sort.Apply
    End If
End Sub
