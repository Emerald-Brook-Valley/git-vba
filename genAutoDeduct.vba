'
' 產生自動扣繳授權清冊
' 1. 注意代繳戶姓名不一定與區權人相同，而是帳號所有人的姓名
' 2. 代繳金額為零者(例如季繳戶在非季繳月份)，可在輸出xlsx檔案後，決定處理方式
' 3. 扣繳失敗的處理方式，目前僅靠銀行通知管委會扣繳失敗戶，需手動加入未銷帳清單內
'
Sub genAutoDeduct()
    Dim billSheet As Worksheet
    Set billSheet = ActiveWorkbook.Sheets("當月帳款")
    Dim autoDeductSheet As Worksheet
    Set autoDeductSheet = ActiveWorkbook.Sheets("自動扣繳")
    Dim rngCell As Range
    Dim sResid As String
    Dim iRow As Integer
    ' ---------------------------
    ' 尋找自動扣繳戶別定義區間
    ' ---------------------------
    Dim auFirst As Integer
    Dim auLast As Integer
    auFirst = FindFirstResid(autoDeductSheet.Range("D:D")) '用戶號碼所在欄位
    auLast = FindLastResid(autoDeductSheet.Range("D:D"), auFirst)
    Dim auResid As Range
    Set auResid = autoDeductSheet.Range("D" & auFirst & ":D" & auLast)
    Dim matchResult As Variant '儲存match結果，有找到會是行數
    Dim sumFee As Long
	Dim deductMonth As String
	deductMonth = ActiveWorkbook.Sheets("基本資料").Range("B27").Value
    ' ---------------------------
    ' 利用用戶號碼，從當月帳款中找到總管理費，填入繳費金額
    ' ---------------------------
    For Each rngCell In auResid
        sResid = rngCell.Value
        iRow = rngCell.Row
        matchResult = Application.Match(sResid, billSheet.Range("A:A"), 0)
        If IsNumeric(matchResult) Then
            sumFee = billSheet.Range("D" & matchResult).Value
            autoDeductSheet.Range("E" & iRow).Value = sumFee '繳款金額所在欄位
			autoDeductSheet.Range("C" & iRow).Value = deductMonth '扣繳月份所在欄位
        Else
            MsgBox "自動扣繳用戶號碼" & sResid & "不正確"
        End If
    Next rngCell
    ' 委託扣款日 - disabled, filled at bank
    'matchResult = Application.Match("委託扣款日:", autoDeductSheet.Range("A:A"), 0)
    'If IsNumeric(matchResult) Then
    '    autoDeductSheet.Range("B" & matchResult).Value = _
    '        ActiveWorkbook.Sheets("基本資料").Range("B22").Value
    'Else
    '    MsgBox "找不到委託扣款日位置，請檢查"
    'End If
    
    ' 輸出xlsx檔
    Call exportSheetAsXlsx("自動扣繳", True)
	' 開啟自動扣繳，刪除繳費=0戶
	Call removeNoFeeDeduct("自動扣繳")
	Call genNotice
    ' 插入完成check
    insertCheck (ActiveWorkbook.Sheets("起始表").Range("G8"))
	Dim logFile As String
	logFile = ActiveWorkbook.Path & "\logfile.txt"
	Call appendLog(logFile, "Finished genAutoDeduct")

End Sub
