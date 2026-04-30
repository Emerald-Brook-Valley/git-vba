'
' 產生自動代扣繳戶的通知單Excel, 以供word做合併
'
Sub genNotice()
    Dim billSheet As Worksheet
    Set billSheet = ActiveWorkbook.Sheets("當月帳款")
    Dim autoDeductSheet As Worksheet
    Set autoDeductSheet = ActiveWorkbook.Sheets("自動扣繳")
    Dim noticeSheet As Worksheet
    Set noticeSheet = ActiveWorkbook.Sheets("代扣通知")
    Dim nameSheet As Worksheet
    Set nameSheet = ActiveWorkbook.Sheets("區權人")
    Dim landSheet As Worksheet
    Set landSheet = ActiveWorkbook.Sheets("坪數")
	Dim basicSheet As Worksheet
	Set basicSheet = ActiveWorkbook.Sheets("基本資料")


    Dim rngCell As Range 'used in autoDeductSheet
    Dim sResid As String
    Dim iRow As Integer 'for autoDeductSheet
	Dim iRowN As Integer 'for noticeSheet
	iRowN = 2 '代扣通知從第2行開始填寫
    ' ---------------------------
    ' 尋找自動扣繳戶別定義區間
    ' ---------------------------
    Dim auFirst As Integer
    Dim auLast As Integer
    auFirst = FindFirstResid(autoDeductSheet.Range("D:D"))
    auLast = FindLastResid(autoDeductSheet.Range("D:D"), auFirst)
    Dim auResid As Range
    Set auResid = autoDeductSheet.Range("D" & auFirst & ":D" & auLast)
    Dim matchResult As Variant '儲存match結果，有找到會是行數
    Dim sumFee As Long
    ' ---------------------------
    ' 利用用戶號碼，蒐集資料，填入代扣通知裡
    ' ---------------------------
	'先清除舊資料
	Dim endRow As Integer
	endRow = noticeSheet.Range("A:A").End(xlDown).Row
	noticeSheet.Range("A2:T" & endRow).ClearContents
    For Each rngCell In auResid
        sResid = rngCell.Value
        iRow = rngCell.Row
		noticeSheet.Range("A" & iRowN).Value = sResid '填入ResidentID
		matchResult = Application.Match(sResid, nameSheet.Range("A:A"), 0)
        noticeSheet.Range("B" & iRowN).Value = nameSheet.Range("B" & matchResult).Value '填入RecipientName
		matchResult = Application.Match(sResid, landSheet.Range("A:A"), 0)
        noticeSheet.Range("C" & iRowN).Value = basicSheet.Range("B31").Value & _
			landSheet.Range("E" & matchResult).Value '填入RecipientAddress
        noticeSheet.Range("E" & iRowN).Value = basicSheet.Range("B22").Value '填入DeductionDate
		noticeSheet.Range("F" & iRowN).Value = basicSheet.Range("B27").Value '填入feeMonth
		'填入元大帳號後四碼AccountLast4
		noticeSheet.Range("G" & iRowN).NumberFormat = "@" 'force to text format
		noticeSheet.Range("G" & iRowN).Value = Right(autoDeductSheet.Range("B" & iRow).Value, 4)
		' data from billSheet
        matchResult = Application.Match(sResid, billSheet.Range("A:A"), 0)
        noticeSheet.Range("D" & iRowN).Value = billSheet.Range("D" & matchResult).Value '填入feeTotal
		noticeSheet.Range("H" & iRowN).Value = billSheet.Range("B" & matchResult).Value '填入landFee
		noticeSheet.Range("I" & iRowN).Value = billSheet.Range("C" & matchResult).Value '填入carFee
		noticeSheet.Range("J" & iRowN).Value = billSheet.Range("E" & matchResult).Value '填入debtFee
		noticeSheet.Range("K" & iRowN).Value = billSheet.Range("J" & matchResult).Value '填入fineFee
		noticeSheet.Range("L" & iRowN).Value = billSheet.Range("K" & matchResult).Value '填入elseFee
		noticeSheet.Range("M" & iRowN).Value = billSheet.Range("L" & matchResult).Value '填入message1
		noticeSheet.Range("N" & iRowN).Value = billSheet.Range("M" & matchResult).Value '填入message2
		noticeSheet.Range("O" & iRowN).Value = billSheet.Range("N" & matchResult).Value '填入message3
		noticeSheet.Range("P" & iRowN).Value = billSheet.Range("O" & matchResult).Value '填入message4
		noticeSheet.Range("U" & iRowN).Value = billSheet.Range("P" & matchResult).Value '填入Notes
		'填寫pdf檔案名稱與位置
		noticeSheet.Range("Q" & iRowN).Value = ActiveWorkbook.Path & "\pdf"
		noticeSheet.Range("R" & iRowN).Value = sResid
		noticeSheet.Range("S" & iRowN).Value = ActiveWorkbook.Path & "\pdf"
		noticeSheet.Range("T" & iRowN).Value = sResid
		iRowN = iRowN + 1
    Next rngCell
	
	Call exportSheetAsXlsx("代扣通知", False)
	' 開啟代扣通知，刪除繳費=0戶
	Call removeNoFee("代扣通知")
	Dim logFile As String
	logFile = ActiveWorkbook.Path & "\logfile.txt"
	Call appendLog(logFile, "Finished genNotice total record " & iRowN-2)
End Sub