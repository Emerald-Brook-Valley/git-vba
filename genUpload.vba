'
' 產生上傳元大銀行易利收系統之Excel檔(10欄格式)
' 1. 僅更新B5的期別資訊，其餘表頭不變
' 2. 行7之標題不變
' 3. 當月收款從第8行開始填入，跳過總管理費=0之繳款戶
' 4. 跳過自動代繳戶
' 5. 新增拆單作業。將a.金額大於兩萬，b.戶別在拆單作業清單裡: 放置於另一個拆單上傳檔
' 6. 新增戶別清單，用來處理pdf檔案命名。(disabled, 改用上傳/拆單檔, 因與整份pdf match)
'
Sub genUpload()
	' 基本資料處理
	Dim basicSheet As Worksheet
    Set basicSheet = ActiveWorkbook.Sheets("基本資料")
    Dim paymentPeriod As String
    paymentPeriod = basicSheet.Range("B27").Value
	Dim uploadSheet As Worksheet
    Set uploadSheet = ActiveWorkbook.Sheets("上傳資料")
	Dim splitSheet As Worksheet
    Set splitSheet = ActiveWorkbook.Sheets("拆單上傳")
    uploadSheet.Range("B5").Value = paymentPeriod
	splitSheet.Range("B5").Value = paymentPeriod
    Dim superMarketDL As String
    superMarketDL = basicSheet.Range("B29").Value
    Dim superMarketExt As String
    superMarketExt = basicSheet.Range("B30").Value
	' 定義內容資料位置
    Dim billSheet As Worksheet
    Set billSheet = ActiveWorkbook.Sheets("當月帳款")
    Dim autoDeductSheet As Worksheet
    Set autoDeductSheet = ActiveWorkbook.Sheets("自動扣繳")
	Dim splitListSheet As Worksheet
    Set splitListSheet = ActiveWorkbook.Sheets("拆單作業")
    Dim nameSheet As Worksheet
    Set nameSheet = ActiveWorkbook.Sheets("區權人")
    Dim areaSheet As Worksheet
    Set areaSheet = ActiveWorkbook.Sheets("坪數")
    Dim rResid313 As Range
    Set rResid313 = billSheet.Range("A4:A316")
    Dim rngCell As Range
    Dim sResid As String
    Dim iRow As Integer
	Dim fillRow As Integer
    fillRow = 8 '填寫上傳資料起始位置
	Dim spliRow As Integer
    spliRow = 8 '填寫上傳資料起始位置
	' pdf命名用之戶別資料
	'Dim uploadListFile As String
	'Dim splitListFile As String
	'uploadListFile = ActiveWorkbook.Path & "\uploadList.txt"
	'splitListFile = ActiveWorkbook.Path & "\splitList.txt"

    ' ---------------------------
    ' 尋找自動扣繳戶別定義區間
    ' ---------------------------
    Dim auFirst As Integer
    Dim auLast As Integer
    auFirst = FindFirstResid(autoDeductSheet.Range("D:D"))
    auLast = FindLastResid(autoDeductSheet.Range("D:D"), auFirst)
    Dim auResid As Range
    Set auResid = autoDeductSheet.Range("D" & auFirst & ":D" & auLast)
	
    ' ---------------------------
    ' 尋找拆單戶定義區間
    ' ---------------------------
	Dim slExist As Boolean
    Dim slFirst As Integer
    Dim slLast As Integer
	Dim slResid As Range
	If isResid(splitListSheet.Range("A4").Value) Then
		slExist = True
		slFirst = FindFirstResid(splitListSheet.Range("A:A"))
		slLast = FindLastResid(splitListSheet.Range("A:A"), slFirst)
		Set slResid = splitListSheet.Range("A" & slFirst & ":A" & slLast)
	Else
		slExist = False
		slFirst = 4
		slLast = 4
		Set slResid = splitListSheet.Range("A" & slFirst & ":A" & slLast)
	End If
    
	' 加上progress bar
	UserForm1.Show vbModeless
    For Each rngCell In rResid313
        sResid = rngCell.Value
        iRow = rngCell.Row
        If billSheet.Range("D" & iRow).Value > 0 Then '繳費大於0元
            If Not IsNumeric(Application.Match(sResid, auResid, 0)) Then '非自動扣繳
				If billSheet.Range("D" & iRow).Value > 20000 Then '繳費大於20,000元
					Call writeUploadCell(iRow, spliRow, sResid, superMarketDL, superMarketExt, paymentPeriod, _
						splitSheet, billSheet, basicSheet, nameSheet, areaSheet)
					'Call AppendWithFSO(splitListFile, sResid & ",")
					spliRow = spliRow + 1
				ElseIf slExist AND IsNumeric(Application.Match(sResid, slResid, 0)) Then '位於拆單作業名單中
					Call writeUploadCell(iRow, spliRow, sResid, superMarketDL, superMarketExt, paymentPeriod, _
						splitSheet, billSheet, basicSheet, nameSheet, areaSheet)
					'Call AppendWithFSO(splitListFile, sResid & ",")
					spliRow = spliRow + 1
				Else
					Call writeUploadCell(iRow, fillRow, sResid, superMarketDL, superMarketExt, paymentPeriod, _
						uploadSheet, billSheet, basicSheet, nameSheet, areaSheet)
					'Call AppendWithFSO(uploadListFile, sResid & ",")
					fillRow = fillRow + 1
				End If
            End If '非自動扣繳
        End If '繳費大於0元
		'依照iRow位置顯示progress bar
		UserForm1.Label2.Width = UserForm1.Label1.Width*iRow/316
		UserForm1.Label3.Caption = format(iRow/316*100, "0.0") & "%完成"
		DoEvents
    Next rngCell
	' 清除progress bar
	Unload UserForm1
	'Beep
	
    ' 輸出xlsx檔
    Call exportSheetAsXls97("上傳資料", True)
	' 檢查拆單上傳是否有資料
	If isResid(ActiveWorkbook.Sheets("拆單上傳").Range("A8").Value) Then
		Call exportSheetAsXls97("拆單上傳", True)
	End If
    ' 插入完成check
    insertCheck (ActiveWorkbook.Sheets("起始表").Range("G7"))
	Dim logFile As String
	logFile = ActiveWorkbook.Path & "\logfile.txt"
	Call appendLog(logFile, "Finished genUpload")
    
End Sub
