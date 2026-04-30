Option Explicit
Public Const comName As String = "翡翠流域"
'Public logFile As String
'Public logNum As Integer
Sub ConfirmNew()
    ' 將工作目錄轉移到Workbook開啟位置
    ChDrive Left(ActiveWorkbook.Path, 1)
    ChDir ActiveWorkbook.Path
    ' 清理workbook舊資料
    ActiveWorkbook.Sheets("當月帳款").Range("J4:P316").ClearContents '行政規費之後
    ActiveWorkbook.Sheets("當月帳款").Range("E4:E316").ClearContents '清除前期欠款
	ActiveWorkbook.Sheets("當月帳款").Range("F4:F316").Value = 0 '清除當期欠款期數
    ActiveWorkbook.Sheets("上傳資料").Range("A8:U350").ClearContents '清除上傳資料
	ActiveWorkbook.Sheets("拆單上傳").Range("A8:U350").ClearContents '清除拆單上傳
    
	' 清除上傳資料與自動扣繳檔案
	Dim logFile As String
	logFile = ActiveWorkbook.Path & "\logfile.txt"
	Call killingFile(ActiveWorkbook.Path & "\上傳資料*.xls", logFile)
	Call killingFile(ActiveWorkbook.Path & "\拆單上傳*.xls", logFile)
	Call killingFile(ActiveWorkbook.Path & "\自動扣繳*.xlsx", logFile)
	Call killingFile(ActiveWorkbook.Path & "\代扣通知*.xlsx", logFile)
	Call killingFile(ActiveWorkbook.Path & "\公告管理費未繳名單*.docm", logFile)
	'Call killingFile(ActiveWorkbook.Path & "\pdf\*.pdf", logFile)
	' normal kill not working for read-only pdf file, do filesystem object method
	Dim fso As Object
	Set fso = CreateObject("Scripting.FileSystemObject")
	Dim pdfPath As String
	pdfPath = ActiveWorkbook.Path & "\pdf\*.pdf"
	If Dir(pdfPath) <> "" Then
		fso.DeleteFile (ActiveWorkbook.Path & "\pdf\*.pdf"), True
		Call appendLog(logFile, "deleted all pdf file under pdf dir")
	Else
		Call appendLog(logFile, "No pdf file under pdf dir")
	End If
	
    MsgBox ("準備完成。請進行下一步驟")
    insertCheck (ActiveWorkbook.Sheets("起始表").Range("G3"))
	Call appendLog(logFile, "工作電腦:" & GetNetworkPCName)
	Call appendLog(logFile, "Finished ConfirmNew")
End Sub

Sub qualYearSeason()
    Call qualResidYear
    Call qualResidSeason
    MsgBox "完成季、年繳資料檢查"
    insertCheck (ActiveWorkbook.Sheets("起始表").Range("G5"))
	Dim logFile As String
	logFile = ActiveWorkbook.Path & "\logfile.txt"
	Call appendLog(logFile, "Finished qualYearSeason")
End Sub

' ---------------------------
' 尋找第一個符合A0101 pattern的位置，回傳row
' ---------------------------
Public Function FindFirstResid(TargetRange As Range) As Integer
    Dim iFound As Boolean
    Dim i As Integer
    Dim sResident As String '戶名變數
    Dim iFirstRow
    
    ' 尋找第一筆符合A0101字樣的戶別資料
    iFound = False
    i = 1
    Do Until iFound = True Or i > 30 '從1往下找30行，找不到就放棄
        sResident = TargetRange(i).Value
        If isResid(sResident) Then
            iFirstRow = i
            iFound = True
        End If
        i = i + 1
    Loop

    If iFound = False Then
        'MsgBox "無住戶資料，請確認"
        FindFirstResid = 0
    Else
        FindFirstResid = iFirstRow
    End If
End Function

' ---------------------------
' 尋找最後一個符合A0101 pattern的位置，回傳row
' ---------------------------
Public Function FindLastResid(TargetRange As Range, iFirstRow As Integer) As Integer
    Dim sResid As String
    Dim iError As Boolean
    Dim i As Integer
    If iFirstRow = 0 Then '該表單無資料
        iError = True
        i = 2
    Else
        iError = False
        i = iFirstRow
    End If
    
    Do Until iError = True Or i > 350
        sResid = TargetRange(i).Value
        If IsEmpty(sResid) Then
            iError = True
        ElseIf Not isResid(sResid) Then
            iError = True
        End If
        i = i + 1
    Loop

    If iError = False Then
        MsgBox "戶數資料過長，請檢查"
    Else
        FindLastResid = i - 2
    End If
End Function

' ---------------------------
' 計算戶數
' ---------------------------
Public Function CountResid(TargetRange As Range, iStartRow As Integer) As Integer
    Dim iEmpty As Boolean
    iEmpty = False
    Dim i As Integer
    i = iStartRow
    Do Until iEmpty = True Or i > 350
        If IsEmpty(TargetRange(i).Value) Then
            iEmpty = True
        End If
        i = i + 1
    Loop

    If iEmpty = False Then
        MsgBox "戶數資料過長，請檢查"
    Else
        CountResid = i - iStartRow - 1
    End If
End Function

' 在指定位置插入check符號
Public Function insertCheck(rngCell As Range)
    With rngCell
        .Value = "P"
        .Font.Name = "Wingdings 2"
        .Font.Color = vbBlue
        .Font.Size = 24
    End With
End Function

' 寫入附加資料於logFile
Public Sub appendLog(logFile0 As String, newLog As String)
	Dim logNum As Integer
	logNum = FreeFile
	Open logFile0 For Append As #logNum
	Print #logNum, vbNewLine & Now & ":" & newLog
	Close #logNum
End Sub

' 回到起始表
Public Sub returnTop()
	Worksheets("起始表").Activate
End Sub

' 刪除檔案，可用wildcard
Public Sub killingFile(fileName As String, logFile0 As String)
	Dim fileDir As String
	On Error Resume Next
	fileDir = Dir(fileName)
    Do While fileDir <> ""
       Call appendLog(logFile0, "deleting " & fileDir)
	   SetAttr fileDir, vbNormal 'remove the read-only attribute
	   Kill(fileDir)
       ' Get the next file name matching the same pattern
       fileDir = Dir()
    Loop
End Sub

' 檢查5位數住戶id格式
Public Function isResid(sResident As String) As Boolean
    If Len(sResident) = 5 And Left(sResident, 1) Like "[A-H]" And IsNumeric(Right(sResident, 4)) Then
		isResid = True
    Else
		isResid = False
	End If
End Function

' append string to a file
Sub AppendWithFSO(filePath As String, sString As String)
    Dim fso As Object
    Dim ts As Object

    Set fso = CreateObject("Scripting.FileSystemObject")    
    ' OpenTextFile parameters: (FileName, IOMode, CreateIfNotExist)
    ' IOMode 8 is for Appending
    Set ts = fso.OpenTextFile(filePath, 8, True)
    ts.WriteLine sString
    ts.Close
End Sub

' 從上傳檔提出戶別清單，分割並命名pdf(python script)
Public Sub uploadList2pdf()
	Call res2csv()
	Call RunPythonInVenv()
End Sub

' get pc name 判斷python path
Public Function GetNetworkPCName() As String
    Dim objNetwork As Object
    Set objNetwork = CreateObject("WScript.Network")
    GetNetworkPCName = objNetwork.ComputerName
End Function
'
' 從上期公告未繳名單，讀取上期累計欠款期數
'
Sub updateSumDebt()
    Dim wsDebt As Worksheet
    Set wsDebt = ActiveWorkbook.Worksheets("公告未繳")
    Dim rngDebtResid As Range
    Dim rngDebtPeriod As Range
    Dim iDebtFirstRow As Integer
    Dim iDebtLastRow As Integer
    iDebtFirstRow = FindFirstResid(wsDebt.Range("A:A"))
	If iDebtFirstRow = 0 Then
		iDebtFirstRow = 4
		iDebtLastRow = 4
	Else
		iDebtLastRow = FindLastResid(wsDebt.Range("A:A"), iDebtFirstRow)
	End If
    Set rngDebtResid = wsDebt.Range("A" & iDebtFirstRow & ":A" & iDebtLastRow)
    Set rngDebtPeriod = wsDebt.Range("B" & iDebtFirstRow & ":B" & iDebtLastRow)
    
    Dim wsBill As Worksheet
    Set wsBill = ActiveWorkbook.Worksheets("當月帳款")
    Dim rngCell As Range
    Dim sResid As String
    Dim iRow As Integer
    Dim matchResult As Variant
    For Each rngCell In wsBill.Range("A4:A316")
        sResid = rngCell.Value
        iRow = rngCell.Row
        matchResult = Application.Match(sResid, rngDebtResid, 0)
        If IsNumeric(matchResult) Then
            wsBill.Range("G" & iRow).Value = rngDebtPeriod(matchResult).Value
        Else
            wsBill.Range("G" & iRow).Value = 0
        End If
    Next rngCell
	Dim logFile As String
	logFile = ActiveWorkbook.Path & "\logfile.txt"
	Call appendLog(logFile, "Finished updateSumDebt")
End Sub
'
' External Workbook to local sheet
' 使用 Application.FileDialog 方法，讓使用者自行選擇要讀取的檔案
' 將讀取的 Excel 檔案存到工作表，以進行後續操作
' 特殊化，針對未銷帳清單
'
Sub importDeficitFile()
    Dim targetWorkbook As Workbook
    Set targetWorkbook = Application.ActiveWorkbook
	Dim response As VbMsgBoxResult
	Dim logFile As String
	logFile = ActiveWorkbook.Path & "\logfile.txt"

    ' ---------------------------
    ' 開啟外部Excel檔案
    ' ---------------------------
    Dim fDialog As FileDialog
    Dim fileName As String
    MsgBox "請選擇上期未銷帳檔案"
    'using FileDialog for user to pick filename
    On Error Resume Next
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    'Show the dialog. -1 means success!
    If fDialog.Show = -1 Then
        Debug.Print fDialog.SelectedItems(1) 'The full path to the file selected by the user
        fileName = fDialog.SelectedItems(1)
		Call appendLog(logFile, "選擇未銷帳檔案" & fDialog.SelectedItems(1))
    Else
		MsgBox ("檔案" & fDialog.SelectedItems(1) & "開啟失敗，權限問題?請檢查後重試")
		Call appendLog(logFile, "檔案" & fDialog.SelectedItems(1) & "開啟失敗，權限問題?請檢查後重試")
		Exit Sub
    End If
    On Error GoTo 0

    ' ---------------------------
    ' 複製新開的未銷帳檔案內容，到目前檔案的"未銷帳"工作表
    ' ---------------------------
    Dim extBook As Workbook
    On Error Resume Next
    Set extBook = Application.Workbooks.Open(fileName, ReadOnly:=True)
    If extBook Is Nothing Or Err.Number <> 0 Then
        MsgBox ("檔案" & fDialog.SelectedItems(1) & "非Excel格式，請檢查後重試")
		Call appendLog(logFile, "檔案" & fDialog.SelectedItems(1) & "非Excel格式，請檢查後重試")
        extBook.Close SaveChanges:=False
        Exit Sub
    End If
    On Error GoTo 0
    If extBook.Sheets.Count > 1 Then
        MsgBox (fDialog.SelectedItems(1) & "不只一張工作表，請確認內容後重試")
		Call appendLog(logFile, fDialog.SelectedItems(1) & "不只一張工作表，請確認內容後重試")
        extBook.Close SaveChanges:=False
        Exit Sub
    End If

    Dim sourceSheet As Worksheet
    Set sourceSheet = extBook.Worksheets(1) '正常元大檔案只有一張工作表
	Dim targetSheet As Worksheet
    Set targetSheet = targetWorkbook.Sheets("未銷帳")

    ' ---------------------------
    ' 先清除舊資料，再複製內容
    ' ---------------------------
	targetSheet.Range("A1:J350").ClearContents
	Dim endRow As Integer
	' Note End xldown will jump to row larger than integer for single row
	If isResid(sourceSheet.Range("A2").Value) Then
		endRow = sourceSheet.Range("A:A").End(xlDown).Row
	Else
		endRow = 1
	End If
	MsgBox "新資料共" & endRow & "行"
	targetSheet.Columns("E").NumberFormat = "0"
    targetSheet.Range("A1", "J" & endRow).Value = sourceSheet.Range("A1", "J" & endRow).Value
    extBook.Close SaveChanges:=False

    ' 轉換內容為table
    Dim tblDeficit As ListObject
    Set tblDeficit = targetSheet.ListObjects.Add(xlSrcRange, targetSheet.Range("A1", "J" & endRow), , xlYes)
    tblDeficit.Name = "未銷帳戶"
    ' 檢查table內容是否符合
    Dim matchRowIndex1 As Variant
    Dim matchRowIndex2 As Variant
    matchRowIndex1 = Application.Match("繳款人識別碼", tblDeficit.HeaderRowRange, 0)
    matchRowIndex2 = Application.Match("繳款總額", tblDeficit.HeaderRowRange, 0)
    If IsError(matchRowIndex1) Or IsError(matchRowIndex2) Then
		Call appendLog(logFile, "未銷帳清單格式錯誤。未找到相關欄位")
		response = MsgBox("未找到相關欄位，請檢查銀行未銷帳清單是否正確。是否離開?", vbYesNo + vbQuestion, "未銷帳清單格式錯誤")
		If response = vbYes Then
			Exit Sub
		End If
    End If
	
	' 資料依照繳款人識別碼排序
	With tblDeficit.Sort
		.SortFields.Clear 'good practice to clear the sort field if there is any previous content
		.SortFields.Add Key:=Range("未銷帳戶[繳款人識別碼]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
		.Header = xlYes ' Assumes your table has a header row
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        ' Apply the sort
        .Apply
	End With
	
	' 合併同戶繳款金額
	Dim dataRow As Integer
	Dim idCol As ListColumn
    Dim feeCol As ListColumn
	Set idCol = tblDeficit.ListColumns("繳款人識別碼")
	Set feeCol = tblDeficit.ListColumns("繳款總額")
	Dim idThisRow As String
	If endRow > 2 Then '至少兩筆未銷帳
		For dataRow = 1 To endRow - 2
			idThisRow = idCol.DataBodyRange(dataRow).Value
			Do While idThisRow = idCol.DataBodyRange(dataRow+1).Value And isResid(idThisRow)
				Call appendLog(logFile, idThisRow & " in row " & dataRow+1 & " and " & dataRow+2)
				' 將戶號相同的兩行金額加到上一行
				feeCol.DataBodyRange(dataRow).Value = feeCol.DataBodyRange(dataRow).Value + feeCol.DataBodyRange(dataRow + 1).Value
				' 刪除下一行
				tblDeficit.DataBodyRange.Rows(dataRow+1).Delete
				'endRow = endRow - 1
			Loop
		Next dataRow
	End If
    ' ---------------------------
    ' 回到起始表
    ' ---------------------------
    targetWorkbook.Sheets("起始表").Activate
    MsgBox "未銷帳資料匯入完成"
    insertCheck (ActiveWorkbook.Sheets("起始表").Range("G4"))
	Call appendLog(logFile, "Finished importDeficitFile")
End Sub
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
'
' 更新季繳名單(M=1,4,7,10)
' 1. 檢查是否於年繳清單內
' 2. 檢查是否於未銷帳清單內
'
Public Sub qualResidSeason()
    Dim iMonth As Integer
    iMonth = ActiveWorkbook.Sheets("基本資料").Range("B16").Value
    Dim sFeePeriod As String
    sFeePeriod = ActiveWorkbook.Sheets("基本資料").Range("B27").Value
    Dim validMonth As Variant
    validMonth = Array(1, 4, 7, 10)
    Dim checkMonth As Variant
    checkMonth = Array(2, 5, 8, 11)
    Dim tblSeason As ListObject
    Set tblSeason = ActiveWorkbook.Sheets("季繳名單").ListObjects("季繳戶別")
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
    ' 更新季繳名單
    ' ---------------------------
    If IsNumeric(Application.Match(iMonth, validMonth, 0)) Then
        'check if list in deficit
        'yes, remove from list, add comment
        For Each rngCell In tblSeason.ListColumns("季繳戶").DataBodyRange
            sResid = rngCell.Value
            ' 檢查未銷帳清單
			'確認未銷帳非空
			If isResid(ActiveWorkbook.Sheets("未銷帳").Range("A2").Value) Then
				If IsNumeric(Application.Match(sResid, tblDeficit.ListColumns("繳款人識別碼").DataBodyRange, 0)) Then
					Call appendLog(logFile, "季繳戶" & sResid & " Found in Deficit. address=" & rngCell.Address)
					ActiveWorkbook.Sheets("季繳名單").Range("D6").Insert Shift:=xlDown
					ActiveWorkbook.Sheets("季繳名單").Range("D6").Value = sResid & "因尚有欠款，自季繳名單移除" & sFeePeriod
					rngCell.ClearContents
				End If
			End If
            ' 檢查年繳清單
            If IsNumeric(Application.Match(sResid, tblYear.ListColumns("年繳戶").DataBodyRange, 0)) Then
				Call appendLog(logFile, "季繳戶" & sResid & " Found in YearList. address=" & rngCell.Address)
                ActiveWorkbook.Sheets("季繳名單").Range("D6").Insert Shift:=xlDown
                ActiveWorkbook.Sheets("季繳名單").Range("D6").Value = sResid & "與年繳名單重複，自季繳名單移除" & sFeePeriod
                rngCell.ClearContents
            End If
        Next rngCell
        tblSeason.Sort.Apply
        ' 不符資格季繳戶清除完畢，更新季繳欄位
        For Each rngCell In rResid313
            If IsNumeric(Application.Match(rngCell.Value, tblSeason.ListColumns("季繳戶").DataBodyRange, 0)) Then
                ActiveWorkbook.Sheets("當月帳款").Range("H" & rngCell.Row).Value = 1
            Else
                ActiveWorkbook.Sheets("當月帳款").Range("H" & rngCell.Row).Value = 0
            End If
        Next rngCell
    ' ---------------------------
    ' 非季繳月份，名單須維持不變
    ' ---------------------------
    Else '非季繳處理月份，檢查清單是否與紀錄不符
        ' 是否有不當新增
        For Each rngCell In tblSeason.ListColumns("季繳戶").DataBodyRange
            sResid = rngCell.Value
            matchResult = Application.Match(sResid, rResid313, 0)
            If IsNumeric(matchResult) Then
                If ActiveWorkbook.Sheets("當月帳款").Range("H" & matchResult + 3).Value = 1 Then
                    'MsgBox sResid & "正確"
                Else
					Call appendLog(logFile, "季繳戶" & sResid & " 不在上月名單中，將被移除")
                    ActiveWorkbook.Sheets("季繳名單").Range("D6").Insert Shift:=xlDown
                    ActiveWorkbook.Sheets("季繳名單").Range("D6").Value = sResid & "非季繳申請期間請勿更新，將自季繳名單移除" & sFeePeriod
                    rngCell.ClearContents
                End If
            Else
				Call appendLog(logFile, "季繳存在無法識別的戶名" & sResid & "或空白")
                'Exit Sub
            End If
        Next rngCell
        ' 是否有不當刪除
        For Each rngCell In ActiveWorkbook.Sheets("當月帳款").Range("H4:H316")
            If rngCell.Value = 1 Then
                sResid = ActiveWorkbook.Sheets("當月帳款").Range("A" & rngCell.Row).Value
                If IsNumeric(Application.Match(sResid, tblSeason.ListColumns("季繳戶").DataBodyRange, 0)) Then
                    'MsgBox sResid & "正確"
                Else
					Call appendLog(logFile, "季繳戶" & sResid & " 不當移除，將被加回")
                    ActiveWorkbook.Sheets("季繳名單").Range("A4").Insert Shift:=xlDown
                    ActiveWorkbook.Sheets("季繳名單").Range("A4").Value = sResid
                    ActiveWorkbook.Sheets("季繳名單").Range("D6").Value = sResid & "非季繳申請期間請勿更新，將自季繳名單加回" & sFeePeriod
                End If
            End If
        Next rngCell
        tblSeason.Sort.Apply
    End If
    ' ---------------------------
    ' 季繳未繳處理
    ' ---------------------------
    If IsNumeric(Application.Match(iMonth, checkMonth, 0)) Then
        For Each rngCell In tblSeason.ListColumns("季繳戶").DataBodyRange
            sResid = rngCell.Value
            ' 檢查未銷帳清單
			'確認未銷帳非空
			If isResid(ActiveWorkbook.Sheets("未銷帳").Range("A2").Value) Then
				If IsNumeric(Application.Match(sResid, tblDeficit.ListColumns("繳款人識別碼").DataBodyRange, 0)) Then
					Call appendLog(logFile, "季繳戶" & sResid & " Found in Deficit. address=" & rngCell.Address)
					matchResult = Application.Match(sResid, rResid313, 0)
					ActiveWorkbook.Sheets("當月帳款").Range("H" & matchResult + 3).Value = -1
					ActiveWorkbook.Sheets("季繳名單").Range("D6").Insert Shift:=xlDown
					ActiveWorkbook.Sheets("季繳名單").Range("D6").Value = sResid & "因季繳未繳，改為月繳戶並加計欠款一期" & sFeePeriod
					rngCell.ClearContents
				End If
			End If
        Next rngCell
        tblSeason.Sort.Apply
    End If
End Sub
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
'
' 將單張工作表另存xlsx檔案
' 第2參數決定是否在檔名加上日期時間
'
Sub exportSheetAsXlsx(sheetName As String, qDateAppend As Boolean)
    Dim CurrentSheet As Worksheet
    Dim SavePath As Variant
    Dim fileName As String
    Set CurrentSheet = ActiveWorkbook.Sheets(sheetName)
	' 將sheet複製到剪貼簿
    CurrentSheet.Copy
    
    ' The newly created workbook becomes the ActiveWorkbook
	Dim NewWB As Workbook
    Set NewWB = ActiveWorkbook
	If qDateAppend Then
		fileName = CurrentSheet.Name & "_" & Format(Now(), "yyyy-mm-dd_hh-mm") & ".xlsx"
	Else
		fileName = CurrentSheet.Name
	End If
    ' Prompt the user for a save location and file name
    SavePath = Application.GetSaveAsFilename(InitialFileName:=fileName, _
                    FileFilter:="Excel Workbook (*.xlsx), *.xlsx", _
                    Title:="Select Save Location and File Name")
                                            
    ' Check if the user cancelled the dialog box
    If SavePath <> False Then
        ' Save the new workbook in the specified format and path
        Application.DisplayAlerts = False ' Suppress the macro-free warning message
        NewWB.SaveAs fileName:=SavePath, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
        Application.DisplayAlerts = True
        
        ' Close the new workbook
        NewWB.Close SaveChanges:=False ' Changes are already saved by SaveAs
        MsgBox sheetName & "成功輸出到: " & vbCrLf & SavePath, vbInformation, "Export Successful"
		Dim logFile As String
		logFile = ActiveWorkbook.Path & "\logfile.txt"
		Call appendLog(logFile, "輸出檔案" & SavePath)
    Else
        ' If the user cancels, close the new workbook without saving
        NewWB.Close SaveChanges:=False
        MsgBox "Export cancelled by user.", vbExclamation, "Cancelled"
    End If
End Sub
'
' 將單張工作表另存xls檔案Excel 97-2003
' 第2參數決定是否在檔名加上日期時間
'
Sub exportSheetAsXls97(sheetName As String, qDateAppend As Boolean)
    Dim CurrentSheet As Worksheet
    Dim SavePath As Variant
    Dim fileName As String
    Set CurrentSheet = ActiveWorkbook.Sheets(sheetName)
	' 將sheet複製到剪貼簿
    CurrentSheet.Copy
    
    ' The newly created workbook becomes the ActiveWorkbook
	Dim NewWB As Workbook
    Set NewWB = ActiveWorkbook
	If qDateAppend Then
		fileName = CurrentSheet.Name & "_" & Format(Now(), "yyyy-mm-dd_hh-mm") & ".xls"
	Else
		fileName = CurrentSheet.Name
	End If
    ' Prompt the user for a save location and file name
    SavePath = Application.GetSaveAsFilename(InitialFileName:=fileName, _
                    FileFilter:="Excel97 Workbook (*.xls), *.xls", _
                    Title:="Select Save Location and File Name")
                                            
    ' Check if the user cancelled the dialog box
    If SavePath <> False Then
        ' Save the new workbook in the specified format and path
        Application.DisplayAlerts = False ' Suppress the macro-free warning message
        NewWB.SaveAs fileName:=SavePath, FileFormat:=56, CreateBackup:=False
        Application.DisplayAlerts = False
        
        ' Close the new workbook
        NewWB.Close SaveChanges:=False ' Changes are already saved by SaveAs
        MsgBox sheetName & "成功輸出到: " & vbCrLf & SavePath, vbInformation, "Export Successful"
		Dim logFile As String
		logFile = ActiveWorkbook.Path & "\logfile.txt"
		Call appendLog(logFile, "輸出檔案" & SavePath)
    Else
        ' If the user cancels, close the new workbook without saving
        NewWB.Close SaveChanges:=False
        MsgBox "Export cancelled by user.", vbExclamation, "Cancelled"
    End If
End Sub
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
' write data of one Resid from other places to upload
Sub writeUploadCell( _
	iRow As Integer, _
	fillRow As Integer, _
	sResid As String, _
	superMarketDL As String, _
	superMarketExt As String, _
	paymentPeriod As String, _
	uploadSheet As Worksheet, _
	billSheet As Worksheet, _
	basicSheet As Worksheet, _
	nameSheet As Worksheet, _
	areaSheet As Worksheet)
	
    uploadSheet.Range("A" & fillRow).Value = sResid
    uploadSheet.Range("B" & fillRow).Value = nameSheet.Range("B" & iRow).Value
    uploadSheet.Range("C" & fillRow).Value = basicSheet.Range("B31").Value & _
        areaSheet.Range("E" & iRow).Value
    uploadSheet.Range("D" & fillRow).Value = billSheet.Range("D" & iRow).Value
    uploadSheet.Range("E" & fillRow).Value = superMarketDL
    uploadSheet.Range("F" & fillRow).Value = superMarketExt
    uploadSheet.Range("G" & fillRow).Value = paymentPeriod
    uploadSheet.Range("H" & fillRow).Value = billSheet.Range("L" & iRow).Value & _
        billSheet.Range("M" & iRow).Value
    uploadSheet.Range("I" & fillRow).Value = billSheet.Range("N" & iRow).Value & _
        billSheet.Range("O" & iRow).Value
    uploadSheet.Range("J" & fillRow).Value = billSheet.Range("B" & iRow).Value
    uploadSheet.Range("K" & fillRow).Value = billSheet.Range("C" & iRow).Value
	' L,M,N空白必須填0
	If billSheet.Range("E" & iRow).Value = "" Then
		uploadSheet.Range("L" & fillRow).Value = 0
	Else
		uploadSheet.Range("L" & fillRow).Value = billSheet.Range("E" & iRow).Value
	End If
	If billSheet.Range("J" & iRow).Value = "" Then
		uploadSheet.Range("M" & fillRow).Value = 0
	Else
		uploadSheet.Range("M" & fillRow).Value = billSheet.Range("J" & iRow).Value
	End If
	If billSheet.Range("K" & iRow).Value = "" Then
		uploadSheet.Range("N" & fillRow).Value = 0
	Else
		uploadSheet.Range("N" & fillRow).Value = billSheet.Range("K" & iRow).Value
	End If
	uploadSheet.Range("I" & fillRow).Value = billSheet.Range("P" & iRow).Value '備註
End Sub'
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
End Sub'
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
	
End Sub'
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
		
		wb.Close SaveChanges:=True
		' Get next file
        xlsxFile = Dir()
    Loop
	
End Sub'
' 從當月帳款更新公告未繳
'
Sub renewDebt()
    Dim wsDebt As Worksheet
    Set wsDebt = ActiveWorkbook.Worksheets("公告未繳")
    Dim iDebtFirstRow As Integer
    Dim iDebtLastRow As Integer
    iDebtFirstRow = FindFirstResid(wsDebt.Range("A:A"))
    iDebtLastRow = FindLastResid(wsDebt.Range("A:A"), iDebtFirstRow)
    ' 清除舊項目
    If iDebtFirstRow <> 0 Then '若非空名單
        wsDebt.Range("A" & iDebtFirstRow & ":D" & iDebtLastRow).ClearContents
    End If
    
    Dim rngCell As Range
    Dim wsBill As Worksheet
    Set wsBill = ActiveWorkbook.Worksheets("當月帳款")
    Dim sResid As String
    Dim iRow As Integer
    Dim iDebt As Long
    Dim iDebtRow As Integer
    iDebtRow = 4

    For Each rngCell In wsBill.Range("A4:A316")
        sResid = rngCell.Value
        iRow = rngCell.Row
        iDebt = wsBill.Range("E" & iRow)
        If iDebt > 15 Then '有欠款，金額15以下不計
            wsDebt.Range("A" & iDebtRow).Value = sResid
            wsDebt.Range("B" & iDebtRow).Value = wsBill.Range("F" & iRow).Value
            wsDebt.Range("C" & iDebtRow).Value = iDebt
            iDebtRow = iDebtRow + 1
        End If
    Next rngCell
End Sub
'
' Word公告管理費未繳名單
'
Sub ancmWordDoc()
	' 更新公告未繳
	Call renewDebt
    ' ---------------------------
    ' 開啟Word寫入檔案
    ' ---------------------------
    Dim wdApp As Word.Application
    Set wdApp = New Word.Application
    wdApp.Visible = True
    Dim wdDoc As Word.Document
    Dim sPeriod As String
    sPeriod = ActiveWorkbook.Worksheets("基本資料").Range("B27").Value
    Dim wordTemplate As String
    wordTemplate = ActiveWorkbook.Path & "\公告範本.dotx"
    Dim wordFile As String
    wordFile = ActiveWorkbook.Path & "\公告管理費未繳名單" & sPeriod & ".docm"
    Set wdDoc = wdApp.Documents.Add(Template:=wordTemplate)
    'wdDoc.Activate

    ' ---------------------------
    ' 工作表相關位置
    ' ---------------------------
    Dim wsDebt As Worksheet
    Set wsDebt = ActiveWorkbook.Worksheets("公告未繳")
    Dim iDebtFirstRow As Integer
    Dim iDebtLastRow As Integer
    iDebtFirstRow = FindFirstResid(wsDebt.Range("A:A"))
    iDebtLastRow = FindLastResid(wsDebt.Range("A:A"), iDebtFirstRow)
    Dim rDebt As Range
    If iDebtFirstRow = 0 Then
        MsgBox "本期無人欠繳，恭喜!"
    Else
        Set rDebt = wsDebt.Range("A" & iDebtFirstRow - 1 & ":D" & iDebtLastRow)
		rDebt.Copy
	
		' ---------------------------
		' 貼上並整理 word 表格
		' ---------------------------
		wdDoc.Paragraphs(3).Range.PasteExcelTable _
			LinkedToExcel:=False, WordFormatting:=True, RTF:=True
		Dim WordTable As Word.Table
		Set WordTable = wdDoc.Tables(1)
		WordTable.AllowAutoFit = True
		With WordTable
			.AutoFormat Format:=wdTableFormatElegant
			.AutoFitBehavior wdAutoFitWindow
			.Range.Font.Size = 14
			.Range.Font.Name = "微軟正黑體"
		End With
		WordTable.Rows.Alignment = wdAlignRowCenter
		WordTable.UpdateAutoFormat
		' ---------------------------
		' 儲存並關閉word檔案
		' ---------------------------
		wdDoc.SaveAs2 fileName:=wordFile ', FileFormat:=wdFormatXMLDocument
		'wdDoc.Close 'SaveChanges:=wdDoNotSaveChanges
		MsgBox "公告未繳名單已儲存於 " & wordFile	
	
	End If
    
    insertCheck (ActiveWorkbook.Sheets("起始表").Range("G9"))
	Dim logFile As String
	logFile = ActiveWorkbook.Path & "\logfile.txt"
	Call appendLog(logFile, "Finished ancmWordDoc")
End Sub
'
' Email 寄送
' 依照 Email 工作表內容寄送郵件
'
Public Sub sendMail()
    Dim mailResid As Range
	Dim wsMail As Worksheet
	Set wsMail = ActiveWorkbook.Worksheets("Email")
	Dim iMailFirstRow As Integer
	Dim iMailLastRow As Integer
	iMailFirstRow = FindFirstResid(wsMail.Range("A:A"))
	iMailLastRow = FindLastResid(wsMail.Range("A:A"), iMailFirstRow)
    Set mailResid = wsMail.Range("A" & iMailFirstRow & ":A" & iMailLastRow)
    Dim rngCell As Range
    Dim sResid As String
    Dim iRow As Integer
    Dim sSubject As String
    Dim sRecipients As String
	Dim sCC As String '如有多個CC mail中間以分號;隔開
    Dim sDetails As String
	Dim sAttachList As String
    Dim sAttachment As String
	Dim countMail As Integer
	countMail = 0
	Dim logFile As String
	logFile = ActiveWorkbook.Path & "\logfile.txt"
	'Dim matchedFiles As New Collection

	' 加上progress bar
	UserForm1.Show vbModeless
    For Each rngCell In mailResid
        sResid = rngCell.Value
        iRow = rngCell.Row
        sSubject = wsMail.Range("B" & iRow)
        sRecipients = wsMail.Range("C" & iRow)
		sCC = wsMail.Range("F" & iRow) '如有多個CC mail中間以分號;隔開
        sDetails = wsMail.Range("D" & iRow)
        'sAttachment = wsMail.Range("E" & iRow)
		'搜尋pdf子目錄，集合所有與戶名相符的pdf，作為附件
		'sAttachList = Dir(ActiveWorkbook.Path & "\pdf\" & sResid & "*.pdf")
		'sAttachment = ""
		If Dir(ActiveWorkbook.Path & "\pdf\" & sResid & "*.pdf") <> "" Then '檢查附件是否存在
			Call SendEmailUsingGmail(sSubject, sRecipients, sCC, sDetails, sResid)
			countMail = countMail + 1
		Else
			Call appendLog(logFile, sResid & " Email附檔不存在")
		End If
		'依照iRow位置顯示progress bar
		UserForm1.Label2.Width = UserForm1.Label1.Width*iRow/iMailLastRow
		UserForm1.Label3.Caption = format(iRow/iMailLastRow*100, "0.0") & "%完成"
		DoEvents
    Next rngCell
	' 清除progress bar
	Unload UserForm1	

    insertCheck (ActiveWorkbook.Sheets("起始表").Range("G10"))
	Call appendLog(logFile, "Finished sendMail count=" & countMail)
	MsgBox "Email共" & countMail & "封寄送完成"
End Sub

'For Early Binding, enable Tools > References > Microsoft CDO for Windows 2000 Library
Public Sub SendEmailUsingGmail(sSubject As String, sRecipients As String, sCC As String, sDetails As String, sResid As String)

	Dim NewMail As Object
	Dim mailConfig As Object
	Dim fields As Variant
	Dim msConfigURL As String
	Dim logFile As String
    logFile = ActiveWorkbook.Path & "\logfile.txt"
	Dim pdfPath As String
	pdfPath = ActiveWorkbook.Path & "\pdf\"
	Dim pdfFile As String
	pdfFile = Dir(pdfPath & sResid & "*.pdf")

	On Error GoTo Err:

	'late binding
	Set NewMail = CreateObject("CDO.Message")
	Set mailConfig = CreateObject("CDO.Configuration")
	'Set NewMail = New CDO.Message
	'Set mailconfig = New CDO.Configuration

	' load all default configurations
	mailConfig.Load -1

	Set fields = mailConfig.fields

	'Set All Email Properties
	With NewMail
		.From = "emeraldbrookvalley@gmail.com"
		.To = sRecipients
		.CC = sCC
		.BCC = ""
		.Subject = sSubject
		.TextBody = sDetails
		'.AddAttachment sAttachment
		Do While Len(pdfFile) > 0
			.AddAttachment pdfPath & pdfFile
			pdfFile = Dir ' Get next file
		Loop
	End With

	msConfigURL = "http://schemas.microsoft.com/cdo/configuration"


	With fields
		.Item(msConfigURL & "/smtpusessl") = True             'Enable SSL Authentication
		.Item(msConfigURL & "/smtpauthenticate") = 1          'SMTP authentication Enabled
		.Item(msConfigURL & "/smtpserver") = "smtp.gmail.com" 'Set the SMTP server details
		.Item(msConfigURL & "/smtpserverport") = 465          'Set the SMTP port Details
		.Item(msConfigURL & "/sendusing") = 2                 'Send using default setting
		.Item(msConfigURL & "/sendusername") = "emeraldbrookvalley" 'Your gmail address
		.Item(msConfigURL & "/sendpassword") = "wldiuoqwigvxkjkr" 'Your password or App Password
		.Update                                               'Update the configuration fields
	End With
	NewMail.Configuration = mailConfig
	NewMail.Send
	   
		'MsgBox "Your email has been sent", vbInformation

	'Exit_Err:
		'Release object memory
		'Set NewMail = Nothing
		'Set mailConfig = Nothing
		'End

	Err:
		Select Case Err.Number
			Case -2147220973  'Could be because of Internet Connection
				MsgBox "Check your internet connection." & vbNewLine & Err.Number & ": " & Err.Description
			Case -2147220975  'Incorrect credentials User ID or password
				MsgBox "Check your login credentials and try again." & vbNewLine & Err.Number & ": " & Err.Description '<- I'm getting to this error
			Case 0 'all is fine, do nothing
			Case Else   'Report other errors
				'MsgBox "Error encountered while sending email." & vbNewLine & Err.Number & ": " & Err.Description
				Call appendLog(logFile, "Error encountered while sending email." & vbNewLine & Err.Number & ": " & Err.Description)
		End Select

		'Resume Exit_Err

End Sub

'
' 取出上傳檔的住戶欄，存成csv檔
' 用來分離元大pdf繳費單
'
Public Sub res2csv()
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    Dim rng As Range
    Dim iFirstRow As Integer
    Dim iLastRow As Integer
    iFirstRow = FindFirstResid(Worksheets(1).Range("A:A"))
    If iFirstRow = 0 Then
        iFirstRow = 4
        iLastRow = 4
    Else
        iLastRow = FindLastResid(Worksheets(1).Range("A:A"), iFirstRow)
    End If
    Set rng = Worksheets(1).Range("A" & iFirstRow & ":A" & iLastRow)
    rng.Select
    Selection.Copy
    Workbooks.Add
    ActiveSheet.Paste
    
    ' Save the new workbook as a CSV file
    Dim listFile As String
    listFile = wb.Path & "\listFile.csv"
    ActiveWorkbook.SaveAs fileName:=listFile, FileFormat:=xlCSV, CreateBackup:=False
    ActiveWorkbook.Close
End Sub'
' 執行 python 程式來分割繳費單並冠上戶名
'
Public Sub RunPythonInVenv()
    Dim objShell As Object
    Dim pythonExe As String
	Dim pyScript As String
    Dim scriptPath As String
    Dim command As String
    Set objShell = VBA.CreateObject("WScript.Shell")
    ChDrive Left(ActiveWorkbook.Path, 1)
    ChDir ActiveWorkbook.Path
    
	' ---------------------------
    ' 開啟包含多戶的pdf繳費單
    ' ---------------------------
	Dim logFile As String
	logFile = ActiveWorkbook.Path & "\logfile.txt"
    Dim fDialog As FileDialog
    Dim fileName As String
    MsgBox "請選擇內含多戶的pdf繳費單"
    'using FileDialog for user to pick filename
    On Error Resume Next
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    'Show the dialog. -1 means success!
    If fDialog.Show = -1 Then
        Debug.Print fDialog.SelectedItems(1) 'The full path to the file selected by the user
        fileName = fDialog.SelectedItems(1)
		Call appendLog(logFile, "選擇內含多戶的pdf繳費單" & fDialog.SelectedItems(1))
    Else
		MsgBox ("檔案" & fDialog.SelectedItems(1) & "開啟pdf失敗，權限問題?請檢查後重試")
		Call appendLog(logFile, "檔案" & fDialog.SelectedItems(1) & "開啟pdf失敗，權限問題?請檢查後重試")
		Exit Sub
    End If
    On Error GoTo 0
	
    ' Path to the python.exe WITHIN your virtual environment folder
	Dim pcName As String
	pcName = GetNetworkPCName
	If pcName = "DESKTOP-P3U4FVV" Then
		pythonExe = "F:\anaconda3\python.exe"
	ElseIf pcName = "DESKTOP-TR30JJO" Then
		pythonExe = "C:\Users\user\AppData\Local\Programs\Python\Python313\python.exe"
	Else
		pythonExe = "C:\Users\user\AppData\Local\Programs\Python\Python313\python.exe"
	End If
    pyScript = ActiveWorkbook.Path & "\split_pdf.py"
    scriptPath = Chr(34) & pyScript & Chr(34)
    ' add /k if you want to pause the cmd window
    'command = "%comspec% /u " & pythonExe & """ """ & scriptPath
	'command = "%comspec% /u " & pythonExe & Chr(32) & scriptPath
    Dim wholePdf As String
	'wholePdf = "20260413144618.pdf"
	wholePdf = fileName
    ' Run the command (0 hides the window, True waits for it to finish)
    'objShell.Run command, 1, True
	Shell """" & pythonExe & """ """ & pyScript & """ """ & wholePdf & """", vbNormalFocus
    
    MsgBox "分戶繳費單製作完成，請見pdf目錄"
End Sub
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