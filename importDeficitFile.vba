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
