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
