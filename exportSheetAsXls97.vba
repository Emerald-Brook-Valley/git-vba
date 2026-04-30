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
