'
' Do not put this in modules
' It needs to be in VBAProject->Microsoft Excel Object->ThisWorkbook (double-click)
' select Workbook instead of general
'
Private Sub Workbook_Open()
	' public logFile not available outside Module, define locally
	Dim logFile As String
	logFile = ThisWorkbook.Path & "\logfile.txt"
	Call appendLog(logFile, "開啟" & ThisWorkbook.Name & "於" & ThisWorkbook.Path)
    MsgBox "目前工作目錄在" & ThisWorkbook.Path
    ' 清除check mark
    ThisWorkbook.Sheets("起始表").Range("G3:G10").ClearContents
End Sub
