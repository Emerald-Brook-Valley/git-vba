'
' 執行 python 程式來合併pdf子目錄下的所有*.pdf檔為單一all.pdf
'
Public Sub mergePdfPython()
    Dim objShell As Object
    Dim pythonExe As String
	Dim pyScript As String
    Dim scriptPath As String
    Dim command As String
    Set objShell = VBA.CreateObject("WScript.Shell")
    ChDrive Left(ActiveWorkbook.Path, 1)
    ChDir ActiveWorkbook.Path
	Dim logFile As String
	logFile = ActiveWorkbook.Path & "\logfile.txt"
   	
    ' Path to the python.exe WITHIN your virtual environment folder
	Dim pcName As String
	pcName = GetNetworkPCName
	If pcName = "DESKTOP-P3U4FVV" Then
		pythonExe = "F:\anaconda3\python.exe"
	ElseIf pcName = "DESKTOP-TR30JJO" Then
		pythonExe = "C:\Users\user\AppData\Local\Programs\Python\Python313\python.exe"
	Else
		pythonExe = "python.exe"
	End If
    pyScript = ActiveWorkbook.Path & "\merge_pdf.py"
    scriptPath = Chr(34) & pyScript & Chr(34)
	Shell """" & pythonExe & """ """ & pyScript & """", vbNormalFocus
    
	Call appendLog(logFile, "mergePdfPython done")
    MsgBox "繳費單合併完成，存成all.pdf"
End Sub
