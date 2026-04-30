'
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
