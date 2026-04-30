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
