Option Explicit
'
' check whether the files in a range do exist
' if not, show the missing filenames
'
Sub checkMissingAttachment()
	Dim ws As Worksheet
	Set ws = ActiveSheet
	Dim iRowStart As Integer
	Dim iRowEnd As Integer
	iRowStart = FindFirstResid(ws.Range("A:A"))
	iRowEnd = FindLastResid(ws.Range("A:A"), iRowStart)
	Dim arr() As String
	Dim cell As range
	Dim count As Integer
	count = 0
	
	' loop all cells content with attachment filenames
	For Each cell In ws.Range("E" & iRowStart & ":E" & iRowEnd)
		If Dir(cell.Value) = "" Then
			ReDim Preserve arr(count)
			arr(count) = Right(cell.Value, 9)
			count = count + 1
		End If
	Next cell
	
	' join all missing filename to single String
	Dim sAllNames As String
	sAllNames = Join(arr, ", ")
	MsgBox("共缺少" & count & "個檔案" & vbCrLf & sAllNames)
	'MsgBox(sAllNames)
End Sub

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

' 檢查5位數住戶id格式
Public Function isResid(sResident As String) As Boolean
    If Len(sResident) = 5 And Left(sResident, 1) Like "[A-H]" And IsNumeric(Right(sResident, 4)) Then
		isResid = True
    Else
		isResid = False
	End If
End Function