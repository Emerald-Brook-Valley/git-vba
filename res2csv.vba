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
End Sub