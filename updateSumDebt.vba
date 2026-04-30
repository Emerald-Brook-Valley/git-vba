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
