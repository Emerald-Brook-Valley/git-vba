'
' 從當月帳款更新公告未繳
'
Sub renewDebt()
    Dim wsDebt As Worksheet
    Set wsDebt = ActiveWorkbook.Worksheets("公告未繳")
    Dim iDebtFirstRow As Integer
    Dim iDebtLastRow As Integer
    iDebtFirstRow = FindFirstResid(wsDebt.Range("A:A"))
    iDebtLastRow = FindLastResid(wsDebt.Range("A:A"), iDebtFirstRow)
    ' 清除舊項目
    If iDebtFirstRow <> 0 Then '若非空名單
        wsDebt.Range("A" & iDebtFirstRow & ":D" & iDebtLastRow).ClearContents
    End If
    
    Dim rngCell As Range
    Dim wsBill As Worksheet
    Set wsBill = ActiveWorkbook.Worksheets("當月帳款")
    Dim sResid As String
    Dim iRow As Integer
    Dim iDebt As Long
    Dim iDebtRow As Integer
    iDebtRow = 4

    For Each rngCell In wsBill.Range("A4:A316")
        sResid = rngCell.Value
        iRow = rngCell.Row
        iDebt = wsBill.Range("E" & iRow)
        If iDebt > 15 Then '有欠款，金額15以下不計
            wsDebt.Range("A" & iDebtRow).Value = sResid
            wsDebt.Range("B" & iDebtRow).Value = wsBill.Range("F" & iRow).Value
            wsDebt.Range("C" & iDebtRow).Value = iDebt
            iDebtRow = iDebtRow + 1
        End If
    Next rngCell
End Sub
