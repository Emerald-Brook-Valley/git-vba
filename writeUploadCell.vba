' write data of one Resid from other places to upload
Sub writeUploadCell( _
	iRow As Integer, _
	fillRow As Integer, _
	sResid As String, _
	superMarketDL As String, _
	superMarketExt As String, _
	paymentPeriod As String, _
	uploadSheet As Worksheet, _
	billSheet As Worksheet, _
	basicSheet As Worksheet, _
	nameSheet As Worksheet, _
	areaSheet As Worksheet)
	
    uploadSheet.Range("A" & fillRow).Value = sResid
    uploadSheet.Range("B" & fillRow).Value = nameSheet.Range("B" & iRow).Value
    uploadSheet.Range("C" & fillRow).Value = basicSheet.Range("B31").Value & _
        areaSheet.Range("E" & iRow).Value
    uploadSheet.Range("D" & fillRow).Value = billSheet.Range("D" & iRow).Value
    uploadSheet.Range("E" & fillRow).Value = superMarketDL
    uploadSheet.Range("F" & fillRow).Value = superMarketExt
    uploadSheet.Range("G" & fillRow).Value = paymentPeriod
    uploadSheet.Range("H" & fillRow).Value = billSheet.Range("L" & iRow).Value & _
        billSheet.Range("M" & iRow).Value
    uploadSheet.Range("I" & fillRow).Value = billSheet.Range("N" & iRow).Value & _
        billSheet.Range("O" & iRow).Value
    uploadSheet.Range("J" & fillRow).Value = billSheet.Range("B" & iRow).Value
    uploadSheet.Range("K" & fillRow).Value = billSheet.Range("C" & iRow).Value
	' L,M,N空白必須填0
	If billSheet.Range("E" & iRow).Value = "" Then
		uploadSheet.Range("L" & fillRow).Value = 0
	Else
		uploadSheet.Range("L" & fillRow).Value = billSheet.Range("E" & iRow).Value
	End If
	If billSheet.Range("J" & iRow).Value = "" Then
		uploadSheet.Range("M" & fillRow).Value = 0
	Else
		uploadSheet.Range("M" & fillRow).Value = billSheet.Range("J" & iRow).Value
	End If
	If billSheet.Range("K" & iRow).Value = "" Then
		uploadSheet.Range("N" & fillRow).Value = 0
	Else
		uploadSheet.Range("N" & fillRow).Value = billSheet.Range("K" & iRow).Value
	End If
	uploadSheet.Range("I" & fillRow).Value = billSheet.Range("P" & iRow).Value '備註
End Sub