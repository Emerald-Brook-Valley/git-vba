'
' Email 寄送
' 依照 Email 工作表內容寄送郵件
'
Public Sub sendMail()
    Dim mailResid As Range
	Dim wsMail As Worksheet
	Set wsMail = ActiveWorkbook.Worksheets("Email")
	Dim iMailFirstRow As Integer
	Dim iMailLastRow As Integer
	iMailFirstRow = FindFirstResid(wsMail.Range("A:A"))
	iMailLastRow = FindLastResid(wsMail.Range("A:A"), iMailFirstRow)
    Set mailResid = wsMail.Range("A" & iMailFirstRow & ":A" & iMailLastRow)
    Dim rngCell As Range
    Dim sResid As String
    Dim iRow As Integer
    Dim sSubject As String
    Dim sRecipients As String
	Dim sCC As String '如有多個CC mail中間以分號;隔開
    Dim sDetails As String
	Dim sAttachList As String
    Dim sAttachment As String
	Dim countMail As Integer
	countMail = 0
	Dim logFile As String
	logFile = ActiveWorkbook.Path & "\logfile.txt"
	'Dim matchedFiles As New Collection

	' 加上progress bar
	UserForm1.Show vbModeless
    For Each rngCell In mailResid
        sResid = rngCell.Value
        iRow = rngCell.Row
        sSubject = wsMail.Range("B" & iRow)
        sRecipients = wsMail.Range("C" & iRow)
		sCC = wsMail.Range("F" & iRow) '如有多個CC mail中間以分號;隔開
        sDetails = wsMail.Range("D" & iRow)
        'sAttachment = wsMail.Range("E" & iRow)
		'搜尋pdf子目錄，集合所有與戶名相符的pdf，作為附件
		'sAttachList = Dir(ActiveWorkbook.Path & "\pdf\" & sResid & "*.pdf")
		'sAttachment = ""
		If Dir(ActiveWorkbook.Path & "\pdf\" & sResid & "*.pdf") <> "" Then '檢查附件是否存在
			Call SendEmailUsingGmail(sSubject, sRecipients, sCC, sDetails, sResid)
			countMail = countMail + 1
		Else
			Call appendLog(logFile, sResid & " Email附檔不存在")
		End If
		'依照iRow位置顯示progress bar
		UserForm1.Label2.Width = UserForm1.Label1.Width*iRow/iMailLastRow
		UserForm1.Label3.Caption = format(iRow/iMailLastRow*100, "0.0") & "%完成"
		DoEvents
    Next rngCell
	' 清除progress bar
	Unload UserForm1	

    insertCheck (ActiveWorkbook.Sheets("起始表").Range("G10"))
	Call appendLog(logFile, "Finished sendMail count=" & countMail)
	MsgBox "Email共" & countMail & "封寄送完成"
End Sub

'For Early Binding, enable Tools > References > Microsoft CDO for Windows 2000 Library
Public Sub SendEmailUsingGmail(sSubject As String, sRecipients As String, sCC As String, sDetails As String, sResid As String)

	Dim NewMail As Object
	Dim mailConfig As Object
	Dim fields As Variant
	Dim msConfigURL As String
	Dim logFile As String
    logFile = ActiveWorkbook.Path & "\logfile.txt"
	Dim pdfPath As String
	pdfPath = ActiveWorkbook.Path & "\pdf\"
	Dim pdfFile As String
	pdfFile = Dir(pdfPath & sResid & "*.pdf")

	On Error GoTo Err:

	'late binding
	Set NewMail = CreateObject("CDO.Message")
	Set mailConfig = CreateObject("CDO.Configuration")
	'Set NewMail = New CDO.Message
	'Set mailconfig = New CDO.Configuration

	' load all default configurations
	mailConfig.Load -1

	Set fields = mailConfig.fields

	'Set All Email Properties
	With NewMail
		.From = "emeraldbrookvalley@gmail.com"
		.To = sRecipients
		.CC = sCC
		.BCC = ""
		.Subject = sSubject

		'force utf-8 encoding here
		.BodyPart.Charset = "utf-8"

		.TextBody = sDetails
		'.AddAttachment sAttachment
		Do While Len(pdfFile) > 0
			.AddAttachment pdfPath & pdfFile
			pdfFile = Dir ' Get next file
		Loop
	End With

	msConfigURL = "http://schemas.microsoft.com/cdo/configuration"


	With fields
		.Item(msConfigURL & "/smtpusessl") = True             'Enable SSL Authentication
		.Item(msConfigURL & "/smtpauthenticate") = 1          'SMTP authentication Enabled
		.Item(msConfigURL & "/smtpserver") = "smtp.gmail.com" 'Set the SMTP server details
		.Item(msConfigURL & "/smtpserverport") = 465          'Set the SMTP port Details
		.Item(msConfigURL & "/sendusing") = 2                 'Send using default setting
		.Item(msConfigURL & "/sendusername") = "emeraldbrookvalley" 'Your gmail address
		.Item(msConfigURL & "/sendpassword") = "wldiuoqwigvxkjkr" 'Your password or App Password
		.Update                                               'Update the configuration fields
	End With
	NewMail.Configuration = mailConfig
	NewMail.Send
	   
		'MsgBox "Your email has been sent", vbInformation

	'Exit_Err:
		'Release object memory
		'Set NewMail = Nothing
		'Set mailConfig = Nothing
		'End

	Err:
		Select Case Err.Number
			Case -2147220973  'Could be because of Internet Connection
				MsgBox "Check your internet connection." & vbNewLine & Err.Number & ": " & Err.Description
			Case -2147220975  'Incorrect credentials User ID or password
				MsgBox "Check your login credentials and try again." & vbNewLine & Err.Number & ": " & Err.Description '<- I'm getting to this error
			Case 0 'all is fine, do nothing
			Case Else   'Report other errors
				'MsgBox "Error encountered while sending email." & vbNewLine & Err.Number & ": " & Err.Description
				Call appendLog(logFile, "Error encountered while sending email." & vbNewLine & Err.Number & ": " & Err.Description)
		End Select

		'Resume Exit_Err

End Sub

