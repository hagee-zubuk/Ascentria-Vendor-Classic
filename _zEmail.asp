<%
DIM	z_SMTP_CONN, z_SMTP_From, z_SMTP_MailingID
z_SMTP_From = "language.services@thelanguagebank.org"
z_SMTP_MailingID = "VendorGen"
z_SMTP_CONN = "Provider=SQLOLEDB;Data Source=10.10.16.35;Initial Catalog=langbank;User Id=langbank;Password=lang#lang;"
z_SMTP_CONN = "Provider=SQLOLEDB;Data Source=ERNIE\SQLEXPRESS;Initial Catalog=langbank;Integrated Security=SSPI;"

DIM z_SMTPServer(1), z_SMTP_Port(1), z_SMTP_User(1), z_SMTP_Pass(1)
z_SMTPServer(0) = "smtp.socketlabs.com"
z_SMTP_Port(0) = 2525
z_SMTP_User(0) = "server3874"
z_SMTP_Pass(0) = "UO2CUSxat9ZmzYD7jkTB"
z_SMTPServer(1) = "smtp.mailgun.org"
z_SMTP_Port(1) = 587
z_SMTP_User(1) = "postmaster@alt.thelanguagebank.org"
z_SMTP_Pass(1) = "d53256ad805ddbcf269221d16db0f6d1"


Function zSendMessage(strTo, strBCC, strSubject, strMSG)
	'SEND EMAIL
	lngIdx = 0
	blnOK = False
	Set mlMail = zSetEmailConfig()
	If Left(Request.ServerVariables("REMOTE_ADDR"), 11) = "192.168.111" Or _
			Left(Request.ServerVariables("REMOTE_ADDR"), 8) = "114.108." Or _
			Left(Request.ServerVariables("REMOTE_ADDR"), 8) = "127.0.0." Or _
			Left(Request.ServerVariables("REMOTE_ADDR"), 3) = "::1" Then 
		'mlMail.To = strTo
		mlMail.To = "hagee@zubuk.com"
		mlMail.Bcc = ""
	Else
		mlMail.To = strTo
		mlMail.Bcc = strBCC
	End If
	mlMail.From = z_SMTP_From
	mlMail.Subject= strSubject
	If (InStr(strMSG, "<html>")>0) Then
		strMSG = "<!doctype html><html lang=""en""><head><meta charset=""utf-8"">" & _
				"<title>" & strSubject & "</title><meta name=""description"" content=""Notification"">" & _
				"<meta name=""author"" content=""Language Services""></head><body>" & vbCrLf & _
				strMSG & vbCrLf & "</body></html>"
	End If
	mlMail.HTMLBody = strMSG
	lngRet = 0
	'mlMail.Configuration.Fields.Update
	mlMail.Fields.Item("urn:schemas:mailheader:X-xsMailingId")	= z_SMTP_MailingID
	mlMail.Fields.Item("urn:schemas:mailheader:MailingId")		= z_SMTP_MailingID
On Error Resume next
	mlMail.Send
	lngRet = Err.Number
On Error Goto 0

	If lngRet = 0 Then
		blnOK = zLogMailMessage(lngRet, mlMail.To, mlMail.Subject, z_SMTPServer(lngIdx), mlMail.HTMLBody, mlMail.Bcc)
		blnOK = True
	Else
		blnOK = zLogMailMessageRem(lngRet, mlMail.To, mlMail.Subject, z_SMTPServer(0) _
					, mlMail.HTMLBody, mlMail.Bcc _
					, "TOTAL FAILURE: " & z_SMTPServer(lngIdx))
		blnOK = True
	End If
	Set mlMail = Nothing
	zSendMessage = lngRet
End Function

Function zSetEmailConfig()
	Set mlMail = Server.CreateObject("CDO.Message")
	Set rsCnf = Server.CreateObject("ADODB.RecordSet")
	rsCnf.Open "SELECT TOP 1 * FROM [conf_email] ORDER BY [ord] ASC, [ts] DESC", z_SMTP_CONN, 1, 3
	If rsCnf.EOF Then
		lngIdx = 0
		With mlMail.Configuration.Fields
			.Item("http://schemas.microsoft.com/cdo/configuration/sendusing")			= 2
			.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver")			= z_SMTPServer(lngIdx)
			.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport")		= z_SMTP_Port(lngIdx)
			.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate")	= 1 'basic (clear-text) authentication
			.Item("http://schemas.microsoft.com/cdo/configuration/sendusername")		= z_SMTP_User(lngIdx)
			.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword")		= z_SMTP_Pass(lngIdx)
		End With
	Else
		With mlMail.Configuration.Fields
			.Item("http://schemas.microsoft.com/cdo/configuration/sendusing")			= 2
			.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate")	= 1 'basic (clear-text) authentication
			.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport")		= Z_CLng(rsCnf("port"))
			strTmp = Z_DoDecrypt(rsCnf("server"))
			If strTmp <> "" Then
				.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver")		= strTmp
			Else 
				.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver")		= Z_FixNull(rsCnf("server"))
			End If
			strTmp = Z_DoDecrypt(rsCnf("user"))
			If strTmp <> "" Then
				.Item("http://schemas.microsoft.com/cdo/configuration/sendusername")	= strTmp
			Else	
				.Item("http://schemas.microsoft.com/cdo/configuration/sendusername")	= Z_FixNull(rsCnf("user"))
			End If
			strTmp = Z_DoDecrypt(rsCnf("pass"))
			If strTmp <> "" Then
				.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword")	= strTmp
			Else
				.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword")	= Z_FixNull(rsCnf("pass"))
			End If
		End With
	End If
	mlMail.Configuration.Fields.Update
	rsCnf.Close
	Set rsCnf = Nothing
	Set zSetEmailConfig = mlMail
End Function

Function zLogMailMessage(lngerr, strto, subject, smtp, body, cc)
	Set rsLog = Server.CreateObject("ADODB.RecordSet")
	rsLog.Open "SELECT TOP 1 * FROM [log_email]", z_SMTP_CONN, 1, 3
	rsLog.AddNew
On Error Resume Next
	rsLog("err") = lngerr
	rsLog("subject") = Left(subject, 200)
	rsLog("org") = Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL")
	rsLog("body") = body
	rsLog("smtp") = smtp
	rsLog("to") = Left(strto, 100)
	rsLog("cc") = cc	
	rsLog.Update
	rsLog.Close
	Set rsLog = Nothing
	zLogMailMessage = True	
	Exit Function
End Function

Function zLogMailMessageRem(lngerr, strto, subject, smtp, body, cc, remk)
	Set rsLog = Server.CreateObject("ADODB.RecordSet")
	rsLog.Open "SELECT TOP 1 * FROM [log_email]", z_SMTP_CONN, 1, 3
	rsLog.AddNew
On Error Resume Next
	rsLog("err") = lngerr
	rsLog("rem") = remk
	rsLog("subject") = subject
	rsLog("org") = Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL")
	rsLog("body") = body
	rsLog("smtp") = smtp
	rsLog("to") = strto
	rsLog("cc") = cc
	rsLog.Update
	rsLog.Close
	Set rsLog = Nothing
	zLogMailMessageRem = True
	Exit Function
End Function

Function zMsgsLastHour(smtp)
	Set rsLog = Server.CreateObject("ADODB.RecordSet")
	dtLsHour   = DateAdd("h", -1, Now)
	strLsHour  = DatePart("yyyy", dtLsHour) & "-" & DatePart("m", dtLsHour) & "-" & _
			DatePart("d", dtLsHour) & " " & FormatDateTime(dtLsHour, 4)

	strSQL = "EXEC [dbo].[CountMessages] '" & strLsHour & "', '" & smtp & "'"

	rsLog.Open strSQL, z_SMTP_CONN, 3, 1
	If rsLog.EOF Then
		zMsgsLastHour = 0
	Else
		zMsgsLastHour = Z_CLng(rsLog("msgs"))
	End If
	rsLog.Close
	Set rsLog = Nothing
End Function

Function zMsgsLastMonth(smtp)
	Set rsLog = Server.CreateObject("ADODB.RecordSet")
	dtLsMonth  = DateAdd("m", -1, Date)
	strLsMonth = DatePart("yyyy", dtLsMonth) & "-" & DatePart("m", dtLsMonth) & "-15"

	strSQL = "EXEC [dbo].[CountMessages] '" & strLsMonth & "', '" & smtp & "'"
	rsLog.Open strSQL, z_SMTP_CONN, 3, 1
	If rsLog.EOF Then
		zMsgsLastMonth = 0
	Else
		zMsgsLastMonth = Z_CLng(rsLog("msgs"))
	End If
	rsLog.Close
	Set rsLog = Nothing
End Function

Function zGetInterpreterEmailByID(xxx)
	zGetInterpreterEmailByID = ""
	Set rsEm = Server.CreateObject("ADODB.RecordSet")
	sqlEm = "SELECT [e-mail] FROM interpreter_T WHERE [index] = " & xxx
	rsEm.Open sqlEm, z_SMTP_CONN, 1, 3
	If Not rsEm.EOF Then
		zGetInterpreterEmailByID = rsEm("e-mail")
	End If
	rsEm.Close
	Set rsEm = Nothing
End Function
%>