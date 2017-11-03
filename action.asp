<%Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_UtilCalendar.asp" -->
<!-- #include file="_Security.asp" -->
<%
Function FileUpload(xxx)
	FileUpload = False
	If xxx = "" Then Exit Function
	Set fso = CreateObject("Scripting.FileSystemObject")
	If fso.FileExists(F604AStr & xxx & ".PDF") Then FileUpload = True
	Set fso = Nothing
End Function
Function GetPrime2(xxx)
	GetPrime2 = ""
	Set rsRP = Server.CreateObject("ADODB.RecordSet")
	sqlRP = "SELECT * FROM interpreter_T WHERE [index] = " & xxx
	rsRP.Open sqlRP, g_strCONNLB, 3, 1
	If Not rsRP.EOF Then
		If rsRP("prime") = 0 Then
			GetPrime2 = rsRP("E-mail")
		ElseIf rsRP("prime") = 1 Or rsRP("prime") = 2 Then
			GetPrime2 = rsRP("E-mail")
		ElseIf rsRP("prime") = 3 Then
			GetPrime2 = CleanFax(Trim(rsRP("Fax"))) & "@emailfaxservice.com" 
		End If
	End If
	rsRP.Close
	set rsRP = Nothing
End Function
Function MyLang(xxx)
	MyLang = 0
	Set rsLang = Server.CreateObject("ADODB.RecordSet")
	sqlLang = "SELECT [index] FROM lang_T WHERE Upper(Lang) = '" & xxx & "'"
	rsLang.Open sqlLang, g_strCONN, 1, 3
	If Not rsLang.EOF Then
		MyLang = rsLang("index")
	End If
	rsLang.Close
	Set rsLAng = Nothing
End Function
Function CleanMe(xxx)
	CleanMe = xxx
	If Not IsNull(xxx) Or xxx <> "" Then CleanMe = Replace(xxx, ",", "''")
End Function
Function LangName(xxx)
	If xxx = "" Then Exit Function
	Set rsLang = Server.CreateObject("ADODB.RecordSet")
	sqlLang = "SELECT lang FROM lang_T WHERE [index] = " & xxx
	rsLang.Open sqlLang, g_strCONN, 1, 3
	If Not rsLang.EOF Then
		LangName = rsLang("lang")
	End If
	rsLang.Close
	Set rsLAng = Nothing
End Function
Function Z_FormatTime(xxx)
	Z_FormatTime = Null
	If xxx <> "" Or Not IsNull(xxx)  Then
		If IsDate(xxx) Then Z_FormatTime = FormatDateTime(xxx, 4) 
	End If
End Function
Function SalitaKo(strLang, IntrID)
	SalitaKo = False
	Set rsSalita = Server.CreateObject("ADODB.RecordSet")
	sqlSalita = "SELECT Language1, Language2, Language3, Language4, Language5 FROM interpreter_T WHERE [index] = " & IntrID 
	rsSalita.Open sqlSalita, g_strCONN, 1, 3
	If Not rsSalita.EOF Then
		If UCase(Trim(rsSalita("Language1"))) = Ucase(Trim(StrLang)) Then SalitaKo = True
		If UCase(Trim(rsSalita("Language2"))) = Ucase(Trim(StrLang)) Then SalitaKo = True
		If UCase(Trim(rsSalita("Language3"))) = Ucase(Trim(StrLang)) Then SalitaKo = True
		If UCase(Trim(rsSalita("Language4"))) = Ucase(Trim(StrLang)) Then SalitaKo = True
		If UCase(Trim(rsSalita("Language5"))) = Ucase(Trim(StrLang)) Then SalitaKo = True
	End If
	rsSalita.Close
	Set rsSalita = Nothing
End Function
Function GetLBLang(xxx)
	GetLBLang = -1
	Set rsLang  =Server.CreateObject("ADODB.RecordSet")
	sqlLang = "SELECT LBID FROM Lang_T WHERE [index] = " & xxx
	rsLang.Open sqlLang, g_strCONN, 3, 1
	If Not rsLang.EOF Then
		GetLBLang = rsLang("LBID")
	End If
	rsLang.Close
	Set rsLang = Nothing
End Function
If Request("ctrl") = 1 Then 'save new appointment
	tmpTS = Now
	'STORE ENTRIES ON COOKIE FOR EDITING AND SAVING ENTRIES
	Response.Cookies("LBREQUEST") = Z_DoEncrypt(Request("txtClilname")	& "|" & _
		Request("txtClifname")	& "|" & Request("selReas")	& "|" & Request("chkCall")	& "|" & Request("selDept")	& "|" & _
		Request("txtCliFon")	& "|" & Request("txtCliMobile")	& "|" & _
		Request("selLang") & "|" & Request("txtAppDate")	& "|" & Request("txtAppTFrom")	& "|" & Request("txtAppTTo") & "|" & Request("txtcom") & "|" & _
		Request("txtClinName")	& "|" & Request("txtTP") & "|" & Request("radioLang") & "|" & Request("chkminor") & "|" & Request("txtparents") & "|" & _
		Request("txtreqname") & "|" & Request("txtLBcom") & "|" & Request("txtIntrcom") & "|" & Request("selGen") & "|" & _
		Request("chkClientAdd")	& "|" & Request("txtCliAddrI") & "|" & Request("txtCliadd") & "|" & Request("txtClicity") & "|" & _
		Request("txtClistate")	& "|" & Request("txtCliZip") & "|" & Request("txtRFon") & "|" & Request("txtdhhsemail") & "|" & _
		Request("txtdhhsFon") & "|" & Request("txtoLang") & "|" & Request("chkCall2") & "|" & Request("txtchrg") & "|" & Request("txtAtrny") & "|" & _
		Request("txtDOB") & "|" & Request("txtPDamount") & "|" & Request("h_tmpfilename") & "|" & Request("chkout") & "|" & Request("chkmed") & "|" & _
		Request("MCnum") & "|" & Request("chkacc") & "|" & Request("chkcomp") & "|" & Request("selIns") & "|" & Request("txtemail") & "|" & _
		Request("MHPnum") & "|" & Request("NHHFnum") & "|" & Request("WSHPnum") & "|" & Request("chkawk")& "|" & Request("mrrec")& "|" & Request("chkleave") & "|" & Request("txtCliCir"))
	'CHECK VALID VALUES
	If Session("myClass") <> 3 Then
		If Request("txtTP") <> "" Then
			If Not IsDate(Request("txtTP")) Then Session("MSG") = Session("MSG") & "<br>ERROR: Invalid Time Paged."
		End If
	End If
	If Session("type") = 5 Then
		If Not IsNumeric(Request("txtPDamount")) Then
			Session("MSG") = Session("MSG") & "<br>ERROR: Invalid Amount requested from court."
		End If
	End If
	If Request("txtAppdate") <> "" Then
		If Not IsDate(Request("txtAppdate")) Then 
			Session("MSG") = Session("MSG") & "<br>ERROR: Invalid Appointment date."
		Else
			If cdate(Request("txtAppdate")) < Date Then Session("MSG") = Session("MSG") & "<br>ERROR: Invalid Appointment date."
		End If
	End If
	If Request("txtAppTFrom") <> "" Then
		If Not IsDate(Request("txtAppTFrom")) Then Session("MSG") = Session("MSG") & "<br>ERROR: Invalid Appointment Time (From:)."
	End If
	If Request("txtAppTTo") <> "" Then
		If Not IsDate(Request("txtAppTTo")) Then Session("MSG") = Session("MSG") & "<br>ERROR: Invalid Appointment Time (To:)."
	End If
	If Request("txtDOB") <> "" Then
		If Not IsDate(Request("txtDOB")) Then Session("MSG") = Session("MSG") & "<br>ERROR: Invalid Date of Birth."
	End If
	If Session("MSG") = "" Then	
		'GET COOKIE OF REQUEST
		tmpEntry = Split(Z_DoDecrypt(Request.Cookies("LBREQUEST")), "|")
		deptRate = GetDeptRate(tmpEntry(4))
		'SAVE ENTRIES
		Set rsMain = Server.CreateObject("ADODB.RecordSet")
		sqlMain = "SELECT * FROM appointment_T WHERE [timestamp] = '" & now & "'"
		rsMain.Open sqlMain, g_strCONN, 1, 3
		rsMain.AddNew
		rsMain("timestamp") = tmpTS
		rsMain("clname") = Z_DoEncrypt(CleanMe(tmpEntry(0)))
		rsMain("cfname") =Z_DoEncrypt( CleanMe(tmpEntry(1)))
		'rsMain("addr") = CleanMe(tmpEntry(2))
		'rsMain("city") = tmpEntry(3)
		'rsMain("state") = Ucase(tmpEntry(4))
		'rsMain("zip") = tmpEntry(5)
		rsMain("Phone") = Z_DoEncrypt(tmpEntry(5))
		rsMain("Mobile") = Z_DoEncrypt(tmpEntry(6))
		If tmpEntry(14) = 0 Then rsMain("langID") = tmpEntry(7)
		If tmpEntry(14) = 1 Then rsMain("langID") = MyLang("PORTUGUESE")
		If tmpEntry(14) = 2 Then rsMain("langID") = MyLang("SPANISH")
		If tmpEntry(14) = 3 Then rsMain("langID") = MyLang("SOMALI")
		If tmpEntry(14) = 4 Then rsMain("langID") = MyLang("AMERICAN SIGN LANGUAGE")
		If tmpEntry(14) = 5 Then rsMain("langID") = MyLang("OTHER")
		If tmpEntry(14) = 5 Then 
			rsMain("oLang") = tmpEntry(30)
		Else
			rsMain("oLang") = ""
		End If
		rsMain("appdate") = tmpEntry(8)
		rsMain("TimeFrom") = tmpEntry(8) & " " & Z_FormatTime(tmpEntry(9))
		rsMain("TimeTo") = tmpEntry(8) & " " & Z_FormatTime(tmpEntry(10))
		rsMain("IntrID") = -1
		rsMain("Comment") = tmpEntry(11)
		rsMain("DOB") = Z_DateNull(tmpEntry(34))
		If Session("myClass") <> 3 Then
			Pagedme = Empty
			If tmpEntry(13) <> "" Then Pagedme = tmpEntry(13)
			rsMain("paged") = Pagedme
			rsMain("clinician") = tmpEntry(12)
		Else
			rsMain("crtroom") = tmpEntry(13)
			rsMain("docknum") = tmpEntry(12)
			rsMain("attny") = tmpEntry(33)
			rsMain("charges") = tmpEntry(32)
			'PD
		End If
		If tmpEntry(2) = "" Then tmpEntry(2) = "0"
		rsMain("reason") = tmpEntry(2)
		rsMain("callme") = False
		If tmpEntry(3) <> "" Then rsMain("callme") = True
		rsMain("InstID") = Session("InstID")
		rsMain("minor") = False
		If tmpEntry(15) <> "" Then rsMain("minor") = True
		rsMain("parents") = tmpEntry(16)
		rsMain("reqName") = tmpEntry(17)
		rsMain("block") = false
		If tmpEntry(31) <> "" Then rsMain("block") = true 'tmpLBCOM = "BLOCK SCHEDULE  "
		rsMain("lbcom") = tmpEntry(18)
		rsMain("intrcom") = tmpEntry(19)
		rsMain("Gender") = tmpEntry(20)
		
		If tmpEntry(21) <> "" Then 
			rsMain("useCadr") = True
			rsMain("capt") = tmpEntry(22)
			rsMain("caddress") = tmpEntry(23)
			rsMain("ccity") = tmpEntry(24)
			rsMain("cstate") = tmpEntry(25)
			rsMain("czip") = tmpEntry(26)
		Else
			rsMain("useCadr") = False
			rsMain("capt") = ""
			rsMain("caddress") = ""
			rsMain("ccity") = ""
			rsMain("cstate") = ""
			rsMain("czip") = ""
		End If
		rsMain("semail") = tmpEntry(28)
		rsMain("sphone") = tmpEntry(29)
		rsMain("rphone") = tmpEntry(27)
		
		rsMain("DeptID") = tmpEntry(4)
		rsMain("ReqID") = Session("ReqID")
		rsMain("InstLB") = Session("InstID")
		If Session("type") = 5 Then
			rsMain("PDemail") = tmpEntry(43)
			rsMain("PDamount") = Z_CZero(tmpEntry(35))
			rsMain("UploadFile") = False
			If FileUpload(tmpEntry(36)) Then 
				rsMain("UploadFile") = True
				rsMain("filename") = tmpEntry(36) & ".PDF"
			End If
		End If
		rsMain("outpatient") = False
		If tmpEntry(37) <> "" Then rsMain("outpatient") = True
		rsMain("hasmed") = False
		If tmpEntry(38) <> "" Then rsMain("hasmed") = True
		rsMain("medicaid") = tmpEntry(39)
		rsMain("meridian") = tmpEntry(44)
		rsMain("nhhealth") = tmpEntry(45)
		rsMain("wellsense") = tmpEntry(46)
		rsMain("acknowledge") = false
		If tmpEntry(47) <> "" Then rsMain("acknowledge") = True
		rsMain("autoacc") = False
		If tmpEntry(40) <> "" Then rsMain("autoacc") = True
		rsMain("wcomp") = False
		If tmpEntry(41) <> "" Then rsMain("wcomp") = True
		rsMain("secins") = Z_FixNull(tmpEntry(42))
		rsMain("vermed") = False
		rsMain("mrrec") = tmpEntry(48)
		rsMain("leavemsg") = False
		If tmpEntry(49) <> "" Then rsMain("leavemsg") = True
		rsMain("Spec_cir") = tmpEntry(50)
		rsMain("UID") = Session("UID")
		rsMain.Update
		'GET ID FOR CONFIRM
		tmpID = rsMain("index")
		rsMain.Close
		Set rsMain = Nothing
		Call AddLog("Appointment " & tmpID & " saved in VENDOR DB.")
		If Session("type") = 5 AND FileUpload(tmpEntry(36)) Then 'save Form on DB
			Set rsFile = Server.CreateObject("ADODB.RecordSet")
			sqlFile = "SELECT * FROM pdf_T"
			rsFile.Open sqlFile, g_strCONN, 1, 3
			rsFile.AddNew
			rsFile("appID") = tmpID
			rsFile("filename") = tmpEntry(36) & ".PDF"
			rsFile("datestamp") = Now
			rsFile.Update
			rsFile.Close
			Set rsFile = Nothing
		End If
		'SAVE APPOINTMENT IN LANGUAGE BANK
		Set rsLB = Server.CreateObject("ADODB.RecordSet")
		sqlLB = "SELECT * FROM Request_T WHERE [timestamp] = '" & Now & "'"
		rsLB.Open sqlLB, g_strCONNLB, 1, 3
		rsLB.AddNew
		rsLB("HPID") = tmpID
		rsLB("timestamp") = tmpTS
		rsLB("reqID") = Session("ReqID")
		rsLB("appdate") =  tmpEntry(8)
		rsLB("appTimeFrom") =  tmpEntry(8) & " " & Z_FormatTime(tmpEntry(9))
		rsLB("appTimeTo") = tmpEntry(8) & " " & Z_FormatTime(tmpEntry(10))
		tmpEmerDateTime = CDate(tmpEntry(8) & " " & Z_FormatTime(tmpEntry(9)))
		If DateDiff("n", Now, tmpEmerDateTime) < 1440 Then rsLB("Emergency") = True
		If tmpEntry(14) = 0 Then rsLB("langID") = GetLBLang(tmpEntry(7))
		If tmpEntry(14) = 1 Then rsLB("langID") = GetLBLang(MyLang("PORTUGUESE"))
		If tmpEntry(14) = 2 Then rsLB("langID") = GetLBLang(MyLang("SPANISH"))
		If tmpEntry(14) = 3 Then rsLB("langID") = GetLBLang(MyLang("SOMALI"))
		If tmpEntry(14) = 4 Then rsLB("langID") = GetLBLang(MyLang("AMERICAN SIGN LANGUAGE"))
		If tmpEntry(14) = 5 Then rsLB("langID") = GetLBLang(MyLang("OTHER"))
		rsLB("clname") = CleanMe(tmpEntry(0))
		rsLB("cfname") = CleanMe(tmpEntry(1))
		rsLB("Cphone") = tmpEntry(5)
		rsLB("DOB") = Z_DateNull(tmpEntry(34))
		rsLB("InstID") = Session("InstID")
		rsLB("DeptID") = tmpEntry(4)
		rsLB("Comment") = tmpEntry(11)
		rsLB("IntrID") = -1
		rsLB("CAphone") =  tmpEntry(6)
		rsLB("LBcomment") = tmpEntry(18)
		rsLB("Intrcomment") = tmpEntry(19)
		rsLB("Gender") = tmpEntry(20)
		rsLB("Child") = False
		If tmpEntry(15) <> "" Then rsLB("Child") = True
		If tmpEntry(21) <> "" Then 
			rsLB("CliAdd") = True
			rsLB("CliAdrI") = tmpEntry(22)
			rsLB("caddress") = tmpEntry(23)
			rsLB("ccity") = tmpEntry(24)
			rsLB("cstate") = tmpEntry(25)
			rsLB("czip") = tmpEntry(26)
		Else
			rsLB("CliAdd") = False
			rsLB("CliAdrI") = ""
			rsLB("caddress") = ""
			rsLB("ccity") = ""
			rsLB("cstate") = ""
			rsLB("czip") = ""
		End If
		If Session("myClass") = 3 Then
			rsLB("DocNum") = tmpEntry(12)
			rsLB("CrtRumNum") = tmpEntry(13)
			rsLB("Comment") = tmpEntry(11) & vbCrlf & " Charge/s: " & tmpEntry(32) & vbCrlf & " Attorney: " & tmpEntry(33)
		End If
		rsLB("InstRate") = deptRate
		If Session("type") = 5 Then
			rsLB("DocNum") = tmpEntry(12)
			rsLB("PDamount") = Z_CZero(tmpEntry(35))
			If FileUpload(tmpEntry(36)) Then 
				rsLB("UploadFile") = True
				rsLB("filename") = tmpEntry(36) & ".PDF"
			End If
		End If
		rsLB("outpatient") = False
		If tmpEntry(37) <> "" Then rsLB("outpatient") = True
		rsLB("hasmed") = False
		If tmpEntry(38) <> "" Then rsLB("hasmed") = True
		rsLB("medicaid") = tmpEntry(39)
		rsLB("meridian") = tmpEntry(44)
		rsLB("nhhealth") = tmpEntry(45)
		rsLB("wellsense") = tmpEntry(46)
		rsLB("acknowledge") = false
		If tmpEntry(47) <> "" Then rsLB("acknowledge") = True
		rsLB("autoacc") = False
		If tmpEntry(40) <> "" Then rsLB("autoacc") = True
		rsLB("wcomp") = False
		If tmpEntry(41) <> "" Then rsLB("wcomp") = True
		rsLB("secins") = Z_FixNull(tmpEntry(42))
		rsLB("vermed") = False
		rsLB("mrrec") = tmpEntry(48)
		rsLB("blocksched") = False
		If tmpEntry(31) <> "" Then rsLB("blocksched") = True
		rsLB("leavemsg") = False
		If tmpEntry(49) <> "" Then rsLB("leavemsg") = True
		rsLB("Spec_cir") = tmpEntry(50)
		rsLB("courtcall") = false
		If tmpEntry(3) <> "" Then rsLB("courtcall") = True
		rsLB.update
		tmpLBID = rsLB("index")
		rsLB.Close
		Set rsLB = Nothing
		Call AddLog("Appointment " & tmpID & " saved in STAFF DB.")
		Call ActiveSage(tmpEntry(4))
		'GET INFO for EMAIL
		tmpInst = GetFacility(Session("InstID"))
		tmpDept = GetDept( tmpEntry(4))
		tmpCName = CleanMe(tmpEntry(0)) & ", " & CleanMe(tmpEntry(1))
		tmpFon = tmpEntry(5)
		tmpMobile = tmpEntry(6)
		tmpAppDate = tmpEntry(8)
		tmpAppTimeFrom = Z_FormatTime(tmpEntry(9))
		tmpAppTime = Z_FormatTime(tmpEntry(9))
		'mrrec = tmpEntry(48)
		If tmpEntry(10) <> "" Then tmpAppTime = tmpApptime & " - " & Z_FormatTime(tmpEntry(10))
		If tmpEntry(14) = 0 Then tmpLang = LangName(tmpEntry(7))
		If tmpEntry(14) = 1 Then tmpLang = "Portuguese"
		If tmpEntry(14) = 2 Then tmpLang = "Spanish"
		If tmpEntry(14) = 3 Then tmpLang = "Somali"
		If tmpEntry(14) = 4 Then tmpLang = "ASL"
		If tmpEntry(14) = 5 Then tmpLang = "Other - " & tmpEntry(30)
		tmpACom = CleanMe(tmpEntry(11))
		tmpLCom = CleanMe(tmpEntry(18))
		If tmpEntry(31) <> "" Then tmpLCom = "BLOCK SCHED" & vbCrLf & tmpLCom
		tmpICom = CleanMe(tmpEntry(19))
		'tmpReas = tmpEntry(2)
		tmpReas = GetReas(Z_Replace(tmpEntry(2),", ", "|"))
		tmpCall = ""
		If tmpEntry(3) <> "" Then tmpCall = "*Call patient to remind of appointment"
		'GET USER
		Set rsUser = Server.CreateObject("ADODB.RecordSet")
		sqlUSer = "SELECT lname, fname FROM User_T WHERE [index] = " & Session("UID")
		rsUser.Open sqlUser, g_strCONN, 3, 1
		If Not rsUser.EOF Then
			tmpUname = rsUser("lname") & ", " & rsUser("fname")
		End If
		rsUser.Close
		Set rsUSer = Nothing
		If Request("txtreqname") <> "" Then tmpUname = Request("txtreqname")
		'SAVE HISTORY IN LB
		TimeNow = Now
		Set rsHist = Server.CreateObject("ADODB.RecordSet")
		sqlHist = "SELECT * FROM History_T WHERE [index] = 0"
		rsHist.Open sqlHist, g_strCONNHist, 1,3 
		rsHist.AddNew
		rsHist("reqID") = tmpLBID
		rsHist("Creator") = Session("GreetMe")
		rsHist("date") = tmpEntry(8)
		rsHist("dateTS") = TimeNow
		rsHist("dateU") = Session("GreetMe")
		rsHist("Stime") = Z_dateNull(tmpEntry(8) & " " & tmpEntry(9))
		rsHist("StimeTS") = TimeNow
		rsHist("StimeU") = Session("GreetMe")
		rsHist("location") = "department address"
		rsHist("locationTS") = TimeNow
		rsHist("locationU") = Session("GreetMe")
		rsHist.Update
		rsHist.Close
		Set rsHist = Nothing
		Call AddLog("Appointment " & tmpID & " saved in HISTORY DB.")
		If SaveHist(tmpLBID, "[HP]main.asp") Then
	
		End If
		Call AddLog("Appointment " & tmpID & " saved in DETAILED HISTORY DB.")
		If tmpEntry(14) = 0 Then langsel = tmpEntry(7)
		If tmpEntry(14) = 1 Then langsel = MyLang("PORTUGUESE")
		If tmpEntry(14) = 2 Then langsel = MyLang("SPANISH")
		If tmpEntry(14) = 3 Then langsel = MyLang("SOMALI")
		If tmpEntry(14) = 4 Then langsel = MyLang("AMERICAN SIGN LANGUAGE")
		If tmpEntry(14) = 5 Then langsel = MyLang("OTHER")

		If GetLBLang(langsel) <> 52 And GetLBLang(langsel) <> 78 And GetLBLang(langsel) <> 81 And  GetLBLang(langsel) <> 90 And  GetLBLang(langsel) <> 85 And tmpEntry(14) <> 4 And tmpEntry(31) = "" Then
			call Z_EmailJob(tmpLBID)
		End If
		Call AddLog("Appointment " & tmpID & " EMAILS SENT TO INTERPRETERS.")
		'EMAIL TO LB STAFF
		'If Request("SubmitMe") = 1 Then
	'on error resume next
			strSubj = "Interpreter Request - " & tmpInst & " - " & tmpDept
			tmpAppDateTime = CDate(tmpAppDate & " " & tmpAppTimeFrom)
			If DateDiff("n", Now, tmpAppDateTime) < 1440 Then strSubj = "URGENT - Interpreter Request - " & tmpInst & " - " & tmpDept
			Set mlMail = CreateObject("CDO.Message")
	mlMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing")= 2
	mlMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "localhost"
	mlMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 26
	mlMail.Configuration.Fields.Update
			mlMail.To = "language.services@thelanguagebank.org"
			'mlMail.Bcc = "sysdump1@zubuk.com"
			mlMail.From = "language.services@thelanguagebank.org"
			mlMail.Subject = strSubj
			strBody = "<table cellpadding='0' cellspacing='0' border='0' align='center'>" & vbCrLf & _
					"<tr><td align='center'>" & vbCrLf & _
						"<img src='http://languagebank.lssne.org/lsslbis/images/LBISLOGOBandW.jpg'>" & vbCrLf & _
					"</td></tr>" & vbCrLf & _
					"<tr><td>&nbsp;</td></tr>" & vbCrLf & _	
					"<tr><td align='center'>" & vbCrLf & _
						"<font size='2' face='trebuchet MS'><b>Interpreter Request:</b></font><br>" & vbCrLf & _
					"</td></tr>" & vbCrLf & _
					"<tr><td>&nbsp;</td></tr>" & vbCrLf & _
					"<tr><td>" & vbCrLf & _
						"<table cellpadding='0' cellspacing='0' border='2' width='100%'>" & vbCrLf & _
							"<tr>" & vbCrLf & _
								"<td align='right' width='225px'>" & vbCrLf & _
									"<font size='2' face='trebuchet MS'>Institution - Department:</font><br>" & vbCrLf & _
								"</td>" & vbCrLf & _
								"<td align='left'>" & vbCrLf & _
									"<font size='2' face='trebuchet MS'>&nbsp;<b>" & tmpInst & " - " & tmpDept & "</b></font><br>" & vbCrLf & _
								"</td>" & vbCrLf & _
							"</tr>" & vbCrLf & _
							"<tr>" & vbCrLf & _
								"<td align='right'>" & vbCrLf & _
									"<font size='2' face='trebuchet MS'>Requesting Person:</font><br>" & vbCrLf & _
								"</td>" & vbCrLf & _
								"<td align='left'>" & vbCrLf & _
									"<font size='2' face='trebuchet MS'>&nbsp;<b>" & tmpUname & "</b></font><br>" & vbCrLf & _
								"</td>" & vbCrLf & _
							"</tr>" & vbCrLf & _
						"</table>" & vbCrLf & _
					"</td></tr>" & vbCrLf & _
					"<tr><td>&nbsp;</td></tr>" & vbCrLf & _	
					"<tr><td>" & vbCrLf & _
						"<table cellpadding='0' cellspacing='0' border='2' width='100%'>" & vbCrLf & _
							"<tr>" & vbCrLf & _
								"<td align='right' width='225px'>" & vbCrLf & _
									"<font size='2' face='trebuchet MS'>HospitalPilot ID:</font><br>" & vbCrLf & _
								"</td>" & vbCrLf & _
								"<td align='left'>" & vbCrLf & _
									"<font size='2' face='trebuchet MS'>&nbsp;<b>" & tmpID & "</b></font><br>" & vbCrLf & _
								"</td>" & vbCrLf & _
							"</tr>" & vbCrLf & _
							"<tr>" & vbCrLf & _
								"<td align='right' width='225px'>" & vbCrLf & _
									"<font size='2' face='trebuchet MS'>LanguageBank ID:</font><br>" & vbCrLf & _
								"</td>" & vbCrLf & _
								"<td align='left'>" & vbCrLf & _
									"<font size='2' face='trebuchet MS'>&nbsp;<b>" & tmpLBID & "</b></font><br>" & vbCrLf & _
								"</td>" & vbCrLf & _
							"</tr>" & vbCrLf & _
							"<tr>" & vbCrLf & _
								"<td align='right'>" & vbCrLf & _
									"<font size='2' face='trebuchet MS'>Client Name:</font><br>" & vbCrLf & _
								"</td>" & vbCrLf & _
								"<td align='left'>" & vbCrLf & _
									"<font size='2' face='trebuchet MS'>&nbsp;<b>" & tmpCName & "&nbsp;" & tmpCall & "</b></font><br>" & vbCrLf & _
								"</td>" & vbCrLf & _
							"</tr>" & vbCrLf & _
							"<tr>" & vbCrLf & _
								"<td align='right'>" & vbCrLf & _
									"<font size='2' face='trebuchet MS'>Phone No.:</font><br>" & vbCrLf & _
								"</td>" & vbCrLf & _
								"<td align='left'>" & vbCrLf & _
									"<font size='2' face='trebuchet MS'>&nbsp;<b>" & tmpFon & "</b></font><br>" & vbCrLf & _
								"</td>" & vbCrLf & _
							"</tr>" & vbCrLf & _
							"<tr>" & vbCrLf & _
								"<td align='right'>" & vbCrLf & _
									"<font size='2' face='trebuchet MS'>Mobile No.:</font><br>" & vbCrLf & _
								"</td>" & vbCrLf & _
								"<td align='left'>" & vbCrLf & _
									"<font size='2' face='trebuchet MS'>&nbsp;<b>" & tmpMobile & "</b></font><br>" & vbCrLf & _
								"</td>" & vbCrLf & _
							"</tr>" & vbCrLf & _
							"<tr>" & vbCrLf & _
								"<td align='right'>" & vbCrLf & _
									"<font size='2' face='trebuchet MS'>Date of Appointment:</font><br>" & vbCrLf & _
								"</td>" & vbCrLf & _
								"<td align='left'>" & vbCrLf & _
									"<font size='2' face='trebuchet MS'>&nbsp;<b>" & tmpAppDate & "</b></font><br>" & vbCrLf & _
								"</td>" & vbCrLf & _
							"</tr>" & vbCrLf & _
							"<tr>" & vbCrLf & _
								"<td align='right'>" & vbCrLf & _
									"<font size='2' face='trebuchet MS'>Time of Appointment:</font><br>" & vbCrLf & _
								"</td>" & vbCrLf & _
								"<td align='left'>" & vbCrLf & _
									"<font size='2' face='trebuchet MS'>&nbsp;<b>" & tmpAppTime & "</b></font><br>" & vbCrLf & _
								"</td>" & vbCrLf & _
							"</tr>" & vbCrLf & _
							"<tr>" & vbCrLf & _
								"<td align='right'>" & vbCrLf & _
									"<font size='2' face='trebuchet MS'>Language:</font><br>" & vbCrLf & _
								"</td>" & vbCrLf & _
								"<td align='left'>" & vbCrLf & _
									"<font size='2' face='trebuchet MS'>&nbsp;<b>" & tmpLang & "</b></font><br>" & vbCrLf & _
								"</td>" & vbCrLf & _
							"</tr>" & vbCrLf & _
							"<tr>" & vbCrLf & _
								"<td align='right' valign='top'>" & vbCrLf & _
									"<font size='2' face='trebuchet MS'>Reason:</font><br>" & vbCrLf & _
								"</td>" & vbCrLf & _
								"<td align='left'>" & vbCrLf & _
									"<font size='2' face='trebuchet MS'>&nbsp;<b>" & tmpReas & "</b></font><br>" & vbCrLf & _
								"</td>" & vbCrLf & _
							"</tr>" & vbCrLf & _
							"<tr>" & vbCrLf & _
								"<td align='right' valign='top'>" & vbCrLf & _
									"<font size='2' face='trebuchet MS'>Appointment Comment:</font><br>" & vbCrLf & _
								"</td>" & vbCrLf & _
								"<td align='left'>" & vbCrLf & _
									"<font size='2' face='trebuchet MS'>&nbsp;<b>" & tmpACom & "</b></font><br>" & vbCrLf & _
								"</td>" & vbCrLf & _
							"</tr>" & vbCrLf & _
							"<tr>" & vbCrLf & _
								"<td align='right' valign='top'>" & vbCrLf & _
									"<font size='2' face='trebuchet MS'>Languagebank Comment:</font><br>" & vbCrLf & _
								"</td>" & vbCrLf & _
								"<td align='left'>" & vbCrLf & _
									"<font size='2' face='trebuchet MS'>&nbsp;<b>" & tmpLCom & "</b></font><br>" & vbCrLf & _
								"</td>" & vbCrLf & _
							"</tr>" & vbCrLf & _
							"<tr>" & vbCrLf & _
								"<td align='right' valign='top'>" & vbCrLf & _
									"<font size='2' face='trebuchet MS'>Interpreter Comment:</font><br>" & vbCrLf & _
								"</td>" & vbCrLf & _
								"<td align='left'>" & vbCrLf & _
									"<font size='2' face='trebuchet MS'>&nbsp;<b>" & tmpICom & "</b></font><br>" & vbCrLf & _
								"</td>" & vbCrLf & _
							"</tr>" & vbCrLf & _
						"</table>" & vbCrLf & _
					"</td></tr>" & vbCrLf & _
					"<tr><td>&nbsp;</td></tr>" & vbCrLf & _
					"<tr><td align='left'>" & vbCrLf & _
					"<font size='1' face='trebuchet MS'>* Please do not reply to this email. This is a computer generated email. Use the information above for questions.</font>" & vbCrLf & _
					"</td></tr>" & vbCrLf & _
				"</table>"
			mlMail.HTMLBody = "<html><body>" & vbCrLf & strBody & vbCrLf & "</body></html>"
		'on error resume next 
'			mlMail.Send
			Session("MSG") = "Request for Interpreter submitted to LanguageBank."
		set mlMail=nothing
		Call AddLog("Appointment " & tmpID & " EMAILS SENT TO INSTITUTION.")
		Call AddLog("Appointment " & tmpID & " SUCCESS.")
		'send email here
		Response.Redirect "reqconfirm.asp?ID=" & tmpID
	Else
		Response.Redirect "main.asp"	
	End If	
ElseIf Request("ctrl") = 2 Then 'save reason
	'SAVE COMMENT
	Set rsCom = Server.CreateObject("ADODB.RecordSet")
	sqlCom = "SELECT * FROM Comment_T WHERE UID = " & Request("HideID")
	rsCom.Open sqlCom, g_strCONN, 1, 3
	If rsCom.EOF Then
		rsCom.AddNew
		rsCom("UID") = Request("HideID")
	End If
	rsCom("comment") = Request("HidCom")
	rsCom.Update
	rsCom.Close
	Set rsCom = Nothing
	'DELETE KEY IF NOT SAME
	Set rsDel = Server.CreateOBject("ADODB.RecordSet")
	sqlDel = "SELECT * FROM Encounter_T WHERE appID = " & Request("HideID")
	rsDel.Open sqlDel, g_strCONN, 1, 3
	If Not rsDel.EOF Then
		myOKey = rsDel("Key")
	End If
	rsDel.Close
	If Cint(MyOKey) <> Cint(Request("HideKey")) Then
		sqlDel = "DELETE * FROM Encounter_T WHERE appID = " & Request("HideID")
		rsDel.Open sqlDel, g_strCONN, 1, 3
	End If
	Set rsDel = Nothing
	'SAVE KEY AND REASON
	Set rsKey = Server.CreateOBject("ADODB.RecordSet")
	sqlKey = "SELECT * FROM Encounter_T WHERE appID = " & Request("HideID") & " AND key = " & Request("HideKey") & " AND reason = " & Request("HideReas")
	rsKey.Open sqlKey, g_strCONN, 1, 3
	If rsKey.EOF Then
		If Request("HideReas") <> 0 Then
			rsKey.AddNew
			rsKey("appID") = Request("HideID")
			rsKey("key") = Request("HideKey")
			rsKey("Reason") = Request("HideReas")
			rsKey.Update
		End If
	End If
	rsKey.Close
	Set rsKey = Nothing
	'SAVE KEY ON APPT TBL
	Set rsSkey = Server.CreateObject("ADODB.RecordSet")
	sqlSkey = "SELECT * FROM Appointment_T WHERE [index] = " & Request("HideID")
	rsSkey.Open sqlSkey, g_strCONN, 1, 3
	If Not rsSkey.EOF Then
		tmpAppDate = rsSkey("appDate")
		rsSkey("key") = Request("HideKey")
		rsSkey.Update
	End If
	rsSkey.Close
	Set rsSkey = Nothing
	'SAVE ACTUAL TIME
	Set rsAtime = Server.CreateObject("ADODB.RecordSet")
	sqlAtime = "SELECT * FROM Appointment_T WHERE [index] = " & Request("HideID")
	rsAtime.Open sqlAtime, g_strCONN, 1, 3
	If Not rsAtime.EOF Then
		If Request("HideAST") <> "" then
			rsAtime("AStime") = Request("HideAST")
		Else
			rsAtime("AStime") = Empty
		End If
		If Request("HideAET") <> "" Then
			rsAtime("AEtime") = Request("HideAET")
		Else
			rsAtime("AEtime") = Empty
		End If
		rsAtime.Update
	End If
	rsAtime.Close
	Set rsAtime = Nothing
	'SAVE ON LB ACTUAL TIME
	Set rsAtime = Server.CreateObject("ADODB.RecordSet")
	sqlAtime = "SELECT * FROM request_T WHERE HPID = " & Request("HideID")
	rsAtime.Open sqlAtime, g_strCONNLB, 1, 3
	If Not rsAtime.EOF Then
		If Request("HideAST") <> "" then
			rsAtime("AStarttime") = Request("HideAST")
		Else
			rsAtime("AStarttime") = Empty
		End If
		If Request("HideAET") <> "" Then
			rsAtime("AEndtime") = Request("HideAET")
		Else
			rsAtime("AEndtime") = Empty
		End If
		rsAtime.Update
	End If
	rsAtime.Close
	Set rsAtime = Nothing
	'SAVE FOLLOW UP TIME and CONFIRMATION
	Set rsFol = Server.CreateObject("ADODB.RecordSet")
	sqlFol = "SELECT * FROM Appointment_T WHERE [index] = " & Request("HideID")
	rsFol.Open sqlFol, g_strCONN, 1, 3
	If Not rsFol.EOF Then
		rsFol("Follow") = Z_Czero(Request("HideFollow"))
		rsFol("Confirm") = Z_Czero(Request("HideConfirm"))
		rsFol.Update
	End If
	rsFol.Close
	Set rsFol = Nothing
	Response.Redirect "Encounter.asp?NewKey=1&txtAppDate=" & tmpAppDate
ElseIf Request("ctrl") = 3 Then 'edit appointment
	'STORE ENTRIES ON COOKIE FOR EDITING AND SAVING ENTRIES
	Response.Cookies("LBREQUEST") = Z_DoEncrypt(Request("txtClilname")	& "|" & _
		Request("txtClifname")	& "|" & Request("selReas")	& "|" & Request("chkCall")	& "|" & Request("selDept")	& "|" & _
		Request("txtCliFon")	& "|" & Request("txtCliMobile")	& "|" & _
		Request("selLang") & "|" & Request("txtAppDate")	& "|" & Request("txtAppTFrom")	& "|" & Request("txtAppTTo") & "|" & Request("txtcom") & "|" & _
		Request("txtClinName")	& "|" & Request("txtTP") & "|" & Request("radioLang") & "|" & Request("chkminor") & _
		"|" & Request("txtparents") & "|" & Request("txtLBcom") & "|" & Request("txtIntrcom") & "|" & Request("selGen") & "|" & _
		Request("chkClientAdd")	& "|" & Request("txtCliAddrI") & "|" & Request("txtCliadd") & "|" & Request("txtClicity") & "|" & _
		Request("txtClistate")	& "|" & Request("txtCliZip") & "|" & Request("txtRFon") & "|" & Request("txtdhhsemail") & "|" & _
		Request("txtdhhsFon") & "|" & Request("txtoLang") & "|" & Request("chkCall2") & "|" & Request("txtchrg") & "|" & Request("txtAtrny") & "|" & _
		Request("txtDOB") & "|" & Request("txtPDamount") & "|" & Request("h_tmpfilename") & "|" & Request("chkout") & "|" & Request("chkmed") & "|" & _
		Request("MCnum") & "|" & Request("chkacc") & "|" & Request("chkcomp") & "|" & Request("selIns") & "|" & Request("txtemail") & "|" & _
		Request("MHPnum") & "|" & Request("NHHFnum") & "|" & Request("WSHPnum") & "|" & Request("chkawk")& "|" & Request("mrrec")& "|" & Request("chkleave") & "|" & Request("txtCliCir"))
	'CHECK VALID VALUES
	If Session("myClass") <> 3 Then
		If Request("txtTP") <> "" Then
			If Not IsDate(Request("txtTP")) Then Session("MSG") = Session("MSG") & "<br>ERROR: Invalid Time Paged."
		End If
	End If
	If Request("txtAppdate") <> "" Then
		If Not IsDate(Request("txtAppdate")) Then Session("MSG") = Session("MSG") & "<br>ERROR: Invalid Appointment date."
	End If
	If Request("txtAppTFrom") <> "" Then
		If Not IsDate(Request("txtAppTFrom")) Then Session("MSG") = Session("MSG") & "<br>ERROR: Invalid Appointment Time (From:)."
	End If
	If Request("txtAppTTo") <> "" Then
		If Not IsDate(Request("txtAppTTo")) Then Session("MSG") = Session("MSG") & "<br>ERROR: Invalid Appointment Time (To:)."
	End If
	If Request("txtDOB") <> "" Then
		If Not IsDate(Request("txtDOB")) Then Session("MSG") = Session("MSG") & "<br>ERROR: Invalid Date of Birth."
	End If
	If Session("MSG") = "" Then	
		'GET COOKIE OF REQUEST
		tmpEntry = Split(Z_DoDecrypt(Request.Cookies("LBREQUEST")), "|")
		'SAVE ENTRIES
		Set rsMain = Server.CreateObject("ADODB.RecordSet")
		sqlMain = "SELECT * FROM appointment_T WHERE [index] = " & Request("ID")
		rsMain.Open sqlMain, g_strCONN, 1, 3
		If Not rsMain.EOF Then
			tmpappdate = rsMain("appDate")
			tmpstime = rsMain("TimeFrom")
			tmpetime = rsMain("TimeTo")
			tmpdatefix = tmpEntry(8)
			lang = rsMain("langID")
			If Year(tmpdatefix) < 1981 Then
				tmpdatefix = Request("txtAppDate")
			End If
			If rsMain("IntrID") > 0 Then 
				If tmpappdate <> tmpEntry(8) Then 
					Session("MSG") = Session("MSG") & "<br>ERROR: You cannot edit appointment date.<br>An interpreter has already been assigned.<br>Please cancel this appointment then clone this appointment to change date."
					rsMain.Close
					Set rsMain = Nothing
					Response.Redirect "main.asp?ID=" & Request("ID")
				End If
				If tmpstime <> tmpdatefix & " " & Z_FormatTime(tmpEntry(9)) Then 
					Session("MSG") = Session("MSG") & "<br>ERROR: You cannot edit appointment start time.<br>An interpreter has already been assigned.<br>Please cancel this appointment then clone this appointment to change start time."
					rsMain.Close
					Set rsMain = Nothing
					Response.Redirect "main.asp?ID=" & Request("ID")
				End If
			End If	
			rsMain("clname") = Z_DoEncrypt(CleanMe(tmpEntry(0)))
			rsMain("cfname") = Z_DoEncrypt(CleanMe(tmpEntry(1)))
			'rsMain("addr") = CleanMe(tmpEntry(2))
			'rsMain("city") = tmpEntry(3)
			'rsMain("state") = Ucase(tmpEntry(4))
			'rsMain("zip") = tmpEntry(5)
			rsMain("Phone") = Z_DoEncrypt(tmpEntry(5))
			rsMain("Mobile") = Z_DoEncrypt(tmpEntry(6))
			If Session("UID") <> 36 Then
				rsMain("langID") = tmpEntry(7)
			End If
			tmpappdate = rsMain("appdate")
			if tmpappdate >= date Then
				rsMain("appdate") = tmpEntry(8)
			Else
				Session("MSG") = Session("MSG") & "<br>ERROR: You cannot edit appointment date.<br>If you wish to have another appointment with the same data please use the clone button."
			End If
			rsMain("TimeFrom") = tmpdatefix & " " & Z_FormatTime(tmpEntry(9))
			rsMain("TimeTo") = tmpdatefix & " " & Z_FormatTime(tmpEntry(10))
			rsMain("Comment") = tmpEntry(11)
			rsMain("DOB") = Z_DateNull(tmpEntry(33))
			If Session("myClass") <> 3 Then
				Pagedme = Empty
				If tmpEntry(13) <> "" Then Pagedme = tmpEntry(13)
				rsMain("paged") = Pagedme
				rsMain("clinician") = tmpEntry(12)
			Else
				rsMain("crtroom") = tmpEntry(13)
				rsMain("docknum") = tmpEntry(12)
				rsMain("attny") = tmpEntry(32)
				rsMain("charges") = tmpEntry(31)
			End If
			If tmpEntry(2) = "" Then tmpEntry(2) = "0"
			rsMain("reason") = tmpEntry(2)
			rsMain("callme") = False
			If tmpEntry(3) <> "" Then rsMain("callme") = True
			rsMain("DeptID") = tmpEntry(4)	
			rsMain("minor") = False
			If tmpEntry(15) <> "" Then rsMain("minor") = True
			rsMain("parents") = tmpEntry(16)
			rsMain("lbcom") = tmpEntry(17)
			rsMain("intrcom") = tmpEntry(18)
			rsMain("Gender") = tmpEntry(19)
		If tmpEntry(20) <> "" Then 
			rsMain("useCadr") = True
			rsMain("capt") = tmpEntry(21)
			rsMain("caddress") = tmpEntry(22)
			rsMain("ccity") = tmpEntry(23)
			rsMain("cstate") = tmpEntry(24)
			rsMain("czip") = tmpEntry(25)
		Else
			rsMain("useCadr") = False
			rsMain("capt") = ""
			rsMain("caddress") = ""
			rsMain("ccity") = ""
			rsMain("cstate") = ""
			rsMain("czip") = ""
		End If
		rsMain("block") = False
		If tmpEntry(30) <> "" Then rsMain("block") = True
		rsMain("semail") = tmpEntry(27)
		rsMain("sphone") = tmpEntry(28)
		rsMain("rphone") = tmpEntry(26)
		rsMain("olang") = tmpEntry(29)
		If Session("type") = 5 Then
			rsMain("PDemail") = tmpEntry(42)
			rsMain("PDamount") = Z_CZero(tmpEntry(34))
			If FileUpload(tmpEntry(35)) Then 
				rsMain("UploadFile") = True
				rsMain("filename") = tmpEntry(35) & ".PDF"
			End If
		End If
		rsMain("outpatient") = False
		If tmpEntry(36) <> "" Then rsMain("outpatient") = True
		rsMain("hasmed") = False
		If tmpEntry(37) <> "" Then rsMain("hasmed") = True
		rsMain("medicaid") = tmpEntry(38)
		rsMain("meridian") = tmpEntry(43)
		rsMain("nhhealth") = tmpEntry(44)
		rsMain("wellsense") = tmpEntry(45)
		rsMain("acknowledge") = false
		If tmpEntry(46) <> "" Then rsMain("acknowledge") = True
		rsMain("autoacc") = False
		If tmpEntry(39) <> "" Then rsMain("autoacc") = True
		rsMain("wcomp") = False
		If tmpEntry(40) <> "" Then rsMain("wcomp") = True
		rsMain("secins") = Z_FixNull(tmpEntry(41))
		rsMain("mrrec") = tmpEntry(47)
		rsMain("leavemsg") = False
		If tmpEntry(48) <> "" Then rsMain("leavemsg") = True
		rsMain("Spec_cir") = tmpEntry(49)
		rsMain.Update
		'GET ID FOR CONFIRM
		tmpID = rsMain("index")
		End If
		rsMain.Close
		Set rsMain = Nothing
		If Session("type") = 5 AND FileUpload(tmpEntry(35)) Then 'save Form on DB
			Set rsFile = Server.CreateObject("ADODB.RecordSet")
			sqlFile = "SELECT * FROM pdf_T"
			rsFile.Open sqlFile, g_strCONN, 1, 3
			rsFile.AddNew
			rsFile("appID") = tmpID
			rsFile("filename") = tmpEntry(35) & ".PDF"
			rsFile("datestamp") = Now
			rsFile.Update
			rsFile.Close
			Set rsFile = Nothing
		End If
		'SAVE APPOINTMENT IN LANGUAGE BANK
			Set rsLB = Server.CreateObject("ADODB.RecordSet")
			sqlLB = "SELECT * FROM Request_T WHERE HPID = " & Request("ID")
			rsLB.Open sqlLB, g_strCONNLB, 1, 3
			If not rsLB.EOF Then 
			if tmpappdate >= date Then rsLB("appdate") =  tmpEntry(8)
			rsLB("appTimeFrom") =  tmpEntry(8) & " " & Z_FormatTime(tmpEntry(9))
			rsLB("appTimeTo") = tmpEntry(8) & " " & Z_FormatTime(tmpEntry(10))
			If Session("UID") <> 36 Then
				rsLB("langID") = GetLBLang(tmpEntry(7))
			End If
			rsLB("clname") = CleanMe(tmpEntry(0))
			rsLB("cfname") = CleanMe(tmpEntry(1))
			rsLB("Cphone") = tmpEntry(5)
			rsLB("DeptID") = tmpEntry(4)
			'rsLB("InstID") = Request("LBID")
			'rsLB("DeptID") = Session("DeptID")
			rsLB("Comment") = tmpEntry(11)
			rsLB("CAphone") =  tmpEntry(6)
			'rsLB("LBcomment") = tmpEntry(17)
			rsLB("Intrcomment") = tmpEntry(18)
			rsLB("Gender") = tmpEntry(19)
			rsLB("DOB") = Z_dateNull(tmpEntry(33))
		If tmpEntry(20) <> "" Then 
			rsLB("CliAdd") = True
			rsLB("CliAdrI") = tmpEntry(21)
			rsLB("caddress") = tmpEntry(22)
			rsLB("ccity") = tmpEntry(23)
			rsLB("cstate") = tmpEntry(24)
			rsLB("czip") = tmpEntry(25)
		Else
			rsLB("CliAdd") = False
			rsLB("CliAdrI") = ""
			rsLB("caddress") = ""
			rsLB("ccity") = ""
			rsLB("cstate") = ""
			rsLB("czip") = ""
		End If
		rsLB("Child") = False
		If tmpEntry(14) <> "" Then rsLB("Child") = True
		If Session("myClass") = 3 Then
			rsLB("DocNum") = tmpEntry(12)
			rsLB("CrtRumNum") = tmpEntry(13)
			rsLB("Comment") = tmpEntry(11) & vbCrlf & " Charge/s: " & tmpEntry(31) & vbCrlf & " Attorney: " & tmpEntry(32)
		End If
		If Session("type") = 5 Then
			rsLB("DocNum") = tmpEntry(12)
			rsLB("PDamount") = Z_CZero(tmpEntry(34))
			If FileUpload(tmpEntry(35)) Then 
				rsLB("UploadFile") = True
				rsLB("filename") = tmpEntry(35) & ".PDF"
			End If
		End If
		rsLB("outpatient") = False
		If tmpEntry(36) <> "" Then rsLB("outpatient") = True
		rsLB("hasmed") = False
		If tmpEntry(37) <> "" Then rsLB("hasmed") = True
		rsLB("medicaid") = tmpEntry(38)
		rsLB("meridian") = tmpEntry(43)
		rsLB("nhhealth") = tmpEntry(44)
		rsLB("wellsense") = tmpEntry(45)
		rsLB("acknowledge") = false
		If tmpEntry(46) <> "" Then rsLB("acknowledge") = True
		rsLB("autoacc") = False
		If tmpEntry(39) <> "" Then rsLB("autoacc") = True
		rsLB("wcomp") = False
		If tmpEntry(40) <> "" Then rsLB("wcomp") = True
		rsLB("secins") = Z_FixNull(tmpEntry(41))
		rsLB("mrrec") = tmpEntry(47)
		rsLB("leavemsg") = False
		If tmpEntry(48) <> "" Then rsLB("leavemsg") = True
		rsLB("Spec_cir") = tmpEntry(49)
		rsLB("BlockSched") = False
		If tmpEntry(30)<> "" Then rsLB("BlockSched") = True
		rsLB("courtcall") = false
		If tmpEntry(3) <> "" Then rsLB("courtcall") = True
		rsLB.Update
		'GET INFO FOR EMAIL
		tmpLBID = rsLB("index")
		tmpInst = GetFacility(Session("InstID"))
		tmpDept = GetDept(rsLB("deptID"))
		tmpCName = CleanMe(rsLB("clname")) & ", " & CleanMe(rsLB("cfname"))
		tmpAppDate = rsLB("appDate")
		tmpAppTime = rsLB("appTimeFrom") & " - " & rsLB("appTimeTo")
		tmpHPID = rsLB("HPID")
		rsLB.Close
		Set rsLB = Nothing
		'GET ID FOR CONFIRM
		'tmpID = rsMain("index")
		End If
		'do not include ASL
		'response.write "TMP: " & tmpEntry(14)
		If tmpEntry(30) = "" Then 'tmpEntry(30) <> ""
			If GetLBLang(tmpEntry(7)) <> 52 And GetLBLang(tmpEntry(7)) <> 78 And GetLBLang(tmpEntry(7)) <> 81 And Z_CZero(tmpEntry(14)) <> 4 Then
				'reset interpreter DB if changed date/time/lang
				If Not (tmpappdate = Z_DateNull(tmpEntry(8)) And stime = Z_DateNull(tmpEntry(8)) & " " & Z_DateNull(tmpEntry(9))) And lang = tmpEntry(7)Then
					Call Z_ResetIntr(tmpLBID)
				ElseIf lang <> tmpEntry(7) Then
					Call Z_ResetIntr2(tmpLBID)
				End If
			End If
		End If
		'rsMain.Close
		'Set rsMain = Nothing
		'SAVE HISTORY IN LB
		TimeNow = Now
		Set rsHist = Server.CreateObject("ADODB.RecordSet")
		sqlHist = "SELECT * FROM History_T WHERE REqID = " & tmpLBID
		rsHist.Open sqlHist, g_strCONNHist, 1,3 
		if NOT rsHist.EOF Then
			If Z_DateNull(rsHist("date")) <> Z_DateNull(tmpEntry(8)) Then
				tmpAppDate = tmpAppDate & " (" & rsHist("date") & ")"
				rsHist("date") = tmpEntry(8)
				rsHist("dateTS") = TimeNow
				rsHist("dateU") = Session("GreetMe")
			End If
			If Z_DateNull(rsHist("Stime")) <> Z_dateNull(tmpEntry(8) & " " & tmpEntry(9)) Then
				tmpAppTime = tmpAppTime & " (" & rsHist("Stime") & ")"
				rsHist("Stime") = Z_dateNull(tmpEntry(8) & " " & tmpEntry(9))
				rsHist("StimeTS") = TimeNow
				rsHist("StimeU") = Session("GreetMe")
			End If
			rsHist.Update
		end if
		rsHist.Close
		Set rsHist = Nothing
		'on error resume next
			Set mlMail = CreateObject("CDO.Message")
			mlMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing")= 2
			mlMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "localhost"
			mlMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 26
			mlMail.Configuration.Fields.Update
			mlMail.To = "language.services@thelanguagebank.org"
			'mlMail.Bcc = "sysdump1@zubuk.com"
			mlMail.From = "language.services@thelanguagebank.org"
			mlMail.Subject = "Interpreter Request (edited) - " & tmpInst & " - " & tmpDept
			strBody = "<table cellpadding='0' cellspacing='0' border='0' align='center'>" & vbCrLf & _
					"<tr><td align='center'>" & vbCrLf & _
						"<img src='http://web04.zubuk.com/lss-lbis-staging/images/LBISLOGOBandW.jpg'>" & vbCrLf & _
					"</td></tr>" & vbCrLf & _
					"<tr><td>&nbsp;</td></tr>" & vbCrLf & _	
					"<tr><td align='center'>" & vbCrLf & _
						"<font size='2' face='trebuchet MS'><b>Interpreter Request:</b></font><br>" & vbCrLf & _
					"</td></tr>" & vbCrLf & _
					"<tr><td>&nbsp;</td></tr>" & vbCrLf & _
					"<tr><td>" & vbCrLf & _
						"<table cellpadding='0' cellspacing='0' border='2' width='100%'>" & vbCrLf & _
							"<tr>" & vbCrLf & _
								"<td align='right' width='225px'>" & vbCrLf & _
									"<font size='2' face='trebuchet MS'>Institution - Department:</font><br>" & vbCrLf & _
								"</td>" & vbCrLf & _
								"<td align='left'>" & vbCrLf & _
									"<font size='2' face='trebuchet MS'>&nbsp;<b>" & tmpInst & " - " & tmpDept & "</b></font><br>" & vbCrLf & _
								"</td>" & vbCrLf & _
							"</tr>" & vbCrLf & _
						"</table>" & vbCrLf & _
					"</td></tr>" & vbCrLf & _
					"<tr><td>&nbsp;</td></tr>" & vbCrLf & _	
					"<tr><td>" & vbCrLf & _
						"<table cellpadding='0' cellspacing='0' border='2' width='100%'>" & vbCrLf & _
							"<tr>" & vbCrLf & _
								"<td align='right' width='225px'>" & vbCrLf & _
									"<font size='2' face='trebuchet MS'>HospitalPilot ID:</font><br>" & vbCrLf & _
								"</td>" & vbCrLf & _
								"<td align='left'>" & vbCrLf & _
									"<font size='2' face='trebuchet MS'>&nbsp;<b>" & tmpHPID & "</b></font><br>" & vbCrLf & _
								"</td>" & vbCrLf & _
							"</tr>" & vbCrLf & _
							"<tr>" & vbCrLf & _
								"<td align='right' width='225px'>" & vbCrLf & _
									"<font size='2' face='trebuchet MS'>LanguageBank ID:</font><br>" & vbCrLf & _
								"</td>" & vbCrLf & _
								"<td align='left'>" & vbCrLf & _
									"<font size='2' face='trebuchet MS'>&nbsp;<b>" & tmpLBID & "</b></font><br>" & vbCrLf & _
								"</td>" & vbCrLf & _
							"</tr>" & vbCrLf & _
							"<tr>" & vbCrLf & _
								"<td align='right'>" & vbCrLf & _
									"<font size='2' face='trebuchet MS'>Client Name:</font><br>" & vbCrLf & _
								"</td>" & vbCrLf & _
								"<td align='left'>" & vbCrLf & _
									"<font size='2' face='trebuchet MS'>&nbsp;<b>" & tmpCName & "</b></font><br>" & vbCrLf & _
								"</td>" & vbCrLf & _
							"</tr>" & vbCrLf & _
							"<tr>" & vbCrLf & _
								"<td align='right'>" & vbCrLf & _
									"<font size='2' face='trebuchet MS'>Date of Appointment:</font><br>" & vbCrLf & _
								"</td>" & vbCrLf & _
								"<td align='left'>" & vbCrLf & _
									"<font size='2' face='trebuchet MS'>&nbsp;<b>" & tmpAppDate & "</b></font><br>" & vbCrLf & _
								"</td>" & vbCrLf & _
							"</tr>" & vbCrLf & _
							"<tr>" & vbCrLf & _
								"<td align='right'>" & vbCrLf & _
									"<font size='2' face='trebuchet MS'>Time of Appointment:</font><br>" & vbCrLf & _
								"</td>" & vbCrLf & _
								"<td align='left'>" & vbCrLf & _
									"<font size='2' face='trebuchet MS'>&nbsp;<b>" & tmpAppTime & "</b></font><br>" & vbCrLf & _
								"</td>" & vbCrLf & _
							"</tr>" & vbCrLf & _
						"</table>" & vbCrLf & _
					"</td></tr>" & vbCrLf & _
					"<tr><td>&nbsp;</td></tr>" & vbCrLf & _
					"<tr><td align='left'>" & vbCrLf & _
					"<font size='1' face='trebuchet MS'>* Value in parenthesis is previous value.</font>" & vbCrLf & _
					"</td></tr>" & vbCrLf & _
				"</table>"
			mlMail.HTMLBody = "<html><body>" & vbCrLf & strBody & vbCrLf & "</body></html>"
		'on error resume next
			mlMail.Send
			Session("MSG") = "Changes for this appointment has been submitted to LanguageBank."
		If SaveHist(tmpLBID, "[HP]main.asp") Then
	
		End If
		'log edit
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set LogMe = fso.OpenTextFile("C:\work\lss-lbis\log\editlogs.txt", 8, True)
		strLog = Now & vbtab & "Appointment EDITED - ID: " & tmpLBID & " by " & Session("GreetMe") & " DATE: " & tmpAppDate & " TIME: " & tmpAppTime
		LogMe.WriteLine strLog
		Set LogMe = Nothing
		Set fso = Nothing
		set mlMail = nothing
		Response.Redirect "reqconfirm.asp?ID=" & tmpID
	Else
		Response.Redirect "main.asp?ID=" & Request("ID")	
	End If	
ElseIf Request("ctrl") = 4 Then 'calendar function
	tmpMonthYear = Split(Request("Hmonth"), " - ")
	tmpMonth = tmpMonthYear(0) & "/01/" & tmpMonthYear(1)
	If IsNumeric(tmpMonthYear(1)) Then
		If Request("dir") = 0 Then
			tmpMonth = DateAdd("m", -1, tmpMonth)
		Else
			tmpMonth = DateAdd("m", 1, tmpMonth)
		End If
	End If
	'Response.Redirect "calendarview.asp?selMonth=" & Month(tmpMonth) & "&txtyear=" & Year(tmpMonth)
	Response.Redirect "calendarview2.asp?selMonth=" & Month(tmpMonth) & "&txtday=1&txtyear=" & Year(tmpMonth)
ElseIf Request("ctrl") = 5 Then 'assign interpreter
	Set rsIntr = Server.CreateObject("ADODB.RecordSet")
	sqlIntr ="SELECT * FROM appointment_T WHERE [index] = " & Request("ID")
	rsIntr.Open sqlIntr, g_strCONN, 1, 3
	If Not rsIntr.EOF Then
		rsIntr("IntrID") = Request("selIntr")
		rsIntr.Update
	End If
	rsIntr.CLose
	Set rsIntr =Nothing
	If Request("selIntr") <> 0 Then Session("MSG") = "Interpreter Assigned."
	Response.Redirect "reqconfirm.asp?ID=" & Request("ID")
ElseIf Request("ctrl") = 6 Then 'print schedule
	Response.Cookies("LBREPORT") = Z_DoEncrypt(1 & "|" & Request("xdate"))
	Response.redirect "calendarview2.asp?rpt=1"
ElseIf Request("ctrl") = 7 Then 'reports
	Response.Cookies("HPREPORT") = Z_DoEncrypt(Request("selrep")	& "|" & Request("txtRepFrom") & "|" & Request("txtRepTo"))
	Call AddLog("REPORT INITIATED...")
	Response.Redirect "report.asp?Rtype=1"
ElseIf Request("ctrl") = 8 Then 'cancel request
	'SET STATUS inLB
	Set rsStat = Server.CreateObject("ADODB.RecordSet")
	sqlStat = "SELECT * FROM request_T WHERE HPID = " & Request("ID")
	rsStat.Open sqlStat, g_strCONNLB, 1, 3
	If Not rsStat.EOF Then
		tmpAppDateTime = rsStat("appTimeFrom")
		tmpAppDateTime2 = rsStat("appTimeTo")
		tmpIntr = Z_CZero(rsStat("intrID"))
		myClass = ClassInt(rsStat("deptID"))
		tmpInstID = rsStat("InstID")
		If DateDiff("n", Now, tmpAppDateTime) < 1440 and tmpIntr > 0 Then
			rsStat("status") = 4
			rsStat("Cancel") = 5
			rsStat("Missed") = 0
			rsStat("Astarttime") = tmpAppDateTime
			rsStat("Aendtime") = tmpAppDateTime
			rsStat("payHrs") = 2
			rsStat("Billable") = 2
			If tmpInstID = 273 Or myClass = 3 Then 'darth leb - court
				rsStat("Billable") = Z_GetBillhrs(tmpAppDateTime, tmpAppDateTime2)
			End If
			rsStat("happen") = 1
			rsStat("payintr") = false
			rsStat("M_Intr") = 0
			rsStat("TT_Intr") = 0
		Else
			rsStat("status") = 3
			rsStat("Cancel") = 5
			rsStat("Missed") = 0
		End If
		If myClass = 3 And DateDiff("n", Now, tmpAppDateTime) < 2880 And DateDiff("n", Now, tmpAppDateTime) > 1440 and tmpIntr > 0 Then 'courts and 24-48 hours cancel
			sStat("status") = 4
			rsStat("Cancel") = 5
			rsStat("Missed") = 0
			rsStat("Astarttime") = tmpAppDateTime
			rsStat("Aendtime") = tmpAppDateTime
			rsStat("payHrs") = 0
			rsStat("showIntr") = 0
			rsStat("Billable") = Z_GetBillhrsCourt(tmpAppDateTime, tmpAppDateTime2)	
		End If
		If tmpIntr > 0 Then 
			IntrName = GetIntrNameLB2(tmpIntr)
			rsStat("LBcomment") = rsStat("LBcomment") & vbCrlF & "Cancelation Email sent to " & IntrName & " on " & now
		Else
			rsStat("LBcomment") = rsStat("LBcomment") & vbCrlF & "Cancelation Email sent to Langbank on " & now
		End If
		
		rsStat.Update
		tmpLBID = rsStat("index")
	End If
	rsStat.Close
	Set rsStat = Nothing
	'SET STATUS in HP
	Set rsStat = Server.CreateObject("ADODB.RecordSet")
	sqlStat = "SELECT * FROM Appointment_T WHERE [index] = " & Request("ID")
	rsStat.Open sqlStat, g_strCONN, 1, 3
	If Not rsStat.EOF Then
		rsStat("status") = 3
		tmpInst = GetInstNameLB(rsStat("InstID"))
		tmpInstID = rsStat("InstID")
		tmpDate = rsStat("appDate")
		tmpTime = rsStat("TimeFrom") & " - " & rsStat("TimeTo")
		tmpTimeFrom = rsStat("TimeFrom")
		tmpTimeTo = rsStat("TimeTo")
		tmpDept = GetDeptNameLB(rsStat("deptID"))
		myClass = ClassInt(rsStat("deptID"))
		tmpName = Z_DoDecrypt(rsStat("clname")) & ", " & Z_DoDecrypt(rsStat("cfname"))
		tmpIntr = Z_Czero(rsStat("IntrID"))
		tmpCity = GetDeptCity(rsStat("DeptID"))
		tmpFname = Z_DoDecrypt(rsStat("cfname"))
		If tmpIntr > 0  Then IntrName = GetIntrNameLB2(tmpIntr)
		rsStat.Update
	End If
	rsStat.Close
	Set rsStat = Nothing
	If SaveHist(tmpLBID, "[HP]reqconfirm.asp") Then
		'log edit
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set LogMe = fso.OpenTextFile("C:\work\lss-lbis\log\editlogs.txt", 8, True)
		strLog = Now & vbtab & "Appointment Cancelled - ID: " & tmpLBID & " by " & Session("GreetMe") & " DATE: " & tmpDate & " TIME: " & tmpTime
		LogMe.WriteLine strLog
		Set LogMe = Nothing
		Set fso = Nothing
	End If
	'SEND EMAIL TO NOTIFY CANCEL TO LB
	Set mlMail = CreateObject("CDO.Message")
	mlMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing")= 2
	mlMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "localhost"
	mlMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 26
	mlMail.Configuration.Fields.Update
	mlMail.To = "language.services@thelanguagebank.org"
	mlMail.From = "language.services@thelanguagebank.org"
	'mlMail.Bcc = "sysdump1@zubuk.com"
	mlMail.Subject = "HospitalPilot - Request Cancellation - Request ID: " & tmpLBID
	strBody = "<img src='http://languagebank.lssne.org/lsslbis/images/LBISLOGOBandW.jpg'><br><br>" & vbCrLf & _
	 "<font size='2' face='trebuchet MS'>Request ID: " & tmpLBID & " (HospitalPilot ID: " & Request("ID") & ") has been CANCELED by " & Session("GreetMe") & ".<br>" & vbCrLf & _
	 "<font size='2' face='trebuchet MS'>Date: " & tmpDate & "<br>" & vbCrLf & _
	 "<font size='2' face='trebuchet MS'>Time: " & tmpTime & "<br>" & vbCrLf & _
	 "<font size='2' face='trebuchet MS'>Department: " & tmpDept & "<br>" & vbCrLf & _
	 "<font size='2' face='trebuchet MS'>Client: " & tmpName & "<br><br>" & vbCrLf & _ 
	 "<font size='1' face='trebuchet MS'>* Please do not reply to this email. This is a computer generated email.</font>" & vbCrLf
	mlMail.HTMLBody = "<html><body>" & vbCrLf & strBody & vbCrLf & "</body></html>"
	mlMail.Send
	 set mlMail=nothing
	 If tmpIntr > 0  Then
	  'SEND EMAIL TO NOTIFY CANCEL TO INTR
	 Set mlMail = CreateObject("CDO.Message")
		mlMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing")= 2
		mlMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "localhost"
		mlMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 26
		mlMail.Configuration.Fields.Update
		mlMail.To = GetPrime2(tmpIntr)
		mlMail.Cc = "language.services@thelanguagebank.org"
		'mlMail.Bcc = "sysdump1@zubuk.com"
		mlMail.From = "language.services@thelanguagebank.org"
		mlMail.Subject = "Appointment Cancellation " & tmpDate & "; " & tmpTime & ", " & tmpCity & " - " &  tmpInst
		strBody = "This is to let you know that appointment on " & _
			 tmpDate & ", " & tmpTime & ", in " & tmpCity & " at " & tmpInst & " for " & tmpFname & " is CANCELED.<br>" & _
			 "If you have any questions please contact the LanguageBank office immediately at 410-6183 or email us at " & _
			 "<a href='mailto:info@thelanguagebank.org'>info@thelanguagebank.org</a>.<br>" & _
			 "E-mail about this cancelation was initiated by " & Session("GreetMe") & ".<br><br>" & _
			 "Thanks,<br>" & _
			 "Language Bank"
		mlMail.HTMLBody = "<html><body>" & vbCrLf & strBody & vbCrLf & "</body></html>"
		mlMail.Send
		Set mlMail = Nothing
	End If
  	Session("MSG") = "NOTICE: Request has been cancelled." 
  	 Response.Redirect "reqconfirm.asp?ID=" & Request("ID")
ElseIf Request("ctrl") = 9 Then 'admin tools - user
 	If Request("selUser") <> 0 Then
 		Set rsUser = Server.CreateObject("ADODB.RecordSet")
 		sqlUser = "SELECT * FROM User_T WHERE [index] = " & Request("selUser")
 		rsUser.Open sqlUser, g_strCONN, 1, 3
 		If Not rsUser.EOF Then
 			rsUser("lname") = Request("txtlname")
 			rsUser("fname") = Request("txtfname")
 			rsUser("user") = Request("txtUser")
 			rsUser("pass") = Z_DoEncrypt(Request("txtpass"))
 			rsUser("type") = Request("selType")
 			rsUser.Update
 		End If
 		rsUser.Close
 		Set rsUser = Nothing
 		Session("MSG") = "User Saved."
 	Else
 		Set rsUser = Server.CreateObject("ADODB.RecordSet")
 		sqlUser = "SELECT * FROM User_T WHERE Upper(user) = '" & Trim(UCase(Request("txtUser"))) & "'"
 		rsUser.Open sqlUser, g_strCONN, 1, 3
 		If rsUser.EOF Then
	 		rsUser.AddNew
			rsUser("lname") = Request("txtlname")
			rsUser("fname") = Request("txtfname")
			rsUser("user") = Request("txtUser")
			rsUser("pass") = Z_DoEncrypt(Request("txtpass"))
			rsUser("type") = Request("selType")
			rsUser.Update
			Session("MSG") = "User Saved."
		Else
			Session("MSG") = "ERROR: Username already exists."
		End If
 		rsUser.Close
 		Set rsUser = Nothing
	End If
	Response.Redirect "a_User.asp?Use=" & Request("selUser")
ElseIf Request("ctrl") = 10 Then 'delete request 
	Set rsReq = Server.CreateObject("ADODB.RecordSet")
	sqlReq = "SELECT * FROM request_T WHERE HPID = " & Request("ID")
	rsReq.Open sqlReq, g_strCONNLB, 1, 3
 	If Not rsReq.EOF Then
 		tmpHPID = rsReq("index")
		rsReq.Delete
		rsReq.Update
	End If
	rsReq.Close
	Set rsReq = Nothing
	Set rsReq = Server.CreateObject("ADODB.RecordSet")
	sqlReq = "DELETE FROM appointment_T WHERE [index] = " & Request("ID")
	rsReq.Open sqlReq, g_strCONN, 1, 3
 	Set rsReq = Nothing
 	'CREATE LOG
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set LogMe = fso.OpenTextFile("C:\work\lss-lbis\logs.txt", 8, True)
	strLog = Now & vbtab & "Appointment DELETED ID: " & Request("ID") & " by " & Session("GreetMe")
	LogMe.WriteLine strLog
	Set LogMe = Nothing
	Set fso = Nothing
	Response.Redirect "main.asp"
ElseIf Request("ctrl") = 11 Then 'admin tools - institution
	If Request("selUser") <> 0 And Request("selInst") <> 0 Then
 		Set rsUser = Server.CreateObject("ADODB.RecordSet")
 		sqlUser = "SELECT * FROM User_T WHERE [index] = " & Request("selUser")
 		rsUser.Open sqlUser, g_strCONN, 1, 3
 		If Not rsUser.EOF Then
 			rsUser("InstID") = Request("selInst")
 			rsUser("ReqLB") = Request("selRP")
 			rsUser.Update
 		End If
 		rsUser.Close
 		Set rsUser = Nothing
	End If
	Session("MSG") = "Institution Saved."
	Response.Redirect "a_Inst.asp?Use=" & Request("selUser")
ElseIf Request("ctrl") = 12 Then 'admin tools - interpreter
	If Request("selUser") <> 0 And Request("selIntr") <> 0 Then
 		Set rsUser = Server.CreateObject("ADODB.RecordSet")
 		sqlUser = "SELECT * FROM User_T WHERE [index] = " & Request("selUser")
 		rsUser.Open sqlUser, g_strCONN, 1, 3
 		If Not rsUser.EOF Then
 			rsUser("intrLB") = Request("selIntr")
 			rsUser.Update
 		End If
 		rsUser.Close
 		Set rsUser = Nothing
	End If
	If Request("selInst") <> 0 Then
		Set rsAss = Server.CreateObject("ADODB.RecordSet")
		sqlAss = "SELECT * FROM InstIntr_T WHERE InstID = " & Request("selInst") & " AND IntrID = " & Request("selIntr")
		rsAss.Open sqlAss, g_strCONN, 1, 3
		If rsAss.EOF Then
			rsAss.AddNew
			rsAss("IntrID") = Request("selIntr")
			rsAss("InstID") = Request("selInst")
			rsAss.Update
		End If
		rsAss.Close
		Set rsAss = Nothing
	End If
	If Request("selUser") <> 0 And Request("selIntr") <> 0 Then
		Set rsAssDel  = Server.CreateObject("ADODB.RecordSet")
		sqlAssDel = "SELECT * FROM InstIntr_T WHERE IntrID = " & Request("selIntr")
		rsAssDel.Open sqlAssDel, g_strCONN, 1, 3
		If Not rsAssDel.EOF Then
			If Request("IntrCtr") <> 0 Then 
				ctr = Request("IntrCtr")
				For i = 0 to ctr 
					tmpctr = Request("chkAss" & i)
					If tmpctr <> "" Then
						strTmp = "index= " & tmpctr 
						rsAssDel.Movefirst
						rsAssDel.Find(strTmp)
						If Not rsAssDel.EOF Then
							rsAssDel.Delete
							rsAssDel.Update
						End If
					End If
				Next
			End If 
		End If
		rsAssDel.Close
		Set rsAssDel = Nothing 
	End If
	Session("MSG") = "Interpreter Saved."
	Response.Redirect "a_Intr.asp?Use=" & Request("selUser")
ElseIf Request("ctrl") = 13 Then 'admin tools - delete user
	Set rsDel =Server.CreateObject("ADODB.RecordSet") 
	sqlDel = "SELECT * FROM User_T WHERE [index] = " & Request("selUser")
	rsDel.Open sqlDel, g_strCONN, 1, 3
	If Not rsDel.EOF Then
		tmptype = rsDel("type")
		tmpInst = rsDel("InstID")
		tmpIntr = rsDel("IntrLB")
		tmpRP = rsDel("ReqLB")
		tmpName = rsDel("Lname") & ", " & rsDel("Fname")
		rsDel.Delete
		rsDel.Update
	End If
	rsDel.CLose
	Set rsDel = Nothing
	If tmptype = 0 Then
		Set rsDels = Server.CreateObject("ADODB.RecordSet") 
		sqlDel = "DELETE FROM InstIntr_T WHERE InstID = " & tmpInst
		rsDels.Open sqlDel, g_strCONN, 1, 3
		SetrsDels = Nothing
	ElseIf tmptype = 1 Then
		Set rsDels = Server.CreateObject("ADODB.RecordSet") 
		sqlDel = "DELETE FROM InstIntr_T WHERE IntrID = " & tmpIntr
		rsDels.Open sqlDel, g_strCONN, 1, 3
		Set rsDels = Nothing
	End If
	Session("MSG") = "User " & tmpName & " has been deleted."
	Response.Redirect "a_user.asp"
ElseIf Request("ctrl") = 14 Then 'admin tools - reason
	If Request("selReas") <> "" Then 'delete reason
		tmpReas = Split(Request("selReas"), ",")
		tmpCtr = Ubound(tmpReas)
		ctr = 0
		Do until ctr = tmpCtr + 1
			Set rsDel = Server.CreateObject("ADODB.RecordSet")
			sqlDel = "DELETE FROM Reason_T WHERE [index] = " & tmpReas(ctr)
			rsDel.Open sqlDel, g_strCONN, 1, 3
			Set rsDel = Nothing
			ctr = ctr + 1
		Loop
		Session("MSG") = Session("MSG") & "<br>Reason/s deleted."
	End If
	If Request("txtReas") <> "" Then 'new reason
		Set rsReas = Server.CreateObject("ADODB.RecordSet")
		sqlReas = "SELECT * FROM Reason_T WHERE Upper(reason) = '" & Ucase(Trim(Request("txtReas"))) &"' AND deptID = " & Request("selInst")
		rsReas.Open sqlReas, g_strCONN, 1, 3
		If rsReas.EOF Then
			rsReas.AddNew
			rsReas("reason") = Trim(Request("txtReas"))
			rsReas("deptID") = Request("selInst")
			rsReas.Update
			Session("MSG") = "New Reason saved."
		Else
			Session("MSG") = Session("MSG") & "<br>ERROR: Reason already exists for this department."
		End If
		rsReas.Close
		Set rsReas = Nothing
	End If
	Response.Redirect "a_reason.asp?use=" & Request("selUser") & "&dept=" & Request("selInst")
ElseIf Request("ctrl") = 15 Then
	If Request("selUser") <> 0 And Request("selInst") <> 0 And Request("selDept") <> 0 And Request("selReq") <> 0 Then
 		Set rsUser = Server.CreateObject("ADODB.RecordSet")
 		sqlUser = "SELECT * FROM User_T WHERE [index] = " & Request("selUser")
 		rsUser.Open sqlUser, g_strCONN, 1, 3
 		If Not rsUser.EOF Then
 			rsUser("InstID") = Request("selInst")
 			rsUser("ReqLB") = Request("selReq")
 			rsUser("DeptLB") = Request("selDept")
 			rsUser.Update
 		End If
 		rsUser.Close
 		Set rsUser = Nothing
 		'ASSOCIATE REQUESTING PERSON
 		If Request("chkAll") <> "" Then
 			Set rsReq = Server.CreateObject("ADODB.RecordSet")
 			sqlReq = "SELECT * FROM reqDept_T WHERE ReqID = " & Request("selReq") & " AND DeptID = " & Request("selDept")
 			rsReq.Open sqlReq, g_strCONNLB, 1, 3
 			If rsReq.EOF Then
 				rsReq.AddNew
 				rsReq("ReqID") = Request("selReq")
 				rsReq("DeptID") = Request("selDept")
 				rsReq.Update
 			End If
 			rsReq.Close
 			Set rsReq = Nothing
 		End If
	End If
	Response.Redirect "a_dept.asp?use=" & Request("selUser")
ElseIf Request("ctrl") = 16 Then
	Response.Cookies("LBREPORT") = Z_DoEncrypt(2 & "|" & Request("xdate") & "|" & Request("type"))
	Response.redirect "calendarview2.asp?rpt=1"
ElseIf Request("ctrl") = 17 Then
	'MEDICAID EDITING
	Response.Cookies("LBREQUEST") = Z_DoEncrypt(Request("txtClilname")	& "|" & _
		Request("txtClifname")	& "|" & Request("selReas")	& "|" & Request("chkCall")	& "|" & Request("selDept")	& "|" & _
		Request("txtCliFon")	& "|" & Request("txtCliMobile")	& "|" & _
		Request("selLang") & "|" & Request("txtAppDate")	& "|" & Request("txtAppTFrom")	& "|" & Request("txtAppTTo") & "|" & Request("txtcom") & "|" & _
		Request("txtClinName")	& "|" & Request("txtTP") & "|" & Request("radioLang") & "|" & Request("chkminor") & _
		"|" & Request("txtparents") & "|" & Request("txtLBcom") & "|" & Request("txtIntrcom") & "|" & Request("selGen") & "|" & _
		Request("chkClientAdd")	& "|" & Request("txtCliAddrI") & "|" & Request("txtCliadd") & "|" & Request("txtClicity") & "|" & _
		Request("txtClistate")	& "|" & Request("txtCliZip") & "|" & Request("txtRFon") & "|" & Request("txtdhhsemail") & "|" & _
		Request("txtdhhsFon") & "|" & Request("txtoLang") & "|" & Request("chkCall2") & "|" & Request("txtchrg") & "|" & Request("txtAtrny") & "|" & _
		Request("txtDOB") & "|" & Request("txtPDamount") & "|" & Request("h_tmpfilename") & "|" & Request("chkout") & "|" & Request("chkmed") & "|" & _
		Request("MCnum") & "|" & Request("chkacc") & "|" & Request("chkcomp") & "|" & Request("selIns") & "|" & Request("MHPnum") & "|" & Request("NHHFnum") & "|" & Request("WSHPnum") & "|" & Request("chkawk"))
	If Request("txtDOB") <> "" Then
		If Not IsDate(Request("txtDOB")) Then Session("MSG") = Session("MSG") & "<br>ERROR: Invalid Date of Birth."
	End If
	If Session("MSG") = "" Then	
		'GET COOKIE OF REQUEST
		tmpEntry = Split(Z_DoDecrypt(Request.Cookies("LBREQUEST")), "|")
		'SAVE ENTRIES
		Set rsMain = Server.CreateObject("ADODB.RecordSet")
		sqlMain = "SELECT * FROM appointment_T WHERE [index] = " & Request("ID")
		rsMain.Open sqlMain, g_strCONN, 1, 3
		If Not rsMain.EOF Then
			rsMain("DOB") = Z_DateNull(tmpEntry(33))
			rsMain("outpatient") = False
			If tmpEntry(36) <> "" Then rsMain("outpatient") = True
			rsMain("hasmed") = False
			If tmpEntry(37) <> "" Then rsMain("hasmed") = True
			rsMain("medicaid") = tmpEntry(38)
			rsMain("autoacc") = False
			If tmpEntry(39) <> "" Then rsMain("autoacc") = True
			rsMain("wcomp") = False
			If tmpEntry(40) <> "" Then rsMain("wcomp") = True
			rsMain("secins") = Z_FixNull(tmpEntry(41))
			rsMain("meridian") = tmpEntry(42)
			rsMain("nhhealth") = tmpEntry(43)
			rsMain("wellsense") = tmpEntry(44)
			rsMain("acknowledge") = false
			If tmpEntry(45) <> "" Then rsMain("acknowledge") = True
			rsMain.Update
			'GET ID FOR CONFIRM
			tmpID = rsMain("index")
		End If
		rsMain.Close
		Set rsMain = Nothing
		Set rsLB = Server.CreateObject("ADODB.RecordSet")
		sqlLB = "SELECT * FROM Request_T WHERE HPID = " & Request("ID")
		rsLB.Open sqlLB, g_strCONNLB, 1, 3
		If not rsLB.EOF Then 
			rsLB("DOB") = Z_dateNull(tmpEntry(33))
			rsLB("outpatient") = False
			If tmpEntry(36) <> "" Then rsLB("outpatient") = True
			rsLB("hasmed") = False
			If tmpEntry(37) <> "" Then rsLB("hasmed") = True
			rsLB("medicaid") = tmpEntry(38)
			rsLB("autoacc") = False
			If tmpEntry(39) <> "" Then rsLB("autoacc") = True
			rsLB("wcomp") = False
			If tmpEntry(40) <> "" Then rsLB("wcomp") = True
			rsLB("secins") = Z_FixNull(tmpEntry(41))
			rsLB("meridian") = tmpEntry(42)
			rsLB("nhhealth") = tmpEntry(43)
			rsLB("wellsense") = tmpEntry(44)
			rsLB("acknowledge") = false
			If tmpEntry(45) <> "" Then rsLB("acknowledge") = True
			rsLB.Update
		End If
		tmpLBID = rsLB("index")
		rsLB.Close
		set rsLB = Nothing
		If SaveHist(tmpLBID, "[HP]main.asp") Then
	
		End If
		Session("MSG") = "Medicaid/MCO Information saved."
		If SaveHist(tmpLBID, "[HP]main.asp") Then
	
		End If
		Response.Redirect "reqconfirm.asp?ID=" & tmpID
	Else
		Response.Redirect "main.asp?ID=" & Request("ID")	
	End If	
End If
%>
<!-- #include file="_closeSQL.asp" -->