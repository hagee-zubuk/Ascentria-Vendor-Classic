<%Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="main_helper.asp" -->
<%Response.AddHeader "Pragma", "No-Cache" %>
<%
'USER CHECK
'response.write "CLASS: " & Session("myClass")
If Cint(Session("type")) = 1 And Session("UID") <> 35 Then
	Session("MSG") = "Error: Invalid user type. Please sign-in again."
	Response.Redirect "default.asp"
End If

CalendarPage = False
MeetTayo = "Checked"
MeetTayo2 = ""
tmpCall = ""
radioLang0 = "checked"
radioLang1 = ""
radioLang2 = ""
myStat = 0
billedako = False
Billedna = ""
chkTeleh = ""
'EDIT APPOINTMENT
If Request("ID") <> "" Then
	EditPage = " - Edit"
	Set rsApp = Server.CreateObject("ADODB.RecordSet")
	sqlApp = "SELECT * FROM Appointment_T WHERE [index] = " & Request("ID")
	rsApp.Open sqlApp, g_strCONN, 3, 1
	If Not rsApp.EOF Then
		myStat = GetStatLB(Request("ID"))
		tmpID = rsApp("index")	
		billedako = ChkBilled(tmpID)
		'If billedako Then Billedna = "disabled"
		tmpClname = Z_DoDecrypt(rsApp("clname")	)
		tmpCfname =  Z_DoDecrypt(rsApp("cfname"))
		tmpdept = rsApp("deptID")
		tmpCFon = Z_DoDecrypt(rsApp("phone"))
		tmpCFon2 = Z_DoDecrypt(rsApp("mobile"))
		tmpLang = rsApp("LangID")
		If tmpLang = 98 Then tmpoLang = rsApp("oLang")
		tmpAppDate = rsApp("appDate")
		tmpFtime = Z_FormatTime(rsApp("TimeFrom"))
		tmpTtime = Z_FormatTime(rsApp("TimeTo")) 
		tmpIntr = rsApp("IntrID")
		tmpCom = rsApp("Comment")
		tmpReas = rsApp("Reason")
		tmpCall = "checked"
		If rsApp("callme") = False Then tmpCall = ""
		chkleave = "checked"
		If rsApp("leavemsg") = False Then chkleave = ""
		tmpsc = rsApp("Spec_cir")			
		tmpminor = ""
		If rsApp("minor") = True Then tmpminor = "checked"
		tmpPar = rsApp("parents") 
		tmpreqname = rsApp("reqname") 
		tmpLBcom = rsApp("lbcom") 
		tmpIntrcom = rsApp("intrcom") 
		If rsApp("Gender") <> vbNull Then
			tmpGender	= rsApp("Gender")
			tmpMale = ""
			tmpFemale = ""
			If tmpGender = 0 Then 
				tmpMale = "SELECTED"
			ElseIf tmpGender = 1 Then 
				tmpFemale = "SELECTED"
			End If
		End If
		tmpDOB = rsApp("DOB") 
		chkUClientadd = ""
		If rsApp("useCadr") = true Then chkUClientadd = "checked"
		If rsApp("telehealth") = TRUE Then chkTeleh = "CHECKED"
		tmpCAdrI =  rsApp("capt")
		tmpCAddr =  rsApp("caddress")
		tmpCity = rsApp("ccity")
		tmpState =  rsApp("cstate")
		tmpZip =  rsApp("czip")
		tmpRFon = rsApp("rphone")
		tmpdemail = rsApp("semail")
		tmpdFon =  rsApp("sphone")
		myInst = rsApp("InstLB")
		tmpcall2 = ""
		If rsApp("block") = true then tmpcall2 = "checked"
		If Session("myClass") <> 3 Then
			tmpCli = rsApp("clinician") 
			tmpPage = Z_FormatTime(rsApp("paged"))
		Else
			tmpCli = rsApp("docknum") 
			tmpPage = rsApp("crtroom")
			tmpChrge = rsApp("charges") 
			tmpAtrny = rsApp("attny")
		End If
		If Session("type") = 5 Then
			tmpPDAmount = rsApp("PDAmount") 
			tmpPDemail = rsApp("PDemail")
			tmpfileuploaded = ""
			If rsApp("uploadfile") Then tmpfileuploaded = "*Form 604A already uploaded. Uploading another file will remove the previous uploaded file." 
			disUpload = ""	
			If rsApp("approvePDF") Then 
				disUpload = "disabled"	
				tmpfileuploaded = "*Form 604A already approved."
			End If
		End If
		chkout = ""
		If rsApp("outpatient") Then chkout = "CHECKED"
		chkmed = ""
		If rsApp("hasmed") Then chkmed = "CHECKED"
		If Trim(rsApp("medicaid")) <> "" Then
			MCNum = rsApp("medicaid")
			'radiomed4 = "checked"
		End If
		If Trim(rsApp("amerihealth")) <> "" Then
			AHMIdNum = rsApp("amerihealth")
			radiomed5 = "checked"
		End If
		If Trim(rsApp("meridian")) <> "" Then
			MHPnum = rsApp("meridian")
			radiomed1 = "checked"
		End If
		If Trim(rsApp("nhhealth")) <> "" Then
			NHHFnum = rsApp("nhhealth")
			radiomed2 = "checked"
		End If
		If Trim(rsApp("wellsense")) <> "" Then
			WSHPnum = rsApp("wellsense")
			radiomed3 = "checked"
		End If
		If Trim(rsApp("medicaid")) <> "" And Trim(rsApp("meridian")) = "" And Trim(rsApp("nhhealth")) = "" And _
		Trim(rsApp("wellsense")) = "" THen radiomed4 = "checked"
		chkawk = ""
		If rsApp("acknowledge") Then chkawk = "Checked"
		chkAppMed = ""
		If rsApp("vermed") Then chkAppMed = "CHECKED"
		chkacc = ""
		If rsApp("autoacc") Then chkacc = "CHECKED"
		chkcomp = ""
		If rsApp("wcomp") Then chkcomp = "CHECKED"
		tmpsecins = rsApp("secins")
		mrrec = rsApp("mrrec")
	End If
	rsApp.Close
	Set rsApp = Nothing
	'CHECK IF ALLOWED TO VIEW
	If Session("InstID") <> myInst Then
		Session("MSG") = "Error: You are not allowed to view this appointment. Please sign-in again."
		Response.Redirect "default.asp"
	End If
End If

If Session("MSG") <> "" And Request("ID") = "" Then
	tmpEntry = Split(Z_DoDecrypt(Request.Cookies("LBREQUEST")), "|")
	tmpClname = tmpEntry(0)
	tmpCfname = tmpEntry(1)
	tmpCFon = tmpEntry(5)
	tmpCFon2 = tmpEntry(6)
	tmpCli = tmpEntry(12)
	tmpPage =tmpEntry(13)
	tmpLang = tmpEntry(7)
	tmpAppDate = tmpEntry(8)
	tmpFtime = tmpEntry(9)
	tmpTtime =tmpEntry(10)
	'tmpIntr = rsApp("IntrID")
	tmpCom = tmpEntry(11)
	'tmpReas = tmpEntry(2)
	tmpDept = tmpEntry(4)
	tmpCall = "checked"
	If tmpEntry(3) = "" Then tmpCall = ""
	radioLang0 = ""
	If tmpEntry(14) = 0 Then radioLang0 = "checked"
	If tmpEntry(14) = 1 Then radioLang1 = "checked"
	If tmpEntry(14) = 2 Then radioLang2 = "checked"
	If tmpEntry(14) = 3 Then radioLang3 = "checked"
	If tmpEntry(14) = 4 Then radioLang4 = "checked"
	If tmpEntry(14) = 5 Then radioLang5 = "checked"
	tmpminor = ""
	If tmpEntry(15) <> "" Then tmpminor = "checked"
	tmppar = tmpEntry(16)
	tmpreqname = tmpEntry(17)
	If tmpEntry(20) >= 0 Then
		tmpGender	= Z_CZero(tmpEntry(20))
		tmpMale = ""
		tmpFemale = ""
		If tmpGender = 0 Then 
			tmpMale = "SELECTED"
		ElseIf tmpGender = 1 Then 
			tmpFemale = "SELECTED"
		End If
	End If

	tmpDOB = tmpEntry(34)
	chkUClientadd = ""
	if Z_Czero(tmpEntry(21)) = 1 then chkUClientadd = "checked"
	tmpCAdrI =  tmpEntry(22)
	tmpCAddr =  tmpEntry(23)
	tmpCity = tmpEntry(24)
	tmpState =  tmpEntry(25)
	tmpZip =  tmpEntry(26)
	tmpRFon = tmpEntry(27)
	tmpdemail = tmpEntry(28)
	tmpdFon = tmpEntry(29)
	tmpCall2 = ""
	if tmpEntry(31) <> "" Then tmpCall2 = "checked"
End If
'CLONE REQUEST
If Request("Clone") <> "" Then
	Set rsClone = Server.CreateObject("ADODB.RecordSet")
	sqlClone = "SELECT * FROM Appointment_T WHERE [index] = " & Request("Clone")
	rsClone.Open sqlClone, g_strCONN, 3, 1
	If Not rsCLone.EOF Then
		tmpReqP = rsClone("ReqID") 
		tmpminor = ""
		If rsClone("minor") = True Then tmpminor = "checked"
		tmpPar = rsClone("parents") 
		tmpClname = Z_DoDecrypt(rsClone("clname")	)
		tmpCfname =  Z_DoDecrypt(rsClone("cfname"))
		tmpdept = rsClone("deptID")
		tmpCFon = Z_DoDecrypt(rsClone("phone"))
		tmpCFon2 = Z_DoDecrypt(rsClone("mobile"))
		If rsClone("useCadr") Then
			chkUClientadd = "checked"
			tmpCAdrI = rsClone("capt")
			tmpCAddr = rsClone("caddress")
			tmpCity = rsClone("ccity")
			tmpState = rsClone("cstate")
			tmpZip = rsClone("czip")
		End If
		If Session("type") <>  5  Then
			If Session("myClass") <> 3 Then
				tmpCli = rsClone("clinician") 
				tmpPage = Z_FormatTime(rsClone("paged"))
			Else
				tmpCli = rsClone("docknum") 
				tmpPage = rsClone("crtroom")
				tmpChrge = rsClone("charges") 
				tmpAtrny = rsClone("attny") 
			end if
		Else
			tmpCli = rsClone("docknum")
			tmpPDAmount = rsClone("PDAmount")
		End If
		tmpLang = rsClone("LangID")
		tmpAppDate = rsClone("appDate")
		tmpFtime = Z_FormatTime(rsClone("TimeFrom"))
		tmpTtime = Z_FormatTime(rsClone("TimeTo")) 
		tmpIntr = rsClone("IntrID")
		tmpCom = rsClone("Comment")
		tmpReas = rsClone("Reason")
		
		If rsClone("Gender") <> vbNull Then
			tmpGender	= Z_CZero(rsClone("Gender"))
			tmpMale = ""
			tmpFemale = ""
			If tmpGender = 0 Then 
				tmpMale = "SELECTED"
			ElseIf tmpGender = 1 Then 
				tmpFemale = "SELECTED"
			End If
		End If
		tmpCall = "checked"
		If rsClone("callme") = False Then tmpCall = ""
		chkleave = "checked"
		If rsClone("leavemsg") = False Then chkleave = ""
		tmpcall2 = "checked"
		if rsclone("block") = false then tmpcall2 = ""
		tmpDept = rsClone("deptID")
		tmpreqname = rsClone("reqname")
		tmpRFon = rsClone("rphone")
		tmpDOB = rsClone("DOB")
		'tmpemail = rsClone("email")
		'tmpHPID = rsClone("HPID")
		If Session("myClass") <> 3 Then
			tmpCli = rsClone("clinician") 
			tmpPage = Z_FormatTime(rsClone("paged"))
		Else
			tmpCli = rsClone("docknum") 
			tmpPage = rsClone("crtroom")
			tmpChrge = rsClone("charges") 
			tmpAtrny = rsClone("attny")
		End If
		If Session("type") = 5 Then
			tmpPDAmount = rsClone("PDAmount") 
			tmpPDemail = rsClone("PDemail")
		
		End If
		chkout = ""
		If rsClone("outpatient") Then chkout = "CHECKED"
		chkmed = ""
		If rsClone("hasmed") Then chkmed = "CHECKED"
		If Trim(rsClone("medicaid")) <> "" Then
			MCNum = rsClone("medicaid")
			'radiomed4 = "checked"
		End If
		If Trim(rsClone("amerihealth")) <> "" Then
			AHMIdNum = rsClone("amerihealth")
			radiomed5 = "checked"
		End If
		If Trim(rsClone("meridian")) <> "" Then
			MHPnum = rsClone("meridian")
			radiomed1 = "checked"
		End If
		If Trim(rsClone("nhhealth")) <> "" Then
			NHHFnum = rsClone("nhhealth")
			radiomed2 = "checked"
		End If
		If Trim(rsClone("wellsense")) <> "" Then
			WSHPnum = rsClone("wellsense")
			radiomed3 = "checked"
		End If
		If Trim(rsClone("medicaid")) <> "" And Trim(rsClone("meridian")) = "" And Trim(rsClone("nhhealth")) = "" And _
		Trim(rsClone("wellsense")) = "" THen radiomed4 = "checked"
		chkawk = ""
		If rsClone("acknowledge") Then chkawk = "Checked"
		chkAppMed = ""
		If rsClone("vermed") Then chkAppMed = "CHECKED"
		chkacc = ""
		If rsClone("autoacc") Then chkacc = "CHECKED"
		chkcomp = ""
		If rsClone("wcomp") Then chkcomp = "CHECKED"
		tmpsecins = rsClone("secins")
		mrrec = rsClone("mrrec")
		tmpsc = rsClone("Spec_cir")		
		chkleave = ""
		If rsClone("leavemsg") Then chkleave = "CHECKED"	
		Session("MSG") = "NOTE: Entries cloned from Request: " & Request("Clone")
	End If
	rsClone.CLose
	Set rsClone = Nothing
End If
'GET allowed mco
Set rsInst = Server.CreateObject("ADODB.RecordSet")
sqlInst = "SELECT * FROM mco_T"
rsInst.Open sqlInst, g_strCONNLB, 3, 1
If Not rsInst.EOF Then
	Do Until rsInst.EOF 
		If rsInst("mco") = "Medicaid" Then
			If rsInst("active") Then
				allowMCO = allowMCO & "ff.rdoMed_Med.disabled = false; " & vbCrLf 
			Else
				allowMCO = allowMCO & "ff.rdoMed_Med.disabled = true; " & vbCrLf 
			End If
		End If
		'If rsInst("mco") = "AmeriHealth" Then
		''	If rsInst("active") Then
		''		allowMCO = allowMCO & "ff.rdoMed_Ame.disabled = false; " & vbCrLf 
		''	Else
		''		allowMCO = allowMCO & "ff.rdoMed_Ame.disabled = true; " & vbCrLf 
		''	End If
		'End If
		If rsInst("mco") = "Meridian" Then
			If rsInst("active") Then
				allowMCO = allowMCO & "ff.rdoMed_Mer.disabled = false; " & vbCrLf 
			Else
				allowMCO = allowMCO & "ff.rdoMed_Mer.disabled = true; " & vbCrLf 
			End If
		End If
		If rsInst("mco") = "NHhealth" Then
			If rsInst("active") Then
				allowMCO = allowMCO & "ff.rdoMed_NHH.disabled = false; " & vbCrLf 
			Else
				allowMCO = allowMCO & "ff.rdoMed_NHH.disabled = true; " & vbCrLf 
			End If
		End If
		If rsInst("mco") = "WellSense" Then
			If rsInst("active") Then
				allowMCO = allowMCO & "ff.rdoMed_Wel.disabled = false; " & vbCrLf 
			Else
				allowMCO = allowMCO & "ff.rdoMed_Wel.disabled = true; " & vbCrLf 
			End If
		End If
		rsInst.MoveNext
	Loop
End If
allowMCO = allowMCO & "ff.rdoMed_Ame.disabled = false; " & vbCrLf 
rsInst.Close
Set rsInst = Nothing
'response.write "DEPT: " & tmpDept
'GET INST V.2
Set rsInst = Server.CreateObject("ADODB.RecordSet")
sqlInst = "SELECT * FROM institution_T WHERE [index] = " & Session("InstID")
rsInst.Open sqlInst, g_strCONNLB, 3, 1
If Not rsInst.EOF Then
	tmpInst = rsInst("Facility")		
End If
rsInst.Close
Set rsInst = Nothing
'GET DEPTS V.2
If Session("type") <> 3 Then
	blnSkipIncl = TRUE
	sqlDept = "SELECT d.[index], d.[dept], d.[BLname]" & _
		", d.[address], d.[BAddress]" & _
		", d.[InstAdrI], d.[drg]" & _
		", d.[City], d.[BCity]" & _
		", d.[State], d.[Bstate]" & _
		", d.[Zip], d.[BZip]" & _
		" FROM [interpreterSQL].[dbo].[xr_user_dept] AS x INNER JOIN [langbank].[dbo].[dept_T] AS d ON x.[dept_id]=d.[index] " & _
		" WHERE x.[user_id]=" & Session("UID")
	Set rsDept = Server.CreateObject("ADODB.RecordSet")
	rsDept.Open sqlDept, g_strCONNLB, 3, 1
	If rsDept.EOF Then
		blnSkipIncl = FALSE
		' WHAT THE FURCK?!? 
		sqlDept = "SELECT * FROM Dept_T WHERE active = 1 and InstID = " & Session("InstID") & " ORDER BY dept"
		rsDept.Close
		rsDept.Open sqlDept, g_strCONNLB, 3, 1
	End If
	Do Until rsDept.EOF	
		If Not RemoveDept(rsDept("index"), blnSkipIncl) Then
			tmpSelDept = ""
			If Z_CZero(tmpDept) = rsDept("index") Then tmpSelDept = "selected"
			If Z_IncludeDept(Session("UID"), rsDept("index"), blnSkipIncl) Then 'special rule for user 509 and 510
				strDept = strDept & "<option " & tmpSelDept & " value='" & rsDept("index") & "'>" & rsDept("dept") & "</option>" & vbCrLf
				tmpAddr = rsDept("InstAdrI") & ", " & rsDept("address") & ", " & rsDept("city") & ", " & rsDept("state") & ", " & rsDept("zip")
				tmpBAddr =  rsDept("Baddress") & ", " & rsDept("Bcity") & ", " & rsDept("Bstate") & ", " & rsDept("Bzip")
				strDept2 = strDept2 & "if (dept == " & rsDept("Index") & ") " & vbCrLf & _
					"{document.frmMain.txtInstAddr.value = """ & tmpAddr & """; " & vbCrLf & _
					"document.frmMain.selDept.value = " & rsDept("Index") & "; " & vbCrLf & _
					"document.frmMain.txtBlname.value = """ & rsDept("BLname") & """; " & vbCrLf & _
					"document.frmMain.h_drg.value = """ & rsDept("drg") & """; " & vbCrLf & _
					"document.frmMain.txtBilAddr.value = """ & tmpBAddr & """; " & vbCrLf
				strDept2 = strDept2 & " }" & vbCrLf
			End If
		End If
		rsDept.MoveNext
	Loop
	rsDept.Close
	Set rsDept = Nothing
Else
	sqlDept = "SELECT d.[index], d.[dept], d.[BLname]" & _
		", d.[address], d.[BAddress]" & _
		", d.[InstAdrI], d.[drg]" & _
		", d.[City], d.[BCity]" & _
		", d.[State], d.[Bstate]" & _
		", d.[Zip], d.[BZip]" & _
		" FROM [interpreterSQL].[dbo].[xr_user_dept] AS x INNER JOIN [langbank].[dbo].[dept_T] AS d ON x.[dept_id]=d.[index] " & _
		" WHERE x.[user_id]=" & Session("UID")
	Set rsDept = Server.CreateObject("ADODB.RecordSet")
	rsDept.Open sqlDept, g_strCONNLB, 3, 1
	If rsDept.EOF Then
		sqlDept = "SELECT * FROM Dept_T WHERE [index] = " & Session("DeptID") & " ORDER BY dept"
		rsDept.Close
		rsDept.Open sqlDept, g_strCONNLB, 3, 1
	End If
	Do Until rsDept.EOF
		tmpSelDept = ""
		tmpDept = Session("DeptID") 
		If Z_CZero(tmpDept) = rsDept("index") Then tmpSelDept = "selected"
		strDept = strDept & "<option " & tmpSelDept & " value='" & rsDept("index") & "'>" & rsDept("dept") & "</option>" & vbCrLf
		
		tmpAddr = rsDept("InstAdrI") & ", " & rsDept("address") & ", " & rsDept("city") & ", " & rsDept("state") & ", " & rsDept("zip")
		tmpBAddr =  rsDept("Baddress") & ", " & rsDept("Bcity") & ", " & rsDept("Bstate") & ", " & rsDept("Bzip")
		strDept2 = strDept2 & "if (dept == " & rsDept("Index") & ") " & vbCrLf & _
			"{document.frmMain.txtInstAddr.value = """ & tmpAddr &"""; " & vbCrLf & _
			"document.frmMain.selDept.value = " & rsDept("Index") & "; " & vbCrLf & _
			"document.frmMain.txtBlname.value = """ & rsDept("BLname") &"""; " & vbCrLf & _
			"document.frmMain.h_drg.value = """ & rsDept("drg") &"""; " & vbCrLf & _
			"document.frmMain.txtBilAddr.value = """ & tmpBAddr &"""; }" & vbCrLf
		
		rsDept.MoveNext
	Loop
	rsDept.Close
	Set rsDept = Nothing
End If

'GET REQUESTER INFO V.2
Set rsRP = Server.CreateObject("ADODB.RecordSet")
sqlRP  = "SELECT Email, phone, pExt, fax FROM Requester_T WHERE [index] = " & Session("ReqID")
rsRP.Open sqlRP, g_strCONNLB, 3, 1
If Not rsRp.EOF Then
	tmpEmail = rsRp("Email")
	tmpIphone = rsRp("phone")
	If rsRP("pExt") <> "" Then tmpIphone = tmpIphone & " ext. " & rsRP("pExt")
	tmpfax = rsRP("fax")
End If
rsRP.Close
Set rsRP = Nothing 
'GET USER
Set rsUser = Server.CreateObject("ADODB.RecordSet")
sqlUSer = "SELECT lname, fname FROM User_T WHERE [index] = " & Session("UID")
rsUser.Open sqlUser, g_strCONN, 3, 1
If Not rsUser.EOF Then
	tmpUname = rsUser("lname") & ", " & rsUser("fname")
End If
rsUser.Close
Set rsUSer = Nothing
'GET AVAILABLE LANGUAGES
Set rsLang = Server.CreateObject("ADODB.RecordSet")
sqlLang = "SELECT [index], [Lang] FROM Lang_T WHERE [index] <> 105 ORDER BY [Lang]"
rsLang.Open sqlLang, g_strCONN, 3, 1
Do Until rsLang.EOF
	tmpL = ""
	If tmpLang = "" Then tmpLang = 0
	If CInt(tmpLang) = rsLang("index") Then tmpL = "selected"
	If Request("ID") = "" Then ' dont allow other lang
		If rsLang("index") <> 98 Then strLang = strLang	& "<option " & tmpL & " value='" & rsLang("Index") & "'>" &  rsLang("Lang") & "</option>" & vbCrlf
	Else
		If tmpLang = 98 And rsLang("index") = 98 Then
			strLang = strLang	& "<option " & tmpL & " value='" & rsLang("Index") & "'>" &  rsLang("Lang") & "</option>" & vbCrlf
		Else
			If rsLang("index") <> 98 Then strLang = strLang	& "<option " & tmpL & " value='" & rsLang("Index") & "'>" &  rsLang("Lang") & "</option>" & vbCrlf
		End If
	End If
	rsLang.MoveNext
Loop
rsLang.Close
Set rsLang = Nothing
'GET REASON
tmpReasKo = split(tmpReas, ", ")
tmpReasCount = Ubound(tmpReasKo)
Set rsDept3 = Server.CreateObject("ADODB.RecordSet")
If Session("type") <> 3 Then
	sqlDept3 = "SELECT * FROM Dept_T WHERE InstID = " & Session("InstID") & " ORDER BY dept"
Else
	sqlDept3 = "SELECT * FROM Dept_T WHERE [index] = " & Session("DeptID") & " ORDER BY dept"
End If
rsDept3.Open sqlDept3, g_strCONNLB, 3, 1
Do Until rsDept3.EOF
	strDept3 = strDept3 & "if(dept == " & rsDept3("index") & "){" & vbCrLf 
	Set rsReas = Server.CreateObject("ADODB.RecordSet")
	sqlReas = "SELECT * FROM Reason_T WHERE deptID = " & rsDept3("index") & " ORDER BY Reason"
	rsReas.Open sqlReas, g_strCONN, 3, 1
	ctr = 0
	Do Until rsReas.EOF
		tmpReas = rsReas("Reason")
		strDept3 = strDept3 & "var ChoiceRes = document.createElement('option');" & vbCrLf & _
			"ChoiceRes.value = " & rsReas("index") & ";" & vbCrLf & _
			"ChoiceRes.appendChild(document.createTextNode(""" & tmpReas & """));" & vbCrLf & _
			"document.frmMain.selReas.appendChild(ChoiceRes);" & vbCrLf
			x = 0
			Do Until x =  tmpReasCount + 1
				If Z_CZero(Trim(tmpReasKo(x))) = Z_CZero(rsReas("index")) Then
					strDept3 = strDept3 & "document.frmMain.selReas[" & ctr & "].selected = true;" & vbCrLf
				End If
				x = x + 1
			Loop
			ctr = ctr + 1
		rsReas.MoveNext
	Loop
	rsReas.Close
	Set rsReas = Nothing
	strDept3 = strDept3 & "}" & vbCrLf 
	rsDept3.MoveNext
Loop
rsDept3.Close
Set rsDept3 = Nothing
If Session("type") = 5 Then 'create temp filename
	tmpFilename = Z_GenerateGUID()
	Do Until GUIDExists(tmpFilename) = False
		tmpFilename = Z_GenerateGUID()
	Loop
End If
%>
<!-- #include file="_closeSQL.asp" -->
<html>
	<head>
		<title>Interpreter Request - Interpreter Request Form</title>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<link href="CalendarControl.css" type="text/css" rel="stylesheet">
		<script src="CalendarControl.js" language="javascript"></script>
		<script language='JavaScript'><!--
	function Left(str, n) {
		if (n <= 0)
		    return "";
		else if (n > String(str).length)
		    return str;
		else
		    return String(str).substring(0,n);
	}
	function RTrim(str) {
        var whitespace = new String(" \t\n\r");
        var s = new String(str);
        if (whitespace.indexOf(s.charAt(s.length-1)) != -1) {
            var i = s.length - 1;       
            while (i >= 0 && whitespace.indexOf(s.charAt(i)) != -1) i--;
            s = s.substring(0, i+1);
        }
        return s;
    }

    function LTrim(str) {
		var whitespace = new String(" \t\n\r");
		var s = new String(str);
		if (whitespace.indexOf(s.charAt(0)) != -1) {
			var j=0, i = s.length;
			while (j < i && whitespace.indexOf(s.charAt(j)) != -1) j++;
			s = s.substring(j, i);
		}
		return s;
    }

    function Trim(str) {
		return RTrim(LTrim(str));
    }

	function maskMe(str,textbox,loc,delim) {
		var locs = loc.split(',');
		for (var i = 0; i <= locs.length; i++) {
			for (var k = 0; k <= str.length; k++) {
				if (k == locs[i]) {
					if (str.substring(k, k+1) != delim) {
						str = str.substring(0,k) + delim + str.substring(k,str.length);
					}
				}
			}
		}
		textbox.value = str;
	}

	function ReqChkMe(xxx) {	
		if (document.frmMain.saveme.value == 'Save') {
	
			if (xxx == 2 || xxx == 3 || xxx == 4) {
				alert("ERROR: You cannot edit canceled/missed appointments.")
				return;
			}
<% If Session("type") <> 5 Then %>
			if (document.frmMain.selDept.value == 0) {
				alert("ERROR: Please provide a department.")
				return;
			}
<% Else %>
			if (document.frmMain.selDept.value == 0) {
				alert("ERROR: Please provide a court.")
				return;
			}
<% End If %>
<% If Session("type") = 3 Or Session("type") = 4 Or Session("type") = 5 Then %>
			if (document.frmMain.txtreqname.value == "") {
				alert("ERROR: Please provide a Requester's Name.")
				return;
			}
<% End If %>
			if (document.frmMain.txtClilname.value == "" || document.frmMain.txtClifname.value == "") {
				alert("ERROR: Please provide client.")
				return;
			}
			if ((document.frmMain.chkcall.checked == true || document.frmMain.chkleave.checked == true) && document.frmMain.txtCliFon.value == "") {
				alert("Please input client's phone number.");
				return;
			}
<% If Session("type") = 5 Then %>
			if (document.frmMain.txtClinName.value == "") {
				alert("ERROR: Please provide Docket Number.")
				return;
			}
			if (document.frmMain.txtPDamount.value == "") {
				alert("ERROR: Please provide Amount requested from court.")
				return;
			}
			if (document.frmMain.txtemail.value == "") {
				alert("ERROR: Please provide email.")
				return;
			}
			//alert(document.frmMain.F1.value);
<% End If %>
<% If Session("myClass") = 4 Or Session("myClass") = 6 Then %>
			if (document.frmMain.mrrec.value == "") {
				alert("ERROR: Please provide Patient MR#.")
				return;
			}
			if (document.frmMain.txtCliCir.value == "") {
				alert("ERROR: Please precaution information or \"N/A\".")
				return;
			}
<% End If %>
<% If Request("ID") = "" Then %>
			if (document.frmMain.radioLang[0].checked == true) {
<% End If %>	
<% If Session("UID") <> 36 Then %> 
				if 	(document.frmMain.selLang.value == 0) {
					alert("ERROR: Please provide language.")
					return;
				}
<% End If %>
<% If Request("ID") = "" Then %>
			}
<% End If %>
<% If Request("ID") = "" Then %>
	<% If Session("InstID") = 108 Then %>
		<% If Session("UID") <> 36 Then %> 
			//if (document.frmMain.radioLang[5].checked == true && document.frmMain.txtoLang.value == "")
			//{
			//		alert("ERROR: Please provide language.")
			//		return;
			//}
		<% End If %>
	<% End If %>
<% Else %>
	<% If Session("UID") <> 36 Then %> 
			if 	(document.frmMain.selLang.value == 98 && document.frmMain.txtoLang.value == "") {
				alert("ERROR: Please provide language.")
				return;
			}
	<% End If %>
<% End If %>
			if 	(document.frmMain.txtAppDate.value == "") {
				alert("ERROR: Please provide appointment date.")
				return;
			}
			if (Trim(document.frmMain.txtCliAdd.value) != "" ||
					Trim(document.frmMain.txtCliCity.value) != "" ||
					Trim(document.frmMain.txtCliState.value) != "" ||
					Trim(document.frmMain.txtCliZip.value) != ""
					) {
				if (document.frmMain.chkClientAdd.checked == false) {
					alert("Alternate Appointment Address detected. If you wish to make this address as the" +
							"appointment address, please check the checkbox beside it. \nTHE INTERPRETER " +
							"WILL BE SENT TO THIS ADDRESS.");
					return;
				}
			}
			if 	(document.frmMain.chkClientAdd.checked == true) {
				if (Trim(document.frmMain.txtCliAdd.value) == "" ||
						Trim(document.frmMain.txtCliCity.value) == "" ||
						Trim(document.frmMain.txtCliState.value) == "" ||
						Trim(document.frmMain.txtCliZip.value) == ""
						) {
					alert("ERROR: Please provide Alternate Appointment's Full Address.")
					return;
				}
			}
<% If Session("InstID") = 108 Then %>
			if (document.frmMain.txtdhhsFon.value == "") {
				alert("ERROR: Please provide assigned DHHS staff email.")
				return;
			}
<% End If %>
			if 	(Trim(document.frmMain.txtAppTFrom.value) == "" || Trim(document.frmMain.txtAppTTo.value) == "") {
				alert("ERROR: Please provide appointment time.")
				return;
			}
			if (document.frmMain.txtAppTFrom.value == "24:00") {
				alert("ERROR: Appointment Time (From:) is invalid (24:00 not accepted)."); 
				return;
			}
			if (document.frmMain.txtAppTTo.value == "24:00") {
				alert("ERROR: Appointment Time (To:) is invalid (24:00 not accepted)."); 
				return;
			}
			var d1 = new Date(document.frmMain.txtAppDate.value + " " + document.frmMain.txtAppTFrom.value);
			var d2 = new Date(document.frmMain.txtAppDate.value + " " + document.frmMain.txtAppTTo.value);
			var stime = d1.getHours();
			var etime = d2.getHours() ;
			var difference = etime - stime;
			if 	(difference < 0) {
				alert("ERROR: Invalid time frame.")
				return;
			}
				
			if (document.frmMain.h_drg.value == "True") {
				if (document.frmMain.chkmed.checked == true) {
					if (document.frmMain.txtDOB.value == "") {
						alert("Please input client's date of birth.")
						return;
					}
					if (document.frmMain.rdoMed_Ame.checked == false &&
							document.frmMain.rdoMed_Med.checked == false &&
							document.frmMain.rdoMed_Mer.checked == false &&
							document.frmMain.rdoMed_NHH.checked == false &&
							document.frmMain.rdoMed_Wel.checked == false ) {
						alert("Please select a Medicaid/MCO.")
						return;
					}
					if (Trim(document.frmMain.MHPnum.value) == "" &&
							document.frmMain.rdoMed_Mer.checked == true) {
						alert("Please input client's Meridian Health Plan number.")
						return;
					}
					if (Trim(document.frmMain.NHHFnum.value) == "" && document.frmMain.rdoMed_NHH.checked == true) {
						alert("Please input client's NH Healthy Families number.")
						return;
					} else {
						if (Trim(document.frmMain.NHHFnum.value) != "") {
							var chrmed = Trim(document.frmMain.NHHFnum.value);
							if (chrmed.length != 11) {
								alert("Invalid NH Healthy Families number length(11).")
								return;
							}
						}
					}
					if (Trim(document.frmMain.AHMIdNum.value) == "" && document.frmMain.rdoMed_Ame.checked == true) {
						alert("Please input client's AmeriHealth Member ID number.")
						return;
					} else {
						if (Trim(document.frmMain.AHMIdNum.value) != "") {
							var chrmed = Trim(document.frmMain.AHMIdNum.value);
							if ((chrmed.length < 8) || (chrmed.length > 9)) {
								alert("Invalid AmeriHealth Member ID number length(8/9).")
								return;
							}
						}
					}
					if (Trim(document.frmMain.WSHPnum.value) == "" && document.frmMain.rdoMed_Wel.checked == true) {
						alert("Please input client's Well Sense Health Plan number.")
						return;
					} else {
						if (Trim(document.frmMain.WSHPnum.value) != "") {
							var chrmed = Trim(document.frmMain.WSHPnum.value);
							if (chrmed.length != 9) {
								alert("Invalid Well Sense Health Plans number length(9).")
								return;
							}
							var str = Left(document.frmMain.WSHPnum.value, 2)
							var res = str.toUpperCase(); 
							if (res != 'NH') {
								alert("Well Sense number MUST contain NH (eg: NHXXXXXXX).")
								return;
							}
						}
					}
					if (Trim(document.frmMain.MCnum.value) == "" && document.frmMain.rdoMed_Med.checked == true) {
						alert("Please input client's Medicaid number.")
						return;
					} else {
						if (Trim(document.frmMain.MCnum.value) != "") {
							var chrmed = Trim(document.frmMain.MCnum.value);
							if (chrmed.length != 11) {
								alert("Invalid Medicaid number length(11).")
								return;
							}
						}
					}
					if (document.frmMain.chkawk.checked == false) {
						alert("Acknowledge statement is required.")
						return;
					}
				} 
				if (	(
							Trim(document.frmMain.MCnum.value) == "" &&
							document.frmMain.rdoMed_Med.checked == true
						) || (
							Trim(document.frmMain.AHMIdNum.value) == "" &&
							document.frmMain.rdoMed_Ame.checked == true
						) || (
							Trim(document.frmMain.MHPnum.value) == "" &&
							document.frmMain.rdoMed_Mer.checked == true
						) || (
							Trim(document.frmMain.NHHFnum.value) == "" &&
							document.frmMain.rdoMed_NHH.checked == true
						) || (
							Trim(document.frmMain.WSHPnum.value) == "" &&
							document.frmMain.rdoMed_Wel.checked == true
						) ){
					var ans = window.confirm("This institution/department is qualified for Medicaid/MCO.\n" + 
							"Please make sure to fill up all necessary information to bill Medicaid/MCO.\n" +
							"Otherwise, this appointment will be billed to this institution.\nCancel to "+
							"enter Medicaid/MCO info.\nOK to continue saving.");
					if (ans) {
						document.frmMain.saveme.value = 'Please Wait...';
<% If Request("ID") = "" Then%>
						document.frmMain.action = 'action.asp?ctrl=1';
<% Else %>
						document.frmMain.action = 'action.asp?ctrl=3&ID=' + <%=Request("ID")%>;
<% End If %>
						document.frmMain.submit();
					} else {
						document.frmMain.saveme.value = 'Save';
					}
				} else {
					var ans = window.confirm("Save Request? Click Cancel to stop.");
					if (ans) {
						document.frmMain.saveme.value = 'Please Wait...';
<% If Request("ID") = "" Then%>
						document.frmMain.action = 'action.asp?ctrl=1';
<% Else %>
						document.frmMain.action = 'action.asp?ctrl=3&ID=' + <%=Request("ID")%>;
<% End If %>
						document.frmMain.submit();
					} else {
						document.frmMain.saveme.value = 'Save';
					}
				}
			} else {
				var ans = window.confirm("Save Request? Click Cancel to stop.");
				if (ans) {
					document.frmMain.saveme.value = 'Please Wait...';
<% If Request("ID") = "" Then%>
					document.frmMain.action = 'action.asp?ctrl=1';
<% Else %>
					document.frmMain.action = 'action.asp?ctrl=3&ID=' + <%=Request("ID")%>;
<% End If %>
					document.frmMain.submit();
				} else {
					document.frmMain.saveme.value = 'Save';
				}
			}
		}
	}

	function bawal(tmpform) {
		var iChars = ",|\"\'";
		var tmp = "";
		for (var i = 0; i < tmpform.value.length; i++) {
			if (iChars.indexOf(tmpform.value.charAt(i)) != -1) {
				alert ("This character is not allowed.");
			  	tmpform.value = tmp;
			  	return;
		  	} else {
		  		tmp = tmp + tmpform.value.charAt(i);
		  	}
		}
	}

	function bawal2(tmpform) {
		var iChars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ abcdefghijklmnopqrstuvwxyz0123456789-,.\'"; //",|\"\'";
		var tmp = "";
		for (var i = 0; i < tmpform.value.length; i++) {
			if (iChars.indexOf(tmpform.value.charAt(i)) != -1) {
			  	tmp = tmp + tmpform.value.charAt(i);
		  	} else {
		  		alert ("This character is not allowed.");
			  	tmpform.value = tmp;
			  	return;	
		  	}
		}
	}

	function bawalletters(tmpform) {
		var iChars = "0123456789";
		var tmp = "";
		for (var i = 0; i < tmpform.value.length; i++) {
			if (iChars.indexOf(tmpform.value.charAt(i)) != -1) {
				tmp = tmp + tmpform.value.charAt(i);
			} else {
				alert ("This character is not allowed.");
			  	tmpform.value = tmp;
			  	return;	
		  	}
		}
	}

	function DeptInfo(dept) {
		if (dept == 0 ) {
			document.frmMain.selDept.value =0;
			document.frmMain.txtInstAddr.value = "";
			document.frmMain.txtBlname.value = "";
			document.frmMain.txtBilAddr.value = "";
			document.frmMain.h_drg.value = "False";
		}
		<%=strDept2%>
	}

<% If Request("ID") = "" Then %>
	function chkLang() {
		if (document.frmMain.radioLang[0].checked == false) {
			document.frmMain.selLang.value = 0;
			document.frmMain.selLang.disabled = true;
		} else {
			document.frmMain.selLang.disabled = false;
		}
		//if (document.frmMain.radioLang[5].checked == false) {
		//	document.frmMain.txtoLang.value = "";
		//	document.frmMain.txtoLang.disabled = true;
		// } else {
		//  document.frmMain.txtoLang.value = "";
		//	document.frmMain.txtoLang.disabled = false;
		//}
	}
<% End If %>
	function IsMinor() {
		if (document.frmMain.chkminor.checked == false) {
			document.frmMain.txtParents.disabled = true;
			document.frmMain.txtParents.value = "";
		} else {
			document.frmMain.txtParents.disabled = false;
		}
	}
	
	function ReasList(dept) {
		if (document.frmMain.selReas != null) {
			document.frmMain.selReas.length = 0;
			<%=strDept3%>
		}
	}
	
	function MyReas() {
	}

<% If Session("UID") <> 36 Then %> 
	function oLang(xxx) {
		//	if (xxx == 98) {
		//		document.frmMain.txtoLang.disabled = false;
		//	} else {
		//		document.frmMain.txtoLang.value = "";
		//		document.frmMain.txtoLang.disabled = true;
		//	}
	}
<% End If %>
	function uploadFile() {
		var tmpfname = "<%=tmpFilename%>";
		newwindow = window.open('upload.asp?hfname=' + tmpfname ,'name','height=150,width=400,scrollbars=1,directories=0,status=1,toolbar=0,resizable=0');
		if (window.focus) {newwindow.focus()}
	}
<% If Session("myClass") <> 3 Then %>
	function DpwedeMed() {
		if (document.frmMain.chkacc.checked == true || document.frmMain.chkcomp.checked == true) {
			alert("This appointment is not eligible for Medicaid/MCO.");
			document.frmMain.chkout.checked = false;
			return;
		}
	}

	function DpwedeIba() {
		if (document.frmMain.chkout.checked == true) {
			if (document.frmMain.chkacc.checked == true || document.frmMain.chkcomp.checked == true) {
				alert("This appointment is not eligible for Auto Accident and/or Worker's compensation.");
				document.frmMain.chkacc.checked = false;
				document.frmMain.chkcomp.checked = false;
				return;
			}
		}
	}

	function disableMCOFields(ff) {
		ff.AHMIdNum.disabled 	= true;
		ff.MHPnum.disabled 		= true;
		ff.NHHFnum.disabled 	= true;
		ff.WSHPnum.disabled 	= true;
		ff.MCnum.disabled 		= true;
	}

	function resetMCOTexts(ff) {
		ff.AHMIdNum.value 		= "";
		ff.MHPnum.value 		= "";
		ff.NHHFnum.value 		= "";
		ff.WSHPnum.value 		= "";
		ff.MCnum.value 			= "";
	}

	function resetMCOFields() {
		var ff = document.frmMain;

		ff.rdoMed_Ame.disabled 	= true;
		ff.rdoMed_Med.disabled 	= true;
		ff.rdoMed_Mer.disabled 	= true;
		ff.rdoMed_NHH.disabled 	= true;
		ff.rdoMed_Wel.disabled 	= true;
		
		ff.rdoMed_Ame.checked 	= false;
		ff.rdoMed_Med.checked 	= false;
		ff.rdoMed_Mer.checked 	= false;
		ff.rdoMed_NHH.checked 	= false;
		ff.rdoMed_Wel.checked 	= false;
		
		ff.chkawk.disabled 		= true;
		
		resetMCOTexts(ff);
		disableMCOFields(ff);
	}

	function OutPatient() {
		var ff = document.frmMain;
		if (ff.chkout.checked == true) {
			ff.chkmed.disabled = false;
		} else {
			ff.chkmed.checked = false;
			ff.chkmed.disabled = true;
			resetMCOFields();
		}
	}

	function HasMedicaid(dept) {
		var ff = document.frmMain;
		if (ff.chkmed.checked == true) {
			<%=allowMCO%>
			ff.chkawk.disabled = false;
		} else {
			resetMCOFields();
		}
	}
	
// TODO: change this ~~!
	function SelPlan() {
		var ff = document.frmMain;
		disableMCOFields(ff);
		if (ff.rdoMed_Ame.checked == true) {
			ff.AHMIdNum.disabled 	= false;
			ff.MHPnum.value 		= "";
			ff.NHHFnum.value 		= "";
			ff.WSHPnum.value 		= "";
			ff.MCnum.disabled 		= false;
		}
		if (ff.rdoMed_Mer.checked == true) {
			ff.MHPnum.disabled 		= false;
			ff.AHMIdNum.value 		= "";
			ff.NHHFnum.value 		= "";
			ff.WSHPnum.value 		= "";
			ff.MCnum.disabled 		= false;
		}
		if (ff.rdoMed_NHH.checked == true) {
			ff.NHHFnum.disabled 	= false;
			ff.AHMIdNum.value 		= "";
			ff.MHPnum.value 		= "";
			ff.WSHPnum.value 		= "";
			ff.MCnum.disabled 		= false;
		}
		if (ff.rdoMed_Wel.checked == true) {
			ff.WSHPnum.disabled 	= false;
			ff.AHMIdNum.value 		= "";
			ff.NHHFnum.value 		= "";
			ff.MHPnum.value 		= "";
			ff.MCnum.disabled 		= false;
		}
		if (ff.rdoMed_Med.checked == true) {
			ff.MCnum.disabled 		= false;
			ff.AHMIdNum.value 		= "";
			ff.NHHFnum.value 		= "";
			ff.WSHPnum.value 		= "";
			ff.MHPnum.value 		= "";
		}
	}

// TODO: change this ~~!
	function Chkdrg(tmpdept) {
		if (document.frmMain.h_drg.value == "False") {
			// document.frmMain.chkmed.checked = false;
			document.frmMain.MCnum.value = "";
			//document.frmMain.chkmed.disabled = true;
			document.frmMain.MCnum.disabled = true;
			document.frmMain.chkacc.disabled = true;
			document.frmMain.chkcomp.disabled = true;
			// document.frmMain.selIns.disabled = true;
			document.frmMain.chkacc.checked = false;
			document.frmMain.chkcomp.checked = false;
			document.frmMain.btnSec.disabled = true;
			document.frmMain.selIns.value = "";
			document.frmMain.chkout.checked = false;
			document.frmMain.chkout.disabled = true;
		} else {
			document.frmMain.chkout.disabled = false;
			document.frmMain.chkmed.disabled = false;
			// document.frmMain.MCnum.disabled = false;
			document.frmMain.chkacc.disabled = false;
			document.frmMain.chkcomp.disabled = false;
			// document.frmMain.selIns.disabled = false;
			document.frmMain.btnSec.disabled = false;
			OutPatient();
			HasMedicaid(tmpdept);
		}
	}

	function SaveMed(xxx) {
		var ff = document.frmMain;
		if (ff.h_drg.value == "True") {
			if (ff.chkmed.checked == true) {
				if (ff.txtDOB.value == "") {
					alert("Please input client's date of birth.")
					return;
				}
				if (ff.rdoMed_Ame.checked == false &&
						ff.rdoMed_Med.checked == false &&
						ff.rdoMed_Mer.checked == false &&
						ff.rdoMed_NHH.checked == false &&
						ff.rdoMed_Wel.checked == false
						) {
					alert("Please select a Medicaid/MCO.")
					return;
				}
				if (	Trim(ff.AHMIdNum.value) == "" &&
						ff.rdoMed_Ame.checked == true
						) {
					alert("Please input client's AmeriHealth Member ID number.")
					return;
				} else {
					if (Trim(ff.AHMIdNum.value) != "") {
						var chrmed = Trim(ff.AHMIdNum.value);
						if (chrmed.length != 9) {
							alert("Invalid AmeriHealth Member ID number length.")
							return;
						}
					}
				}
				if (	Trim(ff.NHHFnum.value) == "" &&
						ff.rdoMed_NHH.checked == true
						) {
					alert("Please input client's NH Healthy Families number.")
					return;
				} else {
					if (Trim(ff.NHHFnum.value) != "") {
						var chrmed = Trim(ff.NHHFnum.value);
						if (chrmed.length != 11) {
							alert("Invalid NH Healthy Families number length.")
							return;
						}
					}
				}
				if (	Trim(ff.WSHPnum.value) == "" &&
						ff.rdoMed_Wel.checked == true
						) {
					alert("Please input client's Well Sense Health Plan number.")
					return;
				}
				if (	Trim(ff.MCnum.value) == "" &&
						ff.rdoMed_Med.checked == true
						) {
					alert("Please input client's Medicaid number.")
					return;
				} else {
					if (Trim(ff.MCnum.value) != "") {
						var chrmed = Trim(ff.MCnum.value);
						if (chrmed.length != 11) {
							alert("Invalid Medicaid number length.")
							return;
						}
					}
				}
				if (ff.chkawk.checked == false) {
					alert("Acknowledge statement is required.")
					return;
				}
			} 
			ff.action = 'action.asp?ctrl=17&ID=' + xxx;
			ff.submit();
		}
	}

	function openSecIns() {
		newwindow = window.open('secins.asp','name','height=800,width=400,scrollbars=1,directories=0,status=1,toolbar=0,resizable=0');
		if (window.focus) {newwindow.focus()}
	}

	function chkleavemsg() {
		if (document.frmMain.chkcall.checked == false) {
			document.frmMain.chkleave.checked = false;
		}
	}
<% end If %>
		//-->
		</script>
		</head>
		<body onload='DeptInfo(<%=Z_CZero(tmpdept)%>); IsMinor(); ReasList(document.frmMain.selDept.value);
			<% If Request("ID") = "" Then %>
				chkLang();
			<% End If %>
			<% If Session("UID") <> 36 Then %>
				oLang(<%=tmpLang %>);
			<% End If %>
			<% If Session("myClass") <> 3 Then %>
				OutPatient(); HasMedicaid(<%=Z_CZero(tmpdept)%>); Chkdrg(<%=Z_CZero(tmpdept)%>); SelPlan();
			<% End If %>
			'>	
			<form method='post' name='frmMain' action='main.asp'>
				<table cellSpacing='0' cellPadding='0' height='100%' width="100%" border='0' class='bgstyle2'>
					<tr>
						<td height='100px'>
							<!-- #include file="_header.asp" -->
						</td>
					</tr>
					<tr>
						<td valign='top' >
							<form name='frmService' method='post' action=''>
								<table cellSpacing='0' cellPadding='0' width="100%" border='0'>
									<!-- #include file="_greetme.asp" -->
									<tr>
										<td class='title' colspan='10' align='center'><nobr> Interpreter Request Form <%=EditPage%></td>
									</tr>
									<tr>
										<td align='center' colspan='10'><nobr>(*) required</td>
									</tr>
									<tr>
										<td>&nbsp;</td>
										<td  align='left'>
											<div name="dErr" style="width:50%; height:55px;OVERFLOW: auto;">
												<table border='0' cellspacing='1'>		
													<tr>
														<td><span class='error'><%=Session("MSG")%></span></td>
													</tr>
												</table>
											</div>
										</td>
									</tr>
									<tr><td>&nbsp;</td></tr>
									<tr>
										<% If Session("type") = 5 Then %>
											<td align='right'>Court:</td>
										<% Else %>
											<td align='right'>Department:</td>
										<% End If %>
										<td>
											<select name='selDept' class='seltxt' onchange='DeptInfo(this.value); ReasList(this.value); Chkdrg(this.value);'>
												<option value='0'>&nbsp;</option>
												<%=strDept%>
											</select>
											<input type="hidden" name="h_drg">
										</td>
									</tr>
									<% If Session("type") = 3 Or Session("type") = 4 Or Session("type") = 5 Then %>
										<tr>
											<td align='right'>*Requester's Name:</td>
											<td>
												<input class='main' size='50' maxlength='50' name='txtreqname'  value='<%=tmpreqname%>'>
												<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">*full name</span>
											</td>
										</tr>
										<tr>
											<td align='right'>Phone:</td>
											<td><input class='main' size='12' maxlength='12' name='txtRFon' value='<%=tmpRFon%>'></td>
										</tr>
										<% If Session("type") = 5 Then %>
											<tr>
												<td align='right'>Email:</td>
												<td>
													<input class='main' size='25' maxlength='50' name='txtemail' value='<%=tmpPDemail%>'>
												</td>
											</tr>
										<% End If
										If Session("type") = 4 Then %>
											<tr><td align='right' valign="top">CC e-Mail:</td>
												<td><input autocomplete="off" class="main" size="50" maxlength="50" id="txtccaddr" name="txtccaddr" value="" />
													<br /><p style="margin-top: 0px; padding-top: 0px;">Specifying an email address or fax number in this field sends a copy of the confirmation to that address</p>
												</td>
											</tr>
									<%	End If
									End If %>
									<tr><td>&nbsp;</td></tr>
									<tr><td colspan='10'><hr align='center' width='75%'></td></tr>
									<tr><td>&nbsp;</td></tr>
									<tr>
										<td class='header' colspan='10'><nobr>Contact Information</td>
									</tr>
									<tr>
										<td align='right'>Institution:</td>
										<td>
											<span class='HighLight'><%=tmpInst%></span>
											<input type='hidden' name='LBID' value='<%=tmpInstLB%>'>
										</td>
									</tr>
									<tr>
										<td align='right' ><nobr>Appointment Address:</td>
										<td ><input name='txtInstAddr' class='hp' size='90' readonly></td>
									</tr>
									<tr>
										<td align='right'>Billed To:</td>
										<td> <input name='txtBlname' class='hp' size='90' readonly></span></td>
									</tr>
									<tr>
										<td align='right'>Billing Address:</td>
										<td ><input name='txtBilAddr' class='hp' size='90' readonly></span></td>
									</tr>
									<tr>
										<td align='right'>Requesting Person:</td>
										<td ><span class='HighLight'><%=tmpUname%></span></td>
									</tr>
									<tr>
										<td align='right'>E-mail:</td>
										<td ><span class='HighLight'><%=tmpEmail%></span></td>
									</tr>
									<tr>
										<td align='right'>Phone No.:</td>
										<td ><span class='HighLight'><%=tmpIphone%></span></td>
									</tr>
									<tr>
										<td align='right'>Fax No.:</td>
										<td ><span class='HighLight'><%=tmpfax%></span></td>
									</tr>
									<tr><td>&nbsp;</td></tr>
									<tr><td colspan='10'><hr align='center' width='75%'></td></tr>
									<tr><td>&nbsp;</td></tr>
									<tr>
										<td colspan='10' class='header'><nobr>Appointment Information</td>
									</tr>
									<% If Request("ID") <> "" Then %>
										<tr>
											<td align='right'>Request ID:</td>
											<td><b><%=tmpID%></b></td>
										</tr>
									<% End If %>
									<tr>
										<td align='right'>*Client Last Name:</td>
										<td>
											<input class='main' size='20' maxlength='20' name='txtClilname'  value="<%=tmpClname%>" onkeyup='bawal2(this);'>&nbsp;*First Name:
											<input class='main' size='20' maxlength='20' name='txtClifname'  value="<%=tmpCfname%>" onkeyup='bawal2(this);'>
											<input type='checkbox' name='chkminor' value='1' <%=tmpminor%> onclick='IsMinor();'>
											&nbsp;Minor
										</td>
									</tr>
									<tr>
										<td align='right'>Parent's Name:</td>
										<td><input class='main' size='50' maxlength='50' name='txtParents' value='<%=tmpPar%>'></td>
									</tr>
									<tr>
										<td align='right'>Gender:</td>
										<td>
											<select class='seltxt' name='selGen' style='width: 75px;'>
												<option value='-1' > &nbsp; </option>
												<option value='0' <%=tmpMale%>>Male</option>
												<option value='1' <%=tmpfeMale%>>Female</option>
											</select>
										</td>
									</tr>
									<tr>
										<td align='right'>DOB:</td>
										<td>
											<input class='main' size='11' maxlength='10' name='txtDOB' value='<%=tmpDOB%>' onKeyUp="javascript:return maskMe(this.value,this,'2,5','/');" onBlur="javascript:return maskMe(this.value,this,'2,5','/');">
											<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">mm/dd/yyyy</span>
										</td>
									</tr>
									<% If Session("myClass") = 4 Or Session("myClass") = 6 Then %>
										<tr>
											<td align='right'>Patient MR #:</td>
											<td>
												<input class='main' size='50' maxlength='50' name='mrrec' value='<%=mrrec%>' onkeyup='bawal(this);'><span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">*if no MR available, write N/A or NA</span>
											</td>
										</tr>
									<% End If %>
									<tr>
										<td align='right'>Apartment/Suite Number:</td>
										<td>
											<input class='main' size='50' maxlength='50' name='txtCliAddrI' value='<%=tmpCAdrI%>' onkeyup='bawal(this);'>
										</td>
									</tr>
									<tr>
										<td align='right'>Alternate Appointment Address:</td>
										<td><nobr>
											<input class='main' size='50' maxlength='50' name='txtCliAdd' value='<%=tmpCAddr%>' onkeyup='bawal(this);'>
											<input type='checkbox' name='chkClientAdd' value='1' <%=chkUClientadd%>>CHECK this box and FILL these fields if appointment address is different from department address
											<br>
											<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">*Do not include apartment, floor, suite, etc. numbers</span>
										</td>
									</tr>
									<tr>
										<td align='right'>City:</td>
										<td>
											<input class='main' size='25' maxlength='25' name='txtCliCity' value='<%=tmpCity%>' onkeyup='bawal(this);'>&nbsp;State:
											<input class='main' size='2' maxlength='2' name='txtCliState' value='<%=tmpState%>' onkeyup='bawal(this);'>&nbsp;Zip:
											<input class='main' size='10' maxlength='10' name='txtCliZip' value='<%=tmpZip%>' onkeyup='bawal(this);'>
										</td>
										
									</tr>
									<tr>
										<td align='right'>Client Phone:</td>
										<td><input class='main' size='12' maxlength='12' name='txtCliFon' value='<%=tmpCFon%>'></td>
									</tr>
									<tr>
										<td>&nbsp;</td>
										<td colspan="3" align="left">
											<input type='checkbox' name='chkcall' value='1' <%=tmpCall%>  onclick='chkleavemsg();'>
											Language Bank Interpreter to provide courtesy reminder call (Please note that this is ONLY courtesy reminder call and patient/client may still not show up to his/her appointment).
											<br><br>
										</td>
									</tr>
									<tr>
										<td>&nbsp;</td>
										<td colspan="3" align="left">
											<input type='checkbox' name='chkleave' value='1' <%=chkleave%> onclick='chkleavemsg();'>
											If a patient/client does not answer the phone and his answering machine/voice mail picks up a call or family member answers the phone, can interpreter provide/give full appointment<br>
											info (date, time, location, name of hospital/clinic/department, providers name) on patient/client voice message or give this info to patient/client’s family member?
											<br><br>
										</td>
									</tr>
									<tr>
										<td align='right' valign='top'>Client Mobile:</td>
										<td>
											<textarea name='txtCliMobile' class='main' onkeyup='bawal(this);' ><%=tmpCFon2%></textarea>
										</td>
									</tr>
									<% If Session("InstID") <> 108 Then %>
										<% If Session("type") <> 5 Then %>
											<% If Session("myClass") <> 3 Then %>
												<tr>
													<td align='right'>Clinician Name:</td>
													<td>
														<input class='main' size='50' maxlength='50' name='txtClinName' value='<%=tmpCli%>'>
													</td>
												</tr>
												<!--<tr>
													<td align='right'>Time Paged:</td>
													<td>
														<input class='main' size='5' maxlength='5' name='txtTP' onKeyUp="javascript:return maskMe(this.value,this,'2,6',':');" onBlur="javascript:return maskMe(this.value,this,'2,6',':');" value='<%=tmpPage%>'>
														<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">24-hour format</span>
													</td>
												</tr>//-->
												
											<% Else %>
												<tr>
													<td align='right'>Docket Number:</td>
													<td>
														<input class='main' size='50' maxlength='50' name='txtClinName' value="<%=tmpCli%>">
													</td>
												</tr>
												<tr>
													<td align='right'>Court Room No:</td>
													<td>
														<input class='main' size='12' maxlength='12' name='txtTP' value='<%=tmpPage%>'>
													</td>
												</tr>
											<% End If %>
										<% Else %>
											<tr>
												<td align='right'>*Docket Number:</td>
												<td>
													<input class='main' size='50' maxlength='50' name='txtClinName' value='<%=tmpCli%>'>
												</td>
											</tr>	
											<tr>
												<td align='right'>*Amount requested from court:</td>
												<td>
													$<input class='main' size='8' maxlength='7' name='txtPDamount' value='<%=tmpPDAmount%>'>
												</td>
											</tr>	
											<tr>
												<td align='right'>Form 604A:</td>
												<td>
													<input type="button" name="btnUp" value="UPLOAD" onclick="uploadFile();" class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" <%=disUpload%>>
													<!--<input  class='main' type="file" name="F1" size="20" class='btn'>
													<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">*PDF format only</span>//-->
													<input type="hidden" name="h_tmpfilename" value='<%=tmpFilename%>'>
													<%=tmpfileuploaded%>
												</td>
											</tr>
										<% End If %>
									<% Else %>
									 	<tr>
											<td align='right'>DHHS assigned staff:</td>
											<td>
												<input class='main' size='50' maxlength='50' name='txtClinName' value='<%=tmpCli%>'>
											</td>
										</tr>
									<tr>
										<td>&nbsp;</td>
										<td align='left'>Phone:
										<input class='main' size='12' maxlength='12' name='txtdhhsemail' value='<%=tmpdemail%>'>
										E-mail:
										<input class='main' size='25' maxlength='50' name='txtdhhsFon' value='<%=tmpdFon%>'>
										</td>
									</tr>
									<% End if %>
									<tr>
										<td align='right'>*Language:</td>
										<td>
											<% If Session("UID") <> 36 Then %>
												<% If Request("ID") = "" Then %>
													<input type='radio' name='radioLang' <%=radioLang0%> value='0' onclick='chkLang();'>
												<% End If %>
												<select class='seltxt' name='selLang'  style='width:100px;' onclick='oLang(this.value);'>
													<option value='0'>&nbsp;</option>
													<%=strLang%>
												</select>
												<% If Request("ID") = "" Then %>
													<input type='radio' name='radioLang' <%=radioLang1%> value='1' onclick='chkLang();'>
													Portuguese
													<input type='radio' name='radioLang' <%=radioLang2%> value='2' onclick='chkLang();'>
													Spanish
													<input type='radio' name='radioLang' <%=radioLang3%> value='3' onclick='chkLang();'>
													Somali
													<input type='radio' name='radioLang' <%=radioLang4%> value='4' onclick='chkLang();'>
													American Sign Language
													<!--<input type='radio' name='radioLang' <%=radioLang5%> value='5' onclick='chkLang();'>
													Other
													<input class='main' size='15' maxlength='25' name='txtoLang'  value='<%=tmpoLang%>'>//-->
												<% Else %>
													<input class='main' size='15' maxlength='25' name='txtoLang' readonly value='<%=tmpoLang%>'>
												<% End If %>
										<% Else %>
											<input type='radio' name='radioLang' disabled value='1' onclick='chkLang();'>
													Portuguese
											<input type='radio' name='radioLang' checked value='2' onclick='chkLang();'>
													Spanish
										<% End If %>
										</td>
									</tr>
									<tr>
										<td align='right'>*Appointment Date:</td>
										<td>
											<input class='main' size='10' maxlength='10' name='txtAppDate'  readonly value='<%=tmpAppDate%>'>
											<input type="button" value="..." title='Calendar' name="cal1" style="width: 19px;"
											onclick="showCalendarControl(document.frmMain.txtAppDate);" class='btnLnk' onmouseover="this.className='hovbtnLnk'" onmouseout="this.className='btnLnk'">
											<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">mm/dd/yyyy</span>
										</td>
									</tr>
									<tr>
										<td align='right'>*Appointment Time:</td>
										<td>
											&nbsp;From:<input class='main' size='5' maxlength='5' name='txtAppTFrom' value='<%=tmpFtime%>' onKeyUp="javascript:return maskMe(this.value,this,'2,6',':');" onBlur="javascript:return maskMe(this.value,this,'2,6',':');" >
											&nbsp;To:<input class='main' size='5' maxlength='5' name='txtAppTTo' value='<%=tmpTtime%>' onKeyUp="javascript:return maskMe(this.value,this,'2,6',':');" onBlur="javascript:return maskMe(this.value,this,'2,6',':');">
											<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">24-hour format</span>
										</td>
									</tr>
									<% If Session("type") <> 5 Then %>
										<% If Session("myClass") <> 3 Then %>
											<tr>
												<td align='right' valign='top'>Reason:</td>
												<td colspan='2'>
													<select  name="selReas" class='seltxt' multiple style="height: 150px; width:175px;">
													</select>
												</td>
											</tr>
										<% Else %>
											<tr>
												<td align='right' valign='top'>Charge/s:</td>
												<td colspan='2'>
													<textarea name='txtchrg' class='main' style='width: 350px;' onkeyup='bawal(this);'><%=tmpChrge%></textarea>
												</td>
											</tr>
											<tr>
												<td align='right' valign='top'>Attorney:</td>
												<td colspan='2'>
													<input class='main' size='50' maxlength='50' name='txtAtrny' value='<%=tmpAtrny%>'>
												</td>
											</tr>
										<% End If %>
									<% End If %>
									<tr>
										<td align='right' valign='top'>&nbsp;</td>
										<td colspan='2'>
											<input type='checkbox' name='chkCall2' value='1' <%=tmpCall2%>>
												&nbsp;Block Schedule
										</td>
									</tr>
									
									<tr><td>&nbsp;</td></tr>
									<% If Session("myClass") <> 3 Then %>
										<% If Session("type") <> 5 Then %>
											<tr>
												<td align='right'><b>For Medicaid/MCO billing:</b></td>
												<td><b>(also fill in) 
													<% If tmpintr > 0 Then %>
														<% If billedako = False Then %>
															<a href='#' onclick='SaveMed(<%=Request("ID")%>);' style="text-decoration: none;">[Save Medicaid/MCO Info]</a>
														<% End If %>
													<% End If %>
													</b></td>
											</tr>
											<tr>
												<td align='right'>Auto Accident:</td>
												<td><input type='checkbox' name='chkacc' value='1' <%=chkacc%> onclick="DpwedeIba();"></td>
											</tr>
											<tr>
												<td align='right'>Worker's Compensation:</td>
												<td><input type='checkbox' name='chkcomp' value='1' <%=chkcomp%> onclick="DpwedeIba();"></td>
											</tr>
											<tr>
												<td align='right'>Outpatient:</td>
												<td><input type='checkbox' name='chkout' value='1' <%=chkout%> onclick="DpwedeMed(); OutPatient();"></td>
											</tr>
											<tr>
												<td align='right'>Has Medicaid/MCO:</td>
												<td>
													<input type='checkbox' name='chkmed' value='1' <%=chkmed%> onclick="HasMedicaid(document.frmMain.selDept.value);">
													Medicaid:<input type='text' class='main' maxlength='14' name='MCnum' value="<%=MCNum%>">
												</td>
											</tr>
											<tr>
												<td align='right'></td>
												<td colspan='3'>
													<input type='radio' id="rdoMed_Ame" name='radiomed' <%=radiomed5%> value='5' onclick='SelPlan();'>
													AmeriHealth
													<input type='text' class='main' maxlength='9' minlength="9" placeholder="AmeriHealth Member ID" name='AHMIdNum' id='AHMIdNum' value="<%=AHMIdNum%>"><br />

													<input type='radio' id="rdoMed_Mer" name='radiomed' <%=radiomed1%> value='1' onclick='SelPlan();'>
													Meridian Health Plan
													<input type='text' class='main' maxlength='14' name='MHPnum' value="<%=MHPnum%>"><br />

													<input type='radio'  id="rdoMed_NHH" name='radiomed' <%=radiomed2%> value='2' onclick='SelPlan();'>
													NH Healthy Families
													<input type='text' class='main' maxlength='14' name='NHHFnum' value="<%=NHHFnum%>" onkeyup='bawalletters(this);'><br />

													<input type='radio' id="rdoMed_Wel" name='radiomed' <%=radiomed3%> value='3' onclick='SelPlan();'>
													Well Sense Health Plan
													<input type='text' class='main' maxlength='14' name='WSHPnum' value="<%=WSHPnum%>"><span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">(Well Sense number MUST contain NH (eg: NHXXXXXXX).)</span> <br />

													<input type='radio' id="rdoMed_Med" name='radiomed' <%=radiomed4%> value='4' onclick='SelPlan();'>
													Medicaid
													<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">(Directly Billed to Medicaid/Straight Medicaid/Non-MCO)</span> 
													<br />

													<br />
													<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">*If MCO is disabled, LanguageBank does not a have contract with that MCO to directly enter an appointment through LanguageBank database.<br />
													Call LanguageBank for additional info</span>
												
													<br /><br />
<!-- NEW for 2020-03-24 xxxxx -->
<div style="display: inline-block; background-color: #fef; border-radius: 6px; width: 250px; padding: 5px 10px 8px 24px; border: 1px dotted #cbc; margin-bottom: 10px;">
	<input type="checkbox" name="chkteleh" id="chkteleh" value="1" <%=chkTeleh%> style="padding-top: 3px;" />
	This is a TELEHEALTH appointment
</div>
<!-- ^^^^^^^^^^^^^^^^^^^^^^^^ -->
										<br /><br />
													<input type='checkbox' name='chkawk' value='1' <%=chkawk%> >
													Acknowledgement Statement:<br /> 
													- On behalf of my organization/institution, I/we agree to accept financial responsibility for this appointment and agree to pay Language Bank for interpretation services provided to us, if MCO or Medicaid declines to pay/cover this appointment.<br />
													- I acknowledge that appointment entered is NOT Auto Accident or Workers Compensation case. On behalf of my organization/institution, I/we agree to reimburse/pay Language Bank if the state or MCO request repayment (if case is to be Auto Accident or Workers Compensation case), 
													<br /><br />
												</td>
												<!--<td><input type='text' class='main' maxlength='14' name='MCnum' value="<%=MCNum%>"></td>//-->
											</tr>
											
											<tr>
												<td align='right'>Secondary Insurance:</td>
												<td>	
													<!--<select class='seltxt' name='selIns'  style='width:100px;' onchange=''>
														<option value='0'>&nbsp;</option>
														
													</select>//-->
													<input type="text" class="main" readonly name="selIns" size="5" value="<%=tmpsecins%>">
													<input type="button" style="width: 19px;" value="..." title="Choose Secondary Insurance" name="btnSec" class="btnLnk" onmouseover="this.className='hovbtnLnk'" onmouseout="this.className='btnLnk'" onclick="openSecIns();">
													<a href="#" onclick="document.frmMain.selIns.value = '';" style="text-decoration: none;">[Clear]</a>
												</td>
											</tr>
										<tr><td>&nbsp;</td></tr>
										<% End If %>
									<% End If %>
									<tr><td colspan='10'><hr align='center' width='75%'></td></tr>
									<tr><td>&nbsp;</td></tr>
									<tr>
										<td align='right' valign='top'>Special Circumstances/Precautions:</td>
										<td>
											<textarea name='txtCliCir' class='main' onkeyup='bawal(this);' style='width: 350px;'><%=tmpSC%></textarea>
											<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">*Precautions (infections, safety, etc.) for this appointment.</span>
										</td>
									</tr>
									<tr>
										<td align='right' valign='top'>Appointment Comment:</td>
										<td colspan='2'>
											<textarea name='txtcom' class='main' style='width: 350px;' onkeyup='bawal(this);'><%=tmpCom%></textarea>
										</td>
									</tr>
									<tr>
										<td align='right' valign='top'>Languagebank Comment:</td>
										<td colspan='2' valign='top'>
											<textarea name='txtLBcom' class='main' style='width: 350px;' onkeyup='bawal(this);'><%=tmpLBCom%></textarea>
											<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">*Language Bank office staff will only get this info (eg: Preferred / Do Not Assign Interpreter, etc)</span>
										</td>
									</tr>
									<!--<tr>
										<td align='right' valign='top'>Interpreter Comment:</td>
										<td colspan='2' valign='top'>
											<textarea name='txtIntrcom' class='main' style='width: 350px;' onkeyup='bawal(this);'><%=tmpIntrCom%></textarea>
											<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">*Interpreters will only get this info</span>
										</td>
									</tr>//-->
									<tr><td>&nbsp;</td></tr>
									<tr><td colspan='10'><hr align='center' width='75%'></td></tr>
									<tr><td>&nbsp;</td></tr>
									<tr>
										<td colspan='10' align='center' height='100px' valign='bottom'>
											<input class='btn' type='button' value='Save' name='saveme' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='ReqChkMe(<%=myStat%>);' <%=Billedna%>>
											<% If Request("ID") = "" Then%>
												<input class='btn' type='Reset' value='Clear' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'">
											<% Else%>
												<input class='btn' type='button' value='Back' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="document.location='reqconfirm.asp?ID=<%=Request("ID")%>'">
											<% End If%>									
										</td>
									</tr>
								</table>
							</form>
						</td>
					</tr>
					<tr>
						<td valign='bottom'>
							<!-- #include file="_footer.asp" -->
						</td>
					</tr>
				</table>
			</form>
		</body>
</html>
<%
If Session("MSG") <> "" Then
	tmpMSG = Replace(Session("MSG"), "<br>", "\n")
%>
<script><!--
	alert("<%=tmpMSG%>");
--></script>
<%
End If
Session("MSG") = ""
%>
