<%Language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="main_helper.asp" -->
<%Response.AddHeader "Pragma", "No-Cache" %>
<%
If Cint(Session("type")) = 1 And Session("UID") <> 35 Then
	Session("MSG") = "Error: Invalid user type. Please sign-in again."
	Response.Redirect "default.asp"
End If
'Response.Write "Class: " & Session("myclass") & "--<br />"
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
			tmpGender	= Z_CZero(rsApp("Gender"))
			tmpMale = ""
			tmpFemale = ""
			If tmpGender = 0 Then 
				tmpMale = "SELECTED"
			ElseIf tmpGender = 1 Then 
				tmpFemale = "SELECTED"
			End If
		End If
		tmpDOB = rsApp("DOB") 
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
		tmpsecins = rsApp("secins")
		mrrec = rsApp("mrrec")
	End If ' NOT EOF -- found the appointment (not the request!)'
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
	tmpCom = tmpEntry(11)
	tmpDept = tmpEntry(4)
	tmpCall = "checked"
	If tmpEntry(3) = "" Then tmpCall = ""
	radioLang0 = ""
	Select Case tmpEntry(14)
		Case 0
			radioLang0 = "checked"
		Case 1
			radioLang1 = "checked"
		Case 2
			radioLang2 = "checked"
		Case 3
			radioLang3 = "checked"
		Case 4
			radioLang4 = "checked"
		Case 5
			radioLang5 = "checked"
	End Select
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
	Set rsDept = Server.CreateObject("ADODB.RecordSet")
	sqlDept = "SELECT * FROM Dept_T WHERE [index] = " & Session("DeptID") & " ORDER BY dept"
	rsDept.Open sqlDept, g_strCONNLB, 3, 1
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
%>
<!-- #include file="_closeSQL.asp" -->
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html;charset=ISO-8859-1"> 
		<title>Interpreter Request - Interpreter Request Form</title>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<link href="CalendarControl.css" type="text/css" rel="stylesheet">
		<script src="CalendarControl.js" language="javascript"></script>
		<script src="main_helper.js" language="javascript"></script>
		<script language='JavaScript'><!--
function ReqChkMe(xxx) {
	if (document.frmMain.saveme.value == 'Save') {		
		if (xxx == 2 || xxx == 3 || xxx == 4) {
			alert("ERROR: You cannot edit canceled/missed appointments.");
			return;
		}
		if (document.frmMain.selDept.value == 0) {
			alert("ERROR: Please provide a department.");
			return;
		}
		if (document.frmMain.txtreqname.value == "") {
			alert("ERROR: Please provide a Requester Name.");
			return;
		}
		if (document.frmMain.txtClilname.value == "" || document.frmMain.txtClifname.value == "") {
			alert("ERROR: Please provide patient.");
			document.frmMain.txtClilname.focus();
			return;
		}
		if ((document.frmMain.chkcall.checked == true || document.frmMain.chkleave.checked == true) && document.frmMain.txtCliFon.value == "") {
			alert("Please input patient's phone number.");
			document.frmMain.txtCliFon.focus();
			return;
		}
		if (document.frmMain.mrrec.value == "") {
			alert("ERROR: Please provide Patient MR#.");
			return;
		}
		if ((document.frmMain.radioLang.value == 0) && (document.frmMain.selLang.value == 0) ) {
			alert("ERROR: Please provide language.");
			document.frmMain.selLang.focus();
			return;
		}
		if 	(document.frmMain.txtAppDate.value == "") {
			alert("ERROR: Specify appointment date.");
			return;
		}
		if 	(Trim(document.frmMain.txtAppTFrom.value) == "" || Trim(document.frmMain.txtAppTTo.value) == "") {
			alert("ERROR: Please provide appointment time.");
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
			alert("ERROR: Invalid time frame.");
			return;
		}
				
		if (document.frmMain.h_drg.value == "True") {
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
		
function DeptInfo(dept) {
	if (dept == 0 )	{
		document.frmMain.selDept.value =0;
		document.frmMain.txtInstAddr.value = "";
		document.frmMain.txtBlname.value = "";
		document.frmMain.txtBilAddr.value = "";
		document.frmMain.h_drg.value = "False";
	}
	<%=strDept2%>
}

function IsMinor() {
	if (document.frmMain.chkminor.checked == false)	{
		document.frmMain.txtParents.disabled = true;
		document.frmMain.txtParents.value = "";
	} else {
		document.frmMain.txtParents.disabled = false;
	}
}
<% If Request("ID") = "" Then %>
function chkLang() {
	if (document.frmMain.radioLang[0].checked == false) {
		document.frmMain.selLang.value = 0;
		document.frmMain.selLang.disabled = true;
	} else {
		document.frmMain.selLang.disabled = false;
	}
}
<% End If %>
// -->
		</script>
		</head>
		<body onload='DeptInfo(<%=Z_CZero(tmpdept)%>); IsMinor();
			<% If Request("ID") = "" Then %>
				chkLang();
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
										<td align='right'>Department:</td>
										<td>
											<select name='selDept' class='seltxt' onchange='DeptInfo(this.value); Chkdrg(this.value);'>
												<option value='0'>&nbsp;</option>
												<%=strDept%>
											</select>
											<input type="hidden" name="h_drg">
										</td>
									</tr>
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
									<tr>
										<td align='right'>Email:</td>
										<td>
											<input class='main' size='25' maxlength='50' name='txtemail' value='<%=tmpPDemail%>'>
										</td>
									</tr>

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
									<tr><td colspan='10' class='header'><nobr>Language</td></tr>
									<tr>
										<td align='right'>*Language:</td>
										<td>
											<% If Request("ID") = "" Then %>
												<input type='radio' name='radioLang' <%=radioLang0%> value='0' onclick='chkLang();'>
											<% End If %>
											<select class='seltxt' name='selLang'  style='width:100px;' >
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
											<% Else %>
												<input class='main' size='15' maxlength='25' name='txtoLang' readonly value='<%=tmpoLang%>'>
											<% End If %>
										</td>
									</tr>
									<tr><td>&nbsp;</td></tr>
									<tr><td colspan='10' class='header'><nobr>Patient Information</td></tr>
									<tr>
										<td align='right'>* Patient Last Name:</td>
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
									<tr>
										<td align='right'>Patient MR #:</td>
										<td>
											<input class='main' size='50' maxlength='50' name='mrrec' value='<%=mrrec%>' onkeyup='bawal(this);'><span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">*if no MR available, write N/A or NA</span>
										</td>
									</tr>
									<tr>
										<td align='right'>Patient Phone:</td>
										<td><input class='main' size='12' maxlength='12' name='txtCliFon' value='<%=tmpCFon%>'></td>
									</tr>
									<tr>
										<td align='right' valign='top'>Patient Mobile:</td>
										<td>
											<textarea name='txtCliMobile' class='main' onkeyup='bawal(this);' ><%=tmpCFon2%></textarea>
										</td>
									</tr>
<tr><td colspan='10'><hr align='center' width='75%'></td></tr>
									<tr><td>&nbsp;</td></tr>
									<tr><td colspan='10' class='header'><nobr>Appointment Information</td></tr>
									<% If Request("ID") <> "" Then %>
										<tr>
											<td align='right'>Request ID:</td>
											<td><b><%=tmpID%></b></td>
										</tr>
									<% End If %>
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
										<td align='right' valign='top'>&nbsp;</td>
										<td colspan='2'>
											<input type='checkbox' name='chkCall2' value='1' <%=tmpCall2%>>
												&nbsp;Block Schedule
										</td>
									</tr>
									
									<tr><td>&nbsp;</td></tr>

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