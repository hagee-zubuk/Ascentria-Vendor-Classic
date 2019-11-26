<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilCalendar.asp" -->
<%
If Session("type") = "" Then
	Session("MSG") = "Sesion has expired. Please sign in again."
	Response.Redirect "default.asp"
End If
Function GetLang(xxx)
	GetLang = ""
	Set rsLang = Server.CreateObject("ADODB.RecordSet")
	sqlLang = "SELECT * FROM Lang_T WHERE [index] = " & xxx
	rsLang.Open sqlLang, g_strCONN, 3, 1
	If Not rsLang.EOF Then
		GetLang = rsLang("Lang")
	End If
	rsLang.Close
	Set rsLang = Nothing
End Function
Function Z_FormatTime(xxx)
	Z_FormatTime = Null
	If xxx <> "" Or Not IsNull(xxx)  Then
		If IsDate(xxx) Then Z_FormatTime = FormatDateTime(xxx, 4) 
	End If
End Function
Function CleanFax(strFax)
	CleanFax = Replace(strFax, "-", "") 
End Function
Function GetLoc(xxx)
	Select Case xxx
		Case 0 
			GetLoc = "Front Door"
		Case 1
			GetLoc = "Cafeteria"
		Case 2
			GetLoc = "Registration Desk"
		Case 3
			GetLoc = "Department"
		Case 4
			GetLoc = "OTHER"
	End Select
End Function
Function GetMyStatus(xxx)
	Select Case xxx
		Case 1
			GetMyStatus = "COMPLETED"
		Case 2
			GetMyStatus = "MISSED"
		Case 3
			GetMyStatus = "CANCELED"
		Case 4
			GetMyStatus = "CANCELED (BILLABLE)"
		Case Else
			GetMyStatus = "PENDING"
	End Select
End Function
'GET INST
'Set rsInst = Server.CreateObject("ADODB.RecordSet")
'sqlInst = "SELECT * FROM appointment_T, Inst_T WHERE appointment_T.index = " & Request("ID") & _
'	" AND InstID = Inst_T.index"
'rsInst.Open sqlInst, g_strCONN, 3, 1
'If Not rsInst.EOF Then
'	tmpInstID = rsInst("Inst_T.index")
'	tmpInst = rsInst("inst")
'	tmpAddr = rsInst("Addr") & ", "& rsInst("City") & ", " & UCase(rsInst("State")) & ", " & rsInst("zip")
'	tmpBill = rsInst("BillTo")
'	tmpBAddr = rsInst("BillAddr") & ", "& rsInst("BillCity") & ", " & UCase(rsInst("BillState")) & ", " & rsInst("Billzip")
'	tmpIphone = rsInst("Inst_T.phone")
'	tmpEmail = rsInst("email")
'End If
'rsInst.Close
'Set rsInst = Nothing
'GET INST and DEPTV2
lngReqID = Z_CLng(Request("ID"))
Set rsInst = Server.CreateObject("ADODB.RecordSet")
sqlInst = "SELECT * FROM Appointment_T WHERE [index] = " & Request("ID")
rsInst.Open sqlInst, g_strCONN, 3, 1
If Not rsInst.EOF Then
	tmpInst = GetFacility(rsInst("InstID"))
	tmpDept = GetDept(rsInst("DeptID"))
	tmpDeptID = rsInst("DeptID")
End If
rsInst.Close
Set rsInst = Nothing
'GET USER
Set rsUser = Server.CreateObject("ADODB.RecordSet")
sqlUSer = "SELECT * FROM User_T WHERE [index] = " & Session("UID")
rsUser.Open sqlUser, g_strCONN, 3, 1
If Not rsUser.EOF Then
	tmpUname = rsUser("lname") & ", " & rsUser("fname")
End If
rsUser.Close
Set rsUSer = Nothing
'GET DEPT V2
Set rsDept = Server.CreateObject("ADODB.RecordSet")
sqldept = "SELECT * FROM dept_T WHERE [index] = " & tmpDeptID
rsDept.Open sqldept, g_strCONNLB, 3, 1
If Not rsDept.EOF Then
	tmpAddr = rsDept("InstAdrI") & ", " & rsDept("address") & ", " & rsDept("city") & ", " & rsDept("state") & ", " & rsDept("zip")
	tmpBill = rsDept("Blname")
	tmpBAddr = rsDept("BAddress") & ", "& rsDept("BCity") & ", " & UCase(rsDept("BState")) & ", " & rsDept("Bzip")
	tmpClass = rsDept("class")
End If
rsDept.Close
Set rsDept = Nothing
'GET APPOINTMENT
Set rsApp = Server.CreateObject("ADODB.RecordSet")
sqlApp = "SELECT * FROM Appointment_T WHERE [index] = " & Request("ID")
rsApp.Open sqlApp, g_strCONN, 3, 1
If Not rsApp.EOF Then
	tmpTS = rsApp("timestamp")
	tmpCName =  Z_DoDecrypt(rsApp("clname")) & ", " & Z_DoDecrypt(rsApp("cfname"))
	'tmpCAddr =  rsApp("Addr") & ", "& rsApp("City") & ", " & UCase(rsApp("State")) & ", " & rsApp("zip")
	tmpFon =  Z_DoDecrypt(rsApp("Phone"))
	tmpMobile =  Z_DoDecrypt(rsApp("Mobile"))
	tmpLang = GetLang(rsApp("LangID")) 
	if rsApp("oLang") <> "" Then tmpLang = tmpLang & " (" & rsApp("oLang") & ")"
	tmpAppDate = rsApp("AppDate")
	tmpAppTime = Z_FormatTime(rsApp("TimeFrom")) & " - " &  Z_FormatTime(rsApp("TimeTo"))
	tmpIntrID = rsApp("IntrID")
	tmpCom = rsApp("comment")
	tmpReas = GetReas(Z_Replace(rsApp("reason"),", ", "|"))
	tmpCall = ""
	If Session("InstID") = 108 Then
		If rsApp("callme") = True Then tmpCall = "* Call client to remind of appointment"
	Else
		If Session("myClass") <> 3 Then
			If rsApp("callme") = True Then tmpCall = "* Call patient to remind of appointment"
		Else
			If rsApp("callme") = True Then tmpCall = "* Call client to remind of appointment"
		End If
	End If
	tmpblock = ""
	if rsapp("block") then tmpblock = "(BLOCK SHEDULE)"
	tmpstat = GetMyStatus(GetStatLB(Request("ID")))'GetMyStatus(rsApp("status"))
	tmpstat2 = GetStatLB(Request("ID"))
	LockMe = ""
	If rsApp("status") = 1 Or rsApp("status") = 4 Then LockMe = "disabled"
	tmpMinor = ""
	If rsApp("minor") = True Then tmpMinor = "* Minor"
	tmpParents = ""
	If rsApp("parents") <> "" Then tmpParents = rsApp("parents")
	tmpReqName = rsApp("reqName")
	tmpReqPhone = rsApp("RPhone")
	tmpDOB = rsApp("DOB")
	If Session("myClass") <> 3 Then
		tmpClin = rsApp("clinician")
	Else
		tmpClin = rsApp("docknum")
		tmpCrt = rsApp("crtroom")
		tmpAttny = rsApp("attny")
		tmpChrg = rsApp("charges")
	End If
	tmpmed = ""
	If rsApp("outpatient") And rsApp("hasmed") Then
		tmpmed = rsApp("medicaid")
		tmpmer = rsApp("meridian")
		tmpnh = rsApp("nhhealth")
		tmpwell = rsApp("wellsense")
		tmpame = rsApp("amerihealth")
	End If
	tmpLBCom = rsApp("lbcom")
	tmpIntrCom = rsApp("intrcom")
	If rsApp("Gender") = vbNull Then
		tmpSex = "Unknown"
	Else
		tmpGender	= Z_CZero(rsApp("Gender"))
		If tmpGender = 0 Then 
			tmpSex = "MALE"
		ElseIf tmpGender = 1 Then 
			tmpSex = "FEMALE"
		End If
	End If
	tmpCAddress = rsApp("capt") & ", " & rsApp("caddress") & ", " & rsApp("ccity") & ", " & rsApp("cstate") & ", " & rsApp("czip")
	myInst = rsApp("InstLB")
	tmpPDamount = rsApp("PDAmount")
	tmpSC = Replace(Z_FixNull(rsApp("spec_cir")), vbcrlf, "<br>")
	If rsApp("uploadfile") Then 
		uploadfileviewLB = "*Unapproved Form 604A has already been uploaded."
		If rsApp("approvePDF") Then
			uploadfileviewLB = "*Form 604A already approved."
		End If
	Else
		uploadfileviewLB = "*Form 604A has not been uploaded."
	End If
End If
rsApp.Close
Set rsApp = Nothing
If Session("type") <> 6 And tmpClass <> 3 Then
	If Session("InstID") <> myInst Then
		Session("MSG") = "Error: You are not allowed to view this appointment. Please sign-in again."
		Response.Redirect "default.asp"
	End If
End If
If tmpIntrID <> 0 Then
	'GET INTERPRETER (CHANGE TO LB INTERPRETER)
	'Set rsIntr = Server.CreateObject("ADODB.RecordSet")
	'sqlIntr = "SELECT * FROM Intr_T  WHERE index = " & tmpIntrID
	'rsIntr.Open sqlIntr, g_strCONN, 3, 1
	'If Not rsIntr.EOF Then
	'	tmpIntrName = rsIntr("lname") & ", " & rsIntr("fname")
	'	tmpIntrAddr = rsIntr("Addr") & ", " & rsIntr("City") & ", " & UCase(rsIntr("State")) & ", " & rsIntr("Zip")
	'End If
	'rsIntr.Close
	'Set rsIntr = Nothing
	Set rsIntr = Server.CreateObject("ADODB.RecordSet")
	sqlIntr = "SELECT * FROM interpreter_T  WHERE [index] = " & tmpIntrID
	rsIntr.Open sqlIntr, g_strCONNLB, 3, 1
	If Not rsIntr.EOF Then
		tmpIntrName = rsIntr("last name") & ", " & rsIntr("first name")
		tmpIntrAddr = rsIntr("Address1") & ", " & rsIntr("City") & ", " & UCase(rsIntr("State")) & ", " & rsIntr("Zip Code")
	End If
	rsIntr.Close
	Set rsIntr = Nothing
End If
%>
<!-- #include file="_closeSQL.asp" -->
<html>
	<head>
		<title>Interpreter Request - Request Confirmation</title>
		<link href="CalendarControl.css" type="text/css" rel="stylesheet">
		<script src="CalendarControl.js" language="javascript"></script>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
		<!--
		function CalendarView(strDate)
		{
			document.frmConfirm.action = 'calendarview2.asp?appDate=' + strDate;
			document.frmConfirm.submit();
		}
		function CancelMe(xxx, yyy)
		{
			if (yyy == 3)
			{
				alert("This request has been canceled already.")
				return;
			}
			var ans = window.confirm("This action will cancel your request.\nAn E-mail will be sent to a LB staff for notification.\nClick Cancel to stop.");
			if (ans)
			{
				document.frmConfirm.action = "action.asp?ctrl=8&ID=" + xxx;
				document.frmConfirm.submit();
			}
		}
		function DeleteMe(xxx)
		{
			var ans = window.confirm("This action will delete your request.\nClick Cancel to stop.");
			if (ans)
			{
				document.frmConfirm.action = "action.asp?ctrl=10&ID=" + xxx;
				document.frmConfirm.submit();
			}
		}
		function CloneMe(xxx)
		{
			document.frmConfirm.action = "main.asp?clone=" + xxx;
			document.frmConfirm.submit();
		}
		function PrintMe(xxx)
		{
			newwindow = window.open('print2.asp?ID=' + xxx ,'','height=800,width=900,scrollbars=1,directories=0,status=0,toolbar=0,resizable=1');
			if (window.focus) {newwindow.focus()}
		}
		function mySurvey(xxx) {
			newwindow = window.open('survey.asp?ID=' + xxx ,'','height=800,width=900,scrollbars=1,directories=0,status=0,toolbar=0,resizable=1');
			if (window.focus) {newwindow.focus()}
		}
		function uploadFile() {
<%
If Session("type") = 5 Then 'create temp filename
	tmpFilename = Z_GenerateGUID()
	Do Until GUIDExists(tmpFilename) = False
		tmpFilename = Z_GenerateGUID()
	Loop
Else
	tmpFilename = "UNUSED"
End If
%>
			var tmpfname = "<%=tmpFilename%>";
			newwindow = window.open('upload2.asp?rid=<%=lngReqID%>&hfname=<%=tmpFilename%>','name','height=150,width=400,scrollbars=1,directories=0,status=1,toolbar=0,resizable=0');
				if (window.focus) {newwindow.focus()}
		}
		-->
		</script>
	</head>
	<body>
		<form method='post' name='frmConfirm'>
			<table cellSpacing='0' cellPadding='0' height='100%' width="100%" border='0' class='bgstyle2'>
				<tr>
					<td valign='top'>
						<!-- #include file="_header.asp" -->
					</td>
				</tr>
				<tr>
					<td valign='top' >
						<table cellSpacing='2' cellPadding='0' width="100%" border='0'>
							<!-- #include file="_greetme.asp" -->
							<tr>
								<td class='title' colspan='2' align='center'><nobr>Request Confirmation</td>
							</tr>
							<% If tmpIntrID > 0 Then %>
								<tr>
									<td colspan='2' align='center'>
										<a href="#" onclick="mySurvey(<%=Request("ID")%>);" style="text-decoration: none;">[Interpreter Feedback]</a>
									</td>
								</tr>
							<% End If %>
							<tr>
								<td width='10px'>
								&nbsp;
								</td>
							</tr>
							<tr>
								<td align='center' colspan='2'><span class='error'><%=Session("MSG")%></span></td>
							</tr>
							<tr>
								<td class='header' colspan='2'><nobr>Contact Information </td>
							</tr>
							<tr>
								<td align='right' width='25%'>Institution:</td>
								<td class='confirm'><%=tmpInst%></td>
							</tr>
							<tr>
								<td align='right' width='25%'>Address:</td>
								<td class='confirm'><%=tmpAddr%></td>
							</tr>
							<tr>
								<td align='right' width='25%'>Billed To:</td>
								<td class='confirm'><%=tmpBill%></td>
							</tr>
							<tr>
								<td align='right' width='25%'>Billing Address:</td>
								<td class='confirm'><%=tmpBAddr%></td>
							</tr>
							<% If Session("type") <> 1 Then %>
								<tr>
									<td align='right' width='25%'>Requesting Person:</td>
									<td class='confirm'><%=tmpUname%></td>
								</tr>
								<tr>
							<% End If %>
								<td align='right' width='25%'>Department:</td>
								<td class='confirm'><%=tmpDept%></td>
							</tr>
							<% If Session("type") <> 1 Then %>
								<tr>
									<td align='right' width='25%'>E-mail:</td>
									<td class='confirm'><%=tmpEmail%></td>
								</tr>
								<tr>
									<td align='right' width='25%'>Phone No.:</td>
									<td class='confirm'><%=tmpIphone%></td>
								</tr>
							<% End If %>
							<tr><td>&nbsp;</td></tr>
							<tr><td colspan='2'><hr align='center' width='75%'></td></tr>
							<tr><td>&nbsp;</td></tr>
							<tr>
								<td colspan='2' class='header'><nobr>Appointment Information</td>
							</tr>
							<tr>
								<td align='right' width='25%'>Request ID:</td>
								<td class='confirm' width='75%' ><%=Request("ID")%>&nbsp;&nbsp; <%=tmpblock%></td>
							</tr>
							<tr>
								<td align='right' width='25%'>Timestamp:</td>
								<td class='confirm' width='75%' ><%=tmpTS%></td>
							</tr>
							<tr>
								<td align='right' width='25%'>Status:</td>
								<td class='confirm' ><%=tmpstat%></td>
							</tr>
							<% If Session("type") = 3 Then %>
							<tr>
								<td align='right' width='25%'>Requester:</td>
								<td class='confirm' ><%=tmpReqName%></td>
							</tr>
							<tr>
								<td align='right' width='25%'>Requester's Phone:</td>
								<td class='confirm' ><%=tmpReqPhone%></td>
							</tr>
							<% End If %>
							<tr><td>&nbsp;</td></tr>
							<tr>
								<td align='right' width='25%'>Client Name:</td>
								<td class='confirm'><%=tmpCName%></td>
							</tr>
							<% If tmpCall <> "" Then %>
								<tr>
									<td>&nbsp;</td>
									<td class='confirm'><%=tmpCall%></td>
								</tr>
							<% End If %>
							<% If tmpMinor <> "" Then %>
								<tr>
									<td>&nbsp;</td>
									<td class='confirm'><%=tmpMinor%></td>
								</tr>
							<% End If %>
							<% If tmpParents <> "" Then %>
								<tr>
									<td align='right' width='25%'>Parent's Name:</td>
									<td class='confirm'><%=tmpParents%></td>
								</tr>
							<% End If %>
							<tr>
								<td align='right'>Gender:</td>
								<td class='confirm'><%=tmpSex%></td>
							</tr>
							<tr>
								<td align='right'>DOB:</td>
								<td class='confirm'><%=tmpDOB%></td>
							</tr>
							<tr>
								<td align='right'>Client Address:</td>
								<td class='confirm'><%=tmpCAddress%></td>
							</tr>
							<tr>
								<td align='right' width='25%'>Phone:</td>
								<td class='confirm'><%=tmpFon%></td>
							</tr>
							<tr>
								<td align='right' width='25%'>Mobile:</td>
								<td class='confirm'><%=tmpMobile%></td>
							</tr>
							<tr>
								<td align='right' width='25%'>Language:</td>
								<td class='confirm'><%=tmpLang%></td>
							</tr>
							<tr>
								<td align='right' width='25%'>Appointment Date:</td>
								<td class='confirm'><%=tmpAppDate%></td>
							</tr>
							<tr>
								<td align='right' width='25%'>Appointment Time:</td>
								<td class='confirm'><%=tmpAppTime%></td>
							</tr>
							<tr>
								<td align='right' width='25%'>									
								<% If Session("InstID") <> 108 Then %>
									<% If Session("type") <> 5 Then %>
										<% If Session("myClass") <> 3 Then %>
											Clinician:
										<% Else %>
											Docket Number:
										<% End If %>
									<% Else %>
											Docket Number:
									<% End If %>
								<% Else %>
									DHHS assigned staff:
								<% End If %>
								<td class='confirm'><%=tmpClin%></td>
							</tr>
							<% If Session("type") <> 5 Then %>
								<% If Session("myClass") <> 3 Then %>
									<tr>
										<td align='right' width='25%' valign='top'>Reason:</td>
										<td class='confirm'><%=tmpReas%></td>
									</tr>
									<tr><td>&nbsp;</td></tr>
									<tr>
										<td align='right'><b>For Medicaid/MCO Billing:</b></td>
										<td align='left'><b>----</b></td>
									</tr>
									<tr>
										<td align='right'>AmeriHealth Member ID Number:</td>
										<td class='confirm'><%=tmpAme%></td>
									</tr>
									<tr>
										<td align='right'>Medicaid Number:</td>
										<td class='confirm'><%=tmpMed%></td>
									</tr>
									<tr>
										<td align='right'>Meridian Number:</td>
										<td class='confirm'><%=tmpMer%></td>
									</tr>
									<tr>
										<td align='right'>NH Health Number:</td>
										<td class='confirm'><%=tmpnh%></td>
									</tr>
									<tr>
										<td align='right'>Well Sense Number:</td>
										<td class='confirm'><%=tmpwell%></td>
									</tr>
								<% Else %>
									<tr>
										<td align='right' width='25%' valign='top'>Court Room No:</td>
										<td class='confirm'><%=tmpCrt%></td>
									</tr>
									<tr>
										<td align='right' width='25%' valign='top'>Attorney:</td>
										<td class='confirm'><%=tmpAttny%></td>
									</tr>
									<tr>
										<td align='right' width='25%' valign='top'>Charge/s:</td>
										<td class='confirm'><%=tmpChrg%></td>
									</tr>
								<% End If %>
							<% Else %>
									<tr>
										<td align='right' width='25%' valign='top'>Amount requested from court:</td>
										<td class='confirm'>$<%=Z_FormatNumber(tmpPDamount, 2)%></td>
									</tr>
									<tr>
										<td align='right'>&nbsp;</td>
										<td class='confirm'><%=uploadfileviewLB%>
											<br />
											<input type="button" name="btnUp" value="UPLOAD" onclick="uploadFile();" class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" <%=disUpload%>>
										</td>
									</tr>
							<% End If %>
							<tr><td>&nbsp;</td></tr>
							<tr><td colspan='2'><hr align='center' width='75%'></td></tr>
							<tr><td>&nbsp;</td></tr>
							<tr>
								<td align='right' width='25%' valign='top'>Comment:</td>
								<td class='confirm'><%=tmpCom%></td>
							</tr>
							<tr><td>&nbsp;</td></tr>
							<tr>
								<td align='right' width='25%' valign='top'>LanguageBank Comment:</td>
								<td class='confirm'><%=tmplbCom%></td>
							</tr>
							<tr><td>&nbsp;</td></tr>
							<tr>
								<td align='right' width='25%' valign='top'>Special Circumstances/Precautions:</td>
								<td class='confirm'><%=tmpsc%></td>
							</tr>
							<tr><td>&nbsp;</td></tr>
							<tr><td colspan='2'><hr align='center' width='75%'></td></tr>
							<tr><td>&nbsp;</td></tr>
							<tr>
							<tr>
								<td colspan='2' class='header'><nobr>Interpreter Information</td>
							</tr>
							<tr>
								<td align='right' width='25%'>Interpreter:</td>
								<td class='confirm'><%=tmpIntrName%></td>
							</tr>
							<tr><td>&nbsp;</td></tr>
							<tr><td colspan='2'><hr align='center' width='75%'></td></tr>
							<tr><td>&nbsp;</td></tr>
							<tr>
								<td colspan='2' align='center' height='100px' valign='bottom'>
									<input class='btn' type='button' style='width: 125px;' value='View in Calendar' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="document.location='calendarview2.asp?appdate=<%=tmpAppDate%>'">
									<% If Session("type") = 0 Or Session("type") = 3 Or Session("type") = 4 Or Session("type") = 5 Then %>
										<input class='btn' type='button' style='width: 125px;' value='Edit' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="document.location='main.asp?ID=<%=Request("ID")%>';" <%=LockMe%>>
										<input class='btn' type='button' style='width: 125px;' value='Cancel Appt.' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='CancelMe(<%=Request("ID")%>, <%=tmpstat2%>);' <%=LockMe%>>
										<% If Session("InstID") = 108 Or Session("myClass") = 3 Then %>	
											<input class='btn' type='button' style='width: 125px;' value='Print' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='PrintMe(<%=Request("ID")%>);'>
										<% End If %>
										<input class='btn' type='button' style='width: 125px;' value='Clone Appt.' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='CloneMe(<%=Request("ID")%>);'>
									<% End If %>
									<% If Session("type") = 2 Then %>
										<input class='btn' type='button' style='width: 125px;' value='Delete' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="DeleteMe(<%=Request("ID")%>);">
									<% End If %>
									<% If Session("type") = 6 Then %>	
										<input class='btn' type='button' style='width: 125px;' value='Print' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='PrintMe(<%=Request("ID")%>);'>
									<% End If %>
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
	tmpMSG = Session("MSG")
%>
<script><!--
	alert("<%=tmpMSG%>");
--></script>
<%
End If
Session("MSG") = ""
%>