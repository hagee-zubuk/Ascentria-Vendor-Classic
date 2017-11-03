<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<%
Function GetLang(xxx)
	GetLang = ""
	Set rsLang = Server.CreateObject("ADODB.RecordSet")
	sqlLang = "SELECT * FROM Lang_T WHERE index = " & xxx
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
'GET INST
Set rsInst = Server.CreateObject("ADODB.RecordSet")
sqlInst = "SELECT * FROM appointment_T, Inst_T WHERE appointment_T.index = " & Request("ID") & _
	" AND InstID = Inst_T.index"
rsInst.Open sqlInst, g_strCONN, 3, 1
If Not rsInst.EOF Then
	tmpInstID = rsInst("Inst_T.index")
	tmpInst = rsInst("inst")
	tmpAddr = rsInst("Addr") & ", "& rsInst("City") & ", " & UCase(rsInst("State")) & ", " & rsInst("zip")
	tmpBill = rsInst("BillTo")
	tmpBAddr = rsInst("BillAddr") & ", "& rsInst("BillCity") & ", " & UCase(rsInst("BillState")) & ", " & rsInst("Billzip")
	tmpIphone = rsInst("Inst_T.phone")
	tmpEmail = rsInst("email")
End If
rsInst.Close
Set rsInst = Nothing
'GET USER
Set rsUser = Server.CreateObject("ADODB.RecordSet")
sqlUSer = "SELECT * FROM User_T WHERE index = " & Session("UID")
rsUser.Open sqlUser, g_strCONN, 3, 1
If Not rsUser.EOF Then
	tmpUname = rsUser("lname") & ", " & rsUser("fname")
End If
rsUser.Close
Set rsUSer = Nothing
'GET DEPT
Set rsIUser = Server.CreateObject("ADODB.RecordSet")
sqlIUSer = "SELECT * FROM InstUser_T WHERE UID = " & Session("UID")
rsIUser.Open sqlIUser, g_strCONN, 3, 1
If Not rsIUser.EOF Then
	tmpDept = rsIUser("dept")
End If
rsIUser.Close
Set rsIUSer = Nothing
'GET APPOINTMENT
Set rsApp = Server.CreateObject("ADODB.RecordSet")
sqlApp = "SELECT * FROM Appointment_T WHERE index = " & Request("ID")
rsApp.Open sqlApp, g_strCONN, 3, 1
If Not rsApp.EOF Then
	tmpCName = rsApp("clname") & ", " & rsApp("cfname")
	'tmpCAddr =  rsApp("Addr") & ", "& rsApp("City") & ", " & UCase(rsApp("State")) & ", " & rsApp("zip")
	tmpFon =  rsApp("Phone") 
	tmpMobile =  rsApp("Mobile") 
	tmpLang = GetLang(rsApp("LangID")) 
	tmpAppDate = rsApp("AppDate")
	tmpAppTime = rsApp("TimeFrom") & " - " &  rsApp("TimeTo")
	tmpIntrID = rsApp("IntrID")
	tmpCom = rsApp("comment")
	tmpReas = rsApp("reason")
	tmpCall = ""
	If rsApp("callme") = True Then tmpCall = "(* Call patient to remind of appointment)"
End If
rsApp.Close
Set rsApp = Nothing
'GET INTERPRETER
'Set rsIntr = Server.CreateObject("ADODB.RecordSet")
'sqlIntr = "SELECT * FROM Intr_T  WHERE index = " & tmpIntrID
'rsIntr.Open sqlIntr, g_strCONN, 3, 1
'If Not rsIntr.EOF Then
'	tmpIntrName = rsIntr("lname") & ", " & rsIntr("fname")
'	tmpIntrAddr = rsIntr("Addr") & ", " & rsIntr("City") & ", " & UCase(rsIntr("State")) & ", " & rsIntr("Zip")
'End If
'rsIntr.Close
'Set rsIntr = Nothing
'GET INTERPRETER
Set rsIntrInst = Server.CreateObject("ADODB.RecordSet")
sqlIntrInst = "SELECT * FROM InstIntr_T WHERE InstID = " & tmpInstID
rsIntrInst.Open sqlIntrInst, g_strCONN, 3, 1
Do Until rsIntrInst.EOF
	Set rsIntr = Server.CreateObject("ADODB.RecordSet")
	sqlIntr = "SELECT * FROM Intr_T WHERE index = " & rsIntrInst("IntrID")
	rsIntr.Open sqlIntr, g_strCONN, 3, 1
	If Not rsIntr.EOF Then
		tmpI = ""
		If tmpIntrID = "" Then tmpIntrID = 0
		If CInt(tmpIntrID) = rsIntr("index") Then tmpI = "selected"
		tmpIntrName = rsIntr("lname") & ", " &rsIntr("fname")
		strIntr = strIntr & "<option " & tmpI & " value='" & rsIntr("index") & "'>" & tmpIntrName & "</option>" & vbCrLf
		strJScript2 = strJScript2 & "if (Intr == " & rsIntr("Index") & ") " & vbCrLf & _
			"{document.frmAssign.selIntr.value = """ & rsIntr("Index") &"""; " & vbCrLf & _
			"document.frmAssign.txtIntrAddr.value = """ & rsIntr("addr") &"""; " & vbCrLf & _
			"document.frmAssign.txtIntrCity.value = """ & rsIntr("City") &"""; " & vbCrLf & _
			"document.frmAssign.txtIntrState.value = """ & rsIntr("State") &"""; " & vbCrLf & _
			"document.frmAssign.txtIntrZip.value = """ & rsIntr("Zip") &"""; }" & vbCrLf 
	End If
	rsIntr.Close
	Set rsIntr = Nothing
	rsIntrInst.MoveNext
Loop
rsIntrInst.Close
Set rsIntrInst =Nothing
%>
<html>
	<head>
		<title>Interpreter Request - Assign Interpreter</title>
		<link href="CalendarControl.css" type="text/css" rel="stylesheet">
		<script src="CalendarControl.js" language="javascript"></script>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
		<!--
		function CalendarView(strDate)
		{
			document.frmAssign.action = 'calendarview2.asp?appDate=' + strDate;
			document.frmAssign.submit();
		}
		function IntrInfo(Intr)
		{	
			if (Intr == 0)
			{
				document.frmAssign.selIntr.value =0;
				document.frmAssign.txtIntrAddr.value = ""; 
				document.frmAssign.txtIntrCity.value = ""; 
				document.frmAssign.txtIntrState.value = ""; 
				document.frmAssign.txtIntrZip.value = ""; 
			}
			<%=strJScript2%>
		}	
		function AssignMe(xxx)
		{
			document.frmAssign.action = 'action.asp?ctrl=5&ID=' + xxx;
			document.frmAssign.submit();
		}							
		-->
		</script>
		<body onload='IntrInfo(<%=tmpIntrID%>);'>
			<form method='post' name='frmAssign'>
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
									<td class='title' colspan='10' align='center'><nobr>Assign Interpreter</td>
								</tr>
								<tr>
									<td align='center' colspan='10' class='RemME'>
									&nbsp;
									</td>
								</tr>
								<tr>
									<td align='center' colspan='10'><span class='error'><%=Session("MSG")%></span></td>
								</tr>
								<tr>
									<td class='header' colspan='10'><nobr>Contact Information </td>
								</tr>
								<tr>
									<td align='right' width='150px'>Request ID:</td>
									<td class='confirm' ><%=Request("ID")%></td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td align='right'>Institution:</td>
									<td class='confirm'><%=tmpInst%></td>
								</tr>
								<tr>
									<td align='right'>Address:</td>
									<td class='confirm'><%=tmpAddr%></td>
								</tr>
								<tr>
									<td align='right'>Billed To:</td>
									<td class='confirm'><%=tmpBill%></td>
								</tr>
								<tr>
									<td align='right'>Billing Address:</td>
									<td class='confirm'><%=tmpBAddr%></td>
								</tr>
								<tr>
									<td align='right'>E-mail:</td>
									<td class='confirm'><%=tmpEmail%></td>
								</tr>
								<tr>
									<td align='right'>Phone No.:</td>
									<td class='confirm'><%=tmpIphone%></td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr><td colspan='10'><hr align='center' width='75%'></td></tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td colspan='10' class='header'><nobr>Appointment Information</td>
								</tr>
								<tr>
									<td align='right'>Client Name:</td>
									<td class='confirm'><%=tmpCName%>&nbsp;<%=tmpCall%></td>
								</tr>
								<tr>
									<td align='right'>Phone:</td>
									<td class='confirm'><%=tmpFon%></td>
								</tr>
								<tr>
									<td align='right'>Mobile:</td>
									<td class='confirm'><%=tmpMobile%></td>
								</tr>
								<tr>
									<td align='right'>Language:</td>
									<td class='confirm'><%=tmpLang%></td>
								</tr>
								<tr>
									<td align='right'>Appointment Date:</td>
									<td class='confirm'><%=tmpAppDate%></td>
								</tr>
								<tr>
									<td align='right'>Appointment Time:</td>
									<td class='confirm'><%=tmpAppTime%></td>
								</tr>
									<tr>
									<td align='right'>Reason:</td>
									<td class='confirm'><%=tmpReas%></td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr><td colspan='10'><hr align='center' width='75%'></td></tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td colspan='10' class='header'><nobr>Interpreter Information</td>
								</tr>
								<tr>
									<td align='right'>Interpreter:</td>
									<td>
										<select class='seltxt' name='selIntr' style='width: 200px;' onchange='JavaScript:IntrInfo(document.frmAssign.selIntr.value);'>
											<option value='0'>&nbsp;</option>
											<%=strIntr%>
										</select>
									</td>
								</tr>
								<tr>
									<td align='right'>Address:</td>
									<td><input class='main' size='50' maxlength='50' readonly name='txtIntrAddr' value='<%=tmpIntrAddr%>'></td>
								</tr>
								<tr>
									<td align='right'>City:</td>
									<td colspan='5'>
										<input class='main' size='25' maxlength='25' readonly name='txtIntrCity' value='<%=tmpIntrCity%>'>&nbsp;State:
										<input class='main' size='2' maxlength='2' readonly name='txtIntrState' value='<%=tmpIntrState%>'>&nbsp;Zip:
										<input class='main' size='10' maxlength='10' readonly name='txtIntrZip' value='<%=tmpIntrZip%>'>
									</td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr><td colspan='10'><hr align='center' width='75%'></td></tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td align='right'>Comment:</td>
									<td class='confirm'><%=tmpCom%></td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr><td colspan='10'><hr align='center' width='75%'></td></tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td colspan='10' align='center' height='100px' valign='bottom'>
										<input class='btn' type='button' style='width: 125px;' value='View in Calendar' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="document.location='calendarview2.asp?appdate=<%=tmpAppDate%>'">
										<input class='btn' type='button' style='width: 125px;' value='Assign' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='AssignMe(<%=Request("ID")%>);'>
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
	</head>
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