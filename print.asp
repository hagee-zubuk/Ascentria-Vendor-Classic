<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilCalendar.asp" -->
<%
server.scripttimeout = 360000
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
tmpReport = Split(Z_DoDecrypt(Request.Cookies("LBREPORT")), "|")
If tmpReport(0) = 1 Then
	tmpDate = tmpReport(1)
	strMSG = "Appointment schedule of " & Session("GreetMe") & " for " & FormatDateTime(Cdate(tmpDate), 1)
	strHead = "<td class='tblgrn'>Time</td>" & vbCrlf & _
		"<td class='tblgrn'>Client</td>" & vbCrlf & _
		"<td class='tblgrn'>Institution</td>" & vbCrlf & _
		"<td class='tblgrn'>Language</td>" & vbCrlf & _
		"<td class='tblgrn'>Phone No.</td>" & vbCrlf & _
		"<td class='tblgrn'>Mobile</td>" & vbCrlf & _
		"<td class='tblgrn'>Key</td>" & vbCrlf & _
		"<td class='tblgrn' colspan='5'>Reason</td>" & vbCrlf 
	Set rsReq = Server.CreateObject("ADODB.RecordSet")
	HPID = GetHPID(Session("UID"))
	sqlReq = "SELECT * FROM appointment_T WHERE appDate = '" & tmpDate & "' AND IntrID = " & Session("IntrID") & " ORDER BY TimeFrom"
	rsReq.Open sqlReq, g_strCONN, 1, 3
	y = 0
	Do Until rsReq.EOF
		kulay = "#FFFFFF"
		If Not Z_IsOdd(y) Then kulay = "#F5F5F5"
			tmptime = rsReq("TimeFrom") & " - " & rsReq("TimeTo")
			tmpClient = Z_DoDecrypt(rsReq("CLname")) & ", " & Z_DoDecrypt(rsReq("CFname"))
			tmpInst = GetFacility(rsReq("InstID"))
			If GetDept(rsReq("DeptID")) <> "N/A" Then tmpInst = tmpInst & " - " & GetDept(rsReq("DeptID"))
			tmpLang = GetLangLB(rsReq("LangID"))
			
		strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & tmptime & "</td>" & vbCrLf & _
			"<td class='tblgrn2'><nobr>" & tmpClient & "</td>" & vbCrLf & _
			"<td class='tblgrn2'><nobr>" & tmpInst & "</td>" & vbCrLf & _
			"<td class='tblgrn2'><nobr>" & tmpLang & "</td>" & vbCrLf & _
			"<td class='tblgrn2'><nobr>" & Z_DoDecrypt(rsReq("phone")) & "</td>" & vbCrLf & _ 
			"<td class='tblgrn2'><nobr>" & Z_DoDecrypt(rsReq("mobile")) & "</td>" & vbCrLf & _ 
			"<td class='tblgrn2' width='100px'><nobr>&nbsp;</td>" & vbCrLf & _
			"<td class='tblgrn2'><nobr>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td></tr>" & vbCrLf
		'y = y + 1
		rsReq.MoveNext
	Loop
	rsReq.Close
	Set rsReq = Nothing
ElseIf tmpReport(0) = 2 Then
	tmpDate = tmpReport(1)
	tmpMonth = Month(tmpDate)
	tmpYear = Year(tmpDate)
	strHead = "<td class='tblgrn'>Date</td>" & vbCrlf & _
			"<td class='tblgrn'>Time</td>" & vbCrlf & _
			"<td class='tblgrn'>Client</td>" & vbCrlf & _
			"<td class='tblgrn'>Language</td>" & vbCrlf & _
			"<td class='tblgrn'>Interpreter</td>" & vbCrlf
	CSVHead = "Date, Time, Client First Name, Client Last Name, Language, Interpreter Last Name, Interpreter First Name"
	If tmpReport(2) = 0 Then
		If Session("type") = 0 Or Session("type") = 4 Or Session("type") = 5 Then
			strMSG = "Appointment schedule of " & GetInstNameLB(Session("InstID")) & " for " & tmpDate
			If Session("InstID") = 93 Then
				sqlRep = "SELECT * FROM appointment_T WHERE appDate = '" & tmpDate & "' AND InstID = 93 ORDER BY TimeFrom"
			ElseIf Session("UID") = 509 Or Session("UID") = 510 Then 'special rule for user 509 and 510
				sqlRep = "SELECT * FROM appointment_T WHERE appDate = '" & tmpDate & "' AND " & _
					"(DeptID = 2306 OR DeptID = 373 OR DeptID = 946 OR DeptID = 2302) ORDER BY TimeFrom"
			ElseIf Session("UID") = 517 Or Session("UID") = 534 Then 'special rule for user 517 and 534
				sqlRep = "SELECT * FROM appointment_T WHERE appDate = '" & tmpDate & "' AND " & _
					"(DeptID = 446 OR DeptID = 322 OR DeptID = 289) ORDER BY TimeFrom"
			Else
				sqlRep = "SELECT * FROM appointment_T WHERE appDate = '" & tmpDate & "' AND InstID = " & Session("InstID") & " ORDER BY TimeFrom"
			End If
		ElseIf Session("type") = 1 Then
			sqlRep = "SELECT * FROM appointment_T WHERE appDate = '" & tmpDate & "' AND IntrID = " & Session("IntrID") & " ORDER BY TimeFrom"
		ElseIf Session("type") = 2 Then
			sqlRep = "SELECT * FROM appointment_T WHERE appDate = '" & tmpDate & "' ORDER BY TimeFrom"
		ElseIf Session("type") = 3 Then
			strMSG = "Appointment schedule of " & GetInstNameLB(Session("InstID")) & " - " & GetDeptNameLB(Session("DeptID")) & " for " & tmpDate
			sqlRep = "SELECT * FROM appointment_T WHERE appDate = '" & tmpDate & "' AND DeptID = " & Session("DeptID") & " ORDER BY TimeFrom"
		ElseIf Session("type") = 6 Then
			strMSG = "Court appointment schedule for " & tmpDate
			'sqlRep = "SELECT * FROM appointment_T WHERE InstID = " & Session("InstID") & " AND DeptID = " & Session("DeptID") & " AND appDate = '" & tmpDate & "' ORDER BY TimeFrom"
			Set fso = CreateObject("Scripting.FileSystemObject")
			Set oFile = fso.OpenTextFile(crtLst, 1)
			Do Until oFile.AtEndOfStream
				oLine = oFile.ReadLine
				strInst = strInst & "InstID = " & oLine & " OR "
			Loop
			Set oFile = Nothing
			Set fso = Nothing
			sqlInst = Mid(strInst, 1, Len(strInst) - 4)
			sqlRep = "SELECT * FROM appointment_T WHERE appDate = '" & tmpDate & "' AND (" & sqlInst & ") ORDER BY TimeFrom"
		End If
	ElseIf tmpReport(2) = 1 Then
		tmpSun = GetSun(tmpDate)
		tmpSat = GetSat(tmpDate)
		If Session("type") = 0 Or Session("type") = 4 Or Session("type") = 5 Then
			strMSG = "Appointment schedule of " & GetInstNameLB(Session("InstID")) & " for " & tmpSun & " to " & tmpSat
			If Session("InstID") = 93 Then
				sqlRep = "SELECT * FROM appointment_T WHERE appDate >= '" & tmpSun & "' AND appDate <= '" & tmpSat & "' AND InstID = 93 ORDER BY TimeFrom"
			ElseIf Session("UID") = 509 Or Session("UID") = 510 Then 'special rule for user 509 and 510
				sqlRep = "SELECT * FROM appointment_T WHERE appDate >= '" & tmpSun & "' AND appDate <= '" & tmpSat & "' AND " & _
					"(DeptID = 2306 OR DeptID = 373 OR DeptID = 946 OR DeptID = 2302) ORDER BY TimeFrom"
			ElseIf Session("UID") = 517 Then 'special rule for user 517
				sqlRep = "SELECT * FROM appointment_T WHERE appDate >= '" & tmpSun & "' AND appDate <= '" & tmpSat & "' AND " & _
					"(DeptID = 446 OR DeptID = 322 OR DeptID = 289) ORDER BY TimeFrom"
			Else
				sqlRep = "SELECT * FROM appointment_T WHERE appDate >= '" & tmpSun & "' AND appDate <= '" & tmpSat & "' AND InstID = " & Session("InstID") & " ORDER BY appDate, TimeFrom"
			End If
		ElseIf Session("type") = 1 Then
			sqlRep = "SELECT * FROM appointment_T WHERE appDate >= '" & tmpSun & "' AND appDate <= '" & tmpSat & "' AND IntrID = " & Session("IntrID") & " ORDER BY appDate, TimeFrom"
		ElseIf Session("type") = 2 Then
			sqlRep = "SELECT * FROM appointment_T WHERE appDate >= '" & tmpSun & "' AND appDate <= '" & tmpSat & "' ORDER BY appDate, TimeFrom"
		ElseIf Session("type") = 3 Then
			strMSG = "Appointment schedule of " & GetInstNameLB(Session("InstID")) & " - " & GetDeptNameLB(Session("DeptID")) & " for " & tmpSun & " to " & tmpSat
			sqlRep = "SELECT * FROM appointment_T WHERE appDate >= '" & tmpSun & "' AND appDate <= '" & tmpSat & "' AND DeptID = " & Session("DeptID") & " ORDER BY appDate, TimeFrom"
		ElseIf Session("type") = 6 Then
			strMSG = "Court appointment schedule for " & tmpSun & " to " & tmpSat
			'sqlRep = "SELECT * FROM appointment_T WHERE InstID = " & Session("InstID") & " AND DeptID = " & Session("DeptID") & " AND appDate = '" & tmpDate & "' ORDER BY TimeFrom"
			Set fso = CreateObject("Scripting.FileSystemObject")
			Set oFile = fso.OpenTextFile(crtLst, 1)
			Do Until oFile.AtEndOfStream
				oLine = oFile.ReadLine
				strInst = strInst & "InstID = " & oLine & " OR "
			Loop
			Set oFile = Nothing
			Set fso = Nothing
			sqlInst = Mid(strInst, 1, Len(strInst) - 4)
			sqlRep = "SELECT * FROM appointment_T WHERE appDate >= '" & tmpSun & "' AND appDate <= '" & tmpSat & "' AND (" & sqlInst & ") ORDER BY appDate, TimeFrom"
		End If
	End If
	
	Set rsRep = Server.CreateObject("ADODB.RecordSet")
	rsRep.Open sqlRep, g_strCONN, 3, 1
	y = 0
	Do Until rsRep.EOF
		If Not Z_HideApp(rsRep("UID")) Or Session("UID") = rsRep("UID") Then
			kulay = "#FFFFFF"
			If Not Z_IsOdd(y) Then kulay = "#F5F5F5"
			mySTat = ""
			If rsRep("status") = 3 Or rsRep("status") = 4 Then mySTat = "(canceled)"
				tmpIntr = GetIntrNameLB(rsRep("IntrID"))
				strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & rsRep("appDate") & "</td>" & vbCrLf & _
						"<td class='tblgrn2'><nobr>" & rsRep("TimeFrom") & " - " & rsRep("TimeTo") & "</td>" & vbCrLf & _
						"<td class='tblgrn2'><nobr>" & Z_DoDecrypt(rsRep("CLname")) & ", " & Z_DoDecrypt(rsRep("CFname")) & " " & mySTat & "</td>" & vbCrLf & _
						"<td class='tblgrn2'><nobr>" & GetLang(rsRep("LangID")) & "</td>" & vbCrLf & _
						"<td class='tblgrn2'><nobr>" & tmpIntr & "</td>" & vbCrLf & _
						"</tr>" & vbCrLf
					CSVBody = CSVBody & rsRep("appDate") & "," & rsRep("TimeFrom") & "," & rsRep("CLname") & "," & _
							rsRep("CFname") & "," & GetLang(rsRep("LangID")) & "," & vbCrLf
			'End If
			rsRep.MoveNext
			y = y + 1
		End If
	Loop
	rsRep.Close
	Set rsRep = Nothing
End If
%>
<html>
	<head>
		<title>Interpreter Request - Schedule</title>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
		<!--
		function exportMe()
		{
			document.frmRep.action = "printreport.asp?csv=1"
			document.frmRep.submit();
		}
		-->
		</script>
		<body>
			<form method='post' name='frmRep'>
				<table cellSpacing='0' cellPadding='0' width="100%"  height='100%' bgColor='white' border='0'>
					<tr>
						<td valign='top'>
							<table bgColor='white' border='0' cellSpacing='0' cellPadding='0' align='center'>
							<tr>
								<td>
									<img src='images/LBISLOGO.jpg' align='center'>
								</td>
							</tr>
							<tr>
								<td align='center'>
									261&nbsp;Sheep&nbsp;Davis&nbsp;Road,&nbsp;Concord,&nbsp;NH&nbsp;03301<br>
									Tel:&nbsp;(603)&nbsp;410-6183&nbsp;|&nbsp;Fax:&nbsp;(603)&nbsp;410-6186
								</td>
							</tr>
							</table>
						</td>
					</tr>
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td valign='top' >
							<table bgColor='white' border='0' cellSpacing='4' cellPadding='0' align='center'>
								<tr bgcolor='#C2AB4B'>
									<td colspan='12' align='center'>
										<b><%=strMSG%><b>
									</td>
								</tr>
								<tr>
									<%=strHead%>
								</tr>
								<%=strBody%>
								<tr><td>&nbsp;</td></tr>
								
							</table>
						</tr>
						<tr><td>&nbsp;</td></tr>
						<tr>
							<td align='left' valign='bottom'>
								<% If tmpReport(0) <> 2 Then %>
									<!-- #include file="_keyguide.asp" -->
								<% End If %>
								<br><br>
								<b>&nbsp;&nbsp;* If needed, please adjust the page orientation of your printer to landscape to view all columns in a single page</b>
								<br>
								<b>&nbsp;&nbsp;* Please SHRED this sheet at the end of the day</b>
							</td>
						</tr>
						<tr>
							<td colspan='12' align='center' height='100px' valign='bottom'>
								<input class='btn' type='button' value='Print' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='print();'>
							</td>
						</tr>
				</table>
			</form>
		</body>
	</head>
</html>
