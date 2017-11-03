<%Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilCalendar.asp" -->
<%
Function SupXREnt(blnPT, intUID, intDep)
	SupXREnt=  True
	If Not blnPT Then Exit Function
	Set rsXR = Server.CreateObject("ADODB.RecordSet")
	rsXR.Open "SELECT [dept_id] FROM [xr_user_dept] WHERE [dept_id]=" & intDep & " AND [user_id]=" & intUID , g_strCONN, 3, 1
	SupXREnt = Not rsXR.EOF
	rsXR.Close
	Set rsXR = Nothing
End Function
Function UseXRTable(uid) 
	UseXRTable = False
	Set rsXR = Server.CreateObject("ADODB.RecordSet")
	rsXR.Open "SELECT COUNT([user_id]) AS [cnt] FROM [xr_user_dept] WHERE [user_id]=" & uid , g_strCONN, 3, 1
	If Not rsXR.EOF Then
		UseXRTable = CBool(rsXR("cnt") > 0)
	End If
	rsXR.Close
	Set rsXR = Nothing
End Function
Function IsBlock(xxx)
	IsBlock = False
	Set rsUser = Server.CreateObject("ADODB.RecordSet")
	sqlUser = "SELECT block FROM appointment_t WHERE [index] = " & xxx
	rsUser.Open sqlUser, g_StrCONN, 3, 1
	If Not rsUser.EOF Then
		'If InStr(Ucase(rsUser("lbcom")), "BLOCK SCHEDULE") > 0 Then IsBlock = True
		'If InStr(Ucase(rsUser("comment")), "BLOCK SCHEDULE") > 0 Then IsBlock = True
		If rsUser("block") Then IsBlock = True
	End If
	rsUser.Close
	Set rsUser = Nothing
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
Function MyStatus(xxx)
	Select Case xxx
		Case 1
			MyStatus = "<font color='#000000' size='+3'>•</font>"
		Case 2
			MyStatus = "<font color='#0000FF' size='+3'>•</font>"
		Case 3
			MyStatus = "<font color='#FF0000' size='+3'>•</font>"
		Case 4
			MyStatus = "<font color='#FF00FF' size='+3'>•</font>"
		Case Else
			MyStatus = ""
	End Select
End Function

If Session("type") = 1 Then HPID = GetHPID(Session("UID"))
CalendarPage = True
tmpReqMonth = Request("selMonth")
tmpReqYear = Request("txtyear")
If Request("appdate") <> "" Then
		tmpAppDate = Split(Request("appdate"), "/")
		tmpReqMonth = tmpAppDate(0)
		tmpReqYear = tmpAppDate(2)
End If
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
	tmp1Day = tmpReqMonth & "/01/" & tmpReqYear
	tmpMonth = MonthName(tmpReqMonth) & " - " & tmpReqYear
End If
'SET CALENDAR
If tmpReqMonth <> "" And tmpReqYear <> "" Then
	tmp1Day = tmpReqMonth & "/01/" & tmpReqYear
	tmpMonth = MonthName(tmpReqMonth) & " - " & tmpReqYear
End If
If tmp1Day = "" Then 
	tmp1Day = Month(Date) & "/01/" & Year(Date)
	tmpMonth = MonthName(Month(Date)) & " - " & Year(Date)
End If
If Not IsDate(tmp1Day) Then 
	tmp1day = Month(Date) & "/01/" & Year(Date)
	tmpMonth = MonthName(Month(Date)) & " - " & Year(Date)
	Session("MSG") = "ERROR: Year inputted is not valid. Set to current month and year."
End If
CorrectMonth = True
tmpToday = tmp1Day
Do While CorrectMonth = True 
	strCal = strCAL & "<tr><td>&nbsp;</td>"
	If WeekdayName(Weekday(tmpToday), True) <> "Sun" Then 
		strCal = strCAL & "<td>&nbsp;</td>"
	Else
		strMonth = Month(tmpToday) 
		strDay = Day(tmpToday)
		strYear = Year(tmpToday)
		tmpBG = CheckApp(tmpToday)
		strCal = strCAL & "<td bgcolor='" & tmpBG & "' class='caltbl' valign='top' onmouseover=""this.className='caltbl2'"" onmouseout=""this.className='caltbl'"" onclick='GoToday(" & strMonth & "," & strDay & "," & strYear & ");'>" & Day(tmpToday) & "<br>" & strApp & "</td>" & vbCrLf
		tmpToday = DateAdd("d", 1, tmpToday)
		If Month(tmp1Day) <> Month(tmpToday) Then Exit Do
	End If	
	If WeekdayName(Weekday(tmpToday), True) <> "Mon" Then 
		strCal = strCAL & "<td>&nbsp;</td>"
	Else
		strMonth = Month(tmpToday) 
		strDay = Day(tmpToday)
		strYear = Year(tmpToday)
		tmpBG = CheckApp(tmpToday)
		strCal = strCAL & "<td bgcolor='" & tmpBG & "' class='caltbl' valign='top' onmouseover=""this.className='caltbl2'"" onmouseout=""this.className='caltbl'"" onclick='GoToday(" & strMonth & "," & strDay & "," & strYear & ");'>" & Day(tmpToday) & "<br>" & strApp & "</td>" & vbCrLf
		tmpToday = DateAdd("d", 1, tmpToday)
		If Month(tmp1Day) <> Month(tmpToday) Then Exit Do
	End If	
	If WeekdayName(Weekday(tmpToday), True) <> "Tue" Then 
		strCal = strCAL & "<td>&nbsp;</td>"
	Else
		strMonth = Month(tmpToday) 
		strDay = Day(tmpToday)
		strYear = Year(tmpToday)
		tmpBG = CheckApp(tmpToday)
		strCal = strCAL & "<td bgcolor='" & tmpBG & "' class='caltbl' valign='top' onmouseover=""this.className='caltbl2'"" onmouseout=""this.className='caltbl'"" onclick='GoToday(" & strMonth & "," & strDay & "," & strYear & ");'>" & Day(tmpToday) & "<br>" & strApp & "</td>" & vbCrLf
		tmpToday = DateAdd("d", 1, tmpToday)
		If Month(tmp1Day) <> Month(tmpToday) Then Exit Do
	End If	
	If WeekdayName(Weekday(tmpToday), True) <> "Wed" Then 
		strCal = strCAL & "<td>&nbsp;</td>"
	Else
		strMonth = Month(tmpToday) 
		strDay = Day(tmpToday)
		strYear = Year(tmpToday)
		tmpBG = CheckApp(tmpToday)
		strCal = strCAL & "<td bgcolor='" & tmpBG & "' class='caltbl' valign='top' onmouseover=""this.className='caltbl2'"" onmouseout=""this.className='caltbl'"" onclick='GoToday(" & strMonth & "," & strDay & "," & strYear & ");'>" & Day(tmpToday) & "<br>" & strApp & "</td>" & vbCrLf
		tmpToday = DateAdd("d", 1, tmpToday)
		If Month(tmp1Day) <> Month(tmpToday) Then Exit Do
	End If	
	If WeekdayName(Weekday(tmpToday), True) <> "Thu" Then 
		strCal = strCAL & "<td>&nbsp;</td>"
	Else
		strMonth = Month(tmpToday) 
		strDay = Day(tmpToday)
		strYear = Year(tmpToday)
		tmpBG = CheckApp(tmpToday)
		strCal = strCAL & "<td bgcolor='" & tmpBG & "' class='caltbl' valign='top' onmouseover=""this.className='caltbl2'"" onmouseout=""this.className='caltbl'"" onclick='GoToday(" & strMonth & "," & strDay & "," & strYear & ");'>" & Day(tmpToday) & "<br>" & strApp & "</td>" & vbCrLf
		tmpToday = DateAdd("d", 1, tmpToday)
		If Month(tmp1Day) <> Month(tmpToday) Then Exit Do
	End If	
	If WeekdayName(Weekday(tmpToday), True) <> "Fri" Then 
		strCal = strCAL & "<td>&nbsp;</td>"
	Else
		strMonth = Month(tmpToday) 
		strDay = Day(tmpToday)
		strYear = Year(tmpToday)
		tmpBG = CheckApp(tmpToday)
		strCal = strCAL & "<td bgcolor='" & tmpBG & "' class='caltbl' valign='top' onmouseover=""this.className='caltbl2'"" onmouseout=""this.className='caltbl'"" onclick='GoToday(" & strMonth & "," & strDay & "," & strYear & ");'>" & Day(tmpToday) & "<br>" & strApp & "</td>" & vbCrLf
		tmpToday = DateAdd("d", 1, tmpToday)
		If Month(tmp1Day) <> Month(tmpToday) Then Exit Do
	End If	
	If WeekdayName(Weekday(tmpToday), True) <> "Sat" Then 
		strCal = strCAL & "<td>&nbsp;</td>"
	Else
		strMonth = Month(tmpToday) 
		strDay = Day(tmpToday)
		strYear = Year(tmpToday)
		tmpBG = CheckApp(tmpToday)
		strCal = strCAL & "<td bgcolor='" & tmpBG & "' class='caltbl' valign='top' onmouseover=""this.className='caltbl2'"" onmouseout=""this.className='caltbl'"" onclick='GoToday(" & strMonth & "," & strDay & "," & strYear & ");'>" & Day(tmpToday) & "<br>" & strApp & "</td>" & vbCrLf
		tmpToday = DateAdd("d", 1, tmpToday)
		If Month(tmp1Day) <> Month(tmpToday) Then Exit Do
	End If	
	strCal = strCAL & "</tr>"
	If Month(tmp1Day) <> Month(tmpToday) Then CorrectMonth = False
Loop
'SET ORGANIZER
If Request("txtday") <> "" Then
	tmpDate = tmpReqMonth & "/" & Request("txtday") & "/" & tmpReqYear
Else
	tmpDate = Date
	If tmpReqMonth <> "" And tmpReqYear <> "" Then tmpDate = tmp1Day
	If Request("appdate") <> "" Then tmpDate = Request("appdate")
End If
Set rsReq = Server.CreateObject("ADODB.RecordSet")
'Response.Write "TYPE:" & Session("type") & "<br />"
If Session("type") = 0 Or Session("type") = 4 Or Session("type") = 5 Then
	'If GetInstID(Session("UID")) = 4 Or GetInstID(Session("UID")) = 6 Or GetInstID(Session("UID")) = 7 Then 'FOR SNHMC
	If Session("InstID") = 93 Then
		sqlReq = "SELECT * FROM appointment_T WHERE appDate = '" & tmpDate & "' AND InstID = 93 ORDER BY TimeFrom"
	ElseIf Session("UID") = 509 Or Session("UID") = 510 Then 'special rule for user 509 and 510
		sqlReq = "SELECT * FROM appointment_T WHERE appDate = '" & tmpDate & "' AND " & _
			"(DeptID = 2306 OR DeptID = 373 OR DeptID = 946 OR DeptID = 2302) ORDER BY TimeFrom"
	ElseIf Session("UID") = 517 Or Session("UID") = 534 Then 'special rule for user 517 and 534
		sqlReq = "SELECT * FROM appointment_T WHERE appDate = '" & tmpDate & "' AND " & _
			"(DeptID = 446 OR DeptID = 322 OR DeptID = 289) ORDER BY TimeFrom"
	ElseIf Session("UID") = 625 Or Session("UID") = 626 Or Session("UID") = 627 Or Session("UID") = 628 Then 'special rule for user 625 - 628
		sqlReq = "SELECT * FROM appointment_T WHERE appDate = '" & tmpDate & "' AND " & _
			"(DeptID = 2466 OR DeptID = 2465) ORDER BY TimeFrom"
	Else
		sqlReq = "SELECT * FROM appointment_T WHERE appDate = '" & tmpDate & "' AND InstID = " & Session("InstID") & " ORDER BY TimeFrom"
	End If
ElseIf Session("type") = 1 Then
	sqlReq = "SELECT * FROM appointment_T WHERE appDate = '" & tmpDate & "' AND IntrID = " & Session("IntrID") & " ORDER BY TimeFrom"
ElseIf Session("type") = 2 Then
	sqlReq = "SELECT * FROM appointment_T WHERE appDate = '" & tmpDate & "' ORDER BY TimeFrom"
ElseIf Session("type") = 3 Then
	sqlReq = "SELECT * FROM appointment_T WHERE appDate = '" & tmpDate & "' AND DeptID = " & Session("DeptID") & " ORDER BY TimeFrom"
ElseIf Session("type") = 6 Then
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set oFile = fso.OpenTextFile(crtLst, 1)
	Do Until oFile.AtEndOfStream
		oLine = oFile.ReadLine
		strInst = strInst & "InstID = " & oLine & " OR "
	Loop
	Set oFile = Nothing
	Set fso = Nothing
	sqlInst = Mid(strInst, 1, Len(strInst) - 4)
	sqlReq = "SELECT * FROM appointment_T WHERE appDate = '" & tmpDate & "' AND (" & sqlInst & ") ORDER BY TimeFrom"
End If
rsReq.Open sqlReq, g_strCONN, 3, 1
' Response.Write "<code>" & sqlReq & "</code><br />User: " & rsReq("UID") & "|" & Session("UID") & "<br />"
blnXRZ = UseXRTable( Session("UID") ) 
If Not rsReq.EOF Then
	Do Until rsReq.EOF
		If (Not Z_HideApp(rsReq("UID")) Or Session("type") = 3) And SupXREnt(blnXRZ, Session("UID"), rsReq("DeptID")) Then
			tmpInst = GetFacility(rsReq("InstID")) & " - " & GetDept(rsReq("DeptID"))
			'If Session("type") <> 1 Then
			'response.write "IntrID: " & rsReq("IntrID") & "<br>"
				tmpIntrID = rsReq("IntrID")
				tmpIntr = GetIntrNameLB(rsReq("IntrID"))
			'Else
			'	If rsReq("mwhere") = 0 Then
			'		tmpIntr = rsReq("mAddr") & ", "& rsReq("mCity") & ", " & UCase(rsReq("mState")) & ", " & rsReq("mzip")
			'	Else
			'		tmpIntr = UCase(GetLoc(rsReq("mlocation")))
			'		If tmpIntr = "OTHER" Then tmpIntr = rsReq("mother")
			'	End If
			'End If
			tmpLang = GetLang(rsReq("LangID"))
			ColLang = ""
			ColLang2 = ""
			If UCase(tmpLang) = "SPANISH" Or UCase(tmpLang) = "PORTUGUESE" Then 
				ColLang = "<font color='#FF9900'>"
				ColLang2 = "</font>"
			End If
			tmpTime = CTime(rsReq("TimeFrom")) 
			If Z_fixnull(rsReq("TimeTo")) <> "" Then 
				strtmpTime = tmpTime & " - " & CTime(rsReq("TimeTo"))
			Else
				strtmpTime = tmpTime
			End If
			'tmpAppTime = rsReq("TimeFrom") 
			'If rsReq("TimeTo") <> Empty Then tmpAppTime = tmpAppTime & " - " & rsReq("TimeTo")
			tmpPhone = "N/A"
			If rsReq("phone") <> "" Then tmpPhone = rsReq("phone") 
			tmpPhone2 = "N/A"
			If rsReq("mobile") <> "" Then tmpPhone2 = rsReq("mobile") 
			tmp12 = cdate("12:00 AM")
			tmp1259 = cdate("12:59 AM")
			tmpID = rsReq("Index")
			cbk = "#F5F5F5"
			If rsReq("InstID") = 273 And rsReq("LangID") = 33 And (Weekday(tmpDate) = 2 Or Weekday(tmpDate) = 3 Or Weekday(tmpDate) = 4 _
				Or Weekday(tmpDate) = 5) Then 
					If IsBlock(tmpID) Then cbk = "#FFFFCE"
			End If	
			tmpstr = "<tr bgcolor='" & cbk & "' onclick='PassMe(" & tmpID & ");'>" & vbCrLf &" <td align='center'>" &  MyStatus(GetStatLB(tmpID)) & "</td><td align='center'><nobr>" & strtmpTime & "</td><td align='center'>" &  Z_DoDecrypt(rsReq("clname")) & ", " & Z_DoDecrypt(rsReq("cfname")) & "</td>" & _
				"<td align='center'>" & tmpInst & "</td><td align='center'>" & ColLang & tmpLang & ColLang2 & "</td>" & _
				"<td align='center'>" & tmpIntr & "</td><td align='center'><nobr>" &  Z_DoDecrypt(tmpPhone) & "</td><td align='center'><nobr>" &  Z_DoDecrypt(tmpPhone2) & "</td>"
			If Z_CZero(tmpIntrID) > 0 Then
				tmpstr = tmpstr & "<td align='center'><input style='width: 75px;' type='button' class='btn' onmouseover=""this.className='hovbtn'"" onmouseout=""this.className='btn'"" value='Feedback' onclick='mySurvey(" & tmpID & ");' ></td>"
			End If
			tmpstr = tmpstr & "</tr>" & vbCrLf
				'response.write tmpTime & " < " &  DateAdd("H", 1, tmp1259) & " " & cbool(tmpTime < DateAdd("H", 1, tmp1259))
			If tmpTime >= tmp12 And tmpTime < tmp1259 Then
				str12a = str12a & tmpstr
				x12a = x12a + 1
			ElseIf tmpTime >= DateAdd("H", 1, tmp12) And tmpTime < DateAdd("H", 1, tmp1259) Then
				str1a = str1a & tmpstr
							'response.write tmpTime
				x1a = x1a + 1
			ElseIf tmpTime >= DateAdd("H", 2, tmp12) And tmpTime < DateAdd("H", 2, tmp1259) Then
				str2a = str2a & tmpstr
				x2a = x2a + 1
			ElseIf tmpTime >= DateAdd("H", 3, tmp12) And tmpTime < DateAdd("H", 3, tmp1259) Then
				str3a = str3a & tmpstr
				x3a = x3a + 1
			ElseIf tmpTime >= DateAdd("H", 4, tmp12) And tmpTime < DateAdd("H", 4, tmp1259) Then
				str4a = str4a & tmpstr
				x4a = x4a + 1
			ElseIf tmpTime >= DateAdd("H", 5, tmp12) And tmpTime < DateAdd("H", 5, tmp1259) Then
				str5a = str5a & tmpstr
				x5a =  x5a + 1
			ElseIf tmpTime >= DateAdd("H", 6, tmp12) And tmpTime < DateAdd("H", 6, tmp1259) Then
				str6a = str6a & tmpstr
				x6a = x6a + 1
			ElseIf tmpTime >= DateAdd("H", 7, tmp12) And tmpTime < DateAdd("H", 7, tmp1259) Then
				str7a = str7a & tmpstr
				x7a = x7a + 1
			ElseIf tmpTime >= DateAdd("H", 8, tmp12) And tmpTime < DateAdd("H", 8, tmp1259) Then
				str8a = str8a & tmpstr
				x8a = x8a + 1
			ElseIf tmpTime >= DateAdd("H", 9, tmp12) And tmpTime < DateAdd("H", 9, tmp1259) Then
				str9a = str9a & tmpstr
				x9a = x9a + 1
			ElseIf tmpTime >= DateAdd("H", 10, tmp12) And tmpTime < DateAdd("H", 10, tmp1259) Then
				str10a = str10a & tmpstr
				x10a = x10a + 1
			ElseIf tmpTime >= DateAdd("H", 11, tmp12) And tmpTime < DateAdd("H", 11, tmp1259) Then
				str11a = str11a & tmpstr
				x11a = x11a + 1
			ElseIf tmpTime >= DateAdd("H", 12, tmp12) And tmpTime < DateAdd("H", 12, tmp1259) Then
				str12p = str12p & tmpstr
				x12p = x12p + 1
			ElseIf tmpTime >= DateAdd("H", 13, tmp12) And tmpTime < DateAdd("H", 13, tmp1259) Then
				str1p = str1p & tmpstr
				x1p = x1p + 1
			ElseIf tmpTime >= DateAdd("H", 14, tmp12) And tmpTime < DateAdd("H", 14, tmp1259) Then
				str2p = str2p & tmpstr
				x2p = x2p + 1
			ElseIf tmpTime >= DateAdd("H", 15, tmp12) And tmpTime < DateAdd("H", 15, tmp1259) Then
				str3p = str3p & tmpstr
				x3p = x3p + 1
			ElseIf tmpTime >= DateAdd("H", 16, tmp12) And tmpTime < DateAdd("H", 16, tmp1259) Then
				str4p = str4p & tmpstr
				x4p = x4p + 1
			ElseIf tmpTime >= DateAdd("H", 17, tmp12) And tmpTime < DateAdd("H", 17, tmp1259) Then
				str5p = str5p & tmpstr
				x5p = x5p + 1
			ElseIf tmpTime >= DateAdd("H", 18, tmp12) And tmpTime < DateAdd("H", 18, tmp1259) Then
				str6p = str6p & tmpstr
				x6p = x6p + 1
			ElseIf tmpTime >= DateAdd("H", 19, tmp12) And tmpTime < DateAdd("H", 19, tmp1259) Then
				str7p = str7p & tmpstr
				x7p = x7p + 1
			ElseIf tmpTime >= DateAdd("H", 20, tmp12) And tmpTime < DateAdd("H", 20, tmp1259) Then
				str8p = str8p & tmpstr
				x8p = x8p + 1
			ElseIf tmpTime >= DateAdd("H", 21, tmp12) And tmpTime < DateAdd("H", 21, tmp1259) Then
				str9p = str9p & tmpstr
				x9p = x9p + 1
			ElseIf tmpTime >= DateAdd("H", 22, tmp12) And tmpTime < DateAdd("H", 22, tmp1259) Then
				str10p = str10p & tmpstr
				x10p = x1a + 1
			ElseIf tmpTime >= DateAdd("H", 23, tmp12) And tmpTime < DateAdd("H", 23, tmp1259) Then
				str11p = str11p & tmpstr
				x11p = x11p + 1
			End If
		End If
		rsReq.MoveNext
	Loop
End If
rsReq.Close
Set rsReq = Nothing

%>
<html>
	<head>
		<title>Interpreter Request - Calendar View</title>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
		<!--
		function PassMe(xxx)
		{
			document.frmCal.action = "reqconfirm.asp?ID=" + xxx;
			document.frmCal.submit();
		}
		function ChangeMonth(xxx)
		{
			document.frmCal.action = "action.asp?ctrl=4&dir=" + xxx;
			document.frmCal.submit();
		}
		function SearchMonth()
		{
			if (document.frmCal.txtyear.value == "")
			{
				alert("ERROR: Year is required.");
			}
			else
			{
				document.frmCal.action = "calendarview2.asp#am7";
				document.frmCal.submit();
			}
		}
		function GoToday(xm, xd, xy)
		{
			document.frmCal.action = "calendarview2.asp?selMonth=" + xm + "&txtday=" + xd + "&txtyear=" + xy + "#am7";
			document.frmCal.submit();
		}
		function PublishMe2()
		{
			var xdate = "<%=tmpDate%>";
			document.frmCal.action = "action.asp?ctrl=16&type=0&xdate=" + xdate;
			document.frmCal.submit();	
		}
		function PublishMe3()
		{
			var xdate = "<%=tmpDate%>";
			document.frmCal.action = "action.asp?ctrl=16&type=1&xdate=" + xdate;
			document.frmCal.submit();	
		}
		function PrintToday()
		{
			var xdate = "<%=tmpDate%>";
			document.frmCal.action = "action.asp?ctrl=6&xdate=" + xdate;
			document.frmCal.submit();
		}
		function PopMe(xxx)
		{
			if (xxx !== 0)
			{
				newwindow = window.open('print.asp','','height=850,width=900,scrollbars=1,directories=0,status=0,toolbar=0,resizable=1');
				if (window.focus) {newwindow.focus()}
			}
		}
		function mySurvey(xxx) {
			newwindow = window.open('survey.asp?ID=' + xxx ,'','height=800,width=900,scrollbars=1,directories=0,status=0,toolbar=0,resizable=1');
			if (window.focus) {newwindow.focus()}
		}
		-->
		</script>
		<style type="text/css">
	 	.container
	      {
	          border: solid 1px black;
	          overflow: auto;
	      }
	      .noscroll
	      {
	          position: relative;
	          background-color: white;
	          top: expression(this.offsetParent.scrollTop);
	      }
	      th
	      {
	          text-align: left;
	      }
		</style>
	</head>
	<body 
		<%If Request("rpt") <> "" Then%>
			onload='PopMe(<%=Request("rpt")%>);'
		<%End If%>
		>
		<form method='post' name='frmCal'>
			<table cellSpacing='0' cellPadding='0' height='100%' width="100%" border='0' class='bgstyle2'>
				<tr>
					<td height='100px' valign='top'>
						<!-- #include file="_header.asp" -->
					</td>
				</tr>
				<tr>
					<td valign='top'>
						<table cellSpacing='2' cellPadding='2' border='0' align='center'>
							<!-- #include file="_greetme.asp" -->
								<tr>
								<td style='width: 75%;' valign='top'>
									<table cellSpacing='2' cellPadding='2' style='width: 100%;' border='0' align='center'>
										<tr>
											<td colspan='2' align='center' class='timeheader'>
												<%=FormatDateTime(Cdate(tmpDate), 1)%>
											</td>
										</tr>
										<tr>
											<td>
												<div class='container' id='cal' style='height: 440px; width:100%; position: relative;'>
													<table cellSpacing='2' cellPadding='0' height='100%' width='100%' border='0' align='left' bgcolor='#FFFFFF'>
														<thead>
															<tr bgcolor='#D4D0C8' class="noscroll">
																<td align='center'  class='time2'>&nbsp;</td>
																<td align='center'  class='time2'>Status</td>
																<td align='center'  class='time2'>Time</td>
																<td align='center'  class='time2'>Client</td>
																<td align='center'  class='time2'>Institution</td>
																<td align='center'  class='time2'>Language</td>
															
																	<td align='center'  class='time2'>Interpreter</td>
															
																<td align='center'  class='time2'>Phone</td>
																<td align='center'  class='time2'>Mobile</td>
																<td align='center'  class='time2'>&nbsp;</td>
															</tr>
														</thead>
														<tbody style="OVERFLOW: auto;">
															<tr bgcolor='#D4D0C8'>
																<td class='time' rowspan='<%=x12a + 1%>' >
																	12&nbsp;AM
																</td>
																<% If str12a = "" Then %>
																	<td colspan='9'>&nbsp;</td>
																<% End If %>
															</tr>
															<%=str12a%>
															<tr bgcolor='#D4D0C8'>
																<td class='time'  rowspan='<%=x1a + 1%>'>
																	1&nbsp;AM
																</td>
																<% If str1a = "" Then %>
																	<td colspan='9'>&nbsp;</td>
																<% End If %>
															</tr>
															<%=str1a%>
															<tr bgcolor='#D4D0C8'>
																<td class='time'  rowspan='<%=x2a + 1%>'>
																	2&nbsp;AM
																</td>
																<% If str2a = "" Then %>
																	<td colspan='9'>&nbsp;</td>
																<% End If %>
															</tr>
															<%=str2a%>
															<tr bgcolor='#D4D0C8'>
																<td class='time'  rowspan='<%=x3a + 1%>'>
																	3&nbsp;AM
																</td>
																<% If str2a = "" Then %>
																	<td colspan='9'>&nbsp;</td>
																<% End If %>
															</tr>
															<%=str3a%>
															<tr bgcolor='#D4D0C8'>
																<td class='time'  rowspan='<%=x4a + 1%>'>
																	4&nbsp;AM
																</td>
																<% If str4a = "" Then %>
																	<td colspan='9'>&nbsp;</td>
																<% End If %>
															</tr>
															<%=str4a%>
															<tr bgcolor='#D4D0C8'>
																<td class='time'  rowspan='<%=x5a + 1%>'>
																	5&nbsp;AM
																</td>
																<% If str5a = "" Then %>
																	<td colspan='9'>&nbsp;</td>
																<% End If %>
															</tr>
															<%=str5a%>
															<tr bgcolor='#D4D0C8'>
																<td class='time'  rowspan='<%=x6a + 1%>'>
																	6&nbsp;AM
																<% If str2a = "" Then %>
																	<td colspan='9'>&nbsp;</td>
																<% End If %>
															</tr>
															<%=str6a%>
															<tr bgcolor='#D4D0C8'>
																<td class='time'  rowspan='<%=x7a + 1%>'>
																	<a name='am7'>7&nbsp;AM</a>
																</td>
																<% If str7a = "" Then %>
																	<td colspan='9'>&nbsp;</td>
																<% End If %>
															</tr>
															<%=str7a%>
															<tr bgcolor='#D4D0C8'>
																<td class='time'  rowspan='<%=x8a + 1%>'>
																	8&nbsp;AM
																</td>
																<% If str8a = "" Then %>
																	<td colspan='9'>&nbsp;</td>
																<% End If %>
															</tr>
																	<%=str8a%>
															<tr bgcolor='#D4D0C8'>
																<td class='time'  rowspan='<%=x9a + 1%>'>
																	9&nbsp;AM
																</td>
																<% If str9a = "" Then %>
																	<td colspan='9'>&nbsp;</td>
																<% End If %>
															</tr>
															<%=str9a%>
															<tr bgcolor='#D4D0C8'>
																<td class='time'  rowspan='<%=x10a + 1%>'>
																	10&nbsp;AM
																</td>
																<% If str10a = "" Then %>
																	<td colspan='9'>&nbsp;</td>
																<% End If %>
															</tr>
															<%=str10a%>
															<tr bgcolor='#D4D0C8'>
																<td class='time'  rowspan='<%=x11a + 1%>'>
																	11&nbsp;AM
																</td>
																<% If str11a = "" Then %>
																	<td colspan='9'>&nbsp;</td>
																<% End If %>
															</tr>
															<%=str11a%>
															<tr bgcolor='#D4D0C8'>
																<td class='time'  rowspan='<%=x12p + 1%>'>
																	12&nbsp;PM
																</td>
																<% If str12p = "" Then %>
																	<td colspan='9'>&nbsp;</td>
																<% End If %>
															</tr>
															<%=str12p%>
															<tr bgcolor='#D4D0C8'>
																<td class='time' rowspan='<%=x1p + 1%>'>
																	1&nbsp;PM
																</td>
																<% If str1p = "" Then %>
																	<td colspan='9'>&nbsp;</td>
																<% End If %>
															</tr>
															<%=str1p%>
															</tr>
															<tr bgcolor='#D4D0C8'>
																<td class='time' rowspan='<%=x2p + 1%>'>
																	2&nbsp;PM
																</td>
																<% If str2p = "" Then %>
																	<td colspan='9'>&nbsp;</td>
																<% End If %>
															</tr>
															<%=str2p%>
															<tr bgcolor='#D4D0C8'>
																<td class='time' rowspan='<%=x3p + 1%>'>
																	3&nbsp;PM
																</td>
																<% If str3p = "" Then %>
																	<td colspan='9'>&nbsp;</td>
																<% End If %>
															</tr>
															<%=str3p%>
															<tr bgcolor='#D4D0C8'>
																<td class='time' rowspan='<%=x4p + 1%>'>
																	4&nbsp;PM
																</td>
																<% If str4p = "" Then %>
																	<td colspan='9'>&nbsp;</td>
																<% End If %>
															</tr>
															<%=str4p%>
															<tr bgcolor='#D4D0C8'>
																<td class='time' rowspan='<%=x5p + 1%>'>
																	5&nbsp;PM
																</td>
																<% If str5p = "" Then %>
																	<td colspan='9'>&nbsp;</td>
																<% End If %>
															</tr>
															<%=str5p%>
															<tr bgcolor='#D4D0C8'>
																<td class='time' rowspan='<%=x6p + 1%>'>
																	6&nbsp;PM
																</td>
																<% If str6p = "" Then %>
																	<td colspan='9'>&nbsp;</td>
																<% End If %>
															</tr>
															<%=str6p%>
															<tr bgcolor='#D4D0C8'>
																<td class='time' rowspan='<%=x7p + 1%>'>
																	7&nbsp;PM
																</td>
																<% If str7p = "" Then %>
																	<td colspan='9'>&nbsp;</td>
																<% End If %>
															</tr>
															<%=str7p%>
															<tr bgcolor='#D4D0C8'>
																<td class='time' rowspan='<%=x8p + 1%>'>
																	8&nbsp;PM
																</td>
																<% If str8p = "" Then %>
																	<td colspan='9'>&nbsp;</td>
																<% End If %>
															<%=str8p%>
															<tr bgcolor='#D4D0C8'>
																<td class='time' rowspan='<%=x9p + 1%>'>
																	9&nbsp;PM
																</td>
																<% If str9p = "" Then %>
																	<td colspan='9'>&nbsp;</td>
																<% End If %>
															<%=str9p%>
															<tr bgcolor='#D4D0C8'>
																<td class='time' rowspan='<%=x10p + 1%>'>
																	10&nbsp;PM
																</td>
																<% If str10p = "" Then %>
																	<td colspan='9'>&nbsp;</td>
																<% End If %>
															<%=str10p%>
															<tr bgcolor='#D4D0C8'>
																<td class='time' rowspan='<%=x11p + 1%>'>
																	11&nbsp;PM
																</td>
																<% If str11p = "" Then %>
																	<td colspan='9'>&nbsp;</td>
																<% End If %>
															<%=str11p%>
														</tbody>
													</table>
												</div>
											</td>
										</tr>
										<tr>
											<td align='right'>
												Legend: <font color='#000000' size='2'>•</font>&nbsp;-&nbsp;Completed&nbsp;<font color='#0000FF' size='2'>•</font>&nbsp;-&nbsp;Missed&nbsp;<font color='#FF0000 ' size='2'>•</font>&nbsp;-&nbsp;Canceled&nbsp;
													<font color='#FF00FF' size='2'>•</font>&nbsp;-&nbsp;Canceled (billable)
											</td>
										</tr>
									</table>
								</td>
								<td valign='top' style='width: 25%;'>	
									<table cellSpacing='2' cellPadding='2' style='width: 95%;' border='0' align='center'>
										<tr>
											<td colspan='7'>
												<table cellSpacing='0' cellPadding='0' align='center' style='width: 100%;' border='0'>
													<tr>
														<td align='left'>
															<input class='btn' type='button' value='&lt&lt' title='Previous Month' style='width: 25px;' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='ChangeMonth(0);'>
														</td>
														<td colspan='7' align='center' class='calheader'>
															<%=tmpMonth%>
														</td>
														<td align='right'>
															<input class='btn' type='button' value='&gt&gt' title='Next Month' style='width: 25px;' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='ChangeMonth(1);'>
														</td>
													</tr>
													<tr>
														<td>&nbsp;</td>
														<td class='calweekday'>Sun</td>
														<td class='calweekday'>Mon</td>
														<td class='calweekday'>Tue</td>
														<td class='calweekday'>Wed</td>
														<td class='calweekday'>Thu</td>
														<td class='calweekday'>Fri</td>
														<td class='calweekday'>Sat</td>
														<td>&nbsp;</td>
													</tr>
													<%=strCal%>
													<tr>
														<td colspan='9' align='center'>
															<table cellSpacing='0' cellPadding='0' style='width: 100%;' border='0' align='center'>
																<tr>
																	<td align='center'>
																		<nobr>Month:
																		<select class='seltxt' style='width: 50px;' name='selMonth'>
																			<option value='01'>Jan</option>
																			<option value='02'>Feb</option>
																			<option value='03'>Mar</option>
																			<option value='04'>Apr</option>
																			<option value='05'>May</option>
																			<option value='06'>Jun</option>
																			<option value='07'>Jul</option>
																			<option value='08'>Aug</option>
																			<option value='09'>Sep</option>
																			<option value='10'>Oct</option>
																			<option value='11'>Nov</option>
																			<option value='12'>Dec</option>
																		</select>
																		Year:
																		<input class='main' name='txtyear' maxlength='4' size='5'>
																		<input class='btn' type='button' value='GO' style='width: 25px;' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='SearchMonth();'>
																	</td>
																</tr>
																<tr><td>&nbsp;</td></tr>
																<tr><td><hr align='center' width='75%'></td></tr>
																<tr><td>&nbsp;</td></tr>
																<% If Session("type") = 1 Then %>
																	<tr>
																		<td align='center'>
																			<input class='btn' type='button' value='Print Schedule' style='width: 100px;' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='PrintToday();'>
																		</td>
																	</tr>
																<% ElseIf Session("type") <> 1 Then %>
																	<tr>
																		<td align='center'>
																			<input class='btn' type='button' value='Print Today' style='width: 100px;' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='PublishMe2();'>
																			&nbsp;&nbsp;
																			<input class='btn' type='button' value='Print Week' style='width: 100px;' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='PublishMe3();'>
																		</td>
																	</tr>
																<% End If %>
															</table>
														</td>
													</tr>
												</table>
											</td>
										</tr>
									</table>
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<input type='hidden' name='Hmonth' value='<%=tmpMonth%>'>
				<tr>
					<td height='50px' valign='bottom'>
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
<script><!--
	//toMiddle("cal");
--></script>