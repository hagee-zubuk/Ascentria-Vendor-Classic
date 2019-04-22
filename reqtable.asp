<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilCalendar.asp" -->
<%
Function Z_FormatTime(xxx)
	Z_FormatTime = Null
	If xxx <> "" Or Not IsNull(xxx)  Then
		If IsDate(xxx) Then Z_FormatTime = FormatDateTime(xxx, 4) 
	End If
End Function
server.scripttimeout = 360000
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

tmpPage = "document.frmTbl."
radioApp = ""
radioID = ""
radioAll = "checked"
radioAss = "checked"
radioUnass = ""
x = 0
If Request.ServerVariables("REQUEST_METHOD") = "POST"  Or Request("action") = 3 Then
	Call AddLog("FIND INITIATED... ")
	If Session("type") = 0 Or Session("type") = 4 Then
		sqlReq = "SELECT * FROM Appointment_T WHERE InstID = " & Session("InstID")
	Else 'If Session("type") = 3 Then
		sqlReq = "SELECT * FROM Appointment_T WHERE DeptID = " & Session("DeptID")
	End If
	'FIND
	If Request("radioStat") = 0 Then
		radioApp = "checked"
		radioID = ""
		radioAll = ""
		If Request("txtFromd8") <> "" Then
			If IsDate(Request("txtFromd8")) Then
				sqlReq = sqlReq & " AND appDate >= '" & Request("txtFromd8") & "' "
				tmpFromd8 = Request("txtFromd8") 
			Else
				Session("MSG") = "ERROR: Invalid Appointment Date Range (From)."
				Response.Redirect "reqtable.asp"
			End If
		End If
		If Request("txtTod8") <> "" Then
			If IsDate(Request("txtTod8")) Then
				sqlReq = sqlReq & " AND appDate <= '" & Request("txtTod8") & "' "
				tmpTod8 = Request("txtTod8")
			Else
				Session("MSG") = "ERROR: Invalid Appointment Date Range (To)."
				Response.Redirect "reqtable.asp"
			End If
		End If
	ElseIf Request("radioStat") = 1 Then
	
	End If
	xLang = Cint(Request("selLang"))
	If xLang <> -1 Then 
		sqlReq = sqlReq & " AND "
		sqlReq = sqlReq & "LangID = " & xLang
	End If
	
		'If Trim(Request("txtclilname")) <> "" Then
		'		sqlReq = sqlReq & " AND Upper(Clname) LIKE '" & Ucase(Z_DoEncrypt(Trim(Request("txtclilname")))) & "%'"
		'	End If
		'	If Trim(Request("txtclifname")) <> "" Then
		'		sqlReq = sqlReq & " AND Upper(Cfname) LIKE '" & Ucase(Z_DoEncrypt(Trim(Request("txtclifname")))) & "%'"
		'	End If


	'xIntr = Cint(Request("selIntr"))
	'If xIntr <> -1 Then 
	'	sqlReq = sqlReq & " AND "
	'	sqlReq = sqlReq & "IntrID = " & xIntr
	'End If
	
	
	'SORT
	

		sqlReq = sqlReq & " ORDER BY appdate"

'End If
'GET REQUESTS
Call AddLog("FIND: " & sqlReq)
Set rsReq = Server.CreateObject("ADODB.RecordSet")
'Response.Write sqlReq
rsReq.Open sqlReq, g_strCONN, 3, 1
x = 1
If Not rsReq.EOF Then
	Do Until rsReq.EOF
		includeme = True
		If Trim(Request("txtclilname")) <> "" Then
			If InStr(Ucase(Trim(Request("txtclilname"))), Ucase(Z_DoDecrypt(Trim(rsReq("Clname"))))) > 0 Then
				includeme = True
			Else
				includeme = False
			End If 
			'sqlReq = sqlReq & " AND Upper(Clname) LIKE '" & Ucase(Z_DoEncrypt(Trim(Request("txtclilname")))) & "%'"
		End If
		If Trim(Request("txtclifname")) <> "" Then
			If InStr(Ucase(Trim(Request("txtclifname"))), Ucase(Z_DoDecrypt(Trim(rsReq("Cfname"))))) > 0 Then
				includeme = True
			Else
				includeme = False
			End If 
			'sqlReq = sqlReq & " AND Upper(Cfname) LIKE '" & Ucase(Z_DoEncrypt(Trim(Request("txtclifname")))) & "%'"
		End If
		If includeme Then
			kulay = ""
			If Not Z_IsOdd(x) Then kulay = "#FBEEB7"
			'GET INSTITUTION
			
				tmpIname = GetInstNameLB(rsReq("InstID"))
				'If rsInst("Department") <> "" Then tmpIname = tmpIname & " <br> " & rsInst("Department")
			
			'GET INTERPRETER INFO
		
				tmpInName = GetIntrNameLB(rsReq("IntrID"))
		
			'GET LANGUAGE
			Set rsLang = Server.CreateObject("ADODB.RecordSet")
			sqlLang  = "SELECT [lang] FROM lang_T WHERE [index] = " & rsReq("LangID")
			rsLang.Open sqlLang , g_strCONN, 3, 1
			If Not rsLang.EOF Then
				tmpSalita = rsLang("lang") 
			Else
				tmpSalita = "N/A"
			End If
			rsLang.Close
			Set rsLang = Nothing 
			
			Stat = MyStatus(GetStatLB(rsReq("Index")))
			myDept =  GetDept(rsReq("DeptID"))
			
		
	
				strtbl = strtbl & "<tr bgcolor='" & kulay & "'>" & vbCrLf & _ 
					"<td class='tblgrn2' width='10px'>" & Stat & "</td>" & vbCrLf & _
					"<td class='tblgrn2' ><input type='hidden' name='ID" & x & "' value='" & rsReq("Index") & "'><a class='link2' href='reqconfirm.asp?ID=" & rsReq("Index") & "'><b>" & rsReq("Index") & "</b></a></td>" & vbCrLf & _
					"<td class='tblgrn2' ><nobr>" & tmpIname & myDept & "</td>" & vbCrLf & _
					"<td class='tblgrn2' >" & tmpSalita & "</td>" & vbCrLf & _
					"<td class='tblgrn2' >" & Z_DoDecrypt(rsReq("clname")) & ", " & Z_Dodecrypt(rsReq("cfname")) & "</td>" & vbCrLf & _
					"<td class='tblgrn2' >" & tmpInName & "</td>" & vbCrLf & _
					"<td class='tblgrn2' >" & rsReq("appDate") & "</td>" & vbCrLf & _
					"<td class='tblgrn2' ><nobr>" & Z_FormatTime(rsReq("TimeFrom")) & " - " & Z_FormatTime(rsReq("TimeTo")) & "</td>" & vbCrLf & _
					"<td class='tblgrn2' ><nobr>" & Z_FormatTime(rsReq("AStime")) & "</td>" & vbCrLf & _
					"<td class='tblgrn2' ><nobr>" & Z_FormatTime(rsReq("AEtime")) & "</td></tr>" & vbCrLf
				
			
	
			x = x + 1
		End If
		rsReq.MoveNext
	Loop
	if x = 1 Then strtbl = "<tr><td colspan='14' align='center'><i>&lt -- No records found. -- &gt</i></td></tr>"
Else
	strtbl = "<tr><td colspan='14' align='center'><i>&lt -- No records found. -- &gt</i></td></tr>"
End If
rsReq.Close
Set rsReq = Nothing
Call AddLog("FIND SUCCESS.")
End If
'SORT
If Request("sType") <> "" Then
	If Request("stype") = 1 Then stype = 2
	If Request("stype") = 2 Then stype = 1
Else
	stype = 1
End If
'FILTER CRITERIA
tmpclilname = Request("txtclilname")
tmpclifname = Request("txtclifname")

'GET AVAILABLE LANGUAGES
Set rsLang = Server.CreateObject("ADODB.RecordSet")
sqlLang = "SELECT [Index], [lang] FROM lang_T WHERE [index] <> 105 ORDER BY [lang]"
rsLang.Open sqlLang, g_strCONN, 3, 1
Do Until rsLang.EOF
	LangSel = ""
	If Cint(Request("selLang")) = rsLang("Index") Then LangSel = "selected"
	strLang = strLang	& "<option value='" & rsLang("Index") & "' " & LangSel & ">" &  rsLang("lang") & "</option>" & vbCrlf
	rsLang.MoveNext
Loop
rsLang.Close
Set rsLang = Nothing
'GET INTERPRETER LIST
'Set rsIntr = Server.CreateObject("ADODB.RecordSet")
'sqlIntr = "SELECT [Index], [last name], [first name] FROM interpreter_T WHERE Active = true ORDER BY [last name], [first name]"
'rsIntr.Open sqlIntr, g_strCONN, 3, 1
'Do Until rsIntr.EOF
'	IntrSel = ""
'	If Cint(Request("selIntr")) = rsIntr("Index") Then IntrSel = "selected"
'	strIntr = strIntr	& "<option value='" & rsIntr("Index") & "' " & IntrSel & ">" & rsIntr("last name") & ", " & rsIntr("first name") & "</option>" & vbCrlf
'	rsIntr.MoveNext
'Loop
'rsIntr.Close
'Set rsIntr = Nothing


%>
<html>
	<head>
		<title>Language Bank - Table Request</title>
		<link href="CalendarControl.css" type="text/css" rel="stylesheet">
		<script src="CalendarControl.js" language="javascript"></script>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
		<!--
		function SortMe(sortnum)
		{
			
		}
		function FindMe()
		{
			document.frmTbl.submit();
		}
		function FixSort()
		{
			document.frmTbl.txtFromd8.disabled = true;
			document.frmTbl.txtTod8.disabled = true;
		
			if (document.frmTbl.radioStat[0].checked == true)
			{
				document.frmTbl.txtFromd8.disabled = false;
				document.frmTbl.txtTod8.disabled = false;
			}
		}
		function CalendarView(strDate)
		{
			document.frmTbl.action = 'calendarview2.asp?appDate=' + strDate;
			document.frmTbl.submit();
		}
		function maskMe(str,textbox,loc,delim)
		{
			var locs = loc.split(',');
			for (var i = 0; i <= locs.length; i++)
			{
				for (var k = 0; k <= str.length; k++)
				{
					 if (k == locs[i])
					 {
						if (str.substring(k, k+1) != delim)
					 	{
					 		str = str.substring(0,k) + delim + str.substring(k,str.length);
		     			}
					}
				}
		 	}
			textbox.value = str
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
	          top:expression(this.offsetParent.scrollTop);
	      }
	      th
	      {
	          text-align: left;
	      }
		</style>
		<body onload='FixSort();'>
			<form method='post' name='frmTbl' action='reqtable.asp'>
				<table cellSpacing='0' cellPadding='0' height='100%' width="100%" border='0' class='bgstyle2'>
					<tr>
						<td height='100px'>
							<!-- #include file="_header.asp" -->
						</td>
					</tr>
					<tr>
						<td valign='top'>
							<table cellSpacing='2' cellPadding='0' width="100%" border='0'>
								<!-- #include file="_greetme.asp" -->
								<tr>
									<td>
										<table cellpadding='0' cellspacing='0' width='100%' border='0'>
											<tr>
												<td align='left'>
													Legend: <font color='#000000' size='+3'>•</font>&nbsp;-&nbsp;completed&nbsp;<font color='#0000FF' size='+3'>•</font>&nbsp;-&nbsp;missed&nbsp;<font color='#FF0000 ' size='+3'>•</font>&nbsp;-&nbsp;Canceled&nbsp;
													<font color='#FF00FF' size='+3'>•</font>&nbsp;-&nbsp;Canceled (billable)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
												</td>
												
													<td>&nbsp;</td>
											
											</tr>
										</table>
									</td>
								</tr>
								<% If Session("MSG") <> "" Then %>
									<tr><td>&nbsp;</td></tr>
									<tr>
										<td colspan='14' align='left'>
											<div name="dErr" style="width:300px; height:40px;OVERFLOW: auto;">
												<table border='0' cellspacing='1'>		
													<tr>
														<td><span class='error'><%=Session("MSG")%></span></td>
													</tr>
												</table>
											</div>
										</td>
									</tr>
									<tr><td>&nbsp;</td></tr>
								<% End If %>
								<tr>
									<td colspan='10' align='left'>
										<div class='container' style='height: 500px; width:1000px; position: relative;'>
											<table class="reqtble" width='100%'>	
												<thead>
													<tr class="noscroll">	
														<td colspan='2' class='tblgrn' onclick='SortMe(1);' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Request ID</td>
														<td class='tblgrn' onclick='SortMe(2);' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Institution</td>
														<td class='tblgrn' onclick='SortMe(3);' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Language</td>
														
															<td class='tblgrn' onclick='SortMe(4);' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Client</td>
											
														<td class='tblgrn' onclick='SortMe(5);' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Interpreter</td>
														<td class='tblgrn' onclick='SortMe(6);' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Appointment Date</td>
														<td class='tblgrn' onclick='SortMe(7);' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Planned Start and End Time</td>
														<td class='tblgrn' onclick='SortMe(8);' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Actual Start Time</td>
														<td class='tblgrn' onclick='SortMe(9);' onmouseover="this.className='tblgrnhover'" onmouseout="this.className='tblgrn'">Actual End Time</td>
													</tr>
												</thead>
												<tbody style="OVERFLOW: auto;">
													<%=strtbl%>
												</tbody>
											</table>
										</div>	
									</td>
								</tr>
							</table>
						</td>
					</tr>
					<tr>
						<td>
							<table width='100%'  border='0'>
								<tr>
									<td>&nbsp;</td>
									<td align='right'>
										<% If x <> 0 Then %>
											<b><u><%=x - 1%></u></b> records &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
										<% End If %>
									</td>
									<td>&nbsp;</td>
								</tr>
							</table>
						</td>
					</tr>
					<tr>
						<td>
							<table cellSpacing='0' cellPadding='0' width='1005px' border='0' style='border: solid 1px;'>
								<tr bgcolor='#FBEEB7'>
									<td align='right' style='border-bottom: solid 1px;'><b>Sort:</b></td>
									<td style='border-right: solid 1px;border-bottom: solid 1px;'>
										<input type='radio' name='radioStat' value='0' <%=radioApp%> onclick='FixSort();'>&nbsp;<b>App. Date Range:</b>
										&nbsp;&nbsp;
										<input class='main' size='10' maxlength='10' name='txtFromd8' value='<%=tmpFromd8%>'>
										&nbsp;-&nbsp;
										<input class='main' size='10' maxlength='10' name='txtTod8' value='<%=tmpTod8%>'>
										<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">mm/dd/yyyy</span>
										&nbsp;&nbsp;
										<input type='radio' name='radioStat' value='2' <%=radioAll%> onclick='FixSort();'>&nbsp;<b>All</b>
									</td>
									<td align='right' style='border-left: solid 1px;' rowspan='3'>
										<input class='btntbl' type='button' value='Find' style='height: 35px;' onmouseover="this.className='hovbtntbl'" onmouseout="this.className='btntbl'" onclick='FindMe();'>
									</td>
									</td>
								</tr>
								<tr bgcolor='#FBEEB7'>
									<td align='left' colspan='2'>
										
										&nbsp;Language:
										<select class='seltxt' style='width: 150px;' name='selLang'>
											<option value='-1'>&nbsp;</option>
											<%=strLang%>
										</select>

											&nbsp;Client:
											<input class='main' size='20' maxlength='20' name='txtclilname' value='<%=tmpclilname%>'>
											&nbsp;,&nbsp;&nbsp;
											<input class='main' size='20' maxlength='20' name='txtclifname' value='<%=tmpclifname%>'>
											<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">Last name, First name</span>

										
										&nbsp;
									</td>
								</tr>
								<!--<tr bgcolor='#FBEEB7'>
									<td align='left' colspan='4'>
										<!--Interpreter:
										<select class='seltxt' name='selIntr'>
											<option value='-1'>&nbsp;</option>
											<%=strIntr%>
										</select>
									
									</td>
									<td>&nbsp;</td>
								</tr>//-->
							</table>
						</td>
					</tr>
					<tr>
						<td height='50px' valign='bottom'>
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
	tmpMSG = Replace(Session("MSG"), "<br>", "\n")
%>
<script><!--
	alert("<%=tmpMSG%>");
--></script>
<%
End If
Session("MSG") = ""
%>