<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilCalendar.asp" -->
<%
Server.ScriptTimeout= 36000
If Session("type") <> 2 Then
	Session("MSG") = "ERROR: User type not allowed."
	Response.Redirect "default.asp"
End If
Function CleanMe(xxx)
	CleanMe = xxx
	If Not IsNull(xxx) Or xxx <> "" Then CleanMe = Replace(xxx, "'", " ")
End Function
MyInst = 0
MyDept = 0
MyRP = 0
If Request.ServerVariables("REQUEST_METHOD") = "POST" Or Request("Use") <> 0 Then
	Set rsUInfo = Server.CreateObject("ADODB.RecordSet")
	sqlUInfo = "SELECT * FROM User_T WHERE index = " & Request("Use")
	rsUInfo.Open sqlUInfo, g_strCONN, 1, 3
	If Not rsUInfo.EOF Then
		MyInst = rsUInfo("InstID")
		MyRP = rsUInfo("ReqLB")
		MyDept = rsUInfo("DeptLB")
	End If
	rsUInfo.Close
	Set rsUInfo = Nothing
	'GET DEPARTMENT
	Set rsInst = Server.CreateObject("ADODB.RecordSet")
	sqlInst = "SELECT * FROM dept_T WHERE InstID = " & MyInst & " ORDER BY dept"
	rsInst.Open sqlInst, g_strCONNLB, 3, 1
	Do Until rsInst.EOF
		tmpInst = ""
		If MyDept = rsInst("index") Then tmpInst = "selected"
		strDept = strDept & "<option " & tmpInst & " value='" & rsInst("Index") & "'>" &  rsInst("dept") & "</option>" & vbCrlf
		rsInst.MoveNext
	Loop
	rsInst.Close
	Set rsInst = Nothing
	'GET REQUESTING PERSON
	Set rsRP = Server.CreateObject("ADODB.RecordSet")
	sqlRP = "SELECT lname, fname, requester_T.[index] as myReqID FROM requester_T, reqdept_T WHERE reqID = requester_T.[index] AND DeptID = " & MyDept & " ORDER BY lname, fname"
	rsRP.Open sqlRP, g_strCONNLB, 3, 1
	Do Until rsRP.EOF
		tmpRP = ""
		If MyRP = rsRP("myReqID") Then tmpRP = "selected"
		tmpRPname = rsRP("lname") & ", " & rsRP("fname")
		strRP = strRP & "<option " & tmpRP & " value='" & rsRP("myReqID") & "'>" &  tmpRPname & "</option>" & vbCrlf
		rsRP.MoveNext
	Loop
	rsRP.Close
	Set rsRP = Nothing
End If
'GET USERS
Set rsUser = Server.CreateObject("ADODB.RecordSet")
sqlUser = "SELECT * FROM User_T WHERE Type = 3 ORDER BY lname, Fname"
rsUser.Open sqlUser,g_strCONN, 3, 1
ctrUser = 0
Do Until rsUser.EOF
	tmpUser = ""
	If Z_CZero(Request("Use")) = rsUser("index") Then tmpUser = "selected"
	UserName = rsUser("lname")
	If rsUser("fname") <> "" Then UserName = UserName & ", " & rsUser("fname")
	strUser = strUser & "<option " & tmpUser & " value='" & rsUser("Index") & "'>" &  UserName & "</option>" & vbCrlf
	rsUser.MoveNext
Loop
rsUser.Close
Set rsUser = Nothing
'GET INST
Set rsInst = Server.CreateObject("ADODB.RecordSet")
sqlInst = "SELECT * FROM institution_T ORDER BY Facility"
rsInst.Open sqlInst, g_strCONNLB, 3, 1
Do Until rsInst.EOF
	tmpInst = ""
	If MyInst = rsInst("index") Then tmpInst = "selected"
	strInst = strInst & "<option " & tmpInst & " value='" & rsInst("Index") & "'>" &  rsInst("Facility") & "</option>" & vbCrlf
	rsInst.MoveNext
Loop
rsInst.Close
Set rsInst = Nothing

'GET AVAILABLE DEPARTMENTS
Set rsInstDept = Server.CreateObject("ADODB.RecordSet")
sqlInstDept = "SELECT * FROM institution_T ORDER BY Facility"
rsInstDept.Open sqlInstDept, g_strCONNLB, 3, 1
Do Until rsInstDept.EOF
	InstDept = rsInstDept("Index")
	strInstDept = strInstDept & "if (inst == " & InstDept & "){" & vbCrLf
	Set rsDeptInst = Server.CreateObject("ADODB.RecordSet")
	sqlDeptInst = "SELECT * FROM dept_T WHERE InstID = " &  InstDept & " ORDER BY Dept"
	rsDeptInst.Open sqlDeptInst, g_strCONNLB, 3, 1
	If Not rsDeptInst.EOF Then
		Do Until rsDeptInst.EOF
			strInstDept = strInstDept & "if (dept != " & rsDeptInst("index") & ")" & vbCrLf & _
				"{var ChoiceInst = document.createElement('option');" & vbCrLf & _
				"ChoiceInst.value = " & rsDeptInst("index") & ";" & vbCrLf & _
				"ChoiceInst.appendChild(document.createTextNode(""" & rsDeptInst("Dept") & """));" & vbCrLf & _
				"document.frmUser.selDept.appendChild(ChoiceInst);} " & vbCrlf
			rsDeptInst.MoveNext
		Loop
	End If
	rsDeptInst.Close
	Set rsDeptInst = Nothing
	rsInstDept.MoveNext
	strInstDept = strInstDept & "}"
Loop
rsInstDept.Close
Set rsInstDept = Nothing
'GET DEPARTMENTS
Set rsDept2 = Server.CreateObject("ADODB.RecordSet")
sqlDept2 = "SELECT * FROM dept_T ORDER BY Dept"
rsDept2.Open sqlDept2, g_strCONNLB, 3, 1
Do Until rsDept2.EOF
	tmpDpt = ""
	If Z_Czero(myDept) = rsDept2("index") Then tmpDpt = "selected"
	DeptName = rsDept2("Dept")
	'If rsInst("Department") <> "" Then InstName = rsInst("Facility") & " - " & rsInst("Department")
	strDept2 = strDept2	& "<option " & tmpDpt & " value='" & rsDept2("Index") & "'>" &  DeptName & "</option>" & vbCrlf
	rsDept2.MoveNext
Loop
rsDept2.Close
Set rsDept2 = Nothing
'GET AVAILABLE REQUESTING PERSON PER DEPARTMENT
Set rsInstReq = Server.CreateObject("ADODB.RecordSet")
sqlInstReq = "SELECT * FROM dept_T ORDER BY dept"
rsInstReq.Open sqlInstReq, g_strCONNLB, 3, 1
Do Until rsInstReq.EOF
	InstReq = rsInstReq("index")
	strInstReqDept = strInstReqDept & "if (dept == " & InstReq & "){" & vbCrLf
	Set rsReqInst = Server.CreateObject("ADODB.RecordSet")
	sqlReqInst = "SELECT lname, fname, requester_T.[index] as myReqID FROM requester_T, reqdept_T WHERE  ReqID = requester_T.[index] AND DeptID = " & InstReq & " ORDER BY lname, fname"
	rsReqInst.Open sqlReqInst, g_strCONNLB, 3, 1
	Do Until rsReqInst.EOF
		tmpReqName = CleanMe(rsReqInst("lname")) & ", " & CleanMe(rsReqInst("fname"))
		strInstReqDept = strInstReqDept	& "if(req != "& rsReqInst("myReqID") & ")" & vbCrLf & _
			"{var ChoiceReq = document.createElement('option');" & vbCrLf & _
			"ChoiceReq.value = " & rsReqInst("myReqID") & ";" & vbCrLf & _
			"ChoiceReq.appendChild(document.createTextNode(""" & tmpReqName & """));" & vbCrLf & _
			"document.frmUser.selReq.appendChild(ChoiceReq);}" & vbCrLf
		rsReqInst.MoveNext
	Loop
	rsReqInst.Close
	Set rsReqInst = Nothing
	rsInstReq.MoveNext
	strInstReqDept = strInstReqDept & "}"
Loop
rsInstReq.Close
Set rsLangIntr = Nothing
'GET REQUESTING PERSON LIST
Set rsReq = Server.CreateObject("ADODB.RecordSet")
sqlReq = "SELECT * FROM requester_T ORDER BY Lname, Fname"
rsReq.Open sqlReq, g_strCONNLB, 3, 1
Do Until rsReq.EOF
	ReqSel = ""
	If myRP = "" Then tmpReqP = -1
	If Z_Czero(myRP) = rsReq("index") Then ReqSel = "selected"
	tmpReqName = CleanMe(rsReq("lname")) & ", " & CleanMe(rsReq("fname"))
	strReq2 = strReq2 & "<option " & ReqSel & " value='" & rsReq("Index") & "'>" & rsReq("Lname") & ", " & rsReq("Fname") & "</option>" & vbCrLf
	strReq = strReq & "{var ChoiceReq = document.createElement('option');" & vbCrLf & _
			"ChoiceReq.value = " & rsReq("index") & ";" & vbCrLf & _
			"ChoiceReq.appendChild(document.createTextNode(""" & tmpReqName & """));" & vbCrLf & _
			"document.frmUser.selReq.appendChild(ChoiceReq);}" & vbCrLf
	rsReq.MoveNext
Loop
rsReq.Close
Set rsReq = Nothing
%>
<html>
	<head>
		<title>Interpreter Request - Admin Tools - DEPARTMENT</title>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
		<!--
		function bawal(tmpform)
		{
			var iChars = ",|\"\'";
			var tmp = "";
			for (var i = 0; i < tmpform.value.length; i++)
			 {
			  	if (iChars.indexOf(tmpform.value.charAt(i)) != -1)
			  	{
			  		alert ("This character is not allowed.");
			  		tmpform.value = tmp;
			  		return;
		  		}
			  	else
		  		{
		  			tmp = tmp + tmpform.value.charAt(i);
		  		}
		  	}
		}
		function SelectUser(xxx)
		{
			if (xxx != 0)
			{
				document.frmUser.action = "a_dept.asp?use=" + xxx;
				document.frmUser.submit();
			}
			else
			{
				document.frmUser.selInst.value=0;
			}
		}
		function SelectInst(xxx)
		{
			if (xxx != 0)
			{
				document.frmUser.action = "a_dept.asp?use=" + xxx;
				document.frmUser.submit();
			}
			else
			{
				document.frmUser.selInst.value=0;
			}
		}
		function DeptChoice(inst, dept)
		{
			var i;
			for(i=document.frmUser.selDept.options.length-1;i>=1;i--)
			{
				if (dept != "undefined")
				{
					if (document.frmUser.selDept.options[i].value != dept)
					{
						document.frmUser.selDept.remove(i);
					}
				}
				else
				{
					document.frmUser.selReq.remove(i);
				}
			}
			<%=strInstDept%>
		}
		function ReqChoice(dept, req)
		{
			 var i;
			for(i=document.frmUser.selReq.options.length-1;i>=1;i--)
			{
				if (req != "undefined")
				{
					if (document.frmUser.selReq.options[i].value != req)
					{
						document.frmUser.selReq.remove(i);
					}
				}
				else
				{
					document.frmUser.selReq.remove(i);
				}
			}
			<%=strInstReqDept%>
		}
		function ReqShowMe()
		{
			if (document.frmUser.chkAll.checked == true) 
			{
				for(i=document.frmUser.selReq.options.length-1;i>=1;i--)
				{
					document.frmUser.selReq.remove(i);
				}
				<%=strReq%>
			}
			else
			{
				ReqChoice(document.frmUser.selDept.value);
			}
		}
		function SaveMe(xxx)
		{
			//check valid values
			if (document.frmUser.selUser.value == 0)
			{
				alert("ERROR: Please select a User Account.")
				return;
			}
			if (document.frmUser.selInst.value == 0)
			{
				alert("ERROR: Please select an Institution.")
				return;
			}
			if (document.frmUser.selDept.value == 0)
			{
				alert("ERROR: Please select a Department.")
				return;
			}
			if (document.frmUser.selReq.value == 0)
			{
				alert("ERROR: Please select a Requesting Person.")
				return;
			}
			document.frmUser.action = "action.asp?ctrl=15";
			document.frmUser.submit();
		}
		 //-->
		</script>
	</head>
	<body onload='DeptChoice(<%=MyInst%>, <%=MyDept%>); ReqChoice(<%=MyDept%>, <%=MyRP%>)'>
		<form method='post' name='frmUser'>
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
							<tr><td>&nbsp;</td></tr>
							<tr>
								<td align='center' colspan='10'><span class='error'><%=Session("MSG")%></span></td>
							</tr>
							<tr><td>&nbsp;</td></tr>
							<tr>
								<td align='right'>User Account:</td>
								<td>
									<select class='seltxt' name='selUser'  style='width:250px;' onchange='SelectUser(this.value);'>
										<option value='0'>&nbsp;</option>
										<%=strUser%>	
									</select>
								</td>
							</tr>
							<tr><td>&nbsp;</td></tr>
							<tr>
								<td class='header' align='right' width='300px'><nobr>INSTITUTION INFORMATION</td>
								<td><hr align='left' width='250px'></td>
							</tr>
							<tr>
								<td align='right'>LanguageBank Institution:</td>
								<td>
									<select class='seltxt' name='selInst'  style='width:250px;' onfocus='DeptChoice(document.frmUser.selInst.value);' onchange='DeptChoice(document.frmUser.selInst.value);'>
										<option value='0'>&nbsp;</option>
										<%=strInst%>	
									</select>
								</td>
							</tr>
							<tr>
								<td align='right'>Department:</td>
								<td>
									<select class='seltxt' name='selDept'  style='width:250px;' onfocus='ReqChoice(document.frmUser.selDept.value); '  onchange='ReqChoice(document.frmUser.selDept.value); '>
										<option value='0'>&nbsp;</option>
										<%=strDept2%>	
									</select>
								</td>
							</tr>
							<tr>
								<td align='right'>LanguageBank Requesting Person:</td>
								<td>
									<select class='seltxt' name='selReq'  style='width:250px;' onchange=''>
										<option value='0'>&nbsp;</option>
										<%=strReq2%>	
									</select>
									<input type='checkbox' name='chkAll' value='1' onclick='ReqShowMe();'>
									Show All
								</td>
							</tr>
							<tr><td>&nbsp;</td></tr>
							<tr><td>&nbsp;</td></tr>
							<tr><td>&nbsp;</td></tr>
							<tr><td>&nbsp;</td></tr>
							<tr><td>&nbsp;</td></tr>
							<tr><td>&nbsp;</td></tr>
							<tr><td>&nbsp;</td></tr>
							<tr><td>&nbsp;</td></tr>
							<tr><td>&nbsp;</td></tr>
							<tr><td>&nbsp;</td></tr>
							<tr><td>&nbsp;</td></tr>
							<tr><td>&nbsp;</td></tr>
							<tr>
								<td colspan='2' align='center' height='100px' valign='bottom'>
									<input class='btn' type='button' style='width: 125px;' value='Save' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="SaveMe(document.frmUser.selUser.value);">
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