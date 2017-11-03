<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilCalendar.asp" -->
<%
If Session("type") <> 2 Then
	Session("MSG") = "ERROR: User type not allowed."
	Response.Redirect "default.asp"
End If
If Request.ServerVariables("REQUEST_METHOD") = "POST" Or Request("Use") <> 0 Then
	Set rsUInfo = Server.CreateObject("ADODB.RecordSet")
	sqlUInfo = "SELECT * FROM dept_T WHERE instID = " & Request("Use") & " ORDER BY dept"
	rsUInfo.Open sqlUInfo, g_strCONNLB, 1, 3
	Do Until rsUInfo.EOF
		tmpDept = ""
		If Z_CZero(Request("dept")) = rsUInfo("index") Then tmpDept = "selected"
		strInst = strInst & "<option " & tmpDept & " value='" & rsUInfo("index") & "'>" & rsUInfo("dept") & "</option>" & vbCrLf
		rsUInfo.MoveNext
	Loop
	rsUInfo.Close
	Set rsUInfo = Nothing
End If
'GET INST
Set rsUser = Server.CreateObject("ADODB.RecordSet")
sqlUser = "SELECT DISTINCT instID FROM User_T WHERE Type = 0 OR Type = 3 OR Type = 4"
rsUser.Open sqlUser,g_strCONN, 3, 1
ctrUser = 0
Do Until rsUser.EOF
	tmpUser = ""
	If Z_CZero(Request("Use")) = rsUser("instID") Then tmpUser = "selected"
	UserName = GetFacility(rsUser("instID"))
	'If rsUser("fname") <> "" Then UserName = UserName & ", " & rsUser("fname")
	If rsUser("instID") <> 0 Then strUser = strUser & "<option " & tmpUser & " value='" & rsUser("instID") & "'>" &  UserName & "</option>" & vbCrlf
	rsUser.MoveNext
Loop
rsUser.Close
Set rsUser = Nothing
'GET REASONS
Set rsReas = Server.CreateObject("ADODB.RecordSet")
sqlReas = "SELECT * FROM Reason_T  ORDER BY deptID, Reason"
rsReas.Open sqlReas, g_strCONN, 3, 1
Do Until rsReas.EOF
	tmpReas = rsReas("Reason")
	strDept3 = strDept3 & "if(dept == " & rsReas("deptID") & "){" & vbCrLf & _
		"var ChoiceRes = document.createElement('option');" & vbCrLf & _
		"ChoiceRes.value = " & rsReas("index") & ";" & vbCrLf & _
		"ChoiceRes.appendChild(document.createTextNode(""" & tmpReas & """));" & vbCrLf & _
		"document.frmUser.selReas.appendChild(ChoiceRes);}" & vbCrLf
		rsReas.MoveNext
Loop
rsReas.Close
Set rsReas = Nothing
%>
<html>
	<head>
		<title>Interpreter Request - Admin Tools - REASON</title>
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
				document.frmUser.action = "a_reason.asp?use=" + xxx;
				document.frmUser.submit();
			}
			else
			{
				document.frmUser.selInst.value=0;
			}
		}
		function SaveMe(xxx)
		{
			//check valid values
			if (document.frmUser.selUser.value == 0)
			{
				alert("ERROR: Please select an Institution.")
				return;
			}
			if (document.frmUser.selInst.value == 0)
			{
				alert("ERROR: Please select a Department.")
				return;
			}
			document.frmUser.action = "action.asp?ctrl=14";
			document.frmUser.submit();
		}
		function GetReas(dept)
		{
			document.frmUser.selReas.length = 0;
			if(dept == 0)
			{
				document.frmUser.selReas.length = 0;
			}
			<%=strDept3%>
		}
		 //-->
		</script>
	</head>
	<body onload='GetReas(document.frmUser.selInst.value);'>
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
								<td align='right'>Institution:</td>
								<td>
									<select class='seltxt' name='selUser'  style='width:250px;' onchange='SelectUser(this.value);'>
										<option value='0'>&nbsp;</option>
										<%=strUser%>	
									</select>
								</td>
							</tr>
							<tr><td>&nbsp;</td></tr>
							<tr>
								<td class='header' align='right' width='300px'><nobr>REASON INFORMATION</td>
								<td><hr align='left' width='250px'></td>
							</tr>
							<tr>
								<td align='right'>LanguageBank Departments:</td>
								<td>
									<select class='seltxt' name='selInst'  style='width:250px;' onchange='GetReas(this.value);'>
										<option value='0'>&nbsp;</option>
										<%=strInst%>	
									</select>
								</td>
							</tr>
							<tr>
								<td>&nbsp;</td>
								<td>
									<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">* Select Reason then click on save button to delete reason.</span>
								</td>
							</tr>
							<tr>
								<td align='right' valign='top'>Reasons:</td>
								<td>
									<select  name="selReas" class='seltxt' multiple style="height: 150px;">
									</select>
								</td>
							</tr>
							<tr>
								<td align='right' valign='top'>New Reason:</td>
								<td>
									<input class='main' size='50' maxlength='50' name='txtReas'>
								</td>
							</tr>
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
	tmpMSG = Replace(Session("MSG"), "<br>", "\n")
%>
<script><!--
	alert("<%=tmpMSG%>");
--></script>
<%
End If
Session("MSG") = ""
%>