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
	sqlUInfo = "SELECT * FROM User_T WHERE index = " & Request("Use")
	rsUInfo.Open sqlUInfo, g_strCONN, 1, 3
	If Not rsUInfo.EOF Then
		MyInst = rsUInfo("InstID")
		MyRP = rsUInfo("ReqLB")
	End If
	rsUInfo.Close
	Set rsUInfo = Nothing
End If
'GET USERS
Set rsUser = Server.CreateObject("ADODB.RecordSet")
sqlUser = "SELECT * FROM User_T WHERE Type = 0 Or Type = 4 ORDER BY lname, Fname"
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
'GET REQUESTING PERSON
Set rsRP = Server.CreateObject("ADODB.RecordSet")
sqlRP = "SELECT * FROM requester_T ORDER BY lname, fname"
rsRP.Open sqlRP, g_strCONNLB, 3, 1
Do Until rsRP.EOF
	tmpRP = ""
	If MyRP = rsRP("index") Then tmpRP = "selected"
	tmpRPname = rsRP("lname") & ", " & rsRP("fname")
	strRP = strRP & "<option " & tmpRP & " value='" & rsRP("Index") & "'>" &  tmpRPname & "</option>" & vbCrlf
	rsRP.MoveNext
Loop
rsRP.Close
Set rsRP = Nothing
%>
<html>
	<head>
		<title>Interpreter Request - Admin Tools - INSTITUTION</title>
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
				document.frmUser.action = "a_inst.asp?use=" + xxx;
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
				alert("ERROR: Please select a User Account.")
				return;
			}
			if (document.frmUser.selInst.value == 0)
			{
				alert("ERROR: Please select an Institution.")
				return;
			}
			if (document.frmUser.selRP.value == 0)
			{
				alert("ERROR: Please select a Requesting Person.")
				return;
			}
			document.frmUser.action = "action.asp?ctrl=11";
			document.frmUser.submit();
		}
		 //-->
		</script>
	</head>
	<body>
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
									<select class='seltxt' name='selInst'  style='width:250px;' onchange='SelectInst(this.value);'>
										<option value='0'>&nbsp;</option>
										<%=strInst%>	
									</select>
								</td>
							</tr>
								<tr>
								<td align='right'>LanguageBank Requesting Person:</td>
								<td>
									<select class='seltxt' name='selRP'  style='width:250px;' onchange='SelectInst(this.value);'>
										<option value='0'>&nbsp;</option>
										<%=strRP%>	
									</select>
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