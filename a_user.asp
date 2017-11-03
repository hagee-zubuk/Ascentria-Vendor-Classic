<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
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
		tmplname = rsUInfo("lname")
		tmpfname = rsUInfo("fname")
		tmpUsername = rsUInfo("user")
		tmpPass = Z_DoDecrypt(rsUInfo("pass"))
		Inst = ""
		Intr = ""
		Admin = ""
		Dept = ""
		Inst2 = ""
		If rsUInfo("type") = 0 Then Inst = "selected"
		If rsUInfo("type") = 1 Then Intr = "selected"
		If rsUInfo("type") = 2 Then Admin = "selected" 
		If rsUInfo("type") = 3 Then Dept = "selected"
		If rsUInfo("type") = 4 Then Inst2 = "selected"
	End If
	rsUInfo.Close
	Set rsUInfo = Nothing
End If
'GET USERS
Set rsUser = Server.CreateObject("ADODB.RecordSet")
sqlUser = "SELECT * FROM User_T ORDER BY lname, Fname"
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
%>
<html>
	<head>
		<title>Interpreter Request - Admin Tools - USER</title>
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
				document.frmUser.action = "a_user.asp?Use=" + xxx;
				document.frmUser.submit();
			}
			else
			{
				document.frmUser.txtlname.value = "";
				document.frmUser.txtfname.value = "";
				document.frmUser.txtUser.value = "";
				document.frmUser.txtPass.value = "";
				document.frmUser.txtConPass.value = "";
				document.frmUser.selType.value = 0;
			}
		}
		function SaveMe(xxx)
		{
			//check valid values
			if (document.frmUser.txtlname.value == "")
			{
				alert("ERROR: User must have least a last name")
				return;
			}
			if (document.frmUser.txtUser.value == "")
			{
				alert("ERROR: User must have a Username")
				return;
			}
			if (document.frmUser.txtPass.value == "")
			{
				alert("ERROR: User must have a Password")
				return;
			}
			if (document.frmUser.txtPass.value != document.frmUser.txtConPass.value)
			{
				alert("ERROR: Password is different from Confirm Password.")
				return;
			}
			document.frmUser.action = "action.asp?ctrl=9";
			document.frmUser.submit();
		}
		function KillMe(xxx)
		{
			if (xxx != 0)
			{
				var ans = window.confirm("Delete user? Click Cancel to stop.")
				if (ans)
				{
					document.frmUser.action = "action.asp?ctrl=13";
					document.frmUser.submit();
				}
			}
			else
			{
				alert("ERROR: Please select a user.")
				return;
			}
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
								<td align='right' width='200px;'>User Account:</td>
								<td>
									<select class='seltxt' name='selUser'  style='width:250px;' onchange='SelectUser(this.value);'>
										<option value='0'>&nbsp;</option>
										<%=strUser%>	
									</select>
									<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">*Leave blank to add new user.</span>
								</td>
							</tr>
							<tr><td>&nbsp;</td></tr>
							<tr>
								<td class='header' align='right'><nobr>USER INFORMATION</td>
								<td><hr align='left' width='250px'></td>
							</tr>
							<tr>
								<td align='right'>Last Name:</td>
								<td>
									<input class='main' size='20' maxlength='20' name='txtlname' value='<%=tmplname%>' onkeyup='bawal(this);'>&nbsp;First Name:
									<input class='main' size='20' maxlength='20' name='txtfname' value='<%=tmpfname%>' onkeyup='bawal(this);'>
								</td>
							</tr>
							<tr>
								<td align='right'>Username:</td>
								<td>
									<input class='main' size='20' maxlength='20' name='txtUser' value='<%=tmpUsername%>' onkeyup='bawal(this);'>
								</td>
							</tr>
							<tr>
								<td align='right'>Password:</td>
								<td>
									<input type='password'  class='main' size='20' maxlength='20' name='txtPass' value='<%=tmpPass%>' onkeyup='bawal(this);'>
								</td>
							</tr>
							<tr>
								<td align='right'>Confirm Password:</td>
								<td>
									<input type='password'  class='main' size='20' maxlength='20' name='txtConPass' value='<%=tmpPass%>' onkeyup='bawal(this);'>
								</td>
							</tr>
							<tr>
								<td align='right'>Type:</td>
								<td>
									<select class='seltxt' name='selType'  style='width:120px;'>
										<option <%=Inst%> value='0'>Institution ONLY</option>
										<option <%=Dept%> value='3'>Department ONLY</option>	
										<option <%=Inst2%> value='4'>Special Institution</option>
										<option <%=Intr%> value='1'>Interpreter</option>
										<option <%=Admin%> value='2'>Administrator</option>
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
							<tr>
								<td colspan='2' align='center' height='100px' valign='bottom'>
									<input class='btn' type='button' style='width: 125px;' value='Save' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="SaveMe(document.frmUser.selUser.value);">
									<input class='btn' type='button' style='width: 125px;' value='Delete' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="KillMe(document.frmUser.selUser.value);">
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