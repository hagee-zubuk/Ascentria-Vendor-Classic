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

%>
<html>
	<head>
		<title>Interpreter Request - Admin Tools</title>
		<link href='style.css' type='text/css' rel='stylesheet'>
		
	</head>
	<body>
		<form method='post' name='frmAdmin' action=''>
			<table cellSpacing='0' cellPadding='0' height='100%' width="100%" border='0' class='bgstyle2'>
					<tr>
						<td valign='top'>
							<!-- #include file="_header.asp" -->
						</td>
					</tr>
					<tr>
						<td valign='top' >
							<table cellSpacing='2' cellPadding='0' width="100%"  border='0'>
								<!-- #include file="_greetme.asp" -->
								<tr>
									<td class='title' colspan='2' align='center'><nobr>Admin Tools</td>
								</tr>
								<tr>
									<td align='center' colspan='2'><span class='error'><%=Session("MSG")%></span></td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td>
										<table cellSpacing='2' cellPadding='0' align='center'  border='0'>
											<tr>
												<td align='center' style='background-COLOR: #336601;'>
													<a href='a_user.asp' class='Admin'>[User Accounts]</a>
												</td>
												<td>
													: Add/edit/delete User accounts
												</td>
											</tr>
											<tr><td>&nbsp;</td></tr>
											<tr>
												<td align='center' style='background-COLOR: #336601;'>
													<a href='a_inst.asp' class='Admin'>[Institution Accounts]</a>
												</td>
												<td>
													: Edit Institution accounts
												</td>
											</tr>
											<tr><td>&nbsp;</td></tr>
											<tr>
												<td align='center' style='background-COLOR: #336601;'>
													<a href='a_dept.asp' class='Admin'>[Department Accounts]</a>
												</td>
												<td>
													: Edit Department accounts
												</td>
											</tr>
											<tr><td>&nbsp;</td></tr>
											<tr>
												<td align='center' style='background-COLOR: #336601;'>
													<a href='a_intr.asp' class='Admin'>[Interpreter Accounts]</a>
												</td>
												<td>
													: Edit Interpreter accounts
												</td>
											</tr>
											<tr><td>&nbsp;</td></tr>
											<tr>
												<td align='center' style='background-COLOR: #336601;'>
													<a href='a_reason.asp' class='Admin'>[Reason]</a>
												</td>
												<td>
													: Add/edit/delete Reasons
												</td>
											</tr>
										</table>
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
