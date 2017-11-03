<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilCalendar.asp" -->
<%
Set rsReq = Server.CreateObject("ADODB.RecordSet")
sqlReq = "SELECT * FROM Appointment_T WHERE [index] = " & Request("ID")
rsReq.Open sqlReq, g_strCONN, 3, 1
If Not rsReq.EOF Then
		ts = rsReq("timestamp")
		If Session("ReqID") = 1083 Then 
			ReqName = rsReq("reqName")
		Else
			ReqName = rsReq("reqName") 'change
		End If
		ReqPhone = rsReq("rphone")
		Inst = GetInstNameLB(rsReq("InstID"))
		dept = GetDeptNameLB(rsReq("DeptID"))
		adr = GetDeptAdr(rsReq("deptID"))
		If rsReq("useCadr") Then 	adr = rsReq("caddress") & ", " & rsReq("capt") & ", " & rsReq("ccity") & ", " & rsReq("cstate") & ", " & rsReq("czip")
		dte = rsReq("appDate")
		tme = z_formattime(rsReq("TimeFrom")) & " - " & z_formattime(rsReq("TimeTo"))	
		cli = Z_DoDecrypt(rsReq("clname")) & ", " & Z_DoDecrypt(rsReq("cfname"))
		tmpLang = GetLang(rsReq("LangID")) 
		if rsReq("oLang") <> "" Then tmpLang = tmpLang & " (" & rsReq("oLang") & ")"
		phne = Z_DoDecrypt(rsReq("Phone"))
		tmpReas = GetReas(Z_Replace(rsReq("reason"),", ", "|"))
		tmpstaff = rsReq("clinician")
		tmpIntr = GetIntrNameLB(rsReq("IntrID"))
		docnum = rsReq("docknum")
		crtnum = rsReq("crtroom")
		chrg = Z_Replace(rsReq("charges"), vbCrLf, "<br>")
End If
rsReq.Close
Set rsReq = Nothing
%>
<html>
	<head>
		<title>Interpreter Request - Information</title>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
		<!--
		
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
								<% If Session("myClass") = 3 Then %>	
									<tr>
										<td  align='right'>Timestamp:</td>
										<td><b><%=ts%></b></td>
									</tr>
								<% End If %>
								<tr>
									<td  align='right'>Requesting Person:</td>
									<td><b><%=ReqName%></b></td>
								</tr>	
								<tr>
									<td  align='right'>Phone Number:</td>
									<td><b><%=ReqPhone%></b></td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td  align='right'>Institution:</td>
									<td><b><%=Inst%></b></td>
								</tr>	
								<tr>
									<td  align='right'>Department:</td>
									<td><b><%=Dept%></b></td>
								</tr>
								<tr>
									<td  align='right'>Address of Appointment:</td>
									<td><b><%=Adr%></b></td>
								</tr>
								<tr>
									<td  align='right'>Date:</td>
									<td><b><%=dte%></b></td>
								</tr>
								<tr>
									<td  align='right'>Time:</td>
									<td><b><%=tme%></b></td>
								</tr>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td  align='right'>Client:</td>
									<td><b><%=cli%></b></td>
								</tr>
								<tr>
									<td  align='right'>Language:</td>
									<td><b><%=tmpLang%></b></td>
								</tr>
								<tr>
									<td  align='right'>Phone Number:</td>
									<td><b><%=phne%></b></td>
								</tr>
								<% If Session("myClass") <> 3 Then %>	
									<tr>
										<td  align='right'>Reason/s:</td>
										<td><b><%=tmpReas%></b></td>
									</tr>
								<% Else %>
									<tr>
										<td  align='right'>Court Room No.:</td>
										<td><b><%=crtnum%></b></td>
									</tr>
									<tr>
										<td  align='right'>Docket Number:</td>
										<td><b><%=docnum%></b></td>
									</tr>
									<tr>
										<td  align='right' valign='top'>Charges:</td>
										<td><b><%=chrg%></b></td>
									</tr>
								<% End If %>
								<tr><td>&nbsp;</td></tr>
								<% If Session("myClass") <> 3 Then %>	
									<tr>
										<td  align='right'>DHHS assigned staff:</td>
										<td><b><%=tmpstaff%></b></td>
									</tr>
								<% End If %>
								<tr>
									<td  align='right'>Interpreter:</td>
									<td><b><%=tmpIntr%></b></td>
								</tr>
							</table>
						</tr>
						<tr><td>&nbsp;</td></tr>
						
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
