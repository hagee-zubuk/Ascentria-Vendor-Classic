<%Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilCalendar.asp" -->
<%
'USER CHECK
If Cint(Session("type")) = 1 Then
	Session("MSG") = "Error: Invalid user type. Please sign-in again."
	Response.Redirect "default.asp"
End If
%>
<html>
	<head>
		<title>Interpreter Request - Reports</title>
		<link href="CalendarControl.css" type="text/css" rel="stylesheet">
		<script src="CalendarControl.js" language="javascript"></script>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
		<!--
		function RepGen()
		{
			if (document.frmReport.selRep.value == 0)
			{
				alert("NOTICE: Please select a report type.")
				return;
			}
			document.frmReport.action = "action.asp?ctrl=7";
			document.frmReport.submit();
		}
		function PopMe(xxx)
		{
			if (xxx != undefined)
			{ 
				newwindow = window.open('printreport.asp' ,'','height=800,width=900,scrollbars=1,directories=0,status=0,toolbar=0,resizable=1');
				if (window.focus) {newwindow.focus()}
			}
		}
		-->
		</script>
		<body onload='PopMe(<%=Request("Rtype")%>);'>
			<form method='post' name='frmReport' action='report.asp'>
				<table cellSpacing='0' cellPadding='0' height='100%' width="100%" class='bgstyle2' border='0'>
					<tr>
						<td height='100px'>
							<!-- #include file="_header.asp" -->
						</td>
					</tr>
					<!-- #include file="_greetme.asp" -->
					<tr><td>&nbsp;</td></tr>
					<tr>
						<td valign='top' >
							<table cellSpacing='4' cellPadding='0' align='center' border='0' class="defborder">
								<tr>
									<td colspan='2' align='center'>
										<b>Report Query</b>
									</td>
								</tr>
								<tr>
									<td align='right'>
										Type:
									</td>
									<td>
										<select class='seltxt' name='selRep'  style='width:200px;' onchange=''>
											<option value='0'>&nbsp;</option>
											<% If Session("type") <> 6 Then %>
												<!--<option value='1' <%=TypeSel1%>>Encounter Summary Report</option>//-->
											<% End If %>
											<% If Session("InstID") = 273 Then %>
												<option value='3' <%=TypeSel3%>>Appointment Summary Report</option>
												<option value='4' <%=TypeSel4%>>Appointment Summary Report (BLOCK)</option>
											<% End If %>
											<% If Session("InstID") = 27 Then %>
												<option value='2' <%=TypeSel2%>>DHMC Report</option>
											<% End If %>
											<% If Session("type") = 6 Then %>
												<option value='5' <%=TypeSel5%>>Court Appointment Cost</option>
												<option value='6' <%=TypeSel6%>>Court Language Cost</option>
												<option value='7' <%=TypeSel7%>>Court Language Frequency</option>
											<% End If %>
											<% If Session("type") <> 6 Then %>
												<option value='8' <%=TypeSel8%>>Activity Report</option>
												<option value='9' <%=TypeSel9%>>Language Use Report</option>
												<option value='10' <%=TypeSel10%>>Missed Appointment Report</option>
												<option value='11' <%=TypeSel11%>>Institution/Department Report</option>
											<% End If %>
										</select>
									</td>
								</tr>
								<tr><td colspan='2'><hr align='center' width='75%'></td></tr>
								<tr>
									<td align='right'>
										Criteria:
									</td>
									<td>
										( leave blank to select all )
									</td>
								</tr>
								<tr>
									<td align='right'>Timeframe:</td>
									<td>
										&nbsp;From:<input class='main' size='10' maxlength='10' name='txtRepFrom' readonly value='<%=tmpRepFrom%>'>
										<input type="button" value="..." title='Calendar' name="cal1" style="width: 19px;"
											onclick="showCalendarControl(document.frmReport.txtRepFrom);" class='btnLnk' onmouseover="this.className='hovbtnLnk'" onmouseout="this.className='btnLnk'">
										&nbsp;To:<input class='main' size='10' maxlength='10' name='txtRepTo' readonly value='<%=tmpRepTo%>'>
										<input type="button" value="..." title='Calendar' name="cal2" style="width: 19px;"
											onclick="showCalendarControl(document.frmReport.txtRepTo);" class='btnLnk' onmouseover="this.className='hovbtnLnk'" onmouseout="this.className='btnLnk'">
									</td>
								</tr>
								<tr><td colspan='2'><hr align='center' width='75%'></td></tr>
								<tr>
									<td>&nbsp;</td>
									<td>
										<input class='btn' type='button' style='width: 200px;' value='Generate' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='RepGen();'>
										<input type='hidden' name='hideID'>
									</td>
								</tr>
								<tr>
									<td colspan='2' align='center'>
										<span class='error'><%=Session("MSG")%></span>
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
	tmpMSG = Replace(Session("MSG"), "<br>", "\n")
%>
<script><!--
	alert("<%=tmpMSG%>");
--></script>
<%
End If
Session("MSG") = ""
%>