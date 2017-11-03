<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_UtilCalendar.asp" -->
<!-- #include file="_Security.asp" -->
<%
	Set rsTbl = Server.CreateObject("ADODB.RecordSet")
	sqlTbl = "SELECT * FROM Appointment_T WHERE IntrID = 0 ORDER BY appDate, TimeFrom"
	rsTbl.Open sqlTbl, g_strCONN, 1, 3
	x = 1
	Do Until rsTbl.EOF
		kulay = ""
		If Not Z_IsOdd(x) Then kulay = "#FBEEB7"
		tmpClient = rsTbl("CLname") & ", " & rsTbl("CFname")
		tmpInst = GetInstName(rsTbl("InstID"))
		tmpLang = GetLang(rsTbl("LangID"))
		strtbl = strtbl & "<tr bgcolor='" & kulay & "'>" & vbCrLf & _ 
			"<td class='tblgrn2' ><input type='hidden' name='ID" & x & "' value='" & rsTbl("Index") & "'><a class='link2' href='assign.asp?ID=" & rsTbl("Index") & "'><b>" & rsTbl("Index") & "</b></a></td>" & vbCrLf & _
			"<td class='tblgrn2' ><nobr>" & tmpInst & "</td>" & vbCrLf & _
			"<td class='tblgrn2' >" & tmpLang & "</td>" & vbCrLf & _
			"<td class='tblgrn2' >" & tmpClient & "</td>" & vbCrLf & _
			"<td class='tblgrn2' >" & rsTbl("appDate") & "</td>" & vbCrLf & _
			"<td class='tblgrn2' ><nobr>" & Z_FormatTime(rsTbl("TimeFrom")) & "</td>" & vbCrLf & _
			"<td class='tblgrn2' ><nobr>" & Z_FormatTime(rsTbl("TimeTo")) & "</td></tr>" & vbCrLf
		x = x + 1
		rsTBL.MoveNext
	Loop
	rsTbl.Close
	Set rsTbl = Nothing
%>
<html>
	<head>
		<title>Interpreter Request - Assign Table</title>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
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
	</head>
	<body>
		<form method='post' name='frmTbl' action=''>
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
								<td colspan='10' align='left'>
									<div class='container' style='height: 500px; width:1000px; position: relative;'>
										<table class="reqtble" width='100%'>	
											<thead>
												<tr class="noscroll">	
													<td class='tblgrn' >Request ID</td>
													<td class='tblgrn' >Institution</td>
													<td class='tblgrn'>Language</td>
													<td class='tblgrn' >Client</td>
													<td class='tblgrn' >Appointment Date</td>
													<td class='tblgrn' >Start Time</td>
													<td class='tblgrn'>End Time</td>
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
					<td height='50px' valign='bottom'>
						<!-- #include file="_footer.asp" -->
					</td>
				</tr>
			</table>
		</form>
	</body>
</html>