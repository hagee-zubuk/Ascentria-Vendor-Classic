<table cellSpacing='0' cellPadding='0' width="100%" border='0' align="center">
	<tr>
		<td valign='top' align="left" rowspan="2" width="75%" height="65px" colspan="10">
			<img src='images/LBISLOGO.jpg' border="0">
		</td>
		<td align="center" width="25%" class="tollnum">
		Toll-Free 844.579.0610
		</td>
	</tr>
	<tr>
		<td>&nbsp;</td>
	</tr>	
	<tr bgcolor='#f68328'>
		<td class="motto" align="center">
			Understand and Be Understood.
		</td>
		<% If Session("type") = 0 Or Session("type") = 4 Or Session("type") = 5 Then %>
			<td align='center' class='head' width='20px'>&nbsp;|&nbsp;</td> 
			<td align='center' width='100px'><a href='main.asp' class='linkv2'>New Request</a></td>
			<% If Session("UID") <> 36 Then %>
				<td align='center' class='head' width='20px'>&nbsp;|&nbsp;</td> 
				<td align='center' width='100px'><a href='reqtable.asp' class='linkv2'>List</a></td>
				<td align='center' class='head' width='20px'>&nbsp;|&nbsp;</td> 
				<td align='center' width='100px'><a href='calendarview2.asp' class='linkv2'>Calendar</a></td>
			<% End If %>
			<% If Session("UID") <> 36 Then %>
				<td align='center' class='head' width='20px'>&nbsp;|&nbsp;</td> 
	 			<td align='center' width='100px'><a href='report.asp' class='linkv2'>Reports</a></td>
	 		<% End If %>
 			<td align='center' class='head' width='20px'>&nbsp;|&nbsp;</td> 
		<% ElseIf Session("type") = 1 Then%>
			<td align='center' class='head' width='20px'>&nbsp;|&nbsp;</td> 
			<td align='center' width='100px'><a href='calendarview2.asp' class='linkv2'>Calendar</a></td>
			<td align='center' class='head' width='20px'>&nbsp;|&nbsp;</td> 
		<% ElseIf Session("type") = 2 Then%>
			<td align='center' class='head' width='20px'>&nbsp;|&nbsp;</td> 
				<td align='center' width='100px'><a href='report.asp' class='linkv2'>Reports</a></td>
   			<td align='center' class='head' width='20px'>&nbsp;|&nbsp;</td> 
			<td align='center' width='100px'><a href='admin.asp' class='linkv2'>Admin Tools</a></td>
			<td align='center' class='head' width='20px'>&nbsp;|&nbsp;</td> 
		<% ElseIf Session("type") = 3 Then%>
			<td align='center' class='head' width='20px'>&nbsp;|&nbsp;</td> 
			<td align='center' width='100px'><a href='main.asp' class='linkv2'>New Request</a></td>
			<td align='center' class='head' width='20px'>&nbsp;|&nbsp;</td> 
			<td align='center' width='100px'><a href='reqtable.asp' class='linkv2'>List</a></td>
			<td align='center' class='head' width='20px'>&nbsp;|&nbsp;</td> 
			<td align='center' width='100px'><a href='calendarview2.asp' class='linkv2'>Calendar</a></td>
 			<td align='center' class='head' width='20px'>&nbsp;|&nbsp;</td> 
 			<td align='center' width='100px'><a href='report.asp' class='linkv2'>Reports</a></td>
 			<td align='center' class='head' width='20px'>&nbsp;|&nbsp;</td> 
 		<% ElseIf Session("type") = 6 Then%>
 			<td align='center' class='head' width='20px'>&nbsp;|&nbsp;</td> 
			<td align='center' width='100px'><a href='calendarview2.asp' class='linkv2'>Calendar</a></td>
			<td align='center' class='head' width='20px'>&nbsp;|&nbsp;</td> 
			<td align='center' width='100px'><a href='report.asp' class='linkv2'>Reports</a></td>
		<% End If %>
		<td align='right'><a href='default.asp?chk=1' class='linkv2'>Sign Out</a>&nbsp;</td>
	</tr>
</table>