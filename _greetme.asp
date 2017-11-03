<tr>
	<td colspan='2' align='center' class='greet'><nobr> --- Welcome&nbsp;&nbsp;<%=Session("GreetMe")%> ---</td>
</tr>
<% If strANN <> "" Then %>
<tr>
	<td colspan='14' align='center' class='greet2'><marquee scrollamount="3"><nobr> >>> ANNOUNCEMENT: <%=strANN%> <<<</marquee></td>
</tr>
<% End If %>