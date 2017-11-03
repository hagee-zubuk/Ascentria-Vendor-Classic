<%
If Session("UID") = "" Then 
	Session("MSG") = "Please Sign-in first/again."	
	Response.redirect "default.asp"
End If
%>