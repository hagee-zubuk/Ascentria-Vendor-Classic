<%
strAnn = ""
Set rsAnn = Server.CreateObject("ADODB.RecordSet")
sqlAnn = "SELECT HP FROM Announce_T"
rsAnn.Open SqlAnn, g_strCONNLB, 3, 1
If Not rsAnn.EOF Then
	strAnn = rsAnn("HP")
End If
rsAnn.Close
Set rsAnn = Nothing
%>