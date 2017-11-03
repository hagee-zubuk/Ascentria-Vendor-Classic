<%Language=VBScript%>
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<%Response.AddHeader "Pragma", "No-Cache" %>
<%
Function GetClass(xxx)
	Set rsClass = Server.CreateObject("ADODB.RecordSet")
	
	sqlClass = "SELECT [class] FROM Dept_T WHERE [InstID] = " & xxx

	rsClass.Open sqlClass, g_strCONNLB, 3, 1
	If Not rsClass.EOF Then
		GetClass = rsClass("Class")
	End If
	rsClass.CLose
	Set rsClass = Nothing
End function
Function GetClass2(xxx)
	Set rsClass = Server.CreateObject("ADODB.RecordSet")
	
	sqlClass = "SELECT [class] FROM Dept_T WHERE [index] = " & xxx

	rsClass.Open sqlClass, g_strCONNLB, 3, 1
	If Not rsClass.EOF Then
		GetClass2 = rsClass("Class")
	End If
	rsClass.CLose
	Set rsClass = Nothing
End function
ValidAko = False
Set rsUser = Server.CreateObject("ADODB.RecordSet")
sqlUser = "SELECT * FROM User_t WHERE upper([User]) = '" & UCase(Request("txtUN")) & "' "
rsUser.Open sqlUser, g_strCONN, 3, 1
response.write sqlUser
If Not rsUser.EOF Then
	If Request("txtPW") = Z_DoDecrypt(rsUser("pass")) Then 
		Session("type") = rsUser("type")
		Session("UID") =  rsUser("index")
		Session("GreetMe") = rsUser("lname")
		If rsUser("fname") <> "" Then Session("GreetMe") = Session("GreetMe") & ",  " & rsUser("fname")
		If rsUser("type") = 0 Or rsUser("type") = 3  Or rsUser("type") = 4 Or rsUser("type") = 5 Then 
			Session("DeptID") = rsUser("DeptLB")
			Session("ReqID") = rsUser("ReqLB")
			Session("InstID") = rsUser("InstID")
			If Z_CZero(rsUser("DeptLB")) > 0 Then
				Session("myClass") = GetClass2(rsUser("DeptLB"))
			Else
				Session("myClass") = GetClass(Session("InstID"))
			End If
		ElseIf rsUser("type") = 1 Then
			Session("IntrID") = rsUser("IntrLB")
		ElseIf rsUser("type") = 6 Then
			Session("myClass") = 3
			'create court list
			Set fso = Server.CreateObject("Scripting.FileSystemObject")
			Set oFile = fso.CreateTextFile(crtLst, True)
			
			Set rsCrt = Server.CreateObject("ADODB.RecordSet")
			sqlCrt = "SELECT DISTINCT(InstID) FROM dept_T WHERE Class = 3"
			rsCrt.Open sqlCrt, g_strCONNLB, 3, 1
			Do Until rsCrt.EOF
				oFile.WriteLine rsCrt("InstID")
				rsCrt.MoveNext
			Loop
			rsCrt.Close
			Set rsCrt = Nothing
			Set ofile = Nothing
			Set fso = Nothing
		End If
		ValidAko = True
	Else
		Session("MSG") = "ERROR: Invalid username and/or password."
	End If
Else
	Session("MSG") = "ERROR: Username and/or password invalid."
End If
rsUser.Close
Set rsUser = Nothing
%>
<!-- #include file="_closeSQL.asp" -->
<%
If ValidAko = True Then
	If Session("type") = 0 Or Session("type") = 4 Or Session("type") = 5 Or Session("type") = 6 Then
		If Session("UID") <> 36 Then
			Response.Redirect "calendarview2.asp"	
		Else
			Response.Redirect "main.asp"
		End If
	ElseIf Session("type") = 1 Then 
		Response.Redirect "calendarview2.asp"
	ElseIf Session("type") = 2 Then 
		Response.Redirect "admin.asp"
	ElseIf Session("type") = 3 Then 
		Response.Redirect "calendarview2.asp"
	End If
Else
	Response.Redirect "default.asp"
End If
%>