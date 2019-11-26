<%
Function Z_IncludeDept(myid, deptid, ft)
	Z_IncludeDept = True
	if ft = TRUE Then Exit Function
	' I SHOULD ENCODE THIS TO A TABLE SOON! WTGDF! '
	If myid <> 509 And myid <> 510 And myid <> 517 And myid <> 534 And myid <> 625 And myid <> 626 And myid <> 627 And myid <> 628 Then Exit Function
	If myid = 509 Or myid = 510 Then
		Z_IncludeDept = False
		If deptid = 2306 Or deptid = 373 Or deptid = 946 Or deptid = 2302 Then Z_IncludeDept = True
	ElseIf myid = 517 Or myid = 534 Then
		Z_IncludeDept = False
		If deptid = 446 or deptid = 322 Or deptid = 289 Then Z_IncludeDept = True
	ElseIf myid = 625 Or myid = 626 Or myid = 627 Or myid = 628 Then
		Z_IncludeDept = False
		If deptid = 2466 or deptid = 2465 Then Z_IncludeDept = True
	End If
End Function

Function RemoveDept(xxx, ft)
	RemoveDept = False
	If ft = TRUE Then Exit Function
	strtmpDept = "," & xxx & ","
	deptID = ",209,811,886,1014,802,893,263,"
	If Instr(deptID, strtmpDept) > 0 Then RemoveDept = True
End Function

Function Z_FormatTime(xxx)
	Z_FormatTime = Null
	If xxx <> "" Or Not IsNull(xxx)  Then
		If IsDate(xxx) Then Z_FormatTime = FormatDateTime(xxx, 4) 
	End If
End Function

Function CleanMe(xxx)
	CleanMe = xxx
	If Not IsNull(xxx) Or xxx <> "" Then CleanMe = Replace(xxx, "'", "''")
End Function

Function Z_FormatTime(xxx)
	Z_FormatTime = Null
	If xxx <> "" Or Not IsNull(xxx)  Then
		If IsDate(xxx) Then Z_FormatTime = FormatDateTime(xxx, 4) 
	End If
End Function

Function ChkBilled(xxx)
	ChkBilled = False
	If xxx = "" Then Exit Function
	Set rsProc = Server.CreateObject("ADODB.RecordSet")
	sqlProc = "SELECT Processed, processedMedicaid FROM request_T WHERE HPID = " & xxx
	rsProc.Open sqlProc, g_strCONNLB, 3, 1
	If Not rsProc.EOF Then
		If Z_FixNull(rsProc("Processed")) <> "" Or Z_FixNull(rsProc("processedMedicaid")) <> "" Then ChkBilled = True
	End If 
	rsProc.Close
	Set rsProc = Nothing
End Function

Function GetStatLB(myid)
	GetStatLB = 0
	Set rsDept = Server.CreateObject("ADODB.RecordSet")
	sqlDept = "SELECT [status] FROM request_T WHERE [HPID] = " & myid & " ORDER BY [timestamp] DESC"
	rsDept.Open sqlDept, g_strCONNLB, 3, 1
	If Not rsDept.EOF Then
		GetStatLB = rsDept("status")
	End If
	rsDept.Close
	Set rsDept = Nothing
End Function
%>