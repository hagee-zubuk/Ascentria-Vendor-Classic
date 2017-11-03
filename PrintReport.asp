<%Language=VBScript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilCalendar.asp" -->
<%
Function isBlock(xxx)
	isBlock = False
	Set rsblk = Server.CreateObject("ADODB.RecordSet")
	sqlblk = "SELECT block from appointment_T WHERE [index] = " & xxx
	rsblk.open sqlblk, g_strCONN, 3, 1
	If not rsblk.eof then
		'if trim(left(Ucase(rsblk("lbcom")), 15)) = "BLOCK SCHEDULE" Then
		'	isBlock = true
		'End if
		if rsblk("block") then isblock = true
	end if
	rsblk.close
	set rsblk = nothing
End Function
Function MyReas2(xxx)
	Set rsReas = Server.CreateObject("ADODB.RecordSet")
	sqlReas = "SELECT * FROM Encounter_T WHERE appID = " & xxx
	rsReas.Open sqlReas, g_strCONN, 3, 1
	Do Until rsReas.EOF 
		MyReas2 = MyReas2 & GetReason2(rsReas("reason"), rsReas("key")) & vbCrLf
		rsReas.MoveNext
	Loop
	rsReas.Close
	Set rsReas = Nothing
End Function
server.scripttimeout = 360000
tmpReport = Split(Z_DoDecrypt(Request.Cookies("HPREPORT")), "|")
If tmpReport(0) = 1 Then
	Set rsRep = Server.CreateObject("ADODB.RecordSet")
	strHead = "<tr><td class='tblgrn'>Reason</td>" & vbCrlf & _
		"<td class='tblgrn'>Count</td></tr>" & vbCrlf
	strMSG = "Encounter Summary "
	sqlRep = "SELECT Encounter_T.reason, Encounter_T.key, COUNT(Encounter_T.reason) AS CountMe FROM Appointment_T, Encounter_T WHERE [Appointment_T.index] = Encounter_T.appID"
	If tmpReport(1) <> "" Then
		sqlRep = sqlRep & " AND Appointment_T.appDate >= '" & tmpReport(1) & "'"
		strMSG = strMSG & " from " & tmpReport(1)
	End If
	If tmpReport(2) <> "" Then
		sqlRep = sqlRep & " AND Appointment_T.appDate <= '" & tmpReport(2) & "'"
		strMSG = strMSG & " to " & tmpReport(2)
	End If
	sqlRep = sqlRep & " AND InstID = " & Session("InstID") & " GROUP BY Encounter_T.key, Encounter_T.reason"
	strMSG = strMSG & " report for " & GetFacility(Session("InstID")) & "."
	rsRep.Open sqlRep, g_strCONN, 3, 1
	tmpKey = 0
	tmpReas = 0
	y = 0
	tmpCount = 0
	Do Until rsRep.EOF
		kulay = "#FFFFFF"
		If Not Z_IsOdd(y) Then kulay = "#F5F5F5"
		strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & GetReason(rsRep("key"), rsRep("reason")) & "</td>" & vbCrLf & _
			"<td class='tblgrn2'>" & rsRep("CountMe") & "</td></tr>" & vbCrLf
		y = y + 1
		 rsRep.MoveNext
	Loop
	rsRep.Close
	Set rsRep = Nothing
ElseIf tmpReport(0) = 2 Then
	ctr = 13
	repCSV = "DHMCreport.CSV"
	Set rsRep = Server.CreateObject("ADODB.RecordSet")
	'strHead = "<tr><td class='tblgrn' rowspan='2'>ID</td>" & vbCrlf & _
	'	"<td class='tblgrn' rowspan='2'>Appointment Date</td>" & vbCrlf & _
	'	"<td class='tblgrn' rowspan='2'>Time Paged</td>" & vbCrlf & _
	'	"<td class='tblgrn' colspan='2'>Requested</td>" & vbCrlf & _
	'	"<td class='tblgrn' colspan='2'>Actual</td>" & vbCrlf & _
	'	"<td class='tblgrn' rowspan='2'>15 mins. increments</td>" & vbCrlf & _
	'	"<td class='tblgrn' rowspan='2'>Department</td>" & vbCrlf & _
	'	"<td class='tblgrn' colspan='2'>Service Codes</td>" & vbCrlf & _
	'	"<td class='tblgrn' rowspan='2'>Language</td>" & vbCrlf & _
	'	"<td class='tblgrn' rowspan='2'>Confirmation Call (mins.)</td>" & vbCrlf & _
	'	"<td class='tblgrn' rowspan='2'>Follow up Call (mins.)</td>" & vbCrlf & _
	'	"<td class='tblgrn' rowspan='2'>Comments</td></tr>" & vbCrlf & _
	'	"<tr><td class='tblgrn'>Start Time</td>" & vbCrlf & _
	'	"<td class='tblgrn'>End Time</td>" & vbCrlf & _
	'	"<td class='tblgrn'>Start Time</td>" & vbCrlf & _
	'	"<td class='tblgrn'>End Time</td>" & vbCrlf & _
	'	"<td class='tblgrn'>Key</td>" & vbCrlf & _
	'	"<td class='tblgrn'>Reasons</td></tr>" & vbCrlf
	CSVHead = "ID,Appointment Date, Time Paged, Requested,, Actual,, 15 mins. increments, Department, Service Codes,, Language, Confirmation Call (mins.), Follow up Call (mins.), Comments"
	CSVHead2 = ",,,Start Time,End Time,Start Time, End Time,,,Key,Reasons"
	
	strMSG = "DHMC Report "
	sqlRep = "SELECT * FROM Appointment_T WHERE InstID = 27 AND LangID = 33 AND NOT isNull(AStime) AND NOT isNull(AEtime)"
	If tmpReport(1) <> "" Then
		sqlRep = sqlRep & " AND appDate >= '" & tmpReport(1) & "'"
		strMSG = strMSG & " from " & tmpReport(1)
	End If
	If tmpReport(2) <> "" Then
		sqlRep = sqlRep & " AND appDate <= '" & tmpReport(2) & "'"
		strMSG = strMSG & " to " & tmpReport(2)
	End If
	sqlRep = sqlRep & " ORDER by appDate"
	strMSG = strMSG & " report." 
	Call AddLog("REPORT: " & tmpReport(0) & " FIND: " & sqlRep)
	rsRep.Open sqlRep, g_strCONN, 3, 1
	y = 0
	If Not rsRep.EOF Then
		Do Until rsRep.EOF
			myKey = GetKey(rsRep("index"))
			myCom = GetCom(rsRep("index"))
			kulay = "#FFFFFF"
			MyReas = MyReas2(rsRep("index"))
			If Not Z_IsOdd(y) Then kulay = "#F5F5F5"
			strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'>" & rsRep("index") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'>" & rsRep("appDate") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'>" & rsRep("paged") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'>" & rsRep("TimeFrom") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'>" & rsRep("TimeTo") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'>" & rsRep("AStime") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'>" & rsRep("AEtime") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'>" & GetInc(rsRep("AStime"), rsRep("AEtime")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'>" & GetDept(rsRep("deptID")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'>" & myKey & "</td>" & vbCrLf & _
				"<td class='tblgrn2'>" & myReas & "</td>" & vbCrLf & _
				"<td class='tblgrn2'>" & GetLang(rsRep("LangID")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'>" & rsRep("Confirm") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'>" & rsRep("Follow") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'>" & myCom & "</td></tr>" & vbCrLf 
			y = y + 1
			strBody = "<tr><td align='center'>--- click on CSV export to view the report ---</td></tr>"
			CSVBody = CSVBody & rsRep("index") & "," & rsRep("appDate") & "," & rsRep("paged") & "," & rsRep("TimeFrom") & "," & rsRep("TimeTo") & "," & _
				rsRep("AStime") & "," & rsRep("AEtime") & "," & GetInc(rsRep("AStime"), rsRep("AEtime")) & "," & GetDept(rsRep("deptID")) & "," & myKey & _
			 	",""" & myReas & """," & GetLang(rsRep("LangID")) & "," & rsRep("Confirm") & "," & rsRep("Follow") & ",""" & myCom & """" & vbCrLf
			 rsRep.MoveNext
		Loop
	Else
		strBody = "<tr><td colspan='20' align='center'><i>--- No records found ---</i></td></tr>"
		strCSV = "--- No records found ---"
	End If
	rsRep.Close
	Set rsRep = Nothing
	Call AddLog("REPORT: " & tmpReport(0) & " SUCCESS.")
	'CONVERT TO CSV
	'Set fso = CreateObject("Scripting.FileSystemObject")
	'Set Prt = fso.CreateTextFile(BackupStr &  RepCSV, True)
	'Prt.WriteLine "LANGUAGE BANK - REPORT"
	'Prt.WriteLine strMSG
	'Prt.WriteLine CSVHead
	'Prt.WriteLine CSVHead2
	'Prt.WriteLine CSVBody
	'Prt.Close	
	'Set Prt = Nothing
	'Set fso = Nothing
	
	'DONWLOAD
	'tmpFile = BackupStr &  RepCSV
	'Set dload = Server.CreateObject("SCUpload.Upload")
	'dload.Download tmpFile
	'Set dload = Nothing
ElseIf tmpReport(0) = 3 Then
	RepCSV =  "PerInstReq" & tmpdate & ".csv" 
	Set rsRep = Server.CreateObject("ADODB.RecordSet")
	strHead = "<td class='tblgrn'>Appointment Date</td>" & vbCrlf & _
		"<td class='tblgrn'>Client's Name</td>" & vbCrlf & _
		"<td class='tblgrn'>Time of Appointment</td>" & vbCrlf & _
		"<td class='tblgrn'>Duration (mins)</td>" & vbCrlf & _
		"<td class='tblgrn'>Instituion - Department</td>" & vbCrlf & _
		"<td class='tblgrn'>Language</td>" & vbCrlf & _
		"<td class='tblgrn'>Interpreter's Name</td>" & vbCrlf & _
		"<td class='tblgrn'>Billed Hours</td>" & vbCrlf & _
		"<td class='tblgrn'>Total Amount</td>" & vbCrlf & _
		"<td class='tblgrn'>Travel Time</td>" & vbCrlf & _
		"<td class='tblgrn'>Mileage</td>" & vbCrlf
	CSVHead = "Appointment Date,Client's Last Name,Client's First Name,Actual Start Time,Actual End Time,Duration (mins),Instituion,Department," & _
		"Language,Interpreter's Last Name, Interpreter's First Name,Billed Hours,Total Amount,Travel Time,Mileage"
	sqlRep = "SELECT * FROM request_T, interpreter_T, institution_T, language_T, dept_T WHERE Dept_T.[index] = [DeptID] AND IntrID = interpreter_T.[index] " & _
		"AND request_T.InstID = institution_T.[index] AND LangID = language_T.[index] AND (request_T.Status = 1 OR request_T.Status = 4)"
	strMSG = "Per-institution request report"
	If tmpReport(1) <> "" Then
		sqlRep = sqlRep & " AND appDate >= '" & tmpReport(1) & "'"
		strMSG = strMSG & " from " & tmpReport(1)
	End If
	If tmpReport(2) <> "" Then
		sqlRep = sqlRep & " AND appDate <= '" & tmpReport(2) & "'"
		strMSG = strMSG & " to " & tmpReport(2)
	End If
	'If tmpReport(3) = "" Then tmpReport(3) = 0
	'If tmpReport(3) <> 0 Then
		sqlRep = sqlRep & " AND institution_T.[index] = " & Session("InstID")
		strMSG = strMSG & " for " & GetInstNameLB(Session("InstID"))
	'End If
	'If tmpReport(9) = "" Then tmpReport(9) = 0
	'If tmpReport(9) <> 0 Then
	'	If tmpReport(6) = "" Then tmpReport(6) = 0
	'	If tmpReport(6) <> 0 Then 
	'		sqlRep = sqlRep & " AND LangID = " & tmpReport(6)
	'	End If
	'	If tmpReport(7) = "" Then tmpReport(7) = 0
	'	If tmpReport(7) <> "0" Then
	'		tmpCli = Split(tmpReport(7), ",")
	'		sqlRep = sqlRep & " AND Clname = '" & Trim(tmpCli(0)) & "' AND Cfname = '" & Trim(tmpCli(1)) & "'"
	'	End If
	'	If tmpReport(8) = "" Then tmpReport(8) = 0
	'	If tmpReport(8) <> 0 Then 
	'		sqlRep = sqlRep & " AND Class = " & tmpReport(8)
	'	End If
	'End If
	sqlRep = sqlRep & " AND HPID <> '' and not Processed is null ORDER BY appDate, AStarttime, Facility, dept, Clname, Cfname"
	Call AddLog("REPORT: " & tmpReport(0) & " FIND: " & sqlRep)
	rsRep.Open sqlRep, g_strCONNLB, 3, 1
	If Not rsRep.EOF Then 
		x = 0
		Do Until rsRep.EOF 
			if not isBlock(rsRep("HPID")) Then
				kulay = "#FFFFFF"
				If Not Z_IsOdd(x) Then kulay = "#F5F5F5"
				tmpCliName = rsRep("Clname") & ", " & rsRep("Cfname")
				appTime = ctime(rsRep("AStarttime")) & " - " & ctime(rsRep("AEndtime"))
				appmin = DateDiff("n", rsRep("AStarttime"), rsRep("AEndtime"))
				tmpFacil = rsRep("Facility") & " - " & rsRep("Dept")
				tmpIName = rsRep("Last name") & ", " & rsRep("first name")
				tmpPay = (rsRep("billable") * rsRep("InstRate")) + Z_CZero(rsRep("TT_Inst")) + Z_CZero(rsRep("M_Inst"))
				strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & rsRep("appDate") & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & tmpCliName & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & appTime & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & appmin & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & tmpFacil & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & GetLangLB(rsRep("LangID")) & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & tmpIName & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & rsRep("billable") & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & Z_FormatNumber(tmpPay, 2) & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & Z_CZero(rsRep("TT_Inst")) & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & Z_CZero(rsRep("M_Inst")) & "</td></tr>" & vbCrLf
					
				CSVBody = CSVBody & rsRep("appDate") & ",""" & rsRep("Clname") & """,""" & rsRep("Cfname") & """," & rsRep("AStarttime") & _
					"," & rsRep("AEndtime") & "," & appmin & ",""" & rsRep("Facility") & """,""" & rsRep("Dept") & """," & GetLangLB(rsRep("LangID")) & _
					",""" & rsRep("Last name") & """,""" & rsRep("first name") & """," & rsRep("billable") & ",""" & Z_FormatNumber(tmpPay, 2) & _
					"""," & Z_CZero(rsRep("TT_Inst")) & "," & Z_CZero(rsRep("M_Inst")) & vbCrLf
				
				x = x + 1
			End If
			rsRep.MoveNext

		Loop
	Else
		strBody = "<tr><td colspan='11' align='center'><i>&lt --- No records found --- &gt</i></td></tr>"
		CSVBody = "< --- No records found --- >"
	End If
	rsRep.Close
	Set rsRep = Nothing
	Call AddLog("REPORT: " & tmpReport(0) & " SUCCESS.")
ElseIf tmpReport(0) = 4 Then
	RepCSV =  "PerInstReqBLOCK" & tmpdate & ".csv" 
	Set rsRep = Server.CreateObject("ADODB.RecordSet")
	strHead = "<td class='tblgrn'>Appointment Date</td>" & vbCrlf & _
		"<td class='tblgrn'>Client's Name</td>" & vbCrlf & _
		"<td class='tblgrn'>Time of Appointment</td>" & vbCrlf & _
		"<td class='tblgrn'>Duration (mins)</td>" & vbCrlf & _
		"<td class='tblgrn'>Instituion - Department</td>" & vbCrlf & _
		"<td class='tblgrn'>Language</td>" & vbCrlf & _
		"<td class='tblgrn'>Interpreter's Name</td>" & vbCrlf & _
		"<td class='tblgrn'>Billed Hours</td>" & vbCrlf & _
		"<td class='tblgrn'>Total Amount</td>" & vbCrlf & _
		"<td class='tblgrn'>Travel Time</td>" & vbCrlf & _
		"<td class='tblgrn'>Mileage</td>" & vbCrlf
	CSVHead = "Appointment Date,Client's Last Name,Client's First Name,Actual Start Time,Actual End Time,Duration (mins),Instituion,Department," & _
		"Language,Interpreter's Last Name, Interpreter's First Name,Billed Hours,Total Amount,Travel Time,Mileage"
	sqlRep = "SELECT * FROM request_T, interpreter_T, institution_T, language_T, dept_T WHERE Dept_T.[index] = [DeptID] AND IntrID = interpreter_T.[index] " & _
		"AND request_T.InstID = institution_T.[index] AND LangID = language_T.[index] AND (request_T.Status = 1 OR request_T.Status = 4)"
	strMSG = "Per-institution request report (BLOCK)"
	If tmpReport(1) <> "" Then
		sqlRep = sqlRep & " AND appDate >= '" & tmpReport(1) & "'"
		strMSG = strMSG & " from " & tmpReport(1)
	End If
	If tmpReport(2) <> "" Then
		sqlRep = sqlRep & " AND appDate <= '" & tmpReport(2) & "'"
		strMSG = strMSG & " to " & tmpReport(2)
	End If
	'If tmpReport(3) = "" Then tmpReport(3) = 0
	'If tmpReport(3) <> 0 Then
		sqlRep = sqlRep & " AND institution_T.[index] = " & Session("InstID")
		strMSG = strMSG & " for " & GetInstNameLB(Session("InstID"))
	'End If
	'If tmpReport(9) = "" Then tmpReport(9) = 0
	'If tmpReport(9) <> 0 Then
	'	If tmpReport(6) = "" Then tmpReport(6) = 0
	'	If tmpReport(6) <> 0 Then 
	'		sqlRep = sqlRep & " AND LangID = " & tmpReport(6)
	'	End If
	'	If tmpReport(7) = "" Then tmpReport(7) = 0
	'	If tmpReport(7) <> "0" Then
	'		tmpCli = Split(tmpReport(7), ",")
	'		sqlRep = sqlRep & " AND Clname = '" & Trim(tmpCli(0)) & "' AND Cfname = '" & Trim(tmpCli(1)) & "'"
	'	End If
	'	If tmpReport(8) = "" Then tmpReport(8) = 0
	'	If tmpReport(8) <> 0 Then 
	'		sqlRep = sqlRep & " AND Class = " & tmpReport(8)
	'	End If
	'End If
	sqlRep = sqlRep & " AND HPID <> '' and not Processed is null ORDER BY appDate, AStarttime, Facility, dept, Clname, Cfname"
	Call AddLog("REPORT: " & tmpReport(0) & " FIND: " & sqlRep)
	rsRep.Open sqlRep, g_strCONNLB, 3, 1
	If Not rsRep.EOF Then 
		x = 0
		Do Until rsRep.EOF 
			if isBlock(rsRep("HPID")) Then
				kulay = "#FFFFFF"
				If Not Z_IsOdd(x) Then kulay = "#F5F5F5"
				tmpCliName = rsRep("Clname") & ", " & rsRep("Cfname")
				appTime = ctime(rsRep("AStarttime")) & " - " & ctime(rsRep("AEndtime"))
				appmin = DateDiff("n", rsRep("AStarttime"), rsRep("AEndtime"))
				tmpFacil = rsRep("Facility") & " - " & rsRep("Dept")
				tmpIName = rsRep("Last name") & ", " & rsRep("first name")
				tmpPay = (rsRep("billable") * rsRep("InstRate")) + Z_CZero(rsRep("TT_Inst")) + Z_CZero(rsRep("M_Inst"))
				strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & rsRep("appDate") & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & tmpCliName & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & appTime & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & appmin & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & tmpFacil & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & GetLangLB(rsRep("LangID")) & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & tmpIName & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & rsRep("billable") & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & Z_FormatNumber(tmpPay, 2) & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & Z_CZero(rsRep("TT_Inst")) & "</td>" & vbCrLf & _
					"<td class='tblgrn2'><nobr>" & Z_CZero(rsRep("M_Inst")) & "</td></tr>" & vbCrLf
					
				CSVBody = CSVBody & rsRep("appDate") & ",""" & rsRep("Clname") & """,""" & rsRep("Cfname") & """," & rsRep("AStarttime") & _
					"," & rsRep("AEndtime") & "," & appmin & ",""" & rsRep("Facility") & """,""" & rsRep("Dept") & """," & GetLangLB(rsRep("LangID")) & _
					",""" & rsRep("Last name") & """,""" & rsRep("first name") & """," & rsRep("billable") & ",""" & Z_FormatNumber(tmpPay, 2) & _
					"""," & Z_CZero(rsRep("TT_Inst")) & "," & Z_CZero(rsRep("M_Inst")) & vbCrLf
				
				x = x + 1
			End If
			rsRep.MoveNext

		Loop
	Else
		strBody = "<tr><td colspan='11' align='center'><i>&lt --- No records found --- &gt</i></td></tr>"
		CSVBody = "< --- No records found --- >"
	End If
	rsRep.Close
	Set rsRep = Nothing
	Call AddLog("REPORT: " & tmpReport(0) & " SUCCESS.")
ElseIf tmpReport(0) = 5 Then
	RepCSV =  "CourtCost" & tmpdate & ".csv"
	strMSG = "Court Appointment cost report "
	strHead = "<td class='tblgrn'>Institution</td>" & vbCrlf & _
		"<td class='tblgrn'>Department</td>" & vbCrlf & _
		"<td class='tblgrn'>Appointment Date</td>" & vbCrlf & _
		"<td class='tblgrn'>Client</td>" & vbCrlf & _
		"<td class='tblgrn'>Language</td>" & vbCrlf & _
		"<td class='tblgrn'>Interpreter</td>" & vbCrlf & _
		"<td class='tblgrn'>Appointment Time</td>" & vbCrlf & _
		"<td class='tblgrn'>Hours</td>" & vbCrlf & _
		"<td class='tblgrn'>Rate</td>" & vbCrlf & _
		"<td class='tblgrn'>Travel Time</td>" & vbCrlf & _
		"<td class='tblgrn'>Mileage</td>" & vbCrlf & _
		"<td class='tblgrn'>Emergency Surcharge</td>" & vbCrlf & _
		"<td class='tblgrn'>Total</td>" & vbCrlf
	CSVHead = "Institution,Department,Appointment Date,Client Last Name,Client First Name,Language," & _
    "Interpreter Last Name,Interpreter First Name,Appointment Start Time,Appointment End Time,Hours," & _
    "Rate,Travel Time,Mileage,Emergency Surcharge,Total"
  Set rsRate = Server.CreateObject("ADODB.RecordSet")
	sqlRate = "SELECT * FROM EmergencyFee_T"
	rsRate.Open sqlRate, g_strCONNLB, 3, 1
	If Not rsRate.EOF Then
	    tmpFeeL = rsRate("FeeLegal")
	End If
	rsRate.Close
	Set rsRate = Nothing
	Set rsRep = Server.CreateObject("ADODB.RecordSet")
	sqlRep = "SELECT * FROM request_T, institution_T, dept_T, interpreter_T " & _
    "WHERE request_T.[instID] = institution_T.[index] AND deptID = dept_T.[index] AND " & _
    "intrID = interpreter_T.[index] AND class = 3 AND NOT processed IS NULL AND Billable > 0 AND (HPID > 0 OR NOT HPID IS NULL) "
   If tmpReport(1) <> "" Then
		sqlRep = sqlRep & "AND appDate >= '" & tmpReport(1) & "' "
		strMSG = strMSG & "from " & tmpReport(1)
	End If
	If tmpReport(2) <> "" Then
		sqlRep = sqlRep & "AND appDate <= '" & tmpReport(2) & "' "
		strMSG = strMSG & " to " & tmpReport(2)
	End If
	sqlRep = sqlRep & "ORDER BY facility, dept, appdate"
	Call AddLog("REPORT: " & tmpReport(0) & " FIND: " & sqlRep)
	rsRep.Open sqlRep, g_strCONNLB, 3, 1		
	InstID2 = 0
	tottotalPay = 0
	x = 0
	Do Until rsRep.EOF
		kulay = "#FFFFFF"
		If Not Z_IsOdd(x) Then kulay = "#F5F5F5"
    If InstID2 <> rsRep("InstID") And InstID2 <> 0 Then
    	strBody = strBody & "<tr><td colspan='12'>&nbsp;</td><td class='tblgrn2'><nobr>" & Z_formatNumber(tottotalPay, 2) & "</td></tr>" & vbCrLf
      CSVBody = CSVBody & ",,,,,,,,,,,,,,," & tottotalPay & vbCrLf
      tottotalPay = 0
    End If
    BillHours = rsRep("Billable")
    If rsRep("emerFEE") Then
        tmpPay = (BillHours * tmpFeeL) + Z_CZero(rsRep("TT_Inst")) + Z_CZero(rsRep("M_Inst"))
    Else
        tmpPay = (BillHours * rsRep("InstRate")) + Z_CZero(rsRep("TT_Inst")) + Z_CZero(rsRep("M_Inst"))
    End If
    totalPay = Z_FormatNumber(tmpPay, 2)
    strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & rsRep("facility") & "</td>" & vbCrLf & _
    	"<td class='tblgrn2'><nobr>" & rsRep("dept") & "</td>" & vbCrLf & _
			"<td class='tblgrn2'><nobr>" & rsRep("appdate") & "</td>" & vbCrLf & _
			"<td class='tblgrn2'><nobr>" & rsRep("clname") & ", " & rsRep("cfname") & "</td>" & vbCrLf & _
			"<td class='tblgrn2'><nobr>" & GetLangLB(rsRep("langID")) & "</td>" & vbCrLf & _
			"<td class='tblgrn2'><nobr>" & rsRep("Last Name") & ", " & rsRep("First Name") & "</td>" & vbCrLf & _
			"<td class='tblgrn2'><nobr>" & CTime(rsRep("AStarttime")) & " - " & CTime(rsRep("AEndtime"))  & "</td>" & vbCrLf & _
			"<td class='tblgrn2'><nobr>" & BillHours & "</td>" & vbCrLf
        
    CSVBody = CSVBody & """" & rsRep("facility") & """,""" & rsRep("dept") & """,""" & rsRep("appdate") & _
	    """,""" & rsRep("clname") & """,""" & rsRep("cfname") & """,""" & GetLangLB(rsRep("langID")) & _
	    """,""" & rsRep("Last Name") & """,""" & rsRep("First Name") & """,""" & _
	    CTime(rsRep("AStarttime")) & """,""" & CTime(rsRep("AEndtime")) & """,""" & _
	    BillHours & ""","""
	    
    If rsRep("emerFEE") Then
    	strBody = strBody & "<td class='tblgrn2'><nobr>" & tmpFeeL & "</td>" & vbCrLf
      CSVBody = CSVBody & tmpFeeL & ""","""
    Else
    	strBody = strBody & "<td class='tblgrn2'><nobr>" & rsRep("InstRate") & "</td>" & vbCrLf
      CSVBody = CSVBody & rsRep("InstRate") & ""","""
    End If
    strBody = strBody & "<td class='tblgrn2'><nobr>" & Z_CZero(rsRep("TT_Inst")) & "</td>" & vbCrLf & _
    	"<td class='tblgrn2'><nobr>" & Z_CZero(rsRep("M_Inst")) & "</td>" & vbCrLf & _
    	"<td class='tblgrn2'><nobr>0.00</td>" & vbCrLf & _
    	"<td class='tblgrn2'><nobr>" & totalPay & "</td></tr>" & vbCrLf	
    CSVBody = CSVBody & Z_CZero(rsRep("TT_Inst")) & """,""" & Z_CZero(rsRep("M_Inst")) & """,""" & _
        "0.00" & """,""" & totalPay & """" & vbCrLf
    tottotalPay = tottotalPay + tmpPay
    InstID2 = rsRep("InstID")
    x = x + 1
    rsRep.MoveNext
	Loop
	CSVBody = CSVBody & ",,,,,,,,,,,,,,," & tottotalPay
	rsRep.Close
	Set rsRep = Nothing
	Call AddLog("REPORT: " & tmpReport(0) & " SUCCESS.")
ElseIf tmpReport(0) = 6 Then
	RepCSV =  "CourtCostLang" & tmpdate & ".csv"
	strMSG = "Court Appointment cost report by Language "
	strHead = "<td class='tblgrn'>Institution</td>" & vbCrlf & _
		"<td class='tblgrn'>Department</td>" & vbCrlf & _
		"<td class='tblgrn'>Appointment Date</td>" & vbCrlf & _
		"<td class='tblgrn'>Client</td>" & vbCrlf & _
		"<td class='tblgrn'>Language</td>" & vbCrlf & _
		"<td class='tblgrn'>Interpreter</td>" & vbCrlf & _
		"<td class='tblgrn'>Appointment Time</td>" & vbCrlf & _
		"<td class='tblgrn'>Hours</td>" & vbCrlf & _
		"<td class='tblgrn'>Rate</td>" & vbCrlf & _
		"<td class='tblgrn'>Travel Time</td>" & vbCrlf & _
		"<td class='tblgrn'>Mileage</td>" & vbCrlf & _
		"<td class='tblgrn'>Emergency Surcharge</td>" & vbCrlf & _
		"<td class='tblgrn'>Total</td>" & vbCrlf
	CSVHead = "Institution,Department,Appointment Date,Client Last Name,Client First Name,Language," & _
    "Interpreter Last Name,Interpreter First Name,Appointment Start Time,Appointment End Time,Hours," & _
    "Rate,Travel Time,Mileage,Emergency Surcharge,Total"
  Set rsRate = Server.CreateObject("ADODB.RecordSet")
	sqlRate = "SELECT * FROM EmergencyFee_T"
	rsRate.Open sqlRate, g_strCONNLB, 3, 1
	If Not rsRate.EOF Then
	    tmpFeeL = rsRate("FeeLegal")
	End If
	rsRate.Close
	Set rsRate = Nothing
	Set rsRep = Server.CreateObject("ADODB.RecordSet")
	sqlRep = "SELECT * FROM request_T, institution_T, dept_T, interpreter_T, language_T " & _
    "WHERE request_T.[instID] = institution_T.[index] AND deptID = dept_T.[index] AND " & _
    "intrID = interpreter_T.[index] AND class = 3 AND NOT processed IS NULL AND Billable > 0 " & _
    "AND langID = language_T.[index] AND (HPID > 0 OR NOT HPID IS NULL) "
   If tmpReport(1) <> "" Then
		sqlRep = sqlRep & "AND appDate >= '" & tmpReport(1) & "' "
		strMSG = strMSG & "from " & tmpReport(1)
	End If
	If tmpReport(2) <> "" Then
		sqlRep = sqlRep & "AND appDate <= '" & tmpReport(2) & "' "
		strMSG = strMSG & " to " & tmpReport(2)
	End If
	sqlRep = sqlRep & "ORDER BY language, facility, dept, appdate"
	Call AddLog("REPORT: " & tmpReport(0) & " FIND: " & sqlRep)
	rsRep.Open sqlRep, g_strCONNLB, 3, 1		
	InstID2 = 0
	tottotalPay = 0
	x = 0
	Do Until rsRep.EOF
		kulay = "#FFFFFF"
		If Not Z_IsOdd(x) Then kulay = "#F5F5F5"
    If InstID2 <> rsRep("LangID") And InstID2 <> 0 Then
    	strBody = strBody & "<tr><td colspan='12'>&nbsp;</td><td class='tblgrn2'><nobr>" & Z_formatNumber(tottotalPay, 2) & "</td></tr>" & vbCrLf
      CSVBody = CSVBody & ",,,,,,,,,,,,,,," & tottotalPay & vbCrLf
      tottotalPay = 0
    End If
    BillHours = rsRep("Billable")
    If rsRep("emerFEE") Then
        tmpPay = (BillHours * tmpFeeL) + Z_CZero(rsRep("TT_Inst")) + Z_CZero(rsRep("M_Inst"))
    Else
        tmpPay = (BillHours * rsRep("InstRate")) + Z_CZero(rsRep("TT_Inst")) + Z_CZero(rsRep("M_Inst"))
    End If
    totalPay = Z_FormatNumber(tmpPay, 2)
    strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & rsRep("facility") & "</td>" & vbCrLf & _
    	"<td class='tblgrn2'><nobr>" & rsRep("dept") & "</td>" & vbCrLf & _
			"<td class='tblgrn2'><nobr>" & rsRep("appdate") & "</td>" & vbCrLf & _
			"<td class='tblgrn2'><nobr>" & rsRep("clname") & ", " & rsRep("cfname") & "</td>" & vbCrLf & _
			"<td class='tblgrn2'><nobr>" & GetLangLB(rsRep("langID")) & "</td>" & vbCrLf & _
			"<td class='tblgrn2'><nobr>" & rsRep("Last Name") & ", " & rsRep("First Name") & "</td>" & vbCrLf & _
			"<td class='tblgrn2'><nobr>" & CTime(rsRep("AStarttime")) & " - " & CTime(rsRep("AEndtime"))  & "</td>" & vbCrLf & _
			"<td class='tblgrn2'><nobr>" & BillHours & "</td>" & vbCrLf
        
    CSVBody = CSVBody & """" & rsRep("facility") & """,""" & rsRep("dept") & """,""" & rsRep("appdate") & _
	    """,""" & rsRep("clname") & """,""" & rsRep("cfname") & """,""" & GetLangLB(rsRep("langID")) & _
	    """,""" & rsRep("Last Name") & """,""" & rsRep("First Name") & """,""" & _
	    CTime(rsRep("AStarttime")) & """,""" & CTime(rsRep("AEndtime")) & """,""" & _
	    BillHours & ""","""
	    
    If rsRep("emerFEE") Then
    	strBody = strBody & "<td class='tblgrn2'><nobr>" & tmpFeeL & "</td>" & vbCrLf
      CSVBody = CSVBody & tmpFeeL & ""","""
    Else
    	strBody = strBody & "<td class='tblgrn2'><nobr>" & rsRep("InstRate") & "</td>" & vbCrLf
      CSVBody = CSVBody & rsRep("InstRate") & ""","""
    End If
    strBody = strBody & "<td class='tblgrn2'><nobr>" & Z_CZero(rsRep("TT_Inst")) & "</td>" & vbCrLf & _
    	"<td class='tblgrn2'><nobr>" & Z_CZero(rsRep("M_Inst")) & "</td>" & vbCrLf & _
    	"<td class='tblgrn2'><nobr>0.00</td>" & vbCrLf & _
    	"<td class='tblgrn2'><nobr>" & totalPay & "</td></tr>" & vbCrLf	
    CSVBody = CSVBody & Z_CZero(rsRep("TT_Inst")) & """,""" & Z_CZero(rsRep("M_Inst")) & """,""" & _
        "0.00" & """,""" & totalPay & """" & vbCrLf
    tottotalPay = tottotalPay + tmpPay
    InstID2 = rsRep("LangID")
    x = x + 1
    rsRep.MoveNext
	Loop
	CSVBody = CSVBody & ",,,,,,,,,,,,,,," & tottotalPay
	rsRep.Close
	Set rsRep = Nothing
	Call AddLog("REPORT: " & tmpReport(0) & " SUCCESS.")
ElseIf tmpReport(0) = 7 Then
	RepCSV =  "CourtCostFreq" & tmpdate & ".csv"
	strMSG = "Court Appointment by Language Frequency "
	Set rsRepA = Server.CreateObject("ADODB.RecordSet")
	strHead = "<td class='tblgrn'>Institution</td>" & vbCrlf & _
		"<td class='tblgrn'>Language</td>" & vbCrlf & _
		"<td class='tblgrn'>Count</td>" & vbCrlf
	CSVHead = "Institution,Language,Count"
	sqlRepA = "SELECT distinct(institution_T.facility), request_T.[instID] FROM request_T, institution_T, dept_T, interpreter_T " & _
	  "WHERE request_T.[instID] = institution_T.[index] AND deptID = dept_T.[index] AND " & _
	  "intrID = interpreter_T.[index] AND class = 3 AND NOT processed IS NULL AND Billable > 0 AND (HPID > 0 OR NOT HPID IS NULL) "
	If tmpReport(1) <> "" Then
		sqlRepA = sqlRepA & "AND appDate >= '" & tmpReport(1) & "' "
		strMSG = strMSG & "from " & tmpReport(1)
	End If
	If tmpReport(2) <> "" Then
		sqlRepA = sqlRepA & "AND appDate <= '" & tmpReport(2) & "' "
		strMSG = strMSG & " to " & tmpReport(2)
	End If
	sqlRepA = sqlRepA & "ORDER BY facility"
	Call AddLog("REPORT: " & tmpReport(0) & " FIND: " & sqlRepA)
	rsRepA.Open sqlRepA, g_strCONNLB, 3, 1
	Do Until rsRepA.EOF
		kulay = "#F5F5F5"
		strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & rsRepA("facility") & "</td></tr>" & vbCrLf
	  CSVBody = CSVBody & """" & rsRepA("facility") & """" & vbCrLf
	  Set rsRep = Server.CreateObject("ADODB.RecordSet")
    sqlrep = "SELECT distinct([language]) AS myLang, langID FROM request_T, dept_T, language_T WHERE " & _
	    "deptID = dept_T.[index] AND langID = language_T.[index] AND " & _
	    "request_T.[instID] = " & rsRepA("instID") & " AND class = 3 AND NOT processed IS NULL " & _
	    "AND Billable > 0 AND (HPID > 0 OR NOT HPID IS NULL) " 
	  If tmpReport(1) <> "" Then
			sqlRep = sqlRep & "AND appDate >= '" & tmpReport(1) & "' "
		End If
		If tmpReport(2) <> "" Then
			sqlRep = sqlRep & "AND appDate <= '" & tmpReport(2) & "' "
		End If
	  sqlrep = sqlrep & "order by [language]"
	 ' response.write sqlRep & "<br>"
    rsRep.Open sqlrep, g_strCONNLB, 3, 1
    Do Until rsRep.EOF
    	strBody = strBody & "<tr><td>&nbsp;</td><td class='tblgrn2'><nobr>" & rsRep("myLang") & "</td>" & vbCrLf
      CSVBody = CSVBody & """" & """,""" & rsRep("myLang") & ""","""
      Set rsNum = Server.CreateObject("ADODB.RecordSet")
      sqlNum = "SELECT COUNT(langID) AS myCount FROM request_T, Dept_T WHERE langID = " & rsRep("langID") & " AND " & _
	      "request_T.[instID] = " & rsRepA("instID") & " AND deptID = dept_T.[index] AND class = 3 AND " & _
	      "NOT processed IS NULL AND Billable > 0 AND (HPID > 0 OR NOT HPID IS NULL) "
	    If tmpReport(1) <> "" Then
				sqlNum = sqlNum & "AND appDate >= '" & tmpReport(1) & "' "
			End If
			If tmpReport(2) <> "" Then
				sqlNum = sqlNum & "AND appDate <= '" & tmpReport(2) & "' "
			End If
      rsNum.Open sqlNum, g_strCONNLB, 3, 1
      strBody = strBody & "<td class='tblgrn2'><nobr>" & rsNum("myCount") & "</td></tr>" & vbCrLf
      CSVBody = CSVBody & rsNum("myCount") & """" & vbCrLf
      rsNum.Close
      Set rsNum = Nothing
      rsRep.MoveNext
    Loop
    rsRep.Close
    Set rsRep = Nothing
    rsRepA.MoveNext
	Loop
	rsRepA.Close
	Set rsRepA = Nothing	
	Call AddLog("REPORT: " & tmpReport(0) & " SUCCESS.")
ElseIf tmpReport(0) = 8 Then 'Activity report
	RepCSV =  "Activity" & tmpdate & ".csv" 
	
	Set rsRep = Server.CreateObject("ADODB.RecordSet")
	strHead = "<td class='tblgrn'>Request ID</td>" & vbCrlf & _ 
		"<td class='tblgrn'>Department</td>" & vbCrlf & _
		"<td class='tblgrn'>Appointment Date</td>" & vbCrlf & _
		"<td class='tblgrn'>Client Name</td>" & vbCrlf & _
		"<td class='tblgrn'>Language</td>" & vbCrlf & _
		"<td class='tblgrn'>Interpreter Name</td>" & vbCrlf & _
		"<td class='tblgrn'>Appointment Time</td>" & vbCrlf & _
		"<td class='tblgrn'>Hours</td>" & vbCrlf & _
		"<td class='tblgrn'>Rate</td>" & vbCrlf & _
		"<td class='tblgrn'>Travel Time</td>" & vbCrlf & _
		"<td class='tblgrn'>Mileage</td>" & vbCrlf & _
		"<td class='tblgrn'>Emergency Surcharge</td>" & vbCrlf & _
		"<td class='tblgrn'>Total</td>" & vbCrlf & _
		"<td class='tblgrn'>STATUS</td>" & vbCrlf 
	
	CSVHead = "Request ID,Department, Appointment Date, Client Last Name, Client First Name, Language, Interpreter Last Name, Interpreter First Name, Appointment Start Time, " & _
		"Appointment End Time, Hours, Rate, Travel Time, Mileage, Emergency Surcharge, Total, STATUS"	
	
	strMSG = "All Activity Report"
	
	sqlRep = "SELECT request_T.[index] as myindex, status, [Last Name], [First Name], Clname, Cfname, AStarttime, AEndtime, " & _
		"Billable, emerFEE, class, TT_Inst, M_Inst, request_T.InstID as myinstID, DeptID, LangID, appDate, InstRate, bilComment, Processed, " & _
		"ProcessedMedicaid FROM request_T, interpreter_T , dept_T WHERE request_T.[instID] <> 479 AND request_T.deptID =  dept_T.[index] " & _
		"AND IntrID = interpreter_T.[index] AND (Status = 1 OR Status = 4 Or Status = 0) AND (HPID > 0 OR NOT HPID IS NULL) "

	If Session("type") = 0 Or Session("type") = 4 Or Session("type") = 5 Then
		sqlRep = sqlRep & " AND request_T.InstID = " & Session("InstID")
		strMSG = strMSG & " for " & GetInst2(Session("InstID"))
	ElseIf Session("type") = 3 Then
		sqlRep = sqlRep & " AND DeptID = " & Session("DeptID")
		strMSG = strMSG & " for " & GetInst2(Session("InstID")) & GetMyDept(Session("DeptID"))
	End if
	If tmpReport(1) <> "" Then
		sqlRep = sqlRep & " AND appDate >= '" & tmpReport(1) & "'"
		strMSG = strMSG & " from " & tmpReport(1)
	End If
	If tmpReport(2) <> "" Then
		sqlRep = sqlRep & " AND appDate <= '" & tmpReport(2) & "'"
		strMSG = strMSG & " to " & tmpReport(2)
	End If
	strMSG = strMSG & "."
	sqlRep = sqlRep & " ORDER BY AppDate DESC"
	Call AddLog("REPORT: " & tmpReport(0) & " FIND: " & sqlRep)
	rsRep.Open sqlRep, g_strCONNLB, 1, 3
	'EMERGENCY RATE
	Set rsRate = Server.CreateObject("ADODB.RecordSet")
	sqlRate = "SELECT * FROM EmergencyFee_T"
	rsRate.Open sqlRate, g_strCONNLB, 3,1
	If Not rsRate.EOF Then
		tmpFeeL = rsRate("FeeLegal")
		tmpFeeO = rsRate("FeeOther")
	End If
	rsRate.Close
	Set rsRate = Nothing
	If Not rsRep.EOF Then 
		x = 0
		Do Until rsRep.EOF
			kulay = "#FFFFFF"
			If Not Z_IsOdd(x) Then kulay = "#F5F5F5"
			strIntrName = rsRep("Last Name") & ",  " & rsRep("First Name")
			strCliName =  rsRep("Clname") & ", " & rsRep("Cfname")
			strATime =  cTime(rsRep("AStarttime")) & " -  " & cTime(rsRep("AEndtime"))
			'totHrs =  DateDiff("n", CDate(rsRep("AStarttime")) , CDate(rsRep("AEndtime")))
			BillHours =  rsRep("Billable")
			'totHrs2 = Z_FormatNumber( totHrs/60, 2)
			If rsRep("emerFEE") = True Then
				If rsRep("class") = 3 Or rsRep("class") = 5 Then
					tmpPay = (BillHours * tmpFeeL) + Z_CZero(rsRep("TT_Inst")) + Z_CZero(rsRep("M_Inst"))
				ElseIf rsRep("class") = 1 Or rsRep("class") = 2 Or rsRep("class") = 4 Then
					tmpPay = (BillHours * rsRep("InstRate")) + Z_CZero(rsRep("TT_Inst")) + Z_CZero(rsRep("M_Inst")) + tmpFeeO
				End If
			Else
				tmpPay = (BillHours * rsRep("InstRate")) + Z_CZero(rsRep("TT_Inst")) + Z_CZero(rsRep("M_Inst"))
			End If
			totalPay = Z_FormatNumber(tmpPay, 2)
			strBody = strBody & "<tr bgcolor='" & kulay & "' onclick='PassMe(" & rsRep("myindex") & ")'>" & _
				"<td class='tblgrn2'><nobr>" & rsRep("myindex") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & Replace(GetMyDept(rsRep("DeptID")), " - ", "") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("appDate") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & strCliName & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & GetLangLB(rsRep("LangID")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & strIntrName & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & strATime & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & BillHours & "</td>" & vbCrLf
				If rsRep("emerFEE") = True Then 
						If rsRep("class") = 3 Or rsRep("class") = 5 Then
							strBody = strBody & "<td class='tblgrn2'><nobr>$" & tmpFeeL & "</td>" & vbCrLf
						Else
							strBody = strBody & "<td class='tblgrn2'><nobr>$" & rsRep("InstRate") & "</td>" & vbCrLf
						End If
				Else
					strBody = strBody & "<td class='tblgrn2'><nobr>$" & rsRep("InstRate") & "</td>" & vbCrLf
				End If
				strBody = strBody & "<td class='tblgrn2'><nobr>$" & Z_CZero(rsRep("TT_Inst")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>$" & Z_CZero(rsRep("M_Inst")) & "</td>" & vbCrLf 
				If rsRep("emerFEE") = True Then 
					If rsRep("class") = 3 Or rsRep("class") = 5 Then
						strBody = strBody & "<td class='tblgrn2'><nobr>$0.00</td>" & vbCrLf
					ElseIf rsRep("class") = 1 Or rsRep("class") = 2 Or rsRep("class") = 4 Then
						strBody = strBody & "<td class='tblgrn2'><nobr>$" & tmpFeeO & "</td>" & vbCrLf
					End If
				Else
					strBody = strBody & "<td class='tblgrn2'><nobr>$0.00</td>" & vbCrLf
				End If
				strBody = strBody & "<td class='tblgrn2'><nobr><b>$" & totalPay & "</b></td>" & vbCrLf
				If IsNull(rsRep("Processed")) And IsNull(rsRep("ProcessedMedicaid")) Then
					strBody = strBody & "<td class='tblgrn2'><nobr>" & GetStat(rsRep("status")) & "</td><tr>" & vbCrLf 
				Else
					strBody = strBody & "<td class='tblgrn2'><nobr>Billed</td><tr>" & vbCrLf 
				End If
		
			CSVBody = CSVBody & """" & rsRep("myindex") & """,""" &  Replace(GetMyDept(rsRep("DeptID")), " - ", "") & """,""" & rsRep("appDate") & """,""" & rsRep("Clname") & """,""" & rsRep("Cfname") &  """,""" & GetLangLB(rsRep("LangID")) & """,""" & rsRep("Last Name") & _
				""",""" & rsRep("First Name") & ""","""  & cTime(rsRep("AStarttime")) & """,""" & cTime(rsRep("AEndtime")) & """,""" & BillHours
				
			If rsRep("emerFEE") = True Then 
				If rsRep("class") = 3 Or rsRep("class") = 5 Then
					CSVBody = CSVBody & """,""" & tmpFeeL
				Else
					CSVBody = CSVBody & """,""" & rsRep("InstRate")
				End If
			Else
				CSVBody = CSVBody & """,""" & rsRep("InstRate")
			end if
			
			CSVBody = CSVBody & """,""" & Z_CZero(rsRep("TT_Inst")) & """,""" & Z_CZero(rsRep("M_Inst")) & ""","""
			
			If rsRep("emerFEE") = True Then 
				If rsRep("class") = 3 Or rsRep("class") = 5 Then
					CSVBody = CSVBody & "0.00"
				ElseIf rsRep("class") = 1 Or rsRep("class") = 2 Or rsRep("class") = 4 Then
					CSVBody = CSVBody & tmpFeeO
				End If
			Else
				CSVBody = CSVBody & "0.00"
			end if
			
			CSVBody = CSVBody & """,""" & totalPay
			If IsNull(rsRep("Processed")) And IsNull(rsRep("ProcessedMedicaid")) Then
				CSVBody = CSVBody & """,""" & GetStat(rsRep("status")) & """" & vbCrLf 
			Else
				CSVBody = CSVBody & """,""" & "Billed" & """" & vbCrLf 
			End If
			x = x + 1
			rsRep.MoveNext
		Loop
	Else
		strBody = "<tr><td colspan='13' align='center'><i>&lt --- No records found --- &gt</i></td></tr>"
		CSVBody = "< --- No records found --- >"
	
	End If
	rsRep.Close
	Set rsRep = Nothing
	Call AddLog("REPORT: " & tmpReport(0) & " SUCCESS.")
ElseIf tmpReport(0) = 9 Then 'Language report
	RepCSV =  "LangFreq" & tmpdate & ".csv"
	strMSG = "Language Frequency report "
	Set rsRepA = Server.CreateObject("ADODB.RecordSet")
	strHead = "<td class='tblgrn'>Language</td>" & vbCrlf & _
		"<td class='tblgrn'>Count</td>" & vbCrlf
	CSVHead = "Language,Count"
	sqlRepA = "SELECT [language], COUNT([langid]) AS [f] " & _
		"FROM [request_t] AS r " & _
		"INNER JOIN [language_t] AS l ON r.[langID]=l.[index] " & _
		"WHERE (status = 1 OR status = 0) AND instID <> 479 " & _
		"AND (HPID > 0 OR NOT HPID IS NULL) "
	If tmpReport(1) <> "" Then
		sqlRepA = sqlRepA & " AND appDate >= '" & tmpReport(1) & "'"
		strMSG = strMSG & " from " & tmpReport(1)
	End If
	If tmpReport(2) <> "" Then
		sqlRepA = sqlRepA & " AND appDate <= '" & tmpReport(2) & "'"
		strMSG = strMSG & " to " & tmpReport(2)
	End If
	If Session("type") = 0 Or Session("type") = 4 Or Session("type") = 5 Then
		sqlRepA = sqlRepA & " AND InstID = " & Session("InstID")
		strMSG = strMSG & " for " & GetInst2(Session("InstID"))
	ElseIf Session("type") = 3 Then
		sqlRepA = sqlRepA & " AND DeptID = " & Session("DeptID")
		strMSG = strMSG & " for " & GetInst2(Session("InstID")) & GetMyDept(Session("DeptID"))
	End if
	sqlRepA = sqlRepA &	"GROUP BY [language] " & _
		"ORDER BY [language] ASC"
	Call AddLog("REPORT: " & tmpReport(0) & " FIND: " & sqlRepA)
	rsRepA.Open sqlRepA, g_strCONNLB, 3, 1
	y = 0
	Do Until rsRepA.EOF
		kulay = "#FFFFFF"
		If Not Z_IsOdd(y) Then kulay = "#F5F5F5"
    strBody = strBody & "<tr  bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & rsRepA("language") & "</td>" & vbCrLf
		CSVBody = CSVBody & """" & rsRepA("language") & ""","""
		strBody = strBody & "<td class='tblgrn2'><nobr>" & rsRepA("f") & "</td></tr>" & vbCrLf
		CSVBody = CSVBody & rsRepA("f") & """" & vbCrLf
		rsRepA.MoveNext
		y = y + 1
	Loop
	rsRepA.Close
	Set rsRepA = Nothing	
	Call AddLog("REPORT: " & tmpReport(0) & " SUCCESS.")
ElseIf tmpReport(0) = 10 Then 'missed
	RepCSV =  "Missed" & tmpdate & ".csv" 
	Set rsRep = Server.CreateObject("ADODB.RecordSet")
	sqlRep = "SELECT request_T.[index] as myindex, Facility, dept, intrID, DeptID, LangID, Clname, Cfname, missed, appDate, appTimeFrom, appTimeTo, Comment FROM request_T, institution_T, Dept_T WHERE request_T.[instID] <> 479 AND institution_T.[index] = request_T.InstID AND dept_T.[index] = DeptID " & _
		"AND Status = 2 AND (HPID > 0 OR NOT HPID IS NULL)"
	strMSG = "Missed appointment report"
	strHead = "<td class='tblgrn'>Request ID</td>" & vbCrlf & _
		"<td class='tblgrn'>Department</td>" & vbCrlf & _
		"<td class='tblgrn'>Language</td>" & vbCrlf & _
		"<td class='tblgrn'>Client</td>" & vbCrlf & _
		"<td class='tblgrn'>Interpreter</td>" & vbCrlf & _
		"<td class='tblgrn'>Appointment Date</td>" & vbCrlf & _
		"<td class='tblgrn'>Start and End Time</td>" & vbCrlf & _
		"<td class='tblgrn'>Comments</td>" & vbCrlf
	CSVHead = "Request ID, Department,Language, Client Last Name, Client First Name, Interpreter Last Name, Interpreter First Name, Appointment Date, Appointment Start Time, " & _
		"Appointment End Time, Comments"
	If Session("type") = 0 Or Session("type") = 4 Or Session("type") = 5 Then
		sqlRep = sqlRep & " AND request_T.InstID = " & Session("InstID")
		strMSG = strMSG & " for " & GetInst2(Session("InstID"))
	ElseIf Session("type") = 3 Then
		sqlRep = sqlRep & " AND DeptID = " & Session("DeptID")
		strMSG = strMSG & " for " & GetInst2(Session("InstID")) & GetMyDept(Session("DeptID"))
	End if
	If tmpReport(1) <> "" Then
		sqlRep = sqlRep & " AND appDate >= '" & tmpReport(1) & "'"
		strMSG = strMSG & " from " & tmpReport(1)
	End If
	If tmpReport(2) <> "" Then
		sqlRep = sqlRep & " AND appDate <= '" & tmpReport(2) & "'"
		strMSG = strMSG & " to " & tmpReport(2)
	End If
	sqlRep = sqlRep & " ORDER BY Facility, appDate, Clname, Cfname"
	Call AddLog("REPORT: " & tmpReport(0) & " FIND: " & sqlRep)
	rsRep.Open sqlRep, g_strCONNLB, 1, 3	
	If Not rsRep.EOF Then
		x = 0
		Do Until rsRep.EOF
			kulay = "#FFFFFF"
			If Not Z_IsOdd(x) Then kulay = "#F5F5F5"
			strBody = strBody & "<tr bgcolor='" & kulay & "' onclick='PassMe(" & rsRep("myindex") & ")'>" & _
				"<td class='tblgrn2'><nobr>" & rsRep("myindex") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("dept") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & GetLangLB(rsRep("LangID")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("Clname") & ", " & rsRep("Cfname") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & GetIntrNameLB(rsRep("intrID")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("appDate") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & ctime(rsRep("appTimeFrom")) & " - " & ctime(rsRep("appTimeTo")) &"</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("Comment") & "</td></tr>" & vbCrLf
			If GetIntrNameLB(rsRep("intrID")) = "N/A" Then 
				intrName = "N/A,"
			Else
				intrName = GetIntrNameLB(rsRep("intrID"))
			End If
			CSVBody = CSVBody & rsRep("myindex") & "," & Replace(GetMyDept(rsRep("DeptID")), " - ", "") & "," & GetLangLB(rsRep("LangID")) & "," & rsRep("Clname") & "," & rsRep("Cfname") &  ","  & _
				intrName & ","  & rsRep("appDate") & ","  & ctime(rsRep("appTimeFrom")) & "," & ctime(rsRep("appTimeTo")) & ",""" & _
				rsRep("Comment") &"""" &  vbCrLf
			rsRep.MoveNext
			x = x + 1
		Loop
	Else
		strBody = "<tr><td colspan='8' align='center'><i>&lt --- No records found --- &gt</i></td></tr>"
		CSVBody = "< --- No records found --- >"
	End If
	rsRep.Close
	Set rsRep = Nothing	
	Call AddLog("REPORT: " & tmpReport(0) & " SUCCESS.")
ElseIf tmpReport(0) = 11 Then 'insti report
	RepCSV =  "PerInstReq" & tmpdate & ".csv" 
	Set rsRep = Server.CreateObject("ADODB.RecordSet")
	strHead = "<td class='tblgrn'>Appointment Date</td>" & vbCrlf & _
		"<td class='tblgrn'>Client's Name</td>" & vbCrlf & _
		"<td class='tblgrn'>Time of Appointment</td>" & vbCrlf & _
		"<td class='tblgrn'>Duration (mins)</td>" & vbCrlf & _
		"<td class='tblgrn'>Department</td>" & vbCrlf & _
		"<td class='tblgrn'>Language</td>" & vbCrlf & _
		"<td class='tblgrn'>Interpreter's Name</td>" & vbCrlf & _
		"<td class='tblgrn'>Billed Hours</td>" & vbCrlf & _
		"<td class='tblgrn'>Total Amount</td>" & vbCrlf & _
		"<td class='tblgrn'>Travel Time</td>" & vbCrlf & _
		"<td class='tblgrn'>Mileage</td>" & vbCrlf
	CSVHead = "Appointment Date,Client's Last Name,Client's First Name,Actual Start Time,Actual End Time,Duration (mins),Department," & _
		"Language,Interpreter's Last Name, Interpreter's First Name,Billed Hours,Total Amount,Travel Time,Mileage"
	sqlRep = "SELECT * FROM request_T, interpreter_T, institution_T, language_T, dept_T WHERE request_T.[instID] <> 479 AND Dept_T.[index] = [DeptID] AND IntrID = interpreter_T.[index] " & _
		"AND request_T.InstID = institution_T.[index] AND LangID = language_T.[index] AND (request_T.Status = 1 OR request_T.Status = 4) AND (HPID > 0 OR NOT HPID IS NULL)"
	strMSG = "Per-institution request report"
	If Session("type") = 0 Or Session("type") = 4 Or Session("type") = 5 Then
		sqlRep = sqlRep & " AND request_T.InstID = " & Session("InstID")
		strMSG = strMSG & " for " & GetInst2(Session("InstID"))
	ElseIf Session("type") = 3 Then
		sqlRep = sqlRep & " AND DeptID = " & Session("DeptID")
		strMSG = strMSG & " for " & GetInst2(Session("InstID")) & GetMyDept(Session("DeptID"))
	End if
	If tmpReport(1) <> "" Then
		sqlRep = sqlRep & " AND appDate >= '" & tmpReport(1) & "'"
		strMSG = strMSG & " from " & tmpReport(1)
	End If
	If tmpReport(2) <> "" Then
		sqlRep = sqlRep & " AND appDate <= '" & tmpReport(2) & "'"
		strMSG = strMSG & " to " & tmpReport(2)
	End If
	sqlRep = sqlRep & " ORDER BY appDate, AStarttime, Facility, dept, Clname, Cfname"
	Call AddLog("REPORT: " & tmpReport(0) & " FIND: " & sqlRep)
	rsRep.Open sqlRep, g_strCONNLB,3, 1
	If Not rsRep.EOF Then 
		x = 0
		Do Until rsRep.EOF 
			kulay = "#FFFFFF"
			If Not Z_IsOdd(x) Then kulay = "#F5F5F5"
			tmpCliName = rsRep("Clname") & ", " & rsRep("Cfname")
			appTime = ctime(rsRep("AStarttime")) & " - " & ctime(rsRep("AEndtime"))
			appmin = DateDiff("n", rsRep("AStarttime"), rsRep("AEndtime"))
			tmpFacil = rsRep("Dept")
			tmpIName = rsRep("Last name") & ", " & rsRep("first name")
			tmpPay = (rsRep("billable") * rsRep("InstRate")) + Z_CZero(rsRep("TT_Inst")) + Z_CZero(rsRep("M_Inst"))
			strBody = strBody & "<tr bgcolor='" & kulay & "'><td class='tblgrn2'><nobr>" & rsRep("appDate") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & tmpCliName & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & appTime & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & appmin & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & tmpFacil & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & GetLangLB(rsRep("LangID")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & tmpIName & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & rsRep("billable") & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & Z_FormatNumber(tmpPay, 2) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & Z_CZero(rsRep("TT_Inst")) & "</td>" & vbCrLf & _
				"<td class='tblgrn2'><nobr>" & Z_CZero(rsRep("M_Inst")) & "</td></tr>" & vbCrLf
				
			CSVBody = CSVBody & rsRep("appDate") & ",""" & rsRep("Clname") & """,""" & rsRep("Cfname") & """," & rsRep("AStarttime") & _
				"," & rsRep("AEndtime") & "," & appmin & ",""" & rsRep("Dept") & """," & GetLangLB(rsRep("LangID")) & _
				",""" & rsRep("Last name") & """,""" & rsRep("first name") & """," & rsRep("billable") & ",""" & Z_FormatNumber(tmpPay, 2) & _
				"""," & Z_CZero(rsRep("TT_Inst")) & "," & Z_CZero(rsRep("M_Inst")) & vbCrLf
			rsRep.MoveNext
			x = x + 1
		Loop
	Else
		strBody = "<tr><td colspan='11' align='center'><i>&lt --- No records found --- &gt</i></td></tr>"
		CSVBody = "< --- No records found --- >"
	End If
	rsRep.Close
	Set rsRep = Nothing
	Call AddLog("REPORT: " & tmpReport(0) & " SUCCESS.")
End If
If Request("csv") = 1 Then
	'CONVERT TO CSV
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set Prt = fso.CreateTextFile(BackupStr &  RepCSV, True)
	Prt.WriteLine "LANGUAGE BANK - REPORT"
	Prt.WriteLine strMSG
	Prt.WriteLine CSVHead
	Prt.WriteLine CSVHead2
	Prt.WriteLine CSVBody
	Prt.Close	
	Set Prt = Nothing
	Set fso = Nothing
	
	'DONWLOAD
	tmpFile = BackupStr &  RepCSV
	Set dload = Server.CreateObject("SCUpload.Upload")
	dload.Download tmpFile
	Set dload = Nothing
End If
%>
<html>
	<head>
		<title>Interpreter Request - Report Result</title>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
		<!--
		function exportMe()
		{
			document.frmResult.action = "printreport.asp?csv=1"
			document.frmResult.submit();
		}
		-->
		</script>
		<body>
			<form method='post' name='frmResult'>
				<table cellSpacing='0' cellPadding='0' width="100%" bgColor='white' border='0'>
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
							<table bgColor='white' border='0' cellSpacing='2' cellPadding='0' align='center'>
								<tr bgcolor='#C2AB4B'>
									<td colspan='15' align='center'>
										<b><%=strMSG%><b>
									</td>
								</tr>
								<%=strHead%>
								<%=strBody%>
								<tr><td>&nbsp;</td></tr>
								<tr>
									<td colspan='15' align='center' height='100px' valign='bottom'>
										<input class='btn' type='button' value='Print' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='print({bShrinkToFit: true});'>
										<input class='btn' type='button' value='CSV Export' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="exportMe();">
									</td>
								</tr>
									<td colspan='15' align='center' height='100px' valign='bottom'>
										* If needed, please adjust the page orientation of your printer to landscape to view all columns in a single page   
									</td>
								<tr>
								</tr>
							</table>	
						</td>
					</tr>
				</table>
			</form>
		</body>
	</head>
</html>