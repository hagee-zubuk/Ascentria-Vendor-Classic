<%
Function AddLog(strmsg)
	tmpPath = "C:\WORK\InterReq\log\"
	tmpFile = Replace(Date, "/", "") 
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set oFile = fso.OpenTextFile(tmpPath & tmpFile & ".log", 8, True)
	oFile.WriteLine Now & vbTab & strmsg
	Set oFile = Nothing
	Set fso = Nothing
End Function
Function GetStatLB(myid)
	GetStatLB = 0
	Set rsDept = Server.CreateObject("ADODB.RecordSet")
	sqlDept = " SELECT [status] FROM request_T WHERE [HPID] = " & myid
	rsDept.Open sqlDept, g_strCONNLB, 3, 1
	If Not rsDept.EOF Then
		GetStatLB = rsDept("status")
	End If
	rsDept.Close
	Set rsDept = Nothing
End Function
Function GetStat(zzz)
	Select Case zzz
		Case 0 GetStat = "Pending"
		Case 1 GetStat = "Completed"
		Case 2 GetStat = "Missed"
		Case 3 GetStat = "Canceled"
		Case 4 GetStat = "Canceled-Billable"
	End Select
End Function
Function GetMyDept(xxx)
	GetMyDept = ""
	Set rsDept = Server.CreateObject("ADODB.RecordSet")
	sqlDept = " SELECT Dept FROM dept_T WHERE [index] = " & xxx
	rsDept.Open sqlDept, g_strCONNLB, 3, 1
	If Not rsDept.EOF Then
		GetMyDept = " - " & rsDept("Dept")
	End If
	rsDept.Close
	Set rsDept = Nothing
End Function
Function GetInst2(zzz)
	GetInst2 = "N/A"
	If zzz = "" Then Exit Function
	Set rsInst = Server.CreateObject("ADODB.RecordSet")
	sqlInst = "SELECT Facility FROM institution_T WHERE [index] = " & zzz
	rsInst.Open sqlInst, g_strCONNLB, 3, 1
	If Not rsInst.EOF Then
		tmpIname = rsInst("Facility") 
		GetInst2 = tmpIname
	End If
	rsInst.Close
	Set rsInst = Nothing
End Function
Function SkedCheckLenient(intrID, UID, appdate, timefrom, timeto)
	SkedCheckLenient = 0
	Meron = 0
	If Not intrID > 0 Then Exit Function
	'check if same start time
	Set rsSked = Server.CreateObject("ADODB.RecordSet")
	sqlSked = "SELECT * FROM Request_T WHERE status = 0 AND IntrID = " & intrID & " AND appDate = '" & appdate & "' AND appTimeFrom = '" & timefrom & "' " & _
		"AND Request_T.[index] <> " & UID
	rsSked.Open sqlSked, g_strCONNLB, 3, 1
	If Not rsSked.EOF Then
		Meron = 1
	End If	 
	rsSked.Close
	Set rsSked = Nothing
	If Meron = 0 Then
		'check if same end time
		Set rsSked = Server.CreateObject("ADODB.RecordSet")
		sqlSked = "SELECT * FROM Request_T WHERE status = 0 AND IntrID = " & intrID & " AND appDate = '" & appdate & "' AND appTimeFrom = '" & timefrom & "' " & _
			"AND Request_T.[index] <> " & UID
		rsSked.Open sqlSked, g_strCONNLB, 3, 1
		If Not rsSked.EOF Then
			Meron = 1
		End If	 
		rsSked.Close
		Set rsSked = Nothing
	End If
	If Meron = 0 Then
		'check if same time
		Set rsSked = Server.CreateObject("ADODB.RecordSet")
		sqlSked = "SELECT * FROM Request_T WHERE status = 0 AND IntrID = " & intrID & " AND appDate = '" & appdate & "' AND appTimeFrom = '" & timefrom & "' " & _
			"AND apptimeto = '" & timeto & "' AND Request_T.[index] <> " & UID
		rsSked.Open sqlSked, g_strCONNLB, 3, 1
		If Not rsSked.EOF Then
			Meron = 1
		End If	 
		rsSked.Close
		Set rsSked = Nothing
	End If
	If Meron = 0 Then
		'check if between an appt start time
		Set rsSked = Server.CreateObject("ADODB.RecordSet")
		sqlSked = "SELECT * FROM Request_T WHERE status = 0 AND IntrID = " & intrID & " AND appDate = '" & appdate & "' " & _
			"AND appTimeFrom <= '" & timefrom & "' " & _
			"AND apptimeto > '" & timefrom & "' AND Request_T.[index] <> " & UID
		rsSked.Open sqlSked, g_strCONNLB, 3, 1
		If Not rsSked.EOF Then
			Meron = 1
		End If	 
		rsSked.Close
		Set rsSked = Nothing
	End If
	If Meron = 0 Then
		'check if between an appt end time
		Set rsSked = Server.CreateObject("ADODB.RecordSet")
		sqlSked = "SELECT * FROM Request_T WHERE status = 0 AND IntrID = " & intrID & " AND appDate = '" & appdate & "' " & _
			"AND appTimeFrom <= '" & timeto & "' " & _
			"AND apptimeto >= '" & timeto & "' AND Request_T.[index] <> " & UID
		rsSked.Open sqlSked, g_strCONNLB, 3, 1
		If Not rsSked.EOF Then
			Meron = 1
		End If	 
		rsSked.Close
		Set rsSked = Nothing
	End If
	If Meron = 0 Then
		'check if overlap app
		Set rsSked = Server.CreateObject("ADODB.RecordSet")
		sqlSked = "SELECT * FROM Request_T WHERE status = 0 AND IntrID = " & intrID & " AND appDate = '" & appdate & "' " & _
			"AND appTimeFrom >= '" & timefrom & "' " & _
			"AND apptimeto <= '" & timeto & "' AND Request_T.[index] <> " & UID
		rsSked.Open sqlSked, g_strCONNLB, 3, 1
		If Not rsSked.EOF Then
			Meron = 1
		End If	 
		rsSked.Close
		Set rsSked = Nothing
	End If
	SkedCheckLenient = Meron
End Function
Function Z_ResetIntr(appID)
	Set rsRes = Server.CreateObject("ADODB.RecordSet")
	rsRes.Open "UPDATE appt_T Set accept = 0 WHERE appID = " & appID, g_strCONNLB, 1, 3
	Set rsRes = Nothing
	'email intr
End Function
Function Z_ResetIntr2(appID)
	Set rsRes = Server.CreateObject("ADODB.RecordSet")
	rsRes.Open "DELETE FROM appt_T WHERE appID = " & appID, g_strCONNLB, 1, 3
	Set rsRes = Nothing
	Call Z_EmailJob(appID) 'include exceptions
	'email intr
End Function
Function Z_EmailJob(AppID)
	ts = now
	DeptID = Z_GetInfoFROMAppID(appID, "DeptID")
	DeptClass = ClassInt(DeptID)
	LangID = Z_GetInfoFROMAppID(appID, "LangID")
	IDtoLang = UCase(GetLangLB(LangID))
	LangSQL = " (Upper(Language1) = '" & IDtoLang & "' OR Upper(Language2) = '" & IDtoLang & "' OR Upper(Language3) = '" & IDtoLang & _
		"' OR Upper(Language4) = '" & IDtoLang & "' OR Upper(Language5) = '" & IDtoLang & "' OR Upper(Language6) = '" & IDtoLang & "') AND Active = 1"
	If DeptClass = 1 Then classSql = " Social = 1"
	If DeptClass = 2 Then classSql = " Private = 1"
	If DeptClass = 3 Then classSql = " Court = 1"
	If DeptClass = 4 Then classSql = " Medical = 1"
	If DeptClass = 5 Then classSql = " Legal = 1"
	If DeptClass = 6 Then classSql = " Mental = 1"
	AppDate = Z_GetInfoFROMAppID(appID, "AppDate")
	'If AppDate > DateAdd("m", 2, Date) Then Exit Function ' do not send if more than 2 months
	InstID = Z_GetInfoFROMAppID(appID, "InstID")
	appTimeFrom = Z_GetInfoFROMAppID(appID, "appTimeFrom")
	tmpAvail = Weekday(AppDate) & "," & Hour(appTimeFrom)
	appTimeTo = Z_GetInfoFROMAppID(appID, "appTimeTo")
	Set rsIntr = Server.CreateObject("ADODB.RecordSet")
	sqlIntr = "SELECT [index] as myIntrID, [e-mail], Phone1, sendonce FROM interpreter_T WHERE" & classSql & " AND" & LangSQL
	rsIntr.Open sqlIntr, g_strCONNLB, 1, 3
	Do Until rsIntr.EOF
		If Not OnVacation(rsIntr("myIntrID"), AppDate) Then
			If Avail(rsIntr("myIntrID"), tmpAvail) And NotRestrict(rsIntr("myIntrID"), InstID, DeptID) Then
				If SkedCheckLenient(rsIntr("myIntrID"), appID, AppDate, appTimeFrom, appTimeTo) = 0 Then
					'send email here
					If Z_FixNull(rsIntr("e-mail")) <> "" Then
						If AppDate < DateAdd("m", 2, Date) Then ' do not send if more than 2 months
							Urgent = ""
							If DateDiff("n", Now, appTimeFrom) >= 0 And DateDiff("n", Now, appTimeFrom) < 1440 Then Urgent = "URGENT"
							If Not rsIntr("sendonce") Or Urgent = "URGENT" Then
								rsIntr("sendonce") = True
								rsIntr.Update
								strBody = "<p>Language Bank has received new request for your language(s) and skills.<br>" & _
										"Please log into the <a href='https://interpreter.thelanguagebank.org/interpreter/'>" & _
										"LB database</a> and let us know if you are available.</p>" & _
										"<font size='1' face='trebuchet MS'>* Please do not reply to this email. This is a " & _
										"computer generated email." & appID & "</font>"
								retVal = zSendMessage(Trim(rsIntr("e-mail")), "language.services@thelanguagebank.org" _
										, "[LBIS] " & Urgent & " New Appointment in the Database", strBody)
							End If
						End If
					End If
					'save to db
					Set rsApp = Server.CreateObject("ADODB.RecordSet")
					rsApp.Open "SELECT * FROM appt_T WHERE timestamp = '" & ts & "'", g_strCONNLB, 1, 3
					rsApp.AddNew
					rsApp("timestamp") = ts
					rsApp("appID") = appID
					rsApp("IntrID") = rsIntr("myIntrID")
					rsApp.Update
					rsApp.Close
					Set rsApp = Nothing
				End If
			End If
		End If
		rsIntr.MoveNext
	Loop
	rsIntr.Close
	Set rsIntr = Nothing
End Function
Function Z_GetInfoFROMAppID(AppID, infoneeded)
	Set rsIntr = Server.CreateObject("ADODB.RecordSet")
	rsIntr.Open "SELECT " & infoneeded & " FROM request_T WHERE [index] = " & AppID, g_strCONNLB, 3, 1
	If Not rsIntr.EOF Then
		Z_GetInfoFROMAppID = rsIntr(infoneeded)
	End If
	rsIntr.Close
	Set rsIntr = Nothing
End Function
Function OnVacation(IntrID, appDate)
	OnVacation = False
	Set rsVac = Server.CreateObject("ADODB.RecordSet")
	sqlVac = "SELECT vacto, vacfrom, vacto2, vacfrom2 FROM interpreter_T WHERE [index] = " & intrID
	rsVac.Open sqlVac, g_strCONNLB, 3, 1
	If Not rsVac.EOF Then
		If Not IsNull(rsVac("vacfrom")) Then
			If appDate >= rsVac("vacfrom") And appDate <= rsVac("vacto") Then 
				OnVacation = True
			End If
		End If
		If onVacation = False Then
			If Not IsNull(rsVac("vacfrom2")) Then
				If appDate >= rsVac("vacfrom2") And appDate <= rsVac("vacto2") Then 
					OnVacation = True
				End If
			End If
		End If
	End If
	rsVac.Close
	Set rsVac = Nothing
End Function
Function SkedCheck(intrID, UID, appdate, timefrom, timeto)
	SkedCheck = 0
	Meron = 0
	If Not intrID > 0 Then Exit Function
	'check if same start time
	Set rsSked = Server.CreateObject("ADODB.RecordSet")
	sqlSked = "SELECT * FROM Request_T WHERE status = 0 AND IntrID = " & intrID & " AND appDate = '" & appdate & "' AND appTimeFrom = '" & timefrom & "' " & _
		"AND Request_T.[index] <> " & UID
	rsSked.Open sqlSked, g_strCONNLB, 3, 1
	If Not rsSked.EOF Then
		Meron = 1
	End If	 
	rsSked.Close
	Set rsSked = Nothing
	If Meron = 0 Then
		'check if same end time
		Set rsSked = Server.CreateObject("ADODB.RecordSet")
		sqlSked = "SELECT * FROM Request_T WHERE status = 0 AND IntrID = " & intrID & " AND appDate = '" & appdate & "' AND appTimeFrom = '" & timefrom & "' " & _
			"AND Request_T.[index] <> " & UID
		rsSked.Open sqlSked, g_strCONNLB, 3, 1
		If Not rsSked.EOF Then
			Meron = 1
		End If	 
		rsSked.Close
		Set rsSked = Nothing
	End If
	If Meron = 0 Then
		'check if same time
		Set rsSked = Server.CreateObject("ADODB.RecordSet")
		sqlSked = "SELECT * FROM Request_T WHERE status = 0 AND IntrID = " & intrID & " AND appDate = '" & appdate & "' AND appTimeFrom = '" & timefrom & "' " & _
			"AND apptimeto = '" & timeto & "' AND Request_T.[index] <> " & UID
		rsSked.Open sqlSked, g_strCONNLB, 3, 1
		If Not rsSked.EOF Then
			Meron = 1
		End If	 
		rsSked.Close
		Set rsSked = Nothing
	End If
	If Meron = 0 Then
		'check if between an appt start time
		Set rsSked = Server.CreateObject("ADODB.RecordSet")
		sqlSked = "SELECT * FROM Request_T WHERE status = 0 AND IntrID = " & intrID & " AND appDate = '" & appdate & "' " & _
			"AND appTimeFrom <= '" & timefrom & "' " & _
			"AND apptimeto > '" & timefrom & "' AND Request_T.[index] <> " & UID
		rsSked.Open sqlSked, g_strCONNLB, 3, 1
		If Not rsSked.EOF Then
			Meron = 1
		End If	 
		rsSked.Close
		Set rsSked = Nothing
	End If
	If Meron = 0 Then
		'check if between an appt end time
		Set rsSked = Server.CreateObject("ADODB.RecordSet")
		sqlSked = "SELECT * FROM Request_T WHERE status = 0 AND IntrID = " & intrID & " AND appDate = '" & appdate & "' " & _
			"AND appTimeFrom <= '" & timeto & "' " & _
			"AND apptimeto >= '" & timeto & "' AND Request_T.[index] <> " & UID
		rsSked.Open sqlSked, g_strCONNLB, 3, 1
		If Not rsSked.EOF Then
			Meron = 1
		End If	 
		rsSked.Close
		Set rsSked = Nothing
	End If
	'If Meron = 0 Then
	'	'check if no gap between appt
	'	Set rsSked = Server.CreateObject("ADODB.RecordSet")
	'	sqlSked = "SELECT * FROM Request_T WHERE status = 0 AND IntrID = " & intrID & " AND appDate = '" & appdate & "' " & _
	'		"AND apptimeto = '" & timefrom & "' " & _
	'		"AND Request_T.[index] <> " & UID
	'	rsSked.Open sqlSked, g_strCONN, 3, 1
	'	If Not rsSked.EOF Then
	'		Meron = 1
	'	End If	 
	'	rsSked.Close
	'	Set rsSked = Nothing
	'End If
	If Meron = 0 Then
		'check if overlap app
		Set rsSked = Server.CreateObject("ADODB.RecordSet")
		sqlSked = "SELECT * FROM Request_T WHERE status = 0 AND IntrID = " & intrID & " AND appDate = '" & appdate & "' " & _
			"AND appTimeFrom >= '" & timefrom & "' " & _
			"AND apptimeto <= '" & timeto & "' AND Request_T.[index] <> " & UID
		rsSked.Open sqlSked, g_strCONNLB, 3, 1
		If Not rsSked.EOF Then
			Meron = 1
		End If	 
		rsSked.Close
		Set rsSked = Nothing
	End If
	If Meron = 0 Then
		'check if next appointment is less than 2 hr
		Set rsSked = Server.CreateObject("ADODB.RecordSet")
		sqlSked = "SELECT * FROM Request_T WHERE status = 0 AND IntrID = " & intrID & " AND appDate = '" & appdate & "' " & _
			"AND apptimefrom >= '" & dateadd("n", -120, timefrom) & "' " & _
			"AND apptimefrom < '" & timefrom & "' " & _
			"AND Request_T.[index] <> " & UID
		rsSked.Open sqlSked, g_strCONNLB, 3, 1
		If Not rsSked.EOF Then
			Do Until rsSked.EOF
				If datediff("n", rsSked("apptimefrom"), rsSked("apptimeto")) < 121 Then
					If dateadd("n", 120, rsSked("apptimefrom")) > timefrom Then 
						Meron = 1
						Exit Do
					End If
				End If			
				rsSked.MoveNext
			Loop
		End If
		rsSked.Close
		Set rsSked = Nothing
	End If
	If Meron = 0 Then
		'check if previous appointment is less than 2 hr
		If datediff("n", timefrom, timeto) < 121 Then
			Set rsSked = Server.CreateObject("ADODB.RecordSet")
			sqlSked = "SELECT * FROM Request_T WHERE status = 0 AND IntrID = " & intrID & " AND appDate = '" & appdate & "' " & _
				"AND apptimefrom >= '" & timeto & "' " & _
				"AND apptimefrom <= '" & dateadd("n", 30, timeto) & "' " & _
				"AND Request_T.[index] <> " & UID
			'If intrID = 431 Then response.write sqlsked & "<br>"
			rsSked.Open sqlSked, g_strCONNLB, 3, 1
			If Not rsSked.EOF Then
				'Do Until rsSked.EOF
				'	If dateadd("n", 30, rsSked("apptimeto")) > timefrom Then 
						Meron = 1
				'		Exit Do
				'	End If
				'	rsSked.MoveNext
				'Loop	
			End If
			rsSked.Close
			Set rsSked = Nothing
		End If
		
		
		
		
	End If
	SkedCheck = Meron
End Function
Function NotRestrict(IntrID, InstID, DeptID)
	NotRestrict = True
	tmpNotRestrict = 1
	Set rsRes = Server.CreateObject("ADODB.RecordSet")
	sqlRes = "SELECT * FROM Restrict_T WHERE IntrID = " & IntrID & " AND InstID = " & InstID
	rsRes.Open sqlRes, g_strCONNLB, 3, 1
	If Not rsRes.EOF Then
		tmpNotRestrict = 0
	End If
	rsRes.Close
	Set rsRes = Nothing
	If tmpNotRestrict = 1 Then
		Set rsRes = Server.CreateObject("ADODB.RecordSet")
		sqlRes = "SELECT * FROM Restrict2_T WHERE IntrID = " & IntrID & " AND DeptID = " & DeptID
		rsRes.Open sqlRes, g_strCONNLB, 3, 1
		If Not rsRes.EOF Then
			tmpNotRestrict = 0
		End If
		rsRes.Close
		Set rsRes = Nothing
	End If
	If tmpNotRestrict = 0 Then NotRestrict = False
End Function
Function Avail(myID, myTime)
	Avail = False
	Set rsAvail = Server.CreateObject("ADODB.RecordSet")
	sqlAvail = "SELECT * FROM Avail_T WHERE intrID = " & myID & " AND Avail = '" & myTime & "'"
	rsAvail.Open sqlAvail, g_strCONNLB, 3, 1
	If Not rsAvail.EOF Then Avail = True
	rsAvail.Close
	set rsAvail = Nothing
	If Avail Then Exit Function
	Set rsAvail2 = Server.CreateObject("ADODB.RecordSet")
	sqlAvail2 = "SELECT * FROM Avail_T WHERE IntrID = " & myID
	rsAvail2.Open sqlAvail2, g_strCONNLB, 3, 1
	If rsAvail2.EOF Then Avail = True
	rsAvail2.Close
	set rsAvail2 = Nothing
End Function
Function ActiveSage(deptID)
	If Z_CZero(deptID) = 0 Then Exit Function
	Set rsDept = Server.CreateObject("ADODB.RecordSet")
	rsDept.Open "UPDATE Dept_T SET SageActive = 1 WHERE [index] = " & deptID, g_strCONNLB, 1, 3
	Set rsDept = Nothing
End Function
Function Z_GetBillhrsCourt(timefrom, timeto) 'finish this
	Z_GetBillhrsCourt = 1.5
	If DateDiff("n", timefrom, timeto) > 90 Then
		tmpBillMin = DateDiff("n", timefrom, timeto)
		timebefore75 = tmpBillMin / 60
		tmpBillHrs = timebefore75 * 0.75
		tmpBillMHrs = Int(tmpBillHrs)
		tmpLen = Len(tmpBillHrs)
		tmpPosDec = Instr(tmpBillHrs, ".")
		tmpBillMMin = Z_FormatNumber("0" & Right(tmpBillHrs, tmpLen - (tmpPosDec - 1)), 2)
		If Cdbl(tmpBillMMin) > 0.009 And  Cdbl(tmpBillMMin) < 0.22 Then
			Z_GetBillhrsCourt = tmpBillMHrs
		ElseIf Cdbl(tmpBillMMin) => 0.22 And  Cdbl(tmpBillMMin) < 0.38 Then
			Z_GetBillhrsCourt = tmpBillMHrs + 0.25
		ElseIf Cdbl(tmpBillMMin) => 0.38 And  Cdbl(tmpBillMMin) < 0.63 Then
			Z_GetBillhrsCourt = tmpBillMHrs + 0.50
		ElseIf Cdbl(tmpBillMMin) => 0.63 And  Cdbl(tmpBillMMin) < 0.88 Then
			Z_GetBillhrsCourt = tmpBillMHrs + 0.75
		ElseIf Cdbl(tmpBillMMin) => 0.88 Then
			Z_GetBillhrsCourt = tmpBillMHrs + 1
		Else
			Z_GetBillhrsCourt = tmpBillMHrs
		End If
	End If
End Function
Function Z_GetBillhrs(timefrom, timeto)
	Z_GetBillhrs = 2
	If DateDiff("n", timefrom, timeto) > 120 Then
		tmpBillMin = DateDiff("n", timefrom, timeto)
		tmpBillHrs = tmpBillMin / 60
		tmpBillMHrs = Int(tmpBillHrs)
		tmpLen = Len(tmpBillHrs)
		tmpPosDec = Instr(tmpBillHrs, ".")
		tmpBillMMin = Z_FormatNumber("0" & Right(tmpBillHrs, tmpLen - (tmpPosDec - 1)), 2)
		If Cdbl(tmpBillMMin) > 0.009 And  Cdbl(tmpBillMMin) <= 0.5 Then
			Z_GetBillhrs = tmpBillMHrs + 0.5
		ElseIf  Cdbl(tmpBillMMin) > 0.5 And  Cdbl(tmpBillMMin) <= 0.99 Then
			Z_GetBillhrs = tmpBillMHrs + 1
		Else
			Z_GetBillhrs = tmpBillMHrs
		End If
	End If
End Function
Function ClassInt(deptid)
	ClassInt = 0
	Set rsClass = Server.CreateObject("ADODB.RecordSet")
	sqlReq = "SELECT class FROM Dept_T WHERE [index] = " & deptid
	rsClass.Open sqlReq, g_strCONNLB, 3, 1
	If not rsClass.EOF Then
		ClassInt = Z_Czero(rsClass("class"))
	End If
	rsClass.Close
	Set rsClass = Nothing
End Function
Function Z_HideApp(uid)
	Z_HideApp = False
	If Z_Czero(uid) = 0 Then Exit Function
	Set rsRate = Server.CreateObject("ADODB.RecordSet")
	sqlRate = "SELECT hideapp FROM user_T WHERE [index] = " & uid
	rsRate.Open sqlRate, g_strCONN, 3, 1
	If Not rsRate.EOF Then
		If rsRate("hideapp") Then Z_HideApp = True
	End If
	rsRate.Close
	Set rsRate = Nothing
End Function
Function GetDeptRate(xxx)
	GetDeptRate = 0
	Set rsRate = Server.CreateObject("ADODB.RecordSet")
	sqlRate = "SELECT defrate FROM dept_T WHERE [index] = " & xxx
	rsRate.Open sqlRate, g_strCONNLB, 3, 1
	If Not rsRate.EOF Then
		GetDeptRate = rsRate("defrate")
	End If
	rsRate.Close
	Set rsRate = Nothing
End Function
Function CTime(tmptime)
	if z_fixnull(tmptime) <> "" Then 
		myTime = Right(tmptime, 11)
		If instr(myTime, "/") > 0 Then
			Ctime = ""
		Else
			Ctime = cdate(myTime)
		End If
	Else
		Ctime = ""
	End If
End Function
Function GetSun(xxx)
	If Weekday(xxx) = 1 Then
		GetSun = xxx
	Else
		tmpDate = xxx
		Do Until Weekday(tmpDate) = 1
			tmpDate = DateAdd("d", "-1", tmpDate)
		Loop
		GetSun = tmpDate
	End If
End Function
Function GetSat(xxx)
	If Weekday(xxx) = 7 Then
		GetSat = xxx
	Else
		tmpDate = xxx
		Do Until Weekday(tmpDate) = 7
			tmpDate = DateAdd("d", 1, tmpDate)
		Loop
		GetSat = tmpDate
	End If
End Function
Function SaveHist(xxx, mypage)
	'SAVE HIST SQL
	tmpHist = ""
	Set rsHist = Server.CreateObject("ADODB.RecordSet")
	Set rsLB = Server.CreateObject("ADODB.RecordSet")
	sqlHist = "SELECT * FROM hist_T WHERE timestamp = '" & Now & "' "
	sqlLB = "SELECT * FROM request_T WHERE [index] = " & xxx
	rsLB.Open sqlLB, g_strCONNLB, 1, 3
	rsHist.Open sqlHist, g_strCONNHist2, 1,3 
	If not rsLB.EOF Then
		rsHist.AddNew
		rsHist("LBID") = xxx
		rsHist("Timestamp") = Now
		rsHist("Author") = Session("GreetMe")
		rsHist("pageused") = mypage
		x = 1
On error resume next
    Do Until x = rsLB.Fields.Count
        tmpHist = tmpHist & """" & rsLB.Fields(x).Value & ""","
        x = x + 1
    Loop
    rsHist("Hist") = trim(tmpHist)
		rsHist.Update
	End If
	rsLB.CLose
	set rsLB = Nothing
	rsHist.Close
	Set rsHist = Nothing
	SaveHist = True
End Function
Function GetDeptCity(xxx)
	GetDeptCity = "N/A"
	Set rsInst = Server.CreateObject("ADODB.RecordSet")
	sqlInst = "SELECT * FROM dept_T WHERE [index] = " & xxx
	rsInst.Open sqlInst, g_strCONNLB, 3, 1
	If Not rsInst.EOF Then
		GetDeptCity = rsInst("city") 
	End If
	rsInst.Close
	Set rsInst = Nothing
End Function
Function GetDeptAdr(xxx)
	GetDeptAdr = "N/A"
	Set rsInst = Server.CreateObject("ADODB.RecordSet")
	sqlInst = "SELECT * FROM dept_T WHERE [index] = " & xxx
	rsInst.Open sqlInst, g_strCONNLB, 3, 1
	If Not rsInst.EOF Then
		GetDeptAdr = rsInst("address") & ", " & rsInst("InstAdrI") & ", " & rsInst("city") & ", " & rsInst("state") & ", " & rsInst("zip")
	End If
	rsInst.Close
	Set rsInst = Nothing
End Function
Function GetInstNameLB(xxx)
	GetInstNameLB = "N/A"
	Set rsInst = Server.CreateObject("ADODB.RecordSet")
	sqlInst = "SELECT * FROM Institution_T WHERE [index] = " & xxx
	rsInst.Open sqlInst, g_strCONNLB, 3, 1
	If Not rsInst.EOF Then
		GetInstNameLB = rsInst("Facility") 
	End If
	rsInst.Close
	Set rsInst = Nothing
End Function
Function GetIntrNameLB(xxx)
	GetInstNameLB = "N/A"
	Set rsInst = Server.CreateObject("ADODB.RecordSet")
	sqlInst = "SELECT * FROM interpreter_T WHERE [index] = " & xxx
	rsInst.Open sqlInst, g_strCONNLB, 3, 1
	If Not rsInst.EOF Then
		GetIntrNameLB = rsInst("last name") & ", " & rsInst("first name") 
	End If
	rsInst.Close
	Set rsInst = Nothing
End Function
Function GetIntrNameLB2(xxx)
	GetInstNameLB2 = "N/A"
	Set rsInst = Server.CreateObject("ADODB.RecordSet")
	sqlInst = "SELECT * FROM interpreter_T WHERE [index] = " & xxx
	rsInst.Open sqlInst, g_strCONNLB, 3, 1
	If Not rsInst.EOF Then
		GetIntrNameLB2 = rsInst("first name") & " " & rsInst("last name") 
	End If
	rsInst.Close
	Set rsInst = Nothing
End Function
Function GetDeptNameLB(xxx)
	GetDeptNameLB = "N/A"
	Set rsInst = Server.CreateObject("ADODB.RecordSet")
	sqlInst = "SELECT * FROM dept_T WHERE [index] = " & xxx
	rsInst.Open sqlInst, g_strCONNLB, 3, 1
	If Not rsInst.EOF Then
		GetDeptNameLB = rsInst("dept") 
	End If
	rsInst.Close
	Set rsInst = Nothing
End Function
Function GetCom(xxx)
	GetCom = ""
	Set rsCom = Server.CreateObject("ADODB.RecordSet")
	sqlKey = "SELECT * FROM Comment_T  WHERE UID = " & xxx
	rsCom.Open sqlKey, g_strCONN, 3, 1
	If Not rsCom.EOF Then
		GetCom = Replace(rsCom("comment"), vbCrlf, " ")
	End If
	rsCom.Close
	Set rsCom = Nothing
End Function
Function GetKey(xxx)
	GetKey = "N/A"
	Set rsKey = Server.CreateObject("ADODB.RecordSet")
	sqlKey = "SELECT * FROM appointment_T WHERE [index] = " & xxx
	rsKey.Open sqlKey, g_strCONN, 3, 1
	If Not rsKey.EOF Then
		If rsKey("Key") = 1 Then GetKey = "Completed"
		If rsKey("Key") = 2 Then GetKey = "Canceled"
	End If
	rsKey.Close
	Set rsKey = Nothing
End Function
Function GetInc(xxx, yyy)
	tmpMin = DateDiff("n", xxx, yyy)
	tmpInc = DateDiff("n", xxx, yyy) / 15
    If (tmpMin Mod 15) = 0 Then
        GetInc = Int(tmpInc)
    Else
        GetInc = Int(tmpInc) + 1
    End If
End Function
Function GetReas(xxx)
	GetReas = ""
	tmpReas = Split(xxx, "|")
	CtrReas = Ubound(tmpReas)
	x = 0
	Do Until x = CtrReas + 1
		Set rsReas = Server.CreateObject("ADODB.RecordSet")
		sqlReas = "SELECT * FROM Reason_T WHERE [index] = " & tmpReas(x)
		rsReas.Open sqlReas, g_strCONN, 3, 1
		If Not rsReas.EOF Then
			GetReas = GetReas & rsReas("reason") & "<br>"
		End If
		rsReas.Close
		Set rsReas = Nothing
		x = x + 1
	Loop
End Function
Function GetFacility(xxx)
	GetFacility = "N/A"
	Set rsFac = Server.CreateObject("ADODB.RecordSet")
	sqlFac = "SELECT * FROM institution_T WHERE [index] = " & xxx
	rsFac.Open sqlFac, g_strCONNLB, 3, 1
	If Not rsFac.EOF Then
		GetFacility = rsFac("facility")
	End If
	rsFac.Close
	Set rsFac = Nothing
End Function
Function GetDept(xxx)
	GetDept = "N/A"
	Set rsFac = Server.CreateObject("ADODB.RecordSet")
	sqlFac = "SELECT * FROM dept_T WHERE [index] = " & xxx
	rsFac.Open sqlFac, g_strCONNLB, 3, 1
	If Not rsFac.EOF Then
		GetDept = rsFac("dept")
	End If
	rsFac.Close
	Set rsFac = Nothing
End Function
Function GetIntrID(xxx)
	'GET INTR ID
	GetIntrID = 0
	Set rsIntr = Server.CreateObject("ADODB.RecordSet")
	sqlIntr = "SELECT * FROM Intr_T WHERE [index] = " & xxx
	rsIntr.Open sqlIntr, g_strCONN, 3, 1
	If Not rsIntr.EOF Then
		HPID = rsIntr("Index")
	End If
	rsIntr.Close
	Set rsIntr = Nothing
	'CHECK IF FROM LB DB
	If HPID <> "" Then
		Set rsLB = Server.CreateObject("ADODB.RecordSet")
		sqlLB = "SELECT * FROM IntrLB_T  WHERE IntrID = " & HPID
		rsLB.Open sqlLB, g_strCONN, 3, 1
		If Not rsLB.EOF Then
			GetIntrID = rsLB("LBIntrID")
		End If
		rsLB.Close
		Set rsLB = Nothing
	End If
End Function
Function GetIntrName(xxx)
	GetIntrName = "N/A"
	Set rsIntr = Server.CreateObject("ADODB.RecordSet")
	sqlIntr = "SELECT * FROM Intr_T WHERE [index] = " & xxx
	rsIntr.Open sqlIntr, g_strCONN, 3, 1
	If Not rsIntr.EOF Then
		GetIntrName = rsIntr("lname") & ", " & rsIntr("fname")
	End If
	rsIntr.Close
	Set rsIntr = Nothing
End Function
Function GetIntrNameLB(xxx)
	GetIntrNameLB = "N/A"
	Set rsIntr = Server.CreateObject("ADODB.RecordSet")
	sqlIntr = "SELECT * FROM interpreter_T WHERE [index] = " & xxx
	rsIntr.Open sqlIntr, g_strCONNLB, 3, 1
	If Not rsIntr.EOF Then
		GetIntrNameLB = rsIntr("last name") & ", " & rsIntr("first name")
	End If
	rsIntr.Close
	Set rsIntr = Nothing
End Function
Function GetLang(xxx)
	GetLang = "N/A"
	Set rsLang = Server.CreateObject("ADODB.RecordSet")
	sqlLang = "SELECT * FROM Lang_T WHERE [index] = " & xxx
	rsLang.Open sqlLang, g_strCONN, 3, 1
	If Not rsLang.EOF Then
		GetLang = rsLang("lang")
	End If
	rsLang.Close
	Set rsLang = Nothing
End Function
Function GetLangLB(xxx)
	GetLangLB = "N/A"
	Set rsLang = Server.CreateObject("ADODB.RecordSet")
	sqlLang = "SELECT * FROM language_T WHERE [index] = " & xxx
	rsLang.Open sqlLang, g_strCONNLB, 3, 1
	If Not rsLang.EOF Then
		GetLangLB = rsLang("language")
	End If
	rsLang.Close
	Set rsLang = Nothing
End Function
Function GetInstName(xxx)
	GetInstName = "N/A"
	Set rsInst = Server.CreateObject("ADODB.RecordSet")
	sqlInst = "SELECT * FROM Inst_T WHERE [index] = " & xxx
	rsInst.Open sqlInst, g_strCONN, 3, 1
	If Not rsInst.EOF Then
		GetInstName = rsInst("Inst") 
	End If
	rsInst.Close
	Set rsInst = Nothing
End Function
Function CheckApp(tmpdate)
	CheckApp = "#FFFFFF"
	Set rsReq = Server.CreateObject("ADODB.RecordSet")
	If Session("type") = 0 Or Session("type") = 4  Or Session("type") = 5 Then
		'sqlReq = "SELECT * FROM Appointment_T WHERE appDate = '" & tmpDate & "' AND InstID = " & Session("InstID") & " ORDER BY TimeFrom"
		If Session("InstID") = 93 Then
			sqlReq = "SELECT * FROM appointment_T WHERE appDate = '" & tmpDate & "' AND InstID = 93 ORDER BY TimeFrom"
		ElseIf Session("UID") = 509 Or Session("UID") = 510 Then 'special rule for user 509 and 510
			sqlReq = "SELECT * FROM appointment_T WHERE appDate = '" & tmpDate & "' AND " & _
				"(DeptID = 2306 OR DeptID = 373 OR DeptID = 946 OR DeptID = 2302) ORDER BY TimeFrom"
		ElseIf Session("UID") = 517 Or Session("UID") = 534 Then 'special rule for user 517 and 534
			sqlReq = "SELECT * FROM appointment_T WHERE appDate = '" & tmpDate & "' AND " & _
				"(DeptID = 446 OR DeptID = 322 OR DeptID = 289) ORDER BY TimeFrom"
		ElseIf Session("UID") = 625 Or Session("UID") = 626 Or Session("UID") = 627 Or Session("UID") = 628 Then 'special rule for user 625 - 628
			sqlReq = "SELECT * FROM appointment_T WHERE appDate = '" & tmpDate & "' AND " & _
				"(DeptID = 2466 OR DeptID = 2465) ORDER BY TimeFrom"
		Else
			sqlReq = "SELECT * FROM appointment_T WHERE appDate = '" & tmpDate & "' AND InstID = " & Session("InstID") & " ORDER BY TimeFrom"
		End If
	ElseIf Session("type") = 1 Then
		sqlReq = "SELECT * FROM Appointment_T WHERE appDate = '" & tmpDate & "' AND IntrID = " & Session("IntrID") & " ORDER BY TimeFrom"
	ElseIf Session("type") = 2 Then
		sqlReq = "SELECT * FROM Appointment_T WHERE appDate = '" & tmpDate & "' ORDER BY TimeFrom"
	ElseIf Session("type") = 3 Then
		sqlReq = "SELECT * FROM Appointment_T WHERE appDate = '" & tmpDate & "' AND DeptID = " & Session("DeptID") & " ORDER BY TimeFrom"
	ElseIf Session("type") = 6 Then
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set oFile = fso.OpenTextFile(crtLst, 1)
		Do Until oFile.AtEndOfStream
			oLine = oFile.ReadLine
			strInst = strInst & "InstID = " & oLine & " OR "
		Loop
		Set oFile = Nothing
		Set fso = Nothing
		sqlInst = Mid(strInst, 1, Len(strInst) - 4)
		sqlReq = "SELECT * FROM appointment_T WHERE appDate = '" & tmpDate & "' AND (" & sqlInst & ") ORDER BY TimeFrom"
	End If
	rsReq.Open sqlReq, g_strCONN, 3, 1
	If Not rsReq.EOF Then
		If Not Z_HideApp(rsReq("UID")) Or Session("type") = 3 Then
			CheckApp = "#FFFFCE"
		End If
	End If
	rsReq.Close
	Set rsReq = Nothing
End Function
Function GetInstID(xxx)
	GetInstID = 0
	Set rsInst = Server.CreateObject("ADODB.RecordSet")
	sqlInst = "SELECT * FROM InstUser_T, Inst_T WHERE UID = " & xxx & _
		" AND InstID = Inst_T.[index]"
	rsInst.Open sqlInst, g_strCONN, 3, 1
	If Not rsInst.EOF Then
		GetInstID = rsInst("InstID")
	End If
	rsInst.Close
	Set rsInst = Nothing	
End Function
Function Z_FormatTime(xxx)
	Z_FormatTime = Null
	If xxx <> "" Or Not IsNull(xxx)  Then
		If IsDate(xxx) Then Z_FormatTime = FormatDateTime(xxx, 4) 
	End If
End Function
Function GetReason(xxx, yyy)
	GetReason = "" 
	Set rsLbl = Server.CreateObject("ADODB.RecordSet")
	If xxx = 1 Then 'ENCOUNTER
		sqlLbl = "SELECT * FROM complete_T WHERE [index] = " & yyy
	ElseIf xxx = 2 Then 'CANCELED
		sqlLbl = "SELECT * FROM cancel_T WHERE [index] = " & yyy
	Else
		Exit Function
	End If
	rsLbl.Open sqlLbl, g_strCONN, 3, 1
	If Not rsLbl.EOF Then
		If xxx = 1 Then
			GetReason = rsLbl("completeReason")
		ElseIf xxx = 2 Then 
			GetReason = rsLbl("cancelReason")
		End If
	End If	
	rsLbl.Close
	Set rsLbl = Nothing
End Function
Function GetIntrNameEnc(xxx)
	tmpNum = Z_CZero(xxx)
	If tmpNum = 999 Then 
		GetIntrNameEnc = "ALL"
	ElseIf xxx = 0 Then
		GetIntrNameEnc = "N/A"
	Else
		Set rsIntr = Server.CreateObject("ADODB.RecordSet")
		sqlIntr = "SELECT * FROM Intr_T WHERE [index] = " & xxx
		rsIntr.Open sqlIntr, g_strCONN, 3, 1
		If Not rsIntr.EOF Then
			GetIntrNameEnc = rsIntr("lname") & ", " & rsIntr("fname")
		End If
		rsIntr.Close
		Set rsIntr = Nothing
	End If
End Function
Function GetIntrNameEnc2(xxx)
	tmpNum = Z_CZero(xxx)
	If tmpNum = 999 Then 
		GetIntrNameEnc2 = "ALL"
	ElseIf xxx = 0 Then
		GetIntrNameEnc2 = "N/A"
	Else
		Set rsIntr = Server.CreateObject("ADODB.RecordSet")
		sqlIntr = "SELECT * FROM Interpreter_T WHERE [index] = " & xxx
		rsIntr.Open sqlIntr, g_strCONNLB, 3, 1
		If Not rsIntr.EOF Then
			GetIntrNameEnc2 = rsIntr("last name") & ", " & rsIntr("first name")
		End If
		rsIntr.Close
		Set rsIntr = Nothing
	End If
End Function
Function GetComment(xxx)
	GetComment = ""
	Set rsCom = Server.CreateObject("ADODB.RecordSet")
	sqlCom = "SELECT * FROM Comment_T WHERE UID = " & xxx
	rsCom.Open sqlCom, g_strCONN, 3, 1
	If Not rsCom.EOF Then
		GetComment = rsCom("comment")
	End If
	rsCom.Close
	Set rsCom = Nothing
End Function
Function GetReason2(xxx,yyy)
	GetReason2 = "N/A"
	Set rsReas = Server.CreateObject("ADODB.RecordSet")
	sqlReas = "SELECT * FROM "
	If Cint(yyy) = 1 Then
		sqlReas = sqlReas & "Complete_T "
	ElseIf Cint(yyy) = 2 Then
		sqlReas = sqlReas & "Cancel_T "
	End If
	sqlReas = sqlReas & "WHERE [index] = " & xxx
	rsReas.Open sqlReas, g_strCONN, 3, 1
	If Not rsReas.EOF Then
		If Cint(yyy) = 1 Then
			GetReason2 = rsReas("completeReason")
		ElseIf Cint(yyy) = 2 Then
			GetReason2 = rsReas("cancelReason")
		End If
	End If
	rsReas.Close
	Set rsReas = Nothing
End Function
Function CutMe(xxx)
	CutMe = xxx
	If Len(xxx) > 41 Then
		CutMe = Left(xxx, 37) & "..."
	End If
End Function
Function GetHPID(xxx)
	GetHPID = 0
	Set rsHPID = Server.CreateObject("ADODB.RecordSet")
	sqlHPID = "SELECT * FROM Intr_T WHERE UID = " & xxx
	rsHPID.Open sqlHPID, g_strCONN, 3, 1
	If Not rsHPID.EOF Then
		GetHPID = rsHPID("index")
	End If
	rsHPID.Close
	Set rsHPID = Nothing
End Function
%>
