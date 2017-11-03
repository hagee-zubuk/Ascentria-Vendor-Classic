<%
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
Function GUIDExists(xxx)
	GUIDExists = False
	If xxx = "" Then Exit Function
	Set fso = CreateObject("Scripting.FileSystemObject")
	If fso.FileExists(F604AStr & xxx & ".PDF") Then GUIDExists = True
	Set fso = Nothing
End Function
Function Z_GenerateGUID()
	Dim objGUID
	Set objGUID = Server.CreateObject("Z_MkGUID.ZGUID")
	Z_GenerateGUID = objGUID.GetGUID()
	Set objGUID = Nothing
End Function
Function Z_Replace(var, del, rpl)
	If IsNull(var) Then 
		Z_Replace = 0
		Exit Function
	End If
	If Instr(var, del) <> 0 Then 
		Z_Replace = Replace(var, del, rpl)
	Else
		Z_Replace  = var
	End If
End Function
Function Z_DateNull(var)
	Dim dblTmp
    'Z_DateNull = False
    If IsNull(var) Then 
    	Z_DateNull = Empty
    ElseIf var = "" Then 
    	Z_DateNull = Empty
    ElseIf Not IsDate(var) Then
  		Z_DateNull = Empty
  	Else
    	Z_DateNull = cdate(var)
    End If
End Function
Function Z_IsOdd2(var)
	Dim dblTmp
    Z_IsOdd2 = False
    If IsNull(var) Then Exit Function
    If var = "" Then Exit Function
    If Not IsNumeric(var) Then Exit Function
    Z_IsOdd2 = CBool((var Mod 2) = 1) Or CBool((var Mod 2) = -1)
End Function

Function Z_CZero(var)
	If IsNull(var) Then 
		Z_CZero = Cdbl(0)
	ElseIf var = "" Then 
		Z_CZero = Cdbl(0)	
	ElseIf Not IsNumeric(var) Then
		Z_CZero = Cdbl(0)
	Else
		Z_CZero = Cdbl(var)
	End If
End Function

Function Z_CEmpty(var)
	If IsNull(var) Then 
		Z_CEmpty = ""
	ElseIf var = "" Then 
		Z_CEmpty = ""	
	Else
		Z_CEmpty = var
	End If
End Function

Function Z_CDate(var)
	If IsNull(var) Then Z_CDate = Empty
	If var = "" Then Z_CDate = Empty
	If IsDate(var) Then Z_CDate = CDate(var)
End Function

Function Z_IsOdd(var)
DIM dblTmp
	Z_IsOdd = False
	If IsNull(var) Then Exit Function
	If var = "" Then Exit Function
	If Not IsNumeric(var) Then Exit Function
	Z_IsOdd = CBool( (var Mod 2) = 1 )
End Function

Function Z_FixNull(vntZ)
	If IsNull(vntZ) Then
		Z_FixNull = ""
	ElseIf IsEmpty(vntZ) Then
		Z_FixNull = ""
	Else
		Z_FixNull = vntZ
	End If
End Function

Function Z_NullFix(vntZ)
	If IsNull(vntZ) Then
		Z_NullFix = Null
	ElseIf Trim(vntZ) = "" Then
		Z_NullFix = Null
	Else
		Z_NullFix = vntZ
	End If
End Function

Function Z_Blank(vntZ)
	Z_Blank = False
	If IsNull(vntZ) Then
		Z_Blank = True
		Exit Function
	ElseIf Trim(vntZ) = "" Then
		Z_Blank = True
	End If
End Function

Function Z_MDYDate(dtDate)
DIM lngTmp, strDay, strTmp
	If Not IsDate(dtDate) Then Exit Function
	lngTmp = Z_CLng(DatePart("m", dtDate))
	If lngTmp < 10 Then Z_MDYDate = "0"
	Z_MDYDate = Z_MDYDate & lngTmp & "/"
	lngTmp = Z_CLng(DatePart("d", dtDate))
	If lngTmp < 10 Then Z_MDYDate = Z_MDYDate & "0"
	Z_MDYDate = Z_MDYDate & lngTmp & "/"
	strTmp = DatePart("yyyy", dtDate)
	Z_MDYDate = Z_MDYDate & Right(strTmp,2)
End Function

Function Z_SFDate(dtDate)
' semiflowery date
DIM lngTmp, strDay, strTmp
	If Not IsDate(dtDate) Then Exit Function
	lngTmp = DatePart("m", dtDate)
	strTmp = MonthName(lngTmp, False) & " "
	strTmp = strTmp & DatePart("d", dtDate)
	lngTmp = Z_CLng(Right(strTmp,1))
	If lngTmp = 1 Then
		strTmp = strTmp & "st"
	ElseIf lngTmp = 2 Then
		strTmp = strTmp & "nd"
	ElseIf lngTmp = 3 Then
		strTmp = strTmp & "rd"
	Else
		strTmp = strTmp & "th"
	End If
	strTmp = strTmp & ", " & DatePart("yyyy", dtDate)
	Z_SFDate = strTmp
End Function

Function Z_DateAdd(dtDate, lngPd)
' returns a Date: lngPd business days from dtDate
DIM	lngAdded, dtTmp, lngDy, lngTmp, lngYr
	If Not IsDate(dtDate) Then Exit Function
	lngPd = Z_CLng(lngPd)
	If lngPd = 0 Then
		Z_DateAdd = Z_SFDate(dtDate)
		Exit Function
	ElseIf lngPd > 0 Then
		lngDy = 1
	ElseIf lngPd < 0 Then
		lngDy = -1
	End If
	dtTmp = dtDate
	lngAdded = 0
	Do While lngAdded < lngPd
		dtTmp = DateAdd("d", lngDy, dtTmp)
		lngTmp = DatePart("w", dtTmp, vbSunday)
		lngYr = DatePart("yyyy", dtTmp)
		'Response.Write "<!-- " & dtTmp & ": " & lngTmp & " -->" & vbCrLf
		If lngTmp > 1 And lngTmp < 7 Then
			lngAdded = lngAdded + 1
			' holiday check
			If dtTmp = CDate("12/25/" & lngYr) Or dtTmp = CDate("1/1/" & lngYr) Or _
					dtTmp = CDate("1/2/" & lngYr) Or dtTmp = CDate("1/19/" & lngYr) Or _
					dtTmp = CDate("2/16/" & lngYr) Or dtTmp = CDate("5/31/" & lngYr) Or _
					dtTmp = CDate("7/5/" & lngYr) Or dtTmp = CDate("9/6/" & lngYr) Or _
					dtTmp = CDate("10/11/" & lngYr) Or dtTmp = CDate("11/25/" & lngYr) Or _
					dtTmp = CDate("11/26/" & lngYr) Or dtTmp = CDate("12/24/" & lngYr) Then
				lngAdded = lngAdded - 1
			End If
		End If
	Loop
	Z_DateAdd = Z_SFDate(dtTmp)
End Function


Function Z_MDYDateAdd(dtDate, lngPd)
' returns a Date: lngPd business days from dtDate
DIM	lngAdded, dtTmp, lngDy, lngTmp, lngYr
	If Not IsDate(dtDate) Then Exit Function
	lngPd = Z_CLng(lngPd)
	If lngPd = 0 Then
		Z_MDYDateAdd = Z_MDYDate(dtDate)
		Exit Function
	ElseIf lngPd > 0 Then
		lngDy = 1
	ElseIf lngPd < 0 Then
		lngDy = -1
	End If
	dtTmp = dtDate
	lngAdded = 0
	Do While lngAdded < lngPd
		dtTmp = DateAdd("d", lngDy, dtTmp)
		lngTmp = DatePart("w", dtTmp, vbSunday)
		lngYr = DatePart("yyyy", dtTmp)
		'Response.Write "<!-- " & dtTmp & ": " & lngTmp & " -->" & vbCrLf
		If lngTmp > 1 And lngTmp < 7 Then
			lngAdded = lngAdded + 1
			' holiday check
			If dtTmp = CDate("12/25/" & lngYr) Or dtTmp = CDate("1/1/" & lngYr) Or _
					dtTmp = CDate("1/2/" & lngYr) Or dtTmp = CDate("1/19/" & lngYr) Or _
					dtTmp = CDate("2/16/" & lngYr) Or dtTmp = CDate("5/31/" & lngYr) Or _
					dtTmp = CDate("7/5/" & lngYr) Or dtTmp = CDate("9/6/" & lngYr) Or _
					dtTmp = CDate("10/11/" & lngYr) Or dtTmp = CDate("11/25/" & lngYr) Or _
					dtTmp = CDate("11/26/" & lngYr) Or dtTmp = CDate("12/24/" & lngYr) Then
				lngAdded = lngAdded - 1
			End If
		End If
	Loop
	Z_MDYDateAdd = Z_MDYDate(dtTmp)
End Function


Function Z_MDYCalDateAdd(dtDate, lngPd)
' returns a Date: lngPd business days from dtDate
DIM	lngAdded, dtTmp, lngDy, lngTmp, lngYr
	If Not IsDate(dtDate) Then Exit Function
	lngPd = Z_CLng(lngPd)
	If lngPd = 0 Then
		Z_MDYCalDateAdd = Z_MDYDate(dtDate)
		Exit Function
	ElseIf lngPd > 0 Then
		lngDy = 1
	ElseIf lngPd < 0 Then
		lngDy = -1
	End If
	dtTmp = dtDate
	lngAdded = 0
	dtTmp = DateAdd("d", lngPd, dtTmp)
	Do While True
		lngTmp = DatePart("w", dtTmp, vbSunday)
		If lngTmp > 1 And lngTmp < 7 Then
			lngYr = DatePart("yyyy", dtTmp)
			' holiday check
			If dtTmp = CDate("12/25/" & lngYr) Or dtTmp = CDate("1/1/" & lngYr) Or _
					dtTmp = CDate("1/2/" & lngYr) Or dtTmp = CDate("1/19/" & lngYr) Or _
					dtTmp = CDate("2/16/" & lngYr) Or dtTmp = CDate("5/31/" & lngYr) Or _
					dtTmp = CDate("7/5/" & lngYr) Or dtTmp = CDate("9/6/" & lngYr) Or _
					dtTmp = CDate("10/11/" & lngYr) Or dtTmp = CDate("11/25/" & lngYr) Or _
					dtTmp = CDate("11/26/" & lngYr) Or dtTmp = CDate("12/24/" & lngYr) Then
				dtTmp = DateAdd("d", 1, dtTmp)
			Else
				Exit Do
			End If
		Else
			dtTmp = DateAdd("d", 1, dtTmp)
		End If
	Loop
	Z_MDYCalDateAdd = Z_MDYDate(dtTmp)
End Function

Function Z_FixPath(path)
	If Right(path,1)<>"\" Then
		Z_FixPath = path & "\"
	Else
		Z_FixPath = path
	End If
End Function

Function Z_FixVRoot(strWD, strBase)
	Dim strRes, i, strArry
	i = (Len(strWD)-Len(g_FilesPath))
	If i > 0 Then 
		strRes = Right(strWD, i)
		strArry = Split(strRes,"\")
		strRes = ""
		For i = 0 to UBound(strArry)
			if strArry(i)<>"" Then strRes= strRes & strArry(i) & "/"
		Next
		Z_FixVRoot = strRes
	End If
End Function

Function Z_CleanExt(name)
	Dim i
	i = InStrRev(name, ".")
	If i>0 Then Z_CleanExt = Left(name, i-1) Else Z_CleanExt = name
End Function

Function Z_GetExt(name)
	Dim i, j
	j = Len(name)
	i = InStrRev(name, ".")
	If i>0 Then Z_GetExt = UCase(Right(name, j-i)) Else Z_GetExt = ""
	Z_GetExt = UCase(Z_GetExt)
End Function

Function Z_GetPath(name)
	Dim i, j
	If Right(name, 1) = "\" Then name = Left(name, Len(name)-1)
	i = InStrRev(name, "\")
	If i>0 Then Z_GetPath = LCase(Left(name, i)) Else Z_GetPath = LCase(name)
End Function

Function Z_GetFilename(name)
	Dim i, j
	j = Len(name)
	i = InStrRev(name, "\")
	If i > 0 Then Z_GetFilename = Right(name, j-i) Else Z_GetFilename = name
End Function

Function Z_FormatNumber(strN, Decimals)
	Dim strTmp
	Z_FormatNumber = ""
	If IsNull(Decimals) Then Exit Function
	If Not IsNumeric(Decimals) Then Exit Function
	If Not IsNull(strN) Then
		strN = Trim(strN)
		If Trim(strN) <> "" Then
			If IsNumeric(strN) Then
				Z_FormatNumber = FormatNumber(strN, Decimals, -1, -1, -1)
			Else
				Z_FormatNumber = strN
			End If
		End If
	End If
End Function

Function Z_FormatNumberNC(strN, Decimals)
	Dim strTmp
	Z_FormatNumberNC = ""
	If IsNull(Decimals) Then Exit Function
	If Not IsNumeric(Decimals) Then Exit Function
	If Not IsNull(strN) Then
		strN = Trim(strN)
		If Trim(strN) <> "" Then
			If IsNumeric(strN) Then
				Z_FormatNumberNC = FormatNumber(strN, Decimals, -1, -1, 0)
			Else
				Z_FormatNumberNC = strN
			End If
		End If
	End If
End Function


Function Z_MapMime(strM)
	strM = UCase(strM)
	Select Case strM
		Case "PDF"
			Z_MapMime = "application/PDF"
		Case "DOC"
		Case "DOT"
			Z_MapMime = "application/msword"
		Case "XLS"
			Z_MapMime = "application/vnd.ms-excel"
		Case "TAR"
			Z_MapMime = "application/x-tar"
		Case "ZIP"
			Z_MapMime = "application/x-zip-compressed"
		Case "TXT"
			Z_MapMime = "text/plain"
		Case else
			Z_MapMime = "application/x-octetstream"
	End Select
End Function

Function Z_CDbl(var)
	If Not IsNull(var) Then
		If IsNumeric(var) Then
			var = Replace(var," ","")
			Z_CDbl = var
			If Len(var)<=10 then Z_CDbl = CDbl(Replace(var,",",""))
		Else
			Z_CDbl = 0.0
		End If
	Else
		Z_CDbl = 0.0
	End If
End Function

Function Z_CLng(var)
DIM lngI, lngZ, blnLeading, strTmp
	If Not IsNull(var) Then
		If var = "" Then
			Z_CLng = 0
			Exit Function
		End If
		If IsNumeric(var) Then
			var = Replace(var, " ", "")
			If Len(var)<=5 Then
				Z_CLng = CLng(Replace(var, ",", ""))
			Else
				Z_CLng = ""
				blnLeading = True
				lngZ = Len(var)
				For lngI = 1 to lngZ
					strTmp = Mid(var,lngI,1)
					If IsNumeric(strTmp) Then
						If Not blnLeading And strTmp = "0" Then
							Z_CLng = Z_CLng & strTmp
						ElseIf blnLeading and strTmp <> "0" Then
							Z_CLng = Z_CLng & strTmp
							blnLeading = False
						ElseIf Not(blnLeading) Or strTmp <> "0" Then
							Z_CLng = Z_CLng & strTmp
						End If
					Else
						Exit For
					End If
				Next
			End If
		Else
			Z_CLng = 0
		End If
	Else
		Z_CLng = 0
	End If
End Function

Function Replca(totest, backup)
	If IsNull(totest) Then
		Replca = Z_FixNull(backup)
	Else
		If Trim(totest)="" Then
			Replca = Z_FixNull(backup)
		Else
			Replca = Trim(totest)
		End If
	End If
End Function

Function Z_DoEncrypt(strZZ)
	DIM objEncrypt
	If Trim(strZZ) <> "" Then
		Set objEncrypt = Server.CreateObject("ZEnc.ZBlowfish")
		Z_DoEncrypt = objEncrypt.Encrypt3(strZZ)
		Set objEncrypt = Nothing
	Else
		Z_DoEncrypt = ""
	End If
End Function

Function Z_DoDecrypt(strZZ)
	DIM objEncrypt
	If Trim(strZZ) <> "" Then
		Set objEncrypt = Server.CreateObject("ZEnc.ZBlowfish")
		Z_DoDecrypt = objEncrypt.Decrypt3(strZZ)
		Set objEncrypt = Nothing
	Else
		Z_DoDecrypt = ""
	End If
End Function

Function Z_SQLCBool(var)
	If IsNull(var) Then
		Z_SQLCBool = 0
	ElseIf var = "" Then
		Z_SQLCBool = 0
	Else
		If CBool(var) Then
			Z_SQLCBool = 1
		Else
			Z_SQLCBool = 0
		End If
	End If
End Function

Function Z_CleanName(vntN)
DIM	strTmp
	strTmp = Replace(vntN,"""","")
	'strTmp = Replace(strTmp,"'","")
	strTmp = Replace(strTmp,"+","")
	strTmp = Replace(strTmp,"=","")
	strTmp = Replace(strTmp,"\","")
	strTmp = Replace(strTmp,"/","")
	strTmp = Replace(strTmp,"[","")
	strTmp = Replace(strTmp,"]","")
	strTmp = Replace(strTmp,";","")
	strTmp = Replace(strTmp,":","")
	strTmp = Replace(strTmp,"<","")
	strTmp = Replace(strTmp,">","")
	strTmp = Replace(strTmp,"?","")
	strTmp = Replace(strTmp,"|","")
	Z_CleanName = strTmp 
End Function
%>