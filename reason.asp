<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilCalendar.asp" -->
<%
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
	'SAVE COMMENT
	Set rsCom = Server.CreateObject("ADODB.RecordSet")
	sqlCom = "SELECT * FROM Comment_T WHERE UID = " & Request("tagoID")
	rsCom.Open sqlCom, g_strCONN, 1, 3
	If rsCom.EOF Then
		rsCom.AddNew
		rsCom("UID") = Request("tagoID")
	End If
	rsCom("comment") = Request("txtCom")
	rsCom.Update
	rsCom.Close
	Set rsCom = Nothing
	'DELETE EXISTING REASON
	ctrI = Request("ctrKey")
	For i = 0 to ctrI 
		tmpctr = Request("chk" & i)
		If tmpctr <> "" Then
			Set rsDelReas = Server.CreateObject("ADODB.RecordSet")
			sqlDelReas = "DELETE * FROM Encounter_T WHERE index = " & tmpctr
			rsDelReas.Open sqlDelReas, g_strCONN, 1, 3
			Set rsDelReas = Nothing
		End If
	Next
	'DELETE KEY IF NOT SAME
	Set rsDel = Server.CreateOBject("ADODB.RecordSet")
	sqlDel = "SELECT * FROM Encounter_T WHERE appID = " & Request("tagoID")
	rsDel.Open sqlDel, g_strCONN, 1, 3
	If Not rsDel.EOF Then
		myOKey = rsDel("Key")
	End If
	rsDel.Close
	If Cint(MyOKey) <> Cint(Request("selKey")) Then
		sqlDel = "DELETE * FROM Encounter_T WHERE appID = " & Request("tagoID")
		rsDel.Open sqlDel, g_strCONN, 1, 3
	End If
	Set rsDel = Nothing
	'SAVE KEY AND REASON
	Set rsKey = Server.CreateOBject("ADODB.RecordSet")
	sqlKey = "SELECT * FROM Encounter_T WHERE appID = " & Request("tagoID") & " AND key = " & Request("selKey") & " AND reason = " & Request("selNew")
	rsKey.Open sqlKey, g_strCONN, 1, 3
	If rsKey.EOF Then
		If Request("selNew") <> 0 Then
			rsKey.AddNew
			rsKey("appID") = Request("tagoID")
			rsKey("key") = Request("selKey")
			rsKey("Reason") = Request("selNew")
			rsKey.Update
		End If
	End If
	rsKey.Close
	Set rsKey = Nothing
	'SAVE KEY ON APPT TBL
	Set rsSkey = Server.CreateObject("ADODB.RecordSet")
	sqlSkey = "SELECT * FROM Appointment_T WHERE index = " & Request("tagoID")
	rsSkey.Open sqlSkey, g_strCONN, 1, 3
	If Not rsSkey.EOF Then
		rsSkey("key") = Request("selKey")
		rsSkey.Update
	End If
	rsSkey.Close
	Set rsSkey = Nothing
	'SAVE ACTUAL TIME
	Set rsAtime = Server.CreateObject("ADODB.RecordSet")
	sqlAtime = "SELECT * FROM Appointment_T WHERE index = " & Request("tagoID")
	rsAtime.Open sqlAtime, g_strCONN, 1, 3
	If Not rsAtime.EOF Then
		If Request("txtAST") <> "" then
			rsAtime("AStime") = Request("txtAST")
		Else
			rsAtime("AStime") = Empty
		End If
		If Request("txtAET") <> "" Then
			rsAtime("AEtime") = Request("txtAET")
		Else
			rsAtime("AEtime") = Empty
		End If
		rsAtime.Update
	End If
	rsAtime.Close
	Set rsAtime = Nothing
	'SAVE ON LB ACTUAL TIME
	Set rsAtime = Server.CreateObject("ADODB.RecordSet")
	sqlAtime = "SELECT * FROM request_T WHERE HPID = " & Request("tagoID")
	rsAtime.Open sqlAtime, g_strCONNLB, 1, 3
	If Not rsAtime.EOF Then
		If Request("txtAST") <> "" then
			rsAtime("AStarttime") = Request("txtAST")
		Else
			rsAtime("AStarttime") = Empty
		End If
		If Request("txtAET") <> "" Then
			rsAtime("AEndtime") = Request("txtAET")
		Else
			rsAtime("AEndtime") = Empty
		End If
		rsAtime.Update
	End If
	rsAtime.Close
	Set rsAtime = Nothing
	'SAVE FOLLOW UP TIME and CONFIRMATION
	Set rsFol = Server.CreateObject("ADODB.RecordSet")
	sqlFol = "SELECT * FROM Appointment_T WHERE index = " & Request("tagoID")
	rsFol.Open sqlFol, g_strCONN, 1, 3
	If Not rsFol.EOF Then
		rsFol("Follow") = Z_Czero(Request("txtFollow"))
		rsFol("Confirm") = Z_Czero(Request("txtConfirm"))
		rsFol.Update
	End If
	rsFol.Close
	Set rsFol = Nothing
	tmpID = Request("tagoID")
	tmpkey = Request("selKey")
Else
	tmpRec = Split(Request("strReq"), "|")
	tmpID = tmpRec(0) 
	tmpkey = tmpRec(1)
End If
'GET APP
Set rsApp = Server.CreateObject("ADODB.RecordSet")
sqlApp = "SELECT * FROM appointment_T  WHERE index = " & tmpID
rsApp.Open sqlApp, g_strCONN, 3, 1
If Not rsApp.EOF Then
	tmpStart = Z_FormatTime(rsApp("timeFrom"))
	tmpEnd = Z_FormatTime(rsApp("timeTo"))
	tmpAST = Z_FormatTime(rsApp("AStime"))
	tmpAET = Z_FormatTime(rsApp("AEtime"))
	tmpFollow = rsApp("Follow")
	tmpConfirm = rsApp("Confirm")
	tmpInst = GetFacility(rsApp("InstID"))
	tmpCli = Z_DoDecrypt(rsApp("Clname")) & ", " & Z_DoDecrypt(rsApp("Cfname"))
	key1 = ""
	key2 = ""
	If rsApp("key") = 1 Then key1 = "selected"
	If rsApp("key") = 2 Then key2 = "selected"
	keyKo = rsApp("key")
End If
rsApp.Close
Set rsApp = Nothing
'GET COMPLETED REASON
Set rsCom = Server.CreateObject("ADODB.RecordSet")
sqlCom = "SELECT * FROM complete_T ORDER BY index"
rsCom.Open sqlCom, g_strCONN, 3, 1
Do Until rsCom.EOF
	strCom1 = strCom1 & "var ChoiceCom = document.createElement('option');" & vbCrLf & _
		"ChoiceCom.value = " & rsCom("index") & ";" & vbCrLf & _
		"ChoiceCom.title = """ & rsCom("completeReason") & """;" & vbCrLf & _
		"ChoiceCom.appendChild(document.createTextNode(""" & rsCom("completeReason") & """));" & vbCrLf & _
		"document.frmReas.selNew.appendChild(ChoiceCom); " & vbCrlf
		strCom = strCom & "<option value='" & rsCom("index")  & "'>" & rsCom("completeReason") & "</option>" & vbCrLf
	rsCom.MoveNext
Loop
rsCom.Close
Set rsCom = Nothing
'GET CANCELED REASON
Set rsCan= Server.CreateObject("ADODB.RecordSet")
sqlCan = "SELECT * FROM cancel_T ORDER BY index"
rsCan.Open sqlCan, g_strCONN, 3, 1
Do Until rsCan.EOF
	strCan1 = strCan1 & "var ChoiceCan = document.createElement('option');" & vbCrLf & _
		"ChoiceCan.value = " & rsCan("index") & ";" & vbCrLf & _
		"ChoiceCan.title = """ & rsCan("cancelReason") & """;" & vbCrLf & _
		"ChoiceCan.appendChild(document.createTextNode(""" & rsCan("cancelReason") & """));" & vbCrLf & _
		"document.frmReas.selNew.appendChild(ChoiceCan); " & vbCrlf
	strCan = strCan & "<option value='" & rsCan("index")  & "'>" & rsCan("cancelReason") & "</option>" & vbCrLf
	rsCan.MoveNext
Loop
rsCan.Close
Set rsCan = Nothing
'GET APP REASON
Set rsAR = Server.CreateObject("ADODB.RecordSet")
sqlAR = "SELECT * FROM Encounter_T WHERE appID = " & tmpID & " AND KEY = " & keyKo & " ORDER BY reason"
rsAR.Open sqlAR, g_strCONN, 1, 3
ctrKey = 0
Do Until rsAR.EOF
	tblReas = tblReas & "<tr><td align='center'>" & vbCrLf & _
		"<input type='checkbox' name='chk" & ctrKey & "' value='" & rsAR("index") & "'></td>" & vbCrLf & _
		"<td class='confirm' title='" & GetReason2(rsAR("reason"), rsAR("key")) & "'>" & CutMe(GetReason2(rsAR("reason"), rsAR("key"))) & "</td>"
	rsAR.MoveNext
	ctrKey = ctrKey + 1
Loop
rsAR.Close
Set rsAR = Nothing
'GET COMMENT
Set rsCom = Server.CreateObject("ADODB.RecordSet")
sqlCom = "SELECT * FROM Comment_T WHERE UID = " & tmpID
rsCom.Open sqlCom, g_strCONN, 1, 3
If Not rsCom.EOF Then
	tmpCom = rsCom("comment")
End If
rsCom.Close
Set rsCom = Nothing
%>
<html>
	<head>
		<title>Interpreter Request - Edit Reason</title>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
		<!--
		function IsNumeric(sText)
		{
		    var ValidChars = "0123456789.";
		    var IsNumber=true;
		    var Char;
			for (i = 0; i < sText.length && IsNumber == true; i++) 
		     { 
			      Char = sText.charAt(i); 
			      if (ValidChars.indexOf(Char) == -1) 
			      {
			      	IsNumber = false;
			      }
		     }
		 	return IsNumber;
		 }
		function IsTime(strTime)
		{
			var strTime3 = /^(\d{1,2})(\:)(\d{1,2})$/;
			var strFormat3 = strTime.match(strTime3);
			if (strFormat3 != null)
			{
				if (strFormat3[1] > 23 || strFormat3[1] < 00)
				{
                  		return false;
                  	}
				if (strFormat3[3] > 59 || strFormat3[3] < 00) 
				{
                   	return false;	
                   }
              }
              return true;
		}
		function maskMe(str,textbox,loc,delim)
		{
			var locs = loc.split(',');
			for (var i = 0; i <= locs.length; i++)
			{
				for (var k = 0; k <= str.length; k++)
				{
					 if (k == locs[i])
					 {
						if (str.substring(k, k+1) != delim)
					 	{
					 		str = str.substring(0,k) + delim + str.substring(k,str.length);
		     			}
					}
				}
		 	}
			textbox.value = str
		}
		function KeyChoice(tmpKey)
		{
			var i, tmpKey;
			document.frmReas.selNew.length = 1;
			if (tmpKey == 1)
			{
				<%=strCom1%>
			}
			if (tmpKey == 2)
			{
				<%=strCan1%>
			}
		}
		function SaveMe(xxx)
		{
			if (IsNumeric(document.frmReas.txtConfirm.value) == false)
			{
				alert("ERROR: Invalid Confirmation call.")
				return;
			}
			if (IsNumeric(document.frmReas.txtFollow.value) == false)
			{
				alert("ERROR: Invalid Follow-up call.")
				return;
			}
			if (IsTime(document.frmReas.txtAST.value) == false)
			{
				alert("ERROR: Invalid Actual Start time.")
				return;
			}
			if (IsTime(document.frmReas.txtAET.value) == false)
			{
				alert("ERROR: Invalid Actual End time.")
				return;
			}
			var ctrKey = <%=ctrKey%>;
			if (document.frmReas.selNew.value == 0 && document.frmReas.selKey.value != 0 && ctrKey == 0)
			{
				alert("ERROR: At least one reason is needed.")
				return;
			}
			window.opener.document.frmEncounter.HidCom.value = document.frmReas.txtCom.value;
			window.opener.document.frmEncounter.HideID.value = document.frmReas.tagoID.value;
			window.opener.document.frmEncounter.HideKey.value = document.frmReas.selKey.value;
			window.opener.document.frmEncounter.HideReas.value = document.frmReas.selNew.value;
			window.opener.document.frmEncounter.HideAST.value = document.frmReas.txtAST.value;
			window.opener.document.frmEncounter.HideAET.value = document.frmReas.txtAET.value;
			window.opener.document.frmEncounter.HideFollow.value = document.frmReas.txtFollow.value;
			window.opener.document.frmEncounter.HideConfirm.value = document.frmReas.txtConfirm.value;
			window.opener.SubmitAko();
			self.close();
		}
		function JustSaveMe(appID, appKey)
		{
			if (IsNumeric(document.frmReas.txtConfirm.value) == false)
			{
				alert("ERROR: Invalid Confirmation call.")
				return;
			}
			if (IsNumeric(document.frmReas.txtFollow.value) == false)
			{
				alert("ERROR: Invalid Follow-up call.")
				return;
			}
			if (IsTime(document.frmReas.txtAST.value) == false)
			{
				alert("ERROR: Invalid Actual Start time.")
				return;
			}
			if (IsTime(document.frmReas.txtAET.value) == false)
			{
				alert("ERROR: Invalid Actual End time.")
				return;
			}
			var ctrKey = <%=ctrKey%>;
			if (document.frmReas.selNew.value == 0 && document.frmReas.selKey.value != 0 && ctrKey == 0)
			{
				alert("ERROR: At least one reason is needed.")
				return;
			}
			var strReq = appID + "|" + appKey;
			document.frmReas.action = "reason.asp?strReq=" + strReq ;
			document.frmReas.submit();
		}
		-->
		</script>
	</head>
	<body bgcolor='#FBF5DB' style="width:100%;height:100%;filter: progid:DXImageTransform.Microsoft.gradient(startColorstr=#FFFFFFF, endColorstr=#FBF5DB);"
		onload='KeyChoice(<%=tmpkey%>);'>
		<form method='post' name='frmReas' action='reason.asp'>
			<table cellpadding='2' cellspacing='2' border='0' align='left' width='100%'>
				<tr bgcolor='#336601'><td colspan='2'>&nbsp;</td></tr>
				<tr><td>&nbsp;</td></tr>
				<tr>
					<td align='right' width='75px'>Institution:</td>
					<td class='confirm'><%=tmpInst%></td>
				</tr>
				<tr>
					<td align='right'>Client:</td>
					<td class='confirm'><%=tmpCli%></td>
				</tr>
				<tr>
					<td align='right'>Time:</td>
					<td class='confirm'><%=tmpStart%> - <%=tmpEnd%></td>
				</tr>
				<tr>
					<td align='right'>Actual Time:</td>
					<td class='confirm'>
						<input class='main' size='5' maxlength='5' name='txtAST' onKeyUp="javascript:return maskMe(this.value,this,'2,6',':');" onBlur="javascript:return maskMe(this.value,this,'2,6',':');" value='<%=tmpAST%>'>
						&nbsp;-&nbsp;
						<input class='main' size='5' maxlength='5' name='txtAET' onKeyUp="javascript:return maskMe(this.value,this,'2,6',':');" onBlur="javascript:return maskMe(this.value,this,'2,6',':');" value='<%=tmpAET%>'>
						<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">24-hour format</span>
					</td>
				</tr>
				<tr>
					<td align='right'>Confirmation Call:</td>
					<td><input class='main' size='6' maxlength='5' name='txtConfirm' value='<%=tmpConfirm%>'></td>
				</tr>
				<tr>
					<td align='right'>Follow-up Call:</td>
					<td><input class='main' size='6' maxlength='5' name='txtFollow' value='<%=tmpFollow%>'></td>
				</tr>
				<tr><td>&nbsp;</td></tr>
				<tr><td colspan='3'><hr align='center' width='75%'></td></tr>
				<tr><td>&nbsp;</td></tr>
				<tr>
					<td align='right'>Key:</td>
					<td class='confirm'>
						<select  class='seltxt' style='width: 77px;' name='selKey' onchange='KeyChoice(this.value);'>
							<option value='0' >&nbsp;</option>
							<option <%=key1%> value='1'>Completed</option>
							<option <%=key2%>  value='2'>Canceled</option>
						</select>	
					</td>
				</tr>
				<tr><td>&nbsp;</td></tr>
				<tr>
					<td align='right' width='75px'>Reason:</td>
					<td class='confirm'>
						<select name='selNew' class='seltxt' style="position: relative; OVERFLOW: auto;">
							<option value='0'>&nbsp;</option>
						</select>
						<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">*NEW</span>
					</td>
				</tr>
				<tr>
					<td>&nbsp;</td>
					<td>
						<table cellspacing='0' cellpadding='0' border='0'>
							<tr>
								<td colspan='2' align='left'>
									<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">&nbsp;*Click checkbox and click on SAVE button to delete reason</span>
								</td>
							</tr>
							<%=tblReas%>
						</table>
					</td>
				</tr>
				<tr><td>&nbsp;</td></tr>
				<tr><td colspan='3'><hr align='center' width='75%'></td></tr>
				<tr><td>&nbsp;</td></tr>
				<tr>
					<td align='right' valign='top' width='75px'>Comment:</td>
					<td class='confirm'>
						<textarea name='txtCom' class='main' style='width: 250px;'><%=tmpCom%></textarea>
					</td>
				</tr>
				<tr><td>&nbsp;</td></tr>
				<tr><td colspan='3'><hr align='center' width='75%'></td></tr>
				<tr><td>&nbsp;</td></tr>
				<tr>
					<td align='center' colspan='2'>
						<input class='btn' type='button' value='Save' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='JustSaveMe(<%=tmpID%>);'>
						<input class='btn' type='button' value='Save and Exit' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick='SaveMe(<%=tmpID%>, <%=tmpKey%>);'>
						<input type='hidden' name='tagoID' value='<%=tmpID%>'>
						<input type='hidden' name='ctrKey' value='<%=ctrKey%>'>
						<input type='hidden' name='MyKey' value='document.frmReas.selNew.value;'>
						<input type='hidden' name='MyOKey' value='document.frmReas.selNew.value;'>
					</td>
				</tr>
			</table> 	
		</form>
	</body>
</html>