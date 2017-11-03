<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilCalendar.asp" -->
<%
dim HidID, BigCtr
BigCtr = 0
HidID = ""
If Request.ServerVariables("REQUEST_METHOD") = "POST" Or Request("NewKey") = 1 Then
	tmpDate = Request("txtAppDate")
	If Session("type") <> 1 Then
		If Request("selIntr") <> 0 Then
			HidID = Request("selIntr")
		Else
			HidID = Request("IID")
		End If
	End If
End If
If tmpDate = "" Then tmpDate = Date
If Session("type") <> 1 Or Session("type") = 3 Then
	Set rsInst = Server.CreateObject("ADODB.RecordSet")
	sqlInst = "SELECT * FROM User_T  WHERE [index] = " & Session("UID")
	rsInst.Open sqlInst, g_strCONN, 3, 1
	If Not rsInst.EOF Then
		If tmpID = "" Then tmpID = rsInst("InstID")
		tmpInst = rsInst("instID")
		tmpDept = rsInst("DeptLB")
	End If
	rsInst.Close
	Set rsInst = Nothing
ElseIf Session("type") = 1 then
	Set rsIntr = Server.CreateObject("ADODB.RecordSet")
	sqlInst = "SELECT * FROM User_T WHERE [index] = " & Session("UID") 
	rsIntr.Open sqlInst, g_strCONN, 3, 1
	If Not rsIntr.EOF Then
		If tmpID = "" Then  tmpID = rsIntr("intrLB")
	End If
	rsIntr.Close
	Set rsIntr = Nothing
End If
'GET INTERPRETER
validIntr = ""
If Session("type") = 1 Then validIntr = "disabled"
validIntr2 = "disabled"
If Session("type") = 1 Then validIntr2 = ""
IntrName = "---ALL---"
If Session("type") = 1 And Z_CZero(HidID) = 0 Then 
	IntrName = UCase(GetIntrNameEnc(GetHPID(Session("UID"))))
	HidID = Session("IntrID")
ElseIf Z_CZero(HidID) = 0 Then
	HidID = 999
ElseIf  Z_CZero(HidID) <> 999 Then
	IntrName = UCase(GetIntrNameEnc2(HidID))
End If
Set rsIntrInst = Server.CreateObject("ADODB.RecordSet")
If Session("type") = 1 then
	sqlIntrInst = "SELECT * FROM InstIntr_T WHERE IntrID = " & tmpID
Else	
	sqlIntrInst = "SELECT * FROM InstIntr_T WHERE InstID = " & tmpID
End If
rsIntrInst.Open sqlIntrInst, g_strCONN, 3, 1
Do Until rsIntrInst.EOF
	Set rsIntr = Server.CreateObject("ADODB.RecordSet")
	sqlIntr = "SELECT * FROM Interpreter_T WHERE [index] = " & rsIntrInst("IntrID")
	rsIntr.Open sqlIntr, g_strCONNLB, 3, 1
	If Not rsIntr.EOF Then
		tmpIntrName = rsIntr("last name") & ", " &rsIntr("first name")
		SelMe = ""
		If HidID = rsIntr("index") Then 
			SelMe = "selected"
			IntrName = tmpIntrName
		End If
		strIntr = strIntr & "<option value='" & rsIntr("index") & "' " & SelMe & ">" & tmpIntrName & "</option>" & vbCrLf
	End If
	rsIntr.Close
	Set rsIntr = Nothing
	rsIntrInst.MoveNext
Loop
rsIntrInst.Close
Set rsIntrInst =Nothing

'GET APPOINTMENTS
readLang = "readonly"
disabledSya = "disabled"
If Session("type") = 1 Then 
	readLang = ""
	disabledSya = ""
End If
Set rsApp = Server.CreateObject("ADODB.RecordSet")
If Z_CZero(HidID) <> 0 And Z_CZero(HidID) <> 999 Then
	If Session("type") = 0 Then
		sqlApp = "SELECT * FROM Appointment_T WHERE IntrID = " & HidID & " AND InstID = " & tmpID & " AND AppDate = '" & tmpDate & "' ORDER BY timeFrom"
	ElseIf Session("type") = 3 Then
		sqlApp = "SELECT * FROM Appointment_T WHERE IntrID = " & HidID & " AND InstID = " & tmpID & " AND AND deptID = " & tmpDept & " AND AppDate = '" & tmpDate & "' ORDER BY timeFrom"
	Else
		sqlApp = "SELECT * FROM Appointment_T WHERE  IntrID = " & HidID & " AND AppDate = '" & tmpDate & "' ORDER BY timeFrom"
	End If
ElseIf Z_CZero(HidID) = 999 Then
	sqlApp = "SELECT * FROM Appointment_T WHERE  InstID = " & tmpID & " AND AppDate = '" & tmpDate & "' ORDER BY timeFrom"
End If
rsApp.Open sqlApp, g_strCONN, 3, 1
ctr = 0
Do Until rsApp.EOF
	'GET KEY
	keyKo = ""
	If rsApp("key") = 1 Then keyKo = "Completed"
	If rsApp("key") = 2 Then keyKo = "Canceled"
	tmpStr = tmpStr & "<tr bgcolor='#F5F5F5' >" & vbCrLf & _
		"<td align='center'><b>"& Z_FormatTime(rsApp("timeFrom")) & "</b></td>" & vbCrLf &  _
		"<td align='center'><b>" & Z_FormatTime(rsApp("timeTo")) & "</b></td>" & vbCrLf &  _
		"<td align='center'><b>" & Z_FormatTime(rsApp("AStime")) & "</b></td>" & vbCrLf &  _
		"<td align='center'><b>" & Z_FormatTime(rsApp("AEtime")) & "</b></td>" & vbCrLf &  _
		"<td align='center'><b>" & Z_FormatTime(rsApp("paged")) & "</b></td>" & vbCrLf &  _
		"<td align='center'><b>" & Z_DoDecrypt(rsApp("Clname")) & ", " & Z_DoDecrypt(rsApp("Cfname")) & "</b></td>" & vbCrLf &  _
		"<td align='center'><b>" & GetFacility(rsApp("InstID")) &  "</b></td>" & vbCrLf & _
		"<td align='center'><b>" & rsApp("clinician") & "</b></td>" & vbCrLf &  _
		"<td align='center'><b>" & rsApp("confirm") & "</b></td>" & vbCrLf &  _
		"<td align='center'><b>" & rsApp("follow") & "</b></td>" & vbCrLf &  _
		"<td align='center'><textarea name='txtcom' class='main' style='width: 150px;' readonly rows='2'>" & GetComment(rsApp("index")) &  "</textarea></td>" & vbCrLf & _
		"<td align='center'><b>" & keyKo &  "</b></td>" & vbCrLf
	'GET REASON
	Set rsReas = Server.CreateObject("ADODB.RecordSet")
	sqlReas = "SELECT * FROM Encounter_T WHERE appID = " & rsApp("index") & " ORDER BY reason"
	rsReas.Open sqlReas, g_strCONN, 3, 1
	rCtr = 14
	Do Until rsReas.EOF
		tmpStr = tmpStr & "<td width='20px' align='center' title='" & GetReason(rsApp("key"), rsReas("reason")) & "'><b>" & rsReas("reason") & "</b></td>" & vbCrLf
		rsReas.MoveNext
		rCtr = rCtr - 1
	Loop
	rsReas.Close
	Set rsReas = Nothing
	'SPACE FILLER
	If rCtr <> 0 Then tmpStr = tmpStr & "<td style='width: 20px;' align='center' colspan='" & rCtr & "'>&nbsp;</td>" & vbCrLf
	'EDIT REASON
	tmpStr = tmpStr & "<td width='40px' align='center'><input " & disabledSya & " type='button' name='btnEdit" & ctr & "' onclick='EditReas(" & rsApp("index") & ", " & rsApp("key") & ");' class='btn' value='...'  onmouseover=""this.className='hovbtn'"" onmouseout=""this.className='btn'"" style='width: 19px;' title='Edit Reason/Comment'></td>" & vbCrLf
	tmpStr = tmpStr & 	"</tr>" & vbCrLf
	ctr = ctr + 1
	rsApp.MoveNext
Loop
rsApp.Close
Set rsApp = Nothing
%>
<html>
	<head>
		<title>Interpreter Request - Encounter Form</title>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<link href="CalendarControl.css" type="text/css" rel="stylesheet">
		<script src="CalendarControl.js" language="javascript"></script>
		<script language='JavaScript'>
		<!--
		function SelectMe(xxx)
		{
			document.frmEncounter.action = 'encounter.asp';
			document.frmEncounter.submit();
		}
		function CalendarView(strIntr)
		{
			document.frmEncounter.action = 'encounter.asp?IID=' + strIntr;
			document.frmEncounter.submit();
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
		function EditReas(appID, appKey)
		{
			strReq = appID + "|" + appKey;
			//var myBrow = navigator.appName;
			//if (myBrow == "Netscape")
			//{
				var newwindow = window.open('reason.asp?strReq=' + strReq,'name','width=400,scrollbars=1,directories=0,status=0,toolbar=0,resizable=0,dependent=1,modal=yes');
			//}
			//else
			//{
			//	var newwindow = window.showModalDialog('reason.asp?strReq=' + strReq,'name','height=400,width=400,scrollbars=0,directories=0,status=0,toolbar=0,resizable=0');
			//}
			if (window.focus) {newwindow.focus()}
		}
		function SubmitAko()
		{
			
			document.frmEncounter.action = 'action.asp?ctrl=2';
			document.frmEncounter.submit();
		}
		//-->
		</script>
		<style type="text/css">
	 	.container
	      {
	          border: solid 1px black;
	          overflow: auto;
	      }
	      .noscroll
	      {
	          position: relative;
	          background-color: white;
	          top: expression(this.offsetParent.scrollTop);
	      }
	      th
	      {
	          text-align: left;
	      }
		</style>
	</head>
	<body>
		<table cellSpacing='0' cellPadding='0' height='100%' width="100%" border='0' class='bgstyle2'>
			<tr>
				<td height='100px'>
					<!-- #include file="_header.asp" -->
				</td>
			</tr>
			<tr>
				<td valign='top' >
					<form name='frmEncounter' method='post' action='encounter.asp'>
						<table cellSpacing='2' cellPadding='0' width="100%" border='0'>
							<!-- #include file="_greetme.asp" -->
							<tr>
								<td align='left'>
									Interpreter:
									<span class='HighLight'><%=IntrName%></span>
									&nbsp;&nbsp;
									<select class='seltxt' name='selIntr'  <%=validIntr%> style='width:100px;' onchange='JavaScript:SelectMe(this.value);'>
										<option value='0'></option>
										<option value='999'>--- ALL ---</option>
										<%=strIntr%>
									</select>
									<input type='hidden' name='HIID' value='<%=HidID%>'>
									<input type='hidden' name='HidReas'>
									<input type='hidden' name='HidCom'>
									<input type='hidden' name='HideID'>
									<input type='hidden' name='HideKey'>
									<input type='hidden' name='HideReas'>
									<input type='hidden' name='HideAST'>
									<input type='hidden' name='HideAET'>
									<input type='hidden' name='HideFollow'>
									<input type='hidden' name='HideConfirm'>
								</td>
								<td align='center'>
									Date:
									<input class='main' size='10' maxlength='10' name='txtAppDate'  readonly value='<%=tmpDate%>' onfocus='JavaScript:CalendarView(document.frmEncounter.HIID.value);'>
									<input type="button" value="..." title='Calendar' name="cal1" style="width: 19px;"
									onclick="showCalendarControl(document.frmEncounter.txtAppDate);" class='btnLnk' onmouseover="this.className='hovbtnLnk'" onmouseout="this.className='btnLnk'">
								</td>
							</tr>
							<tr>
								<td colspan='3'>
									<div class='container' style='height: 440px; width:100%; position: relative;'>
										<table cellSpacing='2' cellPadding='0' width='100%' border='0' align='left' bgcolor='#FFFFFF'>
											<thead>
												<tr bgcolor='#D4D0C8' class="noscroll">
													<td align='center'  height='30px' class='time2'>Start Time</td>
													<td align='center'  class='time2'>End Time</td>
													<td align='center'  class='time2'>Actual Start Time</td>
													<td align='center'  class='time2'>Actual End Time</td>
													<td align='center'  class='time2'>Time Paged</td>
													<td align='center'  class='time2'>Client</td>
													<td align='center'  class='time2'>Institution</td>
													<td align='center'  class='time2'>Clinician Name</td>
													<td align='center'  class='time2'>Confirmation Call</td>
													<td align='center'  class='time2'>Follow-up Call</td>
													<td align='center'  class='time2'>Comment</td>
													<td align='center'  class='time2'>Key</td>
													<td align='center'  class='time2' colspan='14'>Reason</td>
													<td align='center'  class='time2'>&nbsp;</td>
												</tr>
											</thead>
											<tbody style="OVERFLOW: auto;">
												<%=tmpStr%>
											</tbody>
										</table>
									</div>
								</td>
							</tr>
						</table>
					</form>
				</td>
			<tr>
				<td valign='bottom'>
					<!-- #include file="_footer.asp" -->
				</td>
			</tr>
		</table>
	</body>
</html>	
<%
If Session("MSG") <> "" Then
	tmpMSG = Replace(Session("MSG"), "<br>", "\n")
%>
<script><!--
	alert("<%=tmpMSG%>");
--></script>
<%
End If
Session("MSG") = ""
%>