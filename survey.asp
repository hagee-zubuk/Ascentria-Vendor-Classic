<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<!-- #include file="_Security.asp" -->
<!-- #include file="_UtilCalendar.asp" -->
<%
Function Z_SurveyDone(xxx)
	Z_SurveyDone = False
	Set rsDone = Server.CreateObject("ADODB.RecordSet")
	rsDone.Open "SELECT appID FROM Survey_T WHERE appID = " & xxx, g_strCONN, 3, 1
	If Not rsDone.EOF Then Z_SurveyDone = True
	rsDone.Close
	Set rsDone = Nothing
End Function
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
	Set rsSurv = Server.CreateObject("ADODB.RecordSet")
	rsSurv.Open "SELECT * FROM Survey_T WHERE [timestamp] = '" & Now & "'", g_strCONN, 1, 3
	If rsSurv.EOF Then
		rsSurv.AddNew
		rsSurv("timestamp") = Now
		rsSurv("appID") = Request("ID")
		rsSurv("IntrID") = Request("IntrID")
		rsSurv("q1") = Request("q1")
		rsSurv("q2") = Request("q2")
		rsSurv("q3") = Request("q3")
		rsSurv("q4") = Request("q4")
		rsSurv("q5") = Request("q5")
		rsSurv("q6") = Request("q6")
		rsSurv("q7") = Request("q7")
		rsSurv("q8") = Request("q8")
		rsSurv("q9") = Request("q9")
		rsSurv("q10") = Request("q10")
		rsSurv("q11") = Request("q11")
		rsSurv("qcom1") = Request("txtq1")
		rsSurv("qcom2") = Request("txtq2")
		rsSurv("qcom3") = Request("txtq3")
		rsSurv("qcom4") = Request("txtq4")
		rsSurv("qcom5") = Request("txtq5")
		rsSurv("qcom6") = Request("txtq6")
		rsSurv("qcom7") = Request("txtq7")
		rsSurv("qcom8") = Request("txtq8")
		rsSurv("lname") = Request("txtlname")
		rsSurv("fname") = Request("txtfname")
		rsSurv("phone") = Request("txtphone")
		rsSurv("email") = Request("txtemail")
		rsSurv.Update
	End If
	rsSurv.Close
	Set rsSurv = Nothing
	Session("MSG") = "Survey Saved."
End If
disable = 0
If Z_SurveyDone(Request("ID")) Then disable = 1
Set rsApp = Server.CreateObject("ADODB.RecordSet")
rsApp.Open "SELECT [index] AS myID, intrID FROM appointment_T WHERE [index] = " & Request("ID"), g_strCONN, 3, 1
If Not rsApp.EOF Then
	myID = rsApp("myID")
	tmpIntrID = rsApp("intrID")
End If
rsApp.Close
Set rsApp = Nothing
If Z_CZero(tmpIntrID) > 0 Then
	Set rsIntr = Server.CreateObject("ADODB.RecordSet")
	sqlIntr = "SELECT [last name], [first name] FROM interpreter_T  WHERE [index] = " & tmpIntrID
	rsIntr.Open sqlIntr, g_strCONNLB, 3, 1
	If Not rsIntr.EOF Then
		tmpIntrName = rsIntr("last name") & ", " & rsIntr("first name")
	End If
	rsIntr.Close
	Set rsIntr = Nothing
End If
%>
<html>
	<head>
		<title>Interpreter Request - Interpreter Feedback</title>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
			function SaveSurvey(xxx) {
				var meron = <%=disable%>;
				if (meron == 0) {
					document.frmSurvey.action = "survey.asp?ID=" + xxx;
					document.frmSurvey.submit();
				}
				else if (meron == 1) {
					alert("Survey already done for this appointment.");
					return;
				}
			}
		</script>
	</head>
	<body bgcolor='#FBF5DB' style="width:100%;height:100%;filter: progid:DXImageTransform.Microsoft.gradient(startColorstr=#FFFFFFF, endColorstr=#FBF5DB);" >
		<form method='POST' name='frmSurvey'>
			<table border=0 style="width:100%;">
				<tr><td>&nbsp;</td></tr>
				<tr>
					<td  colspan="2" align="center">
						<p><b>
							Thank you for your willingness to fill out this short survey.<br>Your opinion is incredibly valuable for us to evaluate our interpreters and ensure we are providing the best service possible.<br>  
						Please rate the performance of the Language Bank interpreter for this appointment  in the following areas? (1 low, 5 high) </b>
						</p>
					</td>
				</tr>
				<tr><td>&nbsp;</td></tr>
				<tr>
					<td align='right' width='200px'>ID:</td>
					<td class='confirm'><%=myID%></td>
				</tr>
				<tr>
					<td align='right' width='25%'>Interpreter:</td>
					<td class='confirm'><%=tmpIntrName%></td>
				</tr>
				<tr><td>&nbsp;</td></tr>
				<tr>
					
					<td colspan="2" align='center'>
						<table cellSpacing='2' cellPadding='0' style="border: solid 1px;">
							<tr>
								<td>&nbsp;</td>
								<td class='confirm' align='center'>1</td>
								<td class='confirm' align='center'>2</td>
								<td class='confirm' align='center'>3</td>
								<td class='confirm' align='center'>4</td>
								<td class='confirm' align='center'>5</td>
								<td class='confirm' align='center'>Comment</td>
							</tr>
							<tr>
								<td class='confirm' align='right'>1)	Introduced himself/herself and role of the interpreter (Hold Pre-session)</td>
								<td class='confirm' align='center' style='background-color: #FFFFCE;'><input type='radio' name='q1' value='1' checked></td>
								<td class='confirm' align='center' style='background-color: #FFFFCE;'><input type='radio' name='q1' value='2'></td>
								<td class='confirm' align='center' style='background-color: #FFFFCE;'><input type='radio' name='q1' value='3'></td>
								<td class='confirm' align='center' style='background-color: #FFFFCE;'><input type='radio' name='q1' value='4'></td>
								<td class='confirm' align='center' style='background-color: #FFFFCE;'><input type='radio' name='q1' value='5'></td>
								<td class='confirm' align='center' style='background-color: #FFFFCE;'>
									<textarea name='txtq1' class='main' onkeyup='bawal(this);' ></textarea>
								</td>
							</tr>
							<tr>
								<td class='confirm' align='right'>2)	Interpreted everything it was said (all the conversations) during the appointment</td>
								<td class='confirm' align='center' ><input type='radio' name='q2' value='1' checked></td>
								<td class='confirm' align='center' ><input type='radio' name='q2' value='2'></td>
								<td class='confirm' align='center' ><input type='radio' name='q2' value='3'></td>
								<td class='confirm' align='center' ><input type='radio' name='q2' value='4'></td>
								<td class='confirm' align='center' ><input type='radio' name='q2' value='5'></td>
								<td class='confirm' align='center' >
									<textarea name='txtq2' class='main' onkeyup='bawal(this);' ></textarea>
								</td>
							</tr>
							<tr>
								<td class='confirm' align='right'>3)	Able to keep up with the pace of communication</td>
								<td class='confirm' align='center' style='background-color: #FFFFCE;'><input type='radio' name='q3' value='1' checked></td>
								<td class='confirm' align='center' style='background-color: #FFFFCE;'><input type='radio' name='q3' value='2'></td>
								<td class='confirm' align='center' style='background-color: #FFFFCE;'><input type='radio' name='q3' value='3'></td>
								<td class='confirm' align='center' style='background-color: #FFFFCE;'><input type='radio' name='q3' value='4'></td>
								<td class='confirm' align='center' style='background-color: #FFFFCE;'><input type='radio' name='q3' value='5'></td>
								<td class='confirm' align='center' style='background-color: #FFFFCE;'>
									<textarea name='txtq3' class='main' onkeyup='bawal(this);' ></textarea>
								</td>
							</tr>
							<tr>
								<td class='confirm' align='right'>4)	Maintained transparency by keeping either party ( provider or LEP client / patient ) in the loop when communicating with the other for clarification</td>
								<td class='confirm' align='center' ><input type='radio' name='q4' value='1' checked></td>
								<td class='confirm' align='center' ><input type='radio' name='q4' value='2'></td>
								<td class='confirm' align='center' ><input type='radio' name='q4' value='3'></td>
								<td class='confirm' align='center' ><input type='radio' name='q4' value='4'></td>
								<td class='confirm' align='center' ><input type='radio' name='q4' value='5'></td>
								<td class='confirm' align='center' >
									<textarea name='txtq4' class='main' onkeyup='bawal(this);' ></textarea>
								</td>
							</tr>
							<tr>
								<td class='confirm' align='right'>5)	Used the first person while interpreting</td>
								<td class='confirm' align='center' style='background-color: #FFFFCE;'><input type='radio' name='q5' value='1' checked></td>
								<td class='confirm' align='center' style='background-color: #FFFFCE;'><input type='radio' name='q5' value='2'></td>
								<td class='confirm' align='center' style='background-color: #FFFFCE;'><input type='radio' name='q5' value='3'></td>
								<td class='confirm' align='center' style='background-color: #FFFFCE;'><input type='radio' name='q5' value='4'></td>
								<td class='confirm' align='center' style='background-color: #FFFFCE;'><input type='radio' name='q5' value='5'></td>
								<td class='confirm' align='center' style='background-color: #FFFFCE;'>
									<textarea name='txtq5' class='main' onkeyup='bawal(this);' ></textarea>
								</td>
							</tr>
							<tr>
								<td class='confirm' align='right'>6)	Impartiality and boundaries –did not stay alone in room with patient at any time, keeps personal opinions/feelings/believes out of the triadic setting</td>
								<td class='confirm' align='center' ><input type='radio' name='q6' value='1' checked></td>
								<td class='confirm' align='center' ><input type='radio' name='q6' value='2'></td>
								<td class='confirm' align='center' ><input type='radio' name='q6' value='3'></td>
								<td class='confirm' align='center' ><input type='radio' name='q6' value='4'></td>
								<td class='confirm' align='center' ><input type='radio' name='q6' value='5'></td>
								<td class='confirm' align='center' >
									<textarea name='txtq6' class='main' onkeyup='bawal(this);' ></textarea>
								</td>
							</tr>
							<tr>
								<td class='confirm' align='right'>7)	Professionalism –communicated with provider and others with respect</td>
								<td class='confirm' align='center' style='background-color: #FFFFCE;'><input type='radio' name='q7' value='1' checked></td>
								<td class='confirm' align='center' style='background-color: #FFFFCE;'><input type='radio' name='q7' value='2'></td>
								<td class='confirm' align='center' style='background-color: #FFFFCE;'><input type='radio' name='q7' value='3'></td>
								<td class='confirm' align='center' style='background-color: #FFFFCE;'><input type='radio' name='q7' value='4'></td>
								<td class='confirm' align='center' style='background-color: #FFFFCE;'><input type='radio' name='q7' value='5'></td>
								<td class='confirm' align='center' style='background-color: #FFFFCE;'>
									<textarea name='txtq7' class='main' onkeyup='bawal(this);' ></textarea>
								</td>
							</tr>
							<tr>
								<td class='confirm' align='right'>8)	Was interpreter dressed professionally</td>
								<td align="left" colspan="5">
									<select class='seltxt' style="width: 50px;" name="q8">
										<option value="1">Yes</option>
										<option value="0">No</option>
									</select>
								</td>
								<td class='confirm' align='center'>
									<textarea name='txtq8' class='main' onkeyup='bawal(this);' ></textarea>
								</td>
							</tr>
							<tr>
								<td class='confirm' align='right'>9)	Did interpreter arrive on time</td>
								<td align="left" colspan="6" style='background-color: #FFFFCE;'>
									<select class='seltxt' style="width: 50px;" name="q9">
										<option value="1">Yes</option>
										<option value="0">No</option>
									</select>
								</td>
							</tr>
							<tr>
								<td class='confirm' align='right'>10)	Was interpreting wearing LB badge</td>
								<td align="left" colspan="6">
									<select class='seltxt' style="width: 50px;" name="q10">
										<option value="1">Yes</option>
										<option value="0">No</option>
									</select>
								</td>
							</tr>
							<tr>
								<td class='confirm' align='right'>11)	Please feel free to provide additional comments in regards to this interpreter</td>
								<td class='confirm' align='center' colspan="6" style='background-color: #FFFFCE;'>
									<textarea name='txtq11' class='main' onkeyup='bawal(this);' style="width: 100%;" ></textarea>
								</td>
							</tr>
							<tr>
								<td class='confirm' align='right' valign="top">12)	Please provide your contact information if you would like a follow up to your response</td>
								<td align="left" colspan="6">
									<table cellSpacing='2' cellPadding='0' style="border: solid 1px; width: 100%;">
										<tr>
											<td class="confirm" align='right'>Last Name:</td>
											<td align="left">
												<input type="textbox" class="main" name="txtlname" maxlength="50">
											</td>
										</tr>
										<tr>
											<td class="confirm" align='right'>First Name:</td>
											<td align="left">
												<input type="textbox" class="main" name="txtfname" maxlength="50">
											</td>
										</tr>
										<tr>
											<td class="confirm" align='right'>Phone:</td>
											<td align="left">
												<input type="textbox" class="main" name="txtphone" maxlength="50">
											</td>
										</tr>
											<tr>
											<td class="confirm" align='right'>Email:</td>
											<td align="left">
												<input type="textbox" class="main" name="txtemail" maxlength="100">
											</td>
										</tr>
									</table>
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td  colspan="2" align="center">
						<input class='btn' type='button' style='width: 125px;' value='Save' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="SaveSurvey(<%=myID%>);">
						<input class='btn' type='button' style='width: 125px;' value='Close' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'" onclick="window.close();">
						<input type="hidden" name="intrID" value="<%=tmpIntrID%>">
					</td>
				</tr>
			</table>
		</form>
	</body>
</html>
<%
If Session("MSG") <> "" Then
	tmpMSG = Session("MSG")
%>
<script><!--
	alert("<%=tmpMSG%>");
--></script>
<%
End If
Session("MSG") = ""
%>