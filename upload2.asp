<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<%
	h_filename = Request("hfname")
	lngRID = Request("RID")
	disUpload = ""

	If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
		Set oUpload = Server.CreateObject("SCUpload.Upload")
		oUpload.Upload
		lngRID	= oUpload.Form("rid")
		If oUpload.Files.Count = 0 Then
			Set oUpload = Nothing
			Session("MSG") = "Please specify a file to import."
			Response.Redirect "upload2.asp"
		End If
		oFileName = oUpload.Files(1).Item(1).filename
		If Z_GetExt(oFileName) <> "PDF" Then
			Set oUpload = Nothing
			Session("MSG") = "Invalid File."
			Response.Redirect "upload2.asp"
		End If
		' REMOVED -- 2019-03-24 -- no file size limits in code!'
		'oFileSize = oUpload.Files(1).Item(1).Size
		'If oFileSize > 2097152 Then
		''	Set oUpload = Nothing
		''	Session("MSG") = "File is too large."
		''	Response.Redirect "upload.asp"
		'End If
		nFileName = oUpload.Form("h_filename") & ".PDF"
		oUpload.Files(1).Item(1).Save F604AStr, nFileName
		Set oUpload = Nothing

		Set rsTBL = Server.CreateObject("ADODB.RecordSet")
		sqlTBL = "SELECT [index], [approvepdf], [uploadfile], [filename] FROM [appointment_T] WHERE [index] = " & lngRID
		rsTBL.Open sqlTBL, g_strCONN, 1, 3
		If Not rsTBL.EOF Then
			rsTBL("UploadFile") = True
			rsTBL("approvepdf") = False
			rsTBL("filename") = nFileName & ".PDF"
			rsTBL.Update
		End If
		rsTBL.Close
		Set rsTBL = Nothing

		Set rsTBL = Server.CreateObject("ADODB.RecordSet")
		sqlTBL = "SELECT [index], [approvepdf], [uploadfile], [filename] FROM Request_T WHERE [HPID] = " & lngRID
		rsTBL.Open sqlTBL, g_strCONNLB, 1, 3
		If Not rsTBL.EOF Then
			rsTBL("UploadFile") = True
			rsTBL("approvepdf") = False
			rsTBL("filename") = nFileName & ".PDF"
			rsTBL.Update
		End If
		rsTBL.Close
		Set rsTBL = Nothing

		Session("MSG") = "File Saved."
		disUpload = "disabled"
	End If
%>
<html>
	<head>
		<title>Language Bank - Upload</title>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
		<!--
		function closeWin() {
			window.opener.location.reload();
			self.close();
		}
		function uploadFile() {
			if (document.frmUpload.F1.value != "") {
				filestr = document.frmUpload.F1.value.toUpperCase();
				if (filestr.indexOf(".PDF") == -1) {
					alert("ERROR: Incorrect file extension.")
					document.frmUpload.F1.value = "";
					return;
				}
				else {

					document.frmUpload.action = "upload2.asp";
					document.frmUpload.submit();
				}
			}
			else {
				alert("ERROR: Please select a file.")
				return;
			}
		}
		-->
		</script>
	</head>
	<body bgcolor='#FBF5DB' style="width:100%;height:100%;filter: progid:DXImageTransform.Microsoft.gradient(startColorstr=#FFFFFFF, endColorstr=#FBF5DB);" >
		<form method='post' name='frmUpload' enctype="multipart/form-data">
			<input type="hidden" name="RID" id="RID" value="<%=lngRID%>" />
			<table align="center" border="0" width="100%">
				<tr>
					<td class='header' colspan='2'>
						<nobr>FORM 604A --&gt&gt
					</td>
				</tr>
				<tr>
					<td align='center' colspan='10'><span class='error'><%=Session("MSG")%></span></td>
				</tr>
<% If disUpload <> "disabled" Then %>
				<tr>
					<td align="center">
						<input  class='main' type="file" name="F1" size="30" class='btn'>
					</td>
				</tr>
				<tr>
					<td align="left">
						&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">*PDF format only</span><br>
						&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						<span class='formatsmall' onmouseover="this.className='formatbig'" onmouseout="this.className='formatsmall'">*2 MB limit</span>
					</td>
				</tr>
<% End If %>				
				<tr><td>&nbsp;</td></tr>
				<tr>
					<td align="center">
<% If disUpload <> "disabled" Then %>
						<input type="button" name="btnUp" value="UPLOAD" onclick="uploadFile();" <%=disUpload%> class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'">
<% End If %>						
						<input type="button" name="btnClose" value="Close" onclick="closeWin();" class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'">
						<input type="hidden" name="h_filename" value="<%=h_filename%>">
					</td>
				</tr>
			</table>
		</form>
	</body>
</html>
<%
If Session("MSG") <> "" Then
	tmpMSG = Replace(Session("MSG"), "<br>", "\n")
%>
<%
' <script><!--
' 	alert("__tmpMSG__");
' --></script>
'
End If
Session("MSG") = ""
%>
