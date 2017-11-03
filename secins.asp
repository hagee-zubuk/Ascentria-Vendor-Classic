<%language=vbscript%>
<!-- #include file="_Files.asp" -->
<!-- #include file="_Announce.asp" -->
<!-- #include file="_Utils.asp" -->
<%
 Const adOpenForwardOnly = 0
	Const adOpenKeyset      = 1
	Const adOpenDynamic     = 2
	Const adOpenStatic      = 3
	mySheet = "Alphabetical Order"
	my1stCell = "B3"
	myLastCell = "B900"
	my1stCell2 = "A3"
	myLastCell2 = "A900"
	strHeader = "HDR=NO;"
	myXlsFile = secinsPath & "CARRIER CODE LIST.xls"
	Set objExcel = CreateObject( "ADODB.Connection" )
	 objExcel.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
	    myXlsFile & ";Extended Properties=""Excel 8.0;IMEX=1;" & _
	    strHeader & """"
	 Set objRS = CreateObject( "ADODB.Recordset" )
	    strRange = mySheet & "$" & my1stCell & ":" & myLastCell
	    objRS.Open "Select * from [" & strRange & "]", objExcel, adOpenStatic
	 Set objRS2 = CreateObject( "ADODB.Recordset" )
	    strRange2 = mySheet & "$" & my1stCell2 & ":" & myLastCell2
	    objRS2.Open "Select * from [" & strRange2 & "]", objExcel, adOpenStatic
	 i = 0
	 y = 0
	    Do Until objRS.EOF
	
	      '  If IsNull( objRS.Fields(0).Value ) Or Trim( objRS.Fields(0).Value ) = "" Then Exit Do
	
	        For j = 0 To objRS.Fields.Count - 1
	            If Not IsNull( objRS.Fields(j).Value ) Or Trim(objRS.Fields(j).Value) <> "" Then
	            	kulay = "#FFFFFF"
								If Not Z_IsOdd(y) Then kulay = "#F5F5F5"
	           		stroption = stroption & "<tr bgcolor='" & kulay & "' onclick=""PassMe('" & objRS2.Fields(j).Value & "');""><td align='left'>" &  objRS.Fields(j).Value & " (" & objRS2.Fields(j).Value &")</td></tr>" & vbCrLf
	            	'stroption =stroption & "<option value='" & objRS2.Fields(j).Value & "'>" & objRS.Fields(j).Value & "</option>" & vbCrlf
	               'arrData( j, i ) = Trim( objRS.Fields(j).Value )
	              y = y + 1
	            End If
	        Next
	        ' Move to the next row
	        objRS.MoveNext
	        objRS2.MoveNext
	        ' Increment the array "row" number
	        i = i + 1
	    Loop
	 ' Close the file and release the objects
	 	objRS2.Close
    objRS.Close
    objExcel.Close
    Set objRS    = Nothing
    Set objRS2   = Nothing
    Set objExcel = Nothing
%>
<html>
	<head>
		<title>Secondary Insurance</title>
		<link href='style.css' type='text/css' rel='stylesheet'>
		<script language='JavaScript'>
		<!--
		function PassMe(xxx) {
			opener.document.frmMain.selIns.value =  xxx;
			self.close();
		}
		-->
		</script>
	</head>
	<body bgcolor='#FBF5DB' style="width:100%;height:100%;filter: progid:DXImageTransform.Microsoft.gradient(startColorstr=#FFFFFFF, endColorstr=#FBF5DB);" >
		<form method='post' name='frmSec' >
			<table align="center" border="0" width="100%">
				<tr>
					<td class='header' colspan='2'>
						<nobr>Secondary Insurance --&gt&gt
					</td>
				</tr>
				<tr>
					<td>(to select, click on the insurance company)</td>
				</tr>
				<tr>
					<td align="center">
						<table border='0' cellspacing='0' cellpadding='0'>
							<%=stroption%>
						</table>
					</td>
				</tr>
				<tr><td>&nbsp;</td></tr>
				<tr>
					<td align="center">
						
						<input type="button" name="btnClose" value="Close" onclick="self.close();" class='btn' onmouseover="this.className='hovbtn'" onmouseout="this.className='btn'">
						
					</td>
				</tr>
			</table>
		</form>
	</body>
</html>