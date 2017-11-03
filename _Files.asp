<%
DIM 	g_strCONN, g_strDBPath

'g_strDBPath = "C:\work\InterReq\db\interpreter.mdb"
'g_strCONN = "PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=" & g_strDBPath & ";"

'HOSPITAL PILOT
g_strDBPath = "Provider=SQLOLEDB;Data Source=ERNIE\SQLEXPRESS;Initial Catalog=interpreterSQL;Integrated Security=SSPI;"
Set g_strCONN = Server.CreateObject("ADODB.Connection")
g_strCONN.Open g_strDBPath

'FOR LANGUAGE BANK
g_strCONNDB = "Provider=SQLOLEDB;Data Source=ERNIE\SQLEXPRESS;Initial Catalog=langbank;Integrated Security=SSPI;"
Set g_strCONNLB = Server.CreateObject("ADODB.Connection")
g_strCONNLB.Open g_strCONNDB

BackupStr = "C:\WORK\ascentria\vendor\CSV\"

HistoryDB = "Provider=SQLOLEDB;Data Source=ERNIE\SQLEXPRESS;Initial Catalog=HistLangBank;Integrated Security=SSPI;"'"Provider=SQLOLEDB;Data 
Set g_strCONNHist = Server.CreateObject("ADODB.Connection")
g_strCONNHist.Open HistoryDB

'HIST SQL
g_strCONNDB2 = "Provider=SQLOLEDB;Data Source=ERNIE\SQLEXPRESS;Initial Catalog=histLB;Integrated Security=SSPI;"
Set g_strCONNHIST2 = Server.CreateObject("ADODB.Connection")
g_strCONNHIST2.Open g_strCONNDB2

'PUBLIC DEFENDER
F604AStr = "C:\WORK\ascentria\vendor\F604A\" '"\\webserv6\F604A\"
'Court List
crtLst = "C:\WORK\ascentria\vendor\crtlst.txt"
'Secondary insurance
secinsPath = "C:\WORK\ascentria\vendor\insurance\"
%>