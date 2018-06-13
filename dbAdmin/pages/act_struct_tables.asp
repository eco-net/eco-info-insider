<!--#include virtual="/shared/sqlcheck.asp"-->
<%
DIM theComm
IF len(request("drop"))=1 THEN
	Set theComm = Server.CreateObject("ADODB.Command")
	theComm.ActiveConnection = MM_ecoinfo_STRING
	'theComm.ActiveConnection = MM_voresDebat_STRING
	theComm.CommandText = "DROP TABLE " & request("tablename")
	theComm.Execute
	theComm.ActiveConnection.Close
ELSEIF len(request("empty"))=1 THEN
	Set theComm = Server.CreateObject("ADODB.Command")
	theComm.ActiveConnection = MM_ecoinfo_STRING
	'theComm.ActiveConnection = MM_voresDebat_STRING
	theComm.CommandText = "DELETE FROM " & request("tablename")
	theComm.Execute
	theComm.ActiveConnection.Close
END IF
%>