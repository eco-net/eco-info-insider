<!--#include file="lib.asp" -->
<%
DIM comm, nodel
IF LEN(request("state"))<1 THEN
' Confirm Delete
	state=1
ELSEIF CInt(request("state"))=1 THEN
' Delete OR Select other
	' Checking for in-use
	set comm = server.createobject("ADODB.recordset")
	comm.activeconnection=MM_ecoinfo_STRING
	comm.source="SELECT id FROM eiba_maindata  WHERE onhomepage=1 AND id=" & request("id")
	comm.open()
	IF NOT comm.EOF THEN
		state=2
	ELSE
		DoSave "",3
		state=3
	END IF
	comm.close()

ELSEIF CInt(request("state"))=2 THEN
' Delete AND Update
	set comm = server.createobject("ADODB.Command")
	comm.activeconnection=MM_ecoinfo_STRING
	comm.commandText="UPDATE eiba_maindata SET onhomepage=1 WHERE id=" & request("id")
	comm.execute()
	comm.activeconnection.close()
	DoSave "",4
	state=4
END IF
%>