<!--#include virtual="/Connections/ecoinfo.asp" -->
<%
set rsTables = Server.CreateObject("ADODB.Recordset")
rsTables.ActiveConnection = MM_ecoinfo_STRING
rsTables.Source = "SELECT *  FROM INFORMATION_SCHEMA.TABLES"
rsTables.CursorType = 0
rsTables.CursorLocation = 2
rsTables.LockType = 3
rsTables.Open()
i=0
while not rsTables.EOF 
i=i+1
response.write i & ": " & rsTables("TABLE_NAME") & "<br>"
rsTables.MoveNext
wend
rsTables.Close
%>