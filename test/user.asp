<!--#include virtual="/connections/ecoinfo.asp"-->
<%
set rsTables = Server.CreateObject("ADODB.Recordset")
rsTables.ActiveConnection = MM_ecoinfo_STRING

rsTables.Source = "SELECT * FROM eisys_insiderusers"
rsTables.CursorType = 0
rsTables.CursorLocation = 2
rsTables.LockType = 3
rsTables.Open()

if rsTables("email")="kenneth@valeur.com" then
response.write rsTables("orgid") & "-" & rsTables("email")
else
response.write "Dont exist"
end if
rsTables.Close
%>


