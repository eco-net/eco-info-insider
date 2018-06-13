<!--#include virtual="/Connections/ecoinfo.asp" -->
<%setLocale("da")%>
<%Dim mDato,mAntal
mDato=Date
mAntal=0
set rsM = Server.CreateObject("ADODB.Recordset")
rsM.ActiveConnection = MM_ecoinfo_STRING
rsM.Source = "SELECT * FROM eisys_newmember ORDER BY createdate"
rsM.CursorType = 0
rsM.CursorLocation = 2
rsM.LockType = 3
rsM.Open()
if not rsM.EOF then
WHILE not rsM.EOF
mDato=DatePart("d",rsM("createdate")) & "/" & DatePart("m",rsM("createdate")) & "-" & DatePart("yyyy",rsM("createdate"))
mAntal=mAntal+rsM("members")
rsM.MoveNext
wend
end if
%>