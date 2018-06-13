<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include file="lib.asp"-->
<!--#include virtual="/Connections/ecoinfo.asp" -->
<%

'Dim exists,id,name
'set rs = Server.CreateObject("ADODB.Recordset")
'rs.ActiveConnection = MM_ecoinfo_STRING
'rs.Source = "SELECT id,name  FROM eiorg_maindata WHERE name LIKE '%" & request("name") & "%'"
'rs.CursorType = 2
'rs.CursorLocation = 3
'rs.LockType = 2
'rs.Open()
'if not rs.EOF then
'exists=true
'id=rs("id")
'else 
'exists=false
'end if
'rs.Close

'if exists=false then

DoSave 1
'else
'response.redirect "reject.asp?name=" & request("name") & "&id=" & id
'end if
%>
