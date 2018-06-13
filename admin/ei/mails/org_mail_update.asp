<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/Connections/ecoinfo.asp" -->
<!--#include virtual="/shared/common.asp" -->
<%
set rsOrg = Server.CreateObject("ADODB.Recordset")
rsOrg.ActiveConnection = MM_ecoinfo_STRING
rsOrg.Source = "SELECT id, emailaddress1  FROM eiorg_maindata"
rsOrg.CursorType = 0
rsOrg.CursorLocation = 2
rsOrg.LockType = 3
rsOrg.Open()
rsOrg_numRows = 0
%>
<%
i=1
while NOT rsOrg.EOF
theSQL="UPDATE eisys_insiderusers SET email='" & rsOrg("emailaddress1") & "' WHERE orgid=" & rsOrg("id") 
			execCommand(theSQL)
			response.write "nr: " & i & "---" & rsOrg("id") & "<br>"
			i=i+1
rsOrg.MoveNext()
wend

%>


<%
rsOrg.Close()
%>

