<%@LANGUAGE="VBSCRIPT"%>

<!--#include virtual="/Connections/ecoinfo.asp" -->

<%

set rsPageData = Server.CreateObject("ADODB.Recordset")
rsPageData.ActiveConnection = MM_ecoinfo_STRING
rsPageData.Source = "SELECT m.name,m.lastname,m.www FROM eiorg_maindata m  WHERE stopped=0 ORDER BY m.www"
'response.write rsPageData.Source
'response.end
rsPageData.CursorType = 0
rsPageData.CursorLocation = 2
rsPageData.LockType = 3
rsPageData.Open()
rsPageData_numRows = 0
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<%
i=0
o=1
while not rsPageData.EOF
%>
<tr>
	<td><%=o%></td>
    <td><%=rsPageData("name")%></td>
    <td><%=rsPageData("lastname")%></td>
    <td><%=rsPageData("www")%></td>
	<td><%=i%></td>
  </tr>
<%
if Len(rsPageData("www"))>3 then
i=i+1
end if
o=o+1
rsPageData.MoveNext
wend
rsPageData.Close()
%>
</table>