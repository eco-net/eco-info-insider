<%@LANGUAGE="VBSCRIPT"%> 
<!--#include virtual="/Connections/ecoinfo.asp" -->

<%
Dim theText,intern
theText=""
IF LEN(request.cookies("eiuserid"))=0 THEN
intern=1
ELSE
intern=0
END IF
%>
<%
set rsHelp = Server.CreateObject("ADODB.Recordset")
rsHelp.ActiveConnection = MM_ecoinfo_STRING
rsHelp.Source = "SELECT *  FROM eihelp_section WHERE intern=" & intern & " ORDER BY name"
rsHelp.CursorType = 0
rsHelp.CursorLocation = 2
rsHelp.LockType = 3
rsHelp.Open()
rsHelp_numRows = 0

%>
<%
if intern=0 then
theText="Øko-net medarbejdernes hjælpeprogram!"
elseif intern=1 then
theText="Brugernes hjælpeprogram!"
end if

%>
<%
Dim Repeat1__numRows
Repeat1__numRows = -1
Dim Repeat1__index
Repeat1__index = 0
rsHelp_numRows = rsHelp_numRows + Repeat1__numRows
%>
<html>
<head>
<title>Insider Hj&aelig;lp</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="/shared/styles.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000">
<p><b class="contentHeader2">Insiderhj&aelig;lp</b></p>
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
<tr> 
<td width="10%" bgcolor="#eeeeee" height="27"> 
<div align="center"><b>Browse</b><br>
</div>
</td>
<td width="10%" bgcolor="#999999" height="27"> 
<div align="center"><b><a href="/shared/help/help_search.asp">S&oslash;g</a></b><br>
</div>
</td>
<td width="15%" height="27">&nbsp;</td>
<td height="27">&nbsp;</td>
</tr>
<tr> 
<td valign="top" colspan="3" class="listheader" bgcolor="#eeeeee" height="20">&nbsp;</td>
<td width="55%" valign="top" bgcolor="#eeeeee">&nbsp;</td>
</tr>
<tr> 
<td valign="top" width="15%"> 
<% If Not rsHelp.EOF Or Not rsHelp.BOF Then %>
<% 
While ((Repeat1__numRows <> 0) AND (NOT rsHelp.EOF)) 
%>
<a href="help_sec.asp?sec=<%=(rsHelp.Fields.Item("id").Value)%>&secname=<%=rsHelp("name")%>"><%=(rsHelp.Fields.Item("name").Value)%></a> <br>
<% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsHelp.MoveNext()
Wend
%>
<% End If ' end Not rsHelp.EOF Or NOT rsHelp.BOF %>
</td>
<td valign="top" width="15%"> <br>
</td>
<td valign="top" width="5%">&nbsp; </td>
<td width="55%" valign="top"><%=theText%></td>
</tr>
</table>
<p>&nbsp; </p>
</body>
</html>
<%
rsHelp.Close()
%>
