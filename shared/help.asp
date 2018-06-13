<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/Connections/ecoinfo.asp" -->
<%
Dim theText
theText=""
%>
<%
set rsHelp = Server.CreateObject("ADODB.Recordset")
rsHelp.ActiveConnection = MM_ecoinfo_STRING
rsHelp.Source = "SELECT * FROM eihelp_section ORDER BY name"
rsHelp.CursorType = 0
rsHelp.CursorLocation = 2
rsHelp.LockType = 3
rsHelp.Open()
rsHelp_numRows = 0

%>
<%
theText="Øko-net medarbejdernes hjælpeprogram!"
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
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="/shared/styles.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000">
<p>Insiderhj&aelig;lp</p>
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
<tr> 
<td valign="top"> 
<% If Not rsHelp.EOF Or Not rsHelp.BOF Then %>
<% 
While ((Repeat1__numRows <> 0) AND (NOT rsHelp.EOF)) 
%>
<a href="help_sek.asp?sek=<%=(rsHelp.Fields.Item("id").Value)%>"><%=(rsHelp.Fields.Item("name").Value)%></a> <br>
<% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsHelp.MoveNext()
Wend
%>
<% End If ' end Not rsHelp.EOF Or NOT rsHelp.BOF %>

</td>
<td width="50%" valign="top"><b><%=theText%></b></td>
</tr>
</table>
<p>&nbsp; </p>
</body>
</html>
<%
rsHelp.Close()
%>
