<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/Connections/ecoinfo.asp" -->
<%
Dim rsSek__theID
rsSek__theID = "0"
if (request("detail")    <> "") then rsSek__theID = request("detail")   
%>
<%
set rsSek = Server.CreateObject("ADODB.Recordset")
rsSek.ActiveConnection = MM_ecoinfo_STRING
rsSek.Source = "SELECT *  FROM eihelp_detail  WHERE id=" + Replace(rsSek__theID, "'", "''") + ""
rsSek.CursorType = 0
rsSek.CursorLocation = 2
rsSek.LockType = 3
rsSek.Open()
rsSek_numRows = 0
%>
<%
set rsText = Server.CreateObject("ADODB.Recordset")
rsText.ActiveConnection = MM_ecoinfo_STRING
rsText.Source = "SELECT t.id, t.name,t.thetext  FROM eihelp_text t WHERE 0=0  ORDER BY t.name"
if Len(request("detail"))>0 then
rsText.Source = "SELECT t.id, t.name,t.thetext  FROM eihelp_text t WHERE t.detail_id=" & request("detail") & "  ORDER BY t.name"
end if
if Len(request("tekst"))>0 then
rsText.Source = "SELECT t.id, t.name,t.thetext  FROM eihelp_text t WHERE t.id=" & request("tekst") & "  ORDER BY t.name"
end if
rsText.CursorType = 0
rsText.CursorLocation = 2
rsText.LockType = 3
rsText.Open()
rsText_numRows = 0
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = -1
Dim Repeat1__index
Repeat1__index = 0
rsText_numRows = rsText_numRows + Repeat1__numRows
%>
<html>
<head>
<title>Insider Hj&aelig;lp</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="/shared/styles.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000">
<p class="contentHeader2">Insiderhj&aelig;lp</p>
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
<td height="27"><a href="javascript:history.go(-1)">Tilbage</a></td>
</tr>
<tr> 
<td valign="top" colspan="3" class="listheader" bgcolor="#eeeeee" height="20"><%=request("secname")%>/<%=request("brname")%>/<%=request("funcname")%>/<br>
<%=(rsSek.Fields.Item("name").Value)%></td>
<td width="55%" valign="top" bgcolor="#eeeeee">&nbsp;</td>
</tr>
<tr> 
<td valign="top" width="15%">&nbsp;</td>
<td valign="top" colspan="2"> 
<% 
While ((Repeat1__numRows <> 0) AND (NOT rsText.EOF)) 
%>
<% If Not rsText.EOF Or Not rsText.BOF Then %>
<a href="help_detail.asp?sec=<%=request("sec")%>&secname=<%=request("secname")%>&br=<%=request("br")%>&brname=<%=request("brname")%>&func=<%=request("func")%>&funcname=<%=request("funcname")%>&detail=<%=request("detail")%>&tekst=<%=(rsText.Fields.Item("id").Value)%>">>><%=(rsText.Fields.Item("name").Value)%></a><br>
<% End If ' end Not rsText.EOF Or NOT rsText.BOF %>
<% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsText.MoveNext()
Wend
%>
</td>
<td width="55%" valign="top"> 
<%if Len(request("tekst"))>0 then
rsText.MoveFirst %>
<%=(rsText.Fields.Item("thetext").Value)%> 
<%else %>
<%=(rsSek.Fields.Item("gen_text").Value)%> 
<% end if %>
</td>
</tr>
</table>
<p>&nbsp;</p>
</body>
</html>
<%
rsSek.Close()
%>
<%
rsText.Close()
%>
