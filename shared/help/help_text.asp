<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/Connections/ecoinfo.asp" -->
<%
Dim rsSek__theID
rsSek__theID = "0"
if (request("text")     <> "") then rsSek__theID = request("text")    
%>
<%
set rsSek = Server.CreateObject("ADODB.Recordset")
rsSek.ActiveConnection = MM_ecoinfo_STRING
rsSek.Source = "SELECT *  FROM eihelp_text  WHERE id=" + Replace(rsSek__theID, "'", "''") + ""
rsSek.CursorType = 0
rsSek.CursorLocation = 2
rsSek.LockType = 3
rsSek.Open()
rsSek_numRows = 0
%>
<%
set rsText = Server.CreateObject("ADODB.Recordset")
rsText.ActiveConnection = MM_ecoinfo_STRING
rsText.Source = "SELECT t.id, t.name,t.thetext  FROM eihelp_text t LEFT JOIN eihelp_r_text r ON t.id=r.text_id  WHERE 0=0  ORDER BY t.name"
if request("detail")<>"" then
rsText.Source = replace(rsText.Source,"0=0","r.d=" & request("detail"))
elseif request("tekst")<>"" then
rsText.Source = replace(rsText.Source,"0=0","t.id=" & request("tekst"))
end if
rsText.CursorType = 0
rsText.CursorLocation = 2
rsText.LockType = 3
rsText.Open()
rsText_numRows = 0
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
<tr bgcolor="#eeeeee"> 
<td valign="top" colspan="3" class="listheader" height="20"><%=request("secname")%>/<%=request("brname")%>/<%=request("funcname")%>/<%=request("detailname")%>/<br>
<%=(rsSek.Fields.Item("name").Value)%></td>
<td width="55%" valign="top">&nbsp; </td>
</tr>
<tr> 
<td valign="top" width="5%">&nbsp;</td>
<td valign="top" width="15%"> <a href="help_text.asp?text="></a> <br>
</td>
<td valign="top" width="15%">&nbsp; </td>
<td width="55%" valign="top"><%=(rsSek.Fields.Item("thetext").Value)%></td>
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
