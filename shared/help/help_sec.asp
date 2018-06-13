<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/Connections/ecoinfo.asp" -->
<%
Dim rsFolder__theID
rsFolder__theID = "0"
if (request("sec")   <> "") then rsFolder__theID = request("sec")  
%>
<%
set rsFolder = Server.CreateObject("ADODB.Recordset")
rsFolder.ActiveConnection = MM_ecoinfo_STRING
rsFolder.Source = "SELECT *  FROM eihelp_branch  WHERE sec_id=" + Replace(rsFolder__theID, "'", "''") + "  ORDER BY name"
rsFolder.CursorType = 0
rsFolder.CursorLocation = 2
rsFolder.LockType = 3
rsFolder.Open()
rsFolder_numRows = 0

%>
<%
Dim rsSek__theID
rsSek__theID = "0"
if (request("sec")  <> "") then rsSek__theID = request("sec") 
%>
<%
set rsSek = Server.CreateObject("ADODB.Recordset")
rsSek.ActiveConnection = MM_ecoinfo_STRING
rsSek.Source = "SELECT *  FROM eihelp_section  WHERE id=" + Replace(rsSek__theID, "'", "''") + ""
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
if request("sec")<>"" then
rsText.Source = replace(rsText.Source,"0=0","r.s=" & request("sec"))
elseif request("tekst")<>"" then
rsText.Source = replace(rsText.Source,"0=0","t.id=" & request("tekst"))
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
rsFolder_numRows = rsFolder_numRows + Repeat1__numRows
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
<td valign="top" colspan="3" class="listheader" height="20" bgcolor="#eeeeee"><%=(rsSek.Fields.Item("name").Value)%></td>
<td width="55%" valign="top" bgcolor="#eeeeee">&nbsp;</td>
</tr>
<tr> 
<td valign="top" width="15%"> 
<% If Not rsFolder.EOF Or Not rsFolder.BOF Then %>
<% 
While ((Repeat1__numRows <> 0) AND (NOT rsFolder.EOF)) 
%>
<a href="help_br.asp?sec=<%=request("sec")%>&secname=<%=rsSek("name")%>&br=<%=(rsFolder.Fields.Item("id").Value)%>"><%=(rsFolder.Fields.Item("name").Value)%></a> <br>
<% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsFolder.MoveNext()
Wend
%>
<% End If ' end Not rsFolder.EOF Or NOT rsFolder.BOF %>
</td>
<td valign="top" colspan="2"> 
<% If Not rsText.EOF Or Not rsText.BOF Then %>
<a href="help_sec.asp?sec=<%=request("sec")%>&secname=<%=request("secname")%>&tekst=<%=(rsText.Fields.Item("id").Value)%>">>><%=(rsText.Fields.Item("name").Value)%></a> 
<% End If ' end Not rsText.EOF Or NOT rsText.BOF %>
</td>
<td width="55%" valign="top"> 
<%if request("tekst")<>"" then 
rsText.MoveFirst
%>
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
rsFolder.Close()
%>
<%
rsSek.Close()
%>
<%
rsText.Close()
%>
