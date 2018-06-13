<%@LANGUAGE="VBSCRIPT"%> 
<!--#include virtual="/Connections/ecoinfo.asp" -->
<%
Dim theText,intern
theText=""
IF LEN(session("eiuserid"))=0 THEN
intern=1
ELSE
intern=0
END IF
%>
<%
set rsText = Server.CreateObject("ADODB.Recordset")
rsText.ActiveConnection = MM_ecoinfo_STRING
rsText.Source = "SELECT *  FROM eihelp_text WHERE id=0 AND intern=" & intern & " ORDER BY name"
if request("search")<>"" then
rsText.Source = replace(rsText.Source,"id=0","(name LIKE '%" & request("search") & "%' OR thetext LIKE '%" & request("search") & "%')")
elseif request("id")<>"" then
rsText.Source = replace(rsText.Source,"id=0","id=" & request("id"))
end if 

rsText.CursorType = 0
rsText.CursorLocation = 2
rsText.LockType = 3
rsText.Open()
rsText_numRows = 0

%>
<%
set rsSek = Server.CreateObject("ADODB.Recordset")
rsSek.ActiveConnection = MM_ecoinfo_STRING
rsSek.Source = "SELECT *  FROM eihelp_section WHERE id=0 AND intern=" & intern & " ORDER BY name"
if request("search")<>"" then
rsSek.Source = replace(rsSek.Source,"id=0","(name LIKE '%" & request("search") & "%' OR gen_text LIKE '%" & request("search") & "%')")
elseif request("secid")<>"" then
rsSek.Source = replace(rsSek.Source,"id=0","id=" & request("secid"))
end if 
rsSek.CursorType = 0
rsSek.CursorLocation = 2
rsSek.LockType = 3
rsSek.Open()
rsSek_numRows = 0
%>
<%
set rsBr = Server.CreateObject("ADODB.Recordset")
rsBr.ActiveConnection = MM_ecoinfo_STRING
rsBr.Source = "SELECT *  FROM eihelp_branch WHERE id=0 AND intern=" & intern & " ORDER BY name"
if request("search")<>"" then
rsBr.Source = replace(rsBr.Source,"id=0","(name LIKE '%" & request("search") & "%' OR gen_text LIKE '%" & request("search") & "%')")
elseif request("brid")<>"" then
rsBr.Source = replace(rsBr.Source,"id=0","id=" & request("brid"))
end if 
rsBr.CursorType = 0
rsBr.CursorLocation = 2
rsBr.LockType = 3
rsBr.Open()
rsBr_numRows = 0
%>
<%
set rsFunc = Server.CreateObject("ADODB.Recordset")
rsFunc.ActiveConnection = MM_ecoinfo_STRING
rsFunc.Source = "SELECT *  FROM eihelp_func WHERE id=0 AND intern=" & intern & " ORDER BY name"
if request("search")<>"" then
rsFunc.Source = replace(rsFunc.Source,"id=0","(name LIKE '%" & request("search") & "%' OR gen_text LIKE '%" & request("search") & "%')")
elseif request("funcid")<>"" then
rsFunc.Source = replace(rsFunc.Source,"id=0","id=" & request("funcid"))
end if 
rsFunc.CursorType = 0
rsFunc.CursorLocation = 2
rsFunc.LockType = 3
rsFunc.Open()
rsFunc_numRows = 0
%>
<%
set rsDetail = Server.CreateObject("ADODB.Recordset")
rsDetail.ActiveConnection = MM_ecoinfo_STRING
rsDetail.Source = "SELECT *  FROM eihelp_detail WHERE id=0 AND intern=" & intern & " ORDER BY name"
if request("search")<>"" then
rsDetail.Source = replace(rsDetail.Source,"id=0","(name LIKE '%" & request("search") & "%' OR gen_text LIKE '%" & request("search") & "%')")
elseif request("detailid")<>"" then
rsDetail.Source = replace(rsDetail.Source,"id=0","id=" & request("detailid"))
end if 
rsDetail.CursorType = 0
rsDetail.CursorLocation = 2
rsDetail.LockType = 3
rsDetail.Open()
rsDetail_numRows = 0
%>
<%
if intern=0 then
theText="Øko-net medarbejdernes hjælpeprogram!"
elseif intern=1 then
theText="Brugernes hjælpeprogram!"
end if
if request("id")<>"" then
theText=rsText("thetext")
end if
if request("secid")<>"" then
theText=rsSek("gen_text")
end if
if request("brid")<>"" then
theText=rsBr("gen_text")
end if
if request("funcid")<>"" then
theText=rsFunc("gen_text")
end if
if request("detailid")<>"" then
theText=rsDetail("gen_text")
end if
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = -1
Dim Repeat1__index
Repeat1__index = 0
rsText_numRows = rsText_numRows + Repeat1__numRows
%>
<%
Dim Repeat2__numRows
Repeat2__numRows = -1
Dim Repeat2__index
Repeat2__index = 0
rsSek_numRows = rsSek_numRows + Repeat2__numRows
%>
<%
Dim Repeat3__numRows
Repeat3__numRows = -1
Dim Repeat3__index
Repeat3__index = 0
rsBr_numRows = rsBr_numRows + Repeat3__numRows
%>
<%
Dim Repeat4__numRows
Repeat4__numRows = -1
Dim Repeat4__index
Repeat4__index = 0
rsFunc_numRows = rsFunc_numRows + Repeat4__numRows
%>
<%
Dim Repeat5__numRows
Repeat5__numRows = -1
Dim Repeat5__index
Repeat5__index = 0
rsDetail_numRows = rsDetail_numRows + Repeat5__numRows
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
<td width="10%" bgcolor="#cccccc" height="27"> 
<div align="center"><b><a href="/shared/help/help.asp">Browse</a></b><br>
</div>
</td>
<td width="10%" bgcolor="#eeeeee" height="27"> 
<div align="center"><b>S&oslash;g</b><br>
</div>
</td>
<td width="15%" height="27">&nbsp;</td>
<td height="27"><a href="javascript:history.go(-1)">Tilbage</a></td>
</tr>
<tr> 
<td valign="top" colspan="3" class="listheader" bgcolor="#eeeeee" height="20">&nbsp;</td>
<td width="55%" valign="top" bgcolor="#eeeeee">&nbsp;</td>
</tr>
<tr> 
<td valign="top" colspan="2" height="20%"> <br>
<form name="form1" method="post" action="help_search.asp">
<div align="center"> 
<input type="text" name="search" class="formInputobject">
<input type="submit" name="Submit" value="S&oslash;g" class="formSubmitbutton">
</div>
</form>
<p> 
<% If Not rsText.EOF Or Not rsText.BOF Then %>
<% 
While ((Repeat1__numRows <> 0) AND (NOT rsText.EOF)) 
%>
<a href="help_search.asp?id=<%=(rsText.Fields.Item("id").Value)%>"><br>
<%=(rsText.Fields.Item("name").Value)%></a> <br>
<% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsText.MoveNext()
Wend
%>
<% End If ' end Not rsText.EOF Or NOT rsText.BOF %>
<% If Not rsSek.EOF Or Not rsSek.BOF Then %>
<% 
While ((Repeat2__numRows <> 0) AND (NOT rsSek.EOF)) 
%>
<a href="help_search.asp?secid=<%=(rsSek.Fields.Item("id").Value)%>"><%=(rsSek.Fields.Item("name").Value)%></a> <br>
<% 
  Repeat2__index=Repeat2__index+1
  Repeat2__numRows=Repeat2__numRows-1
  rsSek.MoveNext()
Wend
%>
<% End If ' end Not rsSek.EOF Or NOT rsSek.BOF %>
<% If Not rsBr.EOF Or Not rsBr.BOF Then %>
<% 
While ((Repeat3__numRows <> 0) AND (NOT rsBr.EOF)) 
%>
<a href="help_search.asp?brid=<%=(rsBr.Fields.Item("id").Value)%>"><%=(rsBr.Fields.Item("name").Value)%></a> <br>
<% 
  Repeat3__index=Repeat3__index+1
  Repeat3__numRows=Repeat3__numRows-1
  rsBr.MoveNext()
Wend
%>
<% End If ' end Not rsBr.EOF Or NOT rsBr.BOF %>
<% If Not rsFunc.EOF Or Not rsFunc.BOF Then %>
<% 
While ((Repeat4__numRows <> 0) AND (NOT rsFunc.EOF)) 
%>
<a href="help_search.asp?funcid=<%=(rsFunc.Fields.Item("id").Value)%>"><%=(rsFunc.Fields.Item("name").Value)%></a> <br>
<% 
  Repeat4__index=Repeat4__index+1
  Repeat4__numRows=Repeat4__numRows-1
  rsFunc.MoveNext()
Wend
%>
<% End If ' end Not rsFunc.EOF Or NOT rsFunc.BOF %>
<% 
While ((Repeat5__numRows <> 0) AND (NOT rsDetail.EOF)) 
%>
<% If Not rsDetail.EOF Or Not rsDetail.BOF Then %>
<a href="help_search.asp?detailid=<%=(rsDetail.Fields.Item("id").Value)%>"><%=(rsDetail.Fields.Item("name").Value)%></a><br>
<% End If ' end Not rsDetail.EOF Or NOT rsDetail.BOF %>
<% 
  Repeat5__index=Repeat5__index+1
  Repeat5__numRows=Repeat5__numRows-1
  rsDetail.MoveNext()
Wend
%>
</p>
</td>
<td valign="top" width="15%">&nbsp; </td>
<td width="55%" valign="top"><%=theText%> </td>
</tr>
</table>
<p>&nbsp; </p>
</body>
</html>
<%
rsText.Close()
%>
<%
rsSek.Close()
%>
<%
rsBr.Close()
%>
<%
rsFunc.Close()
%>
<%
rsDetail.Close()
%>
