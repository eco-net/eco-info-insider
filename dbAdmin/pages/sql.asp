<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/dbadmin/connections/dbAdmin.asp" -->
<!--#include virtual="/dbadmin/pages/act_sql.asp" -->
<%
Dim rsExecute__MMColParam
rsExecute__MMColParam = "1"
if (Request("MM_EmptyValue") <> "") then rsExecute__MMColParam = Request("MM_EmptyValue")
%>
<%
set rsExecute = Server.CreateObject("ADODB.Recordset")
rsExecute.ActiveConnection = MM_dbAdmin_STRING
rsExecute.Source = "SELECT id, name FROM sqlstatements WHERE isLink = " + Replace(rsExecute__MMColParam, "'", "''") + " ORDER BY name ASC"
rsExecute.CursorType = 0
rsExecute.CursorLocation = 2
rsExecute.LockType = 3
rsExecute.Open()
rsExecute_numRows = 0
%>
<%
Dim rsTemplates__MMColParam
rsTemplates__MMColParam = "2"
if (Request("MM_EmptyValue") <> "") then rsTemplates__MMColParam = Request("MM_EmptyValue")
%>
<%
set rsTemplates = Server.CreateObject("ADODB.Recordset")
rsTemplates.ActiveConnection = MM_dbAdmin_STRING
rsTemplates.Source = "SELECT id, name FROM sqlstatements WHERE isLink = " + Replace(rsTemplates__MMColParam, "'", "''") + " ORDER BY name ASC"
rsTemplates.CursorType = 0
rsTemplates.CursorLocation = 2
rsTemplates.LockType = 3
rsTemplates.Open()
rsTemplates_numRows = 0
%>
<%
Dim rsTemplate__MMColParam
rsTemplate__MMColParam = "1"
if (Request("id") <> "") then rsTemplate__MMColParam = Request("id")
%>
<%
set rsTemplate = Server.CreateObject("ADODB.Recordset")
rsTemplate.ActiveConnection = MM_dbAdmin_STRING
rsTemplate.Source = "SELECT * FROM sqlstatements WHERE id = " + Replace(rsTemplate__MMColParam, "'", "''") + ""
rsTemplate.CursorType = 0
rsTemplate.CursorLocation = 2
rsTemplate.LockType = 3
rsTemplate.Open()
rsTemplate_numRows = 0
%>
<html><!-- #BeginTemplate "/Templates/list.dwt" --><!-- DW6 -->
<head>
<!-- #BeginEditable "doctitle" --> 
<title>Enter sql-statement</title>
<!-- #EndEditable -->
<script src="/dbadmin/shared/list.js"></script>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="/dbadmin/shared/styles.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="3" marginwidth="0" marginheight="3">
<table width="750" border="0" cellspacing="0" cellpadding="0" align="center"><tr>
<td valign="top">
<!-- START Menu -->
<!-- #BeginEditable "menu" --> 
<%curTab=1%>
<!--#include virtual="/dbadmin/shared/menu.asp"-->
<!-- #EndEditable -->
<!-- END Menu -->

<!-- START ContentHeader -->
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
<td background="/dbadmin/images/layout/page_l.gif"><img src="/dbadmin/images/layout/spacer.gif" border="0" width="10" height="1"></td>
<td class="tableheader"  width="98%"> <!-- #BeginEditable "overskrift" -->Enter 
SQL statement<!-- #EndEditable --></td>
<td align="right"> <!-- #BeginEditable "navigation" --> 
<table border="0" cellpadding="0" cellspacing="0">
<form name="form1" method="post" action="">
<tr> 
<td nowrap> Execute: 
<select name="execode" class="inputobject" onChange="form1.action='result.asp?id='+form1.execode.options[form1.execode.selectedIndex].value;form1.submit();">
<option value="0"></option>
<%
While (NOT rsExecute.EOF)
%>
<option value="<%=(rsExecute.Fields.Item("id").Value)%>"><%=(rsExecute.Fields.Item("name").Value)%></option>
<%
  rsExecute.MoveNext()
Wend
If (rsExecute.CursorType > 0) Then
  rsExecute.MoveFirst
Else
  rsExecute.Requery
End If
%>
</select>
</td>
<td nowrap> &nbsp;&nbsp;Get template: 
<select name="getcode" class="inputobject" onChange="form1.action='sql.asp?template=1&id='+form1.getcode.options[form1.getcode.selectedIndex].value;form1.submit();">
<option value="0"></option>
<%
While (NOT rsTemplates.EOF)
%>
<option value="<%=(rsTemplates.Fields.Item("id").Value)%>"><%=(rsTemplates.Fields.Item("name").Value)%></option>
<%
  rsTemplates.MoveNext()
Wend
If (rsTemplates.CursorType > 0) Then
  rsTemplates.MoveFirst
Else
  rsTemplates.Requery
End If
%>
</select>
</td>
</tr>
</form>
</table>
<!-- #EndEditable --> 
</td>
<td background="/dbadmin/images/layout/page_r.gif"><img src="/dbadmin/images/layout/spacer.gif" border="0" width="10" height="1"></td>
</tr>
<tr>
<td background="/dbadmin/images/layout/page_l.gif"><img src="/dbadmin/images/layout/spacer.gif" border="0" width="10" height="3"></td>
<td><img src="/dbadmin/images/layout/spacer.gif" border="0" width="10" height="3"></td>
<td><img src="/dbadmin/images/layout/spacer.gif" border="0" width="10" height="3"></td>
<td background="/dbadmin/images/layout/page_r.gif"><img src="/dbadmin/images/layout/spacer.gif" border="0" width="10" height="3"></td>
</tr>
</table>
<!-- END ContentHeader -->

<!-- START Content -->
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
<td background="/dbadmin/images/layout/page_l.gif"><img src="/dbadmin/images/layout/spacer.gif" border="0" width="10" height="3"></td>
<td width="99%">
<!-- #BeginEditable "indhold" --> 
<form name="theForm" action="result.asp" method="post" class="listtext">
<table class="listtext"><tr valign="top"><td>
Execute the code in a: 
<select name="objType" class="inputobject">
<option value="0" <%if rsTemplate.Fields.Item("ObjectType").Value=0 THEN response.write("Selected")%>>Recordset</option>
<option value="1" <%if rsTemplate.Fields.Item("ObjectType").Value=1 THEN response.write("Selected")%>>Command</option>
</select>

<textarea name="thesql" cols="60" rows="30" wrap="VIRTUAL" class="inputarea"><%=(rsTemplate.Fields.Item("code").Value)%></textarea>
</td>
<td><div align="right"><input type="submit" name="Submit" value="Submit" class="inputsubmit">
</div>
Name: <br>
<input type="text" name="theName" maxlength="50" size="50" class="inputobjectbig" value="<%IF rsTemplate.Fields.Item("isLink").Value<>2 OR len(request("fetch"))=1 THEN 
	response.write rsTemplate.Fields.Item("name").Value
END IF%>">
Category: <br>
<select name="category" class="inputobject">
<option value="0"></option>
<option value="Table-definitions" <%IF (rsTemplate.Fields.Item("isLink").Value<>2 OR len(request("fetch"))=1) AND CStr(rsTemplate.Fields.Item("Category").Value & " ")="Table-definitions " THEN 
	response.write "selected"
END IF%>>Table-definitions</option>
<option value="Index-definitions"<%IF (rsTemplate.Fields.Item("isLink").Value<>2 OR len(request("fetch"))=1) AND CStr(rsTemplate.Fields.Item("Category").Value & " ")="Index-definitions " THEN 
	response.write "selected"
END IF%>>Index-definitions</option>
<option value="Procedure-definitions"<%IF (rsTemplate.Fields.Item("isLink").Value<>2 OR len(request("fetch"))=1) AND CStr(rsTemplate.Fields.Item("Category").Value & " ")="Procedure-definitions " THEN 
	response.write "selected"
END IF%>>Procedure-definitions</option>
<option value="Procedure-examples"<%IF (rsTemplate.Fields.Item("isLink").Value<>2 OR len(request("fetch"))=1) AND CStr(rsTemplate.Fields.Item("Category").Value & " ")="Procedure-examples " THEN 
	response.write "selected"
END IF%>>Procedure-examples</option>
</select>
<br>
Description: <br>
<textarea name="thedescr" cols="15" rows="15" wrap="VIRTUAL" class="inputareahalf"><%IF rsTemplate.Fields.Item("isLink").Value<>2 OR len(request("fetch"))=1 THEN 
	response.write (rsTemplate.Fields.Item("Description").Value)
END IF%></textarea>
Save as:<br>
<table border="0" cellspacing="0" cellpadding="0" width="200"><tr><td valign="top">
<input type="radio" name="saveas" value="0" <%IF len(request("fetch"))=0 OR (len(request("fetch"))=1 AND CINT(rsTemplate.Fields.Item("isLink").Value)=0) THEN response.write "checked"%>>
Normal<br>
<input type="radio" name="saveas" value="2" <%IF len(request("fetch"))=1 AND CINT(rsTemplate.Fields.Item("isLink").Value)=2 THEN response.write "checked"%>>
Template
</td><td valign="top">
<input type="radio" name="saveas" value="1" <%IF len(request("fetch"))=1 AND CINT(rsTemplate.Fields.Item("isLink").Value)=1 THEN response.write "checked"%>>
Executable
<br>
<input type="radio" name="saveas" value="3" <%IF len(request("fetch"))=1 AND CINT(rsTemplate.Fields.Item("isLink").Value)=3 THEN response.write "checked"%>>
Info page
</td></tr></table>

<%IF len(request("fetch"))=1 THEN
	response.write "<div align=""right""><input type=""button"" value=""Execute"" class=""inputsubmit"" onclick=""theForm.action='sql.asp?dosave=1&exe=1';theForm.submit();"">&nbsp;&nbsp;<input type=""button"" value=""Save"" class=""inputsubmit"" onclick=""theForm.action='sql.asp?dosave=1';theForm.submit();""></div>"
END IF%>
<%=themessage%>
<%IF rsTemplate.Fields.Item("isLink").Value=2 AND len(request("fetch"))=0 THEN%>
<br>
<table width="200" border="0" cellpadding="0" cellspacing="0">
<tr> 
<td><img src="/dbadmin/images/layout/spacer.gif" height="1" width="1"></td>
<td colspan="3" class="box1frame" width="100%"><img src="/dbadmin/images/layout/spacer.gif" height="1" width="1"></td>
<td><img src="/dbadmin/images/layout/spacer.gif" height="1" width="1"></td>
</tr>
<tr> 
<td class="box1frame"><img src="/dbadmin/images/layout/spacer.gif" height="1" width="1"></td>
<td colspan="3" background="/dbadmin/images/layout/box1_headerbg.gif" class="box1header">Description</td>
<td class="box1frame"><img src="/dbadmin/images/layout/spacer.gif" height="1" width="1"></td>
</tr>
<tr>
<td colspan="5" class="box1frame"><img src="/dbadmin/images/layout/spacer.gif" height="1" width="1"></td>
</tr>
<tr> 
<td class="box1frame"><img src="/dbadmin/images/layout/spacer.gif" height="4" width="1"></td>
<td class="box1content" colspan="3"><img src="/dbadmin/images/layout/spacer.gif" height="1" width="1"></td>
<td class="box1frame"><img src="/dbadmin/images/layout/spacer.gif" height="1" width="1"></td>
</tr>
<tr> 
<td class="box1frame"><img src="/dbadmin/images/layout/spacer.gif" height="5" width="1"></td>
<td class="box1content"><img src="/dbadmin/images/layout/spacer.gif" height="1" width="4"></td>
<td class="box1content"><%=replace(rsTemplate.Fields.Item("Description").Value & " ",chr(13),"<br>")%> </td>
<td class="box1content"><img src="/dbadmin/images/layout/spacer.gif" height="1" width="4"></td>
<td class="box1frame"><img src="/dbadmin/images/layout/spacer.gif" height="1" width="1"></td>
</tr>
<tr> 
<td class="box1frame"><img src="/dbadmin/images/layout/spacer.gif" height="4" width="1"></td>
<td class="box1content" colspan="3"><img src="/dbadmin/images/layout/spacer.gif" height="1" width="1"></td>
<td class="box1frame"><img src="/dbadmin/images/layout/spacer.gif" height="1" width="1"></td>
</tr>
<tr> 
<td><img src="/dbadmin/images/layout/spacer.gif" height="1" width="1"></td>
<td colspan="3" class="box1frame" width="168"><img src="/dbadmin/images/layout/spacer.gif" height="1" width="1"></td>
<td><img src="/dbadmin/images/layout/spacer.gif" height="1" width="1"></td>
</tr>
</table>
<%END IF%>
<input type="hidden" name="id" value="<%if LEN(request("template"))=0 THEN response.write (request("id"))%>">
<input type="hidden" name="fetch" value="<%=request("fetch")%>">
</td>
</tr></table> 
</form>
<!-- #EndEditable -->
</td>
<td background="/dbadmin/images/layout/page_r.gif"><img src="/dbadmin/images/layout/spacer.gif" border="0" width="10" height="3"></td>
</tr>
</table>
<!-- END content -->

<!-- START Footer -->
<table border="0" width="100%" cellpadding="0" cellspacing="0">
<tr>
<td><img src="/dbadmin/images/layout/page_bl.gif" border="0" width="12" height="12"></td>
<td background="/dbadmin/images/layout/page_b.gif" width="99%"><img src="/dbadmin/images/layout/spacer.gif" border="0" width="1" height="12"></td>
<td><img src="/dbadmin/images/layout/page_br.gif" border="0" width="12" height="12"></td>
</tr>
</table>
</td></tr></table>
</body>
<!-- #EndTemplate --></html>
<%
rsExecute.Close()
%>
<%
rsTemplates.Close()
%>
<%
rsTemplate.Close()
%>
