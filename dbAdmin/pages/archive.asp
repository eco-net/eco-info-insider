<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/dbadmin/connections/dbAdmin.asp" -->
<!--#include virtual="/dbadmin/pages/act_archive.asp" -->

<%
Dim rsSaved__ColParam1
rsSaved__ColParam1 = "x"
if (theCat <> "") then rsSaved__ColParam1 = theCat
%>
<%
Dim rsSaved__ColParam2
rsSaved__ColParam2 = "x"
if (theName <> "") then rsSaved__ColParam2 = theName
%>
<%
Dim rsSaved__ColParam3
rsSaved__ColParam3 = "0"
if (theType <> "") then rsSaved__ColParam3 = theType
%>
<%
Dim rsSaved__ColParam4
rsSaved__ColParam4 = "%"
if (theCode <> "") then rsSaved__ColParam4 = theCode
%>
<%
set rsSaved = Server.CreateObject("ADODB.Recordset")
rsSaved.ActiveConnection = MM_dbAdmin_STRING
rsSaved.Source = "SELECT *  FROM sqlstatements  WHERE category LIKE '" + Replace(rsSaved__ColParam1, "'", "''") + "%' AND isLink=" + Replace(rsSaved__ColParam3, "'", "''") + " AND  (name LIKE '%" + Replace(rsSaved__ColParam2, "'", "''") + "%' OR code LIKE '%" + Replace(rsSaved__ColParam4, "'", "''") + "%')  ORDER BY category, name"
rsSaved.CursorType = 0
rsSaved.CursorLocation = 2
rsSaved.LockType = 3
rsSaved.Open()
rsSaved_numRows = 0
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = -1
Dim Repeat1__index
Repeat1__index = 0
rsSaved_numRows = rsSaved_numRows + Repeat1__numRows
%>
<%
Dim Repeat2__numRows
Repeat2__numRows = -1
Dim Repeat2__index
Repeat2__index = 0
rsSaved_numRows = rsSaved_numRows + Repeat2__numRows
%>
<html><!-- #BeginTemplate "/Templates/list.dwt" --><!-- DW6 -->
<head>
<!-- #BeginEditable "doctitle" --> 
<title>Statement archive</title>
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
<%curTab=3%>
<!--#include virtual="/dbadmin/shared/menu.asp"-->
<!-- #EndEditable -->
<!-- END Menu -->

<!-- START ContentHeader -->
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
<td background="/dbadmin/images/layout/page_l.gif"><img src="/dbadmin/images/layout/spacer.gif" border="0" width="10" height="1"></td>
<td class="tableheader"  width="98%"> <!-- #BeginEditable "overskrift" --> <%=theTitle%> <!-- #EndEditable --></td>
<td align="right"> <!-- #BeginEditable "navigation" --> 
<table border="0" cellspacing="0" cellpadding="0">
<form name="doit" method="post" action="archive.asp">
<tr> 
<td nowrap>Show: 
<select name="thefilter" class="inputobject">
<option value=""></option>
<option value="a">Show all (saved & log)</option>
<option value="Table-definitions">Show table-definitions</option>
<option value="Index-definitions">Show index-definitions</option>
<option value="Procedure-definitions">Show procedure-definitions</option>
<option value="Procedure-examples">Show procedure-examples</option>
<option value="i">Show infopages</option>
<option value=""></option>
<option value="x">Show executable</option>
<option value="t">Show templates</option>
<option value="l">Show log</option>
</select>
containing 
<input type="text" name="theText" class="inputobjectsmall">
<input type="submit" name="go" value="go" class="inputsubmit">
&nbsp;&nbsp;<a href="archive.asp?thefilter=xl">Clear log</a> </td>
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
<% If Not rsSaved.EOF Or Not rsSaved.BOF Then %>
<%IF showLog=0 THEN%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr class="listlabelbg"> 
<td class="formlabelsleft">Name</td>
<td class="formlabelsleft">Category</td>
<td class="formlabelsleft">Saved as</td>
<td class="formlabelsleft">Description</td>
<td class="formlabelsleft"></td>
</tr>
<% 
While ((Repeat1__numRows <> 0) AND (NOT rsSaved.EOF)) 
%>
<tr> 
<td class="listtext"><%IF rsSaved.Fields.Item("name").Value="x" THEN
	response.write replace(rsSaved.Fields.Item("code").Value & " ",chr(13),"<br>")
ELSE
	response.write rsSaved.Fields.Item("name").Value
END IF
%></td>
<td class="listtext"><%=(rsSaved.Fields.Item("category").Value)%></td>
<td class="listtext"> 
<%IF rsSaved.Fields.Item("isLink").Value=1 THEN 
	response.write "Executable"
ELSEIF rsSaved.Fields.Item("isLink").Value=2 THEN
	response.write "Template"
ELSE
	response.write "Normal"
END IF%>
</td>
<td class="listtext"> 
<%response.write left(rsSaved.Fields.Item("Description").Value,50)
IF len(rsSaved.Fields.Item("Description").Value)>50 THEN response.write " ..."%>
</td>
<td class="listtext"> 
<%
response.write "<a href=""sql.asp?id=" & rsSaved.Fields.Item("id").Value & "&fetch=1"">Fetch</a>&nbsp;&nbsp;"
IF rsSaved.Fields.Item("isLink").Value=1 THEN
	response.write "<a href=""result.asp?id=" & rsSaved.Fields.Item("id").Value & """>Execute</a>"
ELSEIF rsSaved.Fields.Item("isLink").Value=2 THEN
	response.write "<a href=""sql.asp?id=" & rsSaved.Fields.Item("id").Value & """>Use</a>"
END IF
	response.write "&nbsp;&nbsp;<a href=""archive.asp?thefilter=" & request("theFilter") & "&delid=" & rsSaved.Fields.Item("id").Value & """><img src=""/dbadmin/images/listdel.gif"" border=""0""></a>"
%>
</td>
</tr>
<tr>
<td colspan="2" height="2"><img src="/dbadmin/images/spacer.gif" height="2" width="1" border="0"></td>
</tr>
<% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsSaved.MoveNext()
Wend
%>
</table>
<%ELSE
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr class="listlabelbg"> 
<td class="formlabelsleft">Code</td>
<td class="formlabelsleft">Datetime</td>
</tr>
<% 
While ((Repeat2__numRows <> 0) AND (NOT rsSaved.EOF)) 
%>
<tr> 
<td class="listtext"><%=replace(rsSaved.Fields.Item("code").Value & " ",chr(13),"<br>")%></td>
<td class="listtext" nowrap><%=(rsSaved.Fields.Item("logtime").Value)%></td>
</tr>
<tr>
<td colspan="2" height="2"><img src="/dbadmin/images/spacer.gif" height="2" width="1" border="0"></td>
</tr>
<% 
  Repeat2__index=Repeat2__index+1
  Repeat2__numRows=Repeat2__numRows-1
  rsSaved.MoveNext()
Wend
%>
</table>
<%END IF%>
<% End If ' end Not rsSaved.EOF Or NOT rsSaved.BOF %>
<%IF request("theFilter")="xl" THEN response.write themessage%>

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
rsSaved.Close()
%>

