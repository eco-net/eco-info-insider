<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/Connections/ecoinfo.asp" -->
<!--include virtual="/dbadmin/connections/voresDebat.asp" -->
<!--#include virtual="/dbadmin/pages/act_struct_tables.asp" -->
<%
Dim rsTables__ColParam
rsTables__ColParam = "%"
if (request("thename") <> "") then rsTables__ColParam = request("thename")
%>
<%
set rsTables = Server.CreateObject("ADODB.Recordset")
rsTables.ActiveConnection = MM_ecoinfo_STRING
'rsTables.ActiveConnection = MM_voresDebat_STRING
rsTables.Source = "SELECT *  FROM INFORMATION_SCHEMA.TABLES  WHERE TABLE_NAME LIKE '%" + Replace(rsTables__ColParam, "'", "''") + "%'"
rsTables.CursorType = 0
rsTables.CursorLocation = 2
rsTables.LockType = 3
rsTables.Open()
rsTables_numRows = 0
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = -1
Dim Repeat1__index
Repeat1__index = 0
rsTables_numRows = rsTables_numRows + Repeat1__numRows
%>
<html><!-- #BeginTemplate "/Templates/list.dwt" --><!-- DW6 -->
<head>
<!-- #BeginEditable "doctitle" --> 
<title>Tables and views</title>
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
<%curTab=4
curSub=1%>
<!--#include virtual="/dbadmin/shared/menu.asp"-->
<!-- #EndEditable -->
<!-- END Menu -->

<!-- START ContentHeader -->
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
<td background="/dbadmin/images/layout/page_l.gif"><img src="/dbadmin/images/layout/spacer.gif" border="0" width="10" height="1"></td>
<td class="tableheader"  width="98%"> <!-- #BeginEditable "overskrift" -->Tables 
and Views in the database<!-- #EndEditable --></td>
<td align="right"> <!-- #BeginEditable "navigation" --> 
<table border="0" cellspacing="0" cellpadding="0">
<form method="post" action="struct_tables.asp">
<tr> 
<td nowrap>Filter on name: 
<input type="text" name="thename" class="inputobject">
<input type="submit" name="Submit" value="go" class="normaltext">
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
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr class="listlabelbg"> 
<td class="formlabelsleft">Name</td>
<td class="formlabelsleft">Type</td>
<td class="formlabelsleft">Actions</td>
</tr>
<% 
While ((Repeat1__numRows <> 0) AND (NOT rsTables.EOF)) 
%>
<tr> 
<td class="listtext"><a href="javascript:openColumns('<%=(rsTables.Fields.Item("TABLE_NAME").Value)%>')"><%=(rsTables.Fields.Item("TABLE_NAME").Value)%></a></td>
<td class="listtext"><%=(rsTables.Fields.Item("TABLE_TYPE").Value)%></td>
<td class="listtext">
<a href="javascript:openColumns('<%=(rsTables.Fields.Item("TABLE_NAME").Value)%>')">Show structure</a>&nbsp;&nbsp;
<a href="javascript:showData('<%=(rsTables.Fields.Item("TABLE_NAME").Value)%>')">Show Data</a>&nbsp;&nbsp;
<a href="javascript:dropTable('<%=(rsTables.Fields.Item("TABLE_NAME").Value)%>')">Drop table</a>&nbsp;&nbsp;
<a href="javascript:showInfo('<%=(rsTables.Fields.Item("TABLE_NAME").Value)%>')">Info</a>&nbsp;&nbsp;
<a href="javascript:showInsert('<%=(rsTables.Fields.Item("TABLE_NAME").Value)%>')">SQL-code</a>&nbsp;&nbsp;
<a href="javascript:showForm('<%=(rsTables.Fields.Item("TABLE_NAME").Value)%>')">Form-code</a>&nbsp;&nbsp;
<a href="javascript:emptyTable('<%=(rsTables.Fields.Item("TABLE_NAME").Value)%>')">Empty table</a>

</td>
</tr>
<tr>
<td colspan="3" height="2"><img src="/dbadmin/images/spacer.gif" height="2" width="1" border="0"></td>
</tr>
<% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsTables.MoveNext()
Wend
%>
</table>
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
rsTables.Close()
%>
