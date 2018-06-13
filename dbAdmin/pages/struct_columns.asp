<%@LANGUAGE="VBSCRIPT"%>
 <!--#include virtual="/Connections/ecoinfo.asp" -->
<!--include virtual="/dbadmin/connections/voresDebat.asp" -->
<%
Dim rsColumns__ColParam
rsColumns__ColParam = "%"
if (request("tablename") <> "") then rsColumns__ColParam = request("tablename")

%>
<%
set rsColumns = Server.CreateObject("ADODB.Recordset")
rsColumns.ActiveConnection = MM_ecoinfo_STRING
'rsColumns.ActiveConnection = MM_voresDebat_STRING
rsColumns.Source = "SELECT *  FROM INFORMATION_SCHEMA.COLUMNS  WHERE Table_Name LIKE '" + Replace(rsColumns__ColParam, "'", "''") + "'"
rsColumns.CursorType = 0
rsColumns.CursorLocation = 2
rsColumns.LockType = 3
rsColumns.Open()
rsColumns_numRows = 0
%>
<%
Dim rsIndexes__ColParam
rsIndexes__ColParam = "%"
if (request("tablename") <> "") then rsIndexes__ColParam = request("tablename")
%>
<%
Dim rsIndexes
Dim rsIndexes_cmd
Dim rsIndexes_numRows

Set rsIndexes_cmd = Server.CreateObject ("ADODB.Command")
rsIndexes_cmd.ActiveConnection = MM_ecoinfo_STRING
rsIndexes_cmd.CommandText = "SELECT * FROM INFORMATION_SCHEMA.KEY_COLUMN_USAGE WHERE Table_Name LIKE ?" 
rsIndexes_cmd.Prepared = true
rsIndexes_cmd.Parameters.Append rsIndexes_cmd.CreateParameter("param1", 200, 1, 255, rsIndexes__ColParam) ' adVarChar

Set rsIndexes = rsIndexes_cmd.Execute
rsIndexes_numRows = 0
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = -1
Dim Repeat1__index
Repeat1__index = 0
rsColumns_numRows = rsColumns_numRows + Repeat1__numRows
%>
<%
Dim Repeat2__numRows
Repeat2__numRows = -1
Dim Repeat2__index
Repeat2__index = 0
rsIndexes_numRows = rsIndexes_numRows + Repeat2__numRows
%>
<html><!-- #BeginTemplate "/Templates/Record.dwt" --><!-- DW6 -->
<head>
<!-- #BeginEditable "doctitle" --> 
<title>Untitled Document</title>
<!-- #EndEditable -->
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="/dbAdmin/shared/styles.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="3" marginwidth="0" marginheight="3">
<table width="99%" border="0" cellspacing="0" cellpadding="0" align="center"><tr>
<td valign="top">
<!-- START PageHeader -->
<table border="0" width="100%" cellpadding="0" cellspacing="0">
<tr>
<td><img src="/dbadmin/images/layout/page_tl.gif" border="0" width="12" height="12"></td>
<td width="99%" background="/dbadmin/images/layout/page_t.gif"><img src="/dbadmin/images/layout/spacer.gif" border="0" width="1" height="10"></td>
<td><img src="/dbadmin/images/layout/page_tr.gif" border="0" width="12" height="12"></td>
</tr>
</table>
<!-- END PageHeader -->

<!-- START ContentHeader -->
<table border="0" width="100%" cellpadding="0" cellspacing="0">
<tr>
<td background="/dbadmin/images/layout/page_l.gif"><img src="/dbadmin/images/layout/spacer.gif" border="0" width="10" height="1"></td>
<td width="98%" valign="top" class="tableheader">
<!-- #BeginEditable "overskrift" -->Columns 
in table '<%=(rsColumns.Fields.Item("TABLE_NAME").Value)%>'<!-- #EndEditable -->
</td>
<td align="right">
<!-- #BeginEditable "navigation" --> <!-- #EndEditable -->
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
<td background="/dbadmin/images/layout/page_l.gif"><img src="/dbadmin/images/layout/spacer.gif" border="0" width="10" height="1"></td>
<td width="98%">
<!-- #BeginEditable "indhold" --> 
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr class="listlabelbg"> 
<td class="formlabelsleft">Name</td>
<td class="formlabelsleft">Datatype</td>
<td class="formlabelsleft">Antal karakterer</td>
<td class="formlabelsleft">Default value</td>
<td class="formlabelsleft">Nullable</td>
</tr>
<% 
While ((Repeat1__numRows <> 0) AND (NOT rsColumns.EOF)) 
%>
<tr> 
<td class="listtext"><%=(rsColumns.Fields.Item("COLUMN_NAME").Value)%></td>
<td class="listtext"><%=(rsColumns.Fields.Item("DATA_TYPE").Value)%></td>
<td class="listtext"><%=(rsColumns.Fields.Item("CHARACTER_MAXIMUM_LENGTH").Value)%></td>
<td class="listtext"><%=(rsColumns.Fields.Item("COLUMN_DEFAULT").Value)%></td>
<td class="listtext"><%=(rsColumns.Fields.Item("IS_NULLABLE").Value)%></td>
</tr>
<% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsColumns.MoveNext()
Wend
%>
</table>
<br>
<% If Not rsIndexes.EOF Or Not rsIndexes.BOF Then %>
<div class="tableheader">Indexes</div>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr class="listlabelbg"> 
<td class="formlabelsleft">Name</td>
<td class="formlabelsleft">Column</td>
<td class="formlabelsleft">&nbsp;</td>
<td class="formlabelsleft">&nbsp;</td>
<td class="formlabelsleft">Nullable</td>
</tr>
<% 
While ((Repeat2__numRows <> 0) AND (NOT rsIndexes.EOF)) 
%>
<tr> 
<td class="listtext"><%=(rsIndexes.Fields.Item("CONSTRAINT_NAME").Value)%></td>
<td class="listtext"><%=(rsIndexes.Fields.Item("COLUMN_NAME").Value)%></td>
<td class="listtext"></td>
<td class="listtext"></td>
<td class="listtext"></td>
</tr>
<% 
  Repeat2__index=Repeat2__index+1
  Repeat2__numRows=Repeat2__numRows-1
  rsIndexes.MoveNext()
Wend
%>
</table>
<% End If ' end Not rsIndexes.EOF Or NOT rsIndexes.BOF %>
<% If rsIndexes.EOF And rsIndexes.BOF Then %>
<div class="tableheader">No Indexes</div>
<% End If ' end rsIndexes.EOF And rsIndexes.BOF %>
<!-- #EndEditable -->
</td>
<td background="/dbadmin/images/layout/page_r.gif"><img src="/dbadmin/images/layout/spacer.gif" border="0" width="10" height="1"></td>
</tr>
</table>
<!-- END content -->

<!-- START PageFooter -->
<table border="0" width="100%" cellpadding="0" cellspacing="0">
<tr>
<td><img src="/dbadmin/images/layout/page_bl.gif" border="0" width="12" height="12"></td>
<td background="/dbadmin/images/layout/page_b.gif" width="99%"><img src="/dbadmin/images/layout/spacer.gif" border="0" width="1" height="12"></td>
<td><img src="/dbadmin/images/layout/page_br.gif" border="0" width="12" height="12"></td>
</tr>
</table>
<!-- END PageFooter -->
</td></tr></table>
</body>
<!-- #EndTemplate --></html>
<%
rsColumns.Close()
%>
<%
rsIndexes.Close()
%>
