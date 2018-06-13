<%@LANGUAGE="VBSCRIPT"%>
<!--include virtual="/dbadmin/connections/voresDebat.asp" -->
<!--#include virtual="/Connections/ecoinfo.asp" -->
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
Dim Repeat1__numRows
Repeat1__numRows = -1
Dim Repeat1__index
Repeat1__index = 0
rsColumns_numRows = rsColumns_numRows + Repeat1__numRows
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
<br>

<form name="form1" method="post" action="">
<textarea name="textfield" rows="15" class="inputareabig">
<% 
response.write "<form name='form1' method='post' action='' enctype='multipart/form-data'>" &_
	"<table border='0' cellspacing='0' cellpadding='0' class='functionstable'><tr>" &_
	"<td class='headerbox'>2</td><td class='functionsheader'>##</td></tr><tr><td>&nbsp;</td>" &_
	"<td class='functionslabel'>Ny ##</td></tr>" & chr(13) & chr(13)



DIM theColumns, theValues, theTable, theUpdate
theColumns=""
theValues=""
theUpdate=""
theTable=rsColumns.Fields.Item("TABLE_NAME").Value
While ((Repeat1__numRows <> 0) AND (NOT rsColumns.EOF)) 

	response.write "<tr><td>&nbsp;</td><td class='functionsitem'><b>" & rsColumns.Fields.Item("COLUMN_NAME").Value & "</b><br>"

	IF InStr(rsColumns.Fields.Item("COLUMN_NAME").Value,"imagepath")>0 THEN
		response.write Server.HTMLEncode("<input type='file' name='file' class='file'>")
	ELSEIF InStr(rsColumns.Fields.Item("COLUMN_NAME").Value,"category")>0 THEN
		response.write Server.HTMLEncode("<select name='" & rsColumns.Fields.Item("COLUMN_NAME").Value & "' class='dropdown'><!--include file=""/includes/menufiles/sel_##.asp""--></select>")
	ELSEIF rsColumns.Fields.Item("COLUMN_NAME").Value="url" THEN
		response.write Server.HTMLEncode("<input type='text' name='url' class='headingbox'><input type='hidden' name='urldescription'>" &_
			"<input type='button' name='Button' value='Vælg ...' class='button'  onClick=""window.open('/general/sel_link.asp','Link','width=300,height=200,top=200,left=200')"">")
	ELSEIF rsColumns.Fields.Item("COLUMN_NAME").Value="body" OR rsColumns.Fields.Item("COLUMN_NAME").Value="description" OR rsColumns.Fields.Item("COLUMN_NAME").Value="content" THEN
		response.write Server.HTMLEncode("<!--include file=""/includes/inc_textarea.htm""-->")
	ELSEIF rsColumns.Fields.Item("COLUMN_NAME").Value="urldescription" OR rsColumns.Fields.Item("COLUMN_NAME").Value="id" THEN
		response.write ""
	ELSE
		IF CINT("0" & rsColumns.Fields.Item("CHARACTER_MAXIMUM_LENGTH").Value)<100 THEN
			response.write Server.HTMLEncode("<input type='text' name='" & rsColumns.Fields.Item("COLUMN_NAME").Value & "' class='headingbox' value=''>")
		ELSEIF CINT("0" & rsColumns.Fields.Item("CHARACTER_MAXIMUM_LENGTH").Value)<250 THEN
			response.write Server.HTMLEncode("<textarea name='resume' class='textarea' rows='3'></textarea>")
		ELSE
			response.write Server.HTMLEncode("<textarea name='resume' class='textarea' rows='4'></textarea>")
		END IF
	END IF
	response.write "</td></tr>" & chr(13)
	
	rsColumns.movenext
Wend

response.write chr(13) & "<tr><td>&nbsp;</td><td class='functionsitem'><div align='center'><input type='button' value='Opret' class='button' onClick=""donew(this.form);"">" &_
	"<input type='button' value='Fortryd' class='button' onClick=""this.form.action='/functions_##.asp';this.form.submit();"">" &_
	"</div></td></tr></table></form>"
%>
</textarea>
</form>

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




