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
<textarea name="textfield" rows="15" class="inputarea">
<% 
DIM theColumns, theValues, theTable, theUpdate
theColumns=""
theValues=""
theUpdate=""
theTable=rsColumns.Fields.Item("TABLE_NAME").Value
While ((Repeat1__numRows <> 0) AND (NOT rsColumns.EOF)) 

	theColumns=theColumns & rsColumns.Fields.Item("COLUMN_NAME").Value

	IF Instr(CSTR(rsColumns.Fields.Item("DATA_TYPE").Value),"char")>0 THEN
		theValues=theValues & chr(13) & """'"" & replace(request(""" & rsColumns.Fields.Item("COLUMN_NAME").Value & """),""'"",""''"") & ""'"""
		theUpdate=theUpdate & chr(13) & """" & rsColumns.Fields.Item("COLUMN_NAME").Value & "='"" & replace(request(""" &_
			rsColumns.Fields.Item("COLUMN_NAME").Value & """),""'"",""''"") & ""'"""

		theValues2=theValues2 & chr(13) & """'"" & replace(UploadRequest.Item(""" & rsColumns.Fields.Item("COLUMN_NAME").Value & """).Item(""Value""),""'"",""''"") & ""'"""
		theUpdate2=theUpdate2 & chr(13) & """" & rsColumns.Fields.Item("COLUMN_NAME").Value & "='"" & replace(UploadRequest.Item(""" &_
			rsColumns.Fields.Item("COLUMN_NAME").Value & """).Item(""Value""),""'"",""''"") & ""'"""
	ELSE
		theValues2=theValues2 & chr(13) & "UploadRequest.Item(""" & rsColumns.Fields.Item("COLUMN_NAME").Value & """).Item(""Value"")"
		IF rsColumns.Fields.Item("COLUMN_NAME").Value<>"id" THEN theUpdate2=theUpdate2 & chr(13) & """" & rsColumns.Fields.Item("COLUMN_NAME").Value & "="" & UploadRequest.Item(""" &_
			rsColumns.Fields.Item("COLUMN_NAME").Value & """).Item(""Value"")"

		theValues=theValues & chr(13) & "request(""" & rsColumns.Fields.Item("COLUMN_NAME").Value & """)"
		IF rsColumns.Fields.Item("COLUMN_NAME").Value<>"id" THEN theUpdate=theUpdate & chr(13) & """" & rsColumns.Fields.Item("COLUMN_NAME").Value & "="" & request(""" &_
			rsColumns.Fields.Item("COLUMN_NAME").Value & """)"
	END IF

	theColumns=theColumns & ","
	theValues=theValues & " & "","" &_"
	theValues2=theValues2 & " & "","" &_"
	IF rsColumns.Fields.Item("COLUMN_NAME").Value<>"id" THEN 
		theUpdate=theUpdate & " & "","" &_"
		theUpdate2=theUpdate2 & " & "","" &_"
	END IF

	Repeat1__index=Repeat1__index+1
	Repeat1__numRows=Repeat1__numRows-1
	rsColumns.MoveNext()
Wend

response.write ("""INSERT INTO " & theTable & "("" &_" & chr(13) &_
	"""" & LEFT(theColumns,LEN(theColumns)-1) & """ &_" & chr(13) & """) VALUES ("" &_"  &_
	LEFT(theValues,LEN(theValues)-8) & " &_" & chr(13) & """)""")
response.write (chr(13) & chr(13) & chr(13))

response.write ("""INSERT INTO " & theTable & "("" &_" & chr(13) &_
	"""" & LEFT(theColumns,LEN(theColumns)-1) & """ &_" & chr(13) & """) VALUES ("" &_"  &_
	LEFT(theValues2,LEN(theValues2)-8) & " &_" & chr(13) & """)""")
response.write (chr(13) & chr(13) & chr(13))

response.write ("""UPDATE " & theTable & " SET "" &_" & LEFT(theUpdate,LEN(theUpdate)-1) & " &_" & chr(13) &_
	""" WHERE id="" & request(""id"")")
response.write (chr(13) & chr(13) & chr(13))
response.write ("""UPDATE " & theTable & " SET "" &_" & LEFT(theUpdate2,LEN(theUpdate2)-1) & " &_" & chr(13) &_
	""" WHERE id="" & request(""id"")")

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




