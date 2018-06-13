<%@LANGUAGE="VBSCRIPT"%>
<!--include virtual="/dbadmin/connections/voresDebat.asp" -->
 <!--#include virtual="/Connections/ecoinfo.asp" -->
<!--#include virtual="/dbadmin/pages/act_cont_data.asp"-->
<html><!-- #BeginTemplate "/Templates/Record.dwt" --><!-- DW6 -->
<head>
<!-- #BeginEditable "doctitle" --> 
<title>Show data</title>
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
<!-- #BeginEditable "overskrift" --> 
Data in table '<%=request("tablename")%>'<!-- #EndEditable -->
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
<form action="cont_data.asp">
<input type="submit" name="Submit" value="Delete" class="inputsubmit">
<input type="hidden" name="delField" value="<%=theDelField%>">
<%
	response.write "<table>" & theheader
	Dim iRowLoop, iColLoop
	For iRowLoop = 0 to UBound(theData, 2)
		response.write "<tr>"
		response.write "<td class=""listtext""><input type=""checkbox"" name=""del"" value=""" & trim(theData(0, iRowLoop)) & """></td>"
		For iColLoop = 0 to UBound(theData, 1)
			Response.Write ("<td class=""listtext"">" & theData(iColLoop, iRowLoop) & "</td>")
		Next 'iColLoop
		Response.Write("</tr>")
	Next 'iRowLoop
	response.write "</table>"	
%>
<input type="submit" name="Submit" value="Delete" class="inputsubmit">
<input type="hidden" name="tablename" value="<%=request("tablename")%>">
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
