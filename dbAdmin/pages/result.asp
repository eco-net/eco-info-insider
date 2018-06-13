<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/dbadmin/connections/dbAdmin.asp"-->
<!--#include virtual="/dbadmin/connections/voresDebat.asp"-->
<!--#include virtual="/dbadmin/pages/act_result.asp"-->
<html><!-- #BeginTemplate "/Templates/list.dwt" --><!-- DW6 -->
<head>
<!-- #BeginEditable "doctitle" --> 
<title>Result</title>
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
<%curTab=2%>
<!--#include virtual="/dbadmin/shared/menu.asp"-->
<!-- #EndEditable -->
<!-- END Menu -->

<!-- START ContentHeader -->
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
<td background="/dbadmin/images/layout/page_l.gif"><img src="/dbadmin/images/layout/spacer.gif" border="0" width="10" height="1"></td>
<td class="tableheader"  width="98%"> <!-- #BeginEditable "overskrift" -->Your 
sql-statement:<br>
<%="<div class=""formlabelsleft"">" & replace(thesql & " ",chr(13),"<br>") & "</div><br>"%>
Returned values:<br><br>
<!-- #EndEditable --></td>
<td align="right"> <!-- #BeginEditable "navigation" --> <!-- #EndEditable --> 
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
<%
IF isRS=1 THEN
'**** There's a recordset
	response.write "<table>" & theheader
	Dim iRowLoop, iColLoop
	For iRowLoop = 0 to UBound(theData, 2)
		response.write "<tr>"
		For iColLoop = 0 to UBound(theData, 1)
			Response.Write ("<td class=""listtext"">" & theData(iColLoop, iRowLoop) & "</td>")
		Next 'iColLoop
		Response.Write("</tr>")
	Next 'iRowLoop
	response.write "</table>"	
ELSE
'**** there's a respons
	response.write theResponse
END IF
%>
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
