<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/shared/sqlcheck.asp"-->
<!--#include virtual="/shared/inc_access.asp" -->
<!--#include virtual="/Connections/ecoinfo.asp" -->

<%
Dim rsPageData__MMColParam
rsPageData__MMColParam = "1"
if (Request.Cookies("eiorgid") <> "") then rsPageData__MMColParam = Request.Cookies("eiorgid")
%>

<%
set rsPageData = Server.CreateObject("ADODB.Recordset")
rsPageData.ActiveConnection = MM_ecoinfo_STRING
rsPageData.Source = "SELECT username, password, email FROM eisys_insiderusers WHERE orgid = " + Replace(rsPageData__MMColParam, "'", "''") + ""
rsPageData.CursorType = 0
rsPageData.CursorLocation = 2
rsPageData.LockType = 3
rsPageData.Open()
rsPageData_numRows = 0
%>
<html><!-- #BeginTemplate "/Templates/noheader.dwt" --><!-- DW6 -->
<head>
<!-- #BeginEditable "doctitle" --> 
<title>Ecoinfo</title>
<!-- #EndEditable --> 
<script src="/shared/common.js"></script>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="/shared/styles.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="7" marginwidth="0" marginheight="7">
<table width="98%" border="0" cellspacing="0" cellpadding="0" align="center" name="Pagetable">
<tr> 
<td background="/shared/graphics/layout/dots_vert.gif" width="1" valign="top"><img src="/shared/graphics/layout/cover_dots.gif" width="1" height="18"></td>
<td width="100%" valign="top"> 
<!-- START HEADER -->
<table width="100%" border="0" cellspacing="0" cellpadding="0" name="Header">
<tr> 
<td valign="top" rowspan="3" width="180" heigth="33"><img src="/shared/graphics/logo.gif" width="180" height="33"></td>
<td height="17">
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr>
<td align="left"><img src="/shared/graphics/logo2.gif" width="21" height="16"></td>
</tr>
</table>
</td>
</tr>
<tr>
<td valign="top" width="99%" background="/shared/graphics/layout/dots_horz.gif" height="1"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
<tr>
<td height="16"><br></td>
</tr>
</table>
<!-- END HEADER -->

<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr> 
<td width="100%" valign="top" name="maincontent">
<table border="0" cellspacing="0" cellpadding="10" width="100%" name="Contentarea">
<tr><td valign="top">
<!-- #BeginEditable "maincontent" -->
 
<table width="90%" height="100%" border="0" cellspacing="0" cellpadding="0" align="center">
<tr> 
<td valign="top"> 
<form name="form1" method="post" action="act_users.asp">
<div align="center"><br>
<span class="contentHeader2">&AElig;ndring af Bruger-oplysninger:</span>
<br>
<br>
<%IF request("done")=1 THEN response.write "Ændringerne er gemt."%>
</div>
<table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
<tr>
<td class="formLabeltext">
Brugernavn:
</td>
<td>
<input type="text" name="username" class="formInputobjectLong" value="<%=(rsPageData.Fields.Item("username").Value)%>">
</td>
</tr>
<tr>
<td class="formLabeltext">
Password:
</td>
<td>
<input type="text" name="password" class="formInputobjectLong" value="<%=(rsPageData.Fields.Item("password").Value)%>">
</td>
</tr>
<tr>
<td class="formLabeltext">
Email:
</td>
<td>
<input type="text" name="email" class="formInputobjectLong" value="<%=(rsPageData.Fields.Item("email").Value)%>">
</td>
</tr>
<tr>
<td class="formLabeltext">&nbsp;</td>
<td><br>
<input type="hidden" name="orgid" value="<%=request.cookies("eiorgid")%>">
<input type="submit" name="Submit" value="Gem" class="formSubmitbutton">
<input type="button" name="Submit2" value="Luk vindue" class="formbutton" onClick="window.close();">
</td>
</tr>
</table>
</form>
</td>
</tr>
</table>
<!-- #EndEditable -->
</td></tr>
</table>
</td>
</tr>
</table>
</td>
<td background="/shared/graphics/layout/dots_vert.gif" width="1" valign="top"><img src="/shared/graphics/layout/cover_dots.gif" width="1" height="18"></td>
</tr>
<tr height="1"> 
<td background="/shared/graphics/layout/dots_horz.gif" height="1" colspan="3"><img src="/shared/graphics/spacer.gif" width="3" height="1"></td>
</tr>
</table>
</body>
<!-- #EndTemplate --></html>
<%
rsPageData.Close()
%>

<!--#include virtual="/shared/log.asp"-->


